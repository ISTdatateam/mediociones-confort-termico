-- =====================================================================
-- MIGRACIÓN: Separación de datos de confort térmico
-- Descripción: Extrae campos específicos de confort de la tabla visitas
--              hacia una nueva tabla ev_confort
-- Fecha: 2024
-- Versión: 1.3 - CORREGIDO DROP COLUMN syntax
-- =====================================================================

START TRANSACTION;

-- =====================================================================
-- 1. CREAR TABLA ev_confort
-- =====================================================================
CREATE TABLE IF NOT EXISTS ev_confort (
    id_ev_confort BIGINT AUTO_INCREMENT PRIMARY KEY COMMENT 'ID único de la evaluación de confort',
    visita_id BIGINT NOT NULL COMMENT 'FK a la visita asociada',

    -- Datos ambientales
    temperatura_dia DECIMAL(5,2) NULL COMMENT 'Temperatura máxima del día en °C',

    -- Equipos utilizados (FK a equipos_medicion)
    equipo_temp VARCHAR(20) NULL COMMENT 'Equipo de medición de temperatura',
    equipo_vel_air VARCHAR(20) NULL COMMENT 'Equipo de medición de velocidad de aire',

    -- Patrones de verificación
    patron_tbs DECIMAL(5,2) NULL DEFAULT 46.4 COMMENT 'Patrón Temperatura Bulbo Seco (°C)',
    ver_tbs_ini DECIMAL(5,2) NULL COMMENT 'Verificación TBS inicial (°C)',

    patron_tbh DECIMAL(5,2) NULL DEFAULT 12.7 COMMENT 'Patrón Temperatura Bulbo Húmedo (°C)',
    ver_tbh_ini DECIMAL(5,2) NULL COMMENT 'Verificación TBH inicial (°C)',

    patron_tg DECIMAL(5,2) NULL DEFAULT 69.8 COMMENT 'Patrón Temperatura de Globo (°C)',
    ver_tg_ini DECIMAL(5,2) NULL COMMENT 'Verificación TG inicial (°C)',

    -- Verificaciones finales (cierre)
    ver_tbs_fin DECIMAL(5,2) NULL COMMENT 'Verificación TBS final (°C)',
    ver_tbh_fin DECIMAL(5,2) NULL COMMENT 'Verificación TBH final (°C)',
    ver_tg_fin DECIMAL(5,2) NULL COMMENT 'Verificación TG final (°C)',

    -- Auditoría
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Fecha de última actualización',

    -- Constraints
    UNIQUE KEY unique_visita (visita_id),
    KEY idx_equipo_temp (equipo_temp),
    KEY idx_equipo_vel (equipo_vel_air),

    CONSTRAINT fk_ev_confort_visitas
        FOREIGN KEY (visita_id) REFERENCES visitas(id_visita)
        ON DELETE CASCADE
        ON UPDATE CASCADE,

    CONSTRAINT fk_ev_confort_equipo_temp
        FOREIGN KEY (equipo_temp) REFERENCES equipos_medicion(id_equipo)
        ON DELETE SET NULL
        ON UPDATE CASCADE,

    CONSTRAINT fk_ev_confort_equipo_vel
        FOREIGN KEY (equipo_vel_air) REFERENCES equipos_medicion(id_equipo)
        ON DELETE SET NULL
        ON UPDATE CASCADE

) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
COMMENT='Datos específicos de evaluaciones de confort térmico';

-- =====================================================================
-- 2. AGREGAR CAMPO tipo_evaluacion A visitas (si no existe)
-- =====================================================================
SET @column_exists = (
    SELECT COUNT(*)
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE()
    AND TABLE_NAME = 'visitas'
    AND COLUMN_NAME = 'tipo_evaluacion'
);

SET @sql = IF(@column_exists = 0,
    'ALTER TABLE visitas ADD COLUMN tipo_evaluacion ENUM(''confort'', ''ventilacion'') NOT NULL DEFAULT ''confort'' COMMENT ''Tipo de evaluación realizada'' AFTER consultor_zonal',
    'SELECT ''La columna tipo_evaluacion ya existe'' AS mensaje'
);

PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

-- Asegurar que todos los registros existentes tengan un tipo
UPDATE visitas
SET tipo_evaluacion = 'confort'
WHERE tipo_evaluacion IS NULL OR tipo_evaluacion = '';

-- =====================================================================
-- 3. MIGRAR DATOS HISTÓRICOS DE CONFORT
-- =====================================================================
INSERT INTO ev_confort (
    visita_id,
    temperatura_dia,
    equipo_temp,
    equipo_vel_air,
    patron_tbs,
    ver_tbs_ini,
    patron_tbh,
    ver_tbh_ini,
    patron_tg,
    ver_tg_ini,
    ver_tbs_fin,
    ver_tbh_fin,
    ver_tg_fin
)
SELECT
    v.id_visita,
    v.temperatura_dia,
    v.equipo_temp,
    v.equipo_vel_air,
    v.patron_tbs,
    v.ver_tbs_ini,
    v.patron_tbh,
    v.ver_tbh_ini,
    v.patron_tg,
    v.ver_tg_ini,
    v.ver_tbs_fin,
    v.ver_tbh_fin,
    v.ver_tg_fin
FROM visitas v
WHERE v.tipo_evaluacion = 'confort'
AND NOT EXISTS (
    SELECT 1
    FROM ev_confort ec
    WHERE ec.visita_id = v.id_visita
)
AND (
    v.temperatura_dia IS NOT NULL
    OR v.equipo_temp IS NOT NULL
    OR v.equipo_vel_air IS NOT NULL
);

-- =====================================================================
-- 4. VERIFICACIÓN ANTES DE ELIMINAR COLUMNAS
-- =====================================================================
SELECT
    'VERIFICACIÓN DE MIGRACIÓN' AS titulo,
    (SELECT COUNT(*) FROM ev_confort) AS 'Registros en ev_confort',
    (SELECT COUNT(*) FROM visitas WHERE tipo_evaluacion = 'confort') AS 'Visitas tipo confort',
    CASE
        WHEN (SELECT COUNT(*) FROM ev_confort) > 0 THEN '✓ OK - Datos migrados'
        WHEN (SELECT COUNT(*) FROM visitas WHERE tipo_evaluacion = 'confort') = 0 THEN '✓ OK - No hay datos de confort'
        ELSE '⚠ ADVERTENCIA - Verificar migración'
    END AS 'Estado';

-- =====================================================================
-- 5. ELIMINAR COLUMNAS DE CONFORT DE visitas
-- =====================================================================

-- Crear procedimiento temporal para eliminar columnas
DELIMITER $$

CREATE PROCEDURE IF NOT EXISTS drop_column_if_exists(
    IN p_table VARCHAR(64),
    IN p_column VARCHAR(64)
)
BEGIN
    DECLARE column_count INT;

    SELECT COUNT(*) INTO column_count
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE()
    AND TABLE_NAME = p_table
    AND COLUMN_NAME = p_column;

    IF column_count > 0 THEN
        SET @sql = CONCAT('ALTER TABLE ', p_table, ' DROP COLUMN ', p_column);
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    END IF;
END$$

DELIMITER ;

-- Eliminar todas las columnas de confort
CALL drop_column_if_exists('visitas', 'temperatura_dia');
CALL drop_column_if_exists('visitas', 'equipo_temp');
CALL drop_column_if_exists('visitas', 'equipo_vel_air');
CALL drop_column_if_exists('visitas', 'patron_tbs');
CALL drop_column_if_exists('visitas', 'ver_tbs_ini');
CALL drop_column_if_exists('visitas', 'patron_tbh');
CALL drop_column_if_exists('visitas', 'ver_tbh_ini');
CALL drop_column_if_exists('visitas', 'patron_tg');
CALL drop_column_if_exists('visitas', 'ver_tg_ini');
CALL drop_column_if_exists('visitas', 'ver_tbs_fin');
CALL drop_column_if_exists('visitas', 'ver_tbh_fin');
CALL drop_column_if_exists('visitas', 'ver_tg_fin');

-- Eliminar el procedimiento temporal
DROP PROCEDURE IF EXISTS drop_column_if_exists;

-- =====================================================================
-- 6. VERIFICACIÓN FINAL
-- =====================================================================
SELECT 'RESUMEN FINAL' AS titulo;

SELECT
    'ev_confort' AS tabla,
    COUNT(*) AS total_registros
FROM ev_confort
UNION ALL
SELECT
    'visitas (confort)' AS tabla,
    COUNT(*) AS total_registros
FROM visitas
WHERE tipo_evaluacion = 'confort'
UNION ALL
SELECT
    'mediciones' AS tabla,
    COUNT(*) AS total_registros
FROM mediciones;

-- Verificar estructura de ev_confort
SELECT 'ESTRUCTURA DE ev_confort' AS titulo;
DESCRIBE ev_confort;

-- =====================================================================
-- COMMIT
-- =====================================================================
COMMIT;

-- En caso de error ejecutar: ROLLBACK;