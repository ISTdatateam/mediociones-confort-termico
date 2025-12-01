-- =====================================================================
-- MIGRACIÓN: Tabla ev_ventilacion
-- Descripción: Crea una tabla para almacenar los equipos utilizados en
--              las evaluaciones de ventilación.
-- Fecha: 2024
-- =====================================================================

START TRANSACTION;

CREATE TABLE IF NOT EXISTS ev_ventilacion (
    id_ev_ventilacion BIGINT AUTO_INCREMENT PRIMARY KEY COMMENT 'ID único de la evaluación de ventilación',
    visita_id BIGINT NOT NULL COMMENT 'FK a la visita asociada',
    equipo_temp VARCHAR(20) NULL COMMENT 'Equipo termohigrómetro utilizado en la evaluación',
    equipo_vel_air VARCHAR(20) NULL COMMENT 'Equipo de medición de velocidad de aire',
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Fecha de última actualización',

    UNIQUE KEY unique_visita (visita_id),
    KEY idx_ev_vent_equipo_temp (equipo_temp),
    KEY idx_ev_vent_equipo_vel (equipo_vel_air),

    CONSTRAINT fk_ev_vent_visitas
        FOREIGN KEY (visita_id) REFERENCES visitas(id_visita)
        ON DELETE CASCADE
        ON UPDATE CASCADE,

    CONSTRAINT fk_ev_vent_equipo_temp
        FOREIGN KEY (equipo_temp) REFERENCES equipos_medicion(id_equipo)
        ON DELETE SET NULL
        ON UPDATE CASCADE,

    CONSTRAINT fk_ev_vent_equipo_vel
        FOREIGN KEY (equipo_vel_air) REFERENCES equipos_medicion(id_equipo)
        ON DELETE SET NULL
        ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
COMMENT='Equipos utilizados en evaluaciones de ventilación';

COMMIT;
