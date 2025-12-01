-- higiene.areas_medicion definition

CREATE TABLE `areas_medicion` (
  `id_area` bigint NOT NULL AUTO_INCREMENT,
  `nombre_area` varchar(100) NOT NULL,
  `categoria_area` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`id_area`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.centros_trabajo definition

CREATE TABLE `centros_trabajo` (
  `cuv` bigint NOT NULL,
  `rut` varchar(13) COLLATE utf8mb4_spanish_ci NOT NULL,
  `razon_social` varchar(100) COLLATE utf8mb4_spanish_ci NOT NULL,
  `rut2` varchar(11) COLLATE utf8mb4_spanish_ci NOT NULL,
  `nombre_ct` varchar(100) COLLATE utf8mb4_spanish_ci NOT NULL,
  `direccion_ct` varchar(200) COLLATE utf8mb4_spanish_ci NOT NULL,
  `comuna_ct` varchar(100) COLLATE utf8mb4_spanish_ci NOT NULL,
  `region_ct` varchar(100) COLLATE utf8mb4_spanish_ci NOT NULL,
  `region_num_ct` int NOT NULL,
  PRIMARY KEY (`cuv`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_spanish_ci;


-- higiene.equipos_medicion definition

CREATE TABLE `equipos_medicion` (
  `id_equipo` varchar(20) NOT NULL,
  `nombre_equipo` varchar(100) NOT NULL,
  `cod_equipo` varchar(50) DEFAULT NULL,
  `n_serie_equipo` varchar(50) DEFAULT NULL,
  `marca_equipo` varchar(50) DEFAULT NULL,
  `modelo_equipo` varchar(50) DEFAULT NULL,
  `estado_equipo` varchar(50) DEFAULT NULL,
  `obs_equipo` text,
  `estado_calibracion` varchar(50) DEFAULT NULL,
  `fecha_calibracion` date DEFAULT NULL,
  `prox_calibracion` date DEFAULT NULL,
  `empresa_certificadora` varchar(100) DEFAULT NULL,
  `num_certificado` varchar(50) DEFAULT NULL,
  `fecha_ingreso` date DEFAULT NULL,
  `observaciones` text,
  `simple_cod` varchar(50) DEFAULT NULL,
  `url_certificado` varchar(500) DEFAULT NULL,
  `equipo_dicc` varchar(10) DEFAULT NULL,
  `tipo` varchar(100) CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci NOT NULL,
  `patron_tbs` float NOT NULL,
  `patron_tbh` float NOT NULL,
  `patron_tg` float NOT NULL,
  PRIMARY KEY (`id_equipo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.motivo_evaluacion definition

CREATE TABLE `motivo_evaluacion` (
  `id_motivo` int NOT NULL AUTO_INCREMENT,
  `nombre_motivo` varchar(100) NOT NULL,
  PRIMARY KEY (`id_motivo`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.posicion_trabajador definition

CREATE TABLE `posicion_trabajador` (
  `id_posicion` int NOT NULL AUTO_INCREMENT,
  `nombre_posicion_trabajador` varchar(100) NOT NULL,
  `categoria_posicion_trabajador` varchar(100) NOT NULL,
  PRIMARY KEY (`id_posicion`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.puestos_trabajo definition

CREATE TABLE `puestos_trabajo` (
  `id_puesto_trabajo` int NOT NULL AUTO_INCREMENT,
  `nombre_puesto_trabajo` varchar(100) NOT NULL,
  `categoria_puesto_trabajo` varchar(100) NOT NULL,
  PRIMARY KEY (`id_puesto_trabajo`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.sector_especifico definition

CREATE TABLE `sector_especifico` (
  `id_sector_especifico` int NOT NULL AUTO_INCREMENT,
  `nombre_sector_especifico` varchar(100) NOT NULL,
  `categoria_sector_especifico` varchar(100) NOT NULL,
  PRIMARY KEY (`id_sector_especifico`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.usuarios definition

CREATE TABLE `usuarios` (
  `email` varchar(100) NOT NULL,
  `pass` varchar(200) NOT NULL,
  `name` varchar(100) NOT NULL,
  `type` int NOT NULL,
  `zonal` varchar(100) DEFAULT NULL,
  `cargo` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`email`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.vestimenta_trabajador definition

CREATE TABLE `vestimenta_trabajador` (
  `id_vestimenta_trabajador` int NOT NULL AUTO_INCREMENT,
  `nombre_vestimenta_trabajador` varchar(100) NOT NULL,
  `categoria_vestimenta_trabajador` varchar(100) CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci NOT NULL,
  PRIMARY KEY (`id_vestimenta_trabajador`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.visitas definition

CREATE TABLE `visitas` (
  `id_visita` bigint NOT NULL AUTO_INCREMENT,
  `cuv_visita` bigint NOT NULL,
  `fecha_visita` date NOT NULL,
  `hora_visita` time NOT NULL,
  `motivo_evaluacion` varchar(50) DEFAULT NULL,
  `nombre_personal_visita` varchar(100) DEFAULT NULL,
  `cargo_personal_visita` varchar(100) DEFAULT NULL,
  `consultor_ist` varchar(100) NOT NULL,
  `note_visita` varchar(1000) DEFAULT NULL,
  `consultor_cargo` varchar(100) DEFAULT NULL,
  `consultor_zonal` varchar(100) DEFAULT NULL,
  `tipo_evaluacion` enum('confort','ventilacion') NOT NULL DEFAULT 'confort' COMMENT 'Tipo de evaluación realizada',
  PRIMARY KEY (`id_visita`),
  KEY `fk_visitas_centro` (`cuv_visita`),
  KEY `fk_visitas_consultor` (`consultor_ist`),
  CONSTRAINT `fk_visitas_centro` FOREIGN KEY (`cuv_visita`) REFERENCES `centros_trabajo` (`cuv`),
  CONSTRAINT `fk_visitas_consultor` FOREIGN KEY (`consultor_ist`) REFERENCES `usuarios` (`email`)
) ENGINE=InnoDB AUTO_INCREMENT=49 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.ev_confort definition

CREATE TABLE `ev_confort` (
  `id_ev_confort` bigint NOT NULL AUTO_INCREMENT COMMENT 'ID único de la evaluación de confort',
  `visita_id` bigint NOT NULL COMMENT 'FK a la visita asociada',
  `temperatura_dia` decimal(5,2) DEFAULT NULL COMMENT 'Temperatura máxima del día en °C',
  `equipo_temp` varchar(20) DEFAULT NULL COMMENT 'Equipo de medición de temperatura',
  `equipo_vel_air` varchar(20) DEFAULT NULL COMMENT 'Equipo de medición de velocidad de aire',
  `patron_tbs` decimal(5,2) DEFAULT '46.40' COMMENT 'Patrón Temperatura Bulbo Seco (°C)',
  `ver_tbs_ini` decimal(5,2) DEFAULT NULL COMMENT 'Verificación TBS inicial (°C)',
  `patron_tbh` decimal(5,2) DEFAULT '12.70' COMMENT 'Patrón Temperatura Bulbo Húmedo (°C)',
  `ver_tbh_ini` decimal(5,2) DEFAULT NULL COMMENT 'Verificación TBH inicial (°C)',
  `patron_tg` decimal(5,2) DEFAULT '69.80' COMMENT 'Patrón Temperatura de Globo (°C)',
  `ver_tg_ini` decimal(5,2) DEFAULT NULL COMMENT 'Verificación TG inicial (°C)',
  `ver_tbs_fin` decimal(5,2) DEFAULT NULL COMMENT 'Verificación TBS final (°C)',
  `ver_tbh_fin` decimal(5,2) DEFAULT NULL COMMENT 'Verificación TBH final (°C)',
  `ver_tg_fin` decimal(5,2) DEFAULT NULL COMMENT 'Verificación TG final (°C)',
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Fecha de última actualización',
  PRIMARY KEY (`id_ev_confort`),
  UNIQUE KEY `unique_visita` (`visita_id`),
  KEY `idx_equipo_temp` (`equipo_temp`),
  KEY `idx_equipo_vel` (`equipo_vel_air`),
  CONSTRAINT `fk_ev_confort_equipo_temp` FOREIGN KEY (`equipo_temp`) REFERENCES `equipos_medicion` (`id_equipo`) ON DELETE SET NULL ON UPDATE CASCADE,
  CONSTRAINT `fk_ev_confort_equipo_vel` FOREIGN KEY (`equipo_vel_air`) REFERENCES `equipos_medicion` (`id_equipo`) ON DELETE SET NULL ON UPDATE CASCADE,
  CONSTRAINT `fk_ev_confort_visitas` FOREIGN KEY (`visita_id`) REFERENCES `visitas` (`id_visita`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Datos específicos de evaluaciones de confort térmico';


-- higiene.ev_ventilacion definition

CREATE TABLE `ev_ventilacion` (
  `id_ev_ventilacion` bigint NOT NULL AUTO_INCREMENT COMMENT 'ID único de la evaluación de ventilación',
  `visita_id` bigint NOT NULL COMMENT 'FK a la visita asociada',
  `equipo_temp` varchar(20) DEFAULT NULL COMMENT 'Equipo termohigrómetro utilizado en la evaluación',
  `equipo_vel_air` varchar(20) DEFAULT NULL COMMENT 'Equipo de medición de velocidad de aire',
  `instru_nombre_1` varchar(150) DEFAULT NULL COMMENT 'Nombre instrumento 1',
  `instru_marca_1` varchar(100) DEFAULT NULL COMMENT 'Marca instrumento 1',
  `instru_modelo_1` varchar(100) DEFAULT NULL COMMENT 'Modelo instrumento 1',
  `instru_nserie_1` varchar(100) DEFAULT NULL COMMENT 'Número de serie instrumento 1',
  `instru_ncertificado_1` varchar(100) DEFAULT NULL COMMENT 'Certificado instrumento 1',
  `instru_nombre_2` varchar(150) DEFAULT NULL COMMENT 'Nombre instrumento 2',
  `instru_marca_2` varchar(100) DEFAULT NULL COMMENT 'Marca instrumento 2',
  `instru_modelo_2` varchar(100) DEFAULT NULL COMMENT 'Modelo instrumento 2',
  `instru_nserie_2` varchar(100) DEFAULT NULL COMMENT 'Número de serie instrumento 2',
  `instru_ncertificado_2` varchar(100) DEFAULT NULL COMMENT 'Certificado instrumento 2',
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Fecha de última actualización',
  PRIMARY KEY (`id_ev_ventilacion`),
  UNIQUE KEY `unique_visita` (`visita_id`),
  KEY `idx_ev_vent_equipo_temp` (`equipo_temp`),
  KEY `idx_ev_vent_equipo_vel` (`equipo_vel_air`),
  CONSTRAINT `fk_ev_vent_equipo_temp` FOREIGN KEY (`equipo_temp`) REFERENCES `equipos_medicion` (`id_equipo`) ON DELETE SET NULL ON UPDATE CASCADE,
  CONSTRAINT `fk_ev_vent_equipo_vel` FOREIGN KEY (`equipo_vel_air`) REFERENCES `equipos_medicion` (`id_equipo`) ON DELETE SET NULL ON UPDATE CASCADE,
  CONSTRAINT `fk_ev_vent_visitas` FOREIGN KEY (`visita_id`) REFERENCES `visitas` (`id_visita`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Equipos utilizados en evaluaciones de ventilación';


-- higiene.mediciones definition

CREATE TABLE `mediciones` (
  `id_medicion` bigint NOT NULL AUTO_INCREMENT,
  `visita_id` bigint NOT NULL,
  `nombre_area` varchar(100) DEFAULT NULL,
  `sector_especifico` varchar(100) DEFAULT NULL,
  `puesto_trabajo` varchar(100) DEFAULT NULL,
  `posicion_trabajador` varchar(65) DEFAULT NULL,
  `vestimenta_trabajador` varchar(65) DEFAULT NULL,
  `t_bul_seco` decimal(4,2) DEFAULT NULL,
  `t_globo` decimal(4,2) DEFAULT NULL,
  `hum_rel` decimal(4,2) DEFAULT NULL,
  `vel_air` decimal(3,2) DEFAULT NULL,
  `ppd` decimal(4,2) DEFAULT NULL,
  `pmv` decimal(3,2) DEFAULT NULL,
  `resultado_medicion` varchar(100) CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci DEFAULT NULL,
  `cond_techumbre` tinyint(1) DEFAULT NULL,
  `obs_techumbre` varchar(1000) DEFAULT NULL,
  `cond_paredes` tinyint(1) DEFAULT NULL,
  `obs_paredes` varchar(1000) DEFAULT NULL,
  `cond_vantanal` tinyint(1) DEFAULT NULL,
  `obs_ventanal` varchar(1000) DEFAULT NULL,
  `cond_aire_acond` tinyint(1) DEFAULT NULL,
  `obs_aire_acond` varchar(1000) DEFAULT NULL,
  `cond_ventiladores` tinyint(1) DEFAULT NULL,
  `obs_ventiladores` varchar(1000) DEFAULT NULL,
  `cond_inyeccion_extraccion` tinyint(1) DEFAULT NULL,
  `obs_inyeccion_extraccion` varchar(1000) DEFAULT NULL,
  `cond_ventanas` tinyint(1) DEFAULT NULL,
  `obs_ventanas` varchar(1000) DEFAULT NULL,
  `cond_puertas` tinyint(1) DEFAULT NULL,
  `obs_puertas` varchar(1000) DEFAULT NULL,
  `cond_otras` tinyint(1) DEFAULT NULL,
  `obs_otras` varchar(1000) DEFAULT NULL,
  `met` float NOT NULL,
  `clo` float NOT NULL,
  `caract_constructivas` varchar(2000) DEFAULT NULL,
  `ingreso_salida_aire` varchar(2000) DEFAULT NULL,
  PRIMARY KEY (`id_medicion`),
  KEY `fk_mediciones_visita` (`visita_id`),
  CONSTRAINT `fk_mediciones_visita` FOREIGN KEY (`visita_id`) REFERENCES `visitas` (`id_visita`)
) ENGINE=InnoDB AUTO_INCREMENT=31 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;


-- higiene.v_areas definition

CREATE TABLE `v_areas` (
  `area_id` varchar(40) NOT NULL,
  `visita_id` bigint NOT NULL,
  `centro_id` bigint NOT NULL,
  `codigo_area` varchar(40) NOT NULL,
  `nombre_area` varchar(120) NOT NULL,
  `uso` varchar(60) DEFAULT NULL,
  `piso_nivel` varchar(60) DEFAULT NULL,
  `largo_m` decimal(8,2) DEFAULT NULL,
  `ancho_m` decimal(8,2) DEFAULT NULL,
  `alto_m` decimal(8,2) DEFAULT NULL,
  `volumen_m3` decimal(10,2) DEFAULT NULL,
  `aforo_permitido` int DEFAULT NULL,
  `m3_porpersona` decimal(10,2) DEFAULT NULL,
  `m3_porpersona_594` decimal(10,2) NOT NULL DEFAULT 10.00,
  `m3_porpersona_cumple` tinyint(1) NOT NULL DEFAULT 0,
  `caudal_inyeccion_total` decimal(12,2) NOT NULL DEFAULT 0.00,
  `caudal_extraccion_total` decimal(12,2) NOT NULL DEFAULT 0.00,
  `m3_porpersona_hora` decimal(12,2) DEFAULT NULL,
  `m3_porpersona_hora_594` decimal(12,2) NOT NULL DEFAULT 20.00,
  `m3_porpersona_hora_cumple` tinyint(1) NOT NULL DEFAULT 0,
  `recambio_hora_594_min` decimal(8,2) NOT NULL DEFAULT 6.00,
  `recambio_hora_594_max` decimal(8,2) NOT NULL DEFAULT 60.00,
  `recambio_hora` decimal(10,2) DEFAULT NULL,
  `recambio_hora_cumple` tinyint(1) NOT NULL DEFAULT 0,
  `ocupacion_habitual` int DEFAULT NULL,
  `ventilacion_tipo` varchar(30) DEFAULT NULL,
  `ventilacion_sistema` varchar(120) DEFAULT NULL,
  `ventilacion_estado` varchar(30) DEFAULT NULL,
  `aberturas` varchar(255) DEFAULT NULL,
  `croquis_url` varchar(500) DEFAULT NULL,
  `observaciones` text,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`area_id`),
  KEY `idx_v_areas_visita` (`visita_id`),
  KEY `idx_v_areas_centro` (`centro_id`),
  CONSTRAINT `fk_v_areas_centro` FOREIGN KEY (`centro_id`) REFERENCES `centros_trabajo` (`cuv`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `fk_v_areas_visita` FOREIGN KEY (`visita_id`) REFERENCES `visitas` (`id_visita`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Áreas evaluadas en visitas de ventilación';


-- higiene.v_puntos_medicion definition

CREATE TABLE `v_puntos_medicion` (
  `punto_id` varchar(60) NOT NULL,
  `evaluacion_id` bigint NOT NULL,
  `area_id` varchar(40) NOT NULL,
  `codigo_punto` varchar(40) NOT NULL,
  `tipo_punto` varchar(20) NOT NULL,
  `ubicacion_detalle` varchar(255) DEFAULT NULL,
  `altura_m` decimal(6,2) DEFAULT NULL,
  `distancia_fuente_m` decimal(6,2) DEFAULT NULL,
  `conducto_largo_cm` decimal(8,2) DEFAULT NULL,
  `conducto_ancho_cm` decimal(8,2) DEFAULT NULL,
  `conducto_diametro` decimal(8,2) DEFAULT NULL,
  `seccion_conducto_cm2` decimal(12,2) DEFAULT NULL,
  `medicion_caudal_1` decimal(8,3) DEFAULT NULL,
  `medicion_caudal_2` decimal(8,3) DEFAULT NULL,
  `medicion_caudal_3` decimal(8,3) DEFAULT NULL,
  `medicion_caudal_4` decimal(8,3) DEFAULT NULL,
  `medicion_caudal_5` decimal(8,3) DEFAULT NULL,
  `medicion_caudal_p` decimal(8,3) DEFAULT NULL,
  `caudal` decimal(12,3) DEFAULT NULL,
  `fecha_hora` datetime DEFAULT NULL,
  `condiciones_ocupacion` int DEFAULT NULL,
  `puertas_ventanas_abiertas` tinyint(1) DEFAULT NULL,
  `temperatura_c` decimal(6,2) DEFAULT NULL,
  `humedad_relativa_pct` decimal(6,2) DEFAULT NULL,
  `croquis_url` varchar(500) DEFAULT NULL,
  `observaciones` text,
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`punto_id`),
  KEY `idx_v_puntos_area` (`area_id`),
  KEY `idx_v_puntos_eval` (`evaluacion_id`),
  CONSTRAINT `fk_v_puntos_area` FOREIGN KEY (`area_id`) REFERENCES `v_areas` (`area_id`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `fk_v_puntos_visita` FOREIGN KEY (`evaluacion_id`) REFERENCES `visitas` (`id_visita`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='Puntos de medición asociados a visitas de ventilación';