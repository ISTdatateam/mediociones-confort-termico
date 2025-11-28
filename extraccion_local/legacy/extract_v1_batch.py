#!/usr/bin/env python3
from pathlib import Path
import openpyxl
import csv
import sys
import math, re
import datetime as dt


SCHEMA = [
    {"row": 4, "col": 7,  "label": "RAZON SOCIAL EMPRESA", "type": "text"},
    {"row": 4, "col": 19, "label": "RUT EMPRESA",            "type": "text"},
    {"row": 6, "col": 7,  "label": "NOMBRE DEL CENTRO DE TRABAJO (CT)", "type": "text"},
    {"row": 6, "col": 19, "label": "HORARIO DE MEDICIÓN",    "type": "time"},   # <-- antes era "text"
    {"row": 8, "col": 7,  "label": "DIRECCIÓN (CT)",         "type": "text"},
    {"row": 10,"col": 7,  "label": "FECHA DE VISITA",        "type": "text"},
    {"row": 14,"col": 7,  "label": "NOMBRE PERSONA EMPRESA", "type": "text"},
    {"row": 14,"col": 19, "label": "NOMBRE PROFESIONAL IST", "type": "text"},
    {"row": 16,"col": 7,  "label": "CARGO PERSONA EMPRESA",  "type": "text"},
    {"row": 16,"col": 19, "label": "CARGO PROFESIONAL IST",  "type": "text"},

    {"row": 21,"col": 2,  "label": "Instru_nombre_1",        "type": "text"},
    {"row": 21,"col": 5,  "label": "Instru_marca_1",         "type": "text"},
    {"row": 21,"col": 7,  "label": "Instru_modelo_1",        "type": "text"},
    {"row": 21,"col": 9,  "label": "Instru_Nserie_1",        "type": "text"},
    {"row": 21,"col": 12, "label": "Instru_Ncertificado_1",  "type": "text"},

    {"row": 22,"col": 2,  "label": "Instru_nombre_2",        "type": "text"},
    {"row": 22,"col": 5,  "label": "Instru_marca_2",         "type": "text"},
    {"row": 22,"col": 7,  "label": "Instru_modelo_2",        "type": "text"},
    {"row": 22,"col": 9,  "label": "Instru_Nserie_2",        "type": "text"},
    {"row": 22,"col": 12, "label": "Instru_Ncertificado_2",  "type": "text"},

    {"row": 31,"col": 2,  "label": "Area_1",                 "type": "text"},
    {"row": 31,"col": 5,  "label": "N_personas_1",           "type": "integer"},  # <-- antes "number"
    {"row": 31,"col": 7,  "label": "Largo_1",                "type": "number"},
    {"row": 31,"col": 9,  "label": "Ancho_1",                "type": "number"},
    {"row": 31,"col": 11, "label": "Alto_1",                 "type": "number"},

    {"row": 35,"col": 2,  "label": "Area_2",                 "type": "text"},
    {"row": 35,"col": 5,  "label": "N_personas_2",           "type": "integer"},  # <-- antes "number"
    {"row": 35,"col": 7,  "label": "Largo_2",                "type": "number"},
    {"row": 35,"col": 9,  "label": "Ancho_2",                "type": "number"},
    {"row": 35,"col": 11, "label": "Alto_2",                 "type": "number"},

    {"row": 39,"col": 2,  "label": "Area_3",                 "type": "text"},
    {"row": 39,"col": 5,  "label": "N_personas_3",           "type": "integer"},  # <-- antes "number"
    {"row": 39,"col": 7,  "label": "Largo_3",                "type": "number"},
    {"row": 39,"col": 9,  "label": "Ancho_3",                "type": "number"},
    {"row": 39,"col": 11, "label": "Alto_3",                 "type": "number"},

    {"row": 43,"col": 2,  "label": "Area_4",                 "type": "text"},
    {"row": 43,"col": 5,  "label": "N_personas_4",           "type": "integer"},  # <-- antes "number"
    {"row": 43,"col": 7,  "label": "Largo_4",                "type": "number"},
    {"row": 43,"col": 9,  "label": "Ancho_4",                "type": "number"},
    {"row": 43,"col": 11, "label": "Alto_4",                 "type": "number"},

    # ===== Área 1 =====
    {"row": 52, "col": 2,  "label": "Area_1_bis", "type": "text"},
    {"row": 52, "col": 6,  "label": "Area_1_desc", "type": "text"},
    {"row": 53, "col": 15, "label": "Area_1_iny_1", "type": "text"},
    {"row": 54, "col": 15, "label": "Area_1_iny_2", "type": "text"},
    {"row": 55, "col": 15, "label": "Area_1_iny_3", "type": "text"},
    {"row": 56, "col": 15, "label": "Area_1_iny_4", "type": "text"},
    {"row": 57, "col": 15, "label": "Area_1_iny_5", "type": "text"},
    {"row": 59, "col": 15, "label": "Area_1_ext_1", "type": "text"},
    {"row": 60, "col": 15, "label": "Area_1_ext_2", "type": "text"},
    {"row": 61, "col": 15, "label": "Area_1_ext_3", "type": "text"},
    {"row": 62, "col": 15, "label": "Area_1_ext_4", "type": "text"},
    {"row": 63, "col": 15, "label": "Area_1_ext_5", "type": "text"},

    {"row": 53, "col": 19, "label": "Area_1_iny_largo_1", "type": "number"},
    {"row": 54, "col": 19, "label": "Area_1_iny_largo_2", "type": "number"},
    {"row": 55, "col": 19, "label": "Area_1_iny_largo_3", "type": "number"},
    {"row": 56, "col": 19, "label": "Area_1_iny_largo_4", "type": "number"},
    {"row": 57, "col": 19, "label": "Area_1_iny_largo_5", "type": "number"},
    {"row": 59, "col": 19, "label": "Area_1_ext_largo_1", "type": "number"},
    {"row": 60, "col": 19, "label": "Area_1_ext_largo_2", "type": "number"},
    {"row": 61, "col": 19, "label": "Area_1_ext_largo_3", "type": "number"},
    {"row": 62, "col": 19, "label": "Area_1_ext_largo_4", "type": "number"},
    {"row": 63, "col": 19, "label": "Area_1_ext_largo_5", "type": "number"},

    {"row": 53, "col": 20, "label": "Area_1_iny_ancho_1", "type": "number"},
    {"row": 54, "col": 20, "label": "Area_1_iny_ancho_2", "type": "number"},
    {"row": 55, "col": 20, "label": "Area_1_iny_ancho_3", "type": "number"},
    {"row": 56, "col": 20, "label": "Area_1_iny_ancho_4", "type": "number"},
    {"row": 57, "col": 20, "label": "Area_1_iny_ancho_5", "type": "number"},
    {"row": 59, "col": 20, "label": "Area_1_ext_ancho_1", "type": "number"},
    {"row": 60, "col": 20, "label": "Area_1_ext_ancho_2", "type": "number"},
    {"row": 61, "col": 20, "label": "Area_1_ext_ancho_3", "type": "number"},
    {"row": 62, "col": 20, "label": "Area_1_ext_ancho_4", "type": "number"},
    {"row": 63, "col": 20, "label": "Area_1_ext_ancho_5", "type": "number"},

    {"row": 53, "col": 21, "label": "Area_1_iny_diam_1", "type": "number"},
    {"row": 54, "col": 21, "label": "Area_1_iny_diam_2", "type": "number"},
    {"row": 55, "col": 21, "label": "Area_1_iny_diam_3", "type": "number"},
    {"row": 56, "col": 21, "label": "Area_1_iny_diam_4", "type": "number"},
    {"row": 57, "col": 21, "label": "Area_1_iny_diam_5", "type": "number"},
    {"row": 59, "col": 21, "label": "Area_1_ext_diam_1", "type": "number"},
    {"row": 60, "col": 21, "label": "Area_1_ext_diam_2", "type": "number"},
    {"row": 61, "col": 21, "label": "Area_1_ext_diam_3", "type": "number"},
    {"row": 62, "col": 21, "label": "Area_1_ext_diam_4", "type": "number"},
    {"row": 63, "col": 21, "label": "Area_1_ext_diam_5", "type": "number"},

    {"row": 53, "col": 23, "label": "Area_1_iny_vel_1_1", "type": "number"},
    {"row": 53, "col": 24, "label": "Area_1_iny_vel_1_2", "type": "number"},
    {"row": 53, "col": 25, "label": "Area_1_iny_vel_1_3", "type": "number"},
    {"row": 53, "col": 26, "label": "Area_1_iny_vel_1_4", "type": "number"},
    {"row": 53, "col": 27, "label": "Area_1_iny_vel_1_5", "type": "number"},
    {"row": 54, "col": 23, "label": "Area_1_iny_vel_2_1", "type": "number"},
    {"row": 54, "col": 24, "label": "Area_1_iny_vel_2_2", "type": "number"},
    {"row": 54, "col": 25, "label": "Area_1_iny_vel_2_3", "type": "number"},
    {"row": 54, "col": 26, "label": "Area_1_iny_vel_2_4", "type": "number"},
    {"row": 54, "col": 27, "label": "Area_1_iny_vel_2_5", "type": "number"},
    {"row": 55, "col": 23, "label": "Area_1_iny_vel_3_1", "type": "number"},
    {"row": 55, "col": 24, "label": "Area_1_iny_vel_3_2", "type": "number"},
    {"row": 55, "col": 25, "label": "Area_1_iny_vel_3_3", "type": "number"},
    {"row": 55, "col": 26, "label": "Area_1_iny_vel_3_4", "type": "number"},
    {"row": 55, "col": 27, "label": "Area_1_iny_vel_3_5", "type": "number"},
    {"row": 56, "col": 23, "label": "Area_1_iny_vel_4_1", "type": "number"},
    {"row": 56, "col": 24, "label": "Area_1_iny_vel_4_2", "type": "number"},
    {"row": 56, "col": 25, "label": "Area_1_iny_vel_4_3", "type": "number"},
    {"row": 56, "col": 26, "label": "Area_1_iny_vel_4_4", "type": "number"},
    {"row": 56, "col": 27, "label": "Area_1_iny_vel_4_5", "type": "number"},
    {"row": 57, "col": 23, "label": "Area_1_iny_vel_5_1", "type": "number"},
    {"row": 57, "col": 24, "label": "Area_1_iny_vel_5_2", "type": "number"},
    {"row": 57, "col": 25, "label": "Area_1_iny_vel_5_3", "type": "number"},
    {"row": 57, "col": 26, "label": "Area_1_iny_vel_5_4", "type": "number"},
    {"row": 57, "col": 27, "label": "Area_1_iny_vel_5_5", "type": "number"},

    {"row": 59, "col": 23, "label": "Area_1_ext_vel_1_1", "type": "number"},
    {"row": 59, "col": 24, "label": "Area_1_ext_vel_1_2", "type": "number"},
    {"row": 59, "col": 25, "label": "Area_1_ext_vel_1_3", "type": "number"},
    {"row": 59, "col": 26, "label": "Area_1_ext_vel_1_4", "type": "number"},
    {"row": 59, "col": 27, "label": "Area_1_ext_vel_1_5", "type": "number"},
    {"row": 60, "col": 23, "label": "Area_1_ext_vel_2_1", "type": "number"},
    {"row": 60, "col": 24, "label": "Area_1_ext_vel_2_2", "type": "number"},
    {"row": 60, "col": 25, "label": "Area_1_ext_vel_2_3", "type": "number"},
    {"row": 60, "col": 26, "label": "Area_1_ext_vel_2_4", "type": "number"},
    {"row": 60, "col": 27, "label": "Area_1_ext_vel_2_5", "type": "number"},
    {"row": 61, "col": 23, "label": "Area_1_ext_vel_3_1", "type": "number"},
    {"row": 61, "col": 24, "label": "Area_1_ext_vel_3_2", "type": "number"},
    {"row": 61, "col": 25, "label": "Area_1_ext_vel_3_3", "type": "number"},
    {"row": 61, "col": 26, "label": "Area_1_ext_vel_3_4", "type": "number"},
    {"row": 61, "col": 27, "label": "Area_1_ext_vel_3_5", "type": "number"},
    {"row": 62, "col": 23, "label": "Area_1_ext_vel_4_1", "type": "number"},
    {"row": 62, "col": 24, "label": "Area_1_ext_vel_4_2", "type": "number"},
    {"row": 62, "col": 25, "label": "Area_1_ext_vel_4_3", "type": "number"},
    {"row": 62, "col": 26, "label": "Area_1_ext_vel_4_4", "type": "number"},
    {"row": 62, "col": 27, "label": "Area_1_ext_vel_4_5", "type": "number"},
    {"row": 63, "col": 23, "label": "Area_1_ext_vel_5_1", "type": "number"},
    {"row": 63, "col": 24, "label": "Area_1_ext_vel_5_2", "type": "number"},
    {"row": 63, "col": 25, "label": "Area_1_ext_vel_5_3", "type": "number"},
    {"row": 63, "col": 26, "label": "Area_1_ext_vel_5_4", "type": "number"},
    {"row": 63, "col": 27, "label": "Area_1_ext_vel_5_5", "type": "number"},

    # ===== Área 2 =====
    {"row": 67, "col": 2,  "label": "Area_2_bis", "type": "text"},
    {"row": 67, "col": 6,  "label": "Area_2_desc", "type": "text"},
    {"row": 68, "col": 15, "label": "Area_2_iny_1", "type": "text"},
    {"row": 69, "col": 15, "label": "Area_2_iny_2", "type": "text"},
    {"row": 70, "col": 15, "label": "Area_2_iny_3", "type": "text"},
    {"row": 71, "col": 15, "label": "Area_2_iny_4", "type": "text"},
    {"row": 72, "col": 15, "label": "Area_2_iny_5", "type": "text"},
    {"row": 74, "col": 15, "label": "Area_2_ext_1", "type": "text"},
    {"row": 75, "col": 15, "label": "Area_2_ext_2", "type": "text"},
    {"row": 76, "col": 15, "label": "Area_2_ext_3", "type": "text"},
    {"row": 77, "col": 15, "label": "Area_2_ext_4", "type": "text"},
    {"row": 78, "col": 15, "label": "Area_2_ext_5", "type": "text"},

    {"row": 68, "col": 19, "label": "Area_2_iny_largo_1", "type": "number"},
    {"row": 69, "col": 19, "label": "Area_2_iny_largo_2", "type": "number"},
    {"row": 70, "col": 19, "label": "Area_2_iny_largo_3", "type": "number"},
    {"row": 71, "col": 19, "label": "Area_2_iny_largo_4", "type": "number"},
    {"row": 72, "col": 19, "label": "Area_2_iny_largo_5", "type": "number"},
    {"row": 74, "col": 19, "label": "Area_2_ext_largo_1", "type": "number"},
    {"row": 75, "col": 19, "label": "Area_2_ext_largo_2", "type": "number"},
    {"row": 76, "col": 19, "label": "Area_2_ext_largo_3", "type": "number"},
    {"row": 77, "col": 19, "label": "Area_2_ext_largo_4", "type": "number"},
    {"row": 78, "col": 19, "label": "Area_2_ext_largo_5", "type": "number"},

    {"row": 68, "col": 20, "label": "Area_2_iny_ancho_1", "type": "number"},
    {"row": 69, "col": 20, "label": "Area_2_iny_ancho_2", "type": "number"},
    {"row": 70, "col": 20, "label": "Area_2_iny_ancho_3", "type": "number"},
    {"row": 71, "col": 20, "label": "Area_2_iny_ancho_4", "type": "number"},
    {"row": 72, "col": 20, "label": "Area_2_iny_ancho_5", "type": "number"},
    {"row": 74, "col": 20, "label": "Area_2_ext_ancho_1", "type": "number"},
    {"row": 75, "col": 20, "label": "Area_2_ext_ancho_2", "type": "number"},
    {"row": 76, "col": 20, "label": "Area_2_ext_ancho_3", "type": "number"},
    {"row": 77, "col": 20, "label": "Area_2_ext_ancho_4", "type": "number"},
    {"row": 78, "col": 20, "label": "Area_2_ext_ancho_5", "type": "number"},

    {"row": 68, "col": 21, "label": "Area_2_iny_diam_1", "type": "number"},
    {"row": 69, "col": 21, "label": "Area_2_iny_diam_2", "type": "number"},
    {"row": 70, "col": 21, "label": "Area_2_iny_diam_3", "type": "number"},
    {"row": 71, "col": 21, "label": "Area_2_iny_diam_4", "type": "number"},
    {"row": 72, "col": 21, "label": "Area_2_iny_diam_5", "type": "number"},
    {"row": 74, "col": 21, "label": "Area_2_ext_diam_1", "type": "number"},
    {"row": 75, "col": 21, "label": "Area_2_ext_diam_2", "type": "number"},
    {"row": 76, "col": 21, "label": "Area_2_ext_diam_3", "type": "number"},
    {"row": 77, "col": 21, "label": "Area_2_ext_diam_4", "type": "number"},
    {"row": 78, "col": 21, "label": "Area_2_ext_diam_5", "type": "number"},

    {"row": 68, "col": 23, "label": "Area_2_iny_vel_1_1", "type": "number"},
    {"row": 68, "col": 24, "label": "Area_2_iny_vel_1_2", "type": "number"},
    {"row": 68, "col": 25, "label": "Area_2_iny_vel_1_3", "type": "number"},
    {"row": 68, "col": 26, "label": "Area_2_iny_vel_1_4", "type": "number"},
    {"row": 68, "col": 27, "label": "Area_2_iny_vel_1_5", "type": "number"},
    {"row": 69, "col": 23, "label": "Area_2_iny_vel_2_1", "type": "number"},
    {"row": 69, "col": 24, "label": "Area_2_iny_vel_2_2", "type": "number"},
    {"row": 69, "col": 25, "label": "Area_2_iny_vel_2_3", "type": "number"},
    {"row": 69, "col": 26, "label": "Area_2_iny_vel_2_4", "type": "number"},
    {"row": 69, "col": 27, "label": "Area_2_iny_vel_2_5", "type": "number"},
    {"row": 70, "col": 23, "label": "Area_2_iny_vel_3_1", "type": "number"},
    {"row": 70, "col": 24, "label": "Area_2_iny_vel_3_2", "type": "number"},
    {"row": 70, "col": 25, "label": "Area_2_iny_vel_3_3", "type": "number"},
    {"row": 70, "col": 26, "label": "Area_2_iny_vel_3_4", "type": "number"},
    {"row": 70, "col": 27, "label": "Area_2_iny_vel_3_5", "type": "number"},
    {"row": 71, "col": 23, "label": "Area_2_iny_vel_4_1", "type": "number"},
    {"row": 71, "col": 24, "label": "Area_2_iny_vel_4_2", "type": "number"},
    {"row": 71, "col": 25, "label": "Area_2_iny_vel_4_3", "type": "number"},
    {"row": 71, "col": 26, "label": "Area_2_iny_vel_4_4", "type": "number"},
    {"row": 71, "col": 27, "label": "Area_2_iny_vel_4_5", "type": "number"},
    {"row": 72, "col": 23, "label": "Area_2_iny_vel_5_1", "type": "number"},
    {"row": 72, "col": 24, "label": "Area_2_iny_vel_5_2", "type": "number"},
    {"row": 72, "col": 25, "label": "Area_2_iny_vel_5_3", "type": "number"},
    {"row": 72, "col": 26, "label": "Area_2_iny_vel_5_4", "type": "number"},
    {"row": 72, "col": 27, "label": "Area_2_iny_vel_5_5", "type": "number"},

    {"row": 74, "col": 23, "label": "Area_2_ext_vel_1_1", "type": "number"},
    {"row": 74, "col": 24, "label": "Area_2_ext_vel_1_2", "type": "number"},
    {"row": 74, "col": 25, "label": "Area_2_ext_vel_1_3", "type": "number"},
    {"row": 74, "col": 26, "label": "Area_2_ext_vel_1_4", "type": "number"},
    {"row": 74, "col": 27, "label": "Area_2_ext_vel_1_5", "type": "number"},
    {"row": 75, "col": 23, "label": "Area_2_ext_vel_2_1", "type": "number"},
    {"row": 75, "col": 24, "label": "Area_2_ext_vel_2_2", "type": "number"},
    {"row": 75, "col": 25, "label": "Area_2_ext_vel_2_3", "type": "number"},
    {"row": 75, "col": 26, "label": "Area_2_ext_vel_2_4", "type": "number"},
    {"row": 75, "col": 27, "label": "Area_2_ext_vel_2_5", "type": "number"},
    {"row": 76, "col": 23, "label": "Area_2_ext_vel_3_1", "type": "number"},
    {"row": 76, "col": 24, "label": "Area_2_ext_vel_3_2", "type": "number"},
    {"row": 76, "col": 25, "label": "Area_2_ext_vel_3_3", "type": "number"},
    {"row": 76, "col": 26, "label": "Area_2_ext_vel_3_4", "type": "number"},
    {"row": 76, "col": 27, "label": "Area_2_ext_vel_3_5", "type": "number"},
    {"row": 77, "col": 23, "label": "Area_2_ext_vel_4_1", "type": "number"},
    {"row": 77, "col": 24, "label": "Area_2_ext_vel_4_2", "type": "number"},
    {"row": 77, "col": 25, "label": "Area_2_ext_vel_4_3", "type": "number"},
    {"row": 77, "col": 26, "label": "Area_2_ext_vel_4_4", "type": "number"},
    {"row": 77, "col": 27, "label": "Area_2_ext_vel_4_5", "type": "number"},
    {"row": 78, "col": 23, "label": "Area_2_ext_vel_5_1", "type": "number"},
    {"row": 78, "col": 24, "label": "Area_2_ext_vel_5_2", "type": "number"},
    {"row": 78, "col": 25, "label": "Area_2_ext_vel_5_3", "type": "number"},
    {"row": 78, "col": 26, "label": "Area_2_ext_vel_5_4", "type": "number"},
    {"row": 78, "col": 27, "label": "Area_2_ext_vel_5_5", "type": "number"},

    # ===== Área 3 =====
    {"row": 79, "col": 2,  "label": "Area_3_bis", "type": "text"},
    {"row": 79, "col": 6,  "label": "Area_3_desc", "type": "text"},
    {"row": 80, "col": 15, "label": "Area_3_iny_1", "type": "text"},
    {"row": 81, "col": 15, "label": "Area_3_iny_2", "type": "text"},
    {"row": 82, "col": 15, "label": "Area_3_iny_3", "type": "text"},
    {"row": 83, "col": 15, "label": "Area_3_iny_4", "type": "text"},
    {"row": 84, "col": 15, "label": "Area_3_iny_5", "type": "text"},
    {"row": 86, "col": 15, "label": "Area_3_ext_1", "type": "text"},
    {"row": 87, "col": 15, "label": "Area_3_ext_2", "type": "text"},
    {"row": 88, "col": 15, "label": "Area_3_ext_3", "type": "text"},
    {"row": 89, "col": 15, "label": "Area_3_ext_4", "type": "text"},
    {"row": 90, "col": 15, "label": "Area_3_ext_5", "type": "text"},

    {"row": 80, "col": 19, "label": "Area_3_iny_largo_1", "type": "number"},
    {"row": 81, "col": 19, "label": "Area_3_iny_largo_2", "type": "number"},
    {"row": 82, "col": 19, "label": "Area_3_iny_largo_3", "type": "number"},
    {"row": 83, "col": 19, "label": "Area_3_iny_largo_4", "type": "number"},
    {"row": 84, "col": 19, "label": "Area_3_iny_largo_5", "type": "number"},
    {"row": 86, "col": 19, "label": "Area_3_ext_largo_1", "type": "number"},
    {"row": 87, "col": 19, "label": "Area_3_ext_largo_2", "type": "number"},
    {"row": 88, "col": 19, "label": "Area_3_ext_largo_3", "type": "number"},
    {"row": 89, "col": 19, "label": "Area_3_ext_largo_4", "type": "number"},
    {"row": 90, "col": 19, "label": "Area_3_ext_largo_5", "type": "number"},

    {"row": 80, "col": 20, "label": "Area_3_iny_ancho_1", "type": "number"},
    {"row": 81, "col": 20, "label": "Area_3_iny_ancho_2", "type": "number"},
    {"row": 82, "col": 20, "label": "Area_3_iny_ancho_3", "type": "number"},
    {"row": 83, "col": 20, "label": "Area_3_iny_ancho_4", "type": "number"},
    {"row": 84, "col": 20, "label": "Area_3_iny_ancho_5", "type": "number"},
    {"row": 86, "col": 20, "label": "Area_3_ext_ancho_1", "type": "number"},
    {"row": 87, "col": 20, "label": "Area_3_ext_ancho_2", "type": "number"},
    {"row": 88, "col": 20, "label": "Area_3_ext_ancho_3", "type": "number"},
    {"row": 89, "col": 20, "label": "Area_3_ext_ancho_4", "type": "number"},
    {"row": 90, "col": 20, "label": "Area_3_ext_ancho_5", "type": "number"},

    {"row": 80, "col": 21, "label": "Area_3_iny_diam_1", "type": "number"},
    {"row": 81, "col": 21, "label": "Area_3_iny_diam_2", "type": "number"},
    {"row": 82, "col": 21, "label": "Area_3_iny_diam_3", "type": "number"},
    {"row": 83, "col": 21, "label": "Area_3_iny_diam_4", "type": "number"},
    {"row": 84, "col": 21, "label": "Area_3_iny_diam_5", "type": "number"},
    {"row": 86, "col": 21, "label": "Area_3_ext_diam_1", "type": "number"},
    {"row": 87, "col": 21, "label": "Area_3_ext_diam_2", "type": "number"},
    {"row": 88, "col": 21, "label": "Area_3_ext_diam_3", "type": "number"},
    {"row": 89, "col": 21, "label": "Area_3_ext_diam_4", "type": "number"},
    {"row": 90, "col": 21, "label": "Area_3_ext_diam_5", "type": "number"},

    {"row": 80, "col": 23, "label": "Area_3_iny_vel_1_1", "type": "number"},
    {"row": 80, "col": 24, "label": "Area_3_iny_vel_1_2", "type": "number"},
    {"row": 80, "col": 25, "label": "Area_3_iny_vel_1_3", "type": "number"},
    {"row": 80, "col": 26, "label": "Area_3_iny_vel_1_4", "type": "number"},
    {"row": 80, "col": 27, "label": "Area_3_iny_vel_1_5", "type": "number"},
    {"row": 81, "col": 23, "label": "Area_3_iny_vel_2_1", "type": "number"},
    {"row": 81, "col": 24, "label": "Area_3_iny_vel_2_2", "type": "number"},
    {"row": 81, "col": 25, "label": "Area_3_iny_vel_2_3", "type": "number"},
    {"row": 81, "col": 26, "label": "Area_3_iny_vel_2_4", "type": "number"},
    {"row": 81, "col": 27, "label": "Area_3_iny_vel_2_5", "type": "number"},
    {"row": 82, "col": 23, "label": "Area_3_iny_vel_3_1", "type": "number"},
    {"row": 82, "col": 24, "label": "Area_3_iny_vel_3_2", "type": "number"},
    {"row": 82, "col": 25, "label": "Area_3_iny_vel_3_3", "type": "number"},
    {"row": 82, "col": 26, "label": "Area_3_iny_vel_3_4", "type": "number"},
    {"row": 82, "col": 27, "label": "Area_3_iny_vel_3_5", "type": "number"},
    {"row": 83, "col": 23, "label": "Area_3_iny_vel_4_1", "type": "number"},
    {"row": 83, "col": 24, "label": "Area_3_iny_vel_4_2", "type": "number"},
    {"row": 83, "col": 25, "label": "Area_3_iny_vel_4_3", "type": "number"},
    {"row": 83, "col": 26, "label": "Area_3_iny_vel_4_4", "type": "number"},
    {"row": 83, "col": 27, "label": "Area_3_iny_vel_4_5", "type": "number"},
    {"row": 84, "col": 23, "label": "Area_3_iny_vel_5_1", "type": "number"},
    {"row": 84, "col": 24, "label": "Area_3_iny_vel_5_2", "type": "number"},
    {"row": 84, "col": 25, "label": "Area_3_iny_vel_5_3", "type": "number"},
    {"row": 84, "col": 26, "label": "Area_3_iny_vel_5_4", "type": "number"},
    {"row": 84, "col": 27, "label": "Area_3_iny_vel_5_5", "type": "number"},

    {"row": 86, "col": 23, "label": "Area_3_ext_vel_1_1", "type": "number"},
    {"row": 86, "col": 24, "label": "Area_3_ext_vel_1_2", "type": "number"},
    {"row": 86, "col": 25, "label": "Area_3_ext_vel_1_3", "type": "number"},
    {"row": 86, "col": 26, "label": "Area_3_ext_vel_1_4", "type": "number"},
    {"row": 86, "col": 27, "label": "Area_3_ext_vel_1_5", "type": "number"},
    {"row": 87, "col": 23, "label": "Area_3_ext_vel_2_1", "type": "number"},
    {"row": 87, "col": 24, "label": "Area_3_ext_vel_2_2", "type": "number"},
    {"row": 87, "col": 25, "label": "Area_3_ext_vel_2_3", "type": "number"},
    {"row": 87, "col": 26, "label": "Area_3_ext_vel_2_4", "type": "number"},
    {"row": 87, "col": 27, "label": "Area_3_ext_vel_2_5", "type": "number"},
    {"row": 88, "col": 23, "label": "Area_3_ext_vel_3_1", "type": "number"},
    {"row": 88, "col": 24, "label": "Area_3_ext_vel_3_2", "type": "number"},
    {"row": 88, "col": 25, "label": "Area_3_ext_vel_3_3", "type": "number"},
    {"row": 88, "col": 26, "label": "Area_3_ext_vel_3_4", "type": "number"},
    {"row": 88, "col": 27, "label": "Area_3_ext_vel_3_5", "type": "number"},
    {"row": 89, "col": 23, "label": "Area_3_ext_vel_4_1", "type": "number"},
    {"row": 89, "col": 24, "label": "Area_3_ext_vel_4_2", "type": "number"},
    {"row": 89, "col": 25, "label": "Area_3_ext_vel_4_3", "type": "number"},
    {"row": 89, "col": 26, "label": "Area_3_ext_vel_4_4", "type": "number"},
    {"row": 89, "col": 27, "label": "Area_3_ext_vel_4_5", "type": "number"},
    {"row": 90, "col": 23, "label": "Area_3_ext_vel_5_1", "type": "number"},
    {"row": 90, "col": 24, "label": "Area_3_ext_vel_5_2", "type": "number"},
    {"row": 90, "col": 25, "label": "Area_3_ext_vel_5_3", "type": "number"},
    {"row": 90, "col": 26, "label": "Area_3_ext_vel_5_4", "type": "number"},
    {"row": 90, "col": 27, "label": "Area_3_ext_vel_5_5", "type": "number"},

    # ===== Área 4 =====
    {"row": 91, "col": 2,  "label": "Area_4_bis", "type": "text"},
    {"row": 91, "col": 6,  "label": "Area_4_desc", "type": "text"},
    {"row": 92, "col": 15, "label": "Area_4_iny_1", "type": "text"},
    {"row": 93, "col": 15, "label": "Area_4_iny_2", "type": "text"},
    {"row": 94, "col": 15, "label": "Area_4_iny_3", "type": "text"},
    {"row": 95, "col": 15, "label": "Area_4_iny_4", "type": "text"},
    {"row": 96, "col": 15, "label": "Area_4_iny_5", "type": "text"},
    {"row": 98, "col": 15, "label": "Area_4_ext_1", "type": "text"},
    {"row": 99, "col": 15, "label": "Area_4_ext_2", "type": "text"},
    {"row": 100,"col": 15, "label": "Area_4_ext_3", "type": "text"},
    {"row": 101,"col": 15, "label": "Area_4_ext_4", "type": "text"},
    {"row": 102,"col": 15, "label": "Area_4_ext_5", "type": "text"},

    {"row": 92, "col": 19, "label": "Area_4_iny_largo_1", "type": "number"},
    {"row": 93, "col": 19, "label": "Area_4_iny_largo_2", "type": "number"},
    {"row": 94, "col": 19, "label": "Area_4_iny_largo_3", "type": "number"},
    {"row": 95, "col": 19, "label": "Area_4_iny_largo_4", "type": "number"},
    {"row": 96, "col": 19, "label": "Area_4_iny_largo_5", "type": "number"},
    {"row": 98, "col": 19, "label": "Area_4_ext_largo_1", "type": "number"},
    {"row": 99, "col": 19, "label": "Area_4_ext_largo_2", "type": "number"},
    {"row": 100,"col": 19, "label": "Area_4_ext_largo_3", "type": "number"},
    {"row": 101,"col": 19, "label": "Area_4_ext_largo_4", "type": "number"},
    {"row": 102,"col": 19, "label": "Area_4_ext_largo_5", "type": "number"},

    {"row": 92, "col": 20, "label": "Area_4_iny_ancho_1", "type": "number"},
    {"row": 93, "col": 20, "label": "Area_4_iny_ancho_2", "type": "number"},
    {"row": 94, "col": 20, "label": "Area_4_iny_ancho_3", "type": "number"},
    {"row": 95, "col": 20, "label": "Area_4_iny_ancho_4", "type": "number"},
    {"row": 96, "col": 20, "label": "Area_4_iny_ancho_5", "type": "number"},
    {"row": 98, "col": 20, "label": "Area_4_ext_ancho_1", "type": "number"},
    {"row": 99, "col": 20, "label": "Area_4_ext_ancho_2", "type": "number"},
    {"row": 100,"col": 20, "label": "Area_4_ext_ancho_3", "type": "number"},
    {"row": 101,"col": 20, "label": "Area_4_ext_ancho_4", "type": "number"},
    {"row": 102,"col": 20, "label": "Area_4_ext_ancho_5", "type": "number"},

    {"row": 92, "col": 21, "label": "Area_4_iny_diam_1", "type": "number"},
    {"row": 93, "col": 21, "label": "Area_4_iny_diam_2", "type": "number"},
    {"row": 94, "col": 21, "label": "Area_4_iny_diam_3", "type": "number"},
    {"row": 95, "col": 21, "label": "Area_4_iny_diam_4", "type": "number"},
    {"row": 96, "col": 21, "label": "Area_4_iny_diam_5", "type": "number"},
    {"row": 98, "col": 21, "label": "Area_4_ext_diam_1", "type": "number"},
    {"row": 99, "col": 21, "label": "Area_4_ext_diam_2", "type": "number"},
    {"row": 100,"col": 21, "label": "Area_4_ext_diam_3", "type": "number"},
    {"row": 101,"col": 21, "label": "Area_4_ext_diam_4", "type": "number"},
    {"row": 102,"col": 21, "label": "Area_4_ext_diam_5", "type": "number"},

    {"row": 92, "col": 23, "label": "Area_4_iny_vel_1_1", "type": "number"},
    {"row": 92, "col": 24, "label": "Area_4_iny_vel_1_2", "type": "number"},
    {"row": 92, "col": 25, "label": "Area_4_iny_vel_1_3", "type": "number"},
    {"row": 92, "col": 26, "label": "Area_4_iny_vel_1_4", "type": "number"},
    {"row": 92, "col": 27, "label": "Area_4_iny_vel_1_5", "type": "number"},
    {"row": 93, "col": 23, "label": "Area_4_iny_vel_2_1", "type": "number"},
    {"row": 93, "col": 24, "label": "Area_4_iny_vel_2_2", "type": "number"},
    {"row": 93, "col": 25, "label": "Area_4_iny_vel_2_3", "type": "number"},
    {"row": 93, "col": 26, "label": "Area_4_iny_vel_2_4", "type": "number"},
    {"row": 93, "col": 27, "label": "Area_4_iny_vel_2_5", "type": "number"},
    {"row": 94, "col": 23, "label": "Area_4_iny_vel_3_1", "type": "number"},
    {"row": 94, "col": 24, "label": "Area_4_iny_vel_3_2", "type": "number"},
    {"row": 94, "col": 25, "label": "Area_4_iny_vel_3_3", "type": "number"},
    {"row": 94, "col": 26, "label": "Area_4_iny_vel_3_4", "type": "number"},
    {"row": 94, "col": 27, "label": "Area_4_iny_vel_3_5", "type": "number"},
    {"row": 95, "col": 23, "label": "Area_4_iny_vel_4_1", "type": "number"},
    {"row": 95, "col": 24, "label": "Area_4_iny_vel_4_2", "type": "number"},
    {"row": 95, "col": 25, "label": "Area_4_iny_vel_4_3", "type": "number"},
    {"row": 95, "col": 26, "label": "Area_4_iny_vel_4_4", "type": "number"},
    {"row": 95, "col": 27, "label": "Area_4_iny_vel_4_5", "type": "number"},
    {"row": 96, "col": 23, "label": "Area_4_iny_vel_5_1", "type": "number"},
    {"row": 96, "col": 24, "label": "Area_4_iny_vel_5_2", "type": "number"},
    {"row": 96, "col": 25, "label": "Area_4_iny_vel_5_3", "type": "number"},
    {"row": 96, "col": 26, "label": "Area_4_iny_vel_5_4", "type": "number"},
    {"row": 96, "col": 27, "label": "Area_4_iny_vel_5_5", "type": "number"},

    {"row": 98, "col": 23, "label": "Area_4_ext_vel_1_1", "type": "number"},
    {"row": 98, "col": 24, "label": "Area_4_ext_vel_1_2", "type": "number"},
    {"row": 98, "col": 25, "label": "Area_4_ext_vel_1_3", "type": "number"},
    {"row": 98, "col": 26, "label": "Area_4_ext_vel_1_4", "type": "number"},
    {"row": 98, "col": 27, "label": "Area_4_ext_vel_1_5", "type": "number"},
    {"row": 99, "col": 23, "label": "Area_4_ext_vel_2_1", "type": "number"},
    {"row": 99, "col": 24, "label": "Area_4_ext_vel_2_2", "type": "number"},
    {"row": 99, "col": 25, "label": "Area_4_ext_vel_2_3", "type": "number"},
    {"row": 99, "col": 26, "label": "Area_4_ext_vel_2_4", "type": "number"},
    {"row": 99, "col": 27, "label": "Area_4_ext_vel_2_5", "type": "number"},
    {"row": 100,"col": 23, "label": "Area_4_ext_vel_3_1", "type": "number"},
    {"row": 100,"col": 24, "label": "Area_4_ext_vel_3_2", "type": "number"},
    {"row": 100,"col": 25, "label": "Area_4_ext_vel_3_3", "type": "number"},
    {"row": 100,"col": 26, "label": "Area_4_ext_vel_3_4", "type": "number"},
    {"row": 100,"col": 27, "label": "Area_4_ext_vel_3_5", "type": "number"},
    {"row": 101,"col": 23, "label": "Area_4_ext_vel_4_1", "type": "number"},
    {"row": 101,"col": 24, "label": "Area_4_ext_vel_4_2", "type": "number"},
    {"row": 101,"col": 25, "label": "Area_4_ext_vel_4_3", "type": "number"},
    {"row": 101,"col": 26, "label": "Area_4_ext_vel_4_4", "type": "number"},
    {"row": 101,"col": 27, "label": "Area_4_ext_vel_4_5", "type": "number"},
    {"row": 102,"col": 23, "label": "Area_4_ext_vel_5_1", "type": "number"},
    {"row": 102,"col": 24, "label": "Area_4_ext_vel_5_2", "type": "number"},
    {"row": 102,"col": 25, "label": "Area_4_ext_vel_5_3", "type": "number"},
    {"row": 102,"col": 26, "label": "Area_4_ext_vel_5_4", "type": "number"},
    {"row": 102,"col": 27, "label": "Area_4_ext_vel_5_5", "type": "number"},
]


DERIVED_FIELDS = [
    "Volumen_1", "m3xpersona_1", "Resultado_1",
    "Volumen_2", "m3xpersona_2", "Resultado_2",
    "Volumen_3", "m3xpersona_3", "Resultado_3",
    "Volumen_4", "m3xpersona_4", "Resultado_4",
    "Estado",
]

def _is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False

def _to_number(v):
    if v is None:
        return float("nan")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s == "":
        return float("nan")
    s1 = s.replace(" ", "")
    # 1.234,56 -> 1234.56
    if re.match(r"^-?\d{1,3}(\.\d{3})*,\d+$", s1):
        s1 = s1.replace(".", "").replace(",", ".")
        try: return float(s1)
        except: return float("nan")
    # 1,234.56 -> 1234.56
    if re.match(r"^-?\d{1,3}(,\d{3})*\.\d+$", s1):
        s1 = s1.replace(",", "")
        try: return float(s1)
        except: return float("nan")
    # 1.234.567 o 1,234,567 -> 1234567
    if re.match(r"^-?\d{1,3}([.,]\d{3})+$", s1):
        s2 = s1.replace(",", "").replace(".", "")
        try: return float(s2)
        except: return float("nan")
    # genérico: coma como decimal
    s2 = s1.replace(",", ".")
    try: return float(s2)
    except: return float("nan")

def _fmt_number(val_float):
    """Devuelve string sin ceros a la derecha ni punto sobrante."""
    if val_float is None or (isinstance(val_float, float) and math.isnan(val_float)):
        return ""
    s = f"{float(val_float):.10f}".rstrip("0").rstrip(".")
    return s if s != "-0" else "0"

def _to_number_str(v):
    """Parsea a float y devuelve string normalizado (mantiene 3.128, 2.98, etc.)."""
    f = _to_number(v)
    return _fmt_number(f)

def _to_int(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int,)):
        return str(int(v))
    if isinstance(v, float):
        return str(int(round(v)))
    s = str(v).strip()
    if s == "":
        return ""
    # quitar palabras/espacios/separadores y dejar números y signo
    s = s.replace(" ", "")
    s = s.replace(".", "").replace(",", "")
    m = re.match(r"^(-?\d+)", s)
    if not m:
        f = _to_number(v)
        return "" if (isinstance(f, float) and math.isnan(f)) else str(int(round(f)))
    return str(int(m.group(1)))

def _to_time(v):
    """Devuelve HH:MM:SS. Acepta datetime, '14:30', '14:30:00', '15:00 HORAS', etc."""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.time)):
        t = v.time() if isinstance(v, dt.datetime) else v
        return t.strftime("%H:%M:%S")
    s = str(v).strip().upper()
    # limpiar palabras comunes (HRS, HORAS, HORA)
    s = re.sub(r"\b(HRS?|HORAS?)\b", "", s)
    s = s.replace(".", ":")
    s = re.sub(r"\s+", "", s)
    # patrones HH:MM o HH:MM:SS
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
    if m:
        h = int(m.group(1)); mnt = int(m.group(2)); sec = int(m.group(3) or 0)
        h = max(0, min(23, h)); mnt = max(0, min(59, mnt)); sec = max(0, min(59, sec))
        return f"{h:02d}:{mnt:02d}:{sec:02d}"
    # si viene 4 dígitos (1430) -> 14:30:00
    m = re.match(r"^(\d{1,2})(\d{2})$", s)
    if m:
        h = int(m.group(1)); mnt = int(m.group(2))
        h = max(0, min(23, h)); mnt = max(0, min(59, mnt))
        return f"{h:02d}:{mnt:02d}:00"
    return ""  # no reconocible

def _cast(v, t):
    if t == "integer":
        return _to_int(v)                  # -> "2", "1", etc.
    if t == "time":
        return _to_time(v)                 # -> "HH:MM:SS"
    if t == "number":
        return _to_number_str(v)           # -> "3.128", "2.98", "5.67"
    # text (incluye fecha V1 como texto plano)
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    return "" if v is None else str(v).strip()

def _fmt_number(val_float):
    """Devuelve string sin ceros a la derecha ni punto sobrante."""
    if val_float is None or (isinstance(val_float, float) and math.isnan(val_float)):
        return ""
    s = f"{float(val_float):.10f}".rstrip("0").rstrip(".")
    return s if s != "-0" else "0"

def _calc_area_metrics(largo, ancho, alto, n_personas):
    """
    largo/ancho/alto/n_personas: ya parseados con _to_number (float) o NaN.
    Retorna (volumen_str, m3xpersona_str, resultado_str) con:
      - volumen = largo * ancho * alto
      - m3xpersona = volumen / n_personas  (si n_personas <= 0 o NaN -> "")
      - resultado = "Cumple" si m3xpersona > 10, si no "No cumple"; si no hay datos -> ""
    """
    # Validar dimensiones
    if any([(isinstance(x, float) and math.isnan(x)) for x in (largo, ancho, alto)]):
        return ("", "", "")
    vol = float(largo) * float(ancho) * float(alto)
    vol_str = _fmt_number(vol)

    # Validar personas
    if n_personas is None or (isinstance(n_personas, float) and math.isnan(n_personas)) or n_personas <= 0:
        return (vol_str, "", "")

    m3pp = vol / float(n_personas)
    m3pp_str = _fmt_number(m3pp)

    # Umbral estricto: > 10 => Cumple; en otro caso => No cumple
    resultado = "Cumple" if m3pp > 10 else "No cumple"
    return (vol_str, m3pp_str, resultado)



def pick_best_sheet(wb):
    best_sheet = None
    best_hits = -1
    for sname in wb.sheetnames:
        ws = wb[sname]
        hits = 0
        for f in SCHEMA:
            v = ws.cell(row=f["row"], column=f["col"]).value
            if not _is_empty(v):
                hits += 1
        if hits > best_hits:
            best_hits = hits
            best_sheet = sname
    return best_sheet, best_hits

def process_file(path: Path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False, keep_links=False)
    except Exception as e:
        return {"__error__": f"No se pudo abrir: {e}", "__source_file__": str(path)}

    sheet, hits = pick_best_sheet(wb)
    if sheet is None:
        return {"__error__": "No se detectó una hoja válida", "__source_file__": str(path)}

    ws = wb[sheet]
    row_dict = {"__source_file__": str(path), "__sheet__": sheet, "__hits__": hits}

    # 1) Guardar valores base según SCHEMA (con casteo)
    for f in SCHEMA:
        raw = ws.cell(row=f["row"], column=f["col"]).value
        row_dict[f["label"]] = _cast(raw, f["type"])

    # 2) Calcular métricas por área (usando los RAW parseados a número)
    # Coordenadas por área (filas base)
    area_rows = {1:31, 2:35, 3:39, 4:43}
    # Columnas (E=5, G=7, I=9, K=11)
    COL_E, COL_G, COL_I, COL_K = 5, 7, 9, 11

    resultados_por_area = []
    for i, r in area_rows.items():
        # Raw values (sin casteo a string) para cálculos
        raw_np = ws.cell(row=r, column=COL_E).value
        raw_l  = ws.cell(row=r, column=COL_G).value
        raw_an = ws.cell(row=r, column=COL_I).value
        raw_al = ws.cell(row=r, column=COL_K).value

        np_num = _to_number(raw_np)
        l_num  = _to_number(raw_l)
        an_num = _to_number(raw_an)
        al_num = _to_number(raw_al)

        vol_str, m3pp_str, res_str = _calc_area_metrics(l_num, an_num, al_num, np_num)

        row_dict[f"Volumen_{i}"]     = vol_str
        row_dict[f"m3xpersona_{i}"]  = m3pp_str
        row_dict[f"Resultado_{i}"]   = res_str
        resultados_por_area.append(res_str)

    # 3) Estado global
    # - Si alguna área es "No cumple" -> Pendiente
    # - Si al menos una es "Cumple" y ninguna "No cumple" -> Cerrado
    # - Si todas vacías -> ""
    if any(r == "No cumple" for r in resultados_por_area):
        row_dict["Estado"] = "Pendiente"
    else:
        hay_cumple = any(r == "Cumple" for r in resultados_por_area)
        row_dict["Estado"] = "Cerrado" if hay_cumple else ""

    return row_dict

def discover_files(input_dir: Path):
    return sorted([p for p in input_dir.rglob("*") if p.suffix.lower() in [".xlsx", ".xlsm"]])

def write_csv(rows, out_csv: Path):
    meta = ["__source_file__", "__sheet__", "__hits__"]
    fields = [f["label"] for f in SCHEMA]
    headers = meta + fields + DERIVED_FIELDS   # <-- agrega derivados
    with out_csv.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for r in rows:
            if "__error__" in r:
                base = {"__source_file__": r.get("__source_file__", ""), "__sheet__": "", "__hits__": -1}
                for k in (fields + DERIVED_FIELDS): base[k] = ""
                w.writerow(base)
            else:
                w.writerow({h: r.get(h, "") for h in headers})

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def _to_excel_number(s):
    """Convierte strings numéricos a float; deja vacío como ''."""
    if s is None:
        return ""
    if isinstance(s, (int, float)):
        return s
    s = str(s).strip()
    if s == "":
        return ""
    try:
        return float(s)
    except:
        return s  # si no es número, dejar como texto

def write_xlsx(rows, out_xlsx: Path, schema, derived_fields):
    """
    rows: lista de dicts (cada archivo = una fila)
    out_xlsx: ruta .xlsx
    schema: SCHEMA original para conocer tipos
    derived_fields: lista DERIVED_FIELDS
    """
    # columnas
    meta = ["__source_file__", "__sheet__", "__hits__"]
    base_fields = [f["label"] for f in schema]
    headers = meta + base_fields + derived_fields

    # mapa de tipos de SCHEMA para casteo selectivo
    type_by_label = {f["label"]: f["type"] for f in schema}

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    # escribe headers
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)

    # escribe filas
    r_out = 2
    for r in rows:
        # aseguremos todas las claves
        row_vals = {h: r.get(h, "") for h in headers}

        for j, h in enumerate(headers, start=1):
            v = row_vals[h]

            # meta: dejar como texto/num natural
            if h in ["__hits__"]:
                # hits siempre numérico si viene
                v2 = _to_excel_number(v)
                ws.cell(row=r_out, column=j, value=v2 if v2 != "" else None)
                continue
            if h in ["__source_file__", "__sheet__"]:
                ws.cell(row=r_out, column=j, value=None if v == "" else v)
                continue

            # campos del schema con tipo
            if h in type_by_label:
                t = type_by_label[h]
                if t == "integer":
                    # forzamos a int si se puede
                    try:
                        v2 = int(float(str(v).replace(",", ".")))
                        ws.cell(row=r_out, column=j, value=v2)
                    except:
                        ws.cell(row=r_out, column=j, value=None if v == "" else v)
                elif t == "number":
                    v2 = _to_excel_number(v)
                    ws.cell(row=r_out, column=j, value=v2 if v2 != "" else None)
                elif t == "time":
                    # dejamos como texto HH:MM:SS
                    ws.cell(row=r_out, column=j, value=None if v == "" else v)
                else:
                    # text / otros
                    ws.cell(row=r_out, column=j, value=None if v == "" else v)
            else:
                # derivados: Volumen_i y m3xpersona_i como número si se puede; Resultado_i/Estado texto
                if h.startswith("Volumen_") or h.startswith("m3xpersona_"):
                    v2 = _to_excel_number(v)
                    ws.cell(row=r_out, column=j, value=v2 if v2 != "" else None)
                else:
                    ws.cell(row=r_out, column=j, value=None if v == "" else v)

        r_out += 1

    # ancho cómodo de columnas
    for j, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(50, len(h) + 2))

    wb.save(out_xlsx)



def main():
    # === RUTAS EN DURO (EDITA AQUÍ) ===
    input_dir = Path(r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba")
    out_xlsx  = Path(r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\resultados_unificados.xlsx")
    # ===================================

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Carpeta no encontrada: {input_dir}")
        sys.exit(1)

    files = discover_files(input_dir)
    if not files:
        print("No se encontraron archivos .xlsx/.xlsm")
        sys.exit(1)

    rows = []
    total = len(files)
    for idx, p in enumerate(files, start=1):
        print(f"[{idx}/{total}] Procesando: {p}")
        res = process_file(p)
        if "__error__" in res:
            print(f"    → ERROR: {res['__error__']}")
        else:
            print(f"    → OK: hoja='{res['__sheet__']}', hits={res['__hits__']}")
        rows.append(res)


    # ⬇️ Exporta a Excel en vez de CSV
    write_xlsx(rows, out_xlsx, SCHEMA, DERIVED_FIELDS)

    print(f"Listo. Archivos procesados: {len(files)}")
    print(f"Excel unificado: {out_xlsx}")

if __name__ == "__main__":
    main()
