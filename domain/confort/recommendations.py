"""Business logic for comfort recommendations."""
from typing import Dict, List


def generar_recomendaciones(
    pmv: float,
    tdb_initial: float,
    tr_initial: float,
    vr: float,
    rh: float,
    met: float,
    clo: float,
) -> List[Dict]:
    recomendaciones: List[Dict] = []
    dif_temp = tr_initial - tdb_initial
    estrategias_base = []
    if pmv > 1.0:
        estrategias_base = [
            {
                "condicion": True,
                "tipo": "enfriamiento",
                "mensaje": "Refrigeración activa requerida",
                "acciones": [
                    "- Implementar [__] equipos enfriadores de aire (según las dimensiones de las áreas), con el fin de enfriar el aire.",
                ],
                "plazo": "[__] meses desde la recepción del presente informe técnico",
            },
            {
                "condicion": dif_temp > 2.0,
                "tipo": "aislamiento",
                "mensaje": "Reducción de carga térmica",
                "acciones": [
                    "- (REVISAR SI CORRESPONDE) Instalar materiales aislantes en techos/paredes",
                    "- (REVISAR SI CORRESPONDE) Implementar protecciones solares reflectivas",
                    "- (REVISAR SI CORRESPONDE) Aislar fuentes de calor radiante",
                ],
                "plazo": "[__] meses desde la recepción del presente informe técnico",
            },
            {
                "condicion": vr < 0.2,
                "tipo": "ventilacion",
                "mensaje": "Aumentar ventilación",
                "acciones": [
                    "- Implementar [__] ventiladores industriales, con el fin de generar corrientes de aire, las cuales ayudarán a mejorar condiciones de confort térmico en dicha área.",
                    "- (REVISAR SI CORRESPONDE) Implementar sistemas de extracción forzada",
                    "- (REVISAR SI CORRESPONDE) Optimizar ventilación cruzada",
                ],
                "plazo": "[__] meses desde la recepción del presente informe técnico",
            },
        ]
    for estrategia in estrategias_base:
        if estrategia["condicion"]:
            recomendaciones.append(
                {
                    "tipo": estrategia["tipo"],
                    "categoria": estrategia["mensaje"],
                    "nivel": "Prioridad 1",
                    "mensaje": estrategia["mensaje"],
                    "acciones": estrategia.get("acciones", []),
                    "plazo": estrategia.get("plazo", "Inmediato"),
                }
            )
    return recomendaciones


__all__ = ["generar_recomendaciones"]
