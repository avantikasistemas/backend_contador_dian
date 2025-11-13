from fastapi import APIRouter, Request, Depends
from sqlalchemy.orm import Session
from Class.Graph import Graph
from Utils.decorator import http_decorator
from Config.db import get_db

graph_router = APIRouter()

@graph_router.post('/graph/enviar-correo', tags=["Graph"], response_model=dict)
@http_decorator
def enviar_correo_reporte(request: Request, db: Session = Depends(get_db)):
    """
    Envía correo con el reporte de facturación electrónica.
    Consulta los últimos datos procesados de DIAN y DMS,
    genera tablas HTML agrupadas por tipo de documento
    """
    response = Graph(db).enviar_correo_reporte()
    return response
