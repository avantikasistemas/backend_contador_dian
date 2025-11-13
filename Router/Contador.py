from fastapi import APIRouter, Request, Depends
from sqlalchemy.orm import Session
from Class.Contador import Contador
from Utils.decorator import http_decorator
from Config.db import get_db

contador_router = APIRouter()

@contador_router.post('/contador/procesar-archivo', tags=["Contador"], response_model=dict)
@http_decorator
def procesar_archivo_excel(request: Request, db: Session = Depends(get_db)):
    """
    Procesa un archivo Excel de la DIAN aplicando filtros y cálculos.
    
    Pasos:
    1. Filtra por tipo de documento (Factura electrónica, Factura electrónica de contingencia, Nota de crédito electrónica)
    2. Filtra por NIT Emisor = 890101977
    3. Agrega columna Tipo-Folio con fórmula basada en Prefijo y Folio
    4. Agrega columna Subtotal = Total - IVA
    5. Agrega columna Naturaleza de la operación basada en Prefijo y Subtotal
    """
    data = getattr(request.state, "json_data", {})
    response = Contador(db).procesar_archivo_excel(data)
    return response

@contador_router.post('/contador/procesar-archivo-dms', tags=["Contador"], response_model=dict)
@http_decorator
def procesar_archivo_dms(request: Request, db: Session = Depends(get_db)):
    """
    Procesa un archivo DMS aplicando cálculos.
    
    Pasos:
    1. Agrega columna tipo_doc_desc_tipo = Tipo Docto. + " " + Número Docto.
    2. Agrega columna Saldo2 con fórmula basada en Tipo Docto. y Saldo Periodo
    """
    data = getattr(request.state, "json_data", {})
    response = Contador(db).procesar_archivo_dms(data)
    return response
