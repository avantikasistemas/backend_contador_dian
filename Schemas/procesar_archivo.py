from pydantic import BaseModel, Field
from typing import Optional

class ProcesarArchivoSchema(BaseModel):
    archivo: str = Field(..., description="Archivo Excel codificado en Base64")
    nombre_archivo: str = Field(..., description="Nombre del archivo Excel")
    tipo_archivo: str = Field(..., description="Tipo de archivo: 'ventas', 'compras', etc.")

    class Config:
        json_schema_extra = {
            "example": {
                "archivo": "base64_string_aqui...",
                "nombre_archivo": "archivo_ventas.xlsx",
                "tipo_archivo": "ventas"
            }
        }
