from Utils.tools import Tools, CustomException
from sqlalchemy import text, func, select, and_
from sqlalchemy.exc import IntegrityError
from datetime import date, datetime
from collections import defaultdict
from typing import List, Dict, Any
import json

class Querys:

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.query_params = dict()

    # Query para desactivar registros anteriores por tipo
    def desactivar_registros_anteriores(self, tipo: int):
        """
        Actualiza el estado a 0 de todos los registros del mismo tipo.
        
        Args:
            tipo (int): 1=DIAN, 2=DMS
        """
        try:
            query = text("""
                UPDATE dbo.intranet_contabilidad_datos_depuracion
                SET estado = 0
                WHERE tipo = :tipo AND estado = 1
            """)
            
            self.db.execute(query, {"tipo": tipo})
            self.db.commit()
            
            return True
        except Exception as e:
            self.db.rollback()
            raise CustomException(f"Error al desactivar registros anteriores: {str(e)}")

    # Query para guardar datos procesados
    def guardar_datos_procesados(self, tipo: int, datos: dict):
        """
        Guarda los datos procesados en la base de datos.
        
        Args:
            tipo (int): 1=DIAN, 2=DMS
            datos (dict): Diccionario con los datos procesados
        """
        try:
            # Convertir el diccionario a JSON string
            datos_json = json.dumps(datos, ensure_ascii=False, default=str)
            
            query = text("""
                INSERT INTO dbo.intranet_contabilidad_datos_depuracion 
                (tipo, datos)
                VALUES (:tipo, :datos)
            """)
            
            self.db.execute(query, {"tipo": tipo, "datos": datos_json})
            self.db.commit()
            
            return True
        except Exception as e:
            self.db.rollback()
            raise CustomException(f"Error al guardar datos procesados: {str(e)}")

    # Query para obtener últimos datos procesados activos
    def obtener_ultimos_datos_procesados(self):
        """
        Obtiene los últimos registros activos de tipo 1 (DIAN) y tipo 2 (DMS).
        
        Returns:
            dict: {"dian": {...}, "dms": {...}}
        """
        try:
            # Obtener último registro DIAN (tipo=1) activo
            query_dian = text("""
                SELECT TOP 1 id, tipo, datos, fecha_creacion
                FROM dbo.intranet_contabilidad_datos_depuracion
                WHERE tipo = 1 AND estado = 1
                ORDER BY fecha_creacion DESC
            """)
            
            resultado_dian = self.db.execute(query_dian).fetchone()
            
            # Obtener último registro DMS (tipo=2) activo
            query_dms = text("""
                SELECT TOP 1 id, tipo, datos, fecha_creacion
                FROM dbo.intranet_contabilidad_datos_depuracion
                WHERE tipo = 2 AND estado = 1
                ORDER BY fecha_creacion DESC
            """)
            
            resultado_dms = self.db.execute(query_dms).fetchone()
            
            datos = {
                "dian": None,
                "dms": None
            }
            
            if resultado_dian:
                datos["dian"] = {
                    "id": resultado_dian[0],
                    "tipo": resultado_dian[1],
                    "datos": json.loads(resultado_dian[2]),
                    "fecha_creacion": resultado_dian[3].isoformat() if resultado_dian[3] else None
                }
            
            if resultado_dms:
                datos["dms"] = {
                    "id": resultado_dms[0],
                    "tipo": resultado_dms[1],
                    "datos": json.loads(resultado_dms[2]),
                    "fecha_creacion": resultado_dms[3].isoformat() if resultado_dms[3] else None
                }
            
            return datos
            
        except Exception as e:
            raise CustomException(f"Error al obtener últimos datos procesados: {str(e)}")
