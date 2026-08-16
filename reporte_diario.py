# ======================================================
# 1. CONFIGURACIÓN
# ======================================================
from datetime import datetime
from pymongo import MongoClient
import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
from PIL import Image
import numpy as np
import smtplib
from email.message import EmailMessage
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font
from PIL import Image
import matplotlib.pyplot as plt
from PIL import Image
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta, timezone
import shutil
import json

def add_logo(ax, logo, x, y, zoom=0.22):
    """Añade el logo al eje solo si logo no es None."""
    if logo is None:
        return
    imagebox = OffsetImage(logo, zoom=zoom)
    ab = AnnotationBbox(
        imagebox,
        (x, y),
        frameon=False,
        xycoords="axes fraction"
    )
    ax.add_artist(ab)


def _abonos_por_dia_guatemala(col, entity_id: str, desde_iso: str, hasta_iso: str) -> dict:
    """
    Devuelve dict { 'YYYY-MM-DD': (cantidad, monto_q, clientes_unicos) } agrupando por día calendario en America/Guatemala.
    `date` en payment debe ser parseable por $dateFromString (ISO con Z u offset).
    """
    pipeline = [
        {"$match": {
            "financial_entity_id": entity_id,
            "date": {"$gte": desde_iso, "$lte": hasta_iso},
        }},
        {"$set": {
            "dt": {"$dateFromString": {"dateString": "$date", "onError": None, "onNull": None}}
        }},
        {"$match": {"dt": {"$ne": None}}},
        {"$set": {
            "fecha_gt": {
                "$dateToString": {
                    "format": "%Y-%m-%d",
                    "date": "$dt",
                    "timezone": "America/Guatemala",
                }
            }
        }},
        {"$group": {
            "_id": "$fecha_gt",
            "cantidad_abonos": {"$sum": 1},
            "monto_abonos": {"$sum": "$amount"},
            "clientes": {"$addToSet": "$user_id"},
        }},
        {"$set": {"clientes_unicos": {"$size": {"$ifNull": ["$clientes", []]}}}},
    ]
    out = {}
    try:
        for doc in col.aggregate(pipeline, allowDiskUse=True):
            k = doc.get("_id")
            if not k:
                continue
            out[str(k)] = (
                int(doc.get("cantidad_abonos") or 0),
                float(doc.get("monto_abonos") or 0) / 100.0,
                int(doc.get("clientes_unicos") or 0),
            )
    except Exception as e:
        print("Aviso: agregación de abonos por día falló:", e)
    return out


def _clientes_abonados_unicos_periodo(col, entity_id: str, desde_iso: str, hasta_iso: str) -> int:
    """Cuenta clientes únicos (distinct user_id) con pagos en el período."""
    pipeline = [
        {"$match": {
            "financial_entity_id": entity_id,
            "date": {"$gte": desde_iso, "$lte": hasta_iso},
        }},
        {"$group": {"_id": "$user_id"}},
        {"$count": "n"},
    ]
    try:
        res = list(col.aggregate(pipeline, allowDiskUse=True))
        return int(res[0]["n"]) if res else 0
    except Exception as e:
        print("Aviso: conteo de clientes abonados (período) falló:", e)
        return 0


def _anulados_por_dia_guatemala(col, entity_id: str, desde_iso: str, hasta_iso: str) -> dict:
    """
    Devuelve dict { 'YYYY-MM-DD': (cantidad, monto_q) } de créditos anulados (deleted_loan)
    agrupando por día calendario en America/Guatemala usando el campo `deleted_at`.
    """
    pipeline = [
        {"$match": {"financial_entity_id": entity_id}},
        {"$set": {
            "dt": {"$convert": {"input": "$deleted_at", "to": "date", "onError": None, "onNull": None}},
        }},
        {"$match": {"dt": {"$ne": None}}},
        {"$match": {
            "$expr": {
                "$and": [
                    {"$gte": ["$dt", {"$dateFromString": {"dateString": desde_iso, "onError": None, "onNull": None}}]},
                    {"$lte": ["$dt", {"$dateFromString": {"dateString": hasta_iso, "onError": None, "onNull": None}}]},
                ]
            }
        }},
        {"$set": {
            "fecha_gt": {
                "$dateToString": {
                    "format": "%Y-%m-%d",
                    "date": "$dt",
                    "timezone": "America/Guatemala",
                }
            }
        }},
        {"$group": {
            "_id": "$fecha_gt",
            "cantidad_anulados": {"$sum": 1},
            "monto_anulados": {"$sum": "$amount"},
        }},
    ]
    out = {}
    try:
        for doc in col.aggregate(pipeline, allowDiskUse=True):
            k = doc.get("_id")
            if not k:
                continue
            out[str(k)] = (
                int(doc.get("cantidad_anulados") or 0),
                float(doc.get("monto_anulados") or 0) / 100.0,
            )
    except Exception as e:
        print("Aviso: agregación de anulados por día falló:", e)
    return out


def _iso_rango_desde_fecha(fecha_str: str) -> tuple[str, str]:
    """Devuelve (inicio_iso, fin_iso) UTC para un día calendario en Guatemala."""
    d = datetime.strptime(fecha_str, "%Y-%m-%d")
    dia_gt = datetime(d.year, d.month, d.day, tzinfo=GUATE_TZ)
    inicio_gt = dia_gt.replace(hour=0, minute=0, second=0, microsecond=0)
    fin_gt = inicio_gt + timedelta(days=1) - timedelta(seconds=1)
    return (
        inicio_gt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        fin_gt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )


def calcular_metricas_dia(
    fecha_str: str,
    lead_col,
    loans_col,
    users_col,
    abonos_por_dia: dict,
    anulados_por_dia: dict,
) -> dict:
    """Métricas de un día desde Mongo (misma lógica que el informe diario)."""
    inicio_d, fin_d = _iso_rango_desde_fecha(fecha_str)

    prospectos_nuevos = lead_col.count_documents({
        "ecosystem": ECOSYSTEM_ID,
        "created_at": {"$gte": inicio_d, "$lte": fin_d},
    })
    clientes_nuevos = users_col.count_documents({
        "ecosystem": ECOSYSTEM_ID,
        "created_at": {"$gte": inicio_d, "$lte": fin_d},
    })

    pipeline_loans = [
        {"$match": {
            "financial_entity_id": ECOSYSTEM_ID,
            "assignment_date": {"$gte": inicio_d, "$lte": fin_d},
        }},
        {"$group": {
            "_id": None,
            "creditos_nuevos": {"$sum": 1},
            "desembolsos": {"$sum": {"$cond": [{"$eq": ["$disbursed", True]}, 1, 0]}},
            "no_desembolsos": {"$sum": {"$cond": [{"$eq": ["$disbursed", False]}, 1, 0]}},
            "monto_desembolsado": {"$sum": {"$cond": [{"$eq": ["$disbursed", True]}, "$amount", 0]}},
        }},
    ]
    resultado_loans = list(loans_col.aggregate(pipeline_loans))
    data_loans = resultado_loans[0] if resultado_loans else {
        "creditos_nuevos": 0, "desembolsos": 0, "no_desembolsos": 0, "monto_desembolsado": 0,
    }

    creditos_nuevos = data_loans["creditos_nuevos"]
    desembolsos = data_loans["desembolsos"]
    no_desembolsos = data_loans["no_desembolsos"]
    monto_desembolsado = data_loans["monto_desembolsado"] / 100
    ticket_promedio = round(monto_desembolsado / desembolsos, 2) if desembolsos else 0
    total_solicitudes = prospectos_nuevos + clientes_nuevos

    pipeline_rechazos = [
        {"$match": {
            "ecosystem": ECOSYSTEM_ID,
            "created_at": {"$gte": inicio_d, "$lte": fin_d},
        }},
        {"$unwind": "$additional_data.variables_data"},
        {"$match": {"additional_data.variables_data.key": {"$in": ["motivos_negacion", "motivoNegacion"]}}},
        {"$group": {"_id": "$additional_data.variables_data.value", "cantidad": {"$sum": 1}}},
    ]
    rechazo_politicas = 0
    for r in lead_col.aggregate(pipeline_rechazos):
        motivo = (r["_id"] or "").lower()
        if ("no cumple" in motivo or "identificación" in motivo or "nis" in motivo or "lista negra" in motivo):
            rechazo_politicas += r["cantidad"]

    rechazos_bd = prospectos_nuevos - rechazo_politicas
    cab, mab, clab = abonos_por_dia.get(fecha_str, (0, 0.0, 0))
    can, _man = anulados_por_dia.get(fecha_str, (0, 0.0))

    return {
        "Fecha": fecha_str,
        "Solicitudes": total_solicitudes,
        "Rechazos - BD": rechazos_bd,
        "Clientes nuevos": clientes_nuevos,
        "Rechazos - Políticas": rechazo_politicas,
        "Créditos nuevos": creditos_nuevos,
        "Desembolsos": desembolsos,
        "No desembolsos": no_desembolsos,
        "Monto desembolsado": monto_desembolsado,
        "Ticket promedio": ticket_promedio,
        "# de abonos recibidos": cab,
        "# de clientes abonados": clab,
        "Monto abonos": mab,
        "Créditos anulados": can,
    }


# ---- MONGO ----
MONGO_URI = "mongodb+srv://camilo:XDwsOiclZ7AvO6G1@fintechmiddlewarecluste.vvbho.mongodb.net/middleware"
DB_NAME = "middleware"

LEAD_COLLECTION = "lead"
LOANS_COLLECTION = "loan"
USERS_COLLECTION = "user"
PAYMENT_COLLECTION = "payment"
DELETED_LOAN_COLLECTION = "deleted_loan"

ECOSYSTEM_ID = "eab85a60-8565-4666-9a83-428a48d0c7dc"

# Evita cuelgues cuando launchd corre sin red estable (Mongo/SMTP)
MONGO_TIMEOUT_MS = 60_000
SMTP_TIMEOUT_SEC = 120

# ======================================================
# 2. CONFIGURACIÓN DE FECHAS (GUATEMALA UTC-6)
# ======================================================

GUATE_TZ = timezone(timedelta(hours=-6))

# Para anulados: empezar desde esta fecha (Guatemala)
FECHA_INICIO_ANULADOS = datetime(2026, 3, 29)

# Permite ejecutar sin editar el archivo (útil para backfills/automatización).
# - REPORTE_FECHA=YYYY-MM-DD fuerza la fecha del informe
# - REPORTE_EMAILS="a@x.com,b@y.com" sobreescribe destinatarios
REPORTE_FECHA_ENV = os.environ.get("REPORTE_FECHA")
REPORTE_EMAILS_ENV = os.environ.get("REPORTE_EMAILS")

# =============================================================================
# PARÉNTESIS — Recuperar un informe de un día que no se ejecutó
# =============================================================================
# Normal: FECHA_INFORME_MANUAL = None y DIAS_ATRAS = 0 (informe del día de hoy).
#
# Si olvidaste ejecutar ayer:
#   1) Pon DIAS_ATRAS = 1  (informe del día anterior en hora Guatemala).
#   2) Ejecuta el script UNA vez.
#   3) Vuelve a poner DIAS_ATRAS = 0.
#
# Si ya pasaron varios días y necesitas una fecha exacta (ej. 2026-03-27):
#   FECHA_INFORME_MANUAL = "2026-03-27"  y DIAS_ATRAS = 0
#   (tiene prioridad sobre DIAS_ATRAS). Tras ejecutar, vuelve a None.
FECHA_INFORME_MANUAL = None  # ej. "2026-03-27" o None

# Cuántos días atrás respecto a “hoy” Guatemala (0 = hoy, 1 = ayer)
DIAS_ATRAS = 0

# Fecha base Guatemala (día del informe)
if REPORTE_FECHA_ENV and not FECHA_INFORME_MANUAL:
    FECHA_INFORME_MANUAL = REPORTE_FECHA_ENV.strip()

if FECHA_INFORME_MANUAL:
    _d = datetime.strptime(FECHA_INFORME_MANUAL.strip(), "%Y-%m-%d").date()
    hoy = datetime(_d.year, _d.month, _d.day, tzinfo=GUATE_TZ)
else:
    hoy = datetime.now(GUATE_TZ) - timedelta(days=DIAS_ATRAS)

hoy_ejecucion = datetime.now() - timedelta(hours=1)


hora_actual = hoy.strftime("%H%M%S")
FECHA_CONSULTA = hoy.strftime("%Y-%m-%d")
#FECHA_CONSULTA= (hoy - timedelta(days=7)).strftime("%Y-%m-%d")
FECHA_LANZAMIENTO = datetime(2026, 3, 2)

if FECHA_INFORME_MANUAL or DIAS_ATRAS:
    print(
        f"[PARÉNTESIS] Informe para el día {FECHA_CONSULTA} (Guatemala). "
        "Tras esta ejecución, pon FECHA_INFORME_MANUAL=None y DIAS_ATRAS=0."
    )

# Inicio del día Guatemala
inicio_dia_gt = hoy.replace(hour=0, minute=0, second=0, microsecond=0)

# Fin del día Guatemala
fin_dia_gt = inicio_dia_gt + timedelta(days=1) - timedelta(seconds=1)

# Convertir a UTC
inicio_dia_utc = inicio_dia_gt.astimezone(timezone.utc)
fin_dia_utc = fin_dia_gt.astimezone(timezone.utc)

# Formato ISO Z
inicio_iso = inicio_dia_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
fin_iso = fin_dia_utc.strftime("%Y-%m-%dT%H:%M:%SZ")

# Rango para ACUMULADO (desde FECHA_LANZAMIENTO hasta fin del día de consulta)
inicio_acum_gt = FECHA_LANZAMIENTO.replace(tzinfo=GUATE_TZ) if FECHA_LANZAMIENTO.tzinfo is None else FECHA_LANZAMIENTO
inicio_acum_utc = inicio_acum_gt.astimezone(timezone.utc)
inicio_acum_iso = inicio_acum_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
fin_acum_iso = fin_iso  # mismo fin que el día de consulta

# Inicio de anulados (desde FECHA_INICIO_ANULADOS 00:00:00 Guatemala)
inicio_anulados_gt = FECHA_INICIO_ANULADOS.replace(tzinfo=GUATE_TZ) if FECHA_INICIO_ANULADOS.tzinfo is None else FECHA_INICIO_ANULADOS
inicio_anulados_utc = inicio_anulados_gt.astimezone(timezone.utc)
inicio_anulados_iso = inicio_anulados_utc.strftime("%Y-%m-%dT%H:%M:%SZ")

print(inicio_iso)
print(fin_iso)

# Logo: ruta relativa al script (funciona aunque ejecutes desde otro directorio)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(SCRIPT_DIR, "logos", "guatecredito.png")
if os.path.exists(LOGO_PATH):
    logo = plt.imread(LOGO_PATH)
else:
    logo = None
    print("Aviso: no se encontró logos/guatecredito.png. Los reportes se generarán sin logo.")
# ======================================================
# CONFIGURACIÓN DE CARPETAS (rutas absolutas respecto al script)
# ======================================================

BASE_DIR = os.path.join(SCRIPT_DIR, "reportes")

# Carpeta general
os.makedirs(BASE_DIR, exist_ok=True)

# Carpeta exclusiva para Excel acumulado
EXCEL_DIR = os.path.join(BASE_DIR, "excel_acumulado")
os.makedirs(EXCEL_DIR, exist_ok=True)

# Carpeta del día (se crea nueva cada día)
CARPETA_DIA = os.path.join(BASE_DIR, FECHA_CONSULTA)
os.makedirs(CARPETA_DIA, exist_ok=True)

# Archivo Excel acumulado (fuente de verdad histórica)
# Nota: históricamente este archivo ha vivido junto al script (SCRIPT_DIR).
ARCHIVO_EXCEL = os.path.join(SCRIPT_DIR, "reporte_general.xlsx")
# Copia opcional organizada dentro de /reportes
ARCHIVO_EXCEL_COPIA = os.path.join(EXCEL_DIR, "reporte_general.xlsx")
# Copia en el proyecto (desarrollo); launchd usa Application Support
EXCEL_PROYECTO = "/Users/productolead/Documents/Laboratorios_LeanCore/Informes Guatecredito/reporte_general.xlsx"


def _sincronizar_excel_desde_proyecto() -> None:
    """Trae el Excel más reciente del repo si launchd corre en Application Support."""
    if not os.path.exists(EXCEL_PROYECTO):
        return
    try:
        if (not os.path.exists(ARCHIVO_EXCEL)) or (
            os.path.getmtime(EXCEL_PROYECTO) > os.path.getmtime(ARCHIVO_EXCEL)
        ):
            shutil.copy2(EXCEL_PROYECTO, ARCHIVO_EXCEL)
    except Exception as e:
        print("Aviso: no se pudo sincronizar Excel desde proyecto:", e)


def _publicar_excel_en_proyecto() -> None:
    try:
        shutil.copy2(ARCHIVO_EXCEL, ARCHIVO_EXCEL_COPIA)
        if os.path.abspath(ARCHIVO_EXCEL) != os.path.abspath(EXCEL_PROYECTO):
            shutil.copy2(ARCHIVO_EXCEL, EXCEL_PROYECTO)
    except Exception as e:
        print("Aviso: no se pudo copiar Excel al proyecto:", e)


_sincronizar_excel_desde_proyecto()

# ---- EMAIL ----
EMAIL_REMITENTE = "jhonatan@leancore.co"
#EMAIL_PASSWORD = "rgdb dldg gpwm urod"
EMAIL_PASSWORD = "usvz uwbv kied nwzu"
EMAIL_DESTINO = [
    "jhonatan@leancore.co",
    "lucas@leancore.co",
    "fsolorzano@bi.com.gt",
    "egarcia@energuate.com",
    "david_rod88@hotmail.com",
    "anderson@leancore.co",
]

if REPORTE_EMAILS_ENV:
    _emails = [e.strip() for e in REPORTE_EMAILS_ENV.split(",") if e.strip()]
    if _emails:
        EMAIL_DESTINO = _emails

# ======================================================
# 3. CONEXIÓN MONGO
# ======================================================

client = MongoClient(
    MONGO_URI,
    serverSelectionTimeoutMS=MONGO_TIMEOUT_MS,
    connectTimeoutMS=MONGO_TIMEOUT_MS,
    socketTimeoutMS=MONGO_TIMEOUT_MS,
)
db = client[DB_NAME]

lead_col = db[LEAD_COLLECTION]
loans_col = db[LOANS_COLLECTION]
users_col = db[USERS_COLLECTION]
payments_col = db[PAYMENT_COLLECTION]
deleted_loans_col = db[DELETED_LOAN_COLLECTION]

# ======================================================
# 4–6. MÉTRICAS DEL DÍA (y barrido de ayer si el informe es de hoy)
# ======================================================

prospectos_acumulados = lead_col.count_documents({
    "ecosystem": ECOSYSTEM_ID,
    "created_at": {"$gte": inicio_acum_iso, "$lte": fin_acum_iso}
})
clientes_acumulados = users_col.count_documents({
    "ecosystem": ECOSYSTEM_ID,
    "created_at": {"$gte": inicio_acum_iso, "$lte": fin_acum_iso}
})

pipeline_loans_acum = [
    {"$match": {
        "financial_entity_id": ECOSYSTEM_ID,
        "assignment_date": {"$gte": inicio_acum_iso, "$lte": fin_acum_iso}
    }},
    {"$group": {
        "_id": None,
        "creditos_nuevos": {"$sum": 1},
        "desembolsos": {"$sum": {"$cond": [{"$eq": ["$disbursed", True]}, 1, 0]}},
        "no_desembolsos": {"$sum": {"$cond": [{"$eq": ["$disbursed", False]}, 1, 0]}},
        "monto_desembolsado": {"$sum": {"$cond": [{"$eq": ["$disbursed", True]}, "$amount", 0]}}
    }}
]
resultado_loans_acum = list(loans_col.aggregate(pipeline_loans_acum))
if resultado_loans_acum:
    data_loans_acum = resultado_loans_acum[0]
else:
    data_loans_acum = {"creditos_nuevos": 0, "desembolsos": 0, "no_desembolsos": 0, "monto_desembolsado": 0}

creditos_acumulados = data_loans_acum["creditos_nuevos"]
desembolsos_acumulados = data_loans_acum["desembolsos"]
no_desembolsos_acumulados = data_loans_acum["no_desembolsos"]
monto_desembolsado_acum = data_loans_acum["monto_desembolsado"] / 100
ticket_promedio_acum = round(monto_desembolsado_acum / desembolsos_acumulados, 2) if desembolsos_acumulados else 0
clientes_abonados_unicos_acum = _clientes_abonados_unicos_periodo(
    payments_col, ECOSYSTEM_ID, inicio_acum_iso, fin_iso
)

pipeline_rechazos_acum = [
    {"$match": {
        "ecosystem": ECOSYSTEM_ID,
        "created_at": {"$gte": inicio_acum_iso, "$lte": fin_acum_iso}
    }},
    {"$unwind": "$additional_data.variables_data"},
    {"$match": {"additional_data.variables_data.key": {"$in": ["motivos_negacion", "motivoNegacion"]}}},
    {"$group": {"_id": "$additional_data.variables_data.value", "cantidad": {"$sum": 1}}}
]
resultados_rechazos_acum = list(lead_col.aggregate(pipeline_rechazos_acum))
rechazo_politicas_acum = 0
for r in resultados_rechazos_acum:
    motivo = (r["_id"] or "").lower()
    cantidad = r["cantidad"]
    if ("no cumple" in motivo or "identificación" in motivo or "nis" in motivo or "lista negra" in motivo):
        rechazo_politicas_acum += cantidad

solicitudes_acumuladas = prospectos_acumulados + clientes_acumulados
rechazos_bd_acum = prospectos_acumulados - rechazo_politicas_acum

# ======================================================
# 6c. PAGOS / ABONOS (colección payment) — por día Guatemala, rango acumulado
# ======================================================
ABONOS_POR_DIA = _abonos_por_dia_guatemala(
    payments_col, ECOSYSTEM_ID, inicio_acum_iso, fin_iso
)

ANULADOS_POR_DIA = _anulados_por_dia_guatemala(
    deleted_loans_col, ECOSYSTEM_ID, inicio_anulados_iso, fin_iso
)

HOY_GT = datetime.now(GUATE_TZ).strftime("%Y-%m-%d")
filas_excel = []

# Si el informe es del día de hoy, refrescar también ayer (créditos que entran tarde al sistema)
if FECHA_CONSULTA == HOY_GT:
    ayer_gt = (datetime.now(GUATE_TZ) - timedelta(days=1)).strftime("%Y-%m-%d")
    filas_excel.append(calcular_metricas_dia(
        ayer_gt, lead_col, loans_col, users_col, ABONOS_POR_DIA, ANULADOS_POR_DIA,
    ))
    print(f"[Barrido automático] Actualizando fila de {ayer_gt} en Excel acumulado.")

filas_excel.append(calcular_metricas_dia(
    FECHA_CONSULTA, lead_col, loans_col, users_col, ABONOS_POR_DIA, ANULADOS_POR_DIA,
))

metricas_dia = filas_excel[-1]
total_solicitudes = metricas_dia["Solicitudes"]
rechazos_bd = metricas_dia["Rechazos - BD"]
clientes_nuevos = metricas_dia["Clientes nuevos"]
rechazo_politicas = metricas_dia["Rechazos - Políticas"]
creditos_nuevos = metricas_dia["Créditos nuevos"]
desembolsos = metricas_dia["Desembolsos"]
no_desembolsos = metricas_dia["No desembolsos"]
monto_desembolsado = metricas_dia["Monto desembolsado"]
ticket_promedio = metricas_dia["Ticket promedio"]
cantidad_abonos = metricas_dia["# de abonos recibidos"]
clientes_abonados_dia = metricas_dia["# de clientes abonados"]
monto_abonos = metricas_dia["Monto abonos"]
creditos_anulados = metricas_dia["Créditos anulados"]

data_reporte = pd.DataFrame([metricas_dia])

# ======================================================
# 8. EXCEL ACUMULADO
# ======================================================

def _leer_excel_acumulado(path: str) -> pd.DataFrame:
    """
    Lee el excel acumulado tolerando distintos offsets de encabezado.
    Devuelve DataFrame con columnas esperadas (incluida 'Fecha') o vacío si no se puede leer.
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    for header in (4, 0, 1, 2, 3, 5, 6, 7, 8, 9, 10):
        try:
            df = pd.read_excel(path, header=header)
        except Exception:
            continue
        if isinstance(df, pd.DataFrame) and "Fecha" in df.columns:
            return df
    return pd.DataFrame()

# Si existe un excel previo en la carpeta organizada pero no en raíz,
# lo migramos a raíz para conservar el histórico.
if (not os.path.exists(ARCHIVO_EXCEL)) and os.path.exists(ARCHIVO_EXCEL_COPIA):
    try:
        shutil.copy2(ARCHIVO_EXCEL_COPIA, ARCHIVO_EXCEL)
    except Exception:
        pass

df_existente = _leer_excel_acumulado(ARCHIVO_EXCEL)
df_existente_copia = _leer_excel_acumulado(ARCHIVO_EXCEL_COPIA)

# Unificar histórico desde ambas ubicaciones (raíz + carpeta organizada)
if (not df_existente.empty) or (not df_existente_copia.empty):
    df_base = pd.concat([df_existente, df_existente_copia], ignore_index=True)

    # Normalizar Fecha a string YYYY-MM-DD cuando venga como datetime
    try:
        df_base["Fecha"] = pd.to_datetime(df_base["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    # Deduplicar por Fecha (si existe en ambos, nos quedamos con el último)
    if "Fecha" in df_base.columns:
        df_base = df_base.dropna(subset=["Fecha"])
        df_base = df_base.drop_duplicates(subset=["Fecha"], keep="last")
        df_base = df_base.sort_values("Fecha")

    for fila in filas_excel:
        df_base = df_base[df_base["Fecha"] != fila["Fecha"]]
    df_final = pd.concat([df_base, pd.DataFrame(filas_excel)], ignore_index=True)
else:
    df_final = pd.DataFrame(filas_excel)

df_final = df_final.copy()

# Asegurar columnas y rellenarlas para todas las fechas (histórico)
for col in ("# de abonos recibidos", "# de clientes abonados", "Monto abonos", "Créditos anulados"):
    if col not in df_final.columns:
        df_final[col] = 0

for idx in df_final.index:
    f = df_final.at[idx, "Fecha"]
    f_str = pd.to_datetime(f, errors="coerce")
    if pd.isna(f_str):
        continue
    f_str = f_str.strftime("%Y-%m-%d")

    c, m, cu = ABONOS_POR_DIA.get(f_str, (0, 0.0, 0))
    df_final.at[idx, "# de abonos recibidos"] = c
    df_final.at[idx, "# de clientes abonados"] = cu
    df_final.at[idx, "Monto abonos"] = m

    c2, _m2 = ANULADOS_POR_DIA.get(f_str, (0, 0.0))
    df_final.at[idx, "Créditos anulados"] = c2

# Asegurar que nunca aparezca esta columna en el Excel
df_final = df_final.drop(columns=["Monto créditos anulados"], errors="ignore")

df_final.to_excel(ARCHIVO_EXCEL, index=False, startrow=4)

# Cargar workbook
wb = load_workbook(ARCHIVO_EXCEL)
ws = wb.active



for column in ws.columns:
    max_length = 0
    column_letter = get_column_letter(column[0].column)

    for cell in column:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass

    adjusted_width = max_length + 2
    ws.column_dimensions[column_letter].width = adjusted_width

fecha_ejecucion = hoy_ejecucion.strftime("%Y-%m-%d %H:%M:%S")

ws["C1"] = "Reporte acumulado Guatecrédito"
ws["C1"].font = Font(bold=True)

ws["C2"] = f"Fecha de ejecución: {fecha_ejecucion}"

# Insertar logo (solo si existe el archivo)
if os.path.exists(LOGO_PATH):
    logo_excel = XLImage(LOGO_PATH)
    logo_excel.width = 120
    logo_excel.height = 60
    ws.add_image(logo_excel, "A1")

wb.save(ARCHIVO_EXCEL)

_publicar_excel_en_proyecto()

# Excel adjunto al correo: acumulado actualizado hasta la fecha del informe (sin días futuros)
ARCHIVO_EXCEL_ADJUNTO = os.path.join(CARPETA_DIA, f"reporte_general_{FECHA_CONSULTA}.xlsx")
try:
    df_envio = df_final.copy()
    df_envio["Fecha"] = pd.to_datetime(df_envio["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d")
    df_envio = df_envio.dropna(subset=["Fecha"])
    df_envio = df_envio[df_envio["Fecha"] <= FECHA_CONSULTA].sort_values("Fecha")
    df_envio.to_excel(ARCHIVO_EXCEL_ADJUNTO, index=False, startrow=4)

    wb2 = load_workbook(ARCHIVO_EXCEL_ADJUNTO)
    ws2 = wb2.active
    ws2["C1"] = "Reporte acumulado Guatecrédito"
    ws2["C1"].font = Font(bold=True)
    ws2["C2"] = f"Fecha de ejecución: {fecha_ejecucion}"

    if os.path.exists(LOGO_PATH):
        logo_excel2 = XLImage(LOGO_PATH)
        logo_excel2.width = 120
        logo_excel2.height = 60
        ws2.add_image(logo_excel2, "A1")

    for column in ws2.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws2.column_dimensions[column_letter].width = max_length + 2

    wb2.save(ARCHIVO_EXCEL_ADJUNTO)
except Exception as e:
    print("Aviso: no se pudo generar Excel adjunto dedicado:", e)
    ARCHIVO_EXCEL_ADJUNTO = ARCHIVO_EXCEL


# ======================================================
# 7. IMAGEN DIARIA
# ======================================================

fig, ax = plt.subplots(figsize=(6,8))
ax.axis("off")

# TITULO
ax.text(0.5, 0.96, "REPORTE DIARIO", ha="center", fontsize=18, weight="bold")
ax.text(0.5, 0.92, f"Fecha: {FECHA_CONSULTA}", ha="center", fontsize=11)
ax.text(0.5, 0.88, f"Fecha de consulta: {fecha_ejecucion}", ha="center", fontsize=11)

add_logo(ax, logo, 0.08, 0.95)
add_logo(ax, logo, 0.92, 0.95)


labels = [
    "Solicitudes",
    "Rechazos - BD",
    "Clientes nuevos",
    "Rechazos - Políticas",
    "Créditos nuevos",
    "Desembolsos",
    "No desembolsos",
    "Monto desembolsado",
    "Ticket promedio",
    "# de abonos recibidos",
    "# de clientes abonados",
    "Monto abonos",
    "Créditos anulados"
    
]

values = [
    total_solicitudes,
    rechazos_bd,
    clientes_nuevos,
    rechazo_politicas,
    creditos_nuevos,
    desembolsos,
    no_desembolsos,
    f"Q{monto_desembolsado:,.2f}",
    f"Q{ticket_promedio:,.2f}",
    cantidad_abonos,
    clientes_abonados_dia,
    f"Q{monto_abonos:,.2f}",
    creditos_anulados
]

y = 0.72
for label, value in zip(labels, values):
    ax.text(0.10, y, label, fontsize=11, weight="bold")
    ax.text(0.80, y, str(value), ha="right", fontsize=11)
    y -= 0.06

imagen_diaria = os.path.join(
    CARPETA_DIA,
    f"reporte_diario_{FECHA_CONSULTA}_{hora_actual}.png"
)

plt.subplots_adjust(top=0.92)
plt.savefig(imagen_diaria, dpi=300)
plt.close()

# ======================================================
# IMAGEN ACUMULADO (sumas desde Excel acumulado)
# ======================================================

fig, ax = plt.subplots(figsize=(6,8))
ax.axis("off")

# Totales acumulados basados en el Excel (df_final)
df_acum = df_final.copy()
try:
    df_acum["Fecha"] = pd.to_datetime(df_acum["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d")
except Exception:
    pass
df_acum = df_acum.dropna(subset=["Fecha"])
df_acum = df_acum[df_acum["Fecha"] <= FECHA_CONSULTA].sort_values("Fecha")

def _sum_col(name: str) -> float:
    return float(pd.to_numeric(df_acum.get(name), errors="coerce").fillna(0).sum())

solicitudes_acumuladas = int(_sum_col("Solicitudes"))
rechazos_bd_acum = int(_sum_col("Rechazos - BD"))
clientes_acumulados = int(_sum_col("Clientes nuevos"))
rechazo_politicas_acum = int(_sum_col("Rechazos - Políticas"))
creditos_acumulados = int(_sum_col("Créditos nuevos"))
desembolsos_acumulados = int(_sum_col("Desembolsos"))
no_desembolsos_acumulados = int(_sum_col("No desembolsos"))
monto_desembolsado_acum = _sum_col("Monto desembolsado")
ticket_promedio_acum = round(monto_desembolsado_acum / desembolsos_acumulados, 2) if desembolsos_acumulados else 0

# TITULO
ax.text(0.5, 0.96, "ACUMULADO", ha="center", fontsize=18, weight="bold")
ax.text(0.5, 0.88, f"Desde: {FECHA_LANZAMIENTO} hasta: {FECHA_CONSULTA}", ha="center", fontsize=11)
ax.text(0.5, 0.84, f"Fecha de consulta: {fecha_ejecucion}", ha="center", fontsize=11)

add_logo(ax, logo, 0.08, 0.95)
add_logo(ax, logo, 0.92, 0.95)

labels_acum = [
    "Solicitudes",
    "Rechazos - BD",
    "Clientes nuevos",
    "Rechazos - Políticas",
    "Créditos nuevos",
    "Desembolsos",
    "No desembolsos",
    "Monto desembolsado",
    "Ticket promedio",
    "# de abonos recibidos",
    "# de clientes abonados",
    "Monto abonos",
    "Créditos anulados"
]

values_acum = [
    solicitudes_acumuladas,
    rechazos_bd_acum,
    clientes_acumulados,
    rechazo_politicas_acum,
    creditos_acumulados,
    desembolsos_acumulados,
    no_desembolsos_acumulados,
    f"Q{monto_desembolsado_acum:,.2f}",
    f"Q{ticket_promedio_acum:,.2f}",
    int(_sum_col("# de abonos recibidos")),
    clientes_abonados_unicos_acum,
    f"Q{_sum_col('Monto abonos'):,.2f}",
    int(_sum_col("Créditos anulados"))
]

y = 0.72
for label, value in zip(labels_acum, values_acum):
    ax.text(0.10, y, label, fontsize=11, weight="bold")
    ax.text(0.80, y, str(value), ha="right", fontsize=11)
    y -= 0.06

imagen_acumulada = os.path.join(
    CARPETA_DIA,
    f"reporte_acumulado_{FECHA_CONSULTA}_{hora_actual}.png"
)

plt.subplots_adjust(top=0.92)
plt.savefig(imagen_acumulada, dpi=300)
plt.close()

# ======================================================
# CONFIRMACIÓN FINAL
# ======================================================
correo_enviado = False
msg = EmailMessage()
msg["Subject"] = f"Reporte Guatecrédito {FECHA_CONSULTA}"
msg["From"] = EMAIL_REMITENTE
msg["To"] = ", ".join(EMAIL_DESTINO)

msg.set_content(
    f"Adjunto reporte diario y acumulado generado.\n\n"
    f"Fecha y hora de ejecución: {FECHA_CONSULTA} - {hora_actual}"
)

with open(imagen_diaria, "rb") as f:
    msg.add_attachment(f.read(), maintype="image", subtype="png", filename="reporte_diario.png")

with open(imagen_acumulada, "rb") as f:
    msg.add_attachment(f.read(), maintype="image", subtype="png", filename="reporte_acumulado.png")

with open(ARCHIVO_EXCEL_ADJUNTO, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="reporte_general.xlsx"
    )

try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=SMTP_TIMEOUT_SEC) as smtp:
        smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
        smtp.send_message(msg)
    correo_enviado = True
except Exception as e:
    print("Error enviando correo:", e)
    correo_enviado = False

print("\n" + "=" * 50)
print("      RESUMEN DE EJECUCIÓN")
print("=" * 50)

# Confirmar Excel
if os.path.exists(ARCHIVO_EXCEL):
    print("✔ Excel generado correctamente")
else:
    print("✘ Error generando Excel")

# Confirmar imagen diaria
if os.path.exists(imagen_diaria):
    print("✔ Imagen diaria generada correctamente")
else:
    print("✘ Error generando imagen diaria")

# Confirmar imagen acumulada
if os.path.exists(imagen_acumulada):
    print("✔ Imagen acumulada generada correctamente")
else:
    print("✘ Error generando imagen acumulada")

if correo_enviado:
    print("✔ Envío de correo realizado correctamente")
else:
    print("✘ Falló el envío de correo")