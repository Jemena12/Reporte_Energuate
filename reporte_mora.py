# ======================================================
# Reporte de usuarios en mora (status "arrear" → informe "mora")
# Basado en la configuración de reporte_diario.py
# ======================================================
from __future__ import annotations

import os
import shutil
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from pymongo import MongoClient

# ---- MONGO (misma fuente que reporte_diario.py) ----
MONGO_URI = "mongodb+srv://camilo:XDwsOiclZ7AvO6G1@fintechmiddlewarecluste.vvbho.mongodb.net/middleware"
DB_NAME = "middleware"
USERS_COLLECTION = "user"
ECOSYSTEM_ID = "eab85a60-8565-4666-9a83-428a48d0c7dc"

# Evita que launchd quede con el job "running" días si Mongo/SMTP no responden
MONGO_TIMEOUT_MS = 60_000
SMTP_TIMEOUT_SEC = 120

GUATE_TZ = timezone(timedelta(hours=-6))

REPORTE_FECHA_ENV = os.environ.get("REPORTE_FECHA")
REPORTE_EMAILS_ENV = os.environ.get("REPORTE_EMAILS")

FECHA_INFORME_MANUAL = None
DIAS_ATRAS = 0

if REPORTE_FECHA_ENV and not FECHA_INFORME_MANUAL:
    FECHA_INFORME_MANUAL = REPORTE_FECHA_ENV.strip()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.join(SCRIPT_DIR, "reportes")
EXCEL_DIR = os.path.join(BASE_DIR, "excel_acumulado")

# Excel maestro con histórico de conteos por día (acumulativo)
ARCHIVO_HISTORICO_MORA = os.path.join(EXCEL_DIR, "reporte_mora_historico.xlsx")

# ---- EMAIL (mismo remitente que reporte_diario.py) ----
EMAIL_REMITENTE = "jhonatan@leancore.co"
EMAIL_PASSWORD = "usvz uwbv kied nwzu"
# Prueba: solo Jhonatan; en producción se puede sobreescribir con REPORTE_EMAILS
EMAIL_DESTINO = ["jhonatan@leancore.co",
                "wilson@leancore.co",
                "anderson@leancore.co"]

if REPORTE_EMAILS_ENV:
    _emails = [e.strip() for e in REPORTE_EMAILS_ENV.split(",") if e.strip()]
    if _emails:
        EMAIL_DESTINO = _emails


def _fecha_informe_guatemala() -> datetime:
    if FECHA_INFORME_MANUAL:
        d = datetime.strptime(str(FECHA_INFORME_MANUAL).strip(), "%Y-%m-%d").date()
        return datetime(d.year, d.month, d.day, tzinfo=GUATE_TZ)
    return datetime.now(GUATE_TZ) - timedelta(days=DIAS_ATRAS)


def _nombre_completo(doc: dict) -> str:
    n = (doc.get("name") or "").strip()
    ln = (doc.get("lastname") or "").strip()
    return f"{n} {ln}".strip() or "(sin nombre)"


def _cargar_historico(path: str) -> list[tuple[str, int]]:
    if not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb["Historico"]
        out: list[tuple[str, int]] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            fecha = str(row[0]).strip()[:10]
            try:
                cnt = int(row[1] or 0)
            except (TypeError, ValueError):
                cnt = 0
            out.append((fecha, cnt))
        wb.close()
        out.sort(key=lambda x: x[0])
        return out
    except Exception as e:
        print("Aviso: no se pudo leer histórico previo:", e)
        return []


def _guardar_historico(path: str, filas: list[tuple[str, int]]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Historico"
    ws.append(["Fecha (Guatemala)", "Usuarios en mora"])
    for fecha, cnt in filas:
        ws.append([fecha, cnt])
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        maxlen = 0
        for cell in col:
            if cell.value is not None:
                maxlen = max(maxlen, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max(maxlen + 2, 12), 40)
    wb.save(path)


def _construir_workbook_informe(
    fecha_display: str,
    fecha_iso: str,
    cantidad: int,
    usuarios: list[dict],
    historico: list[tuple[str, int]],
) -> Workbook:
    wb = Workbook()

    # --- Hoja 1: informe del día (diario) ---
    ws = wb.active
    ws.title = "Informe del día"

    ws["A1"] = f"Fecha: {fecha_display}"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = f"cantidad de usuario: {cantidad}"
    ws["A2"].font = Font(bold=True, size=12)

    ws["A4"] = "DPI"
    ws["B4"] = "Nombre"
    ws["C4"] = "Estado"
    for cell in (ws["A4"], ws["B4"], ws["C4"]):
        cell.font = Font(bold=True)

    row = 5
    for doc in usuarios:
        ws.cell(row=row, column=1, value=str(doc.get("nit") or ""))
        ws.cell(row=row, column=2, value=_nombre_completo(doc))
        ws.cell(row=row, column=3, value="mora")
        row += 1

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 12
    for r in range(1, row):
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=True)

    # --- Hoja 2: histórico acumulativo (conteos por fecha) ---
    wh = wb.create_sheet("Historico acumulativo")
    wh.append(["Fecha (Guatemala)", "Usuarios en mora"])
    wh["A1"].font = Font(bold=True)
    wh["B1"].font = Font(bold=True)
    for fecha, cnt in historico:
        wh.append([fecha, cnt])
    for col in wh.columns:
        letter = get_column_letter(col[0].column)
        maxlen = 0
        for cell in col:
            if cell.value is not None:
                maxlen = max(maxlen, len(str(cell.value)))
        wh.column_dimensions[letter].width = min(max(maxlen + 2, 12), 40)

    # Metadatos discretos
    ws["E1"] = "Fecha ISO (Guatemala)"
    ws["F1"] = fecha_iso
    ws["E1"].font = Font(italic=True, size=9)
    ws["F1"].font = Font(italic=True, size=9)

    return wb


def main() -> None:
    hoy = _fecha_informe_guatemala()
    if REPORTE_FECHA_ENV or FECHA_INFORME_MANUAL or DIAS_ATRAS:
        print(
            f"[PARÉNTESIS] Informe mora para el día {hoy.strftime('%Y-%m-%d')} (Guatemala). "
            "Tras esta ejecución, pon FECHA_INFORME_MANUAL=None y DIAS_ATRAS=0."
        )

    FECHA_CONSULTA = hoy.strftime("%Y-%m-%d")
    fecha_display = hoy.strftime("%d - %m")
    hora_actual = (datetime.now(GUATE_TZ)).strftime("%H%M%S")
    hoy_ejecucion = datetime.now(GUATE_TZ).strftime("%Y-%m-%d %H:%M:%S")

    os.makedirs(BASE_DIR, exist_ok=True)
    os.makedirs(EXCEL_DIR, exist_ok=True)
    carpeta_dia = os.path.join(BASE_DIR, FECHA_CONSULTA)
    os.makedirs(carpeta_dia, exist_ok=True)

    client = MongoClient(
        MONGO_URI,
        serverSelectionTimeoutMS=MONGO_TIMEOUT_MS,
        connectTimeoutMS=MONGO_TIMEOUT_MS,
        socketTimeoutMS=MONGO_TIMEOUT_MS,
    )
    try:
        db = client[DB_NAME]
        users_col = db[USERS_COLLECTION]

        filtro = {"ecosystem": ECOSYSTEM_ID, "status": "arrear"}
        cantidad = users_col.count_documents(filtro)
        cursor = users_col.find(
            filtro,
            {"nit": 1, "name": 1, "lastname": 1, "status": 1},
        ).sort([("name", 1), ("lastname", 1)])
        usuarios = list(cursor)
    finally:
        client.close()

    historico = _cargar_historico(ARCHIVO_HISTORICO_MORA)
    historico_dict = dict(historico)
    historico_dict[FECHA_CONSULTA] = cantidad
    historico_ordenado = sorted(historico_dict.items(), key=lambda x: x[0])
    _guardar_historico(ARCHIVO_HISTORICO_MORA, historico_ordenado)

    wb = _construir_workbook_informe(
        fecha_display=fecha_display,
        fecha_iso=FECHA_CONSULTA,
        cantidad=cantidad,
        usuarios=usuarios,
        historico=historico_ordenado,
    )

    archivo_dia = os.path.join(carpeta_dia, f"reporte_mora_{FECHA_CONSULTA}_{hora_actual}.xlsx")
    wb.save(archivo_dia)

    archivo_adjunto = os.path.join(carpeta_dia, f"reporte_mora_{FECHA_CONSULTA}.xlsx")
    try:
        shutil.copy2(archivo_dia, archivo_adjunto)
    except Exception:
        archivo_adjunto = archivo_dia

    try:
        shutil.copy2(archivo_adjunto, os.path.join(EXCEL_DIR, f"reporte_mora_{FECHA_CONSULTA}.xlsx"))
    except Exception:
        pass

    msg = EmailMessage()
    msg["Subject"] = f"Reporte mora Guatecrédito {FECHA_CONSULTA}"
    msg["From"] = EMAIL_REMITENTE
    msg["To"] = ", ".join(EMAIL_DESTINO)

    cuerpo = (
        "Adjunto reporte de usuarios en mora.\n\n"
        f"Fecha del informe (Guatemala): {FECHA_CONSULTA}\n"
        f"Usuarios en mora: {cantidad}\n"
        f"Fecha y hora de ejecución: {hoy_ejecucion}\n\n"
    )
    msg.set_content(cuerpo)

    with open(archivo_adjunto, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=f"reporte_mora_{FECHA_CONSULTA}.xlsx",
        )

    correo_ok = False
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=SMTP_TIMEOUT_SEC) as smtp:
            smtp.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            smtp.send_message(msg)
        correo_ok = True
    except Exception as e:
        print("Error enviando correo:", e)

    print("\n" + "=" * 50)
    print("      RESUMEN — REPORTE MORA")
    print("=" * 50)
    print(f"Fecha informe: {FECHA_CONSULTA}  |  Usuarios en mora: {cantidad}")
    print(f"Excel del día: {archivo_adjunto}")
    print(f"Histórico maestro: {ARCHIVO_HISTORICO_MORA}")
    print("✔ Envío de correo realizado correctamente" if correo_ok else "✘ Falló el envío de correo")


if __name__ == "__main__":
    main()
