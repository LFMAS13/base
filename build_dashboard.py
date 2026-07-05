#!/usr/bin/env python3
"""
Genera dashboard.html a partir de un Excel de seguimiento de eventos.

Uso:
    python3 build_dashboard.py ruta/al/excel.xlsx [ruta/salida.html]

Si no se da ruta de salida, escribe "dashboard.html" en el directorio actual.
Requiere que "template.html" (el HTML del dashboard) esté en el mismo directorio
que este script, o ajusta TEMPLATE_PATH abajo.

Requiere: pip install openpyxl --break-system-packages
"""
import sys
import json
import datetime
import openpyxl

TEMPLATE_PATH = "template.html"

# Nombres EXACTOS de columnas que espera el dashboard (deben coincidir con tu Excel)
COLUMNS = [
    "Revisión", "Acción Actual", "EVENTO+PONENTE", "Número de evento",
    "Estatus del Evento", "AT", "Tipo de Evento", "Franquicia",
    "Nombre del evento", "Nombre Speaker", "Estatus del Speaker", "ID Vendor",
    "Régimen fiscal", "Fecha del evento", "Year", "Month",
    "Estatus de Contrato", "Número ICD contrato", "Request Navigator",
    "Attendee ID", "Solicitud de excepción", "Logística del speaker",
    "Fecha solicitud del evento", "Fecha de firma de contrato",
    "Nombre del Dueño del evento", "Dirección correo elec. ponente/consultor",
    "Fecha de solicitud de pago en sistema", "Producto/Marca", "CECO", "OI",
    "ACCOUNT", "Estatus de factura", "VALOR DE FACTURA",
    "Estatus del pago de factura", "Fecha de solicitud de factura", "UUID",
    "Posting date",
]

# Columnas que deben quedar como número (el resto se trata como texto/fecha)
NUMERIC_COLUMNS = {
    "Year", "Month", "Número ICD contrato", "Solicitud de excepción",
    "VALOR DE FACTURA",
}


def clean(v):
    """Limpia celdas: convierte #REF!, cadenas vacías y espacios sobrantes en None."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s.upper() == "#REF!":
            return None
        return s
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()[:10]
    if isinstance(v, datetime.time):
        return v.isoformat()
    return v


def clean_number(v):
    v = clean(v)
    if v is None:
        return None
    try:
        n = float(v)
        # si es entero exacto, guardarlo como int para que se vea limpio en el JSON
        return int(n) if n.is_integer() else n
    except (TypeError, ValueError):
        return None


def pick_sheet(wb):
    # Ajusta este nombre si tu pestaña de eventos se llama distinto
    for name in ("Nacionales", "Eventos", "Sheet1"):
        if name in wb.sheetnames:
            return wb[name]
    return wb[wb.sheetnames[0]]


def load_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = pick_sheet(wb)
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [h.strip() if isinstance(h, str) else h for h in rows[0]]
    idx = {h: i for i, h in enumerate(headers)}

    missing = [c for c in COLUMNS if c not in idx]
    if missing:
        print("Aviso: no se encontraron estas columnas en el Excel (se guardarán como null):")
        for m in missing:
            print(f"   - {m}")

    def get(row, name):
        i = idx.get(name)
        return row[i] if i is not None and i < len(row) else None

    records = []
    for row in rows[1:]:
        if not any(clean(v) is not None for v in row):
            continue  # fila totalmente vacía

        rec = {}
        for col in COLUMNS:
            raw = get(row, col)
            rec[col] = clean_number(raw) if col in NUMERIC_COLUMNS else clean(raw)
        records.append(rec)

    return records


def main():
    if len(sys.argv) < 2:
        print("Uso: python3 build_dashboard.py ruta/al/excel.xlsx [salida.html]")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "dashboard.html"

    records = load_rows(xlsx_path)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        template = f.read()

    marker = "const EVENTOS_RAW = "
    start_marker = template.find(marker)
    if start_marker == -1:
        print(f'Error: no encontré "{marker}" en {TEMPLATE_PATH}')
        sys.exit(1)

    json_start = start_marker + len(marker)
    # Usamos el decodificador de JSON para encontrar exactamente dónde termina
    # el arreglo actual, sin importar su contenido (comillas, corchetes, etc).
    decoder = json.JSONDecoder()
    try:
        _, json_end = decoder.raw_decode(template, json_start)
    except json.JSONDecodeError:
        print("Error: no pude leer el arreglo EVENTOS_RAW existente en el template.")
        sys.exit(1)

    data_json = json.dumps(records, ensure_ascii=False)
    output = template[:json_start] + data_json + template[json_end:]

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"OK: {len(records)} filas incrustadas -> {out_path}")


if __name__ == "__main__":
    main()
