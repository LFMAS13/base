#!/usr/bin/env python3
"""
Genera dashboard.html a partir de un Excel de seguimiento de eventos.

Uso:
python3 build_dashboard.py ruta/al/excel.xlsx [ruta/salida.html]

Si no se da ruta de salida, escribe "dashboard.html" en el directorio actual.
Requiere: pip install openpyxl --break-system-packages
"""
import sys
import json
import datetime
import openpyxl

TEMPLATE_PATH = "template.html"


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
return v


def clean_number(v):
v = clean(v)
if v is None:
return None
try:
return float(v)
except (TypeError, ValueError):
return None


def load_rows(xlsx_path):
wb = openpyxl.load_workbook(xlsx_path, data_only=True)
ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb[wb.sheetnames[0]]
rows = list(ws.iter_rows(values_only=True))
headers = [h.strip() if isinstance(h, str) else h for h in rows[0]]
idx = {h: i for i, h in enumerate(headers)}

def get(row, name):
i = idx.get(name)
return row[i] if i is not None and i < len(row) else None

records = []
for row in rows[1:]:
ev_num = clean(get(row, "Número de evento")) or clean(get(row, "EVENTO+PONENTE"))
# fila totalmente vacía / sin ningún dato útil -> se descarta
if not any(clean(v) is not None for v in row):
continue
rec = {
"evNum": ev_num,
"eventName": clean(get(row, "Nombre del evento")),
"speaker": clean(get(row, "Nombre Speaker")),
"speakerStatus": clean(get(row, "Estatus del Speaker")),
"eventStatus": clean(get(row, "Estatus del Evento")),
"areaTerapeutica": clean(get(row, "AT")),
"franquicia": clean(get(row, "Franquicia")),
"fecha": clean(get(row, "Fecha del evento")),
"year": clean(get(row, "Year")),
"month": clean(get(row, "Month")),
"contractStatus": clean(get(row, "Estatus de Contrato")),
"contractNum": clean(get(row, "Número ICD contrato")),
"owner": clean(get(row, "Nombre del Dueño del evento")),
"producto": clean(get(row, "Producto/Marca")),
"invoiceStatus": clean(get(row, "Estatus de factura")),
"invoiceValue": clean_number(get(row, "VALOR DE FACTURA ")),
"paymentStatus": clean(get(row, "Estatus del pago de factura")),
"exception": clean(get(row, "Solicitud de excepción")),
"logisticaSpeaker": clean(get(row, "Logística del speaker")),
}
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

data_json = json.dumps(records, ensure_ascii=False, indent=2)
output = template.replace("/*__DATA__*/", data_json)

with open(out_path, "w", encoding="utf-8") as f:
f.write(output)

print(f"OK: {len(records)} filas incrustadas -> {out_path}")


if __name__ == "__main__":
main()
