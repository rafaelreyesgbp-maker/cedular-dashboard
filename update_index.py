#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_index.py
Descarga los XLS de la carpeta de Google Drive del Impuesto Cedular sobre la
Obtencion de Ingresos por Actividades Empresariales, calcula el acumulado real,
la proyeccion de cierre y los segmentos de omisos, y actualiza index.html con
los resultados (allData) sin sobrescribir datos mejores ya existentes.

Config leida de:
  - FOLDER_ID / METAS / HIST_2025 (constantes abajo, deben coincidir con index.html)
  - DRIVE_API_KEY (variable de entorno, secret de GitHub Actions)
"""

import os
import re
import sys
import json
import math
import unicodedata
from datetime import datetime

import requests
import xlrd

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
FOLDER_ID = "1nFW72GZ3XLmbDasLJD0JY9M7NF9MtB-Y"
API_KEY = os.environ.get("DRIVE_API_KEY", "")
HTML_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")

METAS = {
    1: 8891728, 2: 7958748, 3: 7867639, 4: 8511439,
    5: 8595150, 6: 7849800, 7: 8407930, 8: 8215539,
    9: 7617855, 10: 7846543, 11: 8383464, 12: 7943525,
}

MONTH_NAMES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}

MONTH_LABELS = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# Cedular: RFC=A(0), Contrib=B(1), Periodo=E(4), Recaudacion = N(13) - H(7)
# Verificado contra datos reales: columna E trae periodo YYYYMM (ej. 202601),
# NO un serial de fecha de Excel. N-H coincide con el monto neto observado.
CED_RFC = 0
CED_CONTRIB = 1
CED_PERIODO = 4
CED_N = 13
CED_H = 7

VALID_MONTH_KEYS = {str(i) for i in range(1, 13)}


# --------------------------------------------------------------------------
# GOOGLE DRIVE
# --------------------------------------------------------------------------
def drive_list_files(folder_id, api_key):
    """Lista archivos de la carpeta. Sin filtro MIME (acepta .xls y .xlsx)."""
    url = "https://www.googleapis.com/drive/v3/files"
    params = {
        "q": f"'{folder_id}' in parents and trashed=false",
        "fields": "files(id,name,mimeType)",
        "pageSize": 100,
        "key": api_key,
    }
    r = requests.get(url, params=params, timeout=30)
    if not r.ok:
        raise RuntimeError(f"Drive list error {r.status_code}: {r.text[:300]}")
    return r.json().get("files", [])


def drive_download(file_id, api_key):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
    params = {"alt": "media", "key": api_key}
    r = requests.get(url, params=params, timeout=60)
    if not r.ok:
        raise RuntimeError(f"Drive download error {r.status_code}: {r.text[:300]}")
    return r.content


# --------------------------------------------------------------------------
# DETECCION DE MES POR NOMBRE DE ARCHIVO
# --------------------------------------------------------------------------
def normalize_str(s):
    s = s or ""
    nfkd = unicodedata.normalize("NFD", s)
    only_ascii = "".join(c for c in nfkd if unicodedata.category(c) != "Mn")
    return only_ascii.lower()


def detect_month(filename):
    n = normalize_str(filename)
    for name, num in MONTH_NAMES.items():
        if name in n:
            return num
    return None


# --------------------------------------------------------------------------
# PARSER XLS (xlrd) - formato antiguo
# --------------------------------------------------------------------------
def read_xls_rows(content_bytes):
    book = xlrd.open_workbook(file_contents=content_bytes)
    sheet = book.sheet_by_index(0)
    return [sheet.row_values(r) for r in range(sheet.nrows)]


def _find_data_start(rows, col):
    """Busca la primera fila con RFC valido (>=12 caracteres, no solo letras)
    en hasta 20 filas."""
    for i in range(min(20, len(rows))):
        row = rows[i]
        if col >= len(row):
            continue
        v = str(row[col] or "").strip()
        if len(v) >= 12 and not re.match(r"^[a-zA-Z\s]+$", v):
            return i
    return -1


def _parse_periodo(raw):
    """float -> int -> str, validado a 6 digitos. Si parece serial de fecha
    Excel (~5 digitos), se descarta (se intentara la columna siguiente)."""
    s = str(raw).strip()
    if s == "":
        return None
    if re.match(r"^\d+\.?\d*$", s):
        try:
            s = str(int(float(s)))
        except (ValueError, OverflowError):
            return None
    if len(s) == 6 and s.isdigit():
        return s
    return None


def _to_float(v):
    try:
        if v in ("", None):
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def _parse_cedular(rows):
    start = _find_data_start(rows, CED_RFC)
    if start < 0:
        return []
    out = []
    for row in rows[start:]:
        maxcol = max(CED_RFC, CED_CONTRIB, CED_PERIODO, CED_N, CED_H)
        if len(row) <= maxcol:
            continue
        rfc = str(row[CED_RFC] or "").strip().upper()
        if not rfc or len(rfc) < 12:
            continue
        periodo = _parse_periodo(row[CED_PERIODO])
        if not periodo and len(row) > CED_PERIODO + 1:
            # el indice de PERIODO puede estar desfasado en 1
            periodo = _parse_periodo(row[CED_PERIODO + 1])
        if not periodo:
            continue
        n_val = _to_float(row[CED_N])
        h_val = _to_float(row[CED_H])
        contrib = str(row[CED_CONTRIB] or "").strip()
        out.append({
            "rfc": rfc,
            "periodo": periodo,
            "recaudacion": n_val - h_val,
            "contrib": contrib,
        })
    return out


# --------------------------------------------------------------------------
# LOGICA DE PROYECCION Y OMISOS (equivalente a computeMonth en el HTML)
# --------------------------------------------------------------------------
def prev_period(p):
    p = str(p)
    y = int(p[:4])
    m = int(p[4:])
    m -= 1
    if m == 0:
        m = 12
        y -= 1
    return f"{y}{m:02d}"


def format_period(p):
    s = str(p)
    labels = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    try:
        return labels[int(s[4:6])] + "-" + s[2:4]
    except (ValueError, IndexError):
        return s


RFC_RE = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")


def is_valid_rfc(rfc):
    return bool(RFC_RE.match(rfc or ""))


def median(values):
    if not values:
        return 0
    s = sorted(values)
    n = len(s)
    mid = n // 2
    if n % 2 != 0:
        return s[mid]
    return (s[mid - 1] + s[mid]) / 2


def expected_period(month_num):
    """Periodo esperado, calculado deterministicamente (no inferido de los
    datos): en el archivo del mes M se espera la declaracion del mes anterior.
    Ej. mes vigente = Julio (7) -> periodo esperado = 202606."""
    return prev_period(f"2026{month_num:02d}")


def get_missing_periods(paid_set, dominant, max_back, stop_before=None):
    out = []
    p = str(dominant)
    while p not in paid_set:
        if stop_before and int(p) < int(stop_before):
            break
        out.append(p)
        p = prev_period(p)
        if len(out) >= max_back:
            break
    return out


def compute_month(month_num, all_month_data, metas):
    """Omisos cumplidores: RFC con >=2 meses anteriores pagados que NO
    aparece en el archivo del mes vigente."""
    cur = all_month_data.get(month_num, [])
    # Acumulado incluye TODOS los registros (RFC validos o no)
    acumulado = sum(r["recaudacion"] for r in cur)

    prev_months = [m for m in range(1, month_num) if all_month_data.get(m)]
    n_prev = len(prev_months)

    paid_this_month = {r["rfc"] for r in cur if r["rfc"]}

    rfc_months_paid = {}
    rfc_monthly_amount = {}
    rfc_periods_paid = {}
    rfc_contrib = {}
    for m in prev_months:
        for r in all_month_data.get(m, []):
            if not is_valid_rfc(r["rfc"]):
                continue  # RFC invalidos se excluyen del analisis de omisos
            rfc_months_paid.setdefault(r["rfc"], set()).add(m)
            monthly = rfc_monthly_amount.setdefault(r["rfc"], {})
            monthly[m] = monthly.get(m, 0) + r["recaudacion"]
            if r["periodo"]:
                rfc_periods_paid.setdefault(r["rfc"], set()).add(r["periodo"])
            if r["rfc"] not in rfc_contrib and r["contrib"]:
                rfc_contrib[r["rfc"]] = r["contrib"]

    candidates = [rfc for rfc, months in rfc_months_paid.items() if len(months) >= 2]
    periodo_esperado = expected_period(month_num)

    omisos = []
    for rfc in candidates:
        if rfc in paid_this_month:
            continue  # ya pago este mes
        cnt = len(rfc_months_paid[rfc])
        paid_periods = rfc_periods_paid.get(rfc, set())
        missing = get_missing_periods(paid_periods, periodo_esperado, 12, None)
        if not missing:
            missing = [periodo_esperado]

        amounts = list(rfc_monthly_amount.get(rfc, {}).values())
        median_val = median(amounts)
        avg = round(median_val * len(missing))

        if cnt == n_prev:
            seg = "alta"
        elif cnt >= math.floor(n_prev * 0.75):
            seg = "media"
        elif cnt >= 3:
            seg = "baja"
        else:
            seg = "seguimiento"

        omisos.append({
            "rfc": rfc,
            "contrib": rfc_contrib.get(rfc) or "",
            "count": cnt,
            "avg": avg,
            "nMissing": len(missing),
            "pending": [format_period(p) for p in missing],
            "seg": seg,
        })

    omisos.sort(key=lambda o: -o["avg"])
    esperado = sum(o["avg"] for o in omisos if o["seg"] in ("alta", "media"))
    proyeccion = acumulado + esperado
    meta = metas.get(month_num, 0)

    segments = {}
    for o in omisos:
        s = segments.setdefault(o["seg"], {"count": 0, "monto": 0, "omisos": []})
        s["count"] += 1
        s["monto"] += o["avg"]
        s["omisos"].append({
            "rfc": o["rfc"], "contrib": o["contrib"], "avg": o["avg"],
            "count": o["count"], "nMissing": o["nMissing"], "pending": o["pending"],
        })
    for s in segments.values():
        s["monto"] = round(s["monto"])
        s["omisos"].sort(key=lambda x: -x["avg"])

    return {
        "mes_label": MONTH_LABELS.get(month_num, str(month_num)),
        "mes_num": month_num,
        "meta": meta,
        "periodo_esperado": int(periodo_esperado),
        "ref_months": prev_months,
        "acumulado_real": round(acumulado),
        "total_omisos": len(omisos),
        "total_esperado": round(esperado),
        "proyeccion_cierre": round(proyeccion),
        "meta_cruzada": proyeccion >= meta,
        "pct_acumulado": (acumulado / meta * 100) if meta else 0,
        "pct_proyeccion": (proyeccion / meta * 100) if meta else 0,
        "segmentos": segments,
        "omisos": omisos[:5000],
    }


# --------------------------------------------------------------------------
# ACTUALIZACION DE index.html (proteccion contra regresiones)
# --------------------------------------------------------------------------
def _load_existing_all_data(html_content):
    m = re.search(r"let allData\s*=\s*(\{.*?\});", html_content, re.S)
    if not m:
        return {}
    try:
        data = json.loads(m.group(1))
    except (json.JSONDecodeError, ValueError):
        return {}
    # Filtrar claves invalidas: conservar solo "1"-"12"
    return {k: v for k, v in data.items() if str(k) in VALID_MONTH_KEYS}


def update_html(html_path, new_month_data):
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    existing = _load_existing_all_data(content)
    merged = dict(existing)

    for key, data in new_month_data.items():
        k = str(key)
        old = merged.get(k)
        new_acum = data.get("acumulado_real", 0)
        old_acum = old.get("acumulado_real", 0) if old else -1
        if old is None or new_acum > old_acum:
            merged[k] = data
            print(f"  Mes {k}: actualizado (acumulado_real {old_acum:,} -> {new_acum:,})")
        else:
            print(f"  Mes {k}: se conserva el dato existente (nuevo {new_acum:,} no supera {old_acum:,})")

    # Filtrar de nuevo por seguridad
    merged = {k: v for k, v in merged.items() if k in VALID_MONTH_KEYS}

    all_data_json = json.dumps(merged, ensure_ascii=False, separators=(",", ":"))
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    content_new = re.sub(
        r"let allData\s*=\s*\{.*?\};",
        lambda _: "let allData = " + all_data_json.replace("\\", "\\\\") + ";",
        content, count=1, flags=re.S,
    )
    content_new = re.sub(
        r"var lastUpdated\s*=\s*'.*?';",
        f"var lastUpdated = '{now_str}';",
        content_new, count=1, flags=re.S,
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content_new)
    print(f"HTML actualizado: {html_path}")


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
def main():
    if not API_KEY:
        print("ERROR: variable de entorno DRIVE_API_KEY no definida.")
        sys.exit(1)

    print("Listando archivos en Google Drive...")
    files = drive_list_files(FOLDER_ID, API_KEY)
    print(f"  {len(files)} archivos encontrados en la carpeta.")

    month_files = []
    for f in files:
        num = detect_month(f["name"])
        if num:
            month_files.append({"id": f["id"], "name": f["name"], "num": num})

    if not month_files:
        print("No se encontraron archivos con nombre de mes reconocible. Nada que hacer.")
        sys.exit(0)

    month_files.sort(key=lambda x: x["num"])

    all_month_data = {}
    for mf in month_files:
        print(f"Descargando '{mf['name']}' (mes {mf['num']} - {MONTH_LABELS.get(mf['num'])})...")
        try:
            content = drive_download(mf["id"], API_KEY)
            rows = read_xls_rows(content)
        except Exception as e:
            print(f"  ERROR leyendo '{mf['name']}': {e}")
            continue
        recs = _parse_cedular(rows)
        suma = sum(r["recaudacion"] for r in recs)
        print(f"  {len(recs)} registros parseados. Suma recaudacion = {suma:,.2f}")
        all_month_data.setdefault(mf["num"], [])
        all_month_data[mf["num"]].extend(recs)

    if not all_month_data:
        print("No se pudo parsear ningun archivo. Abortando sin tocar el HTML.")
        sys.exit(1)

    print("Calculando proyecciones y omisos...")
    new_month_data = {}
    for num in sorted(all_month_data.keys()):
        d = compute_month(num, all_month_data, METAS)
        new_month_data[str(num)] = d
        print(
            f"  Mes {num} ({d['mes_label']}): acumulado_real={d['acumulado_real']:,} "
            f"omisos={d['total_omisos']} esperado={d['total_esperado']:,} "
            f"proyeccion={d['proyeccion_cierre']:,} periodo_esperado={d['periodo_esperado']}"
        )

    update_html(HTML_FILE, new_month_data)


if __name__ == "__main__":
    main()
