# -*- coding: utf-8 -*-
"""
Administra la hoja ACCESOS de "LINKS POWER BI.xlsx": quien entra al tablero y
que areas ve.

La contrasena NUNCA se guarda. Se guarda un PBKDF2-HMAC-SHA256 con sal propia
por usuario; el navegador repite la misma cuenta y compara. Ni este script ni
el Excel ni el HTML pueden devolver la contrasena original.

OJO, y esto hay que tenerlo claro: el tablero es un sitio estatico. El filtro
por area es COSMETICO. El Excel con TODOS los links se publica igual y
cualquiera que sepa la URL lo puede bajar entero. Esto ordena la vista, no
oculta nada de verdad. Lo unico que protege los informes en serio es el login
de Power BI.

Uso
---
  python gestionar_accesos.py --listar
  python gestionar_accesos.py --set directorio
  python gestionar_accesos.py --set directorio --areas "DIRECTORIO;AREAS COMUNES"
  python gestionar_accesos.py --areas administracion03 "COMPRAS;AREAS COMUNES"
  python gestionar_accesos.py --borrar administracion07
  python gestionar_accesos.py --semilla          # crea los usuarios sin clave
  python gestionar_accesos.py --desde-archivo claves.txt

  --areas "*"  = ve todo.
  El archivo de --desde-archivo es "usuario<TAB>contrasena" por linea.
  BORRALO apenas termines: es el unico lugar donde quedan en texto plano.
"""
import argparse
import base64
import getpass
import hashlib
import os
import secrets
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(AQUI, "LINKS POWER BI.xlsx")

HOJA = "ACCESOS"
COLS = ["Usuario", "Nombre", "Areas", "Iteraciones", "Salt", "Hash"]
ANCHOS = {"Usuario": 20, "Nombre": 24, "Areas": 46, "Iteraciones": 12,
          "Salt": 26, "Hash": 46}
ITERACIONES = 200000

# Usuarios que existen hoy. Las areas arrancan en blanco a proposito: que las
# complete quien sabe quien tiene que ver que, y no un valor por defecto que
# termine dandole de mas a alguien.
SEMILLA = [
    ("admin",            "Administrador",     "*"),
    ("administracion01", "Administracion 01", ""),
    ("administracion02", "Administracion 02", ""),
    ("administracion03", "Administracion 03", ""),
    ("administracion04", "Administracion 04", ""),
    ("administracion05", "Administracion 05", ""),
    ("administracion06", "Administracion 06", ""),
    ("administracion07", "Administracion 07", ""),
    ("directorio",       "Directorio",        "DIRECTORIO;AREAS COMUNES"),
]


def norm_usuario(u):
    """'Directorio@danodis.onmicrosoft.com' y 'directorio' son el mismo."""
    u = (u or "").strip().lower()
    return u.split("@")[0] if "@" in u else u


def hashear(clave, salt=None, iters=ITERACIONES):
    salt = salt or secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), salt, iters, 32)
    return (base64.b64encode(salt).decode(),
            base64.b64encode(dk).decode(),
            iters)


def leer(path):
    if not os.path.isfile(path):
        print("ERROR: no encuentro '%s'." % path)
        sys.exit(2)
    wb = openpyxl.load_workbook(path)
    if HOJA not in wb.sheetnames:
        return []
    ws = wb[HOJA]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    cab = [str(c or "").strip().lower() for c in filas[0]]
    idx = {c: (cab.index(c.lower()) if c.lower() in cab else None) for c in COLS}
    out = []
    for f in filas[1:]:
        d = {}
        for c in COLS:
            i = idx[c]
            v = f[i] if (i is not None and i < len(f)) else None
            d[c] = "" if v is None else str(v).strip()
        if d["Usuario"]:
            out.append(d)
    return out


def escribir(path, filas):
    """Reescribe solo la hoja ACCESOS. Tableros y Catalogo no se tocan."""
    wb = openpyxl.load_workbook(path)
    if HOJA in wb.sheetnames:
        del wb[HOJA]
    ws = wb.create_sheet(HOJA)
    ws.append(COLS)
    for c in range(1, len(COLS) + 1):
        cel = ws.cell(1, c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="1B4FE5")
        cel.alignment = Alignment(vertical="center")
    for f in filas:
        ws.append([f.get(c, "") for c in COLS])
    for c, nombre in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = ANCHOS.get(nombre, 18)
    ws.freeze_panes = "A2"
    try:
        wb.save(path)
    except PermissionError:
        print("ERROR: no puedo escribir el Excel. Casi seguro lo tenes abierto.")
        sys.exit(3)


def buscar(filas, usuario):
    for f in filas:
        if norm_usuario(f["Usuario"]) == usuario:
            return f
    return None


def pedir_clave(usuario):
    a = getpass.getpass("Contrasena para '%s': " % usuario)
    if not a:
        print("  Vacia: no cambio la clave.")
        return None
    b = getpass.getpass("Repetila: ")
    if a != b:
        print("ERROR: no coinciden. No toque nada.")
        sys.exit(1)
    return a


def aplicar_clave(fila, clave):
    salt, h, it = hashear(clave)
    fila["Salt"], fila["Hash"], fila["Iteraciones"] = salt, h, str(it)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=EXCEL)
    ap.add_argument("--listar", action="store_true")
    ap.add_argument("--semilla", action="store_true",
                    help="Crea los usuarios conocidos, sin contrasena.")
    ap.add_argument("--set", metavar="USUARIO",
                    help="Crea o actualiza un usuario y pide la contrasena.")
    ap.add_argument("--areas", nargs="+", metavar="X",
                    help="Con --set: las areas. Solo: USUARIO \"A;B\".")
    ap.add_argument("--nombre", default=None)
    ap.add_argument("--borrar", metavar="USUARIO")
    ap.add_argument("--desde-archivo", metavar="RUTA",
                    help="Carga en lote: usuario<TAB>contrasena por linea.")
    args = ap.parse_args()

    filas = leer(args.excel)

    if args.listar:
        if not filas:
            print("No hay hoja ACCESOS todavia. Corre --semilla.")
            return 0
        print("%-20s %-24s %-8s %s" % ("USUARIO", "NOMBRE", "CLAVE", "AREAS"))
        for f in filas:
            print("%-20s %-24s %-8s %s" % (
                f["Usuario"], f["Nombre"] or "-",
                "si" if f["Hash"] else "SIN",
                f["Areas"] or "(ninguna)"))
        sin = [f["Usuario"] for f in filas if not f["Hash"]]
        vac = [f["Usuario"] for f in filas if not f["Areas"]]
        if sin:
            print("")
            print("  Sin contrasena (no pueden entrar): " + ", ".join(sin))
        if vac:
            print("  Sin areas (entran y no ven nada): " + ", ".join(vac))
        return 0

    if args.semilla:
        nuevos = 0
        for u, nom, ar in SEMILLA:
            if not buscar(filas, u):
                filas.append({"Usuario": u, "Nombre": nom, "Areas": ar,
                              "Iteraciones": "", "Salt": "", "Hash": ""})
                nuevos += 1
        escribir(args.excel, filas)
        print("Hoja ACCESOS lista. Usuarios nuevos: %d" % nuevos)
        print("Ahora poneles contrasena:  python gestionar_accesos.py --set USUARIO")
        return 0

    if args.borrar:
        u = norm_usuario(args.borrar)
        antes = len(filas)
        filas = [f for f in filas if norm_usuario(f["Usuario"]) != u]
        if len(filas) == antes:
            print("No existe '%s'." % u)
            return 1
        escribir(args.excel, filas)
        print("Borrado '%s'." % u)
        return 0

    if args.desde_archivo:
        if not os.path.isfile(args.desde_archivo):
            print("ERROR: no encuentro '%s'." % args.desde_archivo)
            return 2
        n = 0
        with open(args.desde_archivo, encoding="utf-8-sig") as fh:
            for linea in fh:
                linea = linea.rstrip("\r\n")
                if not linea.strip() or linea.lstrip().startswith("#"):
                    continue
                partes = linea.split("\t") if "\t" in linea else linea.split(None, 1)
                if len(partes) < 2:
                    print("  salteo (sin contrasena): " + linea.strip())
                    continue
                u, clave = norm_usuario(partes[0]), partes[1].strip()
                f = buscar(filas, u)
                if not f:
                    f = {"Usuario": u, "Nombre": "", "Areas": "",
                         "Iteraciones": "", "Salt": "", "Hash": ""}
                    filas.append(f)
                aplicar_clave(f, clave)
                n += 1
                print("  ok  " + u)
        escribir(args.excel, filas)
        print("")
        print("Listo: %d contrasenas cargadas." % n)
        print("BORRA AHORA '%s': es el unico lugar con las claves en texto plano."
              % args.desde_archivo)
        return 0

    if args.set:
        u = norm_usuario(args.set)
        f = buscar(filas, u)
        if not f:
            f = {"Usuario": u, "Nombre": "", "Areas": "",
                 "Iteraciones": "", "Salt": "", "Hash": ""}
            filas.append(f)
            print("Usuario nuevo: " + u)
        if args.nombre is not None:
            f["Nombre"] = args.nombre
        if args.areas:
            f["Areas"] = ";".join(args.areas)
        clave = pedir_clave(u)
        if clave:
            aplicar_clave(f, clave)
        escribir(args.excel, filas)
        print("Guardado '%s'  areas: %s" % (u, f["Areas"] or "(ninguna)"))
        return 0

    if args.areas and len(args.areas) == 2:
        u = norm_usuario(args.areas[0])
        f = buscar(filas, u)
        if not f:
            print("No existe '%s'. Crealo con --set." % u)
            return 1
        f["Areas"] = args.areas[1]
        escribir(args.excel, filas)
        print("'%s' ahora ve: %s" % (u, f["Areas"]))
        return 0

    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
