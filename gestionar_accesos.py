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

# El Excel donde se gestionan las claves en limpio. Vive UN NIVEL ARRIBA, fuera
# de la carpeta del repo, y por eso no se publica: lo que se sube a Vercel es
# solo esta carpeta. Es a proposito y no hay que moverlo adentro.
CLAVES = os.path.join(os.path.dirname(AQUI), "CLAVES TABLERO.xlsx")
HOJA_CLAVES = "CLAVES"
COLS_CLAVES = ["Usuario", "Contrasena", "Cuenta", "Nombre", "Areas"]
ANCHOS_CLAVES = {"Usuario": 24, "Contrasena": 22, "Cuenta": 10, "Nombre": 24,
                 "Areas": 46}

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


SEP_CLAVES = "|"


def partir_claves(txt):
    """Un usuario puede tener varias claves validas, separadas con '|'.

    Sirve para que entre lo mismo con la clave larga de Power BI o con un PIN
    corto. Se eligio '|' y no ';' porque el ';' ya separa las areas y porque
    ninguna clave de la casa lo usa."""
    if isinstance(txt, (list, tuple)):
        partes = list(txt)
    else:
        partes = str(txt or "").split(SEP_CLAVES)
    return [p for p in (str(x).strip() for x in partes) if p]


def aplicar_clave(fila, clave):
    """Guarda una sal y un hash por cada clave valida, separados con ';'.

    Van en la misma celda: el base64 nunca trae ';', asi que no hace falta
    agregar columnas y las hojas viejas de una sola clave siguen andando."""
    salts, hashes = [], []
    for c in partir_claves(clave):
        s, h, it = hashear(c)
        salts.append(s)
        hashes.append(h)
    fila["Salt"] = ";".join(salts)
    fila["Hash"] = ";".join(hashes)
    fila["Iteraciones"] = str(ITERACIONES)


def sumar_clave(fila, clave):
    """Agrega una clave mas a las que el usuario ya tiene.

    Distinto de aplicar_clave, que reemplaza todas. Sirve para sumar un PIN
    corto sin tener que volver a escribir la clave larga."""
    s, h, it = hashear(clave)
    salts = [x for x in fila["Salt"].split(";") if x.strip()]
    hashes = [x for x in fila["Hash"].split(";") if x.strip()]
    # Las viejas se recalcularon con estas iteraciones o no validarian: si la
    # hoja venia con otro valor, se respeta el que ya estaba.
    if hashes and fila["Iteraciones"] and int(fila["Iteraciones"]) != it:
        s, h, _ = hashear(clave, iters=int(fila["Iteraciones"]))
    else:
        fila["Iteraciones"] = str(it)
    salts.append(s)
    hashes.append(h)
    fila["Salt"] = ";".join(salts)
    fila["Hash"] = ";".join(hashes)
    return len(hashes)


def fuera_del_repo(path):
    """El Excel de claves NO puede estar en la carpeta que se publica.

    Todo lo que hay en esa carpeta termina en Vercel via 'git add -A'. Un
    archivo con las contrasenas en limpio ahi adentro quedaria descargable
    desde internet. Antes que confiar en el .gitignore, directamente no se
    deja trabajar con un archivo que este adentro."""
    p = os.path.abspath(path)
    repo = os.path.abspath(AQUI) + os.sep
    if p.startswith(repo):
        print("ERROR: '%s' esta DENTRO de la carpeta que se publica." % p)
        print("  Ese archivo tiene las contrasenas en limpio: ahi adentro se")
        print("  subiria a Vercel y quedaria publico. Movelo afuera, por ej:")
        print("  " + CLAVES)
        return False
    return True


def crear_planilla_claves(path, filas):
    """Arma el Excel de claves con los usuarios que ya existen.

    La columna Contrasena queda VACIA: las escribe quien las sabe. Vacio
    significa 'no la toques', asi que se puede cargar de a una sin borrar
    las demas."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = HOJA_CLAVES
    ws.append(COLS_CLAVES)
    for c in range(1, len(COLS_CLAVES) + 1):
        cel = ws.cell(1, c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="1B4FE5")
        cel.alignment = Alignment(vertical="center")
    for f in filas:
        ws.append([f["Usuario"], "", f["Nombre"], f["Areas"]])
    for c, nombre in enumerate(COLS_CLAVES, start=1):
        ws.column_dimensions[get_column_letter(c)].width = ANCHOS_CLAVES[nombre]
    ws.freeze_panes = "A2"
    # Las instrucciones van en su propia hoja: al pie de CLAVES se leian como
    # si fueran usuarios y daban de alta dos entradas basura.
    wl = wb.create_sheet("LEEME")
    wl.column_dimensions["A"].width = 100
    for i, linea in enumerate([
            "COMO SE USA",
            "",
            "1. Escribi la contrasena en la columna Contrasena de la hoja CLAVES.",
            "2. Guarda y cerra este Excel.",
            "3. Doble clic en 'aplicar claves.bat', en la carpeta Links POWER BI.",
            "4. Publica el tablero con publicar.bat.",
            "",
            "Celda vacia = no se toca, asi que podes cargar de a una.",
            "En Areas: separa con ; , '*' es ver todo, 'NINGUNA' es no ver nada.",
            "",
            "NO muevas este archivo a la carpeta 'Links POWER BI'.",
            "Esa carpeta se publica entera en internet y las claves quedarian",
            "descargables por cualquiera. El script se niega a leerlo si esta ahi.",
    ], start=1):
        c = wl.cell(i, 1)
        c.value = linea
        if i == 1 or linea.startswith("NO muevas"):
            c.font = Font(bold=True)
    wb.save(path)


def leer_planilla_claves(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[HOJA_CLAVES] if HOJA_CLAVES in wb.sheetnames else wb.worksheets[0]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    cab = [str(c or "").strip().lower() for c in filas[0]]
    idx = {}
    for c in COLS_CLAVES:
        # "Contrasena" y "Contraseña" son lo mismo: la ñ se escapa sola.
        alias = [c.lower()] + (["contraseña", "clave", "password"]
                               if c == "Contrasena" else [])
        idx[c] = next((cab.index(a) for a in alias if a in cab), None)
    out = []
    for f in filas[1:]:
        d = {}
        for c in COLS_CLAVES:
            i = idx[c]
            v = f[i] if (i is not None and i < len(f)) else None
            d[c] = "" if v is None else str(v).strip()
        # Un usuario nunca lleva espacios: si los tiene es un comentario
        # suelto en la planilla, no una fila de verdad.
        if d["Usuario"] and " " not in d["Usuario"]:
            out.append(d)
    return out


def resolver_cuenta(valor):
    """La columna Cuenta dice de que cuenta de Power BI cuelga la persona.

    Se escribe el numero (1, 2, 4, 5) y se entiende 'administracion01' y
    compania; tambien se acepta el nombre completo, por si alguna vez cuelga
    de 'directorio' o de una cuenta que no siga esa numeracion."""
    v = str(valor or "").strip()
    if not v:
        return ""
    # Excel guarda los numeros como 2 o como 2.0 segun como se cargaron.
    try:
        return "administracion%02d" % int(float(v))
    except ValueError:
        return norm_usuario(v)


def aplicar_planilla(filas, pedidos):
    """Vuelca la planilla de claves sobre las filas de la hoja ACCESOS.

    Devuelve (claves_cambiadas, datos_cambiados). Campo vacio = no se toca,
    para poder cargar de a uno sin pisar el resto.

    Quien tenga Cuenta hereda las areas de esa cuenta, y se recalcula en cada
    publicacion: cambiando las areas de 'administracion02' cambian de una todos
    los que cuelgan del 2."""
    claves, datos = [], []
    # Las areas de cada cuenta madre, tal como quedan despues de esta pasada.
    madres = {}
    for p in pedidos:
        u = norm_usuario(p["Usuario"])
        if not resolver_cuenta(p.get("Cuenta")):
            madres[u] = p["Areas"]

    for p in pedidos:
        cuenta = resolver_cuenta(p.get("Cuenta"))
        if cuenta:
            heredado = madres.get(cuenta)
            if heredado is None:
                f0 = buscar(filas, cuenta)
                heredado = f0["Areas"] if f0 else None
            if heredado is None:
                print("  AVISO: '%s' cuelga de '%s', que no existe. Lo salteo."
                      % (norm_usuario(p["Usuario"]), cuenta))
                continue
            p = dict(p, Areas=heredado)
        u = norm_usuario(p["Usuario"])
        f = buscar(filas, u)
        if not f:
            f = {"Usuario": u, "Nombre": "", "Areas": "",
                 "Iteraciones": "", "Salt": "", "Hash": ""}
            filas.append(f)
            datos.append(u + " (nuevo)")
        if p["Contrasena"]:
            aplicar_clave(f, p["Contrasena"])
            claves.append(u)
        if p["Nombre"] and p["Nombre"] != f["Nombre"]:
            f["Nombre"] = p["Nombre"]
            datos.append(u + " nombre")
        # Las areas se comparan contra "" a proposito: dejar la celda vacia
        # es no tocar, y para sacarle todo se pone la palabra NINGUNA.
        # Con Cuenta es distinto: ahi manda la cuenta madre siempre, incluso
        # cuando no tiene ninguna area, o el heredero se quedaria con las que
        # tenia antes de colgarse de ella.
        if p["Areas"] or cuenta:
            nuevo = "" if norm(p["Areas"]) == "ninguna" else p["Areas"]
            if nuevo != f["Areas"]:
                f["Areas"] = nuevo
                datos.append(u + " areas")
    return claves, datos


def norm(s):
    return (s or "").strip().lower()


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
    ap.add_argument("--agregar-clave", metavar="USUARIO",
                    help="Suma una clave mas, sin tocar las que ya tiene.")
    ap.add_argument("--quitar-claves", metavar="USUARIO",
                    help="Deja solo la primera clave del usuario.")
    ap.add_argument("--crear-planilla", action="store_true",
                    help="Arma el Excel de claves, fuera de la carpeta publicada.")
    ap.add_argument("--desde-planilla", nargs="?", const=CLAVES, metavar="RUTA",
                    help="Aplica el Excel de claves sobre la hoja ACCESOS.")
    args = ap.parse_args()

    filas = leer(args.excel)

    if args.listar:
        if not filas:
            print("No hay hoja ACCESOS todavia. Corre --semilla.")
            return 0
        print("%-20s %-24s %-8s %s" % ("USUARIO", "NOMBRE", "CLAVE", "AREAS"))
        for f in filas:
            n = len([x for x in f["Hash"].split(";") if x.strip()])
            print("%-20s %-24s %-8s %s" % (
                f["Usuario"], f["Nombre"] or "-",
                ("si" if n == 1 else "si (%d)" % n) if n else "SIN",
                f["Areas"] or "(ninguna)"))
        sin = [f["Usuario"] for f in filas if not f["Hash"]]
        vac = [f["Usuario"] for f in filas if not f["Areas"]]
        if sin:
            print("")
            print("  Sin contrasena (no pueden entrar): " + ", ".join(sin))
        if vac:
            print("  Sin areas (entran y no ven nada): " + ", ".join(vac))
        return 0

    if args.agregar_clave:
        u = norm_usuario(args.agregar_clave)
        f = buscar(filas, u)
        if not f:
            print("No existe '%s'. Crealo con --set." % u)
            return 1
        if not f["Hash"]:
            print("'%s' no tiene ninguna clave todavia. Usa --set." % u)
            return 1
        clave = pedir_clave(u)
        if not clave:
            return 0
        n = sumar_clave(f, clave)
        escribir(args.excel, filas)
        print("'%s' ahora entra con %d claves distintas." % (u, n))
        print("Para que tome efecto en el tablero: publicar.bat")
        return 0

    if args.quitar_claves:
        u = norm_usuario(args.quitar_claves)
        f = buscar(filas, u)
        if not f:
            print("No existe '%s'." % u)
            return 1
        f["Salt"] = f["Salt"].split(";")[0]
        f["Hash"] = f["Hash"].split(";")[0]
        escribir(args.excel, filas)
        print("'%s' quedo con una sola clave: la primera que tenia." % u)
        return 0

    if args.crear_planilla:
        if not fuera_del_repo(CLAVES):
            return 4
        if os.path.isfile(CLAVES):
            print("Ya existe: " + CLAVES)
            print("  No lo piso para no borrarte lo que tengas cargado.")
            return 1
        if not filas:
            print("No hay hoja ACCESOS todavia. Corre --semilla primero.")
            return 1
        crear_planilla_claves(CLAVES, filas)
        print("Listo: " + CLAVES)
        print("  %d usuarios, con la columna Contrasena vacia." % len(filas))
        print("  Escribi las claves ahi y corre 'aplicar claves.bat'.")
        return 0

    if args.desde_planilla:
        ruta = args.desde_planilla
        if not fuera_del_repo(ruta):
            return 4
        if not os.path.isfile(ruta):
            print("ERROR: no encuentro '%s'." % ruta)
            print("  Crealo con: python gestionar_accesos.py --crear-planilla")
            return 2
        pedidos = leer_planilla_claves(ruta)
        if not pedidos:
            print("La planilla esta vacia.")
            return 1
        claves, datos = aplicar_planilla(filas, pedidos)
        if not claves and not datos:
            print("Nada para cambiar: la planilla ya esta aplicada.")
            return 0
        escribir(args.excel, filas)
        if claves:
            print("Claves actualizadas (%d): %s" % (len(claves), ", ".join(claves)))
        if datos:
            print("Datos actualizados: " + ", ".join(datos))
        print("")
        print("Para que tome efecto en el tablero: publicar.bat")
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
