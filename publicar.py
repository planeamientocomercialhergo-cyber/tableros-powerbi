# -*- coding: utf-8 -*-
"""
Sincroniza los links de Power BI y publica el tablero, todo de una.

Hace, en orden:
  1. Corre sincronizar_powerbi.py contra "LINKS POWER BI.xlsx".
  2. Si algo cambio, hace commit y push.
  3. Vercel republica solo en unos segundos.

Si el paso 1 falla, NO publica nada: es mejor dejar el tablero como estaba que
subir un Excel a medio armar.

Se puede correr desde cualquier carpeta (por ejemplo desde el orquestador): la
carpeta del proyecto se resuelve sola, y antes de tocar git se verifica que el
repo sea el del tablero. Si no lo es, corta. Asi no hay forma de commitear por
error en otro repositorio.

Uso
---
  python publicar.py                 # sincroniza y publica
  python publicar.py --solo-sync     # sincroniza y NO publica (para revisar antes)
  python publicar.py --dry-run       # no escribe ni publica: solo informa
  python publicar.py -m "texto"      # mensaje de commit propio
  python publicar.py --proyecto RUTA # apunta a otra copia del proyecto
"""
import argparse
import os
import re
import subprocess
import sys
from datetime import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))

# Ruta UNC, no "H:\...": el orquestador puede correr bajo una cuenta que no
# tenga la unidad mapeada. H: es \\layla\Documentos.
PROYECTO_DEFECTO = (r"\\layla\Documentos\Automatizacion de reportes"
                    r"\PowerBI\Links POWER BI")

# Salvaguarda: el repo del tablero y ninguno mas.
REPO_ESPERADO = "tableros-powerbi"


def resolver_proyecto(indicado):
    """Devuelve la carpeta del proyecto, probando en orden:
    --proyecto, la variable LINKS_PBI_DIR, la carpeta de este script, y por
    ultimo la ruta UNC de siempre."""
    for cand in (indicado,
                 os.environ.get("LINKS_PBI_DIR"),
                 AQUI,
                 PROYECTO_DEFECTO):
        if cand and os.path.isfile(os.path.join(cand, "sincronizar_powerbi.py")):
            return os.path.abspath(cand)
    return None


def git(carpeta, *args):
    """Corre git en la carpeta del proyecto. Devuelve (codigo, salida)."""
    r = subprocess.run(("git",) + args, cwd=carpeta, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def titulo(txt):
    print("")
    print("=" * 64)
    print("  " + txt)
    print("=" * 64)


def repo_correcto(carpeta):
    """True si esa carpeta es el repo del tablero. Evita commitear en otro."""
    cod, raiz = git(carpeta, "rev-parse", "--show-toplevel")
    if cod != 0:
        print("ERROR: '%s' no es un repositorio git." % carpeta)
        return False
    cod, remoto = git(carpeta, "remote", "get-url", "origin")
    if cod != 0:
        print("ERROR: el repo no tiene remoto 'origin'.")
        return False
    if REPO_ESPERADO not in remoto:
        print("ERROR: me niego a publicar, el repo no es el del tablero.")
        print("  carpeta : " + raiz.strip())
        print("  origin  : " + remoto.strip())
        print("  esperaba: algo que contenga '%s'" % REPO_ESPERADO)
        return False
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--solo-sync", action="store_true",
                    help="Sincroniza pero no publica.")
    ap.add_argument("--dry-run", action="store_true",
                    help="No escribe el Excel ni publica.")
    ap.add_argument("-m", "--mensaje", default="",
                    help="Mensaje de commit. Por defecto lleva la fecha.")
    ap.add_argument("--proyecto", default="",
                    help="Carpeta del proyecto, si no es la de este script.")
    args = ap.parse_args()

    proyecto = resolver_proyecto(args.proyecto)
    if not proyecto:
        print("ERROR: no encuentro la carpeta del proyecto.")
        print("  Busque 'sincronizar_powerbi.py' en:")
        for c in (args.proyecto, os.environ.get("LINKS_PBI_DIR"),
                  AQUI, PROYECTO_DEFECTO):
            if c:
                print("    - " + c)
        print("  Pasa la ruta con --proyecto o la variable LINKS_PBI_DIR.")
        return 1
    print("Proyecto: " + proyecto)

    # ---- 1. sincronizar ----------------------------------------------------
    titulo("1/2  Sincronizando los links de Power BI")
    cmd = [sys.executable, os.path.join(proyecto, "sincronizar_powerbi.py")]
    if args.dry_run:
        cmd.append("--dry-run")
    proc = subprocess.run(cmd, cwd=proyecto, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")
    salida = (proc.stdout or "") + (proc.stderr or "")
    print(salida.rstrip())
    if proc.returncode != 0:
        titulo("Se corto: la sincronizacion fallo, no publico nada")
        return proc.returncode

    if args.dry_run:
        titulo("DRY-RUN: no escribi el Excel ni publique nada")
        return 0

    # numeros del resumen, para que el commit diga algo util
    def dato(patron):
        m = re.search(patron + r"[^\d]*(\d+)", salida)
        return m.group(1) if m else "?"
    detalle = "%s informes de Power BI y %s links a mano" % (
        dato("Publicados"), dato("Links a mano preservados"))

    if args.solo_sync:
        titulo("Listo el Excel. No publique (--solo-sync)")
        print("  Cuando quieras publicar: python publicar.py")
        return 0

    # ---- 2. publicar -------------------------------------------------------
    titulo("2/2  Publicando en Vercel")
    if not repo_correcto(proyecto):
        return 1

    cod, cambios = git(proyecto, "status", "--porcelain")
    if cod != 0:
        print("ERROR consultando git:\n" + cambios)
        return 1
    if not cambios.strip():
        print("No hay cambios para publicar: el tablero ya esta al dia.")
        return 0

    print("Cambios detectados:")
    cod, breve = git(proyecto, "status", "-s")
    print(breve.rstrip())

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
    mensaje = args.mensaje or "Actualizo tableros (%s) - %s" % (fecha, detalle)

    for paso, gargs in [("add", ("add", "-A")),
                        ("commit", ("commit", "-m", mensaje)),
                        ("push", ("push",))]:
        cod, out = git(proyecto, *gargs)
        if cod != 0:
            print("")
            print("ERROR en git %s:" % paso)
            print(out.rstrip())
            if paso == "push":
                print("")
                print("El commit quedo hecho local. Cuando se resuelva, "
                      "corre: git push")
            return 1
        if out.strip():
            print(out.rstrip())

    titulo("Publicado. Vercel republica en unos segundos.")
    print("  " + mensaje)
    return 0


if __name__ == "__main__":
    sys.exit(main())
