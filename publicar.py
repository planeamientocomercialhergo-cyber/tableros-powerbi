# -*- coding: utf-8 -*-
"""
Sincroniza los links de Power BI y publica el tablero, todo de una.

Hace, en orden:
  1. Corre sincronizar_powerbi.py contra "LINKS POWER BI.xlsx".
  2. Si algo cambio, hace commit y push.
  3. Vercel republica solo en unos segundos.

Si el paso 1 falla, NO publica nada: es mejor dejar el tablero como estaba que
subir un Excel a medio armar.

Uso
---
  python publicar.py                 # sincroniza y publica
  python publicar.py --solo-sync     # sincroniza y NO publica (para revisar antes)
  python publicar.py --dry-run       # no escribe ni publica: solo informa
  python publicar.py -m "texto"      # mensaje de commit propio
"""
import argparse
import os
import re
import subprocess
import sys
from datetime import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
SYNC = os.path.join(AQUI, "sincronizar_powerbi.py")


def git(*args, **kw):
    """Corre git en la carpeta del proyecto. Devuelve (codigo, salida)."""
    r = subprocess.run(("git",) + args, cwd=AQUI, capture_output=True,
                       text=True, encoding="utf-8", errors="replace", **kw)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def titulo(txt):
    print("")
    print("=" * 64)
    print("  " + txt)
    print("=" * 64)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--solo-sync", action="store_true",
                    help="Sincroniza pero no publica.")
    ap.add_argument("--dry-run", action="store_true",
                    help="No escribe el Excel ni publica.")
    ap.add_argument("-m", "--mensaje", default="",
                    help="Mensaje de commit. Por defecto lleva la fecha.")
    args = ap.parse_args()

    if not os.path.isfile(SYNC):
        print("ERROR: no encuentro sincronizar_powerbi.py al lado de este script.")
        return 1

    # ---- 1. sincronizar ----------------------------------------------------
    titulo("1/2  Sincronizando los links de Power BI")
    cmd = [sys.executable, SYNC] + (["--dry-run"] if args.dry_run else [])
    proc = subprocess.run(cmd, cwd=AQUI, capture_output=True,
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
    cod, _ = git("rev-parse", "--is-inside-work-tree")
    if cod != 0:
        print("ERROR: esta carpeta no es un repositorio git.")
        return 1

    cod, cambios = git("status", "--porcelain")
    if cod != 0:
        print("ERROR consultando git:\n" + cambios)
        return 1
    if not cambios.strip():
        print("No hay cambios para publicar: el tablero ya esta al dia.")
        return 0

    print("Cambios detectados:")
    cod, breve = git("status", "-s")
    print(breve.rstrip())

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
    mensaje = args.mensaje or "Actualizo tableros (%s) - %s" % (fecha, detalle)

    for paso, gargs in [("add", ("add", "-A")),
                        ("commit", ("commit", "-m", mensaje)),
                        ("push", ("push",))]:
        cod, out = git(*gargs)
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
