"""Herramienta de administración LOCAL (se ejecuta en el computador donde corre el sistema).

  python restablecer_clave.py                       -> lista los usuarios registrados
  python restablecer_clave.py "juan esteban"        -> te pide una contraseña nueva para ese usuario
  python restablecer_clave.py --correo tu@mail.com  -> envía un correo de prueba y muestra el error exacto

No requiere correo ni clave maestra: quien tiene acceso a esta carpeta es el administrador.
"""
import argparse
import getpass
import sys

from sqlalchemy import func

import main
import models
from database import SessionLocal


def listar(db):
    filas = db.query(models.Asesor).order_by(models.Asesor.id).all()
    if not filas:
        print("No hay usuarios registrados todavía.")
    for a in filas:
        print(f"[{a.id}] nombre: {a.nombre!r} | usuario: {a.username!r} | correo: {a.email!r} | pedidos: {len(a.pedidos)}")


def buscar(db, texto):
    n = models.normalizar(texto)
    exactas = db.query(models.Asesor).filter(
        (models.Asesor.username == n) | (models.Asesor.nombre_norm == n)
        | (func.lower(models.Asesor.email) == texto.strip().lower())).all()
    if exactas:
        return exactas
    # si no hay coincidencia exacta, acepta parte del usuario o del nombre
    return db.query(models.Asesor).filter(
        models.Asesor.username.contains(n) | models.Asesor.nombre_norm.contains(n)).all()


def main_cli():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("usuario", nargs="?", help="usuario, nombre o correo")
    ap.add_argument("--clave", help="contraseña nueva (si se omite, se pide sin mostrarla)")
    ap.add_argument("--correo", metavar="DESTINO", help="enviar un correo de prueba")
    args = ap.parse_args()

    if args.correo:
        cfg = main.config_correo()
        print(f"Servidor: {cfg.get('host')}:{cfg.get('port', 587)} | usuario: {cfg.get('user')}")
        ok = main.enviar_correo(args.correo, "Prueba de correo · OMS", "Si lees esto, el envío funciona.")
        print("✔ Correo enviado. Revisa la bandeja (y spam)." if ok else "✘ No se envió. Revisa el mensaje de error de arriba.")
        return

    db = SessionLocal()
    if not args.usuario:
        listar(db)
        return
    cuentas = buscar(db, args.usuario)
    if not cuentas:
        print(f"No encontré «{args.usuario}». Usuarios existentes:")
        listar(db)
        sys.exit(1)
    if len(cuentas) > 1:
        print("Hay varias coincidencias, sé más específico:")
        for a in cuentas:
            print(f"  - {a.nombre} (usuario: {a.username})")
        sys.exit(1)
    a = cuentas[0]

    clave = args.clave
    if not clave:
        clave = getpass.getpass(f"Nueva contraseña para {a.nombre}: ")
        if clave != getpass.getpass("Repítela: "):
            sys.exit("Las contraseñas no coinciden.")
    if len(clave) < 8:
        sys.exit("La contraseña debe tener al menos 8 caracteres.")

    a.password_hash = main.hash_clave(clave)
    a.reset_token_hash, a.reset_expira = None, None
    if not a.username:                      # asesor antiguo sin usuario
        a.username = a.nombre_norm
    db.commit()
    print(f"✔ Contraseña actualizada. Inicia sesión con el usuario: {a.username}")


if __name__ == "__main__":
    main_cli()
