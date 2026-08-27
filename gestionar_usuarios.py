"""Administracion de usuarios de Auto Prizma Pro.

Uso:

    python gestionar_usuarios.py listar
    python gestionar_usuarios.py crear
    python gestionar_usuarios.py clave
    python gestionar_usuarios.py desactivar
    python gestionar_usuarios.py activar
    python gestionar_usuarios.py borrar

Todas las opciones piden los datos de forma interactiva.
La contrasena nunca se escribe en pantalla ni queda en el historial
de PowerShell, porque se pide con getpass.
"""

import getpass
import json
import os
import sys

import main as app


def _pedir_usuario(mensaje="Usuario: "):
    valor = input(mensaje).strip().lower()

    if not valor:
        print("Cancelado: no escribiste ningun usuario.")
        sys.exit(1)

    return valor


def _pedir_contrasena():
    primera = getpass.getpass("Contrasena (min 8 caracteres): ")

    if len(primera) < 8:
        print("La contrasena debe tener al menos 8 caracteres.")
        sys.exit(1)

    segunda = getpass.getpass("Repite la contrasena: ")

    if primera != segunda:
        print("Las contrasenas no coinciden.")
        sys.exit(1)

    return primera


def listar():
    usuarios = app._cargar_usuarios()

    if not usuarios:
        print("No hay usuarios creados todavia.")
        print("Crea el primero con: python gestionar_usuarios.py crear")
        return

    print()
    print(f"{'USUARIO':<24} {'NOMBRE':<26} {'ESTADO':<10} TOKEN GOOGLE")
    print("-" * 78)

    for clave in sorted(usuarios):
        registro = usuarios[clave]
        estado = "activo" if registro.get("activo", True) else "desactivado"
        tiene_token = os.path.isfile(app._ruta_token_google(clave))

        print(
            f"{clave:<24} {str(registro.get('nombre') or ''):<26} "
            f"{estado:<10} {'si' if tiene_token else 'no'}"
        )

    print()
    print(f"Total: {len(usuarios)} usuario(s).")


def crear():
    usuario = _pedir_usuario("Usuario nuevo (sin espacios): ")

    if usuario in app._cargar_usuarios():
        print(f"El usuario '{usuario}' ya existe. Usa 'clave' para cambiarle la contrasena.")
        sys.exit(1)

    nombre = input("Nombre para mostrar (Enter para usar el usuario): ").strip()
    contrasena = _pedir_contrasena()

    app._crear_usuario(usuario, contrasena, nombre)
    print(f"Usuario '{usuario}' creado.")


def clave():
    usuario = _pedir_usuario("Usuario al que le cambias la contrasena: ")
    usuarios = app._cargar_usuarios()

    if usuario not in usuarios:
        print(f"El usuario '{usuario}' no existe.")
        sys.exit(1)

    contrasena = _pedir_contrasena()
    app._crear_usuario(usuario, contrasena, usuarios[usuario].get("nombre"))
    print(f"Contrasena de '{usuario}' actualizada.")


def _cambiar_estado(activo):
    usuario = _pedir_usuario()

    with app.USUARIOS_LOCK:
        usuarios = app._cargar_usuarios()

        if usuario not in usuarios:
            print(f"El usuario '{usuario}' no existe.")
            sys.exit(1)

        usuarios[usuario]["activo"] = activo
        app._guardar_usuarios(usuarios)

    print(
        f"Usuario '{usuario}' "
        + ("activado." if activo else "desactivado. Sus sesiones dejan de servir.")
    )


def desactivar():
    _cambiar_estado(False)


def activar():
    _cambiar_estado(True)


def borrar():
    usuario = _pedir_usuario("Usuario a borrar: ")

    with app.USUARIOS_LOCK:
        usuarios = app._cargar_usuarios()

        if usuario not in usuarios:
            print(f"El usuario '{usuario}' no existe.")
            sys.exit(1)

        confirmacion = input(
            f"Escribe '{usuario}' otra vez para confirmar el borrado: "
        ).strip().lower()

        if confirmacion != usuario:
            print("Cancelado.")
            sys.exit(1)

        usuarios.pop(usuario, None)
        app._guardar_usuarios(usuarios)

    ruta_token = app._ruta_token_google(usuario)

    if os.path.isfile(ruta_token):
        os.remove(ruta_token)
        print("Tambien se borro su token de Google.")

    print(f"Usuario '{usuario}' borrado.")
    print(
        "Nota: sus cargues del historial siguen guardados en "
        "data/historial.json, pero ya nadie los vera en la web."
    )


ACCIONES = {
    "listar": listar,
    "crear": crear,
    "clave": clave,
    "desactivar": desactivar,
    "activar": activar,
    "borrar": borrar,
}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in ACCIONES:
        print(__doc__)
        sys.exit(1)

    ACCIONES[sys.argv[1]]()


if __name__ == "__main__":
    main()
