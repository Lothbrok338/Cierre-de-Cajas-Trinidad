#!/usr/bin/env python3
"""Genera el hash pbkdf2_sha256 de una clave para APP_USERS.
Uso: python scripts/hash_clave.py            (pide la clave sin mostrarla)
     python scripts/hash_clave.py --json usuario "Nombre Visible" cajero   (imprime la entrada JSON lista para APP_USERS)"""
import getpass, json, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from app.auth import hash_clave, ROLES


def main():
    args = sys.argv[1:]
    clave = getpass.getpass('Clave: ')
    if clave != getpass.getpass('Repita la clave: '): sys.exit('Las claves no coinciden.')
    if len(clave) < 8: sys.exit('Use al menos 8 caracteres.')
    h = hash_clave(clave)
    if args and args[0] == '--json':
        if len(args) != 4 or args[3] not in ROLES: sys.exit(f'Uso: --json usuario "Nombre" {"|".join(ROLES)}')
        print(json.dumps({args[1]: {'nombre': args[2], 'rol': args[3], 'clave_hash': h}}, ensure_ascii=False))
    else:
        print(h)


if __name__ == '__main__':
    main()
