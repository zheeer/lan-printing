# -*- mode: python ; coding: utf-8 -*-
"""

"""

import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

script = 'print_server2.6.py'
pathex = [os.path.abspath('.')]

hiddenimports = [
    'comtypes',
    'comtypes.client',
    'comtypes.stream',
    'win32com',
    'win32com.client',
    'win32api',
    'win32print',
    'win32con',
    'winreg',
    'wmi',
    'ctypes',
    'ctypes.wintypes',
    'requests',
    'requests_toolbelt',
    'urllib3',
    'certifi',
    'chardet',
    'flask',
    'flask_cors',
    'flask.json',
    'flask.helpers',
    'requests',
    'werkzeug',
    'waitress',
    'pdf2image',
    'pystray',
    'pystray._base',
    'pystray._win32',
    'PIL',
    'PIL.Image',
    'PIL.ImageDraw',
    'PIL.ImageFont',
    'PIL.ImageOps',
    'parse',
]
hiddenimports += collect_submodules('pystray') if os.path.isdir(os.path.join(pathex[0], 'pystray')) else []

datas = []
def add_if_exists(src, dest=None):
    if not dest:
        dest = os.path.basename(src)
    if os.path.exists(src):
        if os.path.isdir(src):
            for root, _, files in os.walk(src):
                for f in files:
                    full = os.path.join(root, f)
                    rel = os.path.relpath(full, src)
                    datas.append((full, os.path.join(dest, rel)))
        else:
            datas.append((src, dest))

add_if_exists(os.path.join('.', 'uploads'), 'uploads')
add_if_exists(os.path.join('.', 'scanned_files'), 'scanned_files')
add_if_exists(os.path.join('.', 'logo.ico'), '.')
add_if_exists(os.path.join('.', 'bootstrap.min.css'), '.')
add_if_exists(os.path.join('.', 'bootstrap.bundle.min.js'), '.')

poppler_dir = os.path.join('.', 'poppler', 'Library', 'bin')

poppler_binaries = []
if os.path.isdir(poppler_dir):
    for root, _, files in os.walk(poppler_dir):
        for f in files:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, poppler_dir)
            dest = os.path.join('poppler_bin', rel)
            poppler_binaries.append((full, dest))

a = Analysis([script],
             pathex=pathex,
             binaries=poppler_binaries,
             datas=datas,
             hiddenimports=hiddenimports,
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='print_server2.6',
          debug=False,
          strip=False,
          upx=True,
          console=True,
          icon='logo.ico' if os.path.exists('logo.ico') else None)

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='print_server2.6')
