# -*- mode: python ; coding: utf-8 -*-
import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)


block_cipher = None

added_files = [
         ('images', 'images' ),
         (path.join(site_packages,"pptx","templates"), "pptx/templates"),
         ]


a = Analysis(['main.py'],
             pathex=['/Users/shinamadamori/Documents/SF/python/slides_v1',
             '/Users/shinamadamori/Documents/SF/python/slides_v1/env/bin/python'],
             binaries=[],
             datas=added_files,
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='OSZ-PPT Converter',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )

app = BUNDLE(exe,
             name='OSZ-PPT Converter.app',
             icon='slides.ico',
             bundle_identifier=None)
