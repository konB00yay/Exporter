# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['Exporter.py'],
             pathex=['/usr/local/lib/python3.9/site-packages', '/Users/konnorbeaulier/PycharmProjects/ExcelImageExporter'],
             binaries=[('libMagick++-7.Q16HDRI.4.dylib', '.'), ('libMagickCore-7.Q16HDRI.8.dylib', '.'), ('libMagickWand-7.Q16HDRI.8.dylib', '.'), ],
             datas=[],
             hiddenimports=['google-api-core', 'google-api-python-client', 'apiclient'],
             hookspath=[],
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
          name='Exporter',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
