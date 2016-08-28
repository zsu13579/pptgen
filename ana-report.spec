# -*- mode: python -*-
a = Analysis(['ana-report.py'],
             pathex=['E:\\djangoWeb\\pptgen'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='ana-report.exe',
          debug=False,
          strip=None,
          upx=True,
          console=True )
