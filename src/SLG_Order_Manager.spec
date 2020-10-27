# -*- mode: python -*-

block_cipher = None


a = Analysis(['SLG_Order_Manager.py'],
             pathex=['C:\\Users\\User\\Dropbox\\-3DMICH-\\Clients\\SLG\\SLG_APP_WIN_V2'],
             binaries=[],
             datas=[],
             hiddenimports=['comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_0_1_0', 'comtypes.gen.UIAutomationClient'],
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
          name='SLG_Order_Manager',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
