# -*- mode: python -*-

block_cipher = None
data_files = Tree(r'C:\Python35\Lib\site-packages\requests', prefix='requests')

a = Analysis(['outlook_sdplus.py'],
             pathex=['H:\Coding\github\Outlook-Manage-Engine'],
             binaries=None,
             datas=None,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          data_files,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='outlook_sdplus',
          debug=False,
          strip=False,
          upx=True,
          console=True )
