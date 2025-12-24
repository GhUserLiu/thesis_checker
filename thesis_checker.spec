# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['checker.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 包含数据文件夹（如果需要）
        ('原始数据', '原始数据'),
        ('二次提交', '二次提交'),
    ],
    hiddenimports=[
        'jieba',
        'sklearn',
        'pandas',
        'openpyxl',
        'numpy',
        'scipy',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='论文选题查重系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加 .ico 图标文件路径
)
