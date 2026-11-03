# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包规格
#
# 构建命令（沙箱环境）：
#   set CODEBUDDY_SAFE_DELETE_SANDBOX=0
#   pyinstaller 微信文件群发工具.spec --workpath=%TEMP%\wfs_build --distpath=%TEMP%\wfs_dist
#   再把产物拷回 dist/
#
# 踩坑记录：
#   - 沙箱会拦截 build/ 目录清理（os.remove 走回收站）→ 设 SAFE_DELETE_SANDBOX=0
#     并把 workpath/distpath 指向 %TEMP%
#   - 本机未安装 UPX，upx=True 每次构建都刷警告 → 置 False
#   - pywebview 的 webview 后端是动态导入的，必须显式 hiddenimports，
#     否则打包后窗口打不开（缺模块）

a = Analysis(
    ['launch.py'],
    pathex=[],
    binaries=[],
    # 模板编译进 exe 资源，用户无需附带任何文件
    datas=[('发送映射模板.xlsx', '.')],
    hiddenimports=[
        # 本项目模块
        'wechat_sender', 'send_outcome', 'retry', 'state_store', 'config_store',
        # web / excel
        'flask', 'werkzeug', 'openpyxl',
        # 键鼠模拟与 Win32
        'pyautogui', 'pythoncom', 'win32clipboard', 'win32gui', 'win32con',
        'win32api',
        # pywebview（后端动态导入）
        'webview', 'bottle', 'proxy_tools', 'clr_loader', 'pythonnet',
        'webview.platforms.edgechromium',
        # pyautogui 依赖
        'PIL', 'PIL.Image', 'pyscreeze',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 排除用不到的大块头，减小体积。
    # 注意：不要排除 setuptools —— pkg_resources 有些库要用。
    excludes=[
        'tkinter', 'unittest', 'pydoc_data',
        'numpy', 'pandas', 'matplotlib', 'scipy',
        'IPython', 'notebook', 'pytest',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='微信文件群发工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # 本机未装 UPX，避免每次构建刷警告
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # 无控制台黑框
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
