# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('xdlovelife.ico', '.'),
        ('README.md', '.'),
        ('chromedriver.exe', '.') if os.path.exists('chromedriver.exe') else None,
    ],
    hiddenimports=[
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.options',
        'selenium.webdriver.common.by',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.common.exceptions',
        'selenium.webdriver.common.action_chains',
        'pandas',
        'openpyxl',
        'win32com.client',
        'win32com',
        'requests',
        'zipfile',
        'json',
        'logging',
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

# 清理datas列表中的None值
a.datas = [(src, dst, type) for src, dst, type in a.datas if src is not None]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AliProductsImport',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='xdlovelife.ico',
    version='file_version_info.txt',
    uac_admin=True,  # 请求管理员权限
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'python*.dll',
        'VCRUNTIME140.dll',
        'MSVCP140.dll',
        'chrome*.exe',
        'chrome*.dll',
    ],
    name='AliProductsImport',
)

# 添加应用程序清单
import subprocess
import os

manifest = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    version="1.0.0.0"
    processorArchitecture="*"
    name="AliProductsImport"
    type="win32"
  />
  <description>阿里巴巴产品导入工具</description>
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
      <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
      <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
      <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
    </application>
  </compatibility>
  <application xmlns="urn:schemas-microsoft-com:asm.v3">
    <windowsSettings>
      <dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>
      <longPathAware xmlns="http://schemas.microsoft.com/SMI/2016/WindowsSettings">true</longPathAware>
    </windowsSettings>
  </application>
</assembly>
'''

manifest_path = os.path.join('build', 'AliProductsImport.manifest')
os.makedirs('build', exist_ok=True)

with open(manifest_path, 'w') as f:
    f.write(manifest)

# 如果存在mt.exe，使用它来嵌入清单
mt_paths = [
    r'C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\mt.exe',
    r'C:\Program Files (x86)\Windows Kits\10\bin\x64\mt.exe',
]

for mt_path in mt_paths:
    if os.path.exists(mt_path):
        try:
            subprocess.run([
                mt_path,
                '-manifest', manifest_path,
                '-outputresource:dist/AliProductsImport/AliProductsImport.exe;#1'
            ], check=True)
            break
        except:
            continue 