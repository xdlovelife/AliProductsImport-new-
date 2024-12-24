import PyInstaller.__main__
import os
import shutil

def build_exe():
    # 清理之前的构建文件
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # PyInstaller参数
    args = [
        'main.py',  # 主程序文件
        '--name=AliProductsImport',  # 生成的exe名称
        '--windowed',  # 使用GUI模式
        '--icon=xdlovelife.ico',  # 应用图标
        '--add-data=xdlovelife.ico;.',  # 添加图标文件
        '--noconfirm',  # 不询问确认
        '--clean',  # 清理临时文件
        '--hidden-import=PyQt6',
        '--hidden-import=selenium',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--collect-data=selenium',
        '--collect-data=PyQt6',
    ]
    
    # 运行PyInstaller
    PyInstaller.__main__.run(args)
    
    # 复制额外文件到dist目录
    dist_dir = os.path.join('dist', 'AliProductsImport')
    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)
    
    # 复制图标文件
    shutil.copy2('xdlovelife.ico', dist_dir)
    
    print("打包完成！")

if __name__ == '__main__':
    build_exe() 