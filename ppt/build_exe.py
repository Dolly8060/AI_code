"""打包 CSNDC PPT Converter 为exe

执行: python build_exe.py
输出: dist/CSNDC_PPT_Converter.exe
"""

import subprocess
import sys
import os

def main():
    print("=" * 50)
    print("CSNDC PPT Converter 打包工具")
    print("=" * 50)
    
    # 检查pyinstaller
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__}")
    except ImportError:
        print("安装 PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    # 检查依赖
    deps = ["beautifulsoup4", "python-pptx", "playwright", "Pillow"]
    for dep in deps:
        try:
            __import__(dep.replace("-", "_").split("4")[0])
            print(f"✓ {dep}")
        except ImportError:
            print(f"安装 {dep}...")
            subprocess.run([sys.executable, "-m", "pip", "install", dep], check=True)
    
    # 打包
    print("\n开始打包...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",           # 单文件
        "--windowed",          # 无控制台窗口
        "--name", "CSNDC_PPT_Converter",  # exe名称
        "--clean",             # 清理临时文件
        "html2ppt_app.py"
    ]
    
    result = subprocess.run(cmd, cwd=os.path.dirname(__file__) or ".")
    
    if result.returncode == 0:
        print("\n" + "=" * 50)
        print("✓ 打包完成!")
        print("  输出: dist/CSNDC_PPT_Converter.exe")
        print("=" * 50)
        print("\n注意: 用户电脑需要安装Edge或Chrome浏览器")
    else:
        print("\n✗ 打包失败")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
