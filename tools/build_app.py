import PyInstaller.__main__
import shutil
import os
import sys
from pathlib import Path

def build():
    # Ensure we are in the project root
    project_root = Path(__file__).resolve().parent.parent
    os.chdir(project_root)
    
    print(f"🚀 Starting build process in: {project_root}")

    # 1. Clean previous builds
    print("🧹 Cleaning build artifacts...")
    for folder in ['build', 'dist']:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"   - Removed {folder}/")
            except Exception as e:
                print(f"   ! Failed to remove {folder}/: {e}")

    # 2. Run PyInstaller
    print("📦 Running PyInstaller...")
    try:
        PyInstaller.__main__.run([
            'MotorReportApp.spec',
            '--clean',
            '--noconfirm',
            '--log-level=WARN'
        ])
        print("✅ Build complete!")
    except Exception as e:
        print(f"❌ Build failed: {e}")
        sys.exit(1)

    # 3. Verify Output
    # Read version to determine expected filename
    try:
        sys.path.insert(0, str(project_root))
        from src._version import VERSION
        exe_name = f"MotorReportApp-{VERSION}.exe"
    except ImportError:
        print("⚠️  Could not read version, assuming default name.")
        exe_name = "MotorReportApp.exe"

    exe_path = project_root / 'dist' / exe_name
    
    if exe_path.exists():
        print(f"🎉 Single-file executable created at: {exe_path}")
    else:
        print(f"⚠️  Build finished but executable not found at: {exe_path}")
        # List contents of dist to help debug
        dist_dir = project_root / 'dist'
        if dist_dir.exists():
            print(f"   Contents of {dist_dir}:")
            for item in dist_dir.iterdir():
                print(f"   - {item.name}")

if __name__ == "__main__":
    build()
