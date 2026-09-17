"""Build the self-contained Windows application from an explicit asset list."""
from pathlib import Path
import importlib.metadata
import shutil
import subprocess
import sys

if sys.platform != 'win32':
    raise SystemExit('Please build on Windows with 64-bit Python 3.12.')
root = Path(__file__).resolve().parent.parent
build = root / 'build'
licenses = build / 'licenses'
licenses.mkdir(parents=True, exist_ok=True)
shutil.copyfile(Path(sys.base_prefix) / 'LICENSE.txt', licenses / 'Python-LICENSE.txt')
for name in ['openpyxl', 'et_xmlfile', 'pyinstaller']:
    distribution = importlib.metadata.distribution(name)
    for file in distribution.files:
        if any(word in file.name.lower() for word in ('license', 'licence', 'copying')):
            if file.suffix.lower() in ('.txt', '.rst', '.python', ''):
                shutil.copyfile(distribution.locate_file(file), licenses / (name + '-' + file.name))

subprocess.run([
    sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--windowed', '--onedir',
    '--name', 'Shiwen', '--distpath', str(build / 'bundle'),
    '--workpath', str(build / 'freeze'), '--specpath', str(build),
    '--paths', str(root), '--add-data', str(root / 'dist') + ';dist',
    '--add-data', str(licenses) + ';licenses', str(root / 'desktop.py'),
], check=True)
print('Application:', build / 'bundle' / 'Shiwen' / 'Shiwen.exe')
