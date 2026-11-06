"""Create a local venv or run the application/tests without PowerShell."""
from pathlib import Path
import subprocess
import sys


def main():
    if sys.platform != 'win32' or sys.version_info[:2] != (3, 12):
        print('Windows Python 3.12 is required.', file=sys.stderr)
        return 1
    base = Path(__file__).resolve().parent
    python = base / '.venv' / 'Scripts' / 'python.exe'
    command = sys.argv[1] if len(sys.argv) > 1 else 'start'
    try:
        if command == 'setup':
            if not python.is_file():
                subprocess.run([sys.executable, '-m', 'venv', str(base / '.venv')], check=True)
            subprocess.run([str(python), '-m', 'pip', 'install', '-r', str(base / 'requirements.txt')], check=True)
            subprocess.run([str(python), '-c', 'import pythoncom, win32com.client; print("Setup completed.")'], check=True)
            return 0
        if not python.is_file():
            print('Run Setup.cmd first.', file=sys.stderr)
            return 1
        if command == 'test':
            return subprocess.call([str(python), '-m', 'unittest', 'discover', '-s', str(base), '-p', 'test_*.py', '-v'])
        if command == 'start':
            return subprocess.call([str(python), str(base / 'excel_search.py'), *sys.argv[2:]])
        raise ValueError(f'Unknown command: {command}')
    except (OSError, subprocess.CalledProcessError, ValueError) as exc:
        print(f'Error: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
