import os
import sys

# Set UTF-8 encoding globally via environment variable (works in windowed mode)
os.environ["PYTHONIOENCODING"] = "utf-8"

import streamlit.web.cli as stcli

def get_base_path():
    """ Get path to resource, works for dev and PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

if __name__ == '__main__':
    # Safely reconfigure stdout/stderr only if console exists
    if sys.stdout is not None:
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr is not None:
        sys.stderr.reconfigure(encoding='utf-8')

    base_dir = get_base_path()
    script_path = os.path.join(base_dir, 'app.py')

    sys.argv = [
        "streamlit",
        "run",
        script_path,
        "--global.developmentMode=false"
    ]

    sys.exit(stcli.main())