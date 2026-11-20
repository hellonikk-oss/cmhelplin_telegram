import os
import sys
import runpy

if __name__ == "__main__":
    code_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cmhelpline_code")
    if code_dir not in sys.path:
        sys.path.insert(0, code_dir)
    target = os.path.join(code_dir, "cmhlpline_all_complain_download.py")
    runpy.run_path(target, run_name="__main__")
