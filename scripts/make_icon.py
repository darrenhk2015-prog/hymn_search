#!/usr/bin/env python3
"""Build assets/app.ico from assets/app.png (run once before packaging)."""
import os
import sys

from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QApplication

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, 'assets')


def main():
    app = QApplication(sys.argv)
    png = os.path.join(ASSETS, 'app.png')
    if not os.path.isfile(png):
        print(f'Missing {png}')
        return 1
    pm = QPixmap(png)
    if pm.isNull():
        print('Failed to load app.png')
        return 1
    ico = os.path.join(ASSETS, 'app.ico')
    if not pm.save(ico, 'ICO'):
        print('Failed to write app.ico')
        return 1
    print(f'Wrote {ico}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
