# -*- coding: utf-8 -*-
"""PrintCalc 打印计价工具 - 启动入口。

用法：
    python main.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from printcalc.gui.app import main as run
    run()


if __name__ == "__main__":
    main()
