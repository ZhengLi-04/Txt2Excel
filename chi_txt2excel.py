#!/usr/bin/env python
"""兼容入口：打开统一工具，并默认使用 CHI 模式。"""

from txt_excel_studio import launch_app


if __name__ == "__main__":
    launch_app(default_profile="chi")
