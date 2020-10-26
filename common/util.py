import os
import sys

HOST = "imap.gmail.com"
USERNAME = "PLEASE_CHANGE_HERE"
PASSWORD = "PLEASE_CHANGE_HERE"


def OS_Checker():
    from sys import platform

    platform_str = str(platform).lower()
    if platform_str == "linux" or platform_str == "linux2":
        os = "linux"
    elif platform_str == "darwin":
        os = "darwin"
    elif platform_str == "win32":
        os = "win"

    return os