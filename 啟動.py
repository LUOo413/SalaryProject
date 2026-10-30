# -*- coding: utf-8 -*-
"""
啟動程式（打包成 exe 用的）
==========================
這支只負責「打開旁邊的 清冊共用.py」，本身不放任何清冊的邏輯。
所以：
  ・exe 只要轉一次
  ・以後改程式 → 換掉 exe 旁邊的 清冊共用.py 或 客戶模組 裡的 .py 就好，不用重新轉 exe

資料夾要長這樣（exe 跟 清冊共用.py 放在一起）：
  薪資印領清冊.exe      ← 這支打包出來的
  清冊共用.py
  客戶模組\
  客戶設定\            （第一次執行後自動產生）

打包（只要做一次；Windows 在命令提示字元、Mac 在終端機，先 cd 到這個資料夾）：
  Windows：pip install pyinstaller openpyxl xlrd rapidocr_onnxruntime pypdfium2 opencc-python-reimplemented
           pyinstaller --onefile --noconsole --collect-all rapidocr_onnxruntime --collect-all opencc --name 薪資印領清冊 啟動.py
  Mac：    python3 -m pip install --break-system-packages pyinstaller openpyxl xlrd rapidocr_onnxruntime pypdfium2 opencc-python-reimplemented
           pyinstaller --onefile --windowed --collect-all rapidocr_onnxruntime --collect-all opencc --name 薪資印領清冊 啟動.py
  做好的檔在 dist 資料夾裡，把它搬到 清冊共用.py 旁邊。
"""
import os
import sys

# ---- 下面這些 import 不是這支要用的，是要讓打包工具把它們一起包進 exe ----
#   清冊共用.py、客戶模組 是 exe 打開後才讀的，打包工具看不到它們用了什麼，
#   所以在這裡先 import 一次。以後 .py 裡用到新的套件，要加在這裡、再轉一次 exe。
import copy, csv, codecs, collections, datetime, decimal, difflib, functools, glob, importlib, importlib.util  # noqa
import itertools, json, math, re, shutil, statistics, string, traceback, types, unicodedata, zipfile  # noqa
import tkinter, tkinter.ttk, tkinter.font, tkinter.filedialog, tkinter.messagebox, tkinter.simpledialog  # noqa
import openpyxl, openpyxl.cell.cell, openpyxl.formula.translate, openpyxl.formula.tokenizer  # noqa
import openpyxl.styles, openpyxl.utils, openpyxl.worksheet.datavalidation, openpyxl.formatting.rule  # noqa
import xlrd  # noqa
try:                                     # 掃描 PDF 辨識用（寰虹這類客戶）；沒裝的話其他客戶照常可以用
    import numpy, cv2, pypdfium2, opencc, onnxruntime, rapidocr_onnxruntime  # noqa
except ImportError:
    pass


def 程式資料夾():
    """exe（或 .py）所在的資料夾；Mac 的 .app 要往外跳到 .app 所在的資料夾"""
    if getattr(sys, "frozen", False):
        路徑 = os.path.dirname(sys.executable)
        if ".app/Contents/MacOS" in 路徑:
            路徑 = 路徑.split(".app/Contents/MacOS")[0]
            路徑 = os.path.dirname(路徑)
        return 路徑
    return os.path.dirname(os.path.abspath(__file__))


def 顯示錯誤(訊息):
    try:
        from tkinter import Tk, messagebox
        r = Tk(); r.withdraw()
        messagebox.showerror("薪資印領清冊", 訊息)
        r.destroy()
    except Exception:
        print(訊息)


def main():
    夾 = 程式資料夾()
    os.environ["PAYROLL_APP_DIR"] = 夾             # 告訴 清冊共用.py：客戶模組、客戶設定 在這個資料夾
    核心 = os.path.join(夾, "清冊共用.py")
    if not os.path.exists(核心):
        顯示錯誤(f"找不到 清冊共用.py\n請把它放在這個資料夾：\n{夾}")
        return
    import importlib.util
    try:
        規格 = importlib.util.spec_from_file_location("清冊共用", 核心)
        模組 = importlib.util.module_from_spec(規格)
        sys.modules["清冊共用"] = 模組               # 客戶模組裡的 import 清冊共用 會拿到這個
        規格.loader.exec_module(模組)
    except SystemExit as e:
        顯示錯誤(str(e.code)); return
    except Exception as e:
        import traceback
        顯示錯誤(f"清冊共用.py 有錯，打不開：\n{type(e).__name__}: {e}\n\n{traceback.format_exc()[-800:]}")
        return
    模組.main()


if __name__ == "__main__":
    main()