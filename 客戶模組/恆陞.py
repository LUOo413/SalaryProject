# -*- coding: utf-8 -*-
"""
客戶模組：恆陞
==============
恆陞國際的薪資檔：一個 Excel 放全年，一個專案一張工作表（公司員工、手工封裝、明潭勤務…），
  另外一張「扣繳總表」是每個專案每個月的合計。
  ・每張專案表已經是清冊的樣子：編號｜身份證號｜所得人姓名｜地址｜明細（薪資總額）｜一月～十二月｜合計
  ・同一個人在好幾個專案都有薪水 → 依姓名加總成一個人
  ・專案的編號每張都從 001 開始，不能用來認人，所以不讀編號（編號用人員名冊的）
  ・沒有扣繳稅額
員工自提退休金：每個月一個勞保局的 txt（沒有標題，用逗號分欄）
  第 3 欄＝年月（11501）、第 5 欄＝姓名、第 7 欄＝身分證號、第 8 欄＝月提繳工資、第 9 欄＝自提率、第 10 欄＝自提金額
  ・txt 只用來知道「誰」自提多少：身分證號不填進清冊（清冊的身分證、地址用人員名冊）
  ・找人：姓名一樣 → 人員名冊上同一個身分證號的人 → 姓名裡有 ＊（罕用字）的，其他字一樣而且只有一個人
驗算：每個月 ＝ 每張專案表的合計列加起來，也跟「扣繳總表」的合計比；自提跟每個 txt 的總和比。
"""
import os
import re
import 清冊共用 as 共

名稱 = "恆陞"
說明 = "一個 Excel 一個專案一張表（全年），自提是每月勞保局 txt"
預設客戶 = ["恆陞國際"]
預設設定 = {
    "薪資總額": "薪資總額", "退休自提": "勞退 txt 第 10 欄", "扣繳稅額": "",
    "自提檔們": [],
    "執行業務另開": False,
}


# ---------------------------------------------------------------- 讀薪資 Excel
# 回傳 (專案們, 總表)：
#   專案們：[{"表", "人們": [(姓名, {月: 金額})], "合計": {月: 金額}}]
#   總表：{月: 金額}（扣繳總表最下面的合計列），沒有就是 None
def 讀薪資(wb):
    專案們, 總表 = [], None
    for ws in wb.worksheets:
        標頭 = None
        for r in range(1, min(ws.max_row, 15) + 1):
            字們 = [共.清字(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
            if "明細" in 字們 and any("姓名" in x for x in 字們) and "一月" in 字們:
                標頭 = (r, 字們)
                break
        if 標頭 is None:
            if "總表" in ws.title:
                總表 = 讀總表(ws)
            continue
        r0, 字們 = 標頭
        姓名欄 = next(i for i, x in enumerate(字們) if "姓名" in x) + 1
        明細欄 = 字們.index("明細") + 1
        月欄 = {m: 字們.index(共.中文月[m - 1]) + 1 for m in range(1, 13) if 共.中文月[m - 1] in 字們}
        專案 = {"表": ws.title.strip(), "人們": [], "合計": None}
        for r in range(r0 + 1, ws.max_row + 1):
            名 = 共.清字(ws.cell(r, 姓名欄).value)
            明細 = 共.清字(ws.cell(r, 明細欄).value)
            值們 = {m: ws.cell(r, c).value for m, c in 月欄.items()}
            if 名:
                if 共.是人名(名) and (not 明細 or "薪資" in 明細):
                    專案["人們"].append((名, {m: 共.數字(v) for m, v in 值們.items()}))
                continue
            if not 明細 and any(isinstance(v, (int, float)) for v in 值們.values()):   # 沒有姓名、有數字 → 合計列
                專案["合計"] = {m: 共.數字(v) for m, v in 值們.items()}
                break
        專案們.append(專案)
    return 專案們, 總表


def 讀總表(ws):
    for r in range(1, ws.max_row + 1):
        if 共.清字(ws.cell(r, 1).value) == "合計":
            for rr in range(1, r):                        # 找「1月份、2月份…」那一列
                字們 = [共.清字(ws.cell(rr, c).value) for c in range(1, ws.max_column + 1)]
                月欄 = {m: 字們.index(f"{m}月份") + 1 for m in range(1, 13) if f"{m}月份" in 字們}
                if 月欄:
                    return {m: 共.數字(ws.cell(r, c).value) for m, c in 月欄.items()}
    return None


# ---------------------------------------------------------------- 讀自提 txt
# 讀自提檔：兩種格式都可以
#   A. 逗號分欄（勞保局下載的）：第 3 欄＝年月、第 5 欄＝姓名、第 7 欄＝身分證、第 10 欄＝自提金額
#   B. Tab 分欄（只有三欄）：姓名、身分證、自提金額（例如 " 1,908 "）；年月從檔名找，例如 …Y11411… → 114 年 11 月
def 讀自提檔(路徑):
    """回傳 [(年, 月, 姓名, 身分證, 自提金額)]"""
    原 = open(路徑, "rb").read()
    for 編碼 in ("utf-8-sig", "cp950", "big5hkscs"):
        try:
            文字 = 原.decode(編碼)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SystemExit(f"{os.path.basename(路徑)} 讀不到（不是 UTF-8 也不是 Big5）")
    結果 = []
    檔名年月 = None
    for 行 in 文字.splitlines():
        if not 行.strip():
            continue
        if "\t" in 行:                                           # B. Tab 分欄
            欄 = [x.strip().strip('"').strip() for x in 行.split("\t")]
            證們 = [x for x in 欄 if re.fullmatch(r"[A-Za-z][0-9A-Da-d]\d{8}", 共.清字(x))]
            金們 = [x for x in 欄 if re.fullmatch(r"-?[\d,]+(\.\d+)?", x.replace(" ", "")) and x not in 證們]
            名們 = [共.清字(x) for x in 欄 if 共.清字(x) and x not in 證們 and x not in 金們]
            if not 金們 or not 名們:
                continue
            if 檔名年月 is None:
                m = re.search(r"Y(\d{3})(\d{2})", os.path.basename(路徑)) or \
                    re.search(r"(?<!\d)(1\d{2})(0[1-9]|1[0-2])(?!\d)", os.path.basename(路徑))
                if not m:
                    raise SystemExit(f"{os.path.basename(路徑)}：檔案裡沒有年月，檔名也找不到（要有像 11501 這樣的年月）")
                檔名年月 = (int(m.group(1)), int(m.group(2)))
            結果.append((檔名年月[0], 檔名年月[1], 名們[0], 共.清字(證們[0]).upper() if 證們 else "",
                       共.數字(金們[-1].replace(" ", ""))))
            continue
        欄 = [x.strip() for x in 行.split(",")]                   # A. 逗號分欄
        if len(欄) < 10 or not 欄[2].isdigit():
            continue
        年月 = 欄[2]
        結果.append((int(年月[:-2]), int(年月[-2:]), 共.清字(欄[4]), 共.清字(欄[6]).upper(), 共.數字(欄[9])))
    return 結果


# 自提的人 → 薪資表上的哪個人。回傳 (薪資表上的姓名, 怎麼對到的)；對不到回傳 (None, 原因)
def 對人(名, 證, 薪資名們, 名冊依證):
    if 名 in 薪資名們:
        return 名, ""
    if 證 and 證 in 名冊依證 and 名冊依證[證] in 薪資名們:
        return 名冊依證[證], f"txt 姓名「{名}」，用人員名冊的身分證對到"
    if "＊" in 名 or "*" in 名:
        樣 = re.compile("^" + "".join("." if ch in "＊*" else re.escape(ch) for ch in 名) + "$")
        像 = [n for n in 薪資名們 if 樣.match(n)]
        if len(像) == 1:
            return 像[0], f"txt 姓名「{名}」有 ＊，其他字一樣的只有這個人"
        if len(像) > 1:
            return None, f"姓名有 ＊，像的人不只一個：{'、'.join(像)}"
    return None, "薪資表沒有這個人"


# ---------------------------------------------------------------- 讀取（產生清冊時共用核心會叫這個）
def 讀取(設, 年, wb, wbf, 清冊欄):
    表們, 缺表, 警告 = [], [], []
    專案們, 總表 = 讀薪資(wb)
    if not 專案們:
        raise SystemExit("薪資檔裡找不到專案表（要有「所得人姓名」「明細」「一月」這些標題）")
    # 薪資：一個月一個資料來源（13 個專案加在一起），合計＝每張專案表合計列的加總
    月們 = sorted({m for p in 專案們 for _, 金 in p["人們"] for m, v in 金.items() if v})
    for m in 月們:
        來源 = 共.資料來源(f"{年}年{m}月", 共.中文月[m - 1], "月薪")
        來源.顯示名稱 = f"{m}月薪資（{len(專案們)} 個專案）"
        來源.總額, 來源.自提, 來源.扣繳 = ["薪資總額"], None, []
        表上 = 0
        for p in 專案們:
            個人和 = 0
            for 名, 金 in p["人們"]:
                if 金.get(m):
                    來源.加(名, 總額=金[m])
                    個人和 += 金[m]
            if p["合計"] is not None:
                表上 += p["合計"].get(m, 0)
                if abs(個人和 - p["合計"].get(m, 0)) > 1:
                    警告.append(f"「{p['表']}」{m}月：個人加總 {個人和:,} ≠ 表上合計 {p['合計'].get(m, 0):,}")
            else:
                警告.append(f"「{p['表']}」找不到合計列，沒有驗算")
        來源.對合計("總額", 表上)
        if 總表 and m in 總表 and abs(總表[m] - 表上) > 1:
            警告.append(f"{m}月：各專案合計 {表上:,} ≠ 扣繳總表 {總表[m]:,}")
        表們.append(來源)
    if 總表 is None:
        警告.append("找不到「扣繳總表」，沒有跟總表核對")

    # 自提：每個 txt 一個資料來源
    薪資名們 = {名 for p in 專案們 for 名, _ in p["人們"]}
    名冊依證 = {}
    if 設.get("人員名冊") and os.path.exists(設["人員名冊"]):
        名冊依證 = {v["身分證號"]: k for k, v in 共.讀前年度(設["人員名冊"]).items() if v.get("身分證號")}
    for 路徑 in 設.get("自提檔們", []):
        if not os.path.exists(路徑):
            缺表.append(os.path.basename(路徑))
            continue
        列們 = 讀自提檔(路徑)
        if not 列們:
            警告.append(f"{os.path.basename(路徑)}：讀不到任何一筆自提")
            continue
        for 月份 in sorted({(y, m) for y, m, *_ in 列們}):
            y, m = 月份
            if y != 年:
                警告.append(f"{os.path.basename(路徑)}：是 {y} 年 {m} 月的自提，不是 {年} 年，沒有讀")
                continue
            這月 = [x for x in 列們 if (x[0], x[1]) == 月份]
            來源 = 共.資料來源(f"{y}年{m}月自提", 共.中文月[m - 1], "月薪")
            來源.顯示名稱 = f"{os.path.basename(路徑)}（{m}月自提）"
            來源.總額, 來源.自提, 來源.扣繳 = [], "第 10 欄", []
            for _, _, 名, 證, 金 in 這月:
                誰, 說 = 對人(名, 證, 薪資名們, 名冊依證)
                if 誰 is None:
                    警告.append(f"{m}月自提對不到人：{名} 自提 {金:,}（{說}），沒有放進清冊")
                    continue
                if 說:
                    警告.append(f"{m}月自提：{誰} ← {說}")
                if m not in 月們 or not any(金2.get(m) for p in 專案們 for n2, 金2 in p["人們"] if n2 == 誰):
                    警告.append(f"{m}月：{誰} 有自提 {金:,} 但這個月沒有薪資，請確認")
                來源.加(誰, 自提=金)
            來源.併入 = True                              # 摘要的「月薪表幾張」不算自提
            來源.對合計("自提", sum(x[4] for x in 這月))      # 對不到的人沒放進來，加總會比 txt 少，驗算會標出來
            表們.append(來源)
    if not 設.get("自提檔們"):
        警告.append("沒有選員工自提的 txt，退休金自提都是空白")
    return 表們, 缺表, 警告


# ---------------------------------------------------------------- 畫面：檔案頁下面多一張「員工自提」卡片
def 頁面(介, 設):
    tk = 介.tk
    from tkinter import filedialog
    檔們 = list(設.get("自提檔們", []))
    介.區塊標題(介.檔案頁, "每月的員工自提（勞保局 txt）",
             右邊=lambda 框: 介.文字按鈕(框, "＋ 加入檔案…", lambda: 加入(), 色=介.PRIMARY_DARK, 移上色=介.TEXT,
                                     bg=介.BG).pack(side="right"))
    卡 = 介.卡片(介.檔案頁)
    清單 = tk.Frame(卡, bg=介.CARD_BG)
    清單.pack(fill="x")
    清單.columnconfigure(0, weight=1)
    介.小字(卡, "一次可以選好幾個 txt（按住 Command／Ctrl 多選）。逗號分欄的自提金額是第 10 欄；只有三欄（Tab 分開）的年月看檔名；"
              "txt 的身分證只用來認人，清冊的身分證、地址用人員名冊。", 寬=620).pack(fill="x", pady=(6, 0))
    狀 = {"文": "", "色": 介.TEXT_GRAY}

    def 畫():
        for w in 清單.winfo_children():
            w.destroy()
        年 = 介.目前年()
        薪資名們 = set()
        if 介.資料["wb"] is not None:
            try:
                薪資名們 = {名 for p in 讀薪資(介.資料["wb"])[0] for 名, _ in p["人們"]}
            except Exception:
                pass
        錯 = 0
        for i, 路徑 in enumerate(檔們):
            tk.Label(清單, text=os.path.basename(路徑), font=介.字["內文"], fg=介.TEXT, bg=介.CARD_BG,
                     anchor="w").grid(row=i * 2, column=0, sticky="w", pady=(4, 0))
            介.文字按鈕(清單, "✕", lambda p=路徑: 移除(p), 移上色=介.ERR).grid(row=i * 2, column=1, sticky="e")
            try:
                列們 = 讀自提檔(路徑)
                月們 = sorted({(y, m) for y, m, *_ in 列們})
                文 = "、".join(f"{y}年{m}月" for y, m in 月們) + f"｜{len(列們)} 人｜自提合計 {sum(x[4] for x in 列們):,}"
                色 = 介.OK
                if 薪資名們:
                    對不到 = [x[2] for x in 列們 if 對人(x[2], x[3], 薪資名們, {})[0] is None]
                    if 對不到:
                        文 += f"｜⚠ {len(對不到)} 人薪資表找不到：{'、'.join(對不到)}"
                        色 = 介.WARN
                if 年 and any(y != 年 for y, _ in 月們):
                    文 += f"｜⚠ 不是 {年} 年"
                    色 = 介.WARN
            except SystemExit as ex:
                文, 色 = f"✗ {ex.code}", 介.ERR
            except Exception as ex:
                文, 色 = f"✗ 讀不到：{ex}", 介.ERR
            if 色 == 介.ERR:
                錯 += 1
            tk.Label(清單, text=文, font=介.字["小"], fg=色, bg=介.CARD_BG, anchor="w", justify="left",
                     wraplength=600).grid(row=i * 2 + 1, column=0, columnspan=2, sticky="w")
        if not 檔們:
            tk.Label(清單, text="還沒有加入自提的 txt", font=介.字["小"], fg=介.WARN, bg=介.CARD_BG).grid(row=0, column=0, sticky="w")
        狀.update(文=("✗ 自提檔有問題" if 錯 else f"自提 {len(檔們)} 個月") if 檔們 else "", 色=介.ERR if 錯 else 介.TEXT_GRAY)

    def 加入():
        起點 = os.path.dirname(檔們[-1]) if 檔們 else os.path.dirname(介.收集().get("薪資檔") or "")
        選 = filedialog.askopenfilenames(parent=介.root, title="選擇每個月的員工自提 txt（可以多選）", initialdir=起點,
                                       filetypes=[("文字檔", "*.txt *.csv"), ("所有檔案", "*.*")])
        for p in 選 or ():
            if p not in 檔們:
                檔們.append(p)
        檔們.sort(key=lambda p: os.path.basename(p))
        畫()

    def 移除(p):
        檔們.remove(p)
        畫()

    def 收集(設):
        設["自提檔們"] = list(檔們)

    def 狀態():
        return {}
    畫()
    return {"收集": 收集, "載入": 畫, "狀態": 狀態}