# -*- coding: utf-8 -*-
"""
客戶模組：粲紘／連宜
====================
粲紘貿易、連宜兩間公司的薪資放在同一個 Excel，每個月一個 Excel（例如 11501.xls、115.02.xlsx）。
  ・畫面上「每月的薪資檔」一次可以選好幾個檔（只讀選進來的檔，資料夾裡其他檔不管）
  ・有的月份兩間公司疊在同一張工作表（上面粲紘、下面連宜），有的月份分成兩張表
    → 不看工作表名稱，看每一段上面的「公司名稱」那一列來分公司
  ・每一段：公司名稱 → 標題（例如「115年02月份薪資印領清冊」）→ 欄位標題 → 每個人 → 合計
  ・怎麼放進清冊（以清冊 115 年為例）：
      「115年○○月份薪資印領清冊」 → 那個月（一月～十二月）
      「114年12月份薪資印領清冊」（去年 12 月） → 前期薪資
      「114年度年終金…」（去年度的年終） → 上期年終（樣板叫前期年終也可以）
      不休假獎金（特休）→ 免稅，不放進清冊（認法：標題或工作表名稱有下面「免稅詞」裡的字）
      其他獎金（績效、尾牙、長照補助…） → 全部加在「端午獎金」那一欄
  ・薪資總額＝應稅總額、退休自提＝勞退自提金額、扣繳稅額＝代扣稅額；每段都跟「合計」列驗算

兩間公司各一個客戶（各出一份清冊），用同一個模組：薪資檔每一段上面的公司名稱含「客戶名稱」的，才算這個客戶的。
"""
import os
import re
import 清冊共用 as 共

名稱 = "粲紘／連宜"
多個薪資檔 = True                 # 告訴共用核心：薪資檔由這個模組自己選（好幾個檔），不用「薪資檔」那一列
說明 = "每月一個 Excel、兩間公司放在一起；依公司名稱分開讀"
預設客戶 = ["粲紘貿易", "連宜"]
預設設定 = {
    "薪資總額": "應稅總額", "退休自提": "勞退自提金額", "扣繳稅額": "代扣稅額",   # 給畫面、樣板看的（實際讀法在下面）
    "執行業務另開": False,
}
# 欄位標題（去掉空白換行後）含這些字 → 讀這一欄
欄位認法 = {
    "姓名": ["員工姓名", "姓名"],
    "編號": ["編號"],
    "證": ["身分證字號", "身分證"],
    "址": ["戶籍地址", "地址"],
    "總額": ["應稅總額"],
    "自提": ["勞退自提金額", "自提"],
    "稅": ["代扣稅額"],
}
免稅詞 = ["不休假", "特休"]           # 獎金表的標題或工作表名稱有這些字 → 免稅，不放進清冊
月薪標題 = re.compile(r"(\d{2,3})年(\d{1,2})月份?薪資印領清冊")


def 公司(設):                         # 用客戶名稱去找：粲紘貿易 → 「粲紘貿易股份有限公司」那一段
    return 共.清字(設.get("客戶"))


# ---------------------------------------------------------------- 掃描一個資料夾
# 回傳每一段（某檔案、某工作表、某公司、某標題）：
#   {"檔", "表", "公司", "標題", "年", "月", "人們": [{名, 編, 總額, 自提, 稅}], "合計": {總額, 自提, 稅}}
def 掃描檔案(路徑):
    段們 = []
    wb = 共.開活頁簿(路徑)
    for ws in wb.worksheets:
        段 = None
        公司名 = 標題 = ""
        for r in range(1, ws.max_row + 1):
            值們 = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 30) + 1)]
            字們 = [共.清字(v) for v in 值們]
            首 = 字們[0] if 字們 else ""
            if "公司" in 首 and not any(isinstance(v, (int, float)) for v in 值們):   # 公司名稱那一列
                公司名, 標題, 段 = 首, "", None
                continue
            找標題 = next((x for x in 字們 if "印領清冊" in x), None)
            if 找標題 and 段 is None:
                標題 = 找標題
                continue
            if any(x in 欄位認法["姓名"] for x in 字們):                      # 欄位標題那一列
                欄 = {}
                for 鍵, 詞們 in 欄位認法.items():
                    for 詞 in 詞們:                                       # 先找完全一樣，再找「含有」
                        c = next((i for i, x in enumerate(字們) if x == 詞), None)
                        if c is None:
                            c = next((i for i, x in enumerate(字們) if 詞 in x and "補充" not in x), None)
                        if c is not None:
                            欄[鍵] = c
                            break
                m = 月薪標題.search(標題)
                段 = {"檔": os.path.basename(路徑), "表": ws.title.strip(), "公司": 公司名, "標題": 標題,
                     "年": int(m.group(1)) if m else None, "月": int(m.group(2)) if m else None,
                     "欄": 欄, "人們": [], "合計": None, "標題列": r}
                段們.append(段)
                continue
            if 段 is None or 段["合計"] is not None:
                continue
            欄 = 段["欄"]
            取 = lambda 鍵: 共.數字(值們[欄[鍵]]) if 鍵 in 欄 and 欄[鍵] < len(值們) else 0
            if 首 == "合計":
                段["合計"] = {k: 取(k) for k in ("總額", "自提", "稅")}
                continue
            名 = 共.清字(值們[欄["姓名"]]) if "姓名" in 欄 else ""
            if not 共.是人名(名):
                continue
            文字 = lambda 鍵: str(值們[欄[鍵]] or "").strip() if 鍵 in 欄 and 欄[鍵] < len(值們) else ""
            段["人們"].append({"名": 名, "編": 文字("編號"), "證": 文字("證"), "址": 文字("址"),
                             "總額": 取("總額"), "自提": 取("自提"), "稅": 取("稅"), "列": r})
    return 段們


_快取 = {}


def 掃描一個(路徑):                  # 同一個檔沒改過就不重讀（畫面每次更新都會掃）
    鍵 = (路徑, os.path.getmtime(路徑))
    if 鍵 not in _快取:
        _快取[鍵] = 掃描檔案(路徑)
    return _快取[鍵]


def 檔案月份(段們):
    """這個 Excel 是哪一年哪個月：看裡面月薪表的標題（例如 115年01月份薪資印領清冊）"""
    月們 = sorted({(s["年"], s["月"]) for s in 段們 if s["月"]})
    return 月們[0] if len(月們) == 1 else (月們 if 月們 else None)


# 找清冊上的欄：名們 依序找，找到第一個有的；清冊欄是 None（畫面還沒選樣板）就先用第一個名字
def 找欄(名們, 清冊欄):
    if 清冊欄 is None:
        return 名們[0]
    欄 = next((n for n in 名們 if n in 清冊欄), None)
    if not 欄:
        raise SystemExit(f"清冊樣板沒有「{'」或「'.join(名們)}」欄")
    return 欄


# 分類：每一段要放進清冊哪一欄。回傳 (欄位, 類型, 說明)；欄位是 None 的不放
def 分類(段, 年, 檔月, 清冊欄=None):
    if 段["月"]:                                         # 月薪表
        if 段["年"] == 年:
            return 共.中文月[段["月"] - 1], "月薪", f"{段['月']}月月薪"
        if 段["年"] == 年 - 1 and 段["月"] == 12:            # 去年 12 月 → 前期薪資
            欄 = 找欄(["前期薪資"], 清冊欄)
            return 欄, "月薪", 欄
        return None, None, f"是 {段['年']} 年 {段['月']} 月的月薪，不是 {年} 年，也不是去年 12 月，沒有讀"
    m = re.search(r"(\d{2,3})年度?年終", 段["標題"])
    if m:                                               # 年終金
        if int(m.group(1)) != 年 - 1:
            return None, None, f"是 {m.group(1)} 年度的年終，不是 {年 - 1} 年度，沒有讀"
        欄 = 找欄(["上期年終", "前期年終"], 清冊欄)
        return 欄, "獎金", 欄
    if any(詞 in 段["標題"] or 詞 in 段["表"] for 詞 in 免稅詞):   # 不休假獎金（特休）：免稅
        return None, "免稅", "免稅，不放進清冊"
    if isinstance(檔月, tuple) and 檔月[0] != 年:          # 其他獎金：去年的檔裡的不算
        return None, None, f"在 {檔月[0]} 年 {檔月[1]} 月的檔裡，不是 {年} 年的，沒有讀"
    m = re.search(r"(\d{2,3})年\d{1,2}月", 段["標題"])
    if m and int(m.group(1)) != 年:
        return None, None, f"標題是 {m.group(1)} 年的，不是 {年} 年，沒有讀"
    欄 = 找欄(["端午獎金"], 清冊欄)
    return 欄, "獎金", 欄


def 簡稱(段):
    標 = re.sub(r"^\d{2,3}年(\d{1,2}月份|度)?", "", 段["標題"]).replace("印領清冊", "") or 段["表"]
    return 標


# ---------------------------------------------------------------- 讀取（產生清冊時共用核心會叫這個）
def 讀取(設, 年, wb, wbf, 清冊欄):
    表們, 缺表, 警告 = [], [], []
    檔們 = [p for p in 設.get("薪資檔們", []) if p]
    if not 檔們:
        raise SystemExit("請在「檔案與年度」頁的「每月的薪資檔」加入這個客戶每個月的 Excel")
    有我的 = False
    免稅們 = []
    看過 = {}
    for 路徑 in 檔們:
        if not os.path.exists(路徑):
            缺表.append(os.path.basename(路徑))
            continue
        段們 = 掃描一個(路徑)
        檔月 = 檔案月份(段們)
        檔名 = os.path.basename(路徑)
        for s in [s for s in 段們 if 公司(設) in s["公司"]]:
            有我的 = True
            欄位, 類型, 說 = 分類(s, 年, 檔月, 清冊欄)
            名 = f"{檔名}「{s['表']}」{簡稱(s)}"
            if 類型 == "免稅":
                免稅們.append(f"{名}（{len(s['人們'])} 人，{sum(p['總額'] for p in s['人們']):,}）")
                continue
            if 欄位 is None:
                警告.append(f"{名}：{說}")
                continue
            if 類型 == "月薪" and s["月"] and s["年"] == 年:
                if s["月"] in 看過:
                    警告.append(f"{年}年{s['月']}月的月薪出現兩次：{看過[s['月']]}、{名}，兩份都讀進來了，請確認")
                看過[s["月"]] = 名
            缺 = [k for k in ("總額", "自提", "稅") if k not in s["欄"]]
            if 缺:
                警告.append(f"{名}找不到欄位：{'、'.join({'總額': '應稅總額', '自提': '勞退自提金額', '稅': '代扣稅額'}[k] for k in 缺)}")
            來源 = 共.資料來源(f"{年}年{s['月']}月" if s["月"] else 名, 欄位, 類型)
            來源.顯示名稱 = 名 + ("" if s["月"] else f" → {說}")
            來源.總額, 來源.自提, 來源.扣繳 = ["應稅總額"], "勞退自提金額", ["代扣稅額"]
            for p in s["人們"]:
                來源.加(p["名"], 總額=p["總額"], 自提=p["自提"], 稅=p["稅"], 編=p["編"], 證=p["證"], 址=p["址"])
            if s["合計"] is None:
                警告.append(f"{名}：找不到「合計」列，沒有驗算")
            else:
                for k in ("總額", "自提", "稅"):
                    來源.對合計(k, s["合計"][k])
            表們.append(來源)
    if 免稅們:
        警告.append("不休假獎金（特休）免稅，沒有放進清冊：" + "、".join(免稅們))
    if not 有我的:
        raise SystemExit(f"選的薪資檔裡都找不到公司名稱含「{公司(設)}」的段落，請檢查客戶名稱")
    表們.sort(key=lambda t: 共.中文月.index(t.欄位) if t.欄位 in 共.中文月 else -1)
    return 表們, 缺表, 警告


# ---------------------------------------------------------------- 畫面：檔案頁下面多兩張卡片
def 頁面(介, 設):
    tk = 介.tk
    from tkinter import filedialog
    檔們 = list(設.get("薪資檔們", []))

    介.區塊標題(介.檔案頁, "每月的薪資檔",
             右邊=lambda 框: 介.文字按鈕(框, "＋ 加入檔案…", lambda: 加入(), 色=介.PRIMARY_DARK, 移上色=介.TEXT,
                                     bg=介.BG).pack(side="right"))
    卡 = 介.卡片(介.檔案頁)
    清單 = tk.Frame(卡, bg=介.CARD_BG)
    清單.pack(fill="x")
    清單.columnconfigure(0, weight=1)
    提示 = 介.小字(卡, f"一次可以選好幾個檔（按住 Command／Ctrl 多選）；只讀這裡列出的檔，而且只讀公司名稱含「{設.get('客戶')}」的段落。\n"
                    "今年的月薪放那個月、去年 12 月薪資放前期薪資、去年度年終金放上期年終，不休假獎金（特休）免稅不放，其他獎金全部加在端午獎金。", 寬=620)
    提示.pack(fill="x", pady=(6, 0))
    狀 = {"文": "", "色": 介.TEXT_GRAY}

    def 畫():
        for w in 清單.winfo_children():
            w.destroy()
        設2, 年 = 介.收集(), 介.目前年()
        有月, 錯, 要看 = set(), 0, 0
        r = 0

        def 小行(文, 色):
            nonlocal r
            tk.Label(清單, text="　" + 文, font=介.字["小"], fg=色, bg=介.CARD_BG, anchor="w",
                     justify="left").grid(row=r, column=0, columnspan=2, sticky="w")
            r += 1

        for 路徑 in 檔們:
            tk.Label(清單, text=os.path.basename(路徑), font=介.字["標籤"], fg=介.TEXT, bg=介.CARD_BG,
                     anchor="w").grid(row=r, column=0, columnspan=2, sticky="w", pady=(8, 2))
            介.文字按鈕(清單, "✕", lambda p=路徑: 移除(p), 移上色=介.ERR).grid(row=r, column=2, sticky="e", pady=(8, 2))
            r += 1
            if not 年:
                continue
            if not os.path.exists(路徑):
                小行("✗ 找不到這個檔", 介.ERR); 錯 += 1; continue
            try:
                段們 = 掃描一個(路徑)
            except SystemExit as ex:
                小行(f"✗ {ex.code}", 介.ERR); 錯 += 1; continue
            except Exception as ex:
                小行(f"✗ 讀不到：{ex}", 介.ERR); 錯 += 1; continue
            我的 = [s for s in 段們 if 公司(設2) in s["公司"]]
            if not 我的:
                小行(f"✗ 沒有「{公司(設2)}」的資料", 介.ERR); 錯 += 1; continue
            檔月 = 檔案月份(段們)
            for s in 我的:
                try:
                    欄位, 類型, 說 = 分類(s, 年, 檔月, 介.資料["樣板欄"] or None)
                except SystemExit as ex:
                    欄位, 類型, 說 = None, None, ex.code
                金額 = sum(p["總額"] for p in s["人們"])
                if s["月"] and s["年"] == 年:      # 今年的月薪
                    if 欄位:
                        有月.add(s["月"])
                    文 = 說
                else:
                    名 = f"{s['年']}年{s['月']}月薪資" if s["月"] else 簡稱(s)
                    文 = f"{名}（{len(s['人們'])} 人，{金額:,}）→ {說}" if 欄位 else f"{名}：{說}"
                if 類型 == "免稅":
                    小行(f"{簡稱(s)}（{len(s['人們'])} 人，{金額:,}）→ 免稅，不放", 介.TEXT_GRAY)
                    continue
                小行(文 if 欄位 else f"⚠ {文}", 介.OK if 欄位 else 介.WARN)
                if not 欄位:
                    要看 += 1
        if not 檔們:
            tk.Label(清單, text="還沒有加入檔案", font=介.字["小"], fg=介.ERR, bg=介.CARD_BG).grid(row=0, column=0, sticky="w")
            狀.update(文="✗ 還沒加入薪資檔", 色=介.ERR)
        elif 錯:
            狀.update(文=f"✗ {錯} 個檔有問題", 色=介.ERR)
        elif 要看:
            狀.update(文=f"⚠ {要看} 張表要確認", 色=介.WARN)
        else:
            狀.update(文=f"✓ {len(檔們)} 個檔｜月薪 {len(有月)} 個月", 色=介.OK)

    def 加入():
        起點 = os.path.dirname(檔們[-1]) if 檔們 else ""
        選 = filedialog.askopenfilenames(parent=介.root, title="選擇每個月的薪資檔（可以多選）", initialdir=起點,
                                       filetypes=[("Excel 檔案", "*.xlsx *.xlsm *.xls"), ("所有檔案", "*.*")])
        for p in 選 or ():
            if p not in 檔們:
                檔們.append(p)
        排序()
        畫()

    def 移除(p):
        檔們.remove(p)
        畫()

    def 排序():                          # 依檔案裡的月份排（獎金自己一個檔的看標題；都沒有的排最後）
        def 鍵(p):
            try:
                月 = 檔案月份(掃描一個(p))
                if isinstance(月, tuple):
                    return 月
                for 段 in 掃描一個(p):
                    m = re.search(r"(\d{2,3})年(\d{1,2})月", 段["標題"])
                    if m:
                        return (int(m.group(1)), int(m.group(2)))
                return (9999, 99)
            except Exception:
                return (9999, 99)
        檔們.sort(key=鍵)

    def 收集(設):
        設["薪資檔們"] = list(檔們)

    def 狀態():
        return {"檔案": (狀["文"], 狀["色"])} if 狀["文"] else {}
    畫()
    return {"收集": 收集, "載入": 畫, "年改了": 畫, "樣板": 畫, "狀態": 狀態}