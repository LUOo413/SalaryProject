# -*- coding: utf-8 -*-
"""
客戶模組：通用
==============
格式比較單純的客戶用這個：一個 Excel、每個月一張工作表（例如 115年3月、11403），
獎金一種一張工作表。在畫面上填「工作表名稱、薪資總額是哪幾欄」就好，不用寫程式。

這個檔案裡的「卡片」（月薪卡、其他表卡、進階卡、職稱卡）其他客戶模組也可以拿去用，
例如七賢、澤利的模組就是用這些卡片組出自己的頁面。
"""
import 清冊共用 as 共
from collections import OrderedDict

名稱 = "通用"
說明 = "一個 Excel，每月一張工作表；在畫面上設定欄位就好"
預設設定 = {}          # 新增客戶時的預設值（沒寫的用共用核心的空白設定）
預設客戶 = []          # 第一次執行時要自動建好的客戶


# ================================================================ 月薪卡
# 月薪表：工作表名稱（{年度}{月}）、薪資總額＝哪幾欄相加、退休自提、扣繳稅額
def 月薪卡(介, parent, 設):
    tk = 介.tk
    卡 = 介.卡片(parent)
    卡.columnconfigure(1, weight=1)
    變 = {k: tk.StringVar(value=設.get(k, "")) for k in ("月薪表", "薪資總額", "退休自提", "扣繳稅額")}
    介.標籤(卡, "工作表名稱").grid(row=0, column=0, sticky="w", padx=(0, 14), pady=4)
    框, 月名框 = 介.輸入框(卡, 變["月薪表"])
    框.grid(row=0, column=1, sticky="ew", pady=4)
    介.小字(卡, "{年度} 換成清冊年度、{月} 換成 1～12、{月2} 換成 01～12。例：{年度}年{月}月 → 115年3月；{年度}{月2} → 11403"
          ).grid(row=1, column=1, sticky="w")
    提示 = 介.小字(卡)
    提示.grid(row=2, column=1, sticky="w", pady=(2, 6))
    下拉們 = {}
    for i, (名, 說) in enumerate((("薪資總額", "哪幾欄相加，用＋連接，例如 應付薪資＋員工酬勞"),
                                ("退休自提", "員工自提勞退那一欄，沒有就空白"),
                                ("扣繳稅額", "代扣所得稅那一欄（可以＋好幾欄），沒有就空白"))):
        介.標籤(卡, 名).grid(row=3 + i * 2, column=0, sticky="w", padx=(0, 14), pady=(4, 0))
        cb = 介.下拉(卡)
        cb.configure(textvariable=變[名])
        cb.grid(row=3 + i * 2, column=1, sticky="ew", pady=(4, 0))
        介.小字(卡, 說).grid(row=4 + i * 2, column=1, sticky="w")
        下拉們[名] = cb
    舊總 = {"v": 變["薪資總額"].get()}

    def 選了(名, cb):                    # 薪資總額選一欄是「加進去」，其他是「換掉」
        def f(e=None):
            v = cb.get()
            if 名 == "薪資總額" and 舊總["v"]:
                舊 = 共.拆詞(舊總["v"])
                if v and 共.清字(v) not in 舊:
                    cb.set("＋".join(舊 + [v]))
            舊總["v"] = 變["薪資總額"].get()
            更新()
        return f
    for 名, cb in 下拉們.items():
        cb.bind("<<ComboboxSelected>>", 選了(名, cb))
        cb.bind("<FocusOut>", lambda e: (舊總.update(v=變["薪資總額"].get()), 更新()))
    月名框.bind("<FocusOut>", lambda e: 介.重讀())
    月名框.bind("<Return>", lambda e: 介.重讀())

    def 收集(設):
        for k, v in 變.items():
            設[k] = v.get().strip()

    def 更新():
        年, wb, 設 = 介.目前年(), 介.資料["wb"], 介.收集()
        模 = 變["月薪表"].get()
        if 年 is None:
            提示.configure(text="年度請填數字", fg=介.ERR); return
        if wb is None:
            提示.configure(text=f"會讀取 {共.表名(模, 年, 1)}～{共.表名(模, 年, 12)}", fg=介.TEXT_GRAY); return
        有 = [m for m in range(1, 13) if 共.找工作表(wb, 共.表名(模, 年, m))]
        if not 有:
            提示.configure(text=f"✗ 找不到 {共.表名(模, 年, 1)} 這樣的工作表，請檢查名稱", fg=介.ERR); return
        文 = f"找到 {共.表名(模, 年, 有[0])}～{共.表名(模, 年, 有[-1])} 共 {len(有)} 張"
        總 = 共.拆詞(變["薪資總額"].get())
        if not 總:
            提示.configure(text=文 + "；⚠ 請在下面選「薪資總額」是哪幾欄", fg=介.WARN); return
        t = 共.月薪資料表(共.找工作表(wb, 共.表名(模, 年, 有[0])), 設)
        if t is None:
            提示.configure(text=文 + f"；✗ 表裡找不到「{總[0]}」", fg=介.ERR); return
        有過 = set()                      # 有的欄位不是每個月都有（例如員工酬勞），整年都沒有才提醒
        for m in 有:
            tt = 共.月薪資料表(共.找工作表(wb, 共.表名(模, 年, m)), 設)
            if tt:
                有過 |= {k for k in 總 + 共.拆詞(變["退休自提"].get()) if 共.找標題(tt, k)}
        缺 = [k for k in 總 + 共.拆詞(變["退休自提"].get()) if k not in 有過]
        列 = 共.讀表資料(t, 介.資料["wbf"], [])
        說 = (f"；{t.工作表.strip()}：{len(列)} 人，薪資總額＝{'＋'.join(t.總額)}"
             + (f"，自提＝{t.自提}" if t.自提 else "") + (f"，扣繳＝{'＋'.join(t.扣繳)}" if t.扣繳 else ""))
        if 缺:
            說 += f"（⚠ 每個月都找不到：{'、'.join(缺)}）"
        提示.configure(text="✓ " + 文 + 說, fg=介.WARN if 缺 else 介.OK)

    def 載入():
        標題 = []
        年, wb = 介.目前年(), 介.資料["wb"]
        if wb is not None and 年:
            for m in range(1, 13):
                ws = 共.找工作表(wb, 共.表名(變["月薪表"].get(), 年, m))
                if ws is None:
                    continue
                t = (共.月薪資料表(ws, 介.收集()) if 共.拆詞(變["薪資總額"].get()) else None) or next(iter(共.找資料表(ws)), None)
                if t is not None:
                    標題 = [k for k in t.標題 if k != t.姓名標題]
                    break
        for cb in 下拉們.values():
            cb.configure(values=標題)
        更新()

    def 狀態():
        文 = 提示.cget("text")
        if 文.startswith("✓"):
            import re
            m = re.search(r"共 \d+ 張", 文)
            return "✓ " + (m.group(0) if m else ""), 介.OK
        if 文.startswith("✗"):
            return "✗ 找不到月薪表", 介.ERR
        return ("⚠ 請確認欄位", 介.WARN) if 文 else ("", 介.TEXT_GRAY)
    return {"收集": 收集, "載入": 載入, "年改了": 更新, "狀態": 狀態}


# ================================================================ 其他表卡
# 其他表：一列一張工作表（前期薪資、年終、各種獎金），選要填進清冊哪一欄
#   有格式＝可以選「獎金／月薪」；有月份＝可以選「另放月份」（每人多一列，金額放在那個月）
#   固定欄＝不能選月份的欄位（例如 上期年終、端午獎金、中秋獎金、本期年終，它們固定在上面欄位）
def 其他表卡(介, parent, 設, 鍵="其他表", 有格式=True, 有月份=True, 固定欄=(), 預設欄="", 說明文字=None):
    tk = 介.tk
    卡 = 介.卡片(parent)
    表格 = tk.Frame(卡, bg=介.CARD_BG)
    表格.pack(fill="x")
    表格.columnconfigure(0, weight=3)
    表格.columnconfigure(1, weight=2)
    頭 = ["工作表", "填入清冊哪一欄"] + (["格式"] if 有格式 else []) + (["另放月份"] if 有月份 else []) + ["讀到的內容"]
    for c, t in enumerate(頭):
        tk.Label(表格, text=t, font=介.字["小"], fg=介.TEXT_GRAY, bg=介.CARD_BG).grid(row=0, column=c, sticky="w", padx=(0, 8))
    介.小字(卡, 說明文字 or ("格式：「獎金」會自動找金額欄和扣繳欄；「月薪」跟月薪表一樣讀（例如前期薪資讀去年 12 月）。"
                         "表名開頭的年份會跟著清冊年度換。樣板沒有的欄位會自動插在「本期年終」前面。"
                         + ("另放月份：每人在薪資總額下面多一列，金額放在選的月份，扣繳也加到那個月。" if 有月份 else ""))
          ).pack(fill="x", pady=(8, 0))
    列們, 號 = [], {"n": 1}
    月選項 = [""] + 共.中文月

    def 檢查(r):
        名 = r["表"].get().strip()
        wb = 介.資料["wb"]
        if not 名 or wb is None:
            r["狀"].configure(text="", fg=介.TEXT_GRAY); return
        ws = 共.找工作表(wb, 名)
        if ws is None:
            r["狀"].configure(text="✗ 薪資檔找不到這張表", fg=介.ERR); return
        設 = 介.收集()
        try:
            if r["型"].get() == "月薪":
                t = 共.月薪資料表(ws, 設)
                if not t:
                    r["狀"].configure(text="✗ 找不到薪資總額的欄位", fg=介.ERR); return
                列 = 共.讀表資料(t, 介.資料["wbf"], [])
                r["狀"].configure(text=f"✓ 月薪格式｜{len(列)} 人｜總額 {sum(x['金']['總額'] for x in 列):,}", fg=介.OK)
                return
            文 = []
            for t in 共.找資料表(ws):
                if not 共.設定欄位(t, "獎金", 設):
                    continue
                列 = 共.讀表資料(t, 介.資料["wbf"], [])
                文.append(f"金額：{'＋'.join(t.總額)}" + (f"｜扣繳：{'＋'.join(t.扣繳)}" if t.扣繳 else "")
                         + f"｜{len(列)} 人｜合計 {sum(x['金']['總額'] for x in 列):,}")
            r["狀"].configure(text=("✓ " + "；".join(文)) if 文 else "✗ 找不到姓名欄或金額欄", fg=介.OK if 文 else 介.ERR)
        except Exception as e:
            r["狀"].configure(text=f"✗ {e}", fg=介.ERR)

    def 月可選(r):                      # 固定欄位（上期年終…）不能另放月份
        if not 有月份:
            return
        固定 = 共.清字(r["欄"].get()) in 固定欄 or r["型"].get() == "月薪"
        if 固定:
            r["月"].set("")
        r["月"].configure(state="disabled" if 固定 else "readonly")

    def 加(模式="", 欄位="", 型="獎金", 月=""):
        n = 號["n"]; 號["n"] += 1
        年 = 介.目前年()
        r = {"模式": 模式, "型": tk.StringVar(value=型 or "獎金")}
        c = 0
        r["表"] = 介.下拉(表格, 介.資料["工作表"])
        r["表"].grid(row=n, column=c, sticky="ew", padx=(0, 8), pady=3); c += 1
        r["欄"] = 介.下拉(表格, 介.資料["樣板欄"], 寬=12)
        r["欄"].grid(row=n, column=c, sticky="ew", padx=(0, 8), pady=3); c += 1
        r["元件"] = [r["表"], r["欄"]]
        if 有格式:
            cb = 介.下拉(表格, ["獎金", "月薪"], 唯讀=True, 寬=5)
            cb.configure(textvariable=r["型"])
            cb.grid(row=n, column=c, sticky="w", padx=(0, 8), pady=3); c += 1
            cb.bind("<<ComboboxSelected>>", lambda e: (月可選(r), 檢查(r)))
            r["元件"].append(cb)
        if 有月份:
            r["月"] = 介.下拉(表格, 月選項, 唯讀=True, 寬=7)
            r["月"].grid(row=n, column=c, sticky="w", padx=(0, 8), pady=3); c += 1
            r["月"].set(共.中文月[int(月) - 1] if str(月).isdigit() and 1 <= int(月) <= 12 else "")
            r["元件"].append(r["月"])
        r["狀"] = tk.Label(表格, font=介.字["小"], bg=介.CARD_BG, anchor="w", justify="left", wraplength=240)
        r["狀"].grid(row=n, column=c, sticky="w", pady=3); c += 1
        刪 = 介.文字按鈕(表格, "✕", lambda: 刪除(r), 移上色=介.ERR)
        刪.grid(row=n, column=c, padx=(4, 0))
        r["元件"] += [r["狀"], 刪]
        r["表"].set(共.表名(模式, 年) if (年 and 模式) else 模式)
        r["欄"].set(欄位 or 預設欄)

        def 名改了(e=None):
            if 介.目前年():
                r["模式"] = 共.轉模式(r["表"].get(), 介.目前年())
            檢查(r)
        for 事件 in ("<<ComboboxSelected>>", "<FocusOut>"):
            r["表"].bind(事件, 名改了)
            r["欄"].bind(事件, lambda e: 月可選(r))
        列們.append(r)
        月可選(r)
        檢查(r)
        return r

    def 刪除(r):
        for w in r["元件"]:
            w.destroy()
        列們.remove(r)

    for x in 設.get(鍵, []):
        加(x[0], x[1] if len(x) > 1 else "", x[2] if len(x) > 2 else "獎金", x[3] if len(x) > 3 else "")

    def 收集(設):
        年 = 介.目前年() or 設.get("年度", 115)
        設[鍵] = []
        for r in 列們:
            if not r["表"].get().strip():
                continue
            x = [共.轉模式(r["表"].get(), 年), r["欄"].get().strip(), r["型"].get() or "獎金"]
            月 = r["月"].get() if 有月份 else ""
            if 月:
                x.append(共.中文月.index(月) + 1)
            設[鍵].append(x)

    def 樣板():
        for r in 列們:
            r["欄"].configure(values=介.資料["樣板欄"])

    def 載入():
        for r in 列們:
            r["表"].configure(values=介.資料["工作表"])
            檢查(r)

    def 年改了():
        年 = 介.目前年()
        if 年:
            for r in 列們:
                if r["模式"]:
                    r["表"].set(共.表名(r["模式"], 年))

    def 狀態():
        有 = [r for r in 列們 if r["表"].get().strip()]
        錯 = sum(1 for r in 有 if str(r["狀"].cget("fg")).upper() == 介.ERR.upper())
        return f"{len(有)} 張表" + (f"｜✗ {錯} 張有問題" if 錯 else ""), (介.ERR if 錯 else 介.TEXT_GRAY)
    return {"收集": 收集, "樣板": 樣板, "載入": 載入, "年改了": 年改了, "狀態": 狀態, "加": 加, "列們": 列們}


# ================================================================ 職稱卡
# 科目分類用的兩張卡片：
#   ① 會計科目：一列一個科目 ─ 計算表上顯示的名稱、清冊編號的開頭；還有「科目寫在清冊哪一欄」
#   ② 職稱 → 會計科目：一列一個職稱（下拉會列出薪資檔裡的職稱），科目從 ① 選
#   有計算表=False：① 不顯示（科目、編號開頭寫在客戶模組的預設設定裡）
#   存檔時 ① 會轉回原本的寫法：計算表分組「6110/6127=6110、5300=5300」、編號開頭「5300=00、6110=02」
def 職稱卡(介, parent, 設, 有計算表=True):
    import re
    tk = 介.tk
    科列們, 開頭 = [], dict(共.拆對照(設.get("編號開頭", "")))
    列們, 號, 人數 = [], {"n": 0}, OrderedDict()           # ② 職稱表的每一列
    寫入欄 = tk.StringVar(value=設.get("科目寫入欄", ""))

    def 目前科目():                       # 現在有哪些會計科目（給職稱的下拉選）
        if 有計算表:
            return [r["科"].get().strip() for r in 科列們 if r["科"].get().strip()]
        return [v for _, v in 共.拆對照(設.get("計算表分組", ""))]

    if 有計算表:
        科卡 = 介.卡片(parent)
        頭 = tk.Frame(科卡, bg=介.CARD_BG)
        頭.pack(fill="x", pady=(0, 4))
        tk.Label(頭, text="會計科目", font=介.字["標籤"], fg=介.TEXT, bg=介.CARD_BG).pack(side="left")
        介.文字按鈕(頭, "＋ 新增科目", lambda: 加科(), 色=介.PRIMARY_DARK).pack(side="right")
        科表 = tk.Frame(科卡, bg=介.CARD_BG)
        科表.pack(fill="x")
        for c, w in ((0, 2), (1, 3), (2, 1)):
            科表.columnconfigure(c, weight=w, uniform="科")
        for c, t in enumerate(("會計科目", "計算表上的名稱", "編號開頭")):
            tk.Label(科表, text=t, font=介.字["小"], fg=介.TEXT_GRAY, bg=介.CARD_BG, anchor="w").grid(
                row=0, column=c, sticky="w", padx=(0, 8))
        介.小字(科卡, "計算表上的名稱：薪資檔每張表的合計下面，這個科目的加總那一列叫什麼；"
                    "好幾個名稱用 / 隔開（例如 6110/6127），空白＝跟科目一樣。\n"
                    "編號開頭：清冊編號的前兩碼，例如 02 → 0201、0202…；空白＝編號不填。", 寬=620).pack(fill="x", pady=(6, 0))
        tk.Frame(科卡, bg=介.BORDER, height=1).pack(fill="x", pady=10)
        欄列 = tk.Frame(科卡, bg=介.CARD_BG)
        欄列.pack(fill="x")
        介.標籤(欄列, "清冊上放科目的欄").pack(side="left", padx=(0, 10))
        科欄 = 介.下拉(欄列, 寬=14)
        科欄.configure(textvariable=寫入欄)
        科欄.pack(side="left")
        介.小字(科卡, "每個人的會計科目會填進清冊的這一欄，方便看出他屬於哪個科目"
                    "（清冊沒有專門的科目欄，所以借用一欄，例如澤利借「扶養人數」）；空白＝不填。", 寬=620).pack(fill="x", pady=(4, 0))
        科號 = {"n": 1}

        def 加科(科="", 名="", 頭碼=""):
            n = 科號["n"]; 科號["n"] += 1
            r = {"科": tk.StringVar(value=科), "名": tk.StringVar(value=名 if 名 != 科 else ""), "頭": tk.StringVar(value=頭碼)}
            元件 = []
            for c, k, 寬 in ((0, "科", 10), (1, "名", 16), (2, "頭", 5)):
                框, e = 介.輸入框(科表, r[k], 寬=寬)
                框.grid(row=n, column=c, sticky="ew", padx=(0, 8), pady=2)
                元件.append(框)
            刪 = 介.文字按鈕(科表, "✕", lambda: 刪科(r), 移上色=介.ERR)
            刪.grid(row=n, column=3, padx=(4, 0))
            r["元件"] = 元件 + [刪]
            r["科"].trace_add("write", lambda *_: 更新全部())
            科列們.append(r)
            更新全部()

        def 刪科(r):
            for w in r["元件"]:
                w.destroy()
            科列們.remove(r)
            更新全部()


    # ② 職稱 → 會計科目
    卡 = 介.卡片(parent)
    if 有計算表:
        卡.master.pack_configure(pady=(10, 0))
    頭2 = tk.Frame(卡, bg=介.CARD_BG)
    頭2.pack(fill="x", pady=(0, 4))
    tk.Label(頭2, text="職稱 → 會計科目", font=介.字["標籤"], fg=介.TEXT, bg=介.CARD_BG).pack(side="left")
    介.文字按鈕(頭2, "＋ 新增職稱", lambda: 加(), 色=介.PRIMARY_DARK).pack(side="right")
    介.小字(卡, "依這裡的對照，在每月合計下面的計算表用職稱加總；清冊也依這裡分成各科目一份", 寬=620).pack(fill="x")
    表 = tk.Frame(卡, bg=介.CARD_BG)
    表.pack(fill="x", pady=(4, 0))
    表.columnconfigure(0, weight=1)
    表.columnconfigure(1, weight=1)

    def 更新(r):
        職, 科 = r["職"].get().strip(), r["科v"].get().strip()
        n = 人數.get(職)
        科們 = 目前科目()
        if not 職:
            r["狀"].configure(text="", fg=介.TEXT_GRAY)
        elif not 科:
            r["狀"].configure(text="⚠ 請選會計科目" + (f"（{n} 人）" if n else ""), fg=介.WARN)
        elif 科們 and 科 not in 科們:
            r["狀"].configure(text="⚠ 上面沒有這個科目", fg=介.WARN)
        else:
            r["狀"].configure(text=f"✓ {n} 人" if n else "薪資檔沒有這個職稱", fg=介.OK if n else 介.TEXT_GRAY)

    def 更新全部():
        for r in 列們:
            r["科"].configure(values=目前科目())
            更新(r)

    def 加(職="", 科=""):
        n = 號["n"]; 號["n"] += 1
        r = {"職": 介.下拉(表, list(人數)), "科v": tk.StringVar(value=科)}
        r["職"].grid(row=n, column=0, sticky="ew", padx=(0, 8), pady=2)
        r["職"].set(職)
        r["科"] = 介.下拉(表, 目前科目(), 寬=10)     # 從上面的科目選
        r["科"].configure(textvariable=r["科v"])
        r["科"].grid(row=n, column=1, sticky="ew", padx=(0, 8), pady=2)
        r["狀"] = tk.Label(表, font=介.字["小"], bg=介.CARD_BG, anchor="w", width=16)
        r["狀"].grid(row=n, column=2, sticky="w")
        刪 = 介.文字按鈕(表, "✕", lambda: 刪除(r), 移上色=介.ERR)
        刪.grid(row=n, column=3)
        r["元件"] = [r["職"], r["科"], r["狀"], 刪]
        r["科v"].trace_add("write", lambda *_: 更新(r))
        for 事件 in ("<<ComboboxSelected>>", "<FocusOut>"):
            r["職"].bind(事件, lambda e: 更新(r))
        列們.append(r)
        更新(r)

    def 刪除(r):
        for w in r["元件"]:
            w.destroy()
        列們.remove(r)

    if 有計算表:                           # 把存的設定放進 ① 科目表
        用過 = set()
        for 名, 科 in 共.拆對照(設.get("計算表分組", "")):
            數 = re.match(r"\d*", 科).group(0)       # 編號開頭是照科目的數字（6110-研發 跟 6110 一起編）
            頭碼 = 開頭.get(科, 開頭.get(數, "") if 數 not in 用過 else "")
            用過.add(數)
            加科(科, 名, 頭碼)
    for 職, 科 in 設.get("職稱科目", []):
        加(職, 科)

    def 收集(設):
        設["職稱科目"] = [[r["職"].get().strip(), r["科v"].get().strip()] for r in 列們 if r["職"].get().strip()]
        if 有計算表:                       # 表格 → 原本的寫法
            分組, 開們 = [], []
            for r in 科列們:
                科 = r["科"].get().strip()
                if not 科:
                    continue
                分組.append(f"{r['名'].get().strip() or 科}={科}")
                if r["頭"].get().strip():
                    數 = re.match(r"\d*", 科).group(0) or 科
                    if 數 not in [x.split("=")[0] for x in 開們]:
                        開們.append(f"{數}={r['頭'].get().strip()}")
            設["計算表分組"] = "、".join(分組)
            設["編號開頭"] = "、".join(開們)
            設["科目寫入欄"] = 寫入欄.get().strip()

    def 載入():                         # 從月薪表數每個職稱有幾個人；薪資檔有、清單沒有的職稱自動加一列
        人數.clear()
        年, wb, 設 = 介.目前年(), 介.資料["wb"], 介.收集()
        if wb is not None and 年 and 共.拆詞(設.get("薪資總額", "")):
            計 = OrderedDict()
            for m in range(1, 13):
                ws = 共.找工作表(wb, 共.表名(設["月薪表"], 年, m))
                t = 共.月薪資料表(ws, 設) if ws is not None else None
                if t is None or not t.標題.get("職稱"):
                    continue
                for rr in range(t.資料起始列, ws.max_row + 1):
                    n = 共.清字(ws.cell(rr, t.姓名欄).value)
                    if 共.是人名(n):
                        計.setdefault(共.清字(ws.cell(rr, t.標題["職稱"]).value), set()).add(n)
            人數.update((k, len(v)) for k, v in 計.items() if k)
        已有 = {r["職"].get().strip() for r in 列們}
        for 職 in 人數:
            if 職 not in 已有:
                加(職, "")
        for r in 列們:
            r["職"].configure(values=list(人數))
        更新全部()

    def 樣板():
        if 有計算表:
            科欄.configure(values=[""] + 介.資料["樣板左欄"])

    def 狀態():
        缺 = sum(1 for r in 列們 if r["職"].get().strip() and not r["科v"].get().strip())
        return (f"⚠ {缺} 個職稱沒有科目", 介.WARN) if 缺 else (f"{len(列們)} 個職稱", 介.TEXT_GRAY)
    return {"收集": 收集, "載入": 載入, "樣板": 樣板, "狀態": 狀態}


# ================================================================ 把好幾張卡片的鉤子合在一起
def 合併(*鉤們, 狀態=None):
    結果 = {}
    for 名 in ("收集", "樣板", "載入", "年改了"):
        fs = [h[名] for h in 鉤們 if h and 名 in h]
        if fs:
            結果[名] = (lambda fs: lambda *a: [f(*a) for f in fs] and None)(fs)
    if 狀態:
        結果["狀態"] = 狀態
    return 結果


# ================================================================ 通用的頁面
def 頁面(介, 設):
    tk = 介.tk
    月頁 = 介.加頁("月薪", "月薪表")
    介.區塊標題(月頁, "月薪表")
    月 = 月薪卡(介, 月頁, 設)

    其頁 = 介.加頁("其他", "其他表")
    其 = {}
    介.區塊標題(其頁, "其他表（前期薪資、年終、各種獎金）",
             右邊=lambda 框: 介.文字按鈕(框, "＋ 新增", lambda: 其["加"](), 色=介.PRIMARY_DARK, 移上色=介.TEXT,
                                     bg=介.BG).pack(side="right"))
    其.update(其他表卡(介, 其頁, 設, 固定欄=("前期薪資", "上期年終", "前期年終", "端午獎金", "中秋獎金", "本期年終")))

    進頁 = 介.加頁("進階", "進階")
    介.區塊標題(進頁, "進階")
    卡 = 介.卡片(進頁)
    卡.columnconfigure(1, weight=1)
    獎稅 = tk.StringVar(value=設.get("獎金扣繳詞", ""))
    執, 科 = tk.BooleanVar(value=bool(設.get("執行業務另開"))), tk.BooleanVar(value=bool(設.get("科目分類")))
    介.標籤(卡, "獎金扣繳詞").grid(row=0, column=0, sticky="w", padx=(0, 14), pady=4)
    框, _ = 介.輸入框(卡, 獎稅)
    框.grid(row=0, column=1, sticky="ew", pady=4)
    介.小字(卡, "獎金表標題含這些字的欄算扣繳稅額（用、分隔）；空白＝獎金不讀扣繳").grid(row=1, column=1, sticky="w")
    介.ttk.Checkbutton(卡, text="執行業務另外放一張「執行業務」工作表（薪資表上寫到「執行業務」的列）",
                      variable=執, style="奶茶.TCheckbutton").grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
    介.ttk.Checkbutton(卡, text="依職稱分會計科目（薪資檔另存一份插入科目計算表，清冊每個科目一個檔）",
                      variable=科, style="奶茶.TCheckbutton", command=lambda: 顯示科目()).grid(
        row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))
    科框 = tk.Frame(進頁, bg=介.BG)
    職 = 職稱卡(介, 科框, 設)
    職載入 = 職["載入"]
    職["載入"] = lambda: 職載入() if 科.get() else None      # 有勾「分科目」才去數職稱

    def 顯示科目():
        if 科.get():
            科框.pack(fill="x", pady=(10, 0))
            職載入()
        else:
            科框.pack_forget()
    顯示科目()

    def 收集進階(設):
        設.update(獎金扣繳詞=獎稅.get().strip(), 執行業務另開=bool(執.get()), 科目分類=bool(科.get()))

    def 狀態():
        return {"月薪": 月["狀態"](), "其他": 其["狀態"](),
                "進階": ("、".join(x for x, v in (("執行業務分開", 執), ("分科目", 科)) if v.get()) or "沒有開", 介.TEXT_GRAY)}
    return 合併(月, 其, 職, {"收集": 收集進階}, 狀態=狀態)
