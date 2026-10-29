# -*- coding: utf-8 -*-
"""
客戶模組：粲紘／連宜
====================
粲紘貿易、連宜兩間公司的薪資放在同一個 Excel，每個月一個 Excel（例如 11501.xls、115.02.xlsx）。
  ・有的月份兩間公司疊在同一張工作表（上面粲紘、下面連宜），有的月份分成兩張表
    → 不看工作表名稱，看每一段上面的「公司名稱」那一列來分公司
  ・每一段：公司名稱 → 標題（例如「115年02月份薪資印領清冊」）→ 欄位標題 → 每個人 → 合計
  ・標題是「○年○○月份薪資印領清冊」的是月薪；其他（特休、績效、尾牙、年終金、長照補助…）先不讀，
    列在匯入檢查的「請確認」裡（之後決定好放哪一欄再加）
  ・薪資總額＝應稅總額、退休自提＝勞退自提金額、扣繳稅額＝代扣稅額；每段都跟「合計」列驗算

兩間公司各一個客戶（各出一份清冊），用同一個模組，靠「公司名稱」分開。
"""
import os
import re
import 清冊共用 as 共

名稱 = "粲紘／連宜"
說明 = "每月一個 Excel、兩間公司放在一起；依公司名稱分開讀"
預設客戶 = ["粲紘貿易", "連宜"]
預設設定 = {
    "公司名稱": "",              # 空白＝用客戶名稱去找（粲紘貿易、連宜）
    "薪資總額": "應稅總額", "退休自提": "勞退自提金額", "扣繳稅額": "代扣稅額",   # 給畫面、樣板看的（實際讀法在下面）
    "執行業務另開": False,
}
# 欄位標題（去掉空白換行後）含這些字 → 讀這一欄
欄位認法 = {
    "姓名": ["員工姓名", "姓名"],
    "編號": ["編號"],
    "總額": ["應稅總額"],
    "自提": ["勞退自提金額", "自提"],
    "稅": ["代扣稅額"],
}
月薪標題 = re.compile(r"(\d{2,3})年(\d{1,2})月份?薪資印領清冊")


def 公司(設):
    return 共.清字(設.get("公司名稱")) or 共.清字(設.get("客戶"))


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
            段["人們"].append({"名": 名, "編": 值們[欄["編號"]] if "編號" in 欄 else "",
                             "總額": 取("總額"), "自提": 取("自提"), "稅": 取("稅"), "列": r})
    return 段們


def 掃描(資料夾):
    段們, 錯 = [], []
    for f in sorted(os.listdir(資料夾)):
        if f.startswith(("~$", ".")) or not f.lower().endswith((".xls", ".xlsx", ".xlsm")):
            continue
        try:
            段們 += 掃描檔案(os.path.join(資料夾, f))
        except SystemExit as e:
            錯.append(f"{f}：{e.code}")
        except Exception as e:
            錯.append(f"{f}：{type(e).__name__} {e}")
    return 段們, 錯


def 資料夾(設):
    return os.path.dirname(設.get("薪資檔") or "")


# ---------------------------------------------------------------- 讀取（產生清冊時共用核心會叫這個）
def 讀取(設, 年, wb, wbf, 清冊欄):
    表們, 缺表, 警告 = [], [], []
    段們, 錯 = 掃描(資料夾(設))
    警告 += [f"這個檔案讀不到：{x}" for x in 錯]
    我的 = [s for s in 段們 if 公司(設) in s["公司"]]
    if not 我的:
        raise SystemExit(f"「{資料夾(設)}」裡的 Excel 都找不到公司名稱含「{公司(設)}」的段落，請檢查「公司名稱」")
    月薪 = sorted([s for s in 我的 if s["年"] == 年 and s["月"]], key=lambda s: s["月"])
    看過 = {}
    for s in 月薪:
        欄位 = 共.中文月[s["月"] - 1]
        if s["月"] in 看過:
            警告.append(f"{年}年{s['月']}月的月薪出現兩次：{看過[s['月']]}、{s['檔']}「{s['表']}」，兩份都讀進來了，請確認")
        看過[s["月"]] = f"{s['檔']}「{s['表']}」"
        缺 = [k for k in ("總額", "自提", "稅") if k not in s["欄"]]
        if 缺:
            警告.append(f"{s['檔']}「{s['表']}」找不到欄位：{'、'.join({'總額': '應稅總額', '自提': '勞退自提金額', '稅': '代扣稅額'}[k] for k in 缺)}")
        來源 = 共.資料來源(f"{年}年{s['月']}月", 欄位, "月薪")
        來源.顯示名稱 = f"{s['檔']}「{s['表']}」{s['公司']}"
        來源.總額, 來源.自提, 來源.扣繳 = ["應稅總額"], "勞退自提金額", ["代扣稅額"]
        for p in s["人們"]:
            來源.加(p["名"], 總額=p["總額"], 自提=p["自提"], 稅=p["稅"], 編=p["編"])
        if s["合計"] is None:
            警告.append(f"{來源.顯示名稱}：找不到「合計」列，沒有驗算")
        else:
            for k in ("總額", "自提", "稅"):
                來源.對合計(k, s["合計"][k])
        表們.append(來源)
    其他 = [s for s in 我的 if not (s["年"] == 年 and s["月"])]
    for s in 其他:
        總 = sum(p["總額"] for p in s["人們"])
        警告.append(f"還沒設定、沒有放進清冊：{s['檔']}「{s['表']}」{s['標題'] or '（沒有標題）'}，"
                  f"{len(s['人們'])} 人，應稅總額 {總:,}")
    return 表們, 缺表, 警告


# ---------------------------------------------------------------- 畫面：檔案頁下面多一張卡片
def 頁面(介, 設):
    tk = 介.tk
    介.區塊標題(介.檔案頁, "公司與每月檔案")
    卡 = 介.卡片(介.檔案頁)
    卡.columnconfigure(1, weight=1)
    名變 = tk.StringVar(value=設.get("公司名稱", ""))
    介.標籤(卡, "公司名稱").grid(row=0, column=0, sticky="w", padx=(0, 14), pady=4)
    框, e = 介.輸入框(卡, 名變)
    框.grid(row=0, column=1, sticky="ew", pady=4)
    介.小字(卡, f"空白＝用客戶名稱（{設.get('客戶')}）去找。薪資檔每一段上面的公司名稱含這幾個字的，才算這個客戶的").grid(
        row=1, column=1, sticky="w")
    介.小字(卡, "「薪資檔」選任何一個月的 Excel 就好：同一個資料夾裡每個月的 Excel 都會讀（.xls、.xlsx 都可以）").grid(
        row=2, column=1, sticky="w", pady=(6, 0))
    結果 = 介.小字(卡, 寬=600)
    結果.grid(row=3, column=1, sticky="w", pady=(6, 0))
    狀 = {"文": "", "色": 介.TEXT_GRAY}

    def 更新(e=None):
        設2 = 介.收集()
        夾, 年 = 資料夾(設2), 介.目前年()
        if not 夾 or not os.path.isdir(夾) or not 年:
            結果.configure(text="", fg=介.TEXT_GRAY); 狀.update(文="", 色=介.TEXT_GRAY); return
        段們, 錯 = 掃描(夾)
        我的 = [s for s in 段們 if 公司(設2) in s["公司"]]
        月們 = sorted({s["月"] for s in 我的 if s["年"] == 年 and s["月"]})
        其他 = [s for s in 我的 if not (s["年"] == 年 and s["月"])]
        if not 我的:
            文 = f"✗ 資料夾裡找不到「{公司(設2)}」"
            結果.configure(text=文, fg=介.ERR); 狀.update(文=文, 色=介.ERR); return
        行 = [f"✓ {年} 年月薪：{'、'.join(f'{m}月' for m in 月們) or '（沒有）'}"]
        if 其他:
            行.append("⚠ 還沒設定的獎金表（先不放清冊）：" + "、".join(f"{s['表']}" for s in 其他))
        if 錯:
            行.append("✗ 讀不到：" + "、".join(錯))
        結果.configure(text="\n".join(行), fg=介.WARN if 其他 or 錯 else 介.OK)
        狀.update(文=f"✓ {len(月們)} 個月", 色=介.OK if 月們 else 介.WARN)
    e.bind("<FocusOut>", 更新)
    e.bind("<Return>", 更新)

    def 收集(設):
        設["公司名稱"] = 名變.get().strip()
    return {"收集": 收集, "載入": 更新, "年改了": 更新}
