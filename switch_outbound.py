#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
switch_outbound.py  🧡  by 夏以昼  (rev.2025‑11‑21b)
-------------------------------------------------
▲ 更新纪录
  * 2025‑11‑21b ❶ 读 CSV 时自动尝试 UTF‑8 → Shift‑JIS (cp932) 双编码，解决 UnicodeDecodeError。
                 ❷ 输出路径后可跟 **任意多个 keyword_mapping.csv**，自动合并。

用途：
    平台导出单行 Switch 套装订单 → 拆分多行出库表
    输出列：存货编码 / 仓库 / 数量 / 单价 / SN码 / 备注

用法：
    python switch_outbound.py orders.csv 出库.xlsx [mapping1.csv mapping2.csv …]

依赖：pandas >=1.2   (pip install pandas openpyxl)
"""

import sys, os, math, re
import pandas as pd
from pathlib import Path

# ------------------ 固定映射 ------------------
CONSOLE_MAP = {
    "Switch2": {
        "国内専用": "4902370553024",
        "マリオカート": "4902370553031",
        "LEGENDS": "4902370553505",
    },
    "Switch強化版": {
        "ネオン": "4902370550733",
        "グレー": "4902370551198",
    },
    "Switch有機EL": {
        "ホワイト": "4902370548495",
        "ネオン": "4902370548501",
    },
}

ACCESSORY_FIXED = {
    "フィルム": {"jan": "98462", "unit_price": 500},
    "ケース": {"jan": "98463", "unit_price": 500},
}

INLINE_MAPPING = [
    {"keyword": "マリオカート", "jan": "4902370553031", "unit_price": 8000},
    {"keyword": "LEGENDS", "jan": "4902370553505", "unit_price": 8000},
]

# -----------------------------------------------------------
# 工具
# -----------------------------------------------------------

def read_csv_auto(path: str | Path):
    """先尝试 utf‑8，再退到 cp932 (Windows‑31J)"""
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="cp932")


def load_keyword_mappings(paths: list[str]):
    """读取一个或多个关键字映射 CSV，并按主机型号分类
    - 文件名包含 "Switch2" → 属于 "Switch2" 的游戏盘表
    - 文件名包含 "強化"      → 属于 "Switch強化版"
    - 文件名包含 "有機" 或 "EL" → 属于 "Switch有機EL"
    若某个型号没有提供文件，则该型号默认没有游戏盘映射（跳过游戏盘行）。
    """
    # 预先为每个机种准备一个空表
    mapping: dict[str, pd.DataFrame] = {
        key: pd.DataFrame(columns=["keyword", "jan", "unit_price"])
        for key in CONSOLE_MAP.keys()
    }

    if not paths:
        return mapping

    col_alias = {
        "キーワード": "keyword",
        "keyword": "keyword",
        "janコード": "jan",
        "jan": "jan",
        "単価": "unit_price",
        "unit_price": "unit_price",
    }

    for p in paths:
        path = Path(p)
        if not path.exists():
            print(f"⚠️  mapping file not found: {p}  (skip)")
            continue

        df = read_csv_auto(path)
        df.columns = [col_alias.get(c, c) for c in df.columns]
        if "unit_price" in df.columns:
            df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce").fillna(0)

        stem = path.stem  # 文件名（不含后缀）
        stem_lower = stem.lower()
        if "switch2" in stem_lower or "2" in stem:
            key = "Switch2"
        elif "強化" in stem or "kyouka" in stem_lower:
            key = "Switch強化版"
        elif "有機" in stem or "el" in stem:
            key = "Switch有機EL"
        else:
            print(f"⚠️  無法從檔名推斷機種: {stem}  → 請在檔名中包含 Switch2 / 強化 / 有機EL，已略過。")
            continue

        mapping[key] = df

    return mapping


def find_console_type(title: str):
    """根据商品名判断机种 (Switch2 / Switch強化版 / Switch有機EL)"""
    ttl = title.lower() if isinstance(title, str) else ""
    for t in CONSOLE_MAP:
        if t.lower() in ttl:
            return t
    return None


def find_console_model(console_type: str | None, text: str | None):
    """在给定文本中，根据机种表里的关键字做『包含匹配』
    - console_type 先由商品名判定（Switch2 / Switch強化版 / Switch有機EL）
    - text 为 商品情報１ 中的内容，只要包含关键字的一部分即可
    """
    if not console_type or text is None:
        return None
    src = text.lower()
    for kw in CONSOLE_MAP[console_type]:
        if kw.lower() in src:
            return kw
    return None
    ttl = title.lower()
    for kw in CONSOLE_MAP[console_type]:
        if kw.lower() in ttl:
            return kw
    return None

# -----------------------------------------------------------
# 主入口
# -----------------------------------------------------------

def main():
    if len(sys.argv) < 3:
        print("Usage: python switch_outbound.py orders.csv 出库.xlsx [mapping1.csv mapping2.csv …]")
        sys.exit(1)

    orders_path, out_path, *mapping_paths = sys.argv[1:]

    # ---------- 1. 读取订单 ----------
    orders_path = Path(orders_path)
    if not orders_path.exists():
        print("❌ 找不到订单文件:", orders_path)
        sys.exit(1)

    if orders_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        orders = pd.read_excel(orders_path, dtype=str, engine="openpyxl").fillna("")
    else:
        orders = read_csv_auto(orders_path)

    # 列别名映射
    alias = {
        "注文番号": "order_id",
        "order id": "order_id",
        "注文ＩＤ": "order_id",
        "商品名": "title",
        "商品名称": "title",
        # 商品情報１／２ 这类列名的变体很多，下面做基础映射，稍后再做一次自动识别
        "商品情報１": "info1",
        "商品情報1": "info1",
        "商品情報 1": "info1",
        "商品情報２": "info2",
        "商品情報2": "info2",
        "商品情報 2": "info2",
        "数量": "qty",
        "個数": "qty",
        "金額": "amount",
        "合計": "amount",
    }
    # 先用字典做一轮简单映射
    orders.columns = [alias.get(c, c) for c in orders.columns]
    # 再对所有列名做一次模糊识别，凡是包含「商品情報」且带 1/２ 的，都归一为 info1/info2
    new_cols = []
    for c in orders.columns:
        if c in ("info1", "info2", "order_id", "title", "qty", "amount"):
            new_cols.append(c)
            continue
        if "商品情報" in c:
            if any(x in c for x in ["1", "１"]):
                new_cols.append("info1")
                continue
            if any(x in c for x in ["2", "２"]):
                new_cols.append("info2")
                continue
        new_cols.append(c)
    orders.columns = new_cols

    orders["qty"] = pd.to_numeric(orders.get("qty", 1), errors="coerce").fillna(1).astype(int)
    orders["amount"] = pd.to_numeric(orders.get("amount", 0), errors="coerce").fillna(0).astype(float)

    # ---------- 2. 读取 keyword 映射 ----------
    kw_map = load_keyword_mappings(mapping_paths)

    # ---------- 3. 拆分逻辑 ----------
    output_rows = []

    for _, o in orders.iterrows():
        qty       = int(o["qty"])
        total_amt = float(o["amount"])
        remain    = total_amt
        order_id  = o["order_id"]
        title     = o["title"]

        info1 = str(o.get("info1", "")).strip()
        info2 = str(o.get("info2", "")).strip()

        # ① 机种：先用商品名判定 Switch2 / Switch強化版 / Switch有機EL
        console_type = find_console_type(title)

        # ② 型号 / 颜色：只用「商品情報１」做包含匹配
        model_source = info1                      # 即使为空，也不再回退到 title
        console_kw   = find_console_model(console_type, model_source)
        console_jan  = CONSOLE_MAP.get(console_type, {}).get(console_kw, "")

        # ---（游戏盘、壳膜、主机行的逻辑保持原样，这里省略）---

                # 机种：先从商品名里抓 Switch2 / 強化版 / 有機EL
        console_type = find_console_type(title)
        # 型号 / 顔色：只从「商品情報１」中判断，不再回退到商品名
        #   Switch2      → 国内専用 / マリオカート / LEGENDS
        #   Switch強化版 → ネオン / グレー
        #   Switch有機EL → ホワイト / ネオン
        model_source = info1          # 只用「商品情報１」
        console_kw   = find_console_model(console_type, model_source)
        console_jan  = CONSOLE_MAP.get(console_type, {}).get(console_kw, "")

        # 游戏盘：从 商品情報２ 里找匹配 keyword 的“数字/标记”，
        # 且使用對應機種的 CSV（Switch2 / 強化版 / 有機EL 各自不同）
        type_df = kw_map.get(console_type)
        if info2 and type_df is not None and not type_df.empty:
            info2_str = str(info2).strip()
            # 先尝试完全相等
            hit_eq = type_df[type_df["keyword"].astype(str).str.strip() == info2_str]
            if len(hit_eq):
                hit = hit_eq.iloc[0]
            else:
                # 若没命中，再尝试“包含关系”（例如 keyword="[1]"，info2 裡有 "[1]")
                mask = type_df["keyword"].astype(str).apply(lambda k: k in info2_str)
                sub  = type_df[mask]
                hit = sub.iloc[0] if len(sub) else None

            if hit is not None:
                game_jan   = hit["jan"]
                game_price = float(hit.get("unit_price", 0))
                output_rows.append({
                    "存货编码": game_jan, "仓库": "", "数量": qty,
                    "单价": game_price, "SN码": "", "备注": order_id,
                })
                remain -= game_price * qty

        # Switch2 配件
        if console_type == "Switch2":
            for meta in ACCESSORY_FIXED.values():
                output_rows.append({
                    "存货编码": meta["jan"], "仓库": "", "数量": qty,
                    "单价": meta["unit_price"], "SN码": "", "备注": order_id,
                })
                remain -= meta["unit_price"] * qty

        # 主机
        unit_price_console = max(math.floor(remain / qty), 0) if qty else 0
        output_rows.append({
            "存货编码": console_jan, "仓库": "", "数量": qty,
            "单价": unit_price_console, "SN码": "", "备注": order_id,
        })

    # ---------- 4. 写出 ----------
    df_out = pd.DataFrame(output_rows, columns=["存货编码", "仓库", "数量", "单价", "SN码", "备注"])
    df_out.to_excel(out_path, index=False, engine="openpyxl")
    print("✅ 生成完成 -->", out_path)


if __name__ == "__main__":
    main()
