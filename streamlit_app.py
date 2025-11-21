import streamlit as st
from pathlib import Path
from io import BytesIO
import tempfile
import sys

import switch_outbound  # 同一目录下的脚本


def run_switch_outbound(in_path: Path, out_path: Path, mapping_paths: list[Path]) -> None:
    """
    调用已有的 switch_outbound.main() 来生成出库单。
    通过临时修改 sys.argv 的方式，把上传的文件路径传给脚本。
    """
    old_argv = sys.argv
    try:
        sys.argv = [
            "switch_outbound.py",
            str(in_path),
            str(out_path),
            *[str(p) for p in mapping_paths],
        ]
        switch_outbound.main()
    finally:
        sys.argv = old_argv


# ----------------- Streamlit UI -----------------

st.set_page_config(page_title="福袋处理工具", layout="centered")

st.title("福袋处理工具")

st.markdown(
    """
### 使用前注意事项

1. **检查商品名关键字**  
   请确认从 CM 下载的订单表格中，**商品名** 内是否包含以下关键字之一：  
   `Switch2`、`Switch強化版`、`Switch有機EL`。  
   如果没有或关键字不一致，请先手动替换后再上传，否则会读取失败。

2. **检查「商品情報１」「商品情報２」**  
   - 「商品情報１」：需要填入对应的 **颜色或型号**（例如 国内専用 / マリオカート / ネオン 等）。  
   - 「商品情報２」：需要填入 **游戏盘的编号**（例如 `[1]`、`[2]` ……）。

3. **Switch 型号对应关键字**

   **Switch2**  
   - 国内専用  
   - マリオカート  
   - LEGENDS  

   **Switch強化版**  
   - ネオン  
   - グレー  

   **Switch有機EL**  
   - ホワイト  
   - ネオン  

以上三点请确认无误后，再上传订单文件生成出库单。
"""
)

st.divider()

# 上传订单文件
uploaded = st.file_uploader("上传订单文件（CSV 或 Excel）", type=["csv", "xlsx", "xls"])

# 固定的三个游戏盘映射表（跟仓库里的文件名一致）
base_dir = Path(__file__).parent
mapping_files = [
    base_dir / "Switch2.csv",
    base_dir / "Switch強化版.csv",
    base_dir / "Switch有機EL.csv",
]

# 检查映射表是否存在
missing = [p.name for p in mapping_files if not p.exists()]
if missing:
    st.error(
        "以下映射文件未找到，请确认它们与本应用在同一文件夹内：\n\n"
        + "\n".join(f"- {name}" for name in missing)
    )

generate_clicked = st.button("生成出库单", disabled=uploaded is None or bool(missing))

if generate_clicked and uploaded is not None and not missing:
    with st.spinner("正在生成出库单，请稍候……"):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # 保存上传的订单到临时文件
            in_path = tmpdir / uploaded.name
            with open(in_path, "wb") as f:
                f.write(uploaded.getbuffer())

            # 生成的出库单路径
            out_path = tmpdir / "出库明细.xlsx"

            # 调用已有脚本进行处理
            try:
                run_switch_outbound(in_path, out_path, mapping_files)
            except Exception as e:
                st.error("生成出库单时出错，请检查终端日志")
                st.exception(e)
            else:
                # 读出生成的 Excel，准备下载
                data = out_path.read_bytes()
                buffer = BytesIO(data)

                default_name = f"出库明细_{Path(uploaded.name).stem}.xlsx"

                st.success("出库单生成完成！")
                st.download_button(
                    "下载出库单 Excel",
                    data=buffer,
                    file_name=default_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
