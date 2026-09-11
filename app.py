import streamlit as st
import pandas as pd
import json
from datetime import datetime

st.set_page_config(page_title="Arinc 飞行计划辅助脚本生成器", layout="centered")
st.title("✈️ Arinc 飞行计划辅助脚本生成器")
st.caption("上传 Excel 航段表 → 生成 PRELIM·PACKAGE 文本 / JS 脚本")

st.info(
    "⚠️ **合规说明**：本工具生成的脚本仅用于**辅助填写**表单字段（飞机号、起飞机场、目的地机场），"
    "**不会自动点击提交**。每次填充后需您手动点击提交按钮，完全符合“浏览器辅助”安全规范。"
)

uploaded_file = st.file_uploader("📂 上传航班计划 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)

        # ---------- 智能列名匹配 ----------
        col_aircraft = None
        col_origin = None
        col_dest = None
        col_date = None

        aircraft_keywords = ['飞机注册号', '注册号', '机号', 'Aircraft', 'Tail']
        origin_keywords = ['出发地', '起飞机场', 'Origin', 'Departure', '机场四字码(起)']
        dest_keywords = ['到达地', '目的地机场', 'Dest', 'Arrival', '机场四字码(到)']
        date_keywords = ['出发日期', '起飞日期', 'Departure Date', 'Date']

        for col in df.columns:
            col_upper = str(col).strip()
            if any(kw in col_upper for kw in aircraft_keywords):
                col_aircraft = col
            elif any(kw in col_upper for kw in origin_keywords):
                col_origin = col
            elif any(kw in col_upper for kw in dest_keywords):
                col_dest = col
            elif any(kw in col_upper for kw in date_keywords):
                col_date = col

        if col_aircraft is None or col_origin is None or col_dest is None:
            st.warning(f"未通过列名匹配，将按位置读取：C列(飞机号)、K列(出发地)、M列(到达地)、G列(出发日期)。\n当前表头：{list(df.columns)}")
            col_aircraft = df.columns[2] if len(df.columns) > 2 else None
            col_origin = df.columns[10] if len(df.columns) > 10 else None
            col_dest = df.columns[12] if len(df.columns) > 12 else None

        if col_date is None:
            col_date = df.columns[6] if len(df.columns) > 6 else None

        if any(v is None for v in [col_aircraft, col_origin, col_dest]):
            st.error("Excel 列数不足，至少需要包含 C(飞机号)、K(出发地)、M(到达地) 三列。")
            st.stop()

        # ---------- 日期格式化工具 ----------
        MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                  'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

        def format_date(d):
            if d is None or pd.isna(d):
                return None
            if isinstance(d, str):
                d = d.strip()
                for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
                            '%Y-%m-%d %H:%M:%S'):
                    try:
                        d = datetime.strptime(d, fmt)
                        break
                    except ValueError:
                        continue
            if hasattr(d, 'day') and hasattr(d, 'month'):
                return f"{d.day:02d}{MONTHS[d.month - 1]}"
            return None

        # ---------- 提取并清洗数据 ----------
        flights = []
        for idx, row in df.iterrows():
            aircraft = row[col_aircraft]
            origin = row[col_origin]
            dest = row[col_dest]
            date_val = row[col_date] if col_date is not None else None

            if pd.notna(aircraft) and pd.notna(origin) and pd.notna(dest):
                aircraft_str = str(aircraft).strip()
                origin_str = str(origin).strip().upper()
                dest_str = str(dest).strip().upper()
                date_str = format_date(date_val)

                if (len(origin_str) == 4 and origin_str.isalpha() and
                    len(dest_str) == 4 and dest_str.isalpha() and
                    not any('\u4e00' <= ch <= '\u9fff' for ch in aircraft_str)):
                    flights.append({
                        "aircraft": aircraft_str,
                        "origin": origin_str,
                        "dest": dest_str,
                        "date": date_str
                    })

        if not flights:
            st.error("❌ 未读取到有效航段数据，请检查列是否包含正确的四字码和注册号。")
            st.stop()

        st.success(f"✅ 成功解析 **{len(flights)}** 条有效航段")

        # ============================================================
        #  功能二（先显示）：PRELIM / PACKAGE 文本（按飞机号分组）
        # ============================================================
        st.header("📝 PRELIM / PACKAGE 文本")
        st.caption("纯文本，手动复制粘贴到 ARINCDirect。按飞机号分组，先列全部 PRELIM，再列全部 PACKAGE。")

        # 按飞机号分组（保持出现顺序）
        grouped = {}
        order = []
        for f in flights:
            if not f["date"]:
                continue
            ac = f["aircraft"]
            if ac not in grouped:
                grouped[ac] = []
                order.append(ac)
            grouped[ac].append(f)

        if not order:
            st.warning("⚠️ 未能生成 PRELIM/PACKAGE 文本，可能是日期列无法解析。")
        else:
            for ac in order:
                items = grouped[ac]
                prelims = [f"PRELIM {f['aircraft']} {f['origin']}-{f['dest']} {f['date']}" for f in items]
                packages = [f"PACKAGE {f['aircraft']} {f['origin']}-{f['dest']} {f['date']}" for f in items]
                block = "\n".join(prelims + packages)

                st.markdown(f"### {ac}")
                st.code(block, language="text")

        # ============================================================
        #  功能一（后显示）：JavaScript 脚本
        # ============================================================
        st.divider()
        st.header("🧩 JavaScript 填表脚本")
        st.caption("在 Arinc 页面控制台粘贴执行，逐条填充。请自行评估合规风险。")

        flights_json = json.dumps(flights, ensure_ascii=False, indent=2)

        js_template = """
// ============================================================
//  Arinc 飞行计划自动填表脚本（逐条模式）
//  使用方法：
//  1. 在 Arinc 页面按 F12 打开控制台
//  2. 粘贴以下全部代码并回车执行
//  3. 每次在控制台输入 fillNext() 并回车，自动填入下一条
//  4. 填完后请手动点击提交按钮
// ============================================================

var flightData = __FLIGHT_DATA__;
var currentIndex = 0;

function getElementByXpath(path) {
    return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
}

function fillFlight(index) {
    if (index < 0 || index >= flightData.length) {
        console.error("索引超出范围 (0 ~ " + (flightData.length - 1) + ")");
        return false;
    }
    var data = flightData[index];
    console.log(`🛫 填充第 ${index+1}/${flightData.length} 条: ${data.aircraft}  ${data.origin} -> ${data.dest}`);

    // 1. 飞机下拉框
    var sel = document.getElementById('Aircraft');
    if (!sel) {
        console.error("未找到 id='Aircraft' 的下拉框");
        return false;
    }
    var found = false;
    for (var opt of sel.options) {
        if (opt.value === data.aircraft) {
            opt.selected = true;
            found = true;
            break;
        }
    }
    if (!found) {
        console.warn("飞机号 " + data.aircraft + " 不在下拉列表中，尝试直接设置值");
        sel.value = data.aircraft;
    }
    sel.dispatchEvent(new Event('change', { bubbles: true }));

    // 2. 起飞机场
    var originInput = getElementByXpath('/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/input[1]');
    if (!originInput) {
        console.error("未找到起飞机场输入框，XPath可能已变化");
        return false;
    }
    originInput.focus();
    originInput.value = data.origin;
    originInput.dispatchEvent(new Event('input', { bubbles: true }));
    originInput.dispatchEvent(new Event('change', { bubbles: true }));
    originInput.dispatchEvent(new Event('blur', { bubbles: true }));
    originInput.blur();

    // 3. 目的地机场
    var destInput = getElementByXpath('/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[4]/td[1]/input[1]');
    if (!destInput) {
        console.error("未找到目的地机场输入框，XPath可能已变化");
        return false;
    }
    destInput.focus();
    destInput.value = data.dest;
    destInput.dispatchEvent(new Event('input', { bubbles: true }));
    destInput.dispatchEvent(new Event('change', { bubbles: true }));
    destInput.dispatchEvent(new Event('blur', { bubbles: true }));
    destInput.blur();

    console.log("✅ 填充完成，请手动点击提交按钮！");
    return true;
}

function fillNext() {
    if (currentIndex >= flightData.length) {
        console.log("🎉 所有航段已全部填充完毕！");
        return;
    }
    if (fillFlight(currentIndex)) {
        currentIndex++;
    }
}

function resetIndex() {
    currentIndex = 0;
    console.log("🔄 索引已重置为 0");
}

console.log("✅ 脚本加载成功，共 " + flightData.length + " 条航段");
console.log("📌 在控制台输入 fillNext() 填充下一条，每次填完请手动提交");
console.log("📌 如需重新开始，输入 resetIndex() 重置索引");
"""

        js_script = js_template.replace("__FLIGHT_DATA__", flights_json)

        st.subheader("📜 生成的 JavaScript 脚本（复制以下全部内容）")
        st.code(js_script, language="javascript")
        st.info("💡 点击代码框右上角的复制图标，或手动全选后 Ctrl+C 复制。")

        with st.expander("📖 详细使用步骤"):
            st.markdown("""
            1. 登录 Arinc 并进入飞行计划制作页面。
            2. 按 **F12** → 点击 **Console（控制台）** 标签。
            3. 将生成的 JS 脚本 **全选复制**，粘贴到控制台后按 **回车** 执行。
            4. 看到 `✅ 脚本加载成功` 后，在控制台输入 **`fillNext()`** 并回车。
            5. 脚本会自动填入 **飞机号、起飞机场、目的地机场**。
            6. **你手动点击页面上的“提交”按钮**。
            7. 回到控制台，再次输入 **`fillNext()`** 继续下一条。
            8. 重复步骤 6-7，直至所有航段制作完毕。
            """)

    except Exception as e:
        st.error(f"❌ 处理文件时发生错误：{e}")
        st.stop()
else:
    st.info("👆 请先上传 Excel 文件以生成脚本")

st.divider()
st.caption("⚙️ 部署于 Streamlit Cloud · 仅生成文本与脚本，均不存储任何数据")
