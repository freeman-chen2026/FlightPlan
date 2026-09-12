import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta

st.set_page_config(page_title="Arinc 飞行计划辅助脚本生成器", layout="centered")
st.title("✈️ Arinc 飞行计划辅助脚本生成器")
st.caption("上传 Excel 航段表 → 生成 检查单文本 / PRELIM·PACKAGE 文本 / JS 脚本")

uploaded_file = st.file_uploader("📂 上传航班计划 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)

        # ---------- 列定位 ----------
        col_aircraft = None
        col_origin = None
        col_dest = None
        col_date = None
        col_time = None
        col_flight_no = None
        col_purpose = None
        col_dep_city = None
        col_arr_city = None
        col_arr_date = None
        col_arr_time = None

        aircraft_keywords = ['飞机注册号', '注册号', '机号', 'Aircraft', 'Tail']
        origin_keywords = ['出发地', '起飞机场', 'Origin', 'Departure', '机场四字码(起)']
        dest_keywords = ['到达地', '目的地机场', 'Dest', 'Arrival', '机场四字码(到)']
        date_keywords = ['出发日期', '起飞日期', 'Departure Date']
        time_keywords = ['计划出发', '起飞时间', '出发时间', 'Departure Time', 'ETD']
        flight_no_keywords = ['航班号', 'Flight No', 'Flight Number']
        purpose_keywords = ['用途', 'Purpose']
        dep_city_keywords = ['出发城市', 'Departure City']
        arr_city_keywords = ['到达城市', 'Arrival City']
        arr_date_keywords = ['到达日期', 'Arrival Date']
        arr_time_keywords = ['预计到达', '到达时间', 'Arrival Time', 'ETA']

        for col in df.columns:
            col_str = str(col).strip()
            if any(kw in col_str for kw in aircraft_keywords) and col_aircraft is None:
                col_aircraft = col
            elif any(kw in col_str for kw in origin_keywords) and col_origin is None:
                col_origin = col
            elif any(kw in col_str for kw in dest_keywords) and col_dest is None:
                col_dest = col
            elif any(kw in col_str for kw in date_keywords) and col_date is None:
                col_date = col
            elif any(kw in col_str for kw in time_keywords) and col_time is None:
                col_time = col
            elif any(kw in col_str for kw in flight_no_keywords) and col_flight_no is None:
                col_flight_no = col
            elif any(kw in col_str for kw in purpose_keywords) and col_purpose is None:
                col_purpose = col
            elif any(kw in col_str for kw in dep_city_keywords) and col_dep_city is None:
                col_dep_city = col
            elif any(kw in col_str for kw in arr_city_keywords) and col_arr_city is None:
                col_arr_city = col
            elif any(kw in col_str for kw in arr_date_keywords) and col_arr_date is None:
                col_arr_date = col
            elif any(kw in col_str for kw in arr_time_keywords) and col_arr_time is None:
                col_arr_time = col

        # 按位置兜底
        def get_col(idx):
            return df.columns[idx] if len(df.columns) > idx else None

        if col_aircraft is None: col_aircraft = get_col(2)
        if col_origin is None:   col_origin = get_col(10)
        if col_dest is None:     col_dest = get_col(12)
        if col_date is None:     col_date = get_col(6)
        if col_time is None:     col_time = get_col(7)
        if col_flight_no is None: col_flight_no = get_col(1)
        if col_purpose is None:  col_purpose = get_col(3)
        if col_dep_city is None: col_dep_city = get_col(11)
        if col_arr_city is None: col_arr_city = get_col(13)
        if col_arr_date is None: col_arr_date = get_col(14)
        if col_arr_time is None: col_arr_time = get_col(15)

        if any(v is None for v in [col_aircraft, col_origin, col_dest]):
            st.error("Excel 列数不足，至少需要包含 C(飞机号)、K(出发地)、M(到达地) 三列。")
            st.stop()

        # ---------- 日期/时间解析 ----------
        MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                  'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

        def parse_date_to_dt(d):
            if d is None or pd.isna(d):
                return None
            if isinstance(d, str):
                d = d.strip()
                for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d %H:%M:%S'):
                    try:
                        return datetime.strptime(d, fmt)
                    except ValueError:
                        continue
                return None
            if hasattr(d, 'year') and hasattr(d, 'month') and hasattr(d, 'day'):
                return datetime(d.year, d.month, d.day)
            return None

        def parse_time_to_str(t):
            if t is None or pd.isna(t):
                return None
            if isinstance(t, str):
                t = t.strip()
                if ':' in t:
                    parts = t.split(':')
                    try:
                        return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
                    except (ValueError, IndexError):
                        return None
                return None
            if hasattr(t, 'hour') and hasattr(t, 'minute'):
                return f"{t.hour:02d}:{t.minute:02d}"
            return None

        def get_utc_date_str(date_val, time_val):
            """北京时间 → UTC 日期字符串 DDMMM"""
            dt_date = parse_date_to_dt(date_val)
            if dt_date is None:
                return None
            tm = parse_time_to_str(time_val)
            if tm:
                h, m = tm.split(':')
                dt_bj = datetime(dt_date.year, dt_date.month, dt_date.day, int(h), int(m))
            else:
                dt_bj = dt_date
            dt_utc = dt_bj - timedelta(hours=8)
            return f"{dt_utc.day:02d}{MONTHS[dt_utc.month - 1]}"

        # ---------- 提取航段（PRELIM/PACKAGE + JS 用） ----------
        flights = []
        for idx, row in df.iterrows():
            aircraft = row[col_aircraft]
            origin = row[col_origin]
            dest = row[col_dest]
            date_val = row[col_date] if col_date is not None else None
            time_val = row[col_time] if col_time is not None else None

            if pd.notna(aircraft) and pd.notna(origin) and pd.notna(dest):
                aircraft_str = str(aircraft).strip()
                origin_str = str(origin).strip().upper()
                dest_str = str(dest).strip().upper()
                date_str = get_utc_date_str(date_val, time_val)

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
            st.error("❌ 未读取到有效航段数据。")
            st.stop()

        st.success(f"✅ 成功解析 **{len(flights)}** 条有效航段")

        # ============================================================
        #  新增：检查单文本
        # ============================================================
        st.divider()
        st.subheader("📋 检查单文本（可粘贴到腾讯文档）")

        checklist_items = []
        for idx, row in df.iterrows():
            try:
                aircraft = row[col_aircraft]
                if pd.isna(aircraft):
                    continue
                ac_str = str(aircraft).strip()
                # 跳过表头 / 含中文
                if any('\u4e00' <= ch <= '\u9fff' for ch in ac_str):
                    continue

                dep_time = row[col_time] if col_time is not None else None
                arr_time = row[col_arr_time] if col_arr_time is not None else None
                if pd.isna(dep_time) or pd.isna(arr_time):
                    continue

                dep_time_str = parse_time_to_str(dep_time)
                arr_time_str = parse_time_to_str(arr_time)
                if not dep_time_str or not arr_time_str:
                    continue

                dep_date = row[col_date] if col_date is not None else None
                arr_date = row[col_arr_date] if col_arr_date is not None else None
                dep_city = row[col_dep_city] if col_dep_city is not None else None
                arr_city = row[col_arr_city] if col_arr_city is not None else None
                purpose = row[col_purpose] if col_purpose is not None else None
                flight_no = row[col_flight_no] if col_flight_no is not None else None

                dep_city_str = str(dep_city).strip() if pd.notna(dep_city) else ""
                arr_city_str = str(arr_city).strip() if pd.notna(arr_city) else ""
                purpose_str = str(purpose).strip() if pd.notna(purpose) else ""
                flight_no_str = str(flight_no).strip() if pd.notna(flight_no) else ""

                dep_dt = parse_date_to_dt(dep_date)
                arr_dt = parse_date_to_dt(arr_date)
                day_diff = 0
                if dep_dt and arr_dt:
                    day_diff = (arr_dt.date() - dep_dt.date()).days
                plus = f" +{day_diff}" if day_diff > 0 else ""

                line1 = f"{ac_str} {dep_time_str} - {arr_time_str}{plus}"
                line2 = f"{dep_city_str} - {arr_city_str}"

                # 调机加 F 前缀
                if "调机" in purpose_str:
                    content = f"F\n{line1}\n{line2}"
                else:
                    content = f"{line1}\n{line2}"

                checklist_items.append({
                    "aircraft": ac_str,
                    "content": content,
                    "flight_no": flight_no_str,
                })
            except Exception:
                continue

        if not checklist_items:
            st.warning("⚠️ 未能生成检查单文本。")
        else:
            st.caption(f"共 {len(checklist_items)} 条，每条可单独复制；底部可一次性复制全部。")

            # 逐条显示
            for i, item in enumerate(checklist_items):
                st.markdown(f"**{i+1}. {item['aircraft']}**")
                st.code(item["content"], language="text")

            # 批量复制
            st.markdown("#### 📦 批量复制（全部计划）")
            all_checklist = "\n\n".join(item["content"] for item in checklist_items)
            st.code(all_checklist, language="text")

        # ============================================================
        #  PRELIM / PACKAGE 文本
        # ============================================================
        st.divider()

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

        preferred_order = [
            "B652Q", "B652S", "B65AP", "MLLIN", "N88AY",
            "B3926", "B652R", "B8105", "B8160", "B8262",
            "B8292", "B8309", "N2QE", "N328LM", "N550DR",
            "N577QT", "N7777U", "N777ZH", "T7178HT", "T7CJK",
            "VPCSZ", "VPCVA"
        ]

        sorted_order = [ac for ac in preferred_order if ac in grouped]
        for ac in order:
            if ac not in sorted_order:
                sorted_order.append(ac)
        order = sorted_order

        if order:
            for ac in order:
                items = grouped[ac]
                prelims = [f"PRELIM {f['aircraft']} {f['origin']}-{f['dest']} {f['date']}" for f in items]
                packages = [f"PACKAGE {f['aircraft']} {f['origin']}-{f['dest']} {f['date']}" for f in items]
                block = "\n".join(prelims + packages)

                st.markdown(f"### {ac}")
                st.code(block, language="text")

        # ============================================================
        #  JavaScript 脚本
        # ============================================================
        st.divider()

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
