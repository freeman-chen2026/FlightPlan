import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json
import re
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from streamlit_autorefresh import st_autorefresh

st.set_page_config(page_title="飞行任务工具", layout="centered")
st.title("飞行任务工具")

BJ_TZ = timezone(timedelta(hours=8))
def now_bj():
    return datetime.now(BJ_TZ).replace(tzinfo=None)

tab1, tab2 = st.tabs(["功能1：检查单/脚本", "功能2：邮件生成"])


# ============================================================
# 功能1：检查单 / 飞行计划脚本
# ============================================================
with tab1:
    uploaded_file = st.file_uploader("📂 上传航班计划 Excel 文件", type=["xlsx"], key="f1_file")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0)

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
                st.error("Excel 列数不足。")
                st.stop()

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

            def get_bj_date_str(date_val):
                dt_date = parse_date_to_dt(date_val)
                if dt_date is None:
                    return None
                return f"{dt_date.day:02d}{MONTHS[dt_date.month - 1]}"

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
                    date_str_utc = get_utc_date_str(date_val, time_val)
                    date_str_bj = get_bj_date_str(date_val)

                    if (len(origin_str) == 4 and origin_str.isalpha() and
                        len(dest_str) == 4 and dest_str.isalpha() and
                        not any('\u4e00' <= ch <= '\u9fff' for ch in aircraft_str)):
                        flights.append({
                            "aircraft": aircraft_str,
                            "origin": origin_str,
                            "dest": dest_str,
                            "date": date_str_utc,
                            "date_bj": date_str_bj,
                        })

            if not flights:
                st.error("❌ 未读取到有效航段数据。")
                st.stop()

            st.success(f"✅ 成功解析 **{len(flights)}** 条有效航段")

            st.divider()

            preferred_order = [
                "B3926", "B8105", "B8160", "B8262", "B8292", "B8309",
                "N2QE", "N328LM", "N550DR", "N577QT", "N7777U", "N777ZH",
                "T7178HT", "T7CJK", "VPCSZ", "VPCVA",
                "B652R", "N88AY", "MLLIN", "B652Q", "B652S", "B65AP"
            ]
            priority_map = {ac: i for i, ac in enumerate(preferred_order)}
            default_priority = len(preferred_order)

            raw_items = []
            for idx, row in df.iterrows():
                try:
                    aircraft = row[col_aircraft]
                    if pd.isna(aircraft):
                        continue
                    ac_str = str(aircraft).strip()
                    if ac_str.upper() in ('N/A', 'NA', 'NONE', 'NULL', ''):
                        continue
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

                    dep_city_str = str(dep_city).strip() if pd.notna(dep_city) else ""
                    arr_city_str = str(arr_city).strip() if pd.notna(arr_city) else ""
                    purpose_str = str(purpose).strip() if pd.notna(purpose) else ""

                    dep_dt = parse_date_to_dt(dep_date)
                    arr_dt = parse_date_to_dt(arr_date)
                    day_diff = 0
                    if dep_dt and arr_dt:
                        day_diff = (arr_dt.date() - dep_dt.date()).days
                    plus = f" +{day_diff}" if day_diff > 0 else ""

                    line1 = f"{ac_str} {dep_time_str} - {arr_time_str}{plus}"
                    line2 = f"{dep_city_str} - {arr_city_str}"

                    if "调机" in purpose_str:
                        content = f"F\n{line1}\n{line2}"
                    else:
                        content = f"{line1}\n{line2}"

                    raw_items.append({
                        "aircraft": ac_str,
                        "dep_dt": dep_dt,
                        "dep_time_str": dep_time_str,
                        "content": content,
                    })
                except Exception:
                    continue

            raw_items.sort(key=lambda x: (
                x["dep_dt"] if x["dep_dt"] else datetime(2100, 1, 1),
                priority_map.get(x["aircraft"], default_priority),
                x["dep_time_str"] or "99:99"
            ))

            rows = []
            prev_date = None
            prev_ac = None
            first = True
            for it in raw_items:
                cur_date = it["dep_dt"].date() if it["dep_dt"] else None
                cur_ac = it["aircraft"]
                if not first and (cur_date != prev_date or cur_ac != prev_ac):
                    rows.append({"type": "blank"})
                rows.append({"type": "data", "content": it["content"]})
                prev_date = cur_date
                prev_ac = cur_ac
                first = False

            if not rows:
                st.warning("⚠️ 未能生成检查单。")
            else:
                rows_json = json.dumps(rows, ensure_ascii=False)

                components.html(
                    f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                    <style>
                      body {{ font-family: -apple-system, "Segoe UI", sans-serif; margin: 0; }}
                      .btn {{
                        padding: 10px 22px; font-size: 16px; cursor: pointer;
                        background: #ff4b4b; color: white; border: none;
                        border-radius: 6px; font-weight: bold;
                      }}
                      .btn:hover {{ background: #e63939; }}
                      .btn:active {{ transform: translateY(1px); }}
                      .status {{ margin-left: 12px; color: #555; font-size: 14px; }}
                    </style>
                    </head>
                    <body>
                    <button class="btn" id="copyBtn">📋 一键复制全部检查单</button>
                    <span class="status" id="status"></span>

                    <script>
                      const rows = {rows_json};

                      function escapeHtml(s) {{
                        return s.replace(/&/g, '&amp;')
                                .replace(/</g, '&lt;')
                                .replace(/>/g, '&gt;');
                      }}

                      const TD_STYLE = "text-align: center; " +
                                       "vertical-align: middle; " +
                                       "font-family: 'Times New Roman', Times, serif; " +
                                       "font-size: 14pt; " +
                                       "border: 1px solid #000000;";

                      function renderRow(row) {{
                        if (row.type === 'blank') {{
                          return '<tr><td style="' + TD_STYLE + '">&nbsp;</td></tr>';
                        }}
                        return '<tr><td style="' + TD_STYLE + '">' +
                               escapeHtml(row.content).split('\\n').join('<br>') +
                               '</td></tr>';
                      }}

                      document.getElementById('copyBtn').addEventListener('click', async () => {{
                        const html =
                          '<table style="border-collapse: collapse;">' +
                          rows.map(renderRow).join('') +
                          '</table>';
                        const plain = rows.map(r => r.type === 'blank' ? '' : r.content).join('\\n\\n');

                        try {{
                          await navigator.clipboard.write([
                            new ClipboardItem({{
                              'text/html':  new Blob([html],  {{type: 'text/html'}}),
                              'text/plain': new Blob([plain], {{type: 'text/plain'}})
                            }})
                          ]);
                          document.getElementById('status').textContent = '✅ 已复制';
                        }} catch (e) {{
                          document.getElementById('status').textContent =
                            '❌ 复制失败：' + e.message;
                        }}
                      }});
                    </script>
                    </body>
                    </html>
                    """,
                    height=80,
                )

            st.divider()

            date_groups = {}
            for f in flights:
                if not f["date_bj"]:
                    continue
                d = f["date_bj"]
                ac = f["aircraft"]
                if d not in date_groups:
                    date_groups[d] = {}
                if ac not in date_groups[d]:
                    date_groups[d][ac] = []
                date_groups[d][ac].append(f)

            MONTHS_MAP = {m: i for i, m in enumerate(MONTHS)}

            def date_bj_sort_key(d):
                if not d or len(d) < 5:
                    return (99, 0)
                try:
                    day = int(d[:2])
                except ValueError:
                    day = 99
                month = MONTHS_MAP.get(d[2:], 0)
                return (month, day)

            sorted_dates = sorted(date_groups.keys(), key=date_bj_sort_key)

            for date_bj in sorted_dates:
                with st.expander(f"📅 {date_bj}", expanded=False):
                    ac_groups = date_groups[date_bj]
                    sorted_acs = sorted(
                        ac_groups.keys(),
                        key=lambda a: priority_map.get(a, default_priority)
                    )
                    for ac in sorted_acs:
                        items = ac_groups[ac]
                        prelims = [f"PRELIM {f['aircraft']} {f['origin']}-{f['dest']} {f['date']}" for f in items]
                        packages = [f"PACKAGE {f['aircraft']} {f['origin']}-{f['dest']} {f['date']}" for f in items]
                        block = "\n".join(prelims) + "\n\n" + "\n".join(packages)

                        st.markdown(f"**{ac}**")
                        st.code(block, language="text")

            st.divider()

            flights_json = json.dumps(flights, ensure_ascii=False, indent=2)

            js_template = """
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
    console.log("🛫 填充第 " + (index+1) + "/" + flightData.length + " 条: " + data.aircraft + "  " + data.origin + " -> " + data.dest);

    var sel = document.getElementById('Aircraft');
    if (!sel) { console.error("未找到 id='Aircraft'"); return false; }
    var found = false;
    for (var opt of sel.options) {
        if (opt.value === data.aircraft) { opt.selected = true; found = true; break; }
    }
    if (!found) { sel.value = data.aircraft; }
    sel.dispatchEvent(new Event('change', { bubbles: true }));

    var originInput = getElementByXpath('/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/input[1]');
    if (!originInput) { console.error("未找到起飞机场输入框"); return false; }
    originInput.focus();
    originInput.value = data.origin;
    originInput.dispatchEvent(new Event('input', { bubbles: true }));
    originInput.dispatchEvent(new Event('change', { bubbles: true }));
    originInput.dispatchEvent(new Event('blur', { bubbles: true }));
    originInput.blur();

    var destInput = getElementByXpath('/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[4]/td[1]/input[1]');
    if (!destInput) { console.error("未找到目的地机场输入框"); return false; }
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
    if (currentIndex >= flightData.length) { console.log("🎉 全部填充完毕！"); return; }
    if (fillFlight(currentIndex)) { currentIndex++; }
}

function resetIndex() { currentIndex = 0; console.log("🔄 索引已重置为 0"); }

console.log("✅ 脚本加载成功，共 " + flightData.length + " 条航段");
console.log("📌 输入 fillNext() 填充下一条");
"""

            js_script = js_template.replace("__FLIGHT_DATA__", flights_json)

            st.subheader("📜 JavaScript 脚本")
            st.code(js_script, language="javascript")

        except Exception as e:
            st.error(f"❌ 处理文件时发生错误：{e}")
            st.stop()
    else:
        st.info("👆 请先上传 Excel 文件以生成脚本")

    st.divider()
    st.caption("⚙️ 部署于 Streamlit Cloud · 仅生成文本与脚本，均不存储任何数据")


# ============================================================
# 功能2：WX AND NOTAM 邮件生成
# ============================================================
with tab2:
    st_autorefresh(interval=600000, key="f2_refresh")

    st.subheader("WX AND NOTAM 邮件生成器")

    PILOT_MAP = {
        "P001": "gengfan@amber-aviation.com",
        "P002": "zhangyongyi@amber-aviation.com",
        "P003": "fmei@amber-aviation.com",
        "P004": "wangbin@amber-aviation.com",
        "P019": "darranhealy@amber-aviation.com",
        "P020": "thaddeusbeebe@amber-aviation.com",
        "P032": "ericlin@amber-aviation.com",
        "P035": "prjackson@amber-aviation.com",
        "P036": "warrenwang@amber-aviation.com",
        "P038": "johnmiao@amber-aviation.com",
        "P039": "yiftahrauch@amber-aviation.com",
        "P044": "rockli@amber-aviation.com",
        "P046": "zhyszhao@amber-aviation.com",
        "P051": "eugene.peng@humbleholding.com",
        "P052": "brian.wu@humbleholding.com",
        "P053": "brwaines@amber-aviation.com",
        "P054": "rbonetti@amber-aviation.com",
        "P056": "krsherren@amber-aviation.com",
        "P057": "ovracz@amber-aviation.com",
        "P059": "kctsai@amber-aviation.com",
        "P061": "qhli@amber-aviation.com",
        "P065": "wsong@amber-aviation.com",
        "P068": "zjzan@amber-aviation.com",
        "P069": "simoneroeder@amber-aviation.com",
        "P070": "hdstamm@amber-aviation.com",
        "P071": "jasonsun@amber-aviation.com",
        "P072": "zyzhu@amber-aviation.com",
        "P074": "smjin@amber-aviation.com",
        "P075": "eduardroski@amber-aviation.com",
        "P077": "andyliu@amber-aviation.com",
        "P078": "fzhang@amber-aviation.com",
        "P079": "wesleywei@amber-aviation.com",
        "P080": "sliu@amber-aviation.com",
        "P081": "richardwu@amber-aviation.com",
        "P082": "frankliu@amber-aviation.com",
        "P083": "xyou@amber-aviation.com",
        "P084": "ymli@amber-aviation.com",
        "P085": "lzhao@amber-aviation.com",
        "P086": "hxzhang@amber-aviation.com",
        "P087": "hesun@amber-aviation.com",
        "P088": "harryma@amber-aviation.com",
        "P089": "xlli@amber-aviation.com",
        "P090": "hdhuang@amber-aviation.com",
        "P091": "mikema@amber-aviation.com",
        "PJZ001": "zzhang@amber-aviation.com",
        "PJZ002": "charlesguo@amber-aviation.com",
        "PJZ004": "leowang@amber-aviation.com",
        "PJZ005": "evawang@amber-aviation.com",
        "PJZ007": "frankxu@amber-aviation.com",
        "PJZ008": "ariayang@amber-aviation.com",
        "W070": "wang_yanhai@163.com",
        "W213": "cshum@tagaviation.com",
        "W267": "naten7@hotmail.com",
        "W268": "pilotlocalizer@gmail.com",
        "W270": "yang_tao2005@aliyun.com",
        "W272": "Andrew.king@aero.bombardier.com",
    }

    if 'f2_flights' not in st.session_state:
        st.session_state.f2_flights = []

    flight_file = st.file_uploader("上传航段表（Excel 或 CSV）", type=["xlsx", "csv"], key="f2_file")
    plan_text = st.text_area("粘贴文本飞行计划", height=200, key="f2_plan",
                             placeholder="B652Q 06:00 - 07:55\n北京大兴 - 上海虹桥\nP035,P039,C051\n\nB65AP 07:30 - 09:00\n日本东京 羽田 - 日本福冈\nP032,P036,C050,M021")

    gen_btn = st.button("生成邮件链接", key="f2_gen")

    if gen_btn:
        if flight_file is None:
            st.error("请先上传航段表")
        elif not plan_text.strip():
            st.error("请粘贴文本飞行计划")
        else:
            try:
                if flight_file.name.lower().endswith('.csv'):
                    raw = pd.read_csv(flight_file, header=None, dtype=str)
                else:
                    raw = pd.read_excel(flight_file, header=None)

                header_row = None
                for i in range(min(len(raw), 10)):
                    vals = [str(v).strip() for v in raw.iloc[i].values if pd.notna(v)]
                    if '飞机注册号' in vals and '出发地' in vals and '到达地' in vals:
                        header_row = i
                        break

                if header_row is None:
                    st.error("未在航段表中找到标题行（需包含：飞机注册号、出发地、到达地）")
                    st.stop()

                df2 = raw.iloc[header_row + 1:].copy()
                df2.columns = [str(v).strip() for v in raw.iloc[header_row].values]
                df2 = df2.reset_index(drop=True)
            except Exception as e:
                st.error(f"读取航段表失败：{e}")
                st.stop()

            def find_col(df, keywords):
                for c in df.columns:
                    cs = str(c).strip()
                    for kw in keywords:
                        if kw in cs:
                            return c
                return None

            col_flight = find_col(df2, ['飞机注册号', '注册号'])
            col_dep = find_col(df2, ['出发地'])
            col_arr = find_col(df2, ['到达地'])
            col_date = find_col(df2, ['出发日期'])
            col_dep_time = find_col(df2, ['计划出发'])
            col_arr_time = find_col(df2, ['预计到达'])

            if not all([col_flight, col_dep, col_arr, col_date, col_dep_time]):
                st.error(f"航段表缺少必要列。当前列名：{list(df2.columns)}")
            else:
                def to_time(v):
                    if v is None:
                        return None
                    if isinstance(v, float) and pd.isna(v):
                        return None
                    if isinstance(v, (pd.Timestamp, datetime)):
                        return v.strftime('%H:%M')
                    if hasattr(v, 'hour') and hasattr(v, 'minute'):
                        return f"{v.hour:02d}:{v.minute:02d}"
                    s = str(v).strip()
                    if not s or s.lower() == 'nan' or s.lower() == 'nat':
                        return None
                    m = re.match(r'^(\d{1,2}):(\d{2})', s)
                    if m:
                        return f"{int(m.group(1)):02d}:{m.group(2)}"
                    return None

                def to_date(v):
                    if v is None:
                        return None
                    if isinstance(v, float) and pd.isna(v):
                        return None
                    if isinstance(v, (pd.Timestamp, datetime)):
                        return v.strftime('%Y-%m-%d')
                    s = str(v).strip()
                    if not s or s.lower() in ('nan', 'nat'):
                        return None
                    try:
                        return pd.to_datetime(s).strftime('%Y-%m-%d')
                    except Exception:
                        return None

                flight_db = []
                for _, row in df2.iterrows():
                    f_no = str(row[col_flight]).strip() if pd.notna(row[col_flight]) else ''
                    if not f_no or f_no.lower() == 'nan':
                        continue
                    date_str = to_date(row[col_date])
                    dep_t = to_time(row[col_dep_time])
                    arr_t = to_time(row[col_arr_time]) if col_arr_time else None
                    if not date_str or not dep_t:
                        continue
                    flight_db.append({
                        'flight': f_no,
                        'dep': str(row[col_dep]).strip().upper() if pd.notna(row[col_dep]) else '',
                        'arr': str(row[col_arr]).strip().upper() if pd.notna(row[col_arr]) else '',
                        'date': date_str,
                        'dep_time': dep_t,
                        'arr_time': arr_t or ''
                    })

                st.success(f"✅ 航段表解析成功，共 {len(flight_db)} 条航段")

                plan_flights = []
                cur = None
                expect_route = False
                expect_crew = False
                flight_re = re.compile(r'^([A-Z0-9]+)\s+(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})')
                for line in plan_text.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    m = flight_re.match(line)
                    if m:
                        if cur:
                            plan_flights.append(cur)
                        cur = {'flight': m.group(1), 'dep_time': m.group(2), 'arr_time': m.group(3), 'crew': []}
                        expect_route = True
                        expect_crew = False
                    elif cur and expect_route and '-' in line:
                        expect_route = False
                        expect_crew = True
                    elif cur and expect_crew and ',' in line:
                        cur['crew'] = [s.strip() for s in line.split(',') if s.strip()]
                        expect_crew = False
                if cur:
                    plan_flights.append(cur)

                MONTHS2 = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
                result = []
                unmatched = []
                for pf in plan_flights:
                    match = None
                    for fd in flight_db:
                        if fd['flight'] == pf['flight'] and fd['dep_time'] == pf['dep_time']:
                            match = fd
                            break
                    if not match:
                        unmatched.append(f"{pf['flight']} {pf['dep_time']}")
                        continue

                    try:
                        dep_dt = datetime.strptime(match['date'] + ' ' + pf['dep_time'], '%Y-%m-%d %H:%M')
                    except Exception:
                        continue

                    pilots = [c for c in pf['crew'] if re.match(r'^(P|W|PJZ)', c, re.I)]
                    recipients = pilots[:2]
                    emails = [PILOT_MAP.get(r) for r in recipients if PILOT_MAP.get(r)]
                    if not emails:
                        unmatched.append(f"{pf['flight']} {pf['dep_time']} (无邮箱)")
                        continue

                    d = datetime.strptime(match['date'], '%Y-%m-%d')
                    date_text = f"{d.day:02d}{MONTHS2[d.month-1]}"
                    subject = f"WX AND NOTAM {pf['flight']} {match['dep']}-{match['arr']} {date_text}"
                    mailto = "mailto:" + ";".join(emails) + "?subject=" + quote(subject)

                    mail_id = f"{pf['flight']}_{match['date']}_{pf['dep_time']}"
                    result.append({
                        'mail_id': mail_id,
                        'dep_dt': dep_dt,
                        'mailto': mailto,
                        'text': f"{pf['flight']} {match['dep']}-{match['arr']} {pf['dep_time']}-{pf['arr_time']} → {', '.join(recipients)}",
                    })

                st.session_state.f2_flights = result
                if unmatched:
                    st.warning("以下航班未匹配到航段表或邮箱：\n" + "\n".join(unmatched))
                if not result:
                    st.warning("未生成任何邮件，请检查航班号和起飞时间是否与航段表一致。")

    if st.session_state.f2_flights:
        now = now_bj()
        shown = 0
        to_delete = []
        for f in st.session_state.f2_flights:
            dep_dt = f['dep_dt']
            if now >= dep_dt:
                continue
            shown += 1
            three_h = dep_dt - timedelta(hours=3)
            if now >= three_h:
                color = '#fff9c4'
            else:
                color = '#f0f0f0'

            col1, col2 = st.columns([12, 1])
            with col1:
                st.markdown(
                    f'<a href="{f["mailto"]}" target="_blank" style="display:block;padding:10px;'
                    f'background:{color};border:1px solid #ccc;border-radius:4px;'
                    f'text-decoration:none;color:#0066cc;">{f["text"]}</a>',
                    unsafe_allow_html=True
                )
            with col2:
                if st.button("✕", key=f"del_{f['mail_id']}", help="删除此条"):
                    to_delete.append(f['mail_id'])

        if to_delete:
            st.session_state.f2_flights = [
                x for x in st.session_state.f2_flights if x['mail_id'] not in to_delete
            ]
            st.rerun()

        if shown == 0:
            st.info("所有邮件都已过期或已手动删除。")
    else:
        st.info("请上传航段表并粘贴文本飞行计划，然后点击生成。")
