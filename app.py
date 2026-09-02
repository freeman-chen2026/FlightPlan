import streamlit as st
import pandas as pd
import json

st.set_page_config(page_title="Arinc 飞行计划辅助脚本生成器", layout="centered")
st.title("✈️ Arinc 飞行计划辅助脚本生成器")
st.caption("上传 Excel 航段表 → 生成 JS 脚本 → 在 Arinc 页面控制台运行，逐条填充")

# 安全提示
st.info(
    "⚠️ **合规说明**：本工具生成的脚本仅用于**辅助填写**表单字段（飞机号、起飞机场、目的地机场），"
    "**不会自动点击提交**。每次填充后需您手动点击提交按钮，完全符合“浏览器辅助”安全规范。"
)

uploaded_file = st.file_uploader("📂 上传航班计划 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 读取Excel（默认第一个sheet）
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 智能定位列（优先按列名，回退到位置 C=2, K=10, M=12）
        # 根据您附件的表头：飞机注册号、出发地、到达地
        if '飞机注册号' in df.columns and '出发地' in df.columns and '到达地' in df.columns:
            aircraft_col = df['飞机注册号']
            origin_col = df['出发地']
            dest_col = df['到达地']
        else:
            # 如果表头名字不一致，按您描述的位置读取（0-based索引）
            aircraft_col = df.iloc[:, 2]   # C列
            origin_col = df.iloc[:, 10]    # K列
            dest_col = df.iloc[:, 12]      # M列
            st.warning("未识别标准列名，已按位置读取：C列(飞机号)、K列(出发地)、M列(到达地)。")

        # 组装有效航班数据
        flights = []
        for idx in range(len(df)):
            a = aircraft_col.iloc[idx]
            o = origin_col.iloc[idx]
            d = dest_col.iloc[idx]
            # 跳过空值或无效值
            if pd.notna(a) and pd.notna(o) and pd.notna(d):
                flights.append({
                    "aircraft": str(a).strip(),
                    "origin": str(o).strip().upper(),
                    "dest": str(d).strip().upper()
                })

        if not flights:
            st.error("❌ 未读取到有效航段数据，请检查Excel列是否正确（飞机注册号、出发地、到达地）。")
            st.stop()

        st.success(f"✅ 成功解析 **{len(flights)}** 条有效航段")

        # 预览数据（前5条）
        with st.expander("📋 预览解析数据（前5条）"):
            preview_df = pd.DataFrame(flights).head(5)
            st.dataframe(preview_df, use_container_width=True)

        # ---------- 生成 JavaScript 脚本 ----------
        # 将 flights 转为 JSON 字符串
        flights_json = json.dumps(flights, ensure_ascii=False, indent=2)

        js_script = f"""
// ============================================================
//  Arinc 飞行计划自动填表脚本
//  使用方法：
//  1. 在 Arinc 页面按 F12 打开开发者工具，切换到 Console（控制台）面板
//  2. 将下方全部代码粘贴进去，按回车执行
//  3. 之后每填完一条，在控制台输入 fillNext() 并回车，即可自动填入下一条
//  4. 每次填完后，请手动点击页面上的提交按钮
// ============================================================

// 存储所有航段数据（共 {len(flights)} 条）
var flightData = {flights_json};

// 当前填充的索引（从0开始）
var currentIndex = 0;

// 辅助函数：通过 XPath 获取元素
function getElementByXpath(path) {{
    return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
}}

// 核心填充函数：根据索引填充表单
function fillFlight(index) {{
    if (index < 0 || index >= flightData.length) {{
        console.error("❌ 索引超出范围！有效范围 0 ~ " + (flightData.length - 1));
        return false;
    }}
    var data = flightData[index];
    console.log(`🛫 正在填充第 ${{index + 1}}/{flightData.length} 条: ${{data.aircraft}}  ${{data.origin}} -> ${{data.dest}}`);

    // 1. 填充飞机注册号（下拉框）
    var aircraftSelect = document.getElementById('Aircraft');
    if (aircraftSelect) {{
        aircraftSelect.value = data.aircraft;
        aircraftSelect.dispatchEvent(new Event('change', {{ bubbles: true }}));
        console.log("  ✅ 飞机号已选: " + data.aircraft);
    }} else {{
        console.error("  ❌ 未找到 id='Aircraft' 的下拉框，请检查页面是否加载完成");
        return false;
    }}

    // 2. 填充起飞机场（使用您提供的 XPath）
    var originInput = getElementByXpath('/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/input[1]');
    if (originInput) {{
        originInput.value = data.origin;
        originInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        originInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        console.log("  ✅ 起飞机场已填: " + data.origin);
    }} else {{
        console.error("  ❌ 未找到起飞机场输入框（XPath 可能已变化）");
        return false;
    }}

    // 3. 填充目的地机场（使用您提供的 XPath）
    var destInput = getElementByXpath('/html/body/div[4]/form/table/tbody/tr/td[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]/table/tbody/tr[4]/td[1]/input[1]');
    if (destInput) {{
        destInput.value = data.dest;
        destInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        destInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        console.log("  ✅ 目的地机场已填: " + data.dest);
    }} else {{
        console.error("  ❌ 未找到目的地机场输入框（XPath 可能已变化）");
        return false;
    }}

    console.log("🎯 填充完成！请手动点击页面上的提交按钮。");
    return true;
}}

// 便捷函数：填充下一条（自动递增索引）
function fillNext() {{
    if (currentIndex >= flightData.length) {{
        console.log("🎉 所有航段已全部填充完毕！");
        return;
    }}
    var success = fillFlight(currentIndex);
    if (success) {{
        currentIndex++;
    }} else {{
        console.warn("⚠️ 填充失败，请检查控制台错误信息，索引停留在 " + currentIndex);
    }}
}}

// 重置索引（如果中途出错想重来，可执行 resetIndex()）
function resetIndex() {{
    currentIndex = 0;
    console.log("🔄 索引已重置为 0");
}}

// 显示帮助信息
console.log("✅ 脚本加载成功！共有 " + flightData.length + " 条航段。");
console.log("📌 在控制台输入 fillNext() 填充下一条，每次填完请手动提交。");
console.log("📌 如需重新开始，输入 resetIndex() 重置索引。");
"""

        # 显示生成的脚本
        st.subheader("📜 生成的 JavaScript 脚本（复制以下全部内容）")
        st.code(js_script, language="javascript")

        # 额外辅助：一键复制按钮（用 st.button 配合 js 复制，但 Streamlit 后端无法直接操作剪贴板，这里提示用户手动 Ctrl+C）
        st.info("💡 请点击代码框左上角的 **复制图标** 或按 **Ctrl+C** 复制全部代码。")

        # 显示使用步骤
        with st.expander("📖 详细使用步骤"):
            st.markdown("""
            1. **登录 Arinc** 并进入飞行计划制作页面。
            2. 按 **F12** 打开开发者工具，点击 **Console（控制台）** 标签。
            3. 将上面生成的 **JavaScript 脚本全部复制**，粘贴到控制台中，按 **回车** 执行。
            4. 看到 `✅ 脚本加载成功！` 提示后，在控制台输入 **`fillNext()`** 并回车。
            5. 脚本会自动填入 **飞机号、起飞机场、目的地机场**。
            6. **你手动点击页面上的提交按钮**，完成本条计划。
            7. 回到控制台，再次输入 **`fillNext()`**，自动填入下一条。
            8. 重复步骤 6-7，直至所有航段制作完毕。
            """)

    except Exception as e:
        st.error(f"❌ 读取文件出错：{e}")
        st.stop()

else:
    st.info("👆 请先上传 Excel 文件以生成脚本")

# 底部说明
st.divider()
st.caption("⚙️ 部署方式：托管于 GitHub，通过 Streamlit Cloud 运行。仅生成前端辅助脚本，不存储任何数据。")
