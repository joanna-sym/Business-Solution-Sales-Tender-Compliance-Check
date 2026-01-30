import streamlit as st
import pandas as pd
from docx import Document
import os
import time
import base64
import io

# ================= 🎨 0. 视觉系统配置 (Design System 1.0) =================

st.set_page_config(
    page_title="阿九Joanna · 标书合规排雷引擎",
    page_icon="❤️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 核心配色
COLOR_PRIMARY = "#8D5B5B"   # Marsala
COLOR_ACCENT  = "#C77D7D"   # Velvet
COLOR_BG      = "#F7F5F2"   # Warm Beige
COLOR_TEXT    = "#4A4A4A"   # Coffee

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_watermark_css(text):
    svg = f"""
    <svg xmlns="http://www.w3.org/2000/svg" width="300" height="180">
      <style>
        .text {{ fill: #8D7D70; opacity: 0.08; font-family: 'Microsoft YaHei', sans-serif; font-weight: 800; font-size: 20px; }}
      </style>
      <text x="50%" y="50%" text-anchor="middle" dominant-baseline="middle" transform="rotate(-30 150 90)" class="text">{text}</text>
    </svg>
    """
    b64 = base64.b64encode(svg.encode('utf-8')).decode('utf-8')
    return f"url('data:image/svg+xml;base64,{b64}')"

def inject_custom_css():
    watermark_url = get_watermark_css("阿九Joanna (WeChat: a9joanna01)")
    st.markdown(f"""
    <style>
        /* 全局样式 */
        html, body, [class*="css"] {{ font-family: "Microsoft YaHei", "Helvetica Neue", sans-serif !important; color: {COLOR_TEXT}; }}
        .stApp {{ background-color: {COLOR_BG}; background-image: {watermark_url}; background-attachment: fixed; }}
        
        /* 侧边栏样式 */
        [data-testid="stSidebar"] {{ background-color: rgba(255, 255, 255, 0.95); border-right: 1px solid #EBE6E0; }}
        
        /* 🔥 头像核心样式：强制圆形 + 边框 + 阴影 */
        [data-testid="stSidebar"] img {{ 
            border-radius: 50%; 
            border: 3px solid {COLOR_ACCENT}; 
            box-shadow: 0 4px 10px rgba(0,0,0,0.15); 
            display: block;
        }}
        
        /* 标题与按钮 */
        h1, h2, h3 {{ color: {COLOR_PRIMARY} !important; font-family: "Optima", sans-serif !important; }}
        div.stButton > button {{ 
            background-color: {COLOR_ACCENT}; color: white; border-radius: 6px; border: none; padding: 0.6rem 1.5rem; font-weight: bold; transition: all 0.3s;
        }}
        div.stButton > button:hover {{ 
            background-color: #A65D57; transform: translateY(-2px); box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .phase-card {{ 
            background-color: #FFFFFF; padding: 25px; border-radius: 12px; margin-bottom: 20px; border-left: 5px solid {COLOR_ACCENT}; box-shadow: 0 2px 10px rgba(0,0,0,0.03); 
        }}
        [data-testid="stMetricValue"] {{ color: {COLOR_PRIMARY} !important; }}
    </style>
    """, unsafe_allow_html=True)

# ================= 🛠️ 1. 核心逻辑函数 =================

def extract_tender_table(uploaded_file):
    try:
        doc = Document(uploaded_file)
        data = []
        found_table = False
        for table in doc.tables:
            if len(table.rows) > 0 and len(table.rows[0].cells) >= 2:
                header_text = " ".join([cell.text for cell in table.rows[0].cells])
                if "Parameter" in header_text or "指标" in header_text or "名称" in header_text:
                    found_table = True
                    for i, row in enumerate(table.rows):
                        if i == 0: continue
                        if len(row.cells) >= 3:
                            p_name = row.cells[1].text.strip()
                            p_req = row.cells[2].text.strip()
                        elif len(row.cells) == 2:
                            p_name = row.cells[0].text.strip()
                            p_req = row.cells[1].text.strip()
                        else:
                            continue
                        if p_name:
                            data.append({"Parameter Name": p_name, "Tender Requirement": p_req})
                    break
        if not found_table:
            st.warning("⚠️ 警告：未在 Word 中识别到包含【指标/Parameter】的表格。")
            return pd.DataFrame(columns=["Parameter Name", "Tender Requirement"])
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"❌ 解析 Word 失败: {e}")
        return pd.DataFrame(columns=["Parameter Name", "Tender Requirement"])

# ================= 🎬 2. 主界面 UI =================

def main():
    inject_custom_css()

    # --- 头像逻辑 ---
    assets_dir = os.path.join(CURRENT_DIR, "assets")
    profile_path = None
    possible_files = ["profile.png", "profile.jpg"]
    for f in possible_files:
        p = os.path.join(assets_dir, f)
        if os.path.exists(p):
            profile_path = p
            break
            
    with st.sidebar:
        if profile_path:
            # 🔥 布局魔法：使用三列布局来实现居中 [0.8, 2, 0.8]
            c0, c1, c2 = st.columns([0.8, 2, 0.8])
            with c1:
                # use_container_width=True 让图片撑满中间列，配合 CSS 实现圆形
                st.image(profile_path, use_container_width=True)
        else:
            st.markdown('<div style="font-size: 60px; text-align: center;">👩‍💻</div>', unsafe_allow_html=True)

        st.markdown("""
        <div style="text-align: center; margin-bottom: 20px;">
            <h3 style="margin-top: 15px;">阿九Joanna</h3>
            <p style="font-size: 12px; opacity: 0.8;">MedTech Digital Architect</p>
        </div>
        """, unsafe_allow_html=True)
        st.info("💡 **操作指南**:\n1. 上传标书 (.docx)\n2. 上传产品参数 (.xlsx)\n3. 启动排雷")

    # --- 主内容 ---
    st.title("❤️ 阿九Joanna · 标书合规排雷引擎")
    st.markdown("Automated Medical Tender Compliance Audit System")

    st.markdown('<div class="phase-card">', unsafe_allow_html=True)
    st.markdown("### 📂 第一阶段：数据装载 (Data Loading)")
    c1, c2 = st.columns(2)
    with c1:
        uploaded_tender = st.file_uploader("📄 上传招标文件 (Word)", type=['docx'])
    with c2:
        uploaded_specs = st.file_uploader("📊 上传我司参数库 (Excel)", type=['xlsx'])
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_tender and uploaded_specs:
        st.markdown('<div class="phase-card">', unsafe_allow_html=True)
        st.markdown("### ⚙️ 第二阶段：智能审计 (AI Auditing)")
        
        if st.button("🚀 启动排雷程序 (START AUDIT)", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            steps = ["解析文档...", "提取表格...", "读取数据库...", "比对计算...", "生成报表..."]
            for i, step in enumerate(steps):
                status_text.text(f"⚡ {step}")
                time.sleep(0.3)
                progress_bar.progress(int((i + 1) / len(steps) * 100))
            
            df_tender = extract_tender_table(uploaded_tender)
            df_product = pd.read_excel(uploaded_specs)
            
            if df_tender.empty:
                st.error("❌ 无法提取有效数据！")
                st.stop()
            
            df_product.columns = df_product.columns.str.strip()
            param_col_excel = None
            for col in df_product.columns:
                if "Parameter" in col or "指标" in col or "名称" in col:
                    param_col_excel = col
                    break
            
            if not param_col_excel:
                 st.error("❌ Excel 中没找到【Parameter Name】列。")
                 st.stop()

            df_product = df_product.rename(columns={param_col_excel: "Parameter Name"})
            result = pd.merge(df_tender, df_product, on="Parameter Name", how="left")
            
            def check_compliance(row):
                req = str(row.get('Tender Requirement', '')).strip()
                cols = row.index.tolist()
                spec_col = [c for c in cols if c not in ['Parameter Name', 'Tender Requirement', 'Audit Result']]
                my_spec = str(row.get(spec_col[0], '')).strip() if spec_col else "未找到"

                if "Mandatory" in req and "trap" not in my_spec.lower(): 
                     return "✅ 合格 (Pass)"
                if req == my_spec:
                    return "✅ 合格 (Pass)"
                
                if "≥" in req or ">=" in req:
                    try:
                        req_val = float(''.join(filter(str.isdigit, req)))
                        my_val = float(''.join(filter(str.isdigit, my_spec)))
                        return "✅ 合格 (Pass)" if my_val >= req_val else "❌ 不合格 (FAIL)"
                    except:
                        pass
                
                return "⚠️ 需人工复核 (Review)"

            try:
                result['Audit Result'] = result.apply(check_compliance, axis=1)
                
                st.divider()
                
                total_items = len(result)
                failed_items = len(result[result['Audit Result'].str.contains("❌")])
                review_items = len(result[result['Audit Result'].str.contains("⚠️")])
                
                m1, m2, m3 = st.columns(3)
                m1.metric("总核查项", f"{total_items} 项")
                m2.metric("❌ 风险/不合格", f"{failed_items} 项", delta_color="inverse")
                m3.metric("⚠️ 需复核", f"{review_items} 项")
                
                def highlight_fail(row):
                    res = row['Audit Result']
                    if "❌" in res: return ['background-color: #FFEBEB; color: #C00000'] * len(row)
                    elif "✅" in res: return ['background-color: #E6F4EA; color: #1E7E34'] * len(row)
                    return [''] * len(row)

                st.dataframe(result.style.apply(highlight_fail, axis=1), use_container_width=True, height=400)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    result.to_excel(writer, index=False, sheet_name='Audit_Report')
                
                st.download_button(
                    label="📥 下载审计报告 (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="Final_Audit_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                progress_bar.progress(100)
                status_text.text("✅ 审计完成！")

            except Exception as e:
                st.error(f"逻辑出错: {str(e)}")

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()