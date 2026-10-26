import os
import io
import streamlit as st
import pandas as pd
import openpyxl

import pdf_processor
import excel_generator
import boq_engine
import schedule_engine

st.set_page_config(
    page_title="SPACEMAN STUDIO - AI BOQ & Schedule Pro",
    page_icon="🪐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for clean engineering aesthetic
st.markdown("""
<style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        font-weight: 600;
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- SECURITY: PASSWORD GATE -----------------
SYSTEM_PASSWORD = None
if hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
    SYSTEM_PASSWORD = st.secrets["APP_PASSWORD"]
else:
    SYSTEM_PASSWORD = os.getenv("APP_PASSWORD", "boq2026")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    pwd = st.session_state.get("password_input", "")
    if pwd == SYSTEM_PASSWORD:
        st.session_state.authenticated = True
    else:
        st.session_state.authenticated = False
        st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาติดต่อผู้ดูแลระบบเพื่อขอรหัสผ่าน")

if not st.session_state.authenticated:
    st.markdown('<div class="main-title">🔒 ระบบเข้าสู่ระบบ SPACEMAN STUDIO</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">เครื่องมือถอด BOQ และสร้าง Project Schedule (เฉพาะทีมงานและผู้ได้รับอนุญาต)</div>', unsafe_allow_html=True)
    
    st.info("💡 **ระบบความปลอดภัย:** เพื่อนๆ สามารถใส่รหัสผ่านที่ได้รับเพื่อเริ่มใช้งานได้ทันที (โดยไม่ต้องใช้ API Key ส่วนตัว)")
    col_p1, col_p2 = st.columns([1, 1.5])
    with col_p1:
        st.text_input("กรอกรหัสผ่านเพื่อเข้าใช้งาน:", type="password", key="password_input", on_change=check_password)
        st.button("🔑 เข้าสู่ระบบ", on_click=check_password, type="primary")
        st.caption("รหัสผ่านเริ่มต้นสำหรับทดสอบ: `boq2026` (สามารถเปลี่ยนได้)")
    st.stop()

# ----------------- SESSION STATE INITIALIZATION -----------------
if "items_data" not in st.session_state:
    st.session_state.items_data = []
if "pdf_doc" not in st.session_state:
    st.session_state.pdf_doc = None
if "pdf_info" not in st.session_state:
    st.session_state.pdf_info = None
if "detected_codes" not in st.session_state:
    st.session_state.detected_codes = []
if "schedule_data" not in st.session_state:
    st.session_state.schedule_data = schedule_engine.get_sample_brewave_schedule()

# ----------------- SECRETS & API KEY RESOLUTION -----------------
server_api_key = ""
if hasattr(st, "secrets") and "OPENCODE_API_KEY" in st.secrets:
    server_api_key = st.secrets["OPENCODE_API_KEY"]
elif os.getenv("OPENCODE_API_KEY"):
    server_api_key = os.getenv("OPENCODE_API_KEY")

# ----------------- SIDEBAR: CONFIGURATION -----------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/blueprint.png", width=64)
    st.title("⚙️ ตั้งค่าระบบ AI & ฟอร์ม")
    
    st.subheader("1. OpenCode Go (AI Engine)")
    
    if server_api_key:
        st.success("🔒 ระบบใช้ API Key กลางของเจ้าของระบบ (เพื่อนๆ กดใช้ได้เลย ไม่ต้องกรอก)")
        api_key_input = server_api_key
    else:
        api_key_input = st.text_input(
            "OpenCode Go API Key",
            value="",
            type="password",
            help="ใส่ API Key จาก OpenCode Go เพื่อให้ AI ถอดแบบแปลน"
        )
    
    base_url_input = st.text_input(
        "Base URL (OpenAI-compatible)",
        value="https://opencode.ai/zen/go/v1",
        help="Endpoint มาตรฐานของ OpenCode Go"
    )
    
    model_choice = st.selectbox(
        "เลือก AI Model",
        options=["deepseek-v4.1-flash", "deepseek-chat", "deepseek-reasoner", "qwen-2.5-72b-instruct", "glm-4-plus", "kimi-latest", "ระบุเอง..."],
        index=0
    )
    if model_choice == "ระบุเอง...":
        model_choice = st.text_input("ชื่อ Model ที่ต้องการ", value="deepseek-v4.1-flash")
        
    st.divider()
    st.subheader("2. แบบฟอร์ม BOQ แม่แบบ")
    template_option = st.radio(
        "เลือก Template",
        options=["ใช้ Blank BOQ_PLC HOUSE .xlsx (ในโฟลเดอร์)", "อัปโหลดไฟล์ Excel (.xlsx) ของฉัน"],
        index=0
    )
    
    custom_template_file = None
    if template_option == "อัปโหลดไฟล์ Excel (.xlsx) ของฉัน":
        custom_template_file = st.file_uploader("เลือกไฟล์ Template Excel", type=["xlsx"])
    
    st.divider()
    st.subheader("3. การคิดค่าดำเนินการ (Overhead)")
    overhead_pct = st.slider("ค่าดำเนินการ (%)", min_value=0.0, max_value=25.0, value=10.0, step=0.5)
    
    st.divider()
    if st.button("🚪 ออกจากระบบ"):
        st.session_state.authenticated = False
        st.rerun()

# ----------------- MAIN CONTENT -----------------
st.markdown('<div class="main-title">🪐 SPACEMAN STUDIO — AI BOQ & Project Schedule</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">เครื่องมืออัจฉริยะสำหรับถอดแบบ BOQ และสร้างไทม์ไลน์โครงการ (Gantt Chart) ส่งมอบลูกค้า</div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "📑 1. แบบก่อสร้าง (Drawings & PDF)", 
    "🤖 2. ประมวลผลถอด BOQ (AI Extract)", 
    "📊 3. ตรวจสอบ & ส่งออก BOQ Excel",
    "📅 4. ไทม์ไลน์งาน (Project Schedule)"
])

# ================= TAB 1: DRAWINGS =================
with tab1:
    col_upload, col_sample = st.columns([2, 1])
    with col_upload:
        uploaded_pdf = st.file_uploader("อัปโหลดไฟล์แบบก่อสร้าง (PDF หรือ DXF)", type=["pdf", "dxf"])
    with col_sample:
        st.write("**หรือทดลองด้วยไฟล์ตัวอย่าง:**")
        if st.button("⚡ โหลด PLC HOUSE_Fortender.pdf (134 หน้า)", use_container_width=True):
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            sample_path = os.path.join(BASE_DIR, "PLC HOUSE_Fortender.pdf")
            if os.path.exists(sample_path):
                st.session_state.pdf_doc = pdf_processor.get_pdf_document(sample_path)
                st.session_state.pdf_info = pdf_processor.get_pdf_summary(st.session_state.pdf_doc)
                st.session_state.detected_codes = pdf_processor.scan_drawing_schedule_and_codes(st.session_state.pdf_doc)
                st.success("✅ โหลดแบบแปลน PLC House เรียบร้อยแล้ว!")
            else:
                st.error("ไม่พบไฟล์ตัวอย่างในไดเรกทอรี")

    if uploaded_pdf is not None:
        st.session_state.pdf_doc = pdf_processor.get_pdf_document(uploaded_pdf)
        st.session_state.pdf_info = pdf_processor.get_pdf_summary(st.session_state.pdf_doc)
        st.session_state.detected_codes = pdf_processor.scan_drawing_schedule_and_codes(st.session_state.pdf_doc)
        st.success(f"✅ โหลดไฟล์ {uploaded_pdf.name} สำเร็จ!")

    if st.session_state.pdf_doc is not None:
        doc = st.session_state.pdf_doc
        info = st.session_state.pdf_info
        codes = st.session_state.detected_codes

        # Metrics bar
        m1, m2, m3 = st.columns(3)
        m1.metric("จำนวนหน้าแบบทั้งหมด", f"{info['total_pages']} หน้า")
        m2.metric("ตรวจพบรหัสเฟอร์นิเจอร์/แบบ", f"{len(codes)} รายการ")
        m3.metric("สถานะไฟล์", "พร้อมสำหรับการประมวลผล")

        st.divider()
        col_preview, col_codes = st.columns([1.2, 1])
        
        with col_preview:
            st.subheader("🔍 พรีวิวแบบแปลน (Drawing Viewer)")
            page_to_view = st.slider("เลือกหน้าแบบแปลนที่ต้องการดู", 1, info["total_pages"], 8)
            with st.spinner("กำลังเรนเดอร์ภาพแบบแปลน..."):
                img = pdf_processor.render_page_image(doc, page_to_view, dpi=120)
                st.image(img, caption=f"หน้า {page_to_view} จาก {info['total_pages']}", use_container_width=True)

        with col_codes:
            st.subheader("📋 รายการรหัสแบบที่ AI ตรวจพบ (Index & Schedule)")
            if codes:
                codes_df = pd.DataFrame(codes)[["code", "type", "page", "context"]]
                codes_df.columns = ["รหัสแบบ", "ประเภท", "หน้าแบบ", "ข้อความประกอบ"]
                st.dataframe(codes_df, use_container_width=True, height=450)
            else:
                st.info("ไม่พบรหัสมาตรฐานอัตโนมัติ คุณสามารถเลือกหน้าเพื่อส่งให้ AI วิเคราะห์ในแท็บที่ 2 ได้")

# ================= TAB 2: AI EXTRACTION =================
with tab2:
    st.subheader("🤖 ให้ AI วิเคราะห์และถอดรายการ BOQ")
    
    col_input, col_ctrl = st.columns([2, 1])
    with col_input:
        extract_mode = st.radio(
            "โหมดการดึงข้อมูล:",
            ["ดึงจากรหัสและสารบัญแบบที่ AI ตรวจพบ (แนะนำ)", "เลือกหน้าที่ต้องการให้ AI อ่านสเปกงาน", "ป้อนข้อความสเปกงานด้วยตนเอง"],
            horizontal=True
        )
        
        custom_spec_text = ""
        selected_pages = [8]
        if extract_mode == "เลือกหน้าที่ต้องการให้ AI อ่านสเปกงาน":
            if st.session_state.pdf_doc is not None:
                selected_pages = st.multiselect(
                    "เลือกหน้าแบบที่ต้องการให้ AI ถอดรายการ:",
                    options=list(range(1, st.session_state.pdf_info["total_pages"] + 1)),
                    default=[8, 16, 20]
                )
        elif extract_mode == "ป้อนข้อความสเปกงานด้วยตนเอง":
            custom_spec_text = st.text_area(
                "วางสเปกงานหรือรายละเอียดวัสดุที่ต้องการให้ AI ถอดเป็น BOQ:",
                value="ตู้ TV ชั้น 1 ไม้ HMR ปิดลามิเนตลายไม้ ขนาด 1.14*2.62*2.70 ม. จำนวน 1 ชุด\nตู้เสื้อผ้าห้องนอนใหญ่ ขนาด 1.30*1.30*2.80 ม. จำนวน 1 ชุด\nงานปูกระเบื้องพื้นห้องน้ำ 6 ตร.ม.",
                height=150
            )

    with col_ctrl:
        st.write("**เริ่มการประมวลผล:**")
        run_ai_btn = st.button("🚀 สั่ง AI ถอด BOQ (OpenCode Go)", type="primary", use_container_width=True)
        demo_btn = st.button("⚡ โหลดข้อมูลทดสอบสำเร็จรูป (PLC House Demo)", use_container_width=True, help="ใช้ทดลองระบบทันทีโดยไม่ต้องใส่ API Key")

    if demo_btn:
        st.session_state.items_data = boq_engine.get_demo_precalculated_items()
        st.success("✅ โหลดข้อมูลทดสอบสำหรับโครงการ PLC House สำเร็จ! (มีทั้งชั้น 1, ชั้น 2, งานระบบ และ Preliminary)")

    if run_ai_btn:
        if not api_key_input:
            st.error("⚠️ ยังไม่ได้ตั้งค่า 'OpenCode Go API Key' (หรือคลิกปุ่ม 'โหลดข้อมูลทดสอบสำเร็จรูป' เพื่อลองเล่นก่อนได้)")
        else:
            with st.spinner(f"กำลังส่งข้อมูลให้ {model_choice} ผ่าน OpenCode Go วิเคราะห์..."):
                try:
                    if extract_mode == "ดึงจากรหัสและสารบัญแบบที่ AI ตรวจพบ (แนะนำ)":
                        if st.session_state.pdf_doc is not None:
                            content_to_send = pdf_processor.extract_pages_text(st.session_state.pdf_doc, [8, 16, 20, 24, 42])
                        else:
                            content_to_send = "รายการรหัสแบบ Built-in โครงการ PLC House"
                    elif extract_mode == "เลือกหน้าที่ต้องการให้ AI อ่านสเปกงาน":
                        content_to_send = pdf_processor.extract_pages_text(st.session_state.pdf_doc, selected_pages)
                    else:
                        content_to_send = custom_spec_text

                    extracted = boq_engine.extract_boq_from_text_or_drawing(
                        api_key=api_key_input,
                        content_text=content_to_send,
                        base_url=base_url_input,
                        model_name=model_choice
                    )
                    st.session_state.items_data = extracted
                    st.success(f"🎉 AI ถอดรายการสำเร็จ! ได้ทั้งหมด {len(extracted)} รายการ")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการเรียก AI: {str(e)}")

    st.divider()
    st.subheader("📝 ตารางตรวจสอบ & ปรับแต่งตัวเลข (Interactive BOQ Editor)")
    st.caption("💡 เพื่อนๆ สามารถคลิกแก้ไขปริมาณ (Qty), หน่วย (Unit), ค่าวัสดุ, ค่าแรง ได้โดยตรงในตารางนี้ ผลรวมจะคำนวณใหม่อัตโนมัติ!")

    if st.session_state.items_data:
        df = pd.DataFrame(st.session_state.items_data)
        
        for col in ["code", "category", "description", "qty", "unit", "material_rate", "labour_rate", "remark"]:
            if col not in df.columns:
                df[col] = "" if col in ["code", "category", "description", "unit", "remark"] else 0.0

        df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(1.0)
        df["material_rate"] = pd.to_numeric(df["material_rate"], errors="coerce").fillna(0.0)
        df["labour_rate"] = pd.to_numeric(df["labour_rate"], errors="coerce").fillna(0.0)
        df["total_amount"] = (df["material_rate"] + df["labour_rate"]) * df["qty"]

        display_df = df[["code", "category", "description", "qty", "unit", "material_rate", "labour_rate", "total_amount", "remark"]].copy()
        display_df.columns = ["รหัสแบบ", "หมวดงาน", "รายละเอียดรายการ", "จำนวน", "หน่วย", "ค่าวัสดุ/หน่วย", "ค่าแรง/หน่วย", "รวมเงิน (บาท)", "หมายเหตุ"]

        edited_df = st.data_editor(
            display_df,
            use_container_width=True,
            num_rows="dynamic",
            height=400,
            column_config={
                "จำนวน": st.column_config.NumberColumn(format="%.2f"),
                "ค่าวัสดุ/หน่วย": st.column_config.NumberColumn(format="฿%d"),
                "ค่าแรง/หน่วย": st.column_config.NumberColumn(format="฿%d"),
                "รวมเงิน (บาท)": st.column_config.NumberColumn(format="฿%d", disabled=True),
            }
        )

        st.session_state.items_data = [
            {
                "code": row["รหัสแบบ"],
                "category": row["หมวดงาน"],
                "description": row["รายละเอียดรายการ"],
                "qty": float(row["จำนวน"]),
                "unit": row["หน่วย"],
                "material_rate": float(row["ค่าวัสดุ/หน่วย"]),
                "labour_rate": float(row["ค่าแรง/หน่วย"]),
                "remark": row["หมายเหตุ"]
            }
            for _, row in edited_df.iterrows()
        ]
    else:
        st.info("ยังไม่มีข้อมูลรายการ BOQ กรุณากดปุ่มด้านบนเพื่อให้ AI เริ่มทำงาน หรือคลิก 'โหลดข้อมูลทดสอบสำเร็จรูป'")

# ================= TAB 3: REVIEW & EXPORT =================
with tab3:
    st.subheader("📊 สรุปงบประมาณและส่งออกไฟล์ Excel")
    
    if st.session_state.items_data:
        items = st.session_state.items_data
        total_mat = sum(float(it.get("material_rate", 0)) * float(it.get("qty", 1)) for it in items)
        total_lab = sum(float(it.get("labour_rate", 0)) * float(it.get("qty", 1)) for it in items)
        subtotal = total_mat + total_lab
        overhead_amount = subtotal * (overhead_pct / 100.0)
        grand_total = subtotal + overhead_amount

        col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
        col_kpi1.metric("รวมค่าวัสดุ", f"{total_mat:,.2f} ฿")
        col_kpi2.metric("รวมค่าแรง", f"{total_lab:,.2f} ฿")
        col_kpi3.metric(f"ค่าดำเนินการ ({overhead_pct}%)", f"{overhead_amount:,.2f} ฿")
        col_kpi4.metric("ราคารวมทั้งสิ้น (Grand Total)", f"{grand_total:,.2f} ฿")

        st.divider()
        col_preview_summary, col_download = st.columns([1.5, 1])
        
        with col_preview_summary:
            st.write("### 📑 ตัวอย่างโครงสร้างใบเสนอราคา (Quotation Summary)")
            summary_preview = pd.DataFrame([
                {"รายการ": "1. Furniture Built-in ชั้น 1", "จำนวน": "1 Lot", "สถานะ": "คำนวณจากแบบแปลน"},
                {"รายการ": "2. Furniture Built-in ชั้น 2", "จำนวน": "1 Lot", "สถานะ": "คำนวณจากแบบแปลน"},
                {"รายการ": "3. งานระบบไฟฟ้า / ดวงโคม / สวิตซ์", "จำนวน": "1 Lot", "สถานะ": "คำนวณจากสเปก"},
                {"รายการ": "4. งานเตรียมการ (Preliminary)", "จำนวน": "1 Lot", "สถานะ": "โพรเทคชั่น + ทำความสะอาด"},
                {"รายการ": "รวมค่างาน (Total Direct Cost)", "จำนวน": "-", "สถานะ": f"{subtotal:,.2f} ฿"},
                {"รายการ": f"ค่าดำเนินการและกำไร ({overhead_pct}%)", "จำนวน": "-", "สถานะ": f"{overhead_amount:,.2f} ฿"},
                {"รายการ": "ราคารวมทั้งสิ้น (Sum Total)", "จำนวน": "-", "สถานะ": f"{grand_total:,.2f} ฿"},
            ])
            st.table(summary_preview)

        with col_download:
            st.write("### 💾 ส่งออกไฟล์ Excel (.xlsx)")
            st.write("ระบบจะนำข้อมูลที่ AI ถอดและคุณตรวจสอบแล้ว หยอดลงในแม่แบบ **Blank BOQ_PLC HOUSE .xlsx** พร้อม **ฝังสูตร Excel จริง** สำหรับส่งมอบลูกค้าหรือผู้รับเหมาทันที")
            
            project_title = st.text_input("ชื่อโครงการ (Project)", value="Reignwood Park (PLC House)")
            customer_title = st.text_input("ชื่อลูกค้า (Customer)", value="คุณลูกค้า")
            
            excel_bytes = excel_generator.populate_boq_excel(
                template_source=custom_template_file,
                items_data=st.session_state.items_data,
                overhead_pct=overhead_pct,
                project_name=project_title,
                customer_name=customer_title
            )

            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ BOQ Excel (.xlsx) ที่สมบูรณ์",
                data=excel_bytes.getvalue(),
                file_name=f"BOQ_{project_title.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            st.success("✨ ไฟล์ Excel มีการผูกสูตรผลรวมและค่าดำเนินการตามโครงสร้างมาตรฐานของบริษัทเรียบร้อยแล้ว!")
    else:
        st.info("กรุณาประมวลผลหรือโหลดข้อมูลในแท็บที่ 2 ก่อนเพื่อดูผลสรุปและดาวน์โหลดไฟล์ Excel")

# ================= TAB 4: PROJECT SCHEDULE =================
with tab4:
    st.subheader("📅 SPACEMAN STUDIO — Project Schedule Generator")
    st.markdown("พิมพ์บอกระยะเวลาและขั้นตอนงานเป็นภาษาพูด แล้วให้ AI จัดตาราง Gantt Chart สีสันตามมาตรฐานพร้อมจุดประชุม (M) และ Opening Date")

    col_sched_input, col_sched_btn = st.columns([2.5, 1])
    
    with col_sched_input:
        default_prompt = (
            "โปรเจกต์ Brewave Aree ออกแบบและตกแต่งภายใน เริ่มวันที่ 09 May 2025\n"
            "วางแผน 7 ขั้นตอน:\n"
            "A. Schematic Design: ทำ Lay-out 1 สัปดาห์ (May W2) นัดประชุม M\n"
            "B. Preliminary design: Rendering 1 สัปดาห์ นัดประชุม M, แก้แบบ 1 สัปดาห์, Final rendering 1 สัปดาห์ นัดประชุม M, Material presentation 2 สัปดาห์ นัดประชุม M\n"
            "C. Design develop: MEP Drawing 2 สัปดาห์, AC Drawing 2 สัปดาห์\n"
            "D. Final drawing 1: Interior joinery drawing 4 สัปดาห์ (June W1-W4)\n"
            "E. Final drawing 2: Interior detail & equipment specs 4 สัปดาห์ (June W4 - July W3)\n"
            "F. BOQ stage: MEP/AC BOQ 1 สัปดาห์ นัดประชุม, Interior construction BOQ 3 สัปดาห์ นัดประชุม, สั่งเฟอร์นิเจอร์ลอยตัว (Loose furniture) 10 สัปดาห์ นัดประชุม 3 ครั้ง\n"
            "G. Construction stage: งานระบบ 6 สัปดาห์ นัดประชุม 2 ครั้ง, งานก่อสร้างภายใน 10 สัปดาห์ นัดประชุมทุก 2 สัปดาห์, งานเก็บดีเฟกต์ 4 สัปดาห์ นัดประชุม 2 ครั้ง\n"
            "กำหนดเปิดร้าน (OPENING DATE): สัปดาห์ที่ 2 ของเดือน October 2025"
        )
        schedule_prompt_text = st.text_area(
            "พิมพ์รายละเอียดไทม์ไลน์ที่ต้องการให้ AI วางแผน:",
            value=default_prompt,
            height=200
        )

    with col_sched_btn:
        st.write("**การสร้างไทม์ไลน์:**")
        ai_sched_btn = st.button("🚀 สั่ง AI สร้างไทม์ไลน์ (OpenCode Go)", type="primary", use_container_width=True)
        sample_sched_btn = st.button("⚡ โหลดตัวอย่าง Brewave Aree ทันที", use_container_width=True, help="โหลดไทม์ไลน์ตามรูปภาพตัวอย่างของ SPACEMAN STUDIO ทันที")

    if sample_sched_btn:
        st.session_state.schedule_data = schedule_engine.get_sample_brewave_schedule()
        st.success("✅ โหลดไทม์ไลน์ตัวอย่าง Brewave Aree เรียบร้อยแล้ว!")

    if ai_sched_btn:
        if not api_key_input:
            st.error("⚠️ ยังไม่ได้ตั้งค่า 'OpenCode Go API Key' (หรือคลิกปุ่ม 'โหลดตัวอย่าง Brewave Aree ทันที' เพื่อลองดูผลลัพธ์ก่อนได้)")
        else:
            with st.spinner(f"กำลังส่งให้ {model_choice} วางแผนไทม์ไลน์ตามมาตรฐาน SPACEMAN STUDIO..."):
                try:
                    parsed_sched = schedule_engine.parse_schedule_prompt(
                        api_key=api_key_input,
                        user_prompt=schedule_prompt_text,
                        base_url=base_url_input,
                        model_name=model_choice
                    )
                    st.session_state.schedule_data = parsed_sched
                    st.success("🎉 AI สร้างไทม์ไลน์งานสำเร็จเรียบร้อย!")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการสร้างไทม์ไลน์: {str(e)}")

    st.divider()
    
    if st.session_state.schedule_data:
        sched_data = st.session_state.schedule_data

        # Top summary info
        c_pname, c_date, c_year, c_open = st.columns(4)
        c_pname.metric("ชื่อโครงการ", sched_data.get("project_name", "-"))
        c_date.metric("วันที่ออกเอกสาร", sched_data.get("date", "-"))
        c_year.metric("ปีที่ดำเนินการ", str(sched_data.get("year", "2025")))
        op_info = sched_data.get("opening_date", {})
        c_open.metric("กำหนดเปิดร้าน (Opening)", f"{op_info.get('month', '')} Week {op_info.get('week', '')}")

        st.subheader("🖥️ ตัวอย่างตาราง Project Schedule (SPACEMAN STUDIO)")
        
        # Render visual HTML Gantt Table
        gantt_html = schedule_engine.render_schedule_html(sched_data)
        st.components.v1.html(gantt_html, height=580, scrolling=True)

        st.divider()
        col_d1, col_d2 = st.columns([2, 1])
        with col_d1:
            st.write("### 💾 ส่งออกไฟล์ Excel (.xlsx)")
            st.write("ตารางจะถูกระบายสีตามแต่ละหมวดหมู่ (เหลือง, เขียว, ฟ้า, ส้ม, ชมพู) พร้อมใส่ตัว **`M`** ในสัปดาห์ที่มีการนัดประชุมลูกค้า และทำแถบเสาสีเข้ม **`OPENING DATE`** เป๊ะตามมาตรฐานบริษัท")
            
        with col_d2:
            sched_excel_bytes = schedule_engine.generate_schedule_excel(sched_data)
            p_clean = sched_data.get("project_name", "Schedule").replace(" ", "_")
            st.download_button(
                label="📥 ดาวน์โหลดไทม์ไลน์ Excel (.xlsx)",
                data=sched_excel_bytes.getvalue(),
                file_name=f"Schedule_{p_clean}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            st.success("✨ ฟอร์ม Excel ได้รับการจัดระยะหน้ากระดาษและฟอนต์พร้อมส่งต่อลูกค้าทันที!")
