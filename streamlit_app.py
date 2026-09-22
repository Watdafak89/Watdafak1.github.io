import streamlit as st


st.set_page_config(page_title="ระบบจัดทำโครงการสอน", page_icon="📚", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@400;500;600&family=Prompt:wght@400;500;600&display=swap');
    :root { --ink: #173e48; --muted: #6f858b; --line: #cbdcdf; --navy: #123f49; --teal: #16aeb0; --blue-soft: #eaf1fb; }
    html, body, [class*="css"] { font-family: 'Prompt', sans-serif; color: var(--ink); }
    .stApp { background-color: #f5fbfb; background-image: linear-gradient(rgba(31,112,119,.07) 1px, transparent 1px), linear-gradient(90deg, rgba(31,112,119,.07) 1px, transparent 1px); background-size: 32px 32px; }
    .block-container { max-width: 1065px; padding-top: 2rem; }
    .hero { min-height: 138px; padding: 24px 36px; margin-bottom: 22px; border: 1px solid #22646d; border-left: 6px solid #5bd1c8; border-radius: 11px; color: white; background: #123f49; text-align: center; box-shadow: 0 12px 20px rgba(18,63,73,.12); }
    .hero h1 { font-family: 'Kanit', sans-serif; margin: 0 0 8px; font-size: 2.25rem; }
    .hero p { margin: 0; color: #c4f1ef; font-size: .85rem; }
    .band { border-left: 4px solid #18a7a7; border-bottom: 1px solid #d7e9e8; color: #176176; font-size: 1.05rem; font-weight: 600; padding: 8px 12px; margin: 12px 0 14px; background: rgba(228,246,243,.28); }
    .card { border: 1px solid #cbdcdf; border-radius: 7px; padding: 18px 20px; background: rgba(245,251,251,.55); min-height: 150px; }
    .card h3 { color: #145363; font-family: 'Kanit', sans-serif; font-weight: 500; margin: 0 0 5px; }
    .card p { color: var(--muted); font-size: .78rem; margin: 0 0 12px; }
    .stButton > button { border-color: #137d84; color: white; background: #16aeb0; border-radius: 7px; }
    .stButton > button:hover { border-color: #0c7278; color: white; background: #119395; }
    [data-testid="stFileUploader"] { background: rgba(255,255,255,.75); border: 1px dashed #b7dfe0; border-radius: 8px; padding: 8px; }
    [data-testid="stTextInput"] input, [data-testid="stSelectbox"] div[data-baseweb="select"] > div { background: var(--blue-soft); border-radius: 8px; }
    </style>
    <div class="hero">
      <h1>ระบบจัดทำโครงการสอน</h1>
      <p>เปลี่ยนแผนการสอนเป็นตารางรายสัปดาห์ ตรวจแก้ข้อมูล และส่งออกตามแบบฟอร์มวิทยาลัย</p>
    </div>
    """,
    unsafe_allow_html=True,
)

top_left, top_right = st.columns(2, gap="large")
with top_left:
    st.markdown('<div class="card"><h3>🔑 Gemini API Key</h3><p>ใส่ API Key แล้วกดบันทึกก่อนวิเคราะห์เอกสาร</p>', unsafe_allow_html=True)
    api_key = st.text_input("Gemini API Key", type="password", placeholder="วาง API Key ของคุณ", label_visibility="collapsed")
    if st.button("บันทึก API Key", key="save_key", use_container_width=True):
        st.success("บันทึก API Key สำหรับเซสชันนี้แล้ว")
    st.markdown("[กดเพื่อรับ Gemini API Key](https://aistudio.google.com/apikey)")
    st.markdown("</div>", unsafe_allow_html=True)

with top_right:
    st.markdown('<div class="card"><h3>📚 ข้อมูลรายวิชา</h3><p>กำหนดข้อมูลหลักสำหรับโครงการสอน</p>', unsafe_allow_html=True)
    course_code = st.text_input("รหัสวิชา", value="21901-2011")
    course_name = st.text_input("ชื่อวิชา", value="การพัฒนาแอปพลิเคชันบนอุปกรณ์เคลื่อนที่")
    level_col, year_col = st.columns(2)
    with level_col:
        level = st.selectbox("ระดับ", ["ปวช.", "ปวส."])
    with year_col:
        year = st.selectbox("ปีที่", ["1", "2", "3"])
    st.markdown("</div>", unsafe_allow_html=True)

documents, schedule = st.columns(2, gap="large")
with documents:
    st.markdown('<div class="band">📚 แบบฟอร์มและโครงการสอน</div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><p>แบบฟอร์มโครงการสอน (.docx)</p>', unsafe_allow_html=True)
    form_file = st.file_uploader("อัปโหลดแบบฟอร์ม", type=["docx"], key="form_file", label_visibility="collapsed")
    st.markdown('<p>แผนการสอน (PDF)</p>', unsafe_allow_html=True)
    lesson_plan = st.file_uploader("อัปโหลดแผนการสอน", type=["pdf"], key="lesson_plan", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

with schedule:
    st.markdown('<div class="band">🎓 กำหนดข้อมูลการเรียน</div>', unsafe_allow_html=True)
    hours_col, weeks_col, semester_col = st.columns(3)
    with hours_col:
        hours = st.number_input("ชั่วโมง/สัปดาห์", min_value=0, value=5, step=1)
    with weeks_col:
        weeks = st.number_input("สัปดาห์/ภาคเรียน", min_value=0, value=18, step=1)
    with semester_col:
        semester = st.text_input("ภาคเรียนที่", value="1/2569")

st.write("")
if st.button("✨ วิเคราะห์แผนการสอน", use_container_width=True, type="primary"):
    if lesson_plan is None:
        st.warning("กรุณาแนบไฟล์แผนการสอน PDF ก่อนเริ่มวิเคราะห์")
    else:
        st.success(f"พร้อมวิเคราะห์ {lesson_plan.name} · {course_code} · {level} ปีที่ {year}")