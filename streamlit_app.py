import hashlib
import json
import re

import streamlit as st

from teaching_plan import PlanError, generate_plan, render_template, validate_plan


st.set_page_config(page_title="ระบบจัดทำโครงการสอน", page_icon="📚", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@400;500;600&family=Prompt:wght@400;500;600&display=swap');
    :root { --ink: #173e48; --muted: #6f858b; --line: #cbdcdf; --navy: #123f49; --teal: #16aeb0; --blue-soft: #eaf1fb; }
    html, body, [class*="css"] { font-family: 'Prompt', sans-serif; color: var(--ink); }
    .stApp { background-color: #f5fbfb; background-image: linear-gradient(rgba(31,112,119,.07) 1px, transparent 1px), linear-gradient(90deg, rgba(31,112,119,.07) 1px, transparent 1px); background-size: 32px 32px; }
    .block-container { max-width: 960px; margin: 0 auto; padding-top: 1.25rem; }
    [data-testid="stToolbar"] { display: none; }
    .hero { min-height: 138px; padding: 24px 36px; margin-bottom: 22px; border: 1px solid #22646d; border-left: 6px solid #5bd1c8; border-radius: 11px; color: white; background: #123f49; text-align: center; box-shadow: 0 12px 20px rgba(18,63,73,.12); }
    .hero h1 { font-family: 'Kanit', sans-serif; margin: 0 0 8px; font-size: 2.25rem; }
    .hero p { margin: 0; color: #c4f1ef; font-size: .85rem; }
    .band { border-left: 4px solid #18a7a7; border-bottom: 1px solid #d7e9e8; color: #176176; font-size: 1.05rem; font-weight: 600; padding: 8px 12px; margin: 12px 0 14px; background: rgba(228,246,243,.28); }
    .field-title { color: #176176; font-size: 1.05rem; font-weight: 600; padding: 8px 12px; margin: 12px 0 8px; border-left: 4px solid #18a7a7; border-bottom: 1px solid #d7e9e8; background: rgba(228,246,243,.28); }
    .card { border: 1px solid #cbdcdf; border-radius: 7px; padding: 18px 20px; background: rgba(245,251,251,.55); min-height: 150px; }
    .api-card [data-testid="stTextInput"] input { border: 2px solid #16aeb0; box-shadow: 0 0 0 3px rgba(22,174,176,.12); background: #fff; }
    .card h3 { color: #145363; font-family: 'Kanit', sans-serif; font-weight: 500; margin: 0 0 5px; text-align: center; }
    .card p { color: var(--muted); font-size: .78rem; margin: 0 0 12px; text-align: center; }
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
    st.markdown('<div class="card api-card"><h3>🔑 Gemini API Key</h3><p>ใส่ API Key แล้วกดบันทึกก่อนวิเคราะห์เอกสาร</p>', unsafe_allow_html=True)
    api_key = st.text_input("Gemini API Key", type="password", placeholder="วาง API Key ของคุณ", label_visibility="collapsed")
    if st.button("บันทึก API Key", key="save_key", use_container_width=True):
        if api_key.strip():
            st.session_state['saved_api_key'] = api_key.strip()
            st.success("บันทึก API Key สำหรับเซสชันนี้แล้ว")
        else:
            st.session_state.pop('saved_api_key', None)
            st.warning("กรุณากรอก API Key")
    st.markdown("[กดเพื่อรับ Gemini API Key](https://aistudio.google.com/apikey)")
    st.caption("เมื่อกดวิเคราะห์ PDF และข้อความใน template จะถูกส่งให้ Google Gemini โดยใช้ API Key ของคุณ")
    st.markdown('<div class="band">📚 แบบฟอร์มโครงการสอน</div>', unsafe_allow_html=True)
    st.markdown('<p class="upload-title">อัปโหลด template DOCX</p>', unsafe_allow_html=True)
    st.caption("ใช้เมื่อเริ่มงาน หรือเมื่อต้องการเปลี่ยนแบบฟอร์ม")
    form_file = st.file_uploader("อัปโหลดแบบฟอร์ม", type=["docx"], key="form_file")
    st.caption("ไม่เกิน 15 MB และต้องมีช่องข้อมูลกับแถวรายสัปดาห์ตามแบบฟอร์ม")
    st.markdown('<p>แผนการสอน (PDF)</p>', unsafe_allow_html=True)
    lesson_plan = st.file_uploader("อัปโหลดแผนการสอน", type=["pdf"], key="lesson_plan", label_visibility="collapsed")
    st.caption("PDF ไม่เกิน 50 MB และ 1,000 หน้า")
    st.markdown("</div>", unsafe_allow_html=True)

with top_right:
    st.markdown('<div class="card"><h3>📚 ข้อมูลรายวิชา</h3><p>กำหนดข้อมูลหลักสำหรับโครงการสอน</p>', unsafe_allow_html=True)
    st.markdown('<div class="field-title">รหัสและชื่อวิชา</div>', unsafe_allow_html=True)
    course_source = st.radio("รหัสและชื่อวิชา", ["อ่านจาก PDF", "กรอกเอง"], horizontal=True, label_visibility="collapsed")
    if course_source == "กรอกเอง":
        course_code = st.text_input("รหัสวิชา", value="20001-104")
        course_name = st.text_input("ชื่อวิชา", value="กฏหมายแรงงาน")
    else:
        course_code = course_name = ''
    st.caption("ช่องที่เว้นว่างจะอ่านจากแผนการสอน PDF")
    curriculum = st.text_input("หลักสูตร", placeholder="อ่านจาก PDF")
    level_col, year_col = st.columns(2)
    with level_col:
        level = st.selectbox("ระดับ", ["อ่านจาก PDF", "ปวช.", "ปวส."])
    with year_col:
        year = st.selectbox("ปีที่", ["อ่านจาก PDF", "1", "2", "3"])
    st.markdown('<div class="band">🎓 กำหนดข้อมูลการเรียน</div>', unsafe_allow_html=True)
    hours_col, weeks_col, semester_col = st.columns(3)
    with hours_col:
        hours = st.number_input("ชั่วโมง/สัปดาห์", min_value=0, max_value=40, value=0, step=1, help="0 = อ่านจาก PDF")
    with weeks_col:
        weeks = st.number_input("สัปดาห์/ภาคเรียน", min_value=1, max_value=52, value=18, step=1)
    with semester_col:
        semester = st.selectbox("ภาคเรียนที่", ["1/2569", "2/2569"])
    st.caption("เลขแผ่นและเลขหน้าเรียงอัตโนมัติเมื่อเปิดเอกสารใน Word")
    st.markdown("</div>", unsafe_allow_html=True)

settings = {
    'code': course_code.strip(), 'subject': course_name.strip(), 'curriculum': curriculum.strip(),
    'level': '' if level == 'อ่านจาก PDF' else level,
    'year_level': '' if year == 'อ่านจาก PDF' else year,
    'h': hours, 'n': weeks, 'term': semester.strip(),
}
template_bytes = form_file.getvalue() if form_file is not None else b''
pdf_bytes = lesson_plan.getvalue() if lesson_plan is not None else b''
fingerprint = hashlib.sha256(
    hashlib.sha256(template_bytes).digest() + hashlib.sha256(pdf_bytes).digest()
    + json.dumps(settings, sort_keys=True, ensure_ascii=False).encode()
).hexdigest()

if st.button("✨ วิเคราะห์แผนการสอน", use_container_width=True, type="primary"):
    st.session_state.pop('generated_plan', None)
    try:
        if form_file is None or lesson_plan is None:
            raise PlanError("กรุณาแนบ template DOCX และแผนการสอน PDF ให้ครบก่อนวิเคราะห์")
        active_key = api_key.strip() or st.session_state.get('saved_api_key', '')
        if not active_key:
            raise PlanError("กรุณากรอก Gemini API Key ก่อนวิเคราะห์")
        with st.status("กำลังสร้างโครงการสอน", expanded=True) as status:
            try:
                plan = generate_plan(active_key, pdf_bytes, template_bytes, settings,
                                     progress=st.write)
                st.write("กำลังเติมข้อมูลลงใน template")
                render_template(template_bytes, plan)
                status.update(label="สร้างโครงการสอนแล้ว", state="complete", expanded=False)
            except Exception:
                status.update(label="สร้างโครงการสอนไม่สำเร็จ", state="error")
                raise
        st.session_state['plan_revision'] = st.session_state.get('plan_revision', 0) + 1
        st.session_state['generated_plan'] = {'fingerprint': fingerprint, 'data': plan.model_dump()}
    except PlanError as exc:
        st.error(str(exc))
    except Exception:
        st.error("สร้างเอกสารไม่สำเร็จ กรุณาตรวจ template แล้วลองใหม่")

result = st.session_state.get('generated_plan')
if result and result['fingerprint'] != fingerprint:
    st.info("ไฟล์หรือข้อมูลรายวิชาเปลี่ยนแล้ว กรุณากดวิเคราะห์ใหม่เพื่อสร้างเอกสารให้ตรงกับข้อมูลล่าสุด")
elif result:
    st.subheader("ตรวจแก้โครงการสอนก่อนดาวน์โหลด")
    data = result['data']
    for note in data['notes']:
        st.info(note)
    revision = st.session_state['plan_revision']
    meta_labels = {'curriculum': 'หลักสูตร', 'code': 'รหัสวิชา', 'subject': 'ชื่อวิชา',
                   'level': 'ระดับ', 'year_level': 'ปีที่', 'h': 'ชั่วโมง/สัปดาห์',
                   'n': 'สัปดาห์/ภาคเรียน', 'term': 'ภาคเรียนที่'}
    with st.expander("ข้อมูลรายวิชาที่จะใส่ในเอกสาร", expanded=True):
        metadata = st.data_editor(
            [{"รายการ": label, "ข้อมูล": str(data[key])} for key, label in meta_labels.items()],
            disabled=['รายการ'], hide_index=True, use_container_width=True,
            key=f'metadata_{revision}',
        )
    edited = st.data_editor(data['weeks'], hide_index=True, use_container_width=True,
        disabled=['w'], key=f'weeks_{revision}', column_config={
            'w': 'สัปดาห์', 't': 'หัวข้อ', 'p': 'Teaching Point',
            'a': 'กิจกรรม', 'm': 'สื่อ', 'e': 'วัดผล',
        })
    try:
        updated = {**data, 'weeks': edited}
        updated.update({key: row['ข้อมูล'] for key, row in zip(meta_labels, metadata)})
        document = render_template(template_bytes, validate_plan(updated))
        safe_code = re.sub(r'[^\w.-]', '_', str(updated['code']))[:60]
        st.download_button("ดาวน์โหลดโครงการสอน Word", data=document,
            file_name=f'โครงการสอน_{safe_code}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True)
        st.caption("ตรวจเนื้อหาและการแบ่งหน้าใน Word ก่อนนำไปใช้งาน")
    except PlanError as exc:
        st.error(str(exc))
