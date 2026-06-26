import streamlit as st
from datetime import time, datetime
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="근골격계유해요인 사전조사 (근로자용)", layout="centered")

st.markdown("""<style>
    .stNumberInput button { display: none !important; }
    .stRadio > div { gap: 0.7rem !important; }
    .stRadio label { padding: 0.3rem 0 !important; }
    [data-testid="stPills"] button {
        min-width: 12em !important;
        padding: 0.8rem 1.2rem !important;
        border-radius: 0.8rem !important;
    }
    [data-testid="stPills"] button p {
        font-size: 1.3rem !important;
        font-weight: 700 !important;
    }
</style>""", unsafe_allow_html=True)

for _k in list(st.session_state.keys()):
    if _k.startswith(("btn_", "ed_", "entry_client_sel")):
        continue
    try:
        st.session_state[_k] = st.session_state[_k]
    except Exception:
        pass

if "stage" not in st.session_state:
    st.session_state.stage = "entry"

def init_default(key, value):
    if key not in st.session_state:
        st.session_state[key] = value

BURDEN_WORKS = [
    {"id":1,"title":"제1호","desc":"**1일 4시간 이상** 집중적으로 자료입력 등을 위해 **키보드 또는 마우스를 조작**하는 작업","note":"1. 키보드나 마우스 조작이 쉼없이 이루어진다(예: 속기사)\n2. 휴식시간을 마음대로 조정할 수 없다."},
    {"id":2,"title":"제2호","desc":"**1일 총 2시간 이상** 목, 어깨, 팔꿈치, 손목 또는 손을 사용하여 **같은 동작을 반복**하는 작업","note":"2분에 10회 이상 반복되어야 하고, 하루 2시간 이상을 포함합니다."},
    {"id":3,"title":"제3호","desc":"**1일 총 2시간 이상** 머리 위에 손이 있거나, 팔꿈치가 어깨 위에 있거나, 팔꿈치를 몸통으로부터 들거나, 팔꿈치를 몸통 뒤쪽에 위치하도록 하는 상태에서 이루어지는 작업","note":"하루 총 2시간 이상 해당 자세가 유지되어야 합니다."},
    {"id":4,"title":"제4호","desc":"지지되지 않은 상태이거나 임의로 자세를 바꿀 수 없는 조건에서, **1일 총 2시간 이상** 목이나 허리를 **구부리거나 비트는** 상태에서 이루어지는 작업","note":"자세를 마음대로 바꿀 수 없는 조건이어야 합니다."},
    {"id":5,"title":"제5호","desc":"**1일 총 2시간 이상** 쪼그리고 앉거나 무릎을 굽힌 자세에서 이루어지는 작업","note":""},
    {"id":6,"title":"제6호","desc":"**1일 총 2시간 이상** 지지되지 않은 상태에서 **1kg 이상**의 물건을 한 손의 손가락으로 집어 옮기거나, **2kg 이상**에 상응하는 힘을 가하여 한 손의 손가락으로 물건을 쥐는 작업","note":"1kg = 우유팩 1개, 2kg 악력 = 무거운 병뚜껑 돌리기 수준"},
    {"id":7,"title":"제7호","desc":"**1일 총 2시간 이상** 지지되지 않은 상태에서 **4.5kg 이상**의 물건을 한 손으로 들거나 동일한 힘으로 쥐는 작업","note":"4.5kg = 2L 생수병 2개 반"},
    {"id":8,"title":"제8호","desc":"**1일 10회 이상 25kg 이상**의 물체를 드는 작업","note":"25kg = 쌀 한 포대"},
    {"id":9,"title":"제9호","desc":"**1일 25회 이상 10kg 이상**의 물체를 무릎 아래에서 들거나, 어깨 위에서 들거나, 팔을 뻗은 상태에서 드는 작업","note":"10kg = 쌀 반포대"},
    {"id":10,"title":"제10호","desc":"**1일 총 2시간 이상**, 분당 2회 이상 **4.5kg 이상**의 물체를 드는 작업","note":""},
    {"id":11,"title":"제11호","desc":"**1일 총 2시간 이상** 시간당 10회 이상 손 또는 무릎을 사용하여 **반복적으로 충격**을 가하는 작업","note":"손바닥으로 치기, 무릎으로 밀기 등"},
]

BODY_PARTS = ["목", "어깨", "팔/팔꿈치", "손/손목/손가락", "허리", "다리/발"]
PARTS_WITH_SIDE = ["어깨", "팔/팔꿈치", "손/손목/손가락", "다리/발"]
FREQ_OPTIONS = ["거의 매일 (~240일/년)", "주 3~4회 (~180일/년)", "주 1~2일 (~80일/년)", "60일 미만 / 단기간"]
FREQ_MAP = {"거의 매일 (~240일/년)": 240, "주 3~4회 (~180일/년)": 180, "주 1~2일 (~80일/년)": 80, "60일 미만 / 단기간": 30}

# ── 구글시트 ──
@st.cache_resource
def get_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    info = json.loads(st.secrets["gcp_service_account_json"])
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds).open_by_url(st.secrets["sheet_url"])

def get_ws(sh, title, headers):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=5000, cols=max(26, len(headers)))
    if not ws.row_values(1):
        ws.append_row(headers)
    return ws

@st.cache_data(ttl=300)
def load_reference():
    try:
        sh = get_sheet()
        ws = sh.worksheet("reference_works")
        return ws.get_all_records()
    except Exception:
        return []

def ref_clients(ref):
    return sorted(set(r.get("업체명", "") for r in ref if r.get("업체명")))

def ref_depts(ref, client):
    return sorted(set(r.get("부서명", "") for r in ref if r.get("업체명") == client and r.get("부서명")))

def ref_procs(ref, client, dept):
    return sorted(set(r.get("공정명", "") for r in ref if r.get("업체명") == client and r.get("부서명") == dept and r.get("공정명")))

def ref_tasks(ref, client, dept, proc):
    return [{"작업명": r["작업명"], "작업설명": r.get("작업설명", "")}
            for r in ref if r.get("업체명") == client and r.get("부서명") == dept
            and r.get("공정명") == proc and r.get("작업명")]

def all_work_names():
    names = []
    descs = {}
    for idx in range(50):
        nk = f"ref_name_{idx}"
        if nk not in st.session_state:
            break
        if st.session_state.get(f"ref_chk_{idx}", False):
            n = st.session_state.get(nk, "")
            if n and n not in names:
                names.append(n)
                descs[n] = st.session_state.get(f"ref_desc_{idx}", "")
    for j in range(st.session_state.get("num_custom_works", 0)):
        n = st.session_state.get(f"custom_work_name_{j}", "").strip()
        if n and n not in names:
            names.append(n)
            descs[n] = st.session_state.get(f"custom_work_desc_{j}", "")
    return names, descs

# ── 저장 ──
def save_worker_data():
    sh = get_sheet()
    sid = datetime.now().strftime("%Y%m%d-%H%M%S")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    stype = st.session_state.get("w_shift", "")
    shift_detail = ""
    if stype == "주간 고정 근무":
        s = st.session_state.get("w_s1_start", time(9,0)).strftime("%H:%M")
        e = st.session_state.get("w_s1_end", time(18,0)).strftime("%H:%M")
        shift_detail = f"{s}-{e} / 주{st.session_state.get('w_days',5)}일"
    elif stype == "2교대":
        shift_detail = (f"A {st.session_state.get('w_a_start',time(6,0)).strftime('%H:%M')}-"
                        f"{st.session_state.get('w_a_end',time(18,0)).strftime('%H:%M')} / "
                        f"B {st.session_state.get('w_b_start',time(18,0)).strftime('%H:%M')}-"
                        f"{st.session_state.get('w_b_end',time(6,0)).strftime('%H:%M')} / "
                        f"주{st.session_state.get('w_hours',40)}h")
    elif stype == "3교대":
        shift_detail = (f"A {st.session_state.get('w_a_start',time(6,0)).strftime('%H:%M')}-"
                        f"{st.session_state.get('w_a_end',time(14,0)).strftime('%H:%M')} / "
                        f"B {st.session_state.get('w_b_start',time(14,0)).strftime('%H:%M')}-"
                        f"{st.session_state.get('w_b_end',time(22,0)).strftime('%H:%M')} / "
                        f"C {st.session_state.get('w_c_start',time(22,0)).strftime('%H:%M')}-"
                        f"{st.session_state.get('w_c_end',time(6,0)).strftime('%H:%M')} / "
                        f"주{st.session_state.get('w_hours',40)}h")
    else:
        shift_detail = st.session_state.get("w_etc_shift", "")
    client = st.session_state.get("w_client", "")
    dept = st.session_state.get("w_dept", "")
    proc = st.session_state.get("w_process", "") or st.session_state.get("w_process_etc", "")
    info_headers = ["제출ID","작성일시","업체명","이름","연령","성별","부서명","팀명","공정",
                    "재직기간_년","재직기간_월","현업무기간_년","현업무기간_월","이전작업",
                    "근무형태","근무상세","휴식_작업분","휴식_휴식분",
                    "여가활동","가사노동","진단여부","진단질병","진단상태",
                    "사고여부","사고부위","육체적부담"]
    info_ws = get_ws(sh, "worker_info", info_headers)
    info_ws.append_row([sid, now, client,
        st.session_state.get("w_name",""), st.session_state.get("w_age",0),
        st.session_state.get("w_gender",""), dept, st.session_state.get("w_team",""), proc,
        st.session_state.get("w_tenure_y",0), st.session_state.get("w_tenure_m",0),
        st.session_state.get("w_task_y",0), st.session_state.get("w_task_m",0),
        st.session_state.get("w_prev_work",""), stype, shift_detail,
        st.session_state.get("w_break_work",0), st.session_state.get("w_break_rest",0),
        st.session_state.get("w_hobby",""), st.session_state.get("w_housework",""),
        st.session_state.get("w_diagnosis",""),
        ", ".join(st.session_state.get("w_diagnosis_diseases",[])),
        st.session_state.get("w_diagnosis_status",""),
        st.session_state.get("w_injury",""),
        ", ".join(st.session_state.get("w_injury_parts",[])),
        st.session_state.get("w_phys_burden","")])
    works, work_descs = all_work_names()
    task_headers = (["제출ID","작성일시","업체명","이름","부서명","공정","작업명","작업설명"]
                    + [f"부담{b['id']}호" for b in BURDEN_WORKS]
                    + ["빈도카테고리","연간작업일수","조사여부"])
    task_ws = get_ws(sh, "worker_tasks", task_headers)
    task_rows = []
    for j, wname in enumerate(works):
        burden_flags = []
        for b in BURDEN_WORKS:
            sel = st.session_state.get(f"burden_{b['id']}_sel", [])
            burden_flags.append("O" if wname in sel else "")
        freq_cat = st.session_state.get(f"freq_cat_{j}")
        annual = FREQ_MAP.get(freq_cat, 0) if freq_cat else 0
        task_rows.append([sid, now, client, st.session_state.get("w_name",""),
            dept, proc, wname, work_descs.get(wname,"")]
            + burden_flags + [freq_cat or "", annual,
            "조사대상" if annual > 60 else "제외(간헐적)"])
    if task_rows:
        task_ws.append_rows(task_rows)
    symp_headers = ["제출ID","작성일시","업체명","이름","부서명",
                    "통증유무","부위","좌우","통증기간","통증강도","증상빈도","지난1주","조치"]
    symp_ws = get_ws(sh, "worker_symptoms", symp_headers)
    has_pain = st.session_state.get("w_has_pain", "아니오")
    if has_pain == "예":
        for part in BODY_PARTS:
            pk = part.replace("/", "_")
            if st.session_state.get(f"pain_chk_{pk}", False):
                side = st.session_state.get(f"symp_side_{pk}", "-") if part in PARTS_WITH_SIDE else "-"
                symp_ws.append_row([sid, now, client, st.session_state.get("w_name",""), dept,
                    "예", part, side,
                    st.session_state.get(f"symp_dur_{pk}",""),
                    st.session_state.get(f"symp_sev_{pk}",""),
                    st.session_state.get(f"symp_freq_{pk}",""),
                    st.session_state.get(f"symp_week_{pk}",""),
                    st.session_state.get(f"symp_treat_{pk}","")])
    else:
        symp_ws.append_row([sid, now, client, st.session_state.get("w_name",""),
                            dept, "아니오", "", "", "", "", "", "", ""])
    return sid

# ===================== 1. 진입 (업체 선택) =====================
def show_entry():
    st.title("2026 근골격계유해요인 사전조사")
    st.markdown("")
    st.write("""
안녕하세요. **2026 근골격계유해요인 사전조사** 페이지에 오신 것을 환영합니다.

해당 조사는 산업안전보건법 제39조(보건조치)와 산업안전보건기준에 관한 규칙 제657조(유해요인 조사)에 따라 3년에 1회 시행되는 조사입니다.
금번 조사는 안전보건 주관부서와 (주)서울산업안전컨설팅의 위탁계약을 통해 이루어집니다. 계약상 비밀유지 조항에 따라 해당 조사 외에는 다른 용도로 활용되지 않습니다. 필요시 안전보건 주관부서와의 별도 협의를 통해 민감정보 제출을 거부할 수 있습니다.

이 조사는 유해요인조사 시작 전 예비조사입니다. 이후 이 조사내용을 바탕으로 본 조사가 시행됩니다.
이 조사는 신체적인 동작과 업무 형태를 기준으로 합니다. 업무의 추상적 특성(내용, 스트레스)보다는 **장소, 도구, 자세, 시간** 등 외견상 식별 가능한 특징에 집중하여 작성해 주시기 바랍니다.

약 **5분에서 10분 정도 소요**될 수 있으니, 집중할 수 있는 환경에서 잘 읽고 답변해 주시기 바랍니다.

    """)
    st.divider()
    st.markdown("### 소속 업체를 선택하세요")
    ref = load_reference()
    clients = ref_clients(ref)
    if clients:
        selected = st.pills("업체 선택", options=clients, key="entry_client_sel",
                            label_visibility="collapsed")
        if selected:
            with st.container(border=True):
                st.markdown(f"#### **{selected}** 소속 근로자가 맞습니까?")
                c1, c2 = st.columns(2)
                if c1.button("예, 시작하기", type="primary", key="btn_entry_yes"):
                    st.session_state.w_client = selected
                    st.session_state.stage = "info"
                    st.rerun()
                if c2.button("아니오, 다시 선택", key="btn_entry_no"):
                    st.session_state.entry_client_sel = None
                    st.rerun()
    else:
        st.warning("현재 진행중인 업체가 없습니다. 관리자에게 문의하세요.")

# ===================== 2. 기본 정보 =====================
def show_info():
    st.title("① 기본 정보")
    client_val = st.session_state.get("w_client", "")
    st.markdown(f"**업체: {client_val}**")
    ref = load_reference()

    st.text_input("이름 *", key="w_name", placeholder="홍길동")
    c1, c2 = st.columns(2)
    c1.number_input("연령 (만)", min_value=15, max_value=80, value=30, key="w_age")
    c2.radio("성별", ["남", "여"], key="w_gender", horizontal=True)

    st.divider()
    st.markdown("#### 소속")
    depts = ref_depts(ref, client_val) if client_val else []
    if depts:
        dept_opts = depts + ["기타 (직접 입력)"]
        dept_sel = st.selectbox("부서명 *", dept_opts, key="w_dept_sel")
        if dept_sel == "기타 (직접 입력)":
            st.text_input("부서명 직접 입력", key="w_dept")
        else:
            st.session_state.w_dept = dept_sel
    else:
        st.text_input("부서명 *", key="w_dept", placeholder="예: 간호부")

    st.text_input("팀명", key="w_team", placeholder="예: 8병동")

    dept_val = st.session_state.get("w_dept", "")
    procs = ref_procs(ref, client_val, dept_val) if client_val and dept_val else []
    if procs:
        proc_opts = procs + ["기타 (직접 입력)"]
        st.selectbox("자신이 속한 공정 *", proc_opts, key="w_process",
                     help="직무규정상 구분되는 인력 유형(예: 간호사, 간호조무사)")
        if st.session_state.get("w_process") == "기타 (직접 입력)":
            st.text_input("공정 직접 입력", key="w_process_etc")
    else:
        st.text_input("공정 *", key="w_process", placeholder="예: 간호사",
                      help="직무규정상 구분되는 인력 유형(예: 간호사, 간호조무사)")

    st.divider()
    st.markdown("#### 경력")
    st.markdown("**현 직장 재직기간**")
    c1, c2 = st.columns(2)
    c1.number_input("년", 0, 50, 0, key="w_tenure_y", format="%d")
    c2.number_input("개월", 0, 11, 0, key="w_tenure_m", format="%d")
    st.markdown("**현재 업무 수행기간**")
    c1, c2 = st.columns(2)
    c1.number_input("년 ", 0, 50, 0, key="w_task_y", format="%d")
    c2.number_input("개월 ", 0, 11, 0, key="w_task_m", format="%d")
    st.text_input("현재 업무 전에 했던 작업", key="w_prev_work", placeholder="예: 비슷한 업무, 물류 창고 상하차, 조리 등 이전 작업")

    st.divider()
    st.markdown("### 근무형태")
    stype = st.radio("", ["주간 고정 근무", "2교대", "3교대", "기타"],
                     key="w_shift", horizontal=True, label_visibility="collapsed")
    if stype == "주간 고정 근무":
        init_default("w_s1_start", time(9,0)); init_default("w_s1_end", time(18,0)); init_default("w_days", 5)
        a, b, c = st.columns(3)
        a.time_input("시작", key="w_s1_start"); b.time_input("종료", key="w_s1_end")
        c.number_input("주 근무일수", 1, 7, key="w_days")
    elif stype == "2교대":
        init_default("w_a_start", time(6,0)); init_default("w_a_end", time(18,0))
        init_default("w_b_start", time(18,0)); init_default("w_b_end", time(6,0))
        init_default("w_hours", 40)
        a, b = st.columns(2)
        a.time_input("A조 시작", key="w_a_start"); b.time_input("A조 종료", key="w_a_end")
        a.time_input("B조 시작", key="w_b_start"); b.time_input("B조 종료", key="w_b_end")
        st.number_input("주 근무시간", 0, 100, key="w_hours")
    elif stype == "3교대":
        init_default("w_a_start", time(6,0)); init_default("w_a_end", time(14,0))
        init_default("w_b_start", time(14,0)); init_default("w_b_end", time(22,0))
        init_default("w_c_start", time(22,0)); init_default("w_c_end", time(6,0))
        init_default("w_hours", 40)
        a, b = st.columns(2)
        a.time_input("A조 시작", key="w_a_start"); b.time_input("A조 종료", key="w_a_end")
        a.time_input("B조 시작", key="w_b_start"); b.time_input("B조 종료", key="w_b_end")
        a.time_input("C조 시작", key="w_c_start"); b.time_input("C조 종료", key="w_c_end")
        st.number_input("주 근무시간", 0, 100, key="w_hours")
    else:
        st.text_input("근무형태 설명", key="w_etc_shift", placeholder="예: 격일, 탄력근무")

    st.markdown("### 휴식시간")
    st.caption("식사시간 제외")
    r1, r2 = st.columns(2)
    r1.number_input("작업 __분마다", 0, 480, 0, step=10, key="w_break_work")
    r2.number_input("__분 휴식", 0, 60, 0, step=5, key="w_break_rest")

    st.divider()
    name_ok = st.session_state.get("w_name", "").strip() != ""
    cc1, cc2 = st.columns(2)
    if cc1.button("← 업체 선택으로"):
        st.session_state.stage = "entry"; st.rerun()
    if cc2.button("다음 →", type="primary", disabled=not name_ok):
        st.session_state.stage = "works"; st.rerun()
    if not name_ok:
        st.caption("※ 이름을 입력하면 다음으로 넘어갑니다.")

# ===================== 3. 작업 입력 =====================
def show_works():
    st.title("② 작업 입력")
    st.markdown("#### 하루 1시간 이상 수행하는 작업을 확인·수정해 주세요.")
    st.caption("매일 수행하지 않더라도 포함해 주세요.")
    st.info("PC 작업은 회계·인사·차팅·총무를 구분하지 않고 하나의 'PC 작업'입니다.")
    client = st.session_state.get("w_client", "")
    dept = st.session_state.get("w_dept", "")
    proc = st.session_state.get("w_process", "") or st.session_state.get("w_process_etc", "")
    ref = load_reference()
    tasks = ref_tasks(ref, client, dept, proc) if client and dept and proc else []
    if tasks:
        st.markdown(f"**[{proc}]** 에 등록된 작업입니다. 해당하지 않는 작업은 **체크를 해제**하세요.")
        for idx, task in enumerate(tasks):
            st.session_state[f"ref_name_{idx}"] = task["작업명"]
            st.session_state[f"ref_desc_{idx}"] = task["작업설명"]
            init_default(f"ref_chk_{idx}", True)
            label = task["작업명"]
            if task["작업설명"]:
                label += f"  —  {task['작업설명']}"
            st.checkbox(label, key=f"ref_chk_{idx}")
    else:
        st.caption("등록된 참고 작업이 없습니다. 아래에 직접 입력해 주세요.")
    st.divider()
    st.markdown("**추가 작업** (위 목록에 없는 작업)")
    init_default("num_custom_works", 1 if not tasks else 0)
    for j in range(st.session_state.num_custom_works):
        with st.container(border=True):
            wc1, wc2 = st.columns([2, 3])
            wc1.text_input(f"작업명 {j+1}", key=f"custom_work_name_{j}", placeholder="예: 환자 간호, 투약 준비")
            wc2.text_input(f"설명 {j+1}", key=f"custom_work_desc_{j}", placeholder="예: 환자의 자세를 변경하거나 의료장비를 조작함")
    a, b = st.columns(2)
    if st.session_state.num_custom_works < 10:
        if a.button("➕ 작업 추가", key="btn_add_wk"):
            st.session_state.num_custom_works += 1; st.rerun()
    if st.session_state.num_custom_works > 0:
        if b.button("➖ 삭제", key="btn_del_wk"):
            st.session_state.num_custom_works = max(0, st.session_state.num_custom_works - 1); st.rerun()
    st.divider()
    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "info"; st.rerun()
    if cc2.button("다음 →", type="primary"):
        st.session_state.stage = "burden_1"; st.rerun()

# ===================== 4~14. 부담작업 =====================
def show_burden(n):
    b = BURDEN_WORKS[n - 1]
    st.title(f"③-{n} {b['title']}")
    st.markdown("##### 수행하시는 업무 중 해당하는 업무가 있다면 선택해 주세요.")
    st.markdown(b["desc"])
    if b["note"]:
        st.info("📋 **판단 가이드**\n\n" + b["note"])
    st.caption("(부연설명 이미지는 추후 추가)")
    works, _ = all_work_names()
    if works:
        st.pills("해당 작업 선택 (여러 개 가능)", options=works,
                 selection_mode="multi", key=f"burden_{b['id']}_sel")
    st.text_input("위 목록에 없는 작업이 있다면 추가",
                  key=f"burden_{b['id']}_extra", placeholder="예: 추가 작업명")
    st.divider()
    cc1, cc2 = st.columns(2)
    prev = f"burden_{n-1}" if n > 1 else "works"
    nxt = f"burden_{n+1}" if n < 11 else "freq"
    if cc1.button("← 이전", key=f"btn_bprev_{n}"):
        st.session_state.stage = prev; st.rerun()
    if cc2.button("해당 없음 / 다음 →", type="primary", key=f"btn_bnext_{n}"):
        st.session_state.stage = nxt; st.rerun()

# ===================== 15. 연간 작업빈도 =====================
def show_freq():
    st.title("④ 연간 작업빈도")
    st.markdown("#### 작업 빈도를 선택해 주세요.")
    st.caption("하루 1시간 이상 수행 기준")
    works, _ = all_work_names()
    if not works:
        st.warning("먼저 ② 작업 입력에서 작업을 선택·입력해 주세요.")
    else:
        for j, wname in enumerate(works):
            with st.container(border=True):
                st.markdown(f"**{wname}**")
                st.pills("빈도", options=FREQ_OPTIONS,
                         key=f"freq_cat_{j}", label_visibility="collapsed")
        st.divider()
        st.markdown("#### 자동 집계")
        rows = []
        for j, wname in enumerate(works):
            cat = st.session_state.get(f"freq_cat_{j}")
            annual = FREQ_MAP.get(cat, 0) if cat else 0
            rows.append({"작업": wname, "연간작업일수": annual,
                         "조사여부": "조사대상" if annual > 60 else "제외(간헐적)"})
        st.dataframe(pd.DataFrame(rows), hide_index=True)
    st.divider()
    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "burden_11"; st.rerun()
    if cc2.button("다음 →", type="primary"):
        st.session_state.stage = "symptom1"; st.rerun()

# ===================== 증상조사 1/2 =====================
def show_symptom1():
    st.title("⑤ 증상조사 (1/2)")
    st.info("자신의 증상을 **과대 또는 과소 평가**하지 않도록 주의하시기 바랍니다.\n"
            "고용노동부 지침에 따라 해당 증상설문조사의 결과는 **근골격계질환의 이환을 부정 또는 입증하는 근거나 반증자료로 활용할 수 없습니다.**")
    st.markdown("#### 1. 규칙적인 여가·취미활동")
    st.markdown("한번에 30분 이상, 1주일에 적어도 2~3회 이상 하는 활동")
    st.radio("", ["게임 등 컴퓨터 관련 활동", "피아노·트럼펫 등 악기연주",
                  "뜨개질·붓글씨 등", "테니스·축구·농구·골프 등 스포츠 활동",
                  "해당사항 없음"], key="w_hobby", index=None)
    st.markdown("")
    st.markdown("#### 2. 하루 평균 가사노동시간")
    st.markdown("밥하기, 빨래, 청소, 2살 미만 아이 돌보기 등")
    st.radio("", ["거의 하지 않는다", "1시간 미만", "1~2시간 미만",
                  "2~3시간 미만", "3시간 이상"], key="w_housework", index=None)
    st.markdown("")
    st.markdown("#### 3. 의사 진단 여부")
    st.markdown("류머티스 관절염 / 당뇨병 / 루프스병 / 통풍 / 알코올중독 등의 질병에 대해")
    st.radio("의사로부터 진단 받은 적이 있습니까?", ["아니오", "예"],
             key="w_diagnosis", horizontal=True, index=None)
    if st.session_state.get("w_diagnosis") == "예":
        st.multiselect("해당 질병", ["류머티스 관절염", "당뇨병", "루프스병", "통풍", "알코올중독"],
                       key="w_diagnosis_diseases")
        st.radio("현재 상태", ["완치", "치료나 관찰 중"],
                 key="w_diagnosis_status", horizontal=True, index=None)
    st.markdown("")
    st.markdown("#### 4. 과거 사고 이력")
    st.markdown("운동 중 혹은 사고(교통사고, 넘어짐, 추락 등)로 인한 상해")
    st.radio("다친 적이 있습니까?", ["아니오", "예"],
             key="w_injury", horizontal=True, index=None)
    if st.session_state.get("w_injury") == "예":
        st.multiselect("상해 부위", ["손/손가락/손목", "팔/팔꿈치", "어깨", "목", "허리", "다리/발"],
                       key="w_injury_parts")
    st.markdown("")
    st.markdown("#### 5. 육체적 부담 정도")
    st.markdown("현재 하시는 일의 육체적 부담 정도는 어떻습니까?")
    st.radio("", ["전혀 힘들지 않음", "견딜만 함", "약간 힘듦", "힘듦", "매우 힘듦"],
             key="w_phys_burden", index=None)
    st.divider()
    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "freq"; st.rerun()
    if cc2.button("다음 →", type="primary"):
        st.session_state.stage = "symptom2"; st.rerun()

# ===================== 증상조사 2/2 + 제출 =====================
def show_symptom2():
    st.title("⑤ 증상조사 (2/2)")
    st.markdown("#### 지난 1년 동안 통증이나 불편함을 느끼신 적이 있습니까?")
    st.markdown("목, 어깨, 팔/팔꿈치, 손/손목/손가락, 허리, 다리/발 중 어느 한 부위에서라도 작업과 관련하여 **통증·쑤심·찌릿찌릿함·뻣뻣함·화끈거림·무감각** 등을 겪은 경우")
    st.radio("", ["아니오", "예"], key="w_has_pain", horizontal=True, index=None)
    selected_parts = []
    if st.session_state.get("w_has_pain") == "예":
        st.markdown("#### 통증 부위를 모두 체크하세요")
        for part in BODY_PARTS:
            pk = part.replace("/", "_")
            st.checkbox(part, key=f"pain_chk_{pk}")
            if st.session_state.get(f"pain_chk_{pk}", False):
                selected_parts.append(part)
        for part in selected_parts:
            pk = part.replace("/", "_")
            with st.expander(f"📍 {part} 상세", expanded=True):
                if part in PARTS_WITH_SIDE:
                    st.markdown("**1. 통증의 구체적 부위**")
                    st.radio("좌우", ["오른쪽", "왼쪽", "양쪽 모두"],
                             key=f"symp_side_{pk}", index=None, label_visibility="collapsed")
                st.markdown("**2. 통증 지속 기간**")
                st.radio("기간", ["1일 미만", "1일~1주일 미만", "1주일~1달 미만",
                              "1달~6개월 미만", "6개월 이상"],
                         key=f"symp_dur_{pk}", index=None, label_visibility="collapsed")
                st.markdown("**3. 아픈 정도**")
                st.radio("강도", ["약한 통증 (불편하나 작업 중 못 느낌)",
                          "중간 통증 (작업 중 있으나 귀가 후 괜찮음)",
                          "심한 통증 (비교적 심하고 귀가 후에도 계속)",
                          "매우 심한 통증 (일상생활 어려움)"],
                         key=f"symp_sev_{pk}", index=None, label_visibility="collapsed")
                st.markdown("**4. 지난 1년간 증상 빈도**")
                st.radio("빈도", ["6개월에 1번", "2~3달에 1번", "1달에 1번",
                          "1주일에 1번", "매일"],
                         key=f"symp_freq_{pk}", index=None, label_visibility="collapsed")
                st.markdown("**5. 지난 1주일 동안에도 증상이 있었습니까?**")
                st.radio("최근", ["아니오", "예"],
                         key=f"symp_week_{pk}", index=None,
                         horizontal=True, label_visibility="collapsed")
                st.markdown("**6. 지난 1년간 통증으로 인한 조치**")
                st.radio("조치", ["병원·한의원 치료", "약국 치료", "병가·산재",
                          "작업 전환", "해당사항 없음"],
                         key=f"symp_treat_{pk}", index=None, label_visibility="collapsed")
    st.divider()
    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "symptom1"; st.rerun()
    if cc2.button("제출하기", type="primary", key="btn_submit"):
        if st.session_state.get("w_has_pain") is None:
            st.error("⚠️ 통증 유무를 선택해 주세요."); return
        if st.session_state.get("w_has_pain") == "예" and selected_parts:
            for part in selected_parts:
                pk = part.replace("/", "_")
                checks = [f"symp_dur_{pk}", f"symp_sev_{pk}",
                          f"symp_freq_{pk}", f"symp_week_{pk}", f"symp_treat_{pk}"]
                if part in PARTS_WITH_SIDE:
                    checks.append(f"symp_side_{pk}")
                for ck in checks:
                    if st.session_state.get(ck) is None:
                        st.error(f"⚠️ **{part}** 의 모든 질문에 답변해 주세요."); return
        try:
            sid = save_worker_data()
            st.session_state.last_sid = sid
            st.session_state.stage = "done"; st.rerun()
        except Exception as e:
            st.error("저장 실패. secrets/시트 공유 설정을 확인하세요."); st.exception(e)

# ===================== 완료 =====================
def show_done():
    st.title("✅ 제출 완료")
    st.success(f"수고하셨습니다! (제출번호: {st.session_state.get('last_sid', '')})")
    st.write("내용을 바꿔야 하면 다시 작성해 제출하거나 안전관리팀에 문의하세요.")
    if st.button("새 작성 (처음으로)"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.session_state.stage = "entry"; st.rerun()

# ===================== 화면 분기 =====================
stage = st.session_state.stage
if stage == "entry": show_entry()
elif stage == "info": show_info()
elif stage == "works": show_works()
elif stage.startswith("burden_"): show_burden(int(stage.split("_")[1]))
elif stage == "freq": show_freq()
elif stage == "symptom1": show_symptom1()
elif stage == "symptom2": show_symptom2()
elif stage == "done": show_done()
else: show_entry()
