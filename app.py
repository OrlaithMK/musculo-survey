import json
import streamlit as st
from datetime import time, datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

ACCESS_CODE = "sch202606"   # ← 실제 배포 때 쓸 접속코드로 바꾸세요

st.set_page_config(page_title="근골격계유해요인 사전조사", layout="centered")

# 화면을 넘겨도 입력값 유지 (버튼/표 위젯 키는 건드리지 않음)
for _k in list(st.session_state.keys()):
    if _k.startswith(("btn_", "ed_")):
        continue
    try:
        st.session_state[_k] = st.session_state[_k]
    except Exception:
        pass

if "stage" not in st.session_state:
    st.session_state.stage = "entry"
if "num_jobs" not in st.session_state:
    st.session_state.num_jobs = 1

MONTHS = ["1월", "2월", "3월", "4월", "5월", "6월",
          "7월", "8월", "9월", "10월", "11월", "12월"]

MUSCLE_GROUPS = [
    ("HAND", "✋", "손목과 손가락을 많이 쓰는 활동", "손가락 집기·쥐기, 손목 굽힘·비틀기"),
    ("ARM",  "💪", "팔을 들거나 뻗는/어깨 위로 팔이 올라가는 활동",     "팔 들기·뻗기, 어깨 위 작업"),
    ("LEG",  "🦵", "쪼그려 앉는 자세", "쪼그려 앉기, 무릎 굽힘"),
    ("NECK", "🙇", "목을 비틀거나 굽히는 자세",          "목 굽히기·젖히기·비틀기"),
    ("BACK", "🔁", "허리를 비틀거나 굽히는 자세",        "허리 굽히기·비틀기"),
]
PART_OPTIONS = [f"{e} {l}" for _c, e, l, _d in MUSCLE_GROUPS]
OPT_TO_LABEL = {f"{e} {l}": l for _c, e, l, _d in MUSCLE_GROUPS}


def init_default(key, value):
    if key not in st.session_state:
        st.session_state[key] = value


def _to_int(v):
    return 0 if pd.isna(v) else int(v)


def all_work_names():
    names = []
    for i in range(st.session_state.num_jobs):
        for j in range(st.session_state.get(f"num_works_{i}", 0)):
            n = st.session_state.get(f"work_name_{i}_{j}", "").strip()
            if n and n not in names:
                names.append(n)
    return names


def freq_rows():
    rows = []
    for i in range(st.session_state.num_jobs):
        jobname = st.session_state.get(f"job_name_{i}", "").strip() or f"공정 {i + 1}"
        for j in range(st.session_state.get(f"num_works_{i}", 0)):
            wname = st.session_state.get(f"work_name_{i}_{j}", "").strip()
            if wname:
                rows.append((f"{i}_{j}", jobname, wname))
    return rows


def shift_text(job):
    t = job.get("근무형태", "")
    if t == "주간 고정 근무":
        return f"주간 {job.get('시프트1_시작','')}-{job.get('시프트1_종료','')} / 주{job.get('주근무일수','')}일"
    if t == "2교대":
        return (f"2교대 A {job.get('시프트1_시작','')}-{job.get('시프트1_종료','')} / "
                f"B {job.get('시프트2_시작','')}-{job.get('시프트2_종료','')} / 주{job.get('주근무시간','')}h")
    if t == "3교대":
        return (f"3교대 A {job.get('시프트1_시작','')}-{job.get('시프트1_종료','')} / "
                f"B {job.get('시프트2_시작','')}-{job.get('시프트2_종료','')} / "
                f"C {job.get('시프트3_시작','')}-{job.get('시프트3_종료','')} / 주{job.get('주근무시간','')}h")
    if t == "기타":
        return f"기타: {job.get('기타_내용','')}"
    return ""


# ===================== 구글시트 연결/저장 =====================
@st.cache_resource
def get_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    info = json.loads(st.secrets["gcp_service_account_json"])
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds).open_by_url(st.secrets["sheet_url"])


def get_ws(sh, title, headers):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=max(26, len(headers)))
    if ws.row_values(1) == []:
        ws.append_row(headers)
    return ws


def save_to_sheet(data):
    sh = get_sheet()
    sub_headers = ["제출ID", "작성일시", "부서명", "팀명", "근로자수", "공정수"]
    work_headers = (["제출ID", "작성일시", "부서명", "팀명", "공정번호", "공정명",
                     "근무형태", "근무시간", "작업명", "작업설명", "사용부위"]
                    + MONTHS + ["연간작업일수", "조사여부"])
    sub_ws = get_ws(sh, "submissions", sub_headers)
    work_ws = get_ws(sh, "works", work_headers)

    sid = datetime.now().strftime("%Y%m%d-%H%M%S")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    sub_ws.append_row([sid, now, data["부서명"], data["팀명"],
                       data["근로자수"], data["공정수"]])

    work_to_parts = {}
    for part, works in data["근육군묶기"].items():
        for w in works:
            work_to_parts.setdefault(w, []).append(part)

    rows = []
    for job in data["공정목록"]:
        st_txt = shift_text(job)
        for wk in job["작업목록"]:
            months = [wk["월별작업일수"].get(m, 0) for m in MONTHS]
            parts = ", ".join(work_to_parts.get(wk["작업명"], []))
            rows.append([sid, now, data["부서명"], data["팀명"], job["공정번호"],
                         job["공정명"], job["근무형태"], st_txt, wk["작업명"],
                         wk["작업설명"], parts] + months
                        + [wk["연간작업일수"], wk["조사여부"]])
    if rows:
        work_ws.append_rows(rows)
    return sid


# ===================== 1. 진입 =====================
def show_entry():
    st.title("2026 근골격계유해요인 사전조사(부서장/관리감독자)")
    st.write(
        """
안녕하세요. **2026 근골격계유해요인 사전조사(부서장/관리감독자)** 페이지에 오신 것을 환영합니다.

해당 조사는 산업안전보건법 제39조(보건조치)와 산업안전보건기준에 관한 규칙 제657조(유해요인 조사)에 따라 3년에 1회 시행되는 조사입니다. 해당 조사는 안전보건 주관부서와 (주)서울산업안전컨설팅의 위탁계약을 통해 이루어집니다. 계약상 비밀유지 조항에 따라 해당 조사 외에는 다른 용도로 활용되지 않습니다. 필요시 안전보건 주관부서와의 별도 협의를 통해 민감정보 제출을 거부할 수 있습니다.

이 조사는 유해요인조사 시작 전 예비조사로, 관리감독자·부서장·팀장을 대상으로 하는 조사입니다. 이후 이 조사내 용을 바탕으로 각 부서원들의 조사가 시행되므로, 사내 직제규정·직무설명서·업무기록지 등을 참고하여 작성해 주시기 바랍니다. 필요하다면 부서 내의 회의를 통해 부서원의 의견을 수렴한 후 진행하실 수 있습니다.

이 조사는 **신체적인 동작과 업무 형태**를 기준으로 합니다. 업무의 추상적 특성(내용·스트레스)보다는 **장소·도구·자세·시간** 등 외견상 식별 가능한 특징에 집중하여 작성해 주시기 바랍니다.
        """
    )
    st.divider()
    code = st.text_input("접속코드 입력", type="password",
                         placeholder="배포받은 접속코드를 입력하세요")
    if st.button("시작하기", type="primary", disabled=(code != ACCESS_CODE)):
        st.session_state.stage = "dept"
        st.rerun()


# ===================== 2. 부서 기본정보 =====================
def show_dept():
    st.title("① 부서 기본 정보")
    st.caption("장소·도구·자세·시간 위주로 작성해 주세요.")

    c1, c2 = st.columns(2)
    c1.text_input("부서명 *", key="dept_name", placeholder="예: 간호부")
    c2.text_input("팀명 (필요시)", key="team_name", placeholder="예: 00병동, 00촬영실")
    st.number_input("소속 근로자 총 인원(명)", min_value=0, step=1, key="worker_count")

    st.divider()
    st.subheader("공정 및 근무시간")
    st.caption("부서 내 업무 중 크게 특성이 달라지는 경우 구분. 부서 내 업무특성이 모두 비슷하다면 공정 1개만 작성")

    for i in range(st.session_state.num_jobs):
        with st.container(border=True):
            st.markdown(f"**공정 {i + 1}**")
            st.text_input("공정명 (20자 이내)", key=f"job_name_{i}",
                          max_chars=20, placeholder="예: 간호사 / 간호조무사 / 보조인력 / 이송 등")
            stype = st.radio("근무형태", ["주간 고정 근무", "2교대", "3교대", "기타"],
                             key=f"shift_{i}", horizontal=True)

            if stype == "주간 고정 근무":
                init_default(f"s1_start_{i}", time(9, 0))
                init_default(f"s1_end_{i}", time(18, 0))
                init_default(f"days_{i}", 5)
                a, b, c = st.columns(3)
                a.time_input("시작 시각", key=f"s1_start_{i}")
                b.time_input("종료 시각", key=f"s1_end_{i}")
                c.number_input("주 근무일수", 1, 7, key=f"days_{i}")
            elif stype == "2교대":
                init_default(f"a_start_{i}", time(6, 0))
                init_default(f"a_end_{i}", time(18, 0))
                init_default(f"b_start_{i}", time(18, 0))
                init_default(f"b_end_{i}", time(6, 0))
                init_default(f"hours_{i}", 40)
                a, b = st.columns(2)
                a.time_input("A조 시작", key=f"a_start_{i}")
                b.time_input("A조 종료", key=f"a_end_{i}")
                a.time_input("B조 시작", key=f"b_start_{i}")
                b.time_input("B조 종료", key=f"b_end_{i}")
                st.number_input("주 근무시간", 0, 100, key=f"hours_{i}")
            elif stype == "3교대":
                init_default(f"a_start_{i}", time(6, 0))
                init_default(f"a_end_{i}", time(14, 0))
                init_default(f"b_start_{i}", time(14, 0))
                init_default(f"b_end_{i}", time(22, 0))
                init_default(f"c_start_{i}", time(22, 0))
                init_default(f"c_end_{i}", time(6, 0))
                init_default(f"hours_{i}", 40)
                a, b = st.columns(2)
                a.time_input("A조 시작", key=f"a_start_{i}")
                b.time_input("A조 종료", key=f"a_end_{i}")
                a.time_input("B조 시작", key=f"b_start_{i}")
                b.time_input("B조 종료", key=f"b_end_{i}")
                a.time_input("C조 시작", key=f"c_start_{i}")
                b.time_input("C조 종료", key=f"c_end_{i}")
                st.number_input("주 근무시간", 0, 100, key=f"hours_{i}")
            else:
                st.text_input("근무형태 설명", key=f"etc_{i}",
                              placeholder="예: 격일 근무, 탄력근무제 등")

    col_add, col_del = st.columns(2)
    if col_add.button("➕ 공정 추가"):
        st.session_state.num_jobs += 1
        st.rerun()
    if st.session_state.num_jobs > 1 and col_del.button("➖ 마지막 공정 삭제"):
        st.session_state.num_jobs -= 1
        st.rerun()

    st.divider()
    dept_ok = st.session_state.get("dept_name", "").strip() != ""
    cc1, cc2 = st.columns(2)
    if cc1.button("← 처음으로"):
        st.session_state.stage = "entry"
        st.rerun()
    if cc2.button("다음: 작업 입력 →", type="primary", disabled=not dept_ok):
        st.session_state.stage = "work"
        st.rerun()
    if not dept_ok:
        st.caption("※ 부서명을 입력하면 다음으로 넘어갈 수 있습니다.")


# ===================== 3. 작업(work) 입력 =====================
def show_work():
    st.title("② 작업(Work) 입력")
    st.write("**각 공정에서 1시간 이상 수행되는 작업의 명칭을 모두 적어 주세요.**")
    st.info("장소/신체 동작을 기준으로 하므로 너무 세세히 나누지 않아도 됩니다. "
            "예를 들어, PC를 이용한 업무는 회계·인사·차팅 등을 구분하지 않고 모두 하나의 'PC 작업'입니다.")

    for i in range(st.session_state.num_jobs):
        job_name = st.session_state.get(f"job_name_{i}", "").strip() or f"공정 {i + 1}"
        st.subheader(f"[{job_name}] 의 작업")

        init_default(f"num_works_{i}", 1)
        for j in range(st.session_state[f"num_works_{i}"]):
            with st.container(border=True):
                wc1, wc2 = st.columns([2, 3])
                wc1.text_input(f"작업명 {j + 1}", key=f"work_name_{i}_{j}",
                               placeholder="예: PC작업, 환자 처치")
                wc2.text_input(f"설명 {j + 1}", key=f"work_desc_{i}_{j}",
                               placeholder="예: 일반적인 PC입력, 환자에 대한 간호행위")

        a, b = st.columns(2)
        if st.session_state[f"num_works_{i}"] < 10:
            if a.button("➕ 작업 추가", key=f"btn_add_work_{i}"):
                st.session_state[f"num_works_{i}"] += 1
                st.rerun()
        else:
            a.caption("작업은 공정당 최대 10개입니다.")
        if st.session_state[f"num_works_{i}"] > 1:
            if b.button("➖ 마지막 작업 삭제", key=f"btn_del_work_{i}"):
                st.session_state[f"num_works_{i}"] -= 1
                st.rerun()
        st.divider()

    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "dept"
        st.rerun()
    if cc2.button("다음 →", type="primary"):
        st.session_state.stage = "muscle"
        st.rerun()


# ===================== 4. 작업별 사용 부위 =====================
def show_muscle():
    st.title("③ 작업별 신체활동 특성")
    st.write("**작업마다, 그 작업을 할 때 해당되는 특성을 눌러 주세요.**")
    st.caption("한 작업에 여러 특성 선택 가능"
               "해당하는 특성이 없으면 누르지 않아도 됩니다.")

    works = all_work_names()
    if not works:
        st.warning("먼저 ② 작업 입력에서 작업을 적어 주세요.")
    else:
        for w in works:
            with st.container(border=True):
                st.markdown(f"**{w}**")
                st.pills("부위 선택", options=PART_OPTIONS, selection_mode="multi",
                         key=f"parts__{w}", label_visibility="collapsed")

    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "work"
        st.rerun()
    if cc2.button("다음 →", type="primary"):
        st.session_state.stage = "freq"
        st.rerun()


# ===================== 5. 연간 작업빈도(월별) =====================
def show_freq():
    st.title("④ 연간 작업빈도")
    st.write("**작업마다 월별 작업일수(0~20일)를 입력해 주세요.**")
    st.caption("연간 합계가 60일을 초과(61일 이상)하면 조사 대상입니다. 60일 이하는 간헐적 작업으로 조사에서 제외됩니다.")

    rows = freq_rows()
    if not rows:
        st.warning("먼저 ② 작업 입력에서 작업을 적어 주세요.")
    else:
        if "freq" not in st.session_state:
            st.session_state.freq = {}

        data = []
        for key, jobname, wname in rows:
            saved = st.session_state.freq.get(key, {})
            row = {"공정": jobname, "작업": wname}
            for m in MONTHS:
                row[m] = int(saved.get(m, 0))
            data.append(row)
        df = pd.DataFrame(data)

        colcfg = {"공정": st.column_config.TextColumn(disabled=True, width="small"),
                  "작업": st.column_config.TextColumn(disabled=True, width="medium")}
        for m in MONTHS:
            colcfg[m] = st.column_config.NumberColumn(min_value=0, max_value=20,
                                                      step=1, format="%d", width="small")

        edited = st.data_editor(df, column_config=colcfg, hide_index=True, key="ed_freq")

        summary = []
        for idx, (key, jobname, wname) in enumerate(rows):
            vals = {m: _to_int(edited.iloc[idx][m]) for m in MONTHS}
            st.session_state.freq[key] = vals
            total = sum(vals.values())
            summary.append({"공정": jobname, "작업": wname, "연간작업일수": total,
                            "조사여부": "조사대상" if total > 60 else "제외(간헐적)"})

        st.divider()
        st.subheader("자동 집계")
        st.dataframe(pd.DataFrame(summary), hide_index=True)

    cc1, cc2 = st.columns(2)
    if cc1.button("← 이전"):
        st.session_state.stage = "muscle"
        st.rerun()
    if cc2.button("다음: 확인 →", type="primary"):
        st.session_state.stage = "review"
        st.rerun()


# ===================== 지금까지 입력 모으기 =====================
def collect_all():
    jobs = []
    for i in range(st.session_state.num_jobs):
        stype = st.session_state.get(f"shift_{i}", "")
        job = {"공정번호": i + 1,
               "공정명": st.session_state.get(f"job_name_{i}", "").strip(),
               "근무형태": stype, "작업목록": []}
        if stype == "주간 고정 근무":
            job["시프트1_시작"] = st.session_state[f"s1_start_{i}"].strftime("%H:%M")
            job["시프트1_종료"] = st.session_state[f"s1_end_{i}"].strftime("%H:%M")
            job["주근무일수"] = st.session_state[f"days_{i}"]
        elif stype in ("2교대", "3교대"):
            job["시프트1_시작"] = st.session_state[f"a_start_{i}"].strftime("%H:%M")
            job["시프트1_종료"] = st.session_state[f"a_end_{i}"].strftime("%H:%M")
            job["시프트2_시작"] = st.session_state[f"b_start_{i}"].strftime("%H:%M")
            job["시프트2_종료"] = st.session_state[f"b_end_{i}"].strftime("%H:%M")
            if stype == "3교대":
                job["시프트3_시작"] = st.session_state[f"c_start_{i}"].strftime("%H:%M")
                job["시프트3_종료"] = st.session_state[f"c_end_{i}"].strftime("%H:%M")
            job["주근무시간"] = st.session_state[f"hours_{i}"]
        elif stype == "기타":
            job["기타_내용"] = st.session_state.get(f"etc_{i}", "").strip()

        for j in range(st.session_state.get(f"num_works_{i}", 0)):
            name = st.session_state.get(f"work_name_{i}_{j}", "").strip()
            if name:
                vals = st.session_state.get("freq", {}).get(f"{i}_{j}", {})
                total = sum(int(v) for v in vals.values())
                job["작업목록"].append({
                    "작업명": name,
                    "작업설명": st.session_state.get(f"work_desc_{i}_{j}", "").strip(),
                    "월별작업일수": vals,
                    "연간작업일수": total,
                    "조사여부": "조사대상" if total > 60 else "제외(간헐적)",
                })
        jobs.append(job)

    muscle = {l: [] for _c, _e, l, _d in MUSCLE_GROUPS}
    for w in all_work_names():
        for opt in st.session_state.get(f"parts__{w}", []):
            lbl = OPT_TO_LABEL.get(opt)
            if lbl:
                muscle[lbl].append(w)
    muscle = {k: v for k, v in muscle.items() if v}

    return {"부서명": st.session_state.get("dept_name", "").strip(),
            "팀명": st.session_state.get("team_name", "").strip(),
            "근로자수": st.session_state.get("worker_count", 0),
            "공정수": st.session_state.num_jobs,
            "공정목록": jobs,
            "근육군묶기": muscle}


# ===================== 6. 확인·수정 + 제출 =====================
def show_review():
    st.title("⑤ 확인 및 제출")
    st.caption("※ 브라우저 '뒤로가기'가 아니라 아래 버튼으로 이동·수정하세요. 제출 후에는 수정할 수 없습니다.")

    data = collect_all()

    st.write(f"**부서:** {data['부서명']}  /  **팀:** {data['팀명'] or '-'}  /  "
             f"**인원:** {data['근로자수']}명  /  **공정수:** {data['공정수']}")

    work_rows = [{"공정": job["공정명"], "작업": wk["작업명"],
                  "연간작업일수": wk["연간작업일수"], "조사여부": wk["조사여부"]}
                 for job in data["공정목록"] for wk in job["작업목록"]]
    if work_rows:
        st.dataframe(pd.DataFrame(work_rows), hide_index=True)
    else:
        st.warning("입력된 작업이 없습니다.")

    if data["근육군묶기"]:
        st.markdown("**부위별 작업**")
        for part, works in data["근육군묶기"].items():
            st.write(f"- {part}: {', '.join(works)}")

    st.divider()
    st.markdown("**수정이 필요하면 해당 단계로 이동하세요**")
    e1, e2, e3, e4 = st.columns(4)
    if e1.button("① 부서정보"):
        st.session_state.stage = "dept"; st.rerun()
    if e2.button("② 작업"):
        st.session_state.stage = "work"; st.rerun()
    if e3.button("③ 부위"):
        st.session_state.stage = "muscle"; st.rerun()
    if e4.button("④ 빈도"):
        st.session_state.stage = "freq"; st.rerun()

    st.divider()
    if st.button("제출하기", type="primary"):
        try:
            sid = save_to_sheet(data)
            st.session_state.last_sid = sid
            st.session_state.stage = "done"
            st.rerun()
        except Exception as e:
            st.error("구글시트 저장에 실패했습니다. secrets(시트 URL·인증키)와 시트 공유 권한을 확인해 주세요.")
            st.exception(e)


# ===================== 7. 제출 완료 =====================
def show_done():
    st.title("✅ 제출 완료")
    st.success(f"제출되었습니다. (제출번호: {st.session_state.get('last_sid', '')})")
    st.write("제출 후에는 이 화면에서 수정할 수 없습니다. 내용을 바꿔야 하면 다시 작성해 제출하시거나 안전관리팀에 문의해 주세요.")
    if st.button("다시 작성 (처음으로)"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.session_state.stage = "entry"
        st.rerun()


# ===================== 화면 분기 =====================
stage = st.session_state.stage
if stage == "entry":
    show_entry()
elif stage == "dept":
    show_dept()
elif stage == "work":
    show_work()
elif stage == "muscle":
    show_muscle()
elif stage == "freq":
    show_freq()
elif stage == "review":
    show_review()
elif stage == "done":
    show_done()
else:
    show_entry()
