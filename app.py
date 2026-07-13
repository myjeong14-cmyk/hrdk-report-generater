import streamlit as st
import time
import os
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Mm
from playwright.sync_api import sync_playwright

DESTINATION_DB = {
    "경북소프트웨어마이스터고등학교": {"dist": 47, "round_dist": 94, "path": "안동→의성→안동"},
    "경북에너지기술고등학교": {"dist": 73, "round_dist": 146, "path": "안동→상주→안동"},
    "경북자연과학고등학교": {"dist": 84, "round_dist": 168, "path": "안동→상주→안동"},
    "경북조리과학고등학교": {"dist": 68, "round_dist": 136, "path": "안동→문경→안동"},
    "경북직업훈련교도소": {"dist": 51, "round_dist": 102, "path": "안동→청송→안동"},
    "경북항공고등학교": {"dist": 47, "round_dist": 94, "path": "안동→영주→안동"},
    "문경공업고등학교": {"dist": 48, "round_dist": 96, "path": "안동→문경→안동"},
    "산림조합중앙회임업인종합연수원": {"dist": 65, "round_dist": 130, "path": "안동→청송→안동"},
    "상주공업고등학교": {"dist": 75, "round_dist": 150, "path": "안동→상주→안동"},
    "상주교도소": {"dist": 63, "round_dist": 126, "path": "안동→상주→안동"},
    "상주중장비운전학원": {"dist": 72, "round_dist": 144, "path": "안동→상주→안동"},
    "상지미래경영고등학교": {"dist": 51, "round_dist": 102, "path": "안동→상주→안동"},
    "의성유니텍고등학교": {"dist": 36, "round_dist": 72, "path": "안동→의성→안동"},
    "한국미래농업고등학교": {"dist": 99, "round_dist": 198, "path": "안동→상주→안동"},
    "한국미래산업고등학교": {"dist": 32, "round_dist": 64, "path": "안동→영주→안동"},
    "한국산림과학고등학교": {"dist": 70, "round_dist": 140, "path": "안동→봉화→안동"},
    "한국철도고등학교": {"dist": 33, "round_dist": 66, "path": "안동→영주→안동"},
    "한국철도공사 영주역": {"dist": 31, "round_dist": 62, "path": "안동→영주→안동"},
    "한국펫고등학교": {"dist": 52, "round_dist": 104, "path": "안동→봉화→안동"},
    "한국폴리텍대학 영주캠퍼스": {"dist": 35, "round_dist": 70, "path": "안동→영주→안동"},
    "현대건설중장비직업전문학원": {"dist": 32, "round_dist": 64, "path": "안동→영주→안동"}
}

# =========================
# 팝업 제거 (최종 안정 버전)
# =========================
def close_popups(page):
    try:
        page.keyboard.press("Escape")
    except:
        pass

    try:
        page.evaluate("""
            () => {
                // 1. 흔한 팝업/레이어 제거
                document.querySelectorAll(
                    '[class*="popup"], [id*="popup"], [class*="layer"], .modal, .dim, .overlay'
                ).forEach(el => el.remove());

                // 2. iframe 팝업 제거 (오피넷 핵심)
                document.querySelectorAll('iframe').forEach(el => el.remove());

                // 3. body 스크롤 잠금 해제
                document.body.style.overflow = 'auto';
                document.body.style.position = 'static';

                // 4. 화면 덮는 div 강제 제거
                document.querySelectorAll('*').forEach(el => {
                    const style = window.getComputedStyle(el);
                    if (
                        style.position === 'fixed' &&
                        parseInt(style.zIndex || 0) > 1000
                    ) {
                        el.remove();
                    }
                });
            }
        """)
    except:
        pass


# =========================
# hover 메뉴
# =========================
def hover_menu(page, text):
    for _ in range(3):
        try:
            close_popups(page)
            page.get_by_text(text, exact=False).first.hover(timeout=3000)
            return True
        except:
            time.sleep(1)
    return False


# =========================
# 안정 클릭 함수 (핵심)
# =========================
def click_menu_safe(page, text):
    for _ in range(5):
        try:
            close_popups(page)
            page.wait_for_timeout(800)

            locator = page.locator(f"text={text}")
            locator.first.wait_for(timeout=3000)
            locator.first.click(force=True)

            return True
        except:
            page.wait_for_timeout(800)

    return False


# =========================
# 기타 유틸
# =========================
def find_matched_map_image(dest_name):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    map_dir = os.path.join(base_dir, "map")
    for ext in [".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".JPEG"]:
        target = os.path.join(map_dir, f"{dest_name}{ext}")
        if os.path.exists(target):
            return target
    return None


def select_daily_tab(page):
    try:
        page.evaluate("""
            () => {
                const els = Array.from(document.querySelectorAll('a, label, span, button, li'));
                const target = els.find(e => e.textContent.trim() === '일간');
                if (target) { target.click(); return true; }
                return false;
            }
        """)
    except:
        pass


def set_opinet_date(page, date_obj):
    year = str(date_obj.year)
    month = str(date_obj.month)
    day = str(date_obj.day)

    page.evaluate(f"""
        (() => {{
            const target = {{ year: '{year}', month: '{month}', day: '{day}' }};

            function trySet(select, value) {{
                const opt = Array.from(select.options).find(
                    o => o.text.trim() === String(value) || o.value === String(value)
                );
                if (opt) {{
                    select.value = opt.value;
                    select.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    return true;
                }}
                return false;
            }}

            const selects = Array.from(document.querySelectorAll('select'))
                .filter(s => s.offsetParent !== null);

            const yearSelects = selects.filter(s =>
                Array.from(s.options).some(o => o.text.trim() === target.year)
            );
            const monthSelects = selects.filter(s =>
                s.options.length <= 13 &&
                Array.from(s.options).some(o => o.text.trim() === target.month)
            );
            const daySelects = selects.filter(s =>
                s.options.length >= 28 &&
                Array.from(s.options).some(o => o.text.trim() === target.day)
            );

            yearSelects.forEach(s => trySet(s, target.year));
            monthSelects.forEach(s => trySet(s, target.month));
            daySelects.forEach(s => trySet(s, target.day));
        }})()
    """)


def click_query(page):
    for sel in ["a:has-text('조회')", "button:has-text('조회')", "input[value='조회']"]:
        try:
            loc = page.locator(sel)
            if loc.count() > 0:
                loc.first.click(force=True)
                return True
        except:
            pass
    return False


def wait_result_update(page):
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except:
        pass
    time.sleep(3)


# =========================
# 오피넷 캡처 핵심
# =========================
def capture_opinet_print_page(target_date_obj, fuel_type):
    filename = "opinet_capture.png"
    oil_price = 1640 if "휘발유" in fuel_type else 1510
    date_str = target_date_obj.strftime("%Y%m%d")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
            )

            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                viewport={"width": 1400, "height": 950}
            )

            page = context.new_page()

            # 1. 평균판매가격(제품별) 페이지로 직접 이동 (메뉴 클릭 대신 URL 직접 접속)
            page.goto(
                "https://www.opinet.co.kr/user/dopospdrg/dopOsPdrgSelect.do",
                wait_until="domcontentloaded",
                timeout=20000
            )
            page.wait_for_timeout(2000)

            close_popups(page)

            # 2. 구분: 일간 탭 선택
            select_daily_tab(page)
            page.wait_for_timeout(800)

            close_popups(page)

            # 3. 날짜 설정 (연/월/일 드롭다운)
            set_opinet_date(page, target_date_obj)
            page.wait_for_timeout(800)

            # 4. 조회
            if not click_query(page):
                raise Exception("조회 버튼 실패")

            wait_result_update(page)

            # 7. 가격 추출
            try:
                text = page.locator("table td").first.inner_text()
                extracted_num = int("".join(filter(str.isdigit, text)))
                if extracted_num > 1000:
                    oil_price = extracted_num
            except:
                pass

            # 8. 인쇄창
            print_page = None

            try:
                with context.expect_page(timeout=7000) as pop:
                    try:
                        page.evaluate("chkPrint();")
                    except:
                        page.get_by_text("화면인쇄").first.click(force=True)

                    page.wait_for_timeout(1500)

                print_page = pop.value
            except:
                if len(context.pages) > 1:
                    print_page = context.pages[-1]

            if print_page is None:
                raise Exception("인쇄창 없음")

            print_page.screenshot(path=filename, full_page=True)

            browser.close()
            return filename, oil_price

    except Exception as e:
        st.error(f"오피넷 캡처 중 에러 발생: {e}")
        try:
            if 'context' in locals() and len(context.pages) > 0:
                context.pages[-1].screenshot(path=filename, full_page=True)
        except:
            pass
        try:
            if 'browser' in locals():
                browser.close()
        except:
            pass
        return filename, oil_price


# =========================
# DOC 생성 (원본 '시외출장 지출 증빙 양식.hwp'에서 추출한
# 실제 폰트/글자크기/여백 값을 그대로 반영)
# =========================
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_run_font(run, font_name, size_pt, bold=False):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def create_docx_report(data_dict, map_image_path, opinet_image_path="opinet_capture.png"):
    doc = Document()

    # 원본 양식 여백: 위 25.6mm, 아래 17.4mm, 좌우 20mm
    for section in doc.sections:
        section.top_margin = Mm(25.6)
        section.bottom_margin = Mm(17.4)
        section.left_margin = Mm(20)
        section.right_margin = Mm(20)

    # 제목 (원본 양식: 맑은 고딕 18pt bold, 가운데 정렬)
    title = doc.add_paragraph()
    title.alignment = 1
    run = title.add_run("시외출장 지출(개인차량) 증빙 내역")
    set_run_font(run, "맑은 고딕", 18, bold=True)

    # 본문 표 (제공된 '시외출장 지출 증빙 양식.hwp'의 실제 표 구조를 그대로 재현)
    # 4열: 라벨 | 값 | 라벨 | 값,  5행 + 지도/오피넷 병합 행 2개
    col_widths = [Mm(39.06), Mm(41.99), Mm(43.46), Mm(43.46)]

    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    table.autofit = False

    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = col_widths[idx]

    # (왼쪽 라벨, 값, 오른쪽 라벨, 값) - 원본 양식에서 왼쪽 라벨은 일반체, 오른쪽 라벨은 굵게
    rows_data = [
        ("운행일시", data_dict["date"], "유류비(원)", f"{data_dict['fuel_cost']:,}"),
        ("출장지", data_dict["path"], "통행료", f"{data_dict['toll']:,}"),
        ("거리(km)", f"{data_dict['distance']} km (왕복)", "일 비", f"{data_dict['daily_allowance']:,}"),
        ("연비(km/ℓ)", str(data_dict["efficiency"]), "식 비", f"{data_dict['meal_allowance']:,}"),
        ("유가(원,오피넷기준)", f"{data_dict['oil_price']:,}", "총 계", f"{data_dict['total_cost']:,}"),
    ]

    for i, (l1, v1, l2, v2) in enumerate(rows_data):
        cells = table.rows[i].cells

        cells[0].text = ""
        r = cells[0].paragraphs[0].add_run(l1)
        set_run_font(r, "맑은 고딕", 12, bold=False)

        cells[1].text = ""
        r = cells[1].paragraphs[0].add_run(str(v1))
        set_run_font(r, "함초롬바탕", 10, bold=False)

        cells[2].text = ""
        r = cells[2].paragraphs[0].add_run(l2)
        set_run_font(r, "맑은 고딕", 12, bold=True)

        cells[3].text = ""
        r = cells[3].paragraphs[0].add_run(str(v2))
        set_run_font(r, "함초롬바탕", 10, bold=False)

    # 경로 네이버지도 스크린샷 (양식과 동일하게 4열 전체 병합된 1개 행)
    map_row = table.add_row()
    map_cell = map_row.cells[0]
    for c in map_row.cells[1:]:
        map_cell = map_cell.merge(c)
    map_cell.paragraphs[0].alignment = 1
    if map_image_path and os.path.exists(map_image_path):
        map_cell.paragraphs[0].add_run().add_picture(map_image_path, width=Mm(160))
    else:
        r = map_cell.paragraphs[0].add_run("경로 네이버지도 스크린샷")
        set_run_font(r, "맑은 고딕", 12, bold=False)

    # 오피넷 스크린샷 (양식과 동일하게 4열 전체 병합된 1개 행)
    opinet_row = table.add_row()
    opinet_cell = opinet_row.cells[0]
    for c in opinet_row.cells[1:]:
        opinet_cell = opinet_cell.merge(c)
    opinet_cell.paragraphs[0].alignment = 1
    if opinet_image_path and os.path.exists(opinet_image_path):
        opinet_cell.paragraphs[0].add_run().add_picture(opinet_image_path, width=Mm(160))
    else:
        r = opinet_cell.paragraphs[0].add_run("오피넷 스크린샷")
        set_run_font(r, "맑은 고딕", 12, bold=False)

    output = "출장지출증빙_보고서.docx"
    doc.save(output)
    return output


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="출장 지출 증빙", layout="centered")

st.title("시외출장 지출 증빙 보고서 생성기")

st.markdown("""
<style>
/* 1. 모든 일반 버튼 및 다운로드 버튼 박스 기본 스타일 */
div.stButton > button, 
div.stDownloadButton > button {
    background-color: #191970 !important;
    height: 3.4em !important;
    border-radius: 12px !important;
    border: none !important;
    box-shadow: 0 4px 8px rgba(56, 182, 255, 0.35) !important;
    transition: all 0.2s ease;
}

/* 2. 모든 버튼 내부의 "안쪽 글씨" 스타일 설정 */
div.stButton > button p, 
div.stDownloadButton > button p {
    color: white !important;
    font-size: 21px !important;
    font-weight: 600 !important;
}

/* 3. 마우스 올렸을 때 효과 (Hover) */
div.stButton > button:hover, 
div.stDownloadButton > button:hover {
    background-color: #6495ED !important;
    box-shadow: 0 6px 12px rgba(56, 182, 255, 0.45) !important;
}
</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    run_date = st.date_input("운행일시", datetime.today())
with col2:
    dest_selection = st.selectbox("도착지", list(DESTINATION_DB.keys()))

col3, col4 = st.columns(2)
with col3:
    fuel_selection = st.radio("연료", ["휘발유 (10.06 km/ℓ)", "경유 (10.16 km/ℓ)"])
with col4:
    toll_input = st.number_input("통행료", 0, step=100)

col5, col6 = st.columns(2)
with col5:
    daily_fee = st.number_input("일비", 0, step=1000)
with col6:
    meal_fee = st.number_input("식비", 25000, step=1000)

matched_img_file = find_matched_map_image(dest_selection)

st.write("---")

if "report_ready" not in st.session_state:
    st.session_state.report_ready = False

if st.button("보고서 생성", use_container_width=True):
    db_info = DESTINATION_DB[dest_selection]
    round_distance = db_info["round_dist"]

    efficiency = 10.06 if "휘발유" in fuel_selection else 10.16

    with st.spinner("오피넷 조회 중..."):
        opinet_img, oil_price = capture_opinet_print_page(run_date, fuel_selection)

    fuel_cost = int(round_distance / efficiency * oil_price)
    total_cost = fuel_cost + toll_input + daily_fee + meal_fee

    report_data = {
        "date": run_date.strftime("%Y년 %m월 %d일"),
        "path": db_info["path"],
        "distance": round_distance,
        "efficiency": efficiency,
        "oil_price": oil_price,
        "fuel_cost": fuel_cost,
        "toll": toll_input,
        "daily_allowance": daily_fee,
        "meal_allowance": meal_fee,
        "total_cost": total_cost
    }

    file = create_docx_report(report_data, matched_img_file, opinet_img)

    st.session_state.report_ready = True
    st.session_state.report_file = file
    st.session_state.opinet_img = opinet_img
    st.session_state.matched_img_file = matched_img_file

if st.session_state.report_ready:
    st.success("보고서 생성이 완료되었습니다.")

    report_file = st.session_state.report_file
    opinet_img = st.session_state.opinet_img
    matched_img_file = st.session_state.matched_img_file

    with open(report_file, "rb") as f:
        st.download_button(
            "보고서 다운로드",
            f,
            file_name="출장지출증빙_보고서.docx",
            use_container_width=True
        )

    col_a, col_b = st.columns(2)
    col_a, col_b = st.columns(2)
    with col_a:
        if matched_img_file and os.path.exists(matched_img_file):
            st.image(matched_img_file, caption="네이버지도 경로")
            with open(matched_img_file, "rb") as f_map:
                import base64
                map_bytes = f_map.read()
                b64 = base64.b64encode(map_bytes).decode()
                href = f'data:application/octet-stream;base64,{b64}'
                btn_html = (
                    f'<a href="{href}" download="{os.path.basename(matched_img_file)}" target="_blank" style="text-decoration: none; width: 100%;">'
                    f'  <button style="'
                    f'      width: 100%;'
                    f'      background-color: transparent;'
                    f'      color: rgb(49, 51, 63);'
                    f'      border: 1px solid rgba(49, 51, 63, 0.2);'
                    f'      border-radius: 0.5rem;'
                    f'      padding: 0.4rem 0.75rem;'
                    f'      font-size: 14px;'
                    f'      font-weight: 400;'
                    f'      line-height: 1.6;'
                    f'      cursor: pointer;'
                    f'      text-align: center;'
                    f'      font-family: inherit;'
                    f'  ">경로 캡처 다운받기</button>'
                    f'</a>'
                )

                st.markdown(btn_html, unsafe_allow_html=True)
        else:
            st.warning("map 폴더에서 일치하는 지도 사진을 찾지 못했습니다.")
    with col_b:
        if os.path.exists(opinet_img):
            st.image(opinet_img, caption="오피넷 화면인쇄 증빙")
            with open(opinet_img, "rb") as f_oil: # 📝 f 대신 f_oil 사용
                import base64
                opinet_bytes = f_oil.read()
                b64 = base64.b64encode(opinet_bytes).decode()
                href = f'data:application/octet-stream;base64,{b64}'
                
                st.markdown(f'''
                    <a href="{href}" download="opinet_capture.png" target="_blank" style="text-decoration: none; width: 100%;">
                        <button style="
                            width: 100%;
                            background-color: transparent;
                            color: rgb(49, 51, 63);
                            border: 1px solid rgba(49, 51, 63, 0.2);
                            border-radius: 0.5rem;
                            padding: 0.4rem 0.75rem;
                            font-size: 14px;
                            font-weight: 400;
                            line-height: 1.6;
                            cursor: pointer;
                            text-align: center;
                            font-family: inherit;
                        ">유가 캡처 다운받기</button>
                    </a>
                ''', unsafe_allow_html=True)
        else:
            st.warning("오피넷 캡처 이미지를 생성하지 못했습니다. 위 에러 메시지를 확인해주세요.")