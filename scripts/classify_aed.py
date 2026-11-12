#!/usr/bin/env python3
"""안산시 AED 설치장소 키워드 규칙 분류 스크립트.

입력 : 공공데이터포털(국립중앙의료원 AED 정보) 안산시 추출 xlsx 의 '원본 데이터' 시트
출력 : data/aed_안산시_분류.csv (UTF-8 BOM)

사용법
    python scripts/classify_aed.py 입력.xlsx data/aed_안산시_분류.csv

분류 결과(접근성 등급)
    A 완전 개방  : 지하철역, 24시 편의점, 대형마트, 병원 응급실, 경찰, 소방
    B 일반 개방  : 관공서·주민센터, 공공·문화·체육·복지시설, 병원, 금융기관, 교통·상업
    C 후순위     : 아파트·공동주택, 학교
    D 기업·공장  : 일반인 접근 제한 가능성 높음 → 앱 기본 노출에서 제외
    U 분류 불명  : 키워드로 판단 불가
    X 이동형     : 구급차 등 → 앱에서 제외

주의: 생성형 AI 분류의 정답이 아니라, 앱 프로토타입용 키워드 규칙이며 현장 검증 전 추정값이다.
"""
import html
import re
import sys

import pandas as pd

DAYS_WEEK = ["mon", "tue", "wed", "thu", "fri"]
# 반월·시화공단 일대로 가정한 동 이름 (주소·설치위치에 포함되면 공단권역 추정). 가정값이므로 검증 필요.
GONGDAN = ["원시동", "목내동", "성곡동", "사사동", "팔곡일동", "팔곡이동", "양상동", "건건동", "부곡동"]
WEIGHT = {"A": 1.0, "B": 0.8, "C": 0.5, "D": 0.2, "U": 0.5, "X": 0.0}

# 키워드 규칙으로 틀린 것을 검토 후 수동으로 바로잡은 기관 (기업 연구소·현장 등)
MANUAL_CORP = r"LG이노텍|삼성물산안산글로벌클라우드센터|현대건설안산성곡동데이터센터"


def s(x):
    return "" if pd.isna(x) else str(x)


def to_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def covers(st, en, a, b):
    """등록 시간 [st,en] (HHMM)이 [a,b]를 포함하는가. 등록 없음/오류는 None."""
    st, en = to_int(st), to_int(en)
    if st is None or en is None or (st == en and st == 0):
        return None
    if en < st:
        en += 2400
    return st <= a and en >= b


def classify_row(r):
    org, place, addr = s(r.org), s(r.buildPlace), s(r.buildAddress)
    txt = org + " " + place
    ev = []

    def has(rx):
        m = re.search(rx, txt)
        if m:
            ev.append(m.group(0))
        return bool(m)

    if has(r"구급차|앰뷸|특수구급"):
        typ, g, sub = "이동형(구급차)", "X", "이동형"
    elif has(r"지하철|\d호선|수인분당선") or (re.search(r"역$", org.strip()) and not re.search(r"소방|경찰", org)):
        typ, g, sub = "지하철역", "A", "지하철역"
    elif has(r"편의점|GS25|CU|세븐일레븐|이마트24|미니스톱"):
        typ, g, sub = "편의점", "A", "편의점"
    elif has(r"홈플러스|이마트|롯데마트|코스트코|하나로마트|트레이더스|대형마트"):
        typ, g, sub = "대형마트", "A", "대형마트"
    elif has(r"응급실|응급의료"):
        typ, g, sub = "병원 응급실", "A", "응급실"
    elif has(r"경찰서|지구대|파출소|치안센터"):
        typ, g, sub = "경찰", "A", "경찰"
    elif has(r"소방서|119|안전센터|구조대"):
        typ, g, sub = "소방", "A", "소방"
    elif has(r"병원|의료원|의원|클리닉"):
        typ, g, sub = "의료기관(응급실 외)", "B", "병원"
    elif has(r"행정복지센터|주민센터|시청|구청|동사무소|출장소|우체국|법원|검찰|세무서|공단$|공사$"):
        typ, g, sub = "관공서·주민센터", "B", "관공서"
    elif has(r"보건소|보건지소|도서관|문화|체육|복지|회관|박물관|경기장|수영장|청소년|경로|센터"):
        typ, g, sub = "공공·문화체육복지", "B", "공공시설"
    elif has(r"은행|농협|신협|새마을금고|금융|증권"):
        typ, g, sub = "금융기관", "B", "금융"
    elif has(r"터미널|백화점|호텔|시장|프라자|쇼핑|몰$|상가"):
        typ, g, sub = "교통·상업", "B", "상업"
    elif has(r"초등|중학|고등|학교|대학|캠퍼스|유치원|학원"):
        typ, g, sub = "학교", "C", "학교"
    elif has(r"아파트|APT|마을|빌리지|푸르지오|래미안|자이|편한세상|힐스테이트|주공|단지|캐슬|주택|오피스텔|기숙사|관리사무소|경비실"):
        typ, g, sub = "아파트·공동주택", "C", "공동주택"
    elif has(r"\(주\)|㈜|주식회사|공장|연구|산업|전자|테크|공단|사업장|제조|물류|기업"):
        typ, g, sub = "기업·공장", "D", "기업"
    else:
        typ, g, sub = "미분류", "U", "미분류"

    # 미분류 보강 규칙
    if g == "U":
        if re.search(r"골프|썰매|스타디움|생활관|체험|홀$|시민홀|미술관|배드민턴|테니스|게이트볼|국궁|공원|올림픽기념|보노마루|인재개발원", txt):
            typ, g, sub = "공공·문화체육복지", "B", "공공시설"
        elif re.search(r"진료소", txt):
            typ, g, sub = "의료기관(응급실 외)", "B", "병원"
        elif re.search(r"지사|지청|교육지원청|운전면허|소년분류|영업소|토지정보|국토정보", txt):
            typ, g, sub = "관공서·주민센터", "B", "관공서"
        elif re.search(r"한국디지털미디어고|초$|중$|고$", org.strip()):
            typ, g, sub = "학교", "C", "학교"
        elif re.search(r"몰렉스|우유|SP삼화|쿠팡|로지스틱스|신축공사|자원회수|가스공사|전력공사|물류", txt):
            typ, g, sub = "기업·공장", "D", "기업"

    # 기업 키워드가 후순위 유형과 겹치는 경우
    corp_kw = bool(re.search(r"\(주\)|㈜|주식회사|공장|연구소|연구원|산업|테크|공단|사업장|제조|물류", txt))
    if g in ("C", "U", "B") and corp_kw and not re.search(r"아파트|학교|병원|센터|시청", txt):
        typ, g, sub = "기업·공장", "D", "기업"

    gd = [d for d in GONGDAN if d in addr or d in txt]
    if g == "D" and gd:
        ev.append("공단권역:" + gd[0])

    # 수동 보정
    if re.search(MANUAL_CORP, org):
        typ, g, sub = "기업·공장", "D", "기업"
        ev.append("수동보정(기업 시설)")

    expo = {"X": "제외(이동형)", "D": "기본 제외(일반인 접근 제한 가능성 높음)"}.get(g, "기본 노출")

    # 시간 요약
    c24 = [covers(r[d + "SttTme"], r[d + "EndTme"], 0, 2359) for d in DAYS_WEEK]
    wd24 = all(c is True for c in c24)
    wdnone = all(c is None for c in c24)
    noon = [covers(r[d + "SttTme"], r[d + "EndTme"], 1200, 1300) for d in DAYS_WEEK]
    wed23 = covers(r.wedSttTme, r.wedEndTme, 2300, 2400)

    lunch = ""
    if typ in ("관공서·주민센터", "공공·문화체육복지", "금융기관") and not wd24 and any(n is True for n in noon):
        lunch = "12:00~13:00 출입 확인 필요"

    return dict(
        번호=r.rnum, 설치기관명=html.unescape(org), 설치위치=html.unescape(place), 주소=html.unescape(addr),
        위도=r.wgs84Lat, 경도=r.wgs84Lon, 시설유형=typ, 세부=sub, 접근성등급=g, 가중치=WEIGHT[g],
        노출정책=expo, 점심시간규칙=lunch,
        평일24시간등록="Y" if wd24 else ("미등록" if wdnone else "N"),
        평일23시등록가능={True: "Y", False: "N", None: "미등록"}[wed23],
        공단권역추정="Y" if gd else "",
        근거키워드=html.unescape(";".join(dict.fromkeys(ev))),
    )


def main(src, dst):
    df = pd.read_excel(src, sheet_name="원본 데이터")
    out = pd.DataFrame([classify_row(r) for _, r in df.iterrows()])
    out.to_csv(dst, index=False, encoding="utf-8-sig")
    print(f"{len(out)}건 저장 → {dst}")
    print(out["접근성등급"].value_counts().to_string())


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
