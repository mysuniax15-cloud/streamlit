import streamlit as st
import pandas as pd
from openai import OpenAI
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from io import BytesIO
from matplotlib.ticker import MaxNLocator, MultipleLocator
from matplotlib.patches import Wedge
from reportlab.lib.colors import HexColor
from zipfile import ZipFile
from matplotlib import colors as mcolors
import time, random
from openai import RateLimitError
import os
import hashlib
import matplotlib.pyplot as plt
import numpy as np
import re
import textwrap
import base64
import unicodedata
import matplotlib
matplotlib.rcParams['font.family'] = 'Malgun Gothic'
matplotlib.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지
from matplotlib.ticker import MaxNLocator
from matplotlib.patches import Wedge
















# ====== Streamlit UI & CSS ======
st.set_page_config(page_title="리더십 분석기", layout="wide")




















# ====== A.X 4.0 API Key 설정 ======
SKT_AX4_KEY = "sktax-XyeKFrq67ZjS4EpsDlrHHXV8it"
client = OpenAI(
    base_url="https://guest-api.sktax.chat/v1",
    api_key=SKT_AX4_KEY
)
















# 대시보드에서 사용할 7개 객관식 컬럼
VISUAL_COLS = [
    "팀원_자긍심",
    "팀원_공동체의식",
    "팀원_상호배려",
    "팀원_내 일 알기",
    "팀원_도전적 목표 설정",
    "팀원_철저하고 즐거운 실행",
    "팀원_지식공유와 역량개발"
]
















# 사용자에게 보여줄 짧은 한글 라벨
LABEL_MAP = {
    "팀원_자긍심": "자긍심",
    "팀원_공동체의식": "공동체의식",
    "팀원_상호배려": "상호배려",
    "팀원_내 일 알기": "내 일 알기",
    "팀원_도전적 목표 설정": "도전목표",
    "팀원_철저하고 즐거운 실행": "즐거운실행",
    "팀원_지식공유와 역량개발": "지식공유"
}




FOCUS_YEAR = 2024
LEADER_NAME_COL = "이름"
LEADER_ID_COL   = "팀장 ID"
LEADER_KEY_COL  = "리더키"   # 내부용 고유키(표시는 안 함)
BAR_WIDTH = 0.42




LEADER_COL = "이름"   # 엑셀 컬럼명에 맞춰 수정




COLOR_SELF_BAR  = "#B91C1C"   # 22~24년 막대(자기 점수)
COLOR_CORP_LINE = "#64748B"   # 전사 평균 라인 (레이더/추이 공통)
COLOR_SELF_LINE = "#2563EB"   # 레이더에서 자기 라인/채움


PRIMARY = "#EE0000"  # 브랜드 레드 (SK Red)

BASE_2024 = "#EA002C"   #SK 고유색
def lighten(hex_color, amount: float) -> str:
    # amount: 0(원색) ~ 1(완전 흰색)
    c = np.array(mcolors.to_rgb(hex_color))
    return mcolors.to_hex(1 - (1 - c) * (1 - amount))




YEAR_COLORS = {
    2022: lighten(BASE_2024, 0.70),  # 가장 옅게
    2023: lighten(BASE_2024, 0.35),  # 중간
    2024: BASE_2024,                 # #EA002C 그대로
}




COLOR_POS = "#3B82F6"   # 강점
COLOR_NEG = "#EF4444"   # 약점
COLOR_NEU = "#10B981"   # 평균권




ICONS = {
    "upload":    "icons/upload.png",
    "search":    "icons/search.png",
    "zip":       "icons/zip.png",
    "download":  "icons/download.png",
    "dashboard": "icons/dashboard.png",
    "list":      "icons/list.png",
}
def _icon(path, w=60):
    if os.path.exists(path):
        st.image(path, width=w)
        
# 카드 느낌을 위한 얇은 테두리 + 여백 (Streamlit 1.30+면 border=True 가능)
def card(title, icon_key=None):
    return st.container(border=True) if "container" in dir(st) else st.container()
















# ==== 0. 문항-카테고리 매핑 읽기 =================================
QUEST_DF   = pd.read_csv("객관식 문항 분류 35문항 (7개 카테고리).csv", encoding="utf-8-sig")
CAT2ITEMS  = {cat: QUEST_DF[cat].dropna().tolist() for cat in QUEST_DF.columns}
ITEM2CAT   = {item: cat for cat, items in CAT2ITEMS.items() for item in items}












# ====== 유튜브 교육 DB 로드 함수 추가 ======
@st.cache_data(ttl=3600, show_spinner=False)
def load_youtube_db(youtube_file: str = "유튜브교육추천정리.xlsx") -> pd.DataFrame:
    try:
        df = pd.read_excel(youtube_file)
    except Exception:
        # 최소 스키마 반환
        return pd.DataFrame(columns=["영상명", "채널명", "객관식분류", "요약"])


    # 컬럼명 정규화
    colmap = {}
    for c in df.columns:
        k = str(c).strip()
        if k in ["영상명", "영상 제목", "동영상명", "Title", "제목"]:
            colmap[c] = "영상명"
        elif k in ["채널명", "채널", "Channel", "채널 이름"]:
            colmap[c] = "채널명"
        elif k in ["객관식분류", "카테고리", "영역", "분류"]:
            colmap[c] = "객관식분류"
        elif k in ["요약", "설명", "소개", "요약설명", "Description"]:
            colmap[c] = "요약"
    df = df.rename(columns=colmap)


    # 필수 컬럼 보정
    for need in ["영상명", "객관식분류", "요약"]:
        if need not in df.columns:
            df[need] = ""
    if "채널명" not in df.columns:
        df["채널명"] = ""


    # 빈 제목 제거 + 필요한 컬럼만
    df = df[df["영상명"].astype(str).str.strip() != ""].copy()
    return df[["영상명", "채널명", "객관식분류", "요약"]]




# ====== (신규) 채널명 주입 유틸 ======
def _annotate_channels(yt_text: str, youtube_df: pd.DataFrame) -> str:
    """LLM 결과의 '유튜브 영상 ①: 제목' 라인에 채널명이 없으면 붙인다."""
    title2ch = {
        str(t).strip().lower(): str(ch).strip()
        for t, ch in zip(youtube_df["영상명"].astype(str), youtube_df["채널명"].astype(str))
        if str(t).strip()
    }
    # '유튜브 영상 ①: ...' / '유튜브 영상 ②: ...' 모두 잡기
    pat = re.compile(r'^(?P<prefix>\s*유튜브\s*영상\s*[①②]?\s*[:：]\s*)(?P<title>[^\n]+)$', re.MULTILINE)
    def repl(m):
        prefix, title = m.group("prefix"), m.group("title").strip()
        # 이미 채널명이 들어가 있으면 그대로 둠
        if "—" in title or "-" in title or "(" in title:
            return m.group(0)
        ch = title2ch.get(title.lower())
        return f"{prefix}{title} — {ch}" if ch else m.group(0)
    return pat.sub(repl, yt_text)


# ====== 유튜브 프로그램 추천 함수 추가 ======
@st.cache_data(
    ttl=3600, show_spinner=False,
    hash_funcs={pd.DataFrame: lambda df: hashlib.md5(df.to_csv(index=False).encode()).hexdigest()}
)
def recommend_youtube(weak_text: str, youtube_df: pd.DataFrame) -> str:
    # 목록 자체에 채널명 포함
    yt_summary = youtube_df[["영상명", "채널명", "객관식분류", "요약"]].to_string(index=False)


    prompt = f"""다음은 팀장의 약점입니다:
{weak_text}


아래 유튜브 목록에서만 고르세요(목록 밖 금지).
각 항목은 '영상명 — 채널명'을 기준으로 식별합니다:
{yt_summary}


가장 적절한 영상 2개를 정확히 아래 형식으로만:
유튜브 영상 ①: <영상명> — <채널명>
이유:
유튜브 영상 ②: <영상명> — <채널명> (없으면 '적합한 추가 영상 없음')
이유:
"""
    messages = [
        {"role": "system", "content": "당신은 유튜브 교육 콘텐츠 큐레이터입니다. 반드시 채널명을 함께 표기합니다."},
        {"role": "user",   "content": prompt}
    ]
    out = chat_ax4(messages)


    # 혹시 모델이 채널명을 빠뜨리면 DB로 보강
    out = _annotate_channels(out, youtube_df)
    return out












# ➊ STEP 1 전용 마크다운 생성 (5원칙 요약 반영)
@st.cache_data(ttl=3600, show_spinner=False)
def generate_step1_md(
    name: str,
    strengths: str,
    weaknesses: str,
    score_summary: str,
    weak_text: str,
    subjective_weak: str | None = None   # ★ 추가
) -> str:
    """
    STEP 1 · Individual Action용 마크다운을 LLM으로 생성합니다.
    - name: 팀장 이름 (예: "홍길동")
    - strengths: 강점 요약 텍스트
    - weaknesses: 약점 요약 텍스트
    - score_summary: 객관식 점수 요약(여러 줄 가능)
    - weak_text: 최저 점수 영역(예: "상호배려" 또는 "가장 낮은 영역: 상호배려")
    """
    _name         = (name or "팀장").strip()
    _strengths    = (strengths or "-").strip()
    _weaknesses   = (weaknesses or "-").strip()
    _score_summary= (score_summary or "-").strip()
    _weak_text    = (weak_text or "선택된 최저 영역").strip()   # ✅ 널가드
    subweak_block = f"\n[주관식 약점 요약]\n{subjective_weak.strip()}\n" if subjective_weak else ""
   
    # ——— 5 Key Principles: 압축 요약본 (프롬프트 내 포함) ———
    principles_condensed = """
[5 Key Principles · 요약]
공통효과: 상대의 personal needs 충족 → 심리적 안전·신뢰 형성 → 더 높은 수준의 소통·협업 가능.




1) 존중(Respect)의 원칙
- 사실 기반 인정과 신뢰 표현으로 자존감 유지·고취.
- 실수는 ‘의도’와 ‘결과’를 분리, 사람 아닌 ‘사실’에 초점.
- Key Actions: 좋은 생각·의도·동기·성과를 구체적으로 인정 / 신뢰 표현 / 구체적인 칭찬과 피드백 제공 / 낙인 금지 / 개선점은 사실로 말하기.




2) 공감(Empathy)의 원칙
- 비판 없는 경청으로 사실과 감정을 모두 이해, 긴장 완화 및 신뢰 강화.
- 동의와 별개로 “그럴 수 있겠다”라는 이해를 표현.
- Key Actions: 사실(Fact)과 느낌(Feeling) 함께 언급 / 상대방 말을 요약하고 반영하며 듣기 / 긍정·부정 감정 모두 다루기.




3) 참여(Participation)의 원칙
- 겸손과 호기심을 기반으로 한 개방형 질문, 발언 기회 제공.
- 존재감·효능감이 향상되면 몰입·창의·실행의지가 향상.
- Key Actions: 도움 요청 / '무엇·어떤'으로 시작하는 개방형 질문 사용 / 참여를 우선 고려.




4) 공유(Sharing)의 원칙
- 생각·감정·이유(Rationale) 투명 공유로 신뢰 형성·오해 예방.
- 과도한 사적 공유는 지양하되 진정성 유지.
- Key Actions: 생각·감정·근거를 명확히 말하기 / 맥락 제공 / 과장 없는 진정성 제공.




5) 지원(Support)의 원칙
- 책임은 당사자에게 유지, 필요한 도움은 구체화.
- ‘대신 해주기’ 함정 회피, 지원 범위·수준 합의.
- Key Actions: 대신하지 않기 / 스스로 하도록 돕기 / 지원 내용·수준을 명확히 약속·이행.
""".strip()








    prompt = f"""
{_name}의 3개년 리더십 서베이 요약 데이터입니다.








팀장의 최저 영역(개선 우선순위): {_weak_text}




- **강점**: {_strengths}
- **약점**: {_weaknesses}
- **객관식 점수 요약**:
{_score_summary}
















{principles_condensed}
principle_list = "존중(Respect)의 원칙, 공감(Empathy)의 원칙, 참여(Participation)의 원칙, 공유(Sharing)의 원칙, 지원(Support)의 원칙"




<RULES>
→ 아래 '목표'·'핵심포인트'·'활동 예시'·'기대효과'는 반드시 {_weak_text} 개선을 중심으로 작성합니다.
→ '핵심포인트' 2개는 반드시 위 '주관식 약점 요약'의 맥락·키워드를 반영합니다.
→ 목표: 최종 상태를 한 문장으로 제시하고 “~을/를 목표로 합니다.”로 마무리합니다.
→ 핵심포인트: 각 줄은 '제목: 설명' 형식입니다.
    - 제목: {_weak_text}와 관련된 아래 `principle_list` 명시. 반드시 아래 `principle_list` 표기를 그대로 사용하며, 괄호 안 영문까지 포함.
    - 설명: 제목에 대한 설명을 100자 이내 완성형 문장 1개로 작성.
    - 반드시 설명 끝에 관련 핵심 키워드 3개를 해시태그(#)로 표시. [잘못된 예] …입니다. ← 해시태그 없음(무효) [올바른 예] …입니다. #주도성 #경청 #성장 ← 해시태그 3개(유효)
    - 같은 원칙 반복 시 하위 초점·동사·키워드 변형.




→ 활동 예시: 팀장 개인의 ‘말’과 ‘행동’을 한 줄에 함께 제시합니다.
    - 큰따옴표(" ") 안의 발화 예시는 반드시 완결된 1문장만 작성.
    - 예시는 30자 이내의 짧은 문장. 불필요한 부연 설명 금지.
    - 팀/조직 차원의 제도·프로세스·회의·가이드라인 등은 금지.
       
팀장 개인의 ‘말’과 ‘행동’을 한 줄에 함께 제시합니다. 팀/조직 차원의 제도·프로세스·회의·가이드라인 등은 금지합니다.
→ 기대효과: {_weak_text}에 대한 팀원·팀의 관찰 가능한 평가/행동 변화만 기술하고 ‘리더십’ 단어는 사용하지 않습니다.
→ 제목이 아닌 모든 문장은 '합니다/됩니다/입니다' 종결형만 사용합니다('한다/해요/해주세요' 금지).
→ 각 줄은 150자 이내로 완성된 문장으로 작성합니다.
→ '목표'에 대한 설명은 마크다운 불릿 `- `로만 작성합니다.
→ 괄호 안은 반드시 다음 중 하나를 그대로 채웁니다: 존중, 공감, 참여, 공유, 지원; 빈 괄호 금지.

→ '핵심포인트'는 정확히 2개만 작성합니다. 3개 이상 금지.
→ '활동 예시'는 정확히 2개만 작성합니다. 초과 금지.
→ '기대효과'는 정확히 1개만 작성합니다. 초과 금지.
→ 아래 <FORMAT> 블록만 출력하고 그 외 텍스트(제목/추가 불릿/빈 줄) 금지.

</RULES>




[KEYWORDS 사용 규칙]
  - 선택된 {_weak_text}의 KEYWORDS만 참고합니다.




<KEYWORDS>
    자긍심: 자부심, 성과, 소속감
    공동체의식: 동등, 협력, 존중, 다양성, 서로 돕기
    상호배려: 경청하는 분위기, 배려, 갈등 조정, 안정감
    내 일 알기: 명확한 역할, 우선순위, 중요성 인식, 목표 공유, 연결고리 제시, 주도성, 책임감
    도전적 목표설정: 도전적 목표 제시, 방향과 자원 제공, 현실성과 실행력 고려, 안정적 성과, 신뢰, 혁신
    철저하고 즐거운 실행: 일방적 목표 지양, 긍정 에너지, 작은 성과 인정, 소통, 적극 지원
    지식 공유 및 역량 개발: 정보 공유, 배우는 문화, 기회 제공, 업무수행 및 성장
</KEYWORDS>




<FORMAT 규칙>
- 제목줄(핵심포인트/활동 예시/기대효과)은 굵게, 앞뒤 빈줄 1줄씩.
- 각 줄 150자 이내, 완성형 문장.
- 본문 문장과 콜론 뒤 설명은 모두 경어체.
- 제목 어구를 본문 첫 줄에 붙여 쓰지 말 것(예: "… 핵심포인트" 금지).



<FORMAT> 안의 줄만 채워서 그대로 반환:
<FORMAT>
#### 🌱즉각적인 개인 실천의 시작



### **핵심포인트**
- **{{핵심포인트 1 제목}}**: {{ (팀장의 최저 점수 영역: {_weak_text})와 연관된 5 Key Principles와 하위 내용을 1개 설명합니다. }}
- **{{핵심포인트 2 제목}}**: {{ (팀장의 최저 점수 영역: {_weak_text})와 연관된 5 Key Principles와 하위 내용을 1개 설명합니다. }}



### **활동 예시**
- **{{활동명 1}}**: {{핵심포인트 1 설명의 5 Key Principles를 기반으로 실천 가능한 개인의 말과 행동을 상세히 작성합니다.}}
- **{{활동명 2}}**: {{핵심포인트 2 설명의 5 Key Principles를 기반으로 실천 가능한 개인의 말과 행동을 상세히 작성합니다.}}





### **기대효과**
- **{{기대효과 1 제목}}**: {{활동명 1 실행 시 나타날 {_weak_text} 관련 긍정적인 팀의 변화를 작성합니다.}}


</FORMAT>
"""
    messages = [
        {"role": "system", "content": "당신은 조직심리·리더십 코치입니다."},
        {"role": "user",   "content": prompt}
    ]
    raw = chat_ax4(messages)
    return _normalize_step2(raw)




# STEP 2 전용 마크다운 생성 함수 추가
@st.cache_data(ttl=3600, show_spinner=False)
def generate_step2_md(trend_str: str, weak_text: str, recos: str = "", subjective_weak: str | None = None) -> str:
    subweak_block = f"\n[주관식 약점 요약]\n{subjective_weak.strip()}\n" if subjective_weak else ""
    prompt = f"""3개년 리더십 점수 추이입니다.
{trend_str}








팀장의 최저 영역(개선 우선순위): {weak_text}
{subweak_block}




<RULES>
→ 아래의 '목표'·'핵심포인트'와 '기대효과'는 반드시 {weak_text} 개선과 영역별 중요 KEYWORDS를 중심으로 작성하세요.
→ 목표: 최종 상태를 한 문장으로 제시하고 “~을/를 목표로 합니다.”로 마무리합니다.
→ '핵심포인트' 2개는 반드시 위 '주관식 약점 요약'의 맥락·키워드를 반영합니다.
→ 핵심포인트: 각 줄은 '제목: 설명' 형식으로 작성합니다.
   - 제목: {weak_text}를 기반으로 추천하는 '교육 유형’만 명시.
   - 설명: 제목에 대한 설명을 100자 이내 완성형 문장 1개로 작성.
   - 반드시 설명 끝에 관련 핵심 키워드 3개를 해시태그(#)로 표시. [잘못된 예] …입니다. ← 해시태그 없음(무효) [올바른 예] …입니다. #주도성 #경청 #성장 ← 해시태그 3개(유효)
→ 활동 예시: {weak_text}에 직접 관련된 교육만 제시합니다(추천된 사내 교육/유튜브 참고).
→ 기대효과: 위 교육 이수 시 {weak_text} 관련 관찰 가능한 개선점을 기술하고 ‘리더십’ 단어는 사용하지 않습니다.
→ 제목이 아닌 모든 문장은 '합니다/됩니다/입니다' 종결형 어미만 사용합니다('한다/해요/해주세요' 금지).
→ 각 줄은 150자 이내로 완성된 문장으로 작성합니다.
→ '목표'에 대한 설명은 마크다운 불릿 `- `로만 작성합니다.

→ '핵심포인트'는 정확히 2개만 작성합니다. 3개 이상 금지.
→ '활동 예시'는 정확히 3개만 작성합니다. 초과 금지.
→ '기대효과'는 정확히 1개만 작성합니다. 초과 금지.
→ 아래 <FORMAT> 블록만 출력하고 그 외 텍스트(제목/추가 불릿/빈 줄) 금지.
</RULES>




[KEYWORDS 사용 규칙]
  - 선택된 {weak_text}의 'KEYWORDS'만 참고하세요.




<KEYWORDS>
    자긍심: 자부심, 성과, 소속감
    공동체의식: 동등,협력,존중,다양성, 서로 돕기
    상호배려: 경청하는 분위기, 배려, 갈등 조정, 안정감
    내 일 알기: 명확한 역할, 우선순위, 중요성 인식, 목표 고유, 연결고리 제시, 주도성, 책임감
    도전적 목표설정: 도전적 목표 제시, 방향과 자원 제공, 현실성과 실행력 모두 고려, 안정적인 성과 중시, 신뢰, 혁신
    철저하고 즐거운 실행: 일방적이지 않은 목표 설정, 긍정 에너지,작은 성과도 인정,소통, 적극 지원
    지식 공유 및 역량 개발: 정보 공유, 배우는 문화, 기회 제공, 업무수행 및 성장
</KEYWORDS>




<FORMAT 규칙>
- 제목줄(핵심포인트/활동 예시/기대효과)은 굵게, 앞뒤 빈줄 1줄씩.
- 각 줄 150자 이내, 완성형 문장.
- 본문 문장과 콜론 뒤 설명은 모두 경어체.
- 제목 어구를 본문 첫 줄에 붙여 쓰지 말 것(예: "… 핵심포인트" 금지).








{recos}








<FORMAT> 안의 줄만 채워서 반환:
<FORMAT>
#### 🧭체계적인 교육을 통한 역량 강화



### **핵심포인트**
- **{{교육 유형 1 제목}}**: {{팀장의 {weak_text}를 기반으로 필요한 교육 유형을 추천합니다.}}
- **{{교육 유형 2 제목}}**: {{팀장의 {weak_text}를 기반으로 필요한 교육 유형을 추천합니다.}}

### **활동 예시**
- **mySUNI 교육 콘텐츠 ①**: {{카드명}} – {{카드소개내용}}
- **mySUNI 교육 콘텐츠 ②**: {{카드명}} – {{카드소개내용}}
- **유튜브 영상 ①**: {{영상명}}/{{채널명}} – {{요약}}

### **기대효과**
- **{{기대효과 1 제목}}**: {{mySUNI 교육 콘텐츠 ① 수강 시 나타날 기대 변화 1을 설명합니다.}}


</FORMAT>
"""
    messages = [
        {"role": "system", "content": "당신은 조직심리·리더십 전문가입니다."},
        {"role": "user",   "content": prompt}
    ]
    return chat_ax4(messages)










# STEP 3 전용 마크다운 생성 함수 추가 (STEP2 함수 바로 아래에 위치)
@st.cache_data(ttl=3600, show_spinner=False)
def generate_step3_md(weak_text: str, org_scope: str | None = None, subjective_weak: str | None = None) -> str:
    scope_hint = f"현재 조직 범위: {org_scope}" if org_scope else "현재 조직 범위 정보 없음"
    subweak_block = f"\n[주관식 약점 요약]\n{subjective_weak.strip()}\n" if subjective_weak else ""
    prompt = f"""다음은 팀/조직 실행 단계(STEP 3)입니다.




팀장의 최저 영역(개선 우선순위): {weak_text}
{scope_hint}
<RULES>
→ 아래 '목표'·'핵심포인트'·'활동 예시'·'기대효과'는 반드시 {weak_text} 개선과 영역별 주요 KEYWORDS를 중심으로 작성하세요.
→ 목표: 최종 상태를 한 문장으로 제시하고 “~을/를 목표로 합니다.”로 마무리합니다.
→ '핵심포인트'는 반드시 '주관식 약점 요약'의 구체적인 맥락을 반영합니다.
→ 핵심포인트: 각 줄은 '제목: 설명' 형식으로 작성합니다.
   - 제목: {weak_text}를 기반으로 팀/조직이 지향해야 할 문화·분위기 방향성을 간결한 명사형으로 작성.
   - 설명: 제목에 대한 설명을 100자 이내 완성형 문장 1개로 작성.
   - 반드시 설명 끝에 관련 핵심 키워드 3개를 해시태그(#)로 표시. [잘못된 예] …입니다. ← 해시태그 없음(무효) [올바른 예] …입니다. #주도성 #경청 #성장 ← 해시태그 3개(유효)
→ 활동 예시: 개인 행동이 아닌 팀/조직 프로세스·제도·활동만 제시하고, 각 항목에 구체적인 내용을 포함합니다. 'SK 캔미팅 진행 시' 문구를 최소 1회 포함해 워크숍 아이디어를 제시합니다.
→ 기대효과: 실행 시 {weak_text} 관련 팀의 행동/문화/분위기 변화를 기술하며 “리더십” 단어는 사용하지 않습니다.
→ 제목이 아닌 모든 문장은 '합니다/됩니다/입니다' 종결형 어미만 사용합니다('한다/해요/해주세요' 금지).
→ 각 줄은 150자 이내로 완성된 문장으로 작성합니다.
→ '목표'에 대한 설명은 마크다운 불릿 `- `로만 작성합니다.
→ '핵심포인트'는 정확히 2개만 작성합니다. 3개 이상 금지.
→ '활동 예시'는 정확히 2개만 작성합니다. 초과 금지.
→ '기대효과'는 정확히 1개만 작성합니다. 초과 금지.
→ 아래 <FORMAT> 블록만 출력하고 그 외 텍스트(제목/추가 불릿/빈 줄) 금지.
</RULES>




[KEYWORDS 사용 규칙]
  - 선택된 {weak_text}의 KEYWORDS만 참고하세요.




<KEYWORDS>
    자긍심: 자부심, 성과, 소속감
    공동체의식: 동등,협력,존중,다양성, 서로 돕기
    상호배려: 경청하는 분위기, 배려, 갈등 조정, 안정감
    내 일 알기: 명확한 역할, 우선순위, 중요성 인식, 목표 공유, 연결고리 제시, 주도성, 책임감
    도전적 목표설정: 도전적 목표 제시, 방향과 자원 제공, 현실성과 실행력 고려, 안정적 성과, 신뢰, 혁신
    철저하고 즐거운 실행: 일방적 목표 지양, 긍정 에너지, 작은 성과 인정, 소통, 적극 지원
    지식 공유 및 역량 개발: 정보 공유, 학습 문화, 기회 제공, 업무수행 및 성장
</KEYWORDS>




<FORMAT 규칙>
- 제목줄(핵심포인트/활동 예시/기대효과)은 굵게, 앞뒤 빈줄 1줄씩.
- 각 줄 150자 이내, 완성형 문장.
- 본문 문장과 콜론 뒤 설명은 모두 경어체.
- 제목 어구를 본문 첫 줄에 붙여 쓰지 말 것(예: "… 핵심포인트" 금지).




<FORMAT> 안의 줄만 채워서 반환:
<FORMAT>
#### 🪄팀/조직 차원의 변화 실천



### **핵심포인트**
- **{{핵심포인트 1 제목}}**: {{팀과 조직이 지향해야 하는 방향성을 설명합니다.}}
- **{{핵심포인트 2 제목}}**: {{팀과 조직이 지향해야 하는 방향성을 설명합니다.}}



### **활동 예시**
- **{{프로세스/활동/제도 1 이름}}**: {{활동 내용을 추천합니다. (SK 캔미팅 진행 시 …)}}
- **{{프로세스/활동/제도 2 이름}}**: {{활동 내용을 추천합니다.}}


### **기대효과**
- **{{기대효과 제목 1}}**: {{팀장의 {weak_text} 관련 팀의 행동/분위기/문화 변화를 구체화합니다.}}

</FORMAT>
"""
    messages = [
        {"role": "system", "content": "당신은 조직 실행·성과관리 전문가입니다."},
        {"role": "user",   "content": prompt}
    ]
    raw = chat_ax4(messages)
    return _normalize_step2(raw)  # STEP2와 동일한 정규화 사용










def load_and_prepare(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    try:
        # 우선 엑셀로 시도
        df = pd.read_excel(uploaded_file)
    except Exception:
        # 실패 시 CSV로 읽기
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file)
        
    # ➊ 팀장 보임일(날짜) → 연도 추출
    df["보임연도"] = pd.to_datetime(df["팀장 보임일"], errors="coerce").dt.year




    # ➋ 연도 컬럼이 숫자가 아니면 변환 (예: 문자열 ‘2023’)
    df["연도"] = pd.to_numeric(df["연도"], errors="coerce")




    # ➌ 보임연도 > 연도 인 행 제거
    df = df[df["보임연도"] <= df["연도"]].copy()




    # ── 여기부터 수정 ─────────────────────────
    org_cols = [c for c in ["회사명","본부","실","팀"] if c in df.columns]
    pos_cols = [c for c in ["직위","직책","직급"] if c in df.columns]  
    needed   = [LEADER_COL, "팀장 ID", "평가자 수"] + VISUAL_COLS + org_cols + pos_cols  




    missing_required = [LEADER_COL, "팀장 ID"] + VISUAL_COLS
    missing = [c for c in missing_required if c not in df.columns]
    if missing:
        raise ValueError(f"엑셀에 다음 필수 컬럼이 없습니다: {missing}")




    df7 = df[needed].copy()




    def to_float(x):
        if isinstance(x, list):
            vals = []
            for v in x:
                try: vals.append(float(v))
                except: pass
            return np.mean(vals) if vals else np.nan
        return pd.to_numeric(x, errors="coerce")
















    for col in VISUAL_COLS:
        df7[col] = df7[col].apply(to_float)
















    return df7


def draw_chip(c, x, y, text, bg="#EA002C", fg="#FFFFFF", pad_x=12, pad_y=6, r=14, font="KOR_FONT_BOLD", size=12):
    c.saveState()
    c.setFillColor(HexColor(bg)); c.setStrokeColor(HexColor(bg))
    w = stringWidth(text, font, size) + pad_x*2
    h = size + pad_y*2
    c.roundRect(x, y-h, w, h, r, fill=1, stroke=0)
    c.setFillColor(HexColor(fg)); c.setFont(font, size)
    c.drawString(x+pad_x, y-h+pad_y-1, text)
    c.restoreState()
    return w, h  # 폭/높이 반환

def draw_card(c, x, y, w, h, r=18):
    return

def draw_section_pill(c, x, y, text):
    # 빨간 둥근 라벨: "3개년 리더십 추이", "종합 다이어그램", "강점 & 약점"
    return draw_chip(c, x, y, text, bg="#EA002C", fg="#FFFFFF", pad_x=16, pad_y=6, r=18, size=13)

def _normalize_step2(md: str) -> str:
    md = re.sub(r'</?FORMAT>', '', md, flags=re.I)




    # 헤딩 마커(#...)를 유지하고, 앞뒤로 빈 줄만 보장
    md = re.sub(r'(?m)^\s*(#{1,6}\s*.+?)\s*$', r'\n\1\n', md)




    # 혹시 LLM이 제목을 헤딩 없이 보냈을 때 보정 (선택)
    md = re.sub(r'(?m)^\s*(핵심 ?포인트|활동 ?예시|기대효과)\s*$', r'\n### \1\n', md)




    md = re.sub(r'\n{3,}', '\n\n', md)
    return md.strip()




@st.cache_data(ttl=3600, show_spinner=False)
def build_step2_with_recos(trend_str, weak_text, edu_file, subjective_weak: str | None = None):
    edu_df = load_edu_db(edu_file)
    yt_df  = load_youtube_db()
    prog = recommend_programs(weak_text, edu_df) if not edu_df.empty else "사내 교육 DB 없음"
    yt   = recommend_youtube(weak_text, yt_df)   if not yt_df.empty  else "유튜브 DB 없음"
    recos_txt = f"[사내 교육 추천]\n{prog}\n\n[유튜브 추천]\n{yt}"
    md = generate_step2_md(trend_str, weak_text, recos_txt, subjective_weak=subjective_weak)  # ★
    return _normalize_step2(md)
    
def _format_sw_text(avg_scores: pd.Series) -> tuple[str, str]:
    top2, bottom2 = pick_strengths_weaknesses(avg_scores, n=2)
    strengths = ", ".join(LABEL_MAP[c] for c in top2.index)
    weaknesses = ", ".join(LABEL_MAP[c] for c in bottom2.index)
    return strengths, weaknesses    
    
def _make_score_summary(avg_scores: pd.Series, overall_mean: pd.Series | None = None) -> str:
    lines = []
    for col in VISUAL_COLS:
        v = float(avg_scores.get(col, np.nan))
        if np.isnan(v): 
            continue
        if overall_mean is not None and col in overall_mean:
            lines.append(f"- {LABEL_MAP[col]}: {v:.2f} (전사 {float(overall_mean[col]):.2f})")
        else:
            lines.append(f"- {LABEL_MAP[col]}: {v:.2f}")
    return "\n".join(lines)
    
def draw_column_divider(c, x, top, bottom, color="#D1D5DB", width=0.8, dash=None):
    """
    c: reportlab canvas
    x: 세로선 x좌표
    top, bottom: 선의 위/아래 y좌표
    color: HEX 문자열
    width: 선 두께(pt)
    dash: (on, off) 점선 패턴. 예: (3, 2)
    """
    c.saveState()
    c.setStrokeColor(HexColor(color))
    c.setLineWidth(width)
    if dash:
        c.setDash(dash[0], dash[1])
    c.line(x, bottom, x, top)
    c.restoreState()



# 이모지/VS16 제거 + 이모지 헤더 라인 정리
EMOJI_RE = re.compile(r'[\U0001F300-\U0001FAFF\u2600-\u27BF\u200d\uFE0F]')

def md_for_pdf(md: str) -> str:
    # 0) 최상단 헤더(🌱/🧭/🪄 포함 여부와 무관)를 한 번만 제거 → 상자 제목과 중복 방지
    md = re.sub(r'^\s*#{1,6}\s*(?:[🌱🧭🪄]\s*)?.*\n+', '', md, count=1)

    # 1) 남아있는 다른 헤더들에서 이모지만 제거(형식은 유지)
    md = re.sub(r'(?m)^\s*#{1,6}\s*[🌱🧭🪄]\s*', '### ', md)

    # 2) 본문 내 이모지 제거
    md = EMOJI_RE.sub('', md)
    return md.strip()



def _img_data_uri(path: str) -> str:
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("utf-8")




# ── 업로드 카드 컴포넌트 ──
def upload_card(title: str, subtitle: str, icon_path: str, key: str, types: list[str],
                right_icon_path: str | None = None, click_to_browse: bool = True):
    left_src  = _img_data_uri(icon_path)
    right_src = _img_data_uri(right_icon_path) if right_icon_path else None


    # 카드 헤더 (오른쪽 아이콘을 <img>로)
    st.markdown(
        f"""
        <div id="wrap-{key}" class="upload-card">
          <div class="uc-head">
            <div class="uc-left">
              <img src="{left_src}" class="uc-icon" />
              <div>
                <div class="uc-title">{title}</div>
                <div class="uc-sub">{subtitle}</div>
              </div>
            </div>
            <div class="uc-right">
              {f'<img id="up-{key}" src="{right_src}" class="uc-right-icon" />' if right_src else '⬆️'}
            </div>
          </div>
        """,
        unsafe_allow_html=True
    )


    # 업로더 본체
    file_obj = st.file_uploader("파일 업로드", type=types, key=key, label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)  # upload-card 닫기


    # 오른쪽 아이콘 클릭 → 업로더 열기
    if right_src and click_to_browse:
        st.markdown(
            f"""
            <script>
            const wrap = document.getElementById("wrap-{key}");
            if (wrap) {{
              const icon = document.getElementById("up-{key}");
              const btn  = wrap.querySelector('[data-testid="stFileUploader"] button');
              if (icon && btn) {{
                icon.style.cursor = "pointer";
                icon.addEventListener("click", () => btn.click());
              }}
            }}
            </script>
            """,
            unsafe_allow_html=True
        )
    return file_obj


def draw_pill(c, x, y, text, fill=None):
    if fill is None:
        fill = PRIMARY
    c.saveState()
    c.setFillColor(HexColor(fill)); c.setStrokeColor(HexColor(fill))
    w = stringWidth(text, "KOR_FONT_BOLD", 12) + 26
    h = 20
    c.roundRect(x, y-h, w, h, 10, stroke=0, fill=1)
    c.setFillColor(HexColor("#FFFFFF"))  # ← reportlab에 안전하게 HexColor로
    c.setFont("KOR_FONT_BOLD", 12)
    c.drawString(x+13, y-14, text)
    c.restoreState()

def draw_kpis(
    c,
    y_top: float,
    val_100: float,
    label: str,
    with_cards: bool = False,
    card_h: int = 64,
    left_x: float | None = None,
    right_x: float | None = None,
    left_w: float | None = None,
    right_w: float | None = None,
):
    """
    y_top: KPI 타이틀(왼/오) 기준 Y
    with_cards: 카드 배경 표시 여부
    card_h: 카드 높이
    left_x/right_x/left_w/right_w: 좌표/폭 (미지정 시 A4 기준으로 자동 계산)
    """
    # 좌우 좌표/폭 자동 계산 (A4, 여백 36pt, 좌열폭 270, 간격 18)
    if None in (left_x, right_x, left_w, right_w):
        W, H = A4
        M = 36
        gap = 18
        left_x  = M if left_x  is None else left_x
        left_w  = 270 if left_w is None else left_w
        right_x = (left_x + left_w + gap) if right_x is None else right_x
        right_w = (W - right_x - M)       if right_w is None else right_w

    # 카드 배경
    if with_cards:
        draw_card(c, left_x,  y_top, left_w,  card_h, r=14)
        draw_card(c, right_x, y_top, right_w, card_h, r=14)
        y_left_title  = y_top - 18
        y_left_value  = y_top - 44
        y_right_title = y_top - 18
        y_right_value = y_top - 36
        y_right_sub   = y_top - 54
    else:
        y_left_title  = y_top + 10
        y_left_value  = y_top + 10
        y_right_title = y_top + 10
        y_right_value = y_top + 10
        y_right_sub   = y_top - 5

    # streamlit run ax4_final.py
    # 왼쪽 KPI
    c.setFont("KOR_FONT_BOLD", 10); c.setFillColor(HexColor("#111827"))              # ← (1) 왼쪽 타이틀 글씨 크기
    c.drawString(left_x+14, y_left_title, "2024년 리더십 평균 점수 (100점 기준)")       # ← (2) X·Y 위치
    c.setFont("KOR_FONT_BOLD", 10); c.setFillColor(HexColor("#EA002C"))                 # ← (3) 숫자(72.1) 크기
    c.drawRightString(left_x+left_w-14, y_left_value, f"{float(val_100):.1f}/100")          # ← (4) X·Y 위치(오른쪽 정렬)

    # 오른쪽 KPI
    c.setFont("KOR_FONT_BOLD", 10); c.setFillColor(HexColor("#111827"))          # ← (5) '학습 권장 영역' 크기
    c.drawString(right_x+14, y_right_title, "학습 권장 영역")                       # ← (6) 위치
    c.setFont("KOR_FONT_BOLD", 10); c.setFillColor(HexColor("#EA002C"))         # ← (7) 빨간 라벨 글씨 크기
    c.drawRightString(right_x+150, y_right_value, f"{label or ''}")
    c.setFont("KOR_FONT_BOLD", 8); c.setFillColor(HexColor("#374151"))                  # ← (9) 회색 보조문구 크기
    c.drawString(right_x+14, y_right_sub, "단계별 개선방안 1→2→3 단계를 수행해볼까요?")     # ← (10) 위치
    c.setFillColor(HexColor("#000000"))



def draw_step_box(c, x, y, w, h, step_no: int, title: str, md_text: str, header_color: str = "#FFFFFF"):
    # 카드틀
    draw_card(c, x, y, w, h, r=18)

    # 좌측 체크아이콘 위치 계산
    cx, cy = x + 18, y - 22 # 기존 좌표 계산

    # 이미지 크기 설정 (원 반지름 14 → 지름 28)
    icon_size = 20

    # 이미지로 대체
    icon_size = 20 # 원 지름 (반지름 14 * 2)
    c.drawImage(
    "check.png",
    cx - icon_size/2, # 중심 좌표에서 반 나눈 값 빼서 좌측 X
    cy + 10 - icon_size/2, # 중심 Y 맞춤
    width=icon_size,
    height=icon_size,
    mask='auto'
    )
    # STEP 칩 + 제목
    sw, sh = draw_chip(c, x+30, y-5, f"STEP{step_no}", bg="#FEE2E2", fg="#EA002C", pad_x=6, pad_y=3, r=12, size=8)
    c.setFont("KOR_FONT_BOLD", 12.5); c.setFillColor(HexColor("#111827"))
    c.drawString(x+25+sw+10, y-17, title)

    # 본문
    left = x + 16
    right = x + w + 5
    max_w = right - left
    cursor = y - 32
    lines = [ln.rstrip() for ln in md_text.strip().splitlines()]
    for raw in lines:
        line = raw.strip()
        if not line:
            cursor -= 2; continue #제목 앞 간격?
        m = re.match(r'^(#{1,6})\s+(.+)$', line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip('* ')
            size = 9 if level <= 3 else 8
            # 제목 위 간격은 위쪽 if not line에서만 조정
            c.setFont("KOR_FONT_BOLD", size); c.drawString(left, cursor, text)
            cursor -= (7 if level <= 3 else 3); continue # ← 제목 아래 간격
        if re.match(r'^[-*0-9]+\.\s+|^[-*]\s+', line):
            bullet_text = re.sub(r'^[-*0-9]+\.\s+|^[-*]\s+', '• ', line)
            cursor = draw_markdown_line(c, left, cursor, bullet_text, max_width=max_w); cursor -= 1; continue #내용 줄간격
        cursor = draw_markdown_line(c, left, cursor, line, max_width=max_w)











def attach_leader_key(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # 1) ID 우선
    df[LEADER_KEY_COL] = df.get(LEADER_ID_COL, "").astype(str)




    # 2) 빈/누락 ID 보정 → 조직경로 합성키
    if {"회사명","본부","실","팀"}.issubset(df.columns):
        empty = df[LEADER_KEY_COL].isin(["", "nan", "None"])
        df.loc[empty, LEADER_KEY_COL] = (
            df.loc[empty, LEADER_NAME_COL].astype(str).fillna("") + "|" +
            df.loc[empty, "회사명"].astype(str).fillna("") + "|" +
            df.loc[empty, "본부"].astype(str).fillna("")  + "|" +
            df.loc[empty, "실"].astype(str).fillna("")    + "|" +
            df.loc[empty, "팀"].astype(str).fillna("")
        )
    return df


def draw_sw_legend(c, x, y, gap=12, label_pad=4, fs=8):
    # gap: 항목 사이 고정 간격, label_pad: 점과 텍스트 사이 간격
    items = [
        (COLOR_POS, "강점 ≤ 40%"),
        (COLOR_NEU, "평균권 40–60%"),
        (COLOR_NEG, "약점 ≥ 60%"),
    ]
    r = 2.8
    for col, txt in items:
        c.setFillColor(HexColor(col))
        c.circle(x+r, y-r, r, fill=1, stroke=0)   # 점
        c.setFillColor(HexColor("#111827"))
        c.setFont("KOR_FONT", fs)
        c.drawString(x + r*2 + label_pad, y-2*r, txt)
        # 다음 항목 x 이동폭 = 텍스트 폭 + 점+패딩 + gap
        move = stringWidth(txt, "KOR_FONT", fs) + (r*2 + label_pad) + gap
        x += move

# 화면 표시에선 ID/리더키 감추기
def hide_ids(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop(columns=[LEADER_ID_COL, LEADER_KEY_COL], errors="ignore")


# 전체글씨크기조절
# streamlit run ax4_final.py
def draw_markdown_line(
    c, x, y, line, max_width,
    fs=7, fs_bold=8, leading=None, justify=False
):
    """
    fs: 일반 폰트 크기(pt)
    fs_bold: 굵은 폰트 크기(기본 fs와 동일)
    leading: 줄간격(기본 fs+3)
    justify: 좌우 정렬(마지막 줄 제외)을 할지 여부
    """
    fs_bold = fs if fs_bold is None else fs_bold
    leading = (fs + 3) if leading is None else leading

    # 1) '**굵게**' 토큰 분할 → (text, fontName, fontSize)
    parts = re.split(r'(\*\*.+?\*\*)', line)
    segs = []
    for p in parts:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            segs.append((p.strip("*"), "KOR_FONT_BOLD", fs_bold))
        else:
            segs.append((p, "KOR_FONT", fs))

    # 2) 공백 보존 토큰화 (단어/공백 분리)
    def toks(txt):
        # \s+는 공백 묶음, \S+는 공백 아닌 묶음
        return re.findall(r'\S+|\s+', txt)

    tokens = []
    for txt, fn, fz in segs:
        for t in toks(txt):
            tokens.append((t, fn, fz))

    # 3) 줄 나누기(폭 기준). 긴 토큰은 글자 단위로 쪼갬.
    lines = [[]]                  # 각 줄은 [(text, font, size), ...]
    cur_w = 0.0
    for t, fn, fz in tokens:
        tw = stringWidth(t, fn, fz)
        if t == "\n":
            lines.append([]); cur_w = 0.0
            continue

        # 토큰 전체가 안 들어가는데 현재 줄이 비어있지 않으면 줄바꿈
        if cur_w > 0 and cur_w + tw > max_width:
            # 공백이면 버리고 새 줄로
            if t.isspace():
                lines.append([]); cur_w = 0.0
                continue
            # 긴 단어(공백 없음): 글자 단위로 쪼개서 넣기
            for ch in list(t):
                chw = stringWidth(ch, fn, fz)
                if cur_w > 0 and cur_w + chw > max_width:
                    lines.append([]); cur_w = 0.0
                lines[-1].append((ch, fn, fz))
                cur_w += chw
            continue

        # 정상 추가
        lines[-1].append((t, fn, fz))
        cur_w += tw

    # 4) 그리기 (옵션: 좌우정렬)
    for idx, L in enumerate(lines):
        cx = x
        # 좌우정렬: 마지막 줄은 제외
        extra = 0.0
        if justify and idx < len(lines)-1:
            text_w = sum(stringWidth(t, f, s) for t, f, s in L)
            space_cnt = sum(1 for t, _, _ in L if t.isspace())
            if space_cnt and text_w < max_width:
                extra = (max_width - text_w) / space_cnt

        for t, fn, fz in L:
            c.setFont(fn, fz)
            c.drawString(cx, y, t)
            adv = stringWidth(t, fn, fz)
            # 공백이면 여분 간격 배분
            if extra and t.isspace():
                adv += extra
            cx += adv

        y -= leading
    return y








# ===== 주관식 컬럼 상수 =====
SUBJECTIVE_STRENGTH_COL = "주관식 강점"
SUBJECTIVE_WEAK_COL     = "주관식 약점"




def get_subjectives_for_leader(raw_df: pd.DataFrame, leader_key: str) -> tuple[str, str]:
    """
    특정 리더의 주관식 강점/약점 텍스트를 모두 모아 하나의 문자열로 반환.
    (연도별 여러 행이 있어도 합칩니다)
    """
    if raw_df is None or not leader_key or LEADER_KEY_COL not in raw_df.columns:
        return "", ""
    rows = raw_df[raw_df[LEADER_KEY_COL] == leader_key]
    def _collect(col):
        if col not in raw_df.columns: 
            return ""
        return " ".join(rows[col].dropna().astype(str).tolist()).strip()
    return _collect(SUBJECTIVE_STRENGTH_COL), _collect(SUBJECTIVE_WEAK_COL)




@st.cache_data(ttl=3600, show_spinner=False)
def gen_sw_comment_from_subjective(
    category_key: str,       # 예: "팀원_상호배려"
    category_label: str,     # 예: "상호배려"
    sw_type: str,            # "강점" 또는 "약점"
    subj_strength: str,      # 그 리더의 주관식 강점 전체 텍스트
    subj_weak: str           # 그 리더의 주관식 약점 전체 텍스트
) -> list[str]:
    """
    주관식 텍스트를 바탕으로 카테고리 맞춤 코멘트 2문장을 생성합니다.
    주관식이 없거나 실패 시 기본 문구를 쓰지 않고 빈 리스트를 반환합니다.
    """
    ref_text = (subj_strength if sw_type == "강점" else subj_weak or "").strip()
    if not ref_text:
        ref_text = "<<주관식 없음>>"




    prompt = f"""다음은 팀장의 주관식 {sw_type} 응답입니다:
{ref_text}




카테고리: {category_label} (원본 키: {category_key})




규칙:
- '{category_label}' 카테고리에 한정하여 {sw_type} 코멘트 2문장을 제안합니다.
- 각각 한 문장, 120자 이내, 한국어, '합니다/됩니다/입니다' 종결.
- 주관식 텍스트의 단서(상황·행동·빈도·영향)를 반영하고 모호한 일반화는 피합니다.
- 비난·낙인 표현 금지, 관찰 가능한 행동 중심.
- 근거가 약하면 범위를 좁혀 신중하게 기술합니다.




형식(두 줄로만 반환):
- 문장1
- 문장2
"""
    messages = [
        {"role": "system", "content": "당신은 HR 코치입니다. 간결하고 구체적으로 작성합니다."},
        {"role": "user", "content": prompt}
    ]
    try:
        out = chat_ax4(messages)
        lines = [re.sub(r"^\s*[-•]\s*", "", ln).strip() for ln in out.splitlines() if ln.strip()]
        lines = [ln for ln in lines if len(ln) >= 3][:2]
        return lines  # 0~2줄
    except Exception:
        return []    




# ====== 내장 교육 DB 함수 ======
@st.cache_data(ttl=3600, show_spinner=False)
def load_edu_db(edu_file=None):
    if edu_file is not None:
        return pd.read_excel(edu_file)
    return pd.read_excel("리더십 카테고리 정보.xlsx")












def _safe_filename(name: str) -> str:
    """윈도우/맥 공통으로 파일명 안전하게 정리"""
    name = unicodedata.normalize("NFKC", name)
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = name.strip().strip(".")
    return name or "report"




def _make_trend_fig_no_st(sel_df: pd.DataFrame, overall_trend: pd.Series,
                          year_col="연도", score_col="Composite",
                          all_years=(2022, 2023, 2024), y_tick_step=10):
    per_year = sel_df.groupby(year_col)[score_col].mean().reindex(all_years).apply(to_percent)
    corp     = overall_trend.reindex(all_years).apply(to_percent)
    x = np.arange(len(all_years))
    fig, ax = plt.subplots(figsize=(8, 4))
    
    for i, y in enumerate(per_year.values):
        if pd.notna(y):
            yr  = all_years[i]
            col = YEAR_COLORS.get(yr, BASE_2024)
            ax.bar(x[i], y, width=BAR_WIDTH, color=col, edgecolor=col)  # ← 동일
            ax.text(x[i], y + 1.2, f"{y:.0f}", ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.plot(x, corp.values, marker="o", linewidth=2.2,
            color=COLOR_CORP_LINE, label="전사 평균")
    for i, y in enumerate(corp.values):
        if pd.notna(y):
            ax.text(x[i], y+1.2, f"{y:.0f}", ha="center", va="bottom", fontsize=9, color="#333")
    ax.set_xticks(x); ax.set_xticklabels([str(y) for y in all_years])
    ax.set_ylim(0, 100); ax.yaxis.set_major_locator(MultipleLocator(y_tick_step))
    ax.grid(True, axis="y", linestyle="--", alpha=0.4)
    ax.set_xlabel("Years", fontsize=8, loc="right"); ax.legend(loc="upper left")
    plt.tight_layout()
    return fig




def _make_radar_compare_fig_no_st(sel_score: dict, ref_score: dict, title=""):
    order = [LABEL_MAP[c] for c in VISUAL_COLS]




    sel_vals = [to_percent(float(sel_score.get(k, np.nan))) for k in order]
    ref_vals = [to_percent(float(ref_score.get(k, np.nan))) for k in order]
    sel_vals += sel_vals[:1]; ref_vals += ref_vals[:1]




    angles = np.linspace(0, 2*np.pi, len(order), endpoint=False).tolist(); angles += angles[:1]




    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi/2); ax.set_theta_direction(-1)




    ax.set_ylim(0, 100)
    ax.set_yticks([0,20,40,60,80,100])
    ax.set_yticklabels(["0","20","40","60","80","100"], fontsize=10, fontweight="bold", fontfamily="Malgun Gothic")




    ax.set_xticks(angles[:-1]); ax.set_xticklabels(order, fontsize=12, fontweight="bold", fontfamily="Malgun Gothic")
    ax.tick_params(axis='x', pad=20)




    col_ref, col_sel = COLOR_CORP_LINE, COLOR_SELF_LINE
    ax.plot(angles, ref_vals, linewidth=2, marker='o', color=col_ref, alpha=0.9)
    ax.fill(angles, ref_vals, color=col_ref, alpha=0.15)
    ax.plot(angles, sel_vals, linewidth=2.5, marker='o', color=col_sel, alpha=1.0)
    ax.fill(angles, sel_vals, color=col_sel, alpha=0.25)




    ax.xaxis.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.yaxis.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.set_title(title, fontsize=18, pad=20)
    plt.tight_layout()
    return fig




def classify_by_top_band(top_pct: float) -> tuple[str, str]:
    """
    상위 비율(top_pct) 기준으로 (라벨, 색) 반환
    - 40% ≤ top_pct ≤ 60%  → ('평균권', 초록)
    - top_pct < 40%        → ('강점', 파랑)
    - top_pct > 60%        → ('약점', 빨강)
    """
    if 40.0 <= float(top_pct) <= 60.0:
        return "평균권", COLOR_NEU
    return ("강점", COLOR_POS) if float(top_pct) < 40.0 else ("약점", COLOR_NEG)




def to_top_pct(p):
    p = float(p)
    return max(0.0, min(100.0, 100.0 - p))  # 0~100로 클램프




def _gauges_for_leader(avg_scores: pd.Series, overall_mean: pd.Series, overall_series: dict,
                       leader_key: str | None = None, raw_df: pd.DataFrame | None = None):
    figs, metas = [], []
    subj_strength, subj_weak = ("","")
    if leader_key and raw_df is not None:
        subj_strength, subj_weak = get_subjectives_for_leader(raw_df, leader_key)




    top2, bottom2 = pick_strengths_weaknesses(avg_scores)
    targets = list(top2.items()) + list(bottom2.items())




    for col_name, raw_score in targets:
        perc     = to_percent(raw_score)
        overall  = overall_mean[col_name]
        pct_rank = percentile_rank(raw_score, overall_series[col_name])
        top_pct  = to_top_pct(pct_rank)
        tag, color = classify_by_top_band(top_pct)




        base_tag = "강점" if col_name in top2.index else "약점"
        desc_key = "강점" if base_tag == "강점" else "약점"




        lines = gen_sw_comment_from_subjective(
            category_key=col_name,
            category_label=LABEL_MAP[col_name],
            sw_type=desc_key,
            subj_strength=subj_strength,
            subj_weak=subj_weak
        )




        fig = draw_semi_gauge(perc, is_strength=(desc_key=="강점"), color_override=color)
        figs.append(fig)
        metas.append((perc, LABEL_MAP[col_name], pct_rank, tag, color, lines[:2]))
    return figs, metas




@st.cache_data(show_spinner=False)
def make_pdf_for_leader(leader_key: str, survey_df: pd.DataFrame,
                        overall_trend: pd.Series,
                        overall_mean: pd.Series,
                        overall_series: dict,
                        include_llm_text: bool = True,
                        edu_file=None,
                        weak_text: str | None = None,
                        org_scope: str | None = None):

    sel = survey_df[survey_df[LEADER_KEY_COL] == leader_key].copy()
    name = sel[LEADER_NAME_COL].iloc[0] if not sel.empty else leader_key
    sel["Composite"] = sel[VISUAL_COLS].mean(axis=1)
    
    # ▶ NEW: 7개 영역 평균 총점(100점 기준) 계산
    overall_100 = round(to_percent(float(sel["Composite"].mean())), 1)

    # ✅ 여기서 미리 메타 뽑기 (위로 이동)
    leader_meta_sel = extract_leader_meta(sel)
    display_name = (leader_meta_sel or {}).get("name", name)

    # 평균/도표 준비
    team_24, overall_24 = get_scores_for_radar(sel, survey_df, RADAR_COMPARE_YEAR)
    sel_y = sel[sel["연도"] == FOCUS_YEAR]
    if sel_y.empty:     # 안전 폴백
        sel_y = sel
    avg_scores = sel_y[VISUAL_COLS].mean()
    total_score_100 = round(to_percent(float(avg_scores.mean())), 1)

    # ★ 최저 영역 라벨/점수는 여기서 '먼저' 계산 (LLM 호출 전에 필요)
    wl = avg_scores.idxmin()
    weak_label_only = LABEL_MAP[wl]
    weak_score_1to5 = float(avg_scores[wl])

    score_dict = { LABEL_MAP[c]: float(team_24[c]) for c in VISUAL_COLS }
    overall_score_dict = { LABEL_MAP[c]: float(overall_24[c]) for c in VISUAL_COLS }

    trend_fig = _make_trend_fig_no_st(sel, overall_trend)
    radar_fig = _make_radar_compare_fig_no_st(
        score_dict, overall_score_dict, title=f"{RADAR_COMPARE_YEAR}년 기준"
    )

    gauge_figs, gauge_meta = _gauges_for_leader(
        avg_scores, overall_mean, overall_series,
        leader_key=leader_key,
        raw_df=st.session_state.get("raw_df")
    )

    # 3개년 추이 계산
    trend = sel.groupby("연도")["Composite"].mean().sort_index()
    _ = "|".join(f"{int(y)}:{float(s):.4f}" for y, s in trend.reindex([2022, 2023, 2024]).items())

    # ★ LLM 호출 시 최저 영역 라벨을 함께 전달
    trend_comment = make_trend_commentary_via_llm_from_series(trend, weakest_label=weak_label_only).strip()

    strengths_txt, weaknesses_txt = _format_sw_text(avg_scores)
    score_summary_txt = _make_score_summary(avg_scores, overall_mean)

    # 주관식 수집
    subj_strength_txt, subj_weak_txt = get_subjectives_for_leader(st.session_state.get("raw_df"), leader_key)

    # STEP1
    step1_md = generate_step1_md(
        name=display_name,
        strengths=strengths_txt,
        weaknesses=weaknesses_txt,
        score_summary=score_summary_txt,
        weak_text=weak_label_only,
        subjective_weak=subj_weak_txt
    )

    # STEP2/3에 넘길 약점 라벨 확정
    if not weak_text:
        weak_text = weak_label_only

    trend_str = "\n".join(f"{y}년: {s:.2f}점" for y, s in trend.items())

    step2_md = build_step2_with_recos(
        trend_str, weak_text, edu_file,
        subjective_weak=subj_weak_txt
    )
    step3_md = generate_step3_md(
        weak_text=weak_text, org_scope=org_scope,
        subjective_weak=subj_weak_txt
    )

    # ▼ PDF 전용으로 정리
    step1_pdf = md_for_pdf(step1_md)
    step2_pdf = md_for_pdf(step2_md)
    step3_pdf = md_for_pdf(step3_md)

    rf_res   = calc_biggest_rise_fall(sel)
    rf_items = format_rise_fall_items(rf_res)

    pdf_bytes = build_dashboard_pdf(
    leader_name = name,
    radar_fig   = radar_fig,
    gauge_figs  = gauge_figs,
    gauge_meta  = gauge_meta,
    trend_fig   = trend_fig,
    trend_comment = trend_comment,
    rise_fall_items = rf_items,  # ← 여기!
    step1_md    = step1_pdf,
    step2_md    = step2_pdf,
    step3_md    = step3_pdf,
    profile_img = None,
    leader_meta = leader_meta_sel,
    total_score_100 = total_score_100,
    weakest_label = weak_label_only,
    weakest_score_1to5 = weak_score_1to5
    )
    
    plt.close(trend_fig); plt.close(radar_fig)
    for f in gauge_figs: plt.close(f)
    return name, pdf_bytes




def make_zip_for_leaders(leader_keys: list[str],
                         survey_df,
                         overall_trend,
                         overall_mean,
                         overall_series,
                         include_llm_text: bool = True,
                         edu_file=None,
                         org_scope: str | None = None,
                         weakest_label_by_key: dict[str, str] | None = None):
    """여러 리더의 PDF를 zip으로 묶어 반환"""
    mem = BytesIO()
    with ZipFile(mem, "w") as zf:
        for key in leader_keys:
            # 약점 라벨이 있으면 weak_text 구성, 없으면 None (→ 내부 기본 로직 사용)
            wt = None
            if weakest_label_by_key:
                wk = weakest_label_by_key.get(key)
                if wk:
                    wt = wk




            name, pdf = make_pdf_for_leader(
                key, survey_df, overall_trend, overall_mean, overall_series,
                include_llm_text=include_llm_text,
                edu_file=edu_file,
                weak_text=wt,                # ★ 리더별 약점 전달
                org_scope=org_scope          # ★ 조직 범위 전달
            )
            zf.writestr(f"{_safe_filename(name)}_dashboard.pdf", pdf)
    mem.seek(0)
    return mem.getvalue()
# ============================================================================




# 레이더(종합 다이어그램) 비교에 사용할 연도
RADAR_COMPARE_YEAR = 2024




def get_scores_for_radar(sel_df: pd.DataFrame, survey_df: pd.DataFrame, year: int):
    """
    레이더용 평균 점수를 '해당 연도' 기준으로 반환
    - sel_df: 선택 팀장 전체 원본(여러 연도 포함 가능)
    - survey_df: 현재 범위 전체 원본(여러 연도 포함 가능)
    return: (team_mean, overall_mean)  # 둘 다 Series(7개 항목)
    """
    # 팀장: 해당 연도만 평균
    sel_y = sel_df[sel_df["연도"] == year]
    team_mean = sel_y[VISUAL_COLS].mean()




    # 전사: 해당 연도 데이터만 → 팀장별 평균 → 전체 평균
    all_y = survey_df[survey_df["연도"] == year]
    by_leader_y = all_y.groupby(LEADER_KEY_COL)[VISUAL_COLS].mean()
    overall_mean = by_leader_y.mean()




    return team_mean, overall_mean
















def extract_leader_meta(df_for_one_leader: pd.DataFrame) -> dict:
    """선택된 리더의 메타정보를 안전하게 추출"""
    if df_for_one_leader is None or df_for_one_leader.empty:
        return {}




    row = df_for_one_leader.iloc[0]
    def get(*candidates, default=""):
        for c in candidates:
            if c in df_for_one_leader.columns:
                val = row.get(c, "")
                if pd.notna(val) and str(val).strip():
                    return str(val)
        return default




    return {
        "name":     get(LEADER_NAME_COL),
        "id":       get(LEADER_ID_COL),
        "company":  get("회사명"),
        "hq":       get("본부"),
        "dept":     get("실"),
        "team":     get("팀"),
        "position": get("직위", "직책", "직급"),
    }












# ====== STEP 추천 생성 함수 ======
@st.cache_data(ttl=86400, show_spinner=False)
def generate_step_content(step_title: str, objective: str, items: list[str] | None = None) -> str:
    item_block = ""
    if items:
        item_block = "\n\n---\n이 단계와 관련된 세부 객관식 문항:\n" + "\n".join(f"- {q}" for q in items)
    prompt = f"""다음은 리더십 개선 로드맵의 {step_title} 단계입니다.
목표: {objective}{item_block}




다음 3가지를 bullet로 작성:
- 핵심포인트 (3~5개)
- 활동 예시 (2~3개)
- 기대효과 (2~3개)
"""
    messages = [
        {"role": "system","content":"당신은 조직심리 및 리더십 전문가입니다."},
        {"role": "user", "content": prompt}
    ]
    return chat_ax4(messages)


def draw_kpi_boxes_behind(
    c, y_top,
    card_h=64,               # 박스 높이
    top_offset=18,           # 박스 상단을 y_top보다 얼마나 올릴지(+)
    left_x=None, right_x=None, left_w=None, right_w=None,
    r=16, fill="#FFFFFF", stroke="#E5E7EB",
    shadow=True, shadow_offset=(3, -3),
    border_width=1.2
):
    """
    draw_kpis(with_cards=False)로 텍스트를 그대로 그리기 전에
    '배경 카드'만 두 장(좌/우) 깔아준다. draw_card 시그니처는 건드리지 않음.
    """
    # A4 기본 레이아웃(네 코드와 동일 계산)
    if None in (left_x, right_x, left_w, right_w):
        W, H = A4
        M = 36
        gap = 18
        left_x  = M if left_x  is None else left_x
        left_w  = 270 if left_w is None else left_w
        right_x = (left_x + left_w + gap) if right_x is None else right_x
        right_w = (W - right_x - M)       if right_w is None else right_w

    y_card_top = y_top + top_offset  # 카드 상단 Y (텍스트와 맞물림용)

    def _round_box(x, y, w, h):
        # 그림자
        if shadow and shadow_offset:
            dx, dy = shadow_offset
            c.saveState()
            c.setFillColor(HexColor("#F3F4F6"))
            c.setStrokeColor(HexColor("#F3F4F6"))
            c.roundRect(x + dx, y - h + dy, w, h, r, fill=1, stroke=0)
            c.restoreState()
        # 본 카드
        c.saveState()
        c.setFillColor(HexColor(fill))
        c.setStrokeColor(HexColor(stroke))
        c.setLineWidth(border_width)
        c.roundRect(x, y - h, w, h, r, fill=1, stroke=1)
        c.restoreState()

    # 좌/우 박스만 그림 (텍스트는 draw_kpis가 기존 좌표로 그리게 둠)
    _round_box(left_x,  y_card_top, left_w,  card_h)
    _round_box(right_x, y_card_top, right_w, card_h)



def clean_md(md: str) -> list[str]:
    md = re.sub(r"[▶🔸•]", "-", md)      # 아이콘 → 하이픈
    md = re.sub(r"\{|\}", "", md)        # placeholder 중괄호 제거
    md = re.sub(r"\s{2,}", " ", md)      # 공백 정리
    lines = []
    for line in md.splitlines():
        if not line.strip():
            lines.append("")  # 빈 줄 유지
            continue
        # 긴 줄 wrap (한글 50자 기준)
        while stringWidth(line, "KOR_FONT", 9) > 240:   # 폰트 8 pt · 줄이면 왼쪽으로 감 , 한 줄에 허용할 최대 가로폭
            cut = 36                                    # 36자 단위로 wrap 시도
            while stringWidth(line[:cut], "KOR_FONT", 9) < 220 and cut < len(line): # 이 줄에서 허용할 최대 가로 길이
                cut += 1
            lines.append(line[:cut])
            line = line[cut:]
        lines.append(line)
    return lines
# streamlit run ax4_final.py








# ③ 유틸 함수: 점수→백분율, 게이지 그리기, Top‧Bottom 추출
def to_percent(v: float) -> float:
    """1~5 스케일이면 0~100으로 변환, 이미 0~100이면 그대로."""
    return (v - 1) / 4 * 100 if v <= 5 else v
















def percentile_rank(value: float, dist: pd.Series) -> float:
    """value가 dist 내에서 상위 몇 %에 해당하는지 반환 (0~100, 높을수록 우수)"""
    return 100 * (dist < value).mean()








def render_sw_legend_streamlit():
    st.markdown(
        f"""
        <div style="margin:4px 0 8px 0; font-size:12px; color:#6B7280;">
          분류 기준(상위 비율):
          <span style="display:inline-flex; align-items:center; gap:14px; margin-left:6px;">
            <span style="width:10px; height:10px; border-radius:50%; background:{COLOR_POS}; display:inline-block;"></span>
            강점 ≤ 40%
            <span style="width:10px; height:10px; border-radius:50%; background:{COLOR_NEU}; display:inline-block; margin-left:10px;"></span>
            평균권 40–60%
            <span style="width:10px; height:10px; border-radius:50%; background:{COLOR_NEG}; display:inline-block; margin-left:10px;"></span>
            약점 ≥ 60%
          </span>
        </div>
        """,
        unsafe_allow_html=True
    )







def render_strength_weakness(
        avg_scores, overall_mean, overall_series,
        leader_key: str | None = None, raw_df: pd.DataFrame | None = None,
        return_figs: bool = False, return_meta: bool = False):




    subj_strength, subj_weak = ("","")
    if leader_key and raw_df is not None:
        subj_strength, subj_weak = get_subjectives_for_leader(raw_df, leader_key)




    figs, metas = [], []
    top2, bottom2 = pick_strengths_weaknesses(avg_scores)
    targets = list(top2.items()) + list(bottom2.items())




    for row_start in range(0, len(targets), 2):
        cols = st.columns(2)
        for i in range(2):
            if row_start + i >= len(targets): break
            col_name, raw_score = targets[row_start + i]
            perc     = to_percent(raw_score)
            overall  = overall_mean[col_name]
            pct_rank = percentile_rank(raw_score, overall_series[col_name])  # 0~100 (높을수록 상위)
            top_pct  = to_top_pct(pct_rank)                                  # 100 - pct_rank
            tag, color = classify_by_top_band(top_pct)




            # 강/약점 코멘트 생성은 '상/하위 2개' 기반(기존과 동일)
            base_tag = "강점" if (row_start + i) < 2 else "약점"
            desc_key = "강점" if base_tag == "강점" else "약점"




            lines = gen_sw_comment_from_subjective(
                category_key=col_name,
                category_label=LABEL_MAP[col_name],
                sw_type=desc_key,
                subj_strength=subj_strength,
                subj_weak=subj_weak
            )




            with cols[i]:
                fig = draw_semi_gauge(perc, is_strength=(desc_key=="강점"), color_override=color)
                st.pyplot(fig)
                st.markdown(f"<div style='text-align:center; font-size:20px; font-weight:700;'>{int(round(perc))}점</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='text-align:center; font-size:17px; font-weight:600;'>{LABEL_MAP[col_name]}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='text-align:center; font-size:12px; color:#6B7280;'>기준: 전사 평균 {overall:.2f}점</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='text-align:center; font-size:14px; color:{color};'>상위 {top_pct:.0f}% · {tag}</div>", unsafe_allow_html=True)
                for ln in lines[:2]:
                    st.markdown(f"- {ln}")




            figs.append(fig)
            metas.append((perc, LABEL_MAP[col_name], pct_rank, tag, color, lines[:2]))  # PDF용 메타
    outputs = ()
    if return_figs: outputs += (figs,)
    if return_meta: outputs += (metas,)
    return outputs[0] if len(outputs)==1 else outputs
















def draw_semi_gauge(perc: float, is_strength: bool = True, color_override: str | None = None):
    """
    perc : 0-100
    is_strength : 기존 호환용 (미사용 시에도 안전)
    color_override : 특정 색을 강제로 사용 (파랑/빨강/초록 등)
    """
    color_fg = color_override if color_override else ("#3B82F6" if is_strength else "#EF4444")
    color_bg = "#E5E7EB"
    line_col = "#111827"
    R, width = 1.0, 0.3
    fill_a = 180 * np.clip(perc, 0, 100) / 100.0




    fig, ax = plt.subplots(figsize=(2.2, 1.3), subplot_kw={"aspect": "equal"})
    ax.add_patch(Wedge((0,0), R, 180 - fill_a, 180, facecolor=color_fg, width=width, lw=0, zorder=2))
    if fill_a < 180:
        ax.add_patch(Wedge((0,0), R, 0, 180 - fill_a, facecolor=color_bg, width=width, lw=0, zorder=1))
    ax.add_patch(Wedge((0,0), R, 0, 180, facecolor="none", edgecolor=line_col, lw=1.3, zorder=3))
    ax.set_xlim(-1.05, 1.05); ax.set_ylim(0, 1.05); ax.axis("off")
    plt.tight_layout()
    return fig






def calc_biggest_rise_fall(sel_df: pd.DataFrame,
                           year_col: str = "연도",
                           cols: list[str] | None = None,
                           allowed_years=(2022, 2023, 2024),
                           eps: float = 1e-6):
    """
    최근 2개 연도(예: 2023→2024) 간 카테고리별 변화(Δ) 분석.
    eps: 0으로 볼 허용 오차.
    return 예:
      {
        'prev': 2023, 'curr': 2024, 'status': 'both'|'only_up'|'only_down'|'no_change',
        'up_label': '지식공유', 'up_delta': 0.3,                 # 존재 시
        'up_min_label': '상호배려', 'up_min_delta': 0.1,         # 상승만일 때 최소 상승
        'down_label': '공동체의식', 'down_delta': -0.9,          # 존재 시(최대 하락)
        'down_min_label': '자긍심', 'down_min_delta': -0.1       # 하락만일 때 최소 하락(=덜 하락)
      }
    """
    if cols is None:
        cols = VISUAL_COLS
    if sel_df is None or sel_df.empty:
        return None

    by_year = sel_df.groupby(year_col)[cols].mean()
    years = [y for y in allowed_years if y in by_year.index]
    if len(years) < 2:
        years = sorted(by_year.index.tolist())
    if len(years) < 2:
        return None

    y_prev, y_curr = years[-2], years[-1]
    delta = (by_year.loc[y_curr] - by_year.loc[y_prev]).dropna()
    if delta.empty:
        return None

    pos = delta[delta >  eps]
    neg = delta[delta < -eps]
    # 0 근처는 보합으로 간주
    res = {"prev": int(y_prev), "curr": int(y_curr)}

    if len(pos) == 0 and len(neg) == 0:
        res["status"] = "no_change"
        return res

    if len(pos) > 0 and len(neg) > 0:
        res["status"] = "both"
        up_key = pos.idxmax()
        dn_key = neg.idxmin()  # 가장 음수가 큰(=가장 하락)
        res["up_label"] = LABEL_MAP.get(up_key, up_key)
        res["up_delta"] = float(pos[up_key])
        res["down_label"] = LABEL_MAP.get(dn_key, dn_key)
        res["down_delta"] = float(neg[dn_key])
        return res

    if len(pos) > 0 and len(neg) == 0:
        res["status"] = "only_up"
        up_max_key = pos.idxmax()
        up_min_key = pos.idxmin()  # 가장 낮게(작게) 상승
        res["up_label"] = LABEL_MAP.get(up_max_key, up_max_key)
        res["up_delta"] = float(pos[up_max_key])
        res["up_min_label"] = LABEL_MAP.get(up_min_key, up_min_key)
        res["up_min_delta"] = float(pos[up_min_key])
        return res

    # len(pos)==0 and len(neg)>0
    res["status"] = "only_down"
    dn_max_key = neg.idxmin()   # 가장 하락
    dn_min_key = neg.idxmax()   # 가장 적게 하락(0에 가까움)
    res["down_label"] = LABEL_MAP.get(dn_max_key, dn_max_key)
    res["down_delta"] = float(neg[dn_max_key])
    res["down_min_label"] = LABEL_MAP.get(dn_min_key, dn_min_key)
    res["down_min_delta"] = float(neg[dn_min_key])
    return res



def format_rise_fall_items(res) -> list[tuple[str, str]]:
    if not res:
        return []
    p, c = res["prev"], res["curr"]
    s = res.get("status", "both")
    items: list[tuple[str,str]] = []

    if s == "both":
        items.append(("up",   f"가장 상승한 영역: {res['up_label']} ({abs(res['up_delta']):.1f}점)"))
        items.append(("down", f"가장 하락한 영역: {res['down_label']} ({abs(res['down_delta']):.1f}점)"))

    elif s == "only_up":
        items.append(("up",     f"가장 상승한 영역: {res['up_label']} ({abs(res['up_delta']):.1f}점)"))
        items.append(("up_min", f"가장 낮게 상승한 영역: {res['up_min_label']} ({abs(res['up_min_delta']):.1f}점)"))
        # note 항목 제거

    elif s == "only_down":
        items.append(("down",     f"가장 하락한 영역: {res['down_label']} ({abs(res['down_delta']):.1f}점)"))
        items.append(("down_min", f"가장 적게 하락한 영역: {res['down_min_label']} ({abs(res['down_min_delta']):.1f}점)"))
        # note 항목 제거

    elif s == "no_change":
        items.append(("flat", f"전 영역 보합: {p}→{c} 모든 항목 변화 없음 (0.0점)"))

    return items

def render_rise_fall_html(items: list[tuple[str,str]]) -> str:
    up_col   = COLOR_SELF_LINE   # "#2563EB"
    down_col = COLOR_NEG         # "#EF4444"
    note_col = "#6B7280"

    rows = []
    for kind, text in items:
        if kind in ("up","up_min"):
            icon = f"<span style='color:{up_col}; font-weight:900;'>▲</span>"
        elif kind in ("down","down_min"):
            icon = f"<span style='color:{down_col}; font-weight:900;'>▼</span>"
        elif kind == "note":
            icon = f"<span style='color:{note_col}; font-weight:900;'>※</span>"
        else:  # flat
            icon = ""
        rows.append(f"<div>{icon} {text}</div>")
    return (
        "<div style='margin:8px 4px 0; font-size:16px; color:#111827; line-height:1.7; font-weight:500;'>"
        + "".join(rows) + "</div>"
    )
    

def draw_rise_fall_line(c, x, y, kind: str, text: str, max_width: float) -> float:
    # 아이콘 & 색
    if kind in ("up","up_min"):
        c.setFillColor(HexColor(COLOR_SELF_LINE)); icon = "▲"
    elif kind in ("down","down_min"):
        c.setFillColor(HexColor(COLOR_NEG)); icon = "▼"
    elif kind == "note":
        c.setFillColor(HexColor("#6B7280")); icon = "※"
    else:  # flat
        icon = ""

    # 아이콘
    if icon:
        c.setFont("KOR_FONT_BOLD", 6)
        c.drawString(x, y, icon)
        x += 12  # 아이콘 뒤 여백

    # 본문
    c.setFillColor(HexColor("#111827"))
    c.setFont("KOR_FONT", 8)
    return draw_markdown_line(c, x, y, text, max_width)

def pick_strengths_weaknesses(avg: pd.Series, n: int = 2):
    """평균 점수 Series → (강점 2, 약점 2) 반환"""
    sorted_ = avg.sort_values(ascending=False)
    return sorted_.head(n), sorted_.tail(n)
















def fig_to_png_bytes(fig) -> bytes:
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=250, bbox_inches="tight", facecolor="white")
    buf.seek(0)
    return buf.getvalue()




def draw_hero_pill(
    c, W, H, name: str, leader_meta: dict | None = None, y: float | None = None,
    title_size: int = 26, pad_x: int = 18, pad_y: int = 8, r: int = 16,
    brand: str = "#EA002C", show_shadow: bool = True   # ← 옵션 추가
) -> float:
    if y is None:
        y = H - 60

    left_txt  = f"{name} 팀장의"
    right_txt = " Red Re:born"
    title_font = "KOR_FONT_BOLD"

    left_w  = stringWidth(left_txt,  title_font, title_size)
    right_w = stringWidth(right_txt, title_font, title_size)
    title_w = left_w + right_w

    pill_w = title_w + pad_x * 2
    pill_h = title_size + pad_y * 2
    x = (W - pill_w) / 2
    top_y = y
    bottom_y = y - pill_h

    # ✅ (복구) 빨간 그림자(뒤쪽 라운드) — 원하면 show_shadow=False로 끄기
    if show_shadow:
        c.saveState()
        c.setFillColor(HexColor(brand))
        c.roundRect(x + 8, bottom_y - 10, pill_w - 16, pill_h + 12, r, fill=1, stroke=0)
        c.restoreState()

    # ✅ (복구) 흰색 알약(빨간 보더)
    c.saveState()
    c.setFillColor(HexColor("#FFFFFF"))
    c.setStrokeColor(HexColor(brand))
    c.setLineWidth(1.2)
    c.roundRect(x, bottom_y, pill_w, pill_h, r, fill=1, stroke=1)
    c.restoreState()

    # streamlit run ax4_final.py
    # 타이틀 텍스트 (항상 카드 도형 *뒤에* 그려서 겹침 방지)
    tx = x + pad_x
    ty = bottom_y + pad_y + (title_size - 22)        # 1) 제목:  "류동현 팀장의  Red Re:born"
    c.setFont(title_font, title_size); c.setFillColor(HexColor("#111827"))
    c.drawString(tx, ty, left_txt)
    c.setFillColor(HexColor(brand))
    c.drawString(tx + left_w, ty, right_txt)

    # 메타/안내
    meta_parts = []
    if leader_meta:
        if leader_meta.get("id"):       meta_parts.append(f"ID: {leader_meta['id']}")
        if leader_meta.get("company"):  meta_parts.append(f"회사: {leader_meta['company']}")
        if leader_meta.get("team"):     meta_parts.append(f"팀: {leader_meta['team']}")
        if leader_meta.get("position"): meta_parts.append(f"직위:{leader_meta['position']}")

    meta_y = bottom_y - 20          # 2) 메타:  "ID: … · 회사: … · 팀: … · 직위: …"
    tip_y  = meta_y - 10            # 3) 안내:  "좌측 정보를 모두…"
    c.setFont("KOR_FONT", 9);  c.setFillColor(HexColor("#9CA3AF"))
    c.drawCentredString(W/2, meta_y, " · ".join(meta_parts))
    c.setFont("KOR_FONT_BOLD", 9); c.setFillColor(HexColor("#111827"))
    c.drawCentredString(W/2, tip_y, "좌측 정보를 모두 읽은 후 단계별 개선방안을 확인해주세요 :)")

    return tip_y - 8

# === [강점&약점] 게이지 2x2: 윗줄/아랫줄 개별 조정 ===
GA_ROW_SHIFT_TOP    = 0    # 윗줄 전체를 통째로 위(+)/아래(-)로 밀기
GA_ROW_SHIFT_BOTTOM = 0    # 아랫줄 전체를 통째로 위(+)/아래(-)로 밀기

GA_TOP_CFG = {   # 윗줄(게이지 2개)
    "IMG_DX": 0, "IMG_DY": 233, "IMG_H": 50, "IMG_W_MARGIN": 12,            # IMG_DY 값 ↑ → 이미지가 더 아래로
    "TITLE_DX": 22, "TITLE_DY": 238,                                        # TITLE_DX,TITLE_DY ↑ 더 아래로
    "SUB_DX": 26, "SUB_DY": 248,                                             # SUB_DX,SUB_DY ↑ 더 아래로
    "BULLET_DX": 0, "BULLET_FIRST_DY": 258, "BULLET_LINE_GAP": 9,           #코멘트 조절 ↑ 더 아래로
    "WRAP_COLS": 20,                                                        # 한 줄 글자수 기준(↑ → 줄개수 감소)
    "COL_GAP": 18,   # ← 윗줄 두 게이지 사이 간격(총량, pt)
}
GA_BOTTOM_CFG = {  # 아랫줄(게이지 2개)
    "IMG_DX": 0, "IMG_DY": 260, "IMG_H": 50, "IMG_W_MARGIN": 12,            # IMG_DY 값 ↑ → 이미지가 더 아래로
    "TITLE_DX": 22, "TITLE_DY": 265,                                         # TITLE_DX,TITLE_DY ↑ 더 아래로
    "SUB_DX": 26, "SUB_DY": 275,                                            # SUB_DX,SUB_DY ↑ 더 아래로
    "BULLET_DX": 0, "BULLET_FIRST_DY": 285, "BULLET_LINE_GAP": 9,            #코멘트 조절 ↑ 더 아래로 
    "WRAP_COLS": 20,                                                        # 한 줄 글자수 기준(↑ → 줄개수 감소)
    "COL_GAP": 18,    # ← 아랫줄 두 게이지 사이 간격(총량, pt)
}

def build_dashboard_pdf(
    leader_name: str,
    radar_fig,
    gauge_figs,
    gauge_meta,
    trend_fig,
    trend_comment: str,
    step1_md: str,
    step2_md: str,
    step3_md: str,
    profile_img=None,
    leader_meta: dict | None = None,
    total_score_100: float | None = None,   # 좌 KPI
    weakest_label: str | None = None,       # 우 KPI
    weakest_score_1to5: float | None = None,
    rise_fall_items: list[tuple[str,str]] | None = None,
    ) -> bytes:

    W, H = A4
    M = 36
    LEFT_X  = M
    COL_L_W = 270
    GAP     = 18
    RIGHT_X = LEFT_X + COL_L_W + GAP
    COL_R_W = W - RIGHT_X - M

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # ─ 0) 상단 히어로(알약) ─
    display_name = (leader_meta or {}).get("name", leader_name) or leader_name
    y_cursor = H - 9
    y_cursor = draw_hero_pill(
    c, W, H,
    name=display_name,
    leader_meta=leader_meta,
    y=y_cursor,
    title_size=26, pad_x=18, pad_y=8, r=16,
    brand="#EA002C",
    show_shadow=True   # 빨간 그림자까지 보이게. 붉은 띠가 거슬리면 False로.
    )

    # ─ 1) KPI 2장(원하는 Y에 배치) ─
    val_100 = float(total_score_100 or 0.0)
    kpi_label = weakest_label or ""
    # ① 먼저 박스만 깔기 (텍스트 좌표는 그대로)     
    draw_kpi_boxes_behind(
    c,
    y_top=y_cursor - 20,   # 텍스트 기준 y (기존 값 그대로)
    card_h=29,             # 박스크기조절
    top_offset=20,         # 박스 상단 y 미세 이동 (텍스트는 그대로)
    left_x=LEFT_X,         # 좌 박스 X
    left_w=280,            # 좌 박스 너비
    right_x=RIGHT_X,       # 우 박스 X
    right_w=COL_R_W,       # 우 박스 너비
    r=16
    )

    # ② 기존 텍스트 그대로 
    draw_kpis(
        c, y_top=y_cursor - 20,
        val_100=val_100,
        label=kpi_label,
        with_cards=False,                 
        card_h=64,
        left_x=LEFT_X, right_x=RIGHT_X,
        left_w=COL_L_W, right_w=COL_R_W
    )
    # KPI 높이(64) + 아래 여백(36)만큼 커서 이동
    y_cursor -= 24

    # ─ 2) 1행: 좌[3개년 추이] / 우[STEP1] ─
    # (아래 절대좌표는 기존 시안 유지. 원하면 y_cursor 기반으로 바꿔도 OK)
    draw_section_pill(c, LEFT_X+8, H-120, "3개년 리더십 추이")
    TREND_H = 162
    trend_img = ImageReader(BytesIO(fig_to_png_bytes(trend_fig)))
    draw_card(c, LEFT_X, H-168- TREND_H, COL_L_W, TREND_H, r=18)
    c.drawImage(trend_img, LEFT_X+12, H-143- TREND_H + 20, width=COL_L_W-24, height=TREND_H-24, mask='auto')

    # 추이 코멘트
    y = H-152- TREND_H + 24          # ← 코멘트 첫 줄의 시작 Y
    for raw_line in (trend_comment or "").splitlines():
        ln = raw_line.strip()
        if not ln:
            continue
        ln = re.sub(r'^\s*[•\-]\s*', '', ln)
        y = draw_markdown_line(c, LEFT_X+16, y, "• " + ln, max_width=COL_L_W-32) - 2
    # ⬇ 추가: 파란/빨간 삼각형 라인
    if rise_fall_items:
        y -= 2                      # ← 코멘트와 △/▽ 사이 간격
        for kind, text in rise_fall_items:
            y = draw_rise_fall_line(
                    c,
                    LEFT_X + 16,     # ← X 좌표(왼쪽 여백). 더 오른쪽으로 밀고싶으면 +값을 키워요 (예: +24)
                    y+3,               # ← 시작 Y 좌표
                    kind,
                    text,
                    max_width=COL_L_W - 32  # ← 줄 폭(랩핑 폭). 더 짧게 감싸고 싶으면 값을 줄이세요.
                ) - 2               # ← 각 항목 뒤에 추가로 내리는 간격(줄 간격). 0~6 정도로 조절

    draw_section_pill(c, RIGHT_X+8, H-120, "단계별 개선방안")
    draw_step_box(c, RIGHT_X, H-145, COL_R_W, 200, 1, "즉각적인 개인 실천의 시작", step1_md)

    # ─ 3) 2행: 좌[종합 다이어그램] / 우[STEP2] ─
    draw_section_pill(c, LEFT_X+8, H-345, "종합 다이어그램")                 # 종합다이어그램전체조절(칩) Y
    RADAR_H = 210                                                            # 카드/이미지 높이
    radar_img = ImageReader(BytesIO(fig_to_png_bytes(radar_fig)))
    draw_card(c, LEFT_X, H-408- RADAR_H, COL_L_W, RADAR_H, r=18)                # 카드 박스 X/Y/W/H
    c.drawImage(radar_img, LEFT_X+20,H-371- RADAR_H + 36, width=COL_L_W-40, height=RADAR_H-36, mask='auto') # ← 이미지 X/Y 이미지 W/H

    draw_step_box(c, RIGHT_X, H-370, COL_R_W, 230, 2, "체계적인 교육을 통한 역량 강화", step2_md) #스텝2위치조절

    # ─ 4) 3행: 좌[강점 & 약점] / 우[STEP3] ─
    draw_section_pill(c, LEFT_X+8, H-552, "강점 & 약점")  #강점약점전체조절
    draw_sw_legend(c, x=LEFT_X+120, y=H-562, gap=7, label_pad=3,fs=5) #범례

    GA_H = 220                                                  # ← 블록 높이
    draw_card(c, LEFT_X, 240, COL_L_W, GA_H, r=18)

    # 게이지 4개 (2x2)
    cell_w, cell_h = (COL_L_W-32)/2, (GA_H-32)/2
    for i, (fig, meta) in enumerate(zip(gauge_figs[:4], gauge_meta[:4])):
        col, row = i % 2, i // 2                          # col: 0(좌)/1(우), row: 0(윗줄)/1(아랫줄)

        row_cfg   = GA_TOP_CFG if row == 0 else GA_BOTTOM_CFG
        row_shift = GA_ROW_SHIFT_TOP if row == 0 else GA_ROW_SHIFT_BOTTOM
        col_gap   = row_cfg.get("COL_GAP", 0)              # ← (신규) 이 행에서 두 칼럼 사이 추가 간격(총량, pt)

        # 셀 기준점(gx, gy)
        #  - col==0(왼쪽)일 때는 -col_gap/2만큼 왼쪽으로,
        #  - col==1(오른쪽)일 때는 +col_gap/2만큼 오른쪽으로 밀어 여백을 늘린다.
        gx = LEFT_X + 16 + col * cell_w + (col_gap/2 if col == 1 else -col_gap/2)  # ← 가로 간격(행별) 조절
        gy = 240 + GA_H - 16 - row * cell_h + row_shift                            # ← 세로 행 위치(행별) 조절

        # 1) 게이지 이미지
        img = ImageReader(BytesIO(fig_to_png_bytes(fig)))
        c.drawImage(
            img,
            gx + row_cfg.get("IMG_DX", 0),                  # 이미지의 미세 X 이동(+ 오른쪽)
            gy - row_cfg["IMG_DY"],                         # 이미지의 미세 Y 오프셋(값 ↑ → 아래)
            width=cell_w - row_cfg["IMG_W_MARGIN"],         # 이미지 폭(↑ 마진 → 폭↓ → 가운데 여백 커짐)
            height=row_cfg["IMG_H"],                        # 이미지 높이
            mask='auto'
        )

        # 2) 메타 텍스트 (제목/서브)
        perc, lbl, pct_rank, tag, color_hex, desc_lines = meta
        c.setFont("KOR_FONT_BOLD", 9); c.setFillColor(HexColor("#111827"))
        c.drawString(gx + row_cfg["TITLE_DX"], gy - row_cfg["TITLE_DY"], f"{int(round(perc))}점  {lbl}")  # TITLE_DX/DY로 미세 조정

        c.setFont("KOR_FONT", 7); c.setFillColor(HexColor(color_hex))
        top_pct = to_top_pct(pct_rank)
        c.drawString(gx + row_cfg["SUB_DX"], gy - row_cfg["SUB_DY"], f"상위 {top_pct:.0f}% · {tag}")      # SUB_DX/DY로 미세 조정

        # 3) 코멘트(불릿) — 문장 기준, 첫 줄만 불릿, 이후 줄은 들여쓰기
        c.setFillColor(HexColor("#000000")); c.setFont("KOR_FONT", 7)
        ty = gy - row_cfg["BULLET_FIRST_DY"]

        # 들여쓰기 픽셀값(설정값 없으면 글꼴 기준 자동 계산)
        indent_dx = row_cfg.get("BULLET_INDENT_DX")
        if indent_dx is None:
            indent_dx = stringWidth("•  ", "KOR_FONT", 7)  # 불릿+공백 폭

        for sent in (desc_lines or [])[:2]:                     # ← 문장 단위
            wraps = textwrap.wrap(sent, row_cfg["WRAP_COLS"])   # 화면 폭에 맞춰 감싸기
            for j, seg in enumerate(wraps):
                if j == 0:
                    # 첫 줄만 불릿 표시
                    c.drawString(gx + row_cfg["BULLET_DX"], ty, "• " + seg)
                else:
                    # 이후 줄은 불릿 없이 들여쓰기
                    c.drawString(gx + row_cfg["BULLET_DX"] + indent_dx, ty, seg)
                ty -= row_cfg["BULLET_LINE_GAP"]



    draw_step_box(c, RIGHT_X, 230, COL_R_W, 230, 3, "팀/조직 차원의 변화", step3_md)  #스텝3위치조절

    # ─ 완료 ─
    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()















def draw_gauge_text(c, x, y, perc, label, pct_rank, tag, color_hex):
    """
    c : ReportLab canvas
    (x, y) : 왼쪽 하단 기준
    perc : 0~100
    label : '자긍심' 등
    pct_rank : 0~100
    tag : '강점'·'약점'·'개선 필요' 등
    color_hex : '#EF4444'
    """
    c.setFont("KOR_FONT", 9)
    c.drawCentredString(x + 60, y - 12, f"{int(round(perc))}점")
    c.drawCentredString(x + 60, y - 25, label)
    top_pct = to_top_pct(pct_rank)
    c.setFillColor(color_hex)
    c.drawCentredString(x + 60, y - 38, f"상위 {top_pct:.0f}% · {tag}")
    c.setFillColor(HexColor("#000000"))








if 'survey_df' not in st.session_state:
    st.session_state.survey_df = None
   




# ====== 한글 폰트 등록 ======
FONT_PATH = "C:/Windows/Fonts/malgun.ttf"
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont("KOR_FONT", FONT_PATH))
else:
    raise FileNotFoundError("한글 폰트를 찾을 수 없습니다.")




# 볼드용 폰트도 등록
BOLD_PATH = "C:/Windows/Fonts/malgunbd.ttf"
if os.path.exists(BOLD_PATH):
    pdfmetrics.registerFont(TTFont("KOR_FONT_BOLD", BOLD_PATH))
else:
    # 대체로 굴림(bold) 등 다른 볼드 폰트로 대체하셔도 됩니다
    pdfmetrics.registerFont(TTFont("KOR_FONT_BOLD", FONT_PATH))
























def chat_ax4(messages, max_retries=4, base_delay=1.2):
    for i in range(max_retries):
        try:
            r = client.chat.completions.create(model="ax4", messages=messages)
            return r.choices[0].message.content.strip()
        except RateLimitError:
            if i == max_retries - 1:
                raise
            time.sleep(base_delay * (2 ** i) + random.random())
        except Exception:
            if i == max_retries - 1:
                raise
            time.sleep(base_delay * (2 ** i) + random.random())












# ====== 교육 프로그램 추천 함수 ======
@st.cache_data(
    ttl=3600, show_spinner=False,
    hash_funcs={pd.DataFrame: lambda df: hashlib.md5(df.to_csv(index=False).encode()).hexdigest()}
)
def recommend_programs(weak_text, edu_df):
    edu_summary = edu_df[["카드명","카테고리분류","카드소개내용"]].to_string(index=False)
    prompt = f"""다음은 팀장의 약점입니다:
{weak_text}




아래는 추천 가능한 사내 교육 목록입니다:
{edu_summary}




이 중 최적 2개를 다음 형식으로:
프로그램명:
이유:
팁:
"""
    messages = [
        {"role": "system", "content": "당신은 전문 리더십 교육 컨설턴트입니다."},
        {"role": "user", "content": prompt}
    ]
    return chat_ax4(messages)








# 레이더 차트 그리는 함수 정의
def plot_radar_7(score_dict: dict, title: str = ""):
     # 1) 라벨 / 값 준비
    labels = list(score_dict.keys())
    values = list(score_dict.values())
    values = values + values[:1]
    # 2) 각도 계산
    angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist()
    angles += angles[:1]
    # 3) 플롯 준비
    fig, ax = plt.subplots(figsize=(6,6), subplot_kw=dict(polar=True))
    # 12시 방향부터, 시계방향으로 그리기
    ax.set_theta_offset(np.pi/2)
    ax.set_theta_direction(-1)
    # 4) 데이터 그리기 (선 + 채우기)
    ax.plot(angles, values, marker='o', color='#EE0000', linewidth=2)
    ax.fill(angles, values, color='#EE0000', alpha=0.2)
    # 5) 축 라벨(카테고리) 설정
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=12, fontweight='bold' ,fontfamily='Malgun Gothic')
    ax.tick_params(axis='x', pad=20)












    # 6) 반경축(1~5) 설정
    ax.set_ylim(1, 5)
    ax.set_yticks([1,2,3,4,5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"],  # ★ 눈금 라벨 표시
                       fontweight='bold',
                       fontsize=10,
                       fontfamily='Malgun Gothic')




    # 7) 그리드 스타일
    ax.xaxis.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.yaxis.grid(True, color='gray', linestyle='--', linewidth=0.5)
   
    # 8) 제목
    ax.set_title(title, fontsize=18, pad=20)
   
    # 9) 출력
    plt.tight_layout()
    st.pyplot(fig)
    return fig 


#3개년프롬프트
def make_trend_commentary_via_llm_from_series(trend_series: pd.Series, weakest_label: str) -> str:
    """
    trend_series: index=연도(int), values=1~5 스케일 평균 (예: 2022~2024)
    weakest_label: 7개 객관식 분류 중 최저 영역의 '라벨'(예: '상호배려')
    출력: 정확히 3문장 (존댓말). 3문장 구조:
      1) 각 연도 값(100점 기준)과 흐름 사실 요약
      2) 연도 간 변화폭 숫자로만 명시
      3) 최저 영역(weakest_label) 우선 개선 문장
    """
    # ① 2022~2024만 추출 + NaN 제거
    ser = trend_series.reindex([2022, 2023, 2024]).dropna()
    if ser.empty:
        # 예외: 사용 가능한 데이터가 없음
        return f"(데이터 없음) 현재 7개 객관식 분류 중 '{weakest_label}' 영역을 우선 개선하면 종합 점수 상승에 기여할 것으로 보입니다."

    # ② 100점 환산 (1~5 → 0~100)
    ser_100 = ser.apply(to_percent)  # (v-1)/4*100
    # trend_key 예: "2022:72.5|2023:68.1|2024:71.2"
    trend_key = "|".join(f"{int(y)}:{float(s):.1f}" for y, s in ser_100.items())

    # ③ 프롬프트 (정확한 분석 3문장 + 추측 금지 + 3문째는 최저 영역 개선 문장)
    prompt_soft = f"""
    다음은 연도별 점수(100점 기준)입니다:
    {trend_key}

    [작성 규칙]
    - 정확히 3문장, 한국어 존댓말, 부드러운 톤으로 작성합니다.
    - 번호·제목·불릿·이모지·원인 추정 금지, 제공된 수치만 사용합니다.
    - 모든 숫자는 소수점 한 자리까지 표기하고 ‘점’ 단위를 붙입니다.

    [형식]
    1문장: 연도별 값을 반드시 오름차순으로 나열하고 문장을 ‘…확인됩니다.’로 마무리합니다. 2022년 데이터가 없을 경우 2023, 2024년 순대로 설명합니다. 2022,2023,2024 데이터가 모두 있을 경우 반드시 순서대로 설명합니다.
    예) "2022년은 87.2점, 2023년은 76.1점, 2024년은 68.6점으로 확인됩니다."
    2문장: 바로 전 연도 대비 변화폭을 절대값으로 제시하고,
        증가면 ‘높아졌습니다’, 감소면 ‘낮아졌습니다’, 0이면 ‘변화가 없습니다’라고 씁니다.
    예) "2023년 대비 2024년에 7.5점 낮아졌습니다."
    (연도가 3개면 "2022년 대비 2023년에 …, 2023년 대비 2024년에 …"처럼 쉼표로 구분합니다.)
    3문장: "종합 점수 향상을 위해 ‘{weakest_label}’을(를) 먼저 보완하시면 좋겠습니다."로 작성합니다.
    """.strip()

    messages = [
        {"role": "system", "content": "당신은 산술 계산을 정확히 수행하는 데이터 분석가입니다. 제공된 수치 외 추론을 하지 않습니다."},
        {"role": "user", "content": prompt_soft}
    ]

    # ④ LLM 시도
    try:
        out = chat_ax4(messages).strip()
        if out:
            # 혹시 3문장 초과/미달이면 안전하게 절삭/보정
            sents = [s.strip() for s in re.split(r'(?<=[.。])\s+', out) if s.strip()]
            if len(sents) >= 3:
                return " ".join(sents[:3])
            # 부족하면 아래 보수 로직으로 대체
    except Exception:
        pass

    # ⑤ 보수(fallback): 규칙에 맞춰 직접 3문장 생성
    years = list(ser_100.index.astype(int))
    vals = [float(f"{v:.1f}") for v in ser_100.values]

    # 1문장
    parts = [f"{years[i]}년 {vals[i]:.1f}점" for i in range(len(years))]
    sent1 = f"{', '.join(parts)}으로 집계되었습니다."

    # 2문장 (변화폭)
    if len(vals) >= 2:
        deltas = []
        for i in range(1, len(vals)):
            diff = round(vals[i] - vals[i-1], 1)
            sign = "+" if diff > 0 else ""
            if diff == 0:
                deltas.append(f"{years[i-1]}→{years[i]}은 변화 없음")
            else:
                deltas.append(f"{years[i-1]}→{years[i]}은 {sign}{diff:.1f}점")
        sent2 = f"{' , '.join(deltas)} 변화가 있었습니다."
    else:
        sent2 = "비교할 연간 변화 데이터가 없어 변화 분석은 불가합니다."

    # 3문장 (최저 영역 개선 메시지)
    sent3 = f"7개 객관식 분류 중 '{weakest_label}' 영역을 우선 개선하면 종합 점수가 크게 올라갈 것으로 보입니다."

    return " ".join([sent1, sent2, sent3])




   
def plot_yearly_bar_with_avg(sel_df: pd.DataFrame,
                             overall_trend: pd.Series,
                             year_col: str = "연도",
                             score_col: str = "Composite",
                             all_years = (2022, 2023, 2024),
                             y_tick_step: float = 10):
    # 1~5 → 0~100 환산
    per_year = sel_df.groupby(year_col)[score_col].mean().reindex(all_years).apply(to_percent)
    corp     = overall_trend.reindex(all_years).apply(to_percent)




    x = np.arange(len(all_years))
    fig, ax = plt.subplots(figsize=(8, 4))




    # 막대: 연도별 색 적용 + 라벨 Bold
    for i, y in enumerate(per_year.values):
        if pd.notna(y):
            yr  = all_years[i]
            col = YEAR_COLORS.get(yr, BASE_2024)
            ax.bar(x[i], y, width=BAR_WIDTH, color=col, edgecolor=col)  # ← 여기만 변경
            ax.text(x[i], y + 1.2, f"{y:.0f}", ha="center", va="bottom", fontsize=9, fontweight="bold")








    # 전사 평균 라인(색 고정)
    ax.plot(x, corp.values, marker="o", linewidth=2.2,
            color=COLOR_CORP_LINE, label="전사 평균")




    for i, y in enumerate(corp.values):
        if pd.notna(y):
            ax.text(x[i], y + 1.2, f"{y:.0f}", ha="center", va="bottom", fontsize=9, color="#333")




    ax.set_xticks(x); ax.set_xticklabels([str(y) for y in all_years])
    ax.set_ylim(0, 100)                                      # ★ 고정: 0–100
    ax.yaxis.set_major_locator(MultipleLocator(y_tick_step)) # 기본 10점 간격
    ax.grid(True, axis="y", linestyle="--", alpha=0.4)
    ax.set_xlabel("Years", fontsize=8, loc="right")
    ax.set_ylabel("(100점 만점)")
    ax.legend(loc="upper left")




    plt.tight_layout()
    st.pyplot(fig)
    return fig












def plot_radar_compare_7(sel_score: dict, ref_score: dict, title: str = "",
                         legend_labels=("선택 리더", "전사 평균"), show_values: bool = False):
    order = [LABEL_MAP[c] for c in VISUAL_COLS]




    sel_vals = [to_percent(float(sel_score.get(k, np.nan))) for k in order]
    ref_vals = [to_percent(float(ref_score.get(k, np.nan))) for k in order]
    sel_vals += sel_vals[:1]; ref_vals += ref_vals[:1]




    angles = np.linspace(0, 2*np.pi, len(order), endpoint=False).tolist(); angles += angles[:1]




    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi/2); ax.set_theta_direction(-1)




    ax.set_ylim(0, 100)
    ax.set_yticks([0,20,40,60,80,100])
    ax.set_yticklabels(["0","20","40","60","80","100"], fontsize=10, fontweight="bold", fontfamily="Malgun Gothic")




    ax.set_xticks(angles[:-1]); ax.set_xticklabels(order, fontsize=12, fontweight="bold", fontfamily="Malgun Gothic")
    ax.tick_params(axis='x', pad=20)




    col_ref, col_sel = COLOR_CORP_LINE, COLOR_SELF_LINE
    ax.plot(angles, ref_vals, linewidth=2, marker='o', color=col_ref, alpha=0.9, label=legend_labels[1])
    ax.fill(angles, ref_vals, color=col_ref, alpha=0.15)
    ax.plot(angles, sel_vals, linewidth=2.5, marker='o', color=col_sel, alpha=1.0, label=legend_labels[0])
    ax.fill(angles, sel_vals, color=col_sel, alpha=0.25)




    if show_values:
        for ang, v in zip(angles[:-1], sel_vals[:-1]):
            ax.text(ang, v + 1.5, f"{v:.0f}", ha="center", va="center", fontsize=9)




    ax.xaxis.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.yaxis.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.set_title(title, fontsize=18, pad=20)
    ax.legend(loc="upper left", bbox_to_anchor=(0.05, 1.05), frameon=False)
    plt.tight_layout()
    st.pyplot(fig)
    return fig




##분홍
# --- 색 변수(원하면 바꿔가며 실험) ---
#ACCENT = "#FEE1E8"  # 박스 배경
#HOVER  = "#FFF4F7"  # 아주 옅은 핑크
#BORDER = "#F9A8C7"  # 살짝 진한 핑크


# 사이드바(연분홍 테마)
#PINK_BG    = "#FEE1E8"   # 배경
#PINK_EDGE  = "#FFE9F0"   # 버튼/보더 포인트
#PINK_HOVER = "#FDA4AF"   # hover

##초록
# 박스/컴포넌트 기본
#ACCENT = "#62BFBC"  # 박스 배경
#HOVER  = "#3AA5A1"  # 아주 옅은 핑크 (hover/선택)
#BORDER = "#4FB0AD"  # 살짝 진한 핑크 (보더/점선)

# 사이드바
#PINK_BG    = "#CAF0EE"  # 배경 (조금 더 연함)
#PINK_EDGE  = "#9FD7D6"  # 버튼/보더 포인트 (유지)
#PINK_HOVER = "#5AB8B5"  # hover (유지)

#파랑
#ACCENT = "#93BEF5"  # 박스 배경
#HOVER  = "#EAF3FE"  # 아주 옅은 하늘색 (hover/선택)
#BORDER = "#6AA7EC"  # 살짝 진한 블루 (보더/점선)

# 사이드바 (블루 톤)
#PINK_BG    = "#D9E5F4"  # 배경
#PINK_EDGE  = "#BFD4EE"  # 버튼/보더 포인트
#PINK_HOVER = "#8FB5EF"  # hover



# streamlit run ax4_final.py
# CSS

# --- 색 변수(원하면 바꿔가며 실험) ---
# 박스/컴포넌트 기본 (블루 톤)
ACCENT = "#62BFBC"  # 박스 배경
HOVER  = "#3AA5A1"  # 아주 옅은 핑크 (hover/선택)
BORDER = "#4FB0AD"  # 살짝 진한 핑크 (보더/점선)

# 사이드바
PINK_BG    = "#CAF0EE"  # 배경 (조금 더 연함)
PINK_EDGE  = "#9FD7D6"  # 버튼/보더 포인트 (유지)
PINK_HOVER = "#5AB8B5"  # hover (유지)

st.markdown(f"""
<style>
/* ─ 기존 스타일 유지 ─ */
[data-testid="stSidebar"] * {{ color:#F2F2F2 !important; }}
[data-testid="stSidebar"] svg {{ fill:#F2F2F2 !important; }}
[data-testid="stSidebar"] input[type="radio"]:checked + label > div > div:first-child {{ color:#EE0000 !important; }}
textarea {{ background:#0A192F !important; color:#F2F2F2 !important; border:1px solid #EE0000 !important; border-radius:6px !important; }}
[data-testid="stTextArea"] * {{ color:#F2F2F2 !important; font-weight:500 !important; }}
textarea::placeholder{{ color:#FFD700 !important; opacity:1 !important; }}
input[placeholder="예: 홍희,11001"]{{ background:#E8F6FF !important; border:1px solid #1D4ED8 !important; color:#0F172A !important; }}
input[placeholder="예: 홍희,11001"]::placeholder{{ color:#1D4ED8 !important; }}
.block-container{{ padding-top:1.2rem; }}

/* ===== Upload Card ===== */
.upload-card{{ background:#fff; border:1px solid #E5E7EB; border-radius:12px; padding:18px 16px 14px; margin:8px 0 18px; box-shadow:0 1px 2px rgba(0,0,0,.05); }}
.uc-head{{ display:flex; align-items:center; justify-content:space-between; gap:12px; }}
.uc-left{{ display:flex; align-items:center; gap:14px; }}
.uc-icon{{ width:64px; height:64px; object-fit:contain; }}
.uc-title{{ font-size:24px; font-weight:800; color:#111827; line-height:1; }}
.uc-sub{{ margin-top:6px; color:#374151; }}
.uc-right{{ font-size:32px; opacity:.6; }}

.upload-card [data-testid="stFileUploader"] > div > div{{ background:#fff !important; border:1.6px solid #EA002C !important; border-radius:10px !important; }}
.upload-card [data-testid="stFileUploader"] *{{ color:#111 !important; }}
.upload-card [data-testid="stFileUploader"] button{{ background:#EA002C !important; color:#fff !important; border-radius:8px !important; font-weight:700; }}
.upload-card [data-testid="stFileUploader"] > div{{ margin-top:10px; }}
.uc-right-icon{{ width:32px; height:32px; object-fit:contain; opacity:.9; }}
.uc-right-icon:hover{{ opacity:1; transform:translateY(-1px); transition:.15s; }}

.zip-card{{ position:relative; height:180px; background:#fff; border:1px solid #D1D5DB; border-radius:12px; box-shadow:0 1px 2px rgba(0,0,0,.05); margin-bottom:8px; }}
.zip-top-icon{{ position:absolute; top:16px; left:16px; width:72px; height:72px; object-fit:contain; }}
.zip-title{{ position:absolute; top:22px; left:110px; font-size:22px; font-weight:800; color:#111827; line-height:1.2; }}
.zip-btm-icon{{ position:absolute; right:16px; bottom:16px; width:72px; height:72px; object-fit:contain; opacity:.95; }}

/* ───────── 파란 박스(셀렉트/드롭존)만 연주황으로 ───────── */
/* 1) 셀렉트박스(닫힌 입력창) */
.stSelectbox [data-baseweb="select"] > div {{
  background:{ACCENT} !important;
  border:1.5px solid {ACCENT} !important;
  color:#111 !important;
  border-radius:8px !important;
}}
/* 입력/플레이스홀더 글자색 & 화살표 아이콘 */
.stSelectbox [data-baseweb="select"] input,
.stSelectbox [data-baseweb="select"] span {{ color:#111 !important; }}
.stSelectbox svg {{ fill:#111 !important; }}

/* 펼친 메뉴 hover / 선택색 */
.stSelectbox [data-baseweb="menu"] [role="option"]:hover,
.stSelectbox [data-baseweb="menu"] [role="option"][aria-selected="true"] {{
  background:{HOVER} !important;
  color:#111 !important;
}}

/* 2) 파일 업로더 드롭존 */
[data-testid="stFileUploaderDropzone"] {{
  background:{ACCENT} !important;
  border:1.6px dashed {BORDER} !important;
  border-radius:10px !important;
  color:#111 !important;
}}
[data-testid="stFileUploaderDropzone"] * {{ color:#111 !important; }}

/* ───────── 사이드바 연분홍 테마(맨 아래에 둬서 덮어쓰기) ───────── */
section[data-testid="stSidebar"] {{ background:{PINK_BG} !important; }}
section[data-testid="stSidebar"] * {{ color:#111 !important; }}
section[data-testid="stSidebar"] svg {{ fill:#111 !important; }}

/* 사이드바 라디오/버튼/입력 톤 */
section[data-testid="stSidebar"] input[type="radio"]:checked + label > div > div:first-child {{
  color:{PINK_EDGE} !important;
}}
section[data-testid="stSidebar"] button,
section[data-testid="stSidebar"] [data-baseweb="button"] {{
  background:{PINK_EDGE} !important; 
  color:#fff !important; 
  border-radius:8px !important;
}}
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] select,
section[data-testid="stSidebar"] textarea {{
  background:#FFF1F2 !important;
  border:1px solid {PINK_EDGE} !important;
  color:#111 !important;
}}
section[data-testid="stSidebar"] button:hover {{
  filter:brightness(0.98);
  box-shadow:0 0 0 2px {PINK_HOVER} inset;
}}

:root {{
  --brand:{PRIMARY}; --ink:#111827; --muted:#6B7280; --card:#FFFFFF; --soft:#F7F8FA; --shadow:0 8px 24px rgba(17,24,39,.06);
}}
/* 시안 전용 배지/카드/스텝 헤더 */
.rr-badge {{
  display:inline-flex; align-items:center; gap:.5rem;
  background:var(--brand); color:#fff; padding:.55rem 1rem; border-radius:999px;
  box-shadow: var(--shadow); font-weight:800; letter-spacing:.2px;
}}
.rr-kpi {{
  background:#fff; border:1px solid #EEF2F7; border-radius:16px; padding:18px 22px; box-shadow:var(--shadow);
}}
.rr-kpi .title {{ color:var(--muted); font-weight:800; }}
.rr-kpi .value {{ font-weight:900; font-size:2rem; color:var(--brand); }}
.rr-card {{
  background:#fff; border:1px solid #EEF2F7; border-radius:22px; padding:18px; box-shadow:var(--shadow);
}}
.rr-step-head {{
  display:flex; align-items:center; gap:.6rem; font-weight:900; font-size:1.2rem; color:var(--ink);
}}
.rr-step-chip {{
  background:#fff; border:2px solid var(--brand); color:var(--brand);
  padding:.25rem .7rem; border-radius:999px; font-weight:900; font-size:.85rem;
}}
.rr-rule {{
  width:100%; height:2px; margin:6px 0 12px 0;
  background:linear-gradient(90deg, var(--brand), rgba(238,0,0,.15));
  border-radius:999px;
}}
/* KPI를 좌(라벨) ↔ 우(숫자)로 배치 */
.rr-kpi .row{{display:flex;align-items:center;justify-content:space-between;gap:.75rem;}}
.rr-kpi .label{{font-weight:800;color:#111827;}}
.rr-kpi .num{{font-weight:900;font-size:2rem;color:var(--brand);}}
.rr-kpi .sub{{margin-top:.35rem;color:#374151;font-weight:600;}} 
.rr-kpi .num small{{font-weight:700;}}

.hero-wrap{{margin:8px 0 14px 0;}}
.hero-pill{{
  position:relative; display:inline-block; background:#fff;
  border:1px solid #EA002C; border-radius:16px; padding:8px 16px;
  font-weight:900; font-size:26px;
}}
.hero-pill:after{{
  content:""; position:absolute; left:8px; right:8px; bottom:-10px;
  height:22px; background:#EA002C; border-radius:16px; z-index:-1;
}}
.hero-meta{{color:#9CA3AF; font-weight:700; margin-top:6px;}}
.hero-tip{{margin-top:6px; font-weight:800;}}
.rr-step-text {{ font-size:16px; line-height:1.9; }}
.rr-step-text p,
.rr-step-text li {{ font-size:16px; line-height:1.9; }}
.rr-step-text h3 {{ font-size:18px; margin:10px 0 6px; }}  /* 섹션 제목도 살짝 키움(선택) */
/* 강점/평균권/약점 범례 글씨 크기 줄이기 */
.sw-legend, .sw-legend .item{{
  font-size: 20px;      /* ← 원하는 크기(예: 12px, 11px, 10px) */
  line-height: 1.1;
}}
/* 점 크기도 같이 줄이고 싶으면 */
.sw-legend .dot{{ width: 15px; height: 15px; }}  /* 기존이 10px이면 축소 */
</style>
""", unsafe_allow_html=True)


# 카드 여닫이 유틸 (시안 카드 감싸기)
def rr_open():  st.markdown("<div class='rr-card'>", unsafe_allow_html=True)
def rr_close(): st.markdown("</div>", unsafe_allow_html=True)






























# 사이드바 메뉴
with st.sidebar:
    import os, streamlit as st

    logo_top = os.path.join(os.path.dirname(__file__), "sk_logo_top.png")
    if os.path.exists(logo_top):
        st.image(logo_top, use_container_width=True)  # ← 변경!
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    logo1 = os.path.join(os.path.dirname(__file__), "sk_logo.png")
    if os.path.exists(logo1):
        st.image(logo1, use_container_width=True)     # ← 변경!
    st.markdown("## 메뉴")
    menu = st.radio("메뉴", ["📊 팀장 대시보드 생성"])
    st.markdown("---")
  
    








# ───────── 헤더 ─────────
st.markdown(
    """
    <h1 style="text-align:center; color:#EA002C; font-weight:800; margin: 6px 0 18px;">
      SK Red Re:Born
    </h1>
    """,
    unsafe_allow_html=True,
)




# (선택) 설명/사용법
c_exp1, c_exp2 = st.columns([1,1])




with c_exp1:
    with card("설명"):
        st.markdown('<div class="card-title"><b>설명</div>', unsafe_allow_html=True)
        st.markdown(
        '<div class="card-sub">서베이 데이터를 업로드한 뒤 소속을 선택하면<br>'
        '리더십 현황과 개선방안을 한 눈에 확인할 수 있는 리더십 대시보드가 제공됩니다.</div>',
        unsafe_allow_html=True
        )


with c_exp2:
    with card("사용법"):
        st.markdown('<div class="card-title"><b>사용법</div>', unsafe_allow_html=True)
        st.markdown(
        '<div class="card-sub">'
        '① 좌측의 서베이 파일과 교육 프로그램(선택) 칸에서 &#39;Browse files&#39; 클릭하여 파일 업로드<br>'
        '② 업로드 후 우측에서 조직 범위 선택<br>'
        '③ 팀장별 PDF 생성 또는 실 단위 ZIP 다운로드'
        '</div>',
        unsafe_allow_html=True
        )




# ───────── 1행: (좌) 서베이 업로드 / (우) 검색 ─────────
c1, c2 = st.columns([1,1], gap="large")




with c1:
    survey_icon = os.path.join(os.path.dirname(__file__), "survey_upload.png")
    right_icon  = os.path.join(os.path.dirname(__file__), "upload.png")
    uploaded_file = upload_card(
    title="서베이 파일 업로드",
    subtitle="PDF 또는 Excel/CSV 파일을 업로드하세요.",
    icon_path=survey_icon,
    right_icon_path=right_icon,      # ← 여기!
    key="survey_upload",
    types=["pdf","xlsx","csv"]
    )




    # ↓↓↓ 기존 처리 로직 그대로
    extracted_text = ""
    if uploaded_file and uploaded_file.name.endswith(".pdf"):
        with pdfplumber.open(uploaded_file) as pdf:
            extracted_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
        st.success("PDF 텍스트 추출 완료")
    elif uploaded_file and uploaded_file.name.endswith((".xlsx","csv")):
        raw_df = (pd.read_excel(uploaded_file) if uploaded_file.name.endswith(".xlsx")
                    else pd.read_csv(uploaded_file))
        raw_df["보임연도"] = pd.to_datetime(raw_df["팀장 보임일"], errors="coerce").dt.year
        raw_df["연도"]    = pd.to_numeric(raw_df["연도"], errors="coerce")
        raw_df = raw_df[raw_df["보임연도"] <= raw_df["연도"]].copy()
        raw_df = attach_leader_key(raw_df)          # 리더키 부여
        st.session_state.raw_df = raw_df
        st.session_state.org_df = raw_df.copy()
        try:
            df7 = load_and_prepare(uploaded_file)
            df7["연도"] = raw_df["연도"]
            st.session_state.survey_df = attach_leader_key(df7)
            st.success("✅ 데이터 전처리 및 로드 완료")
        except Exception as e:
            st.error(f"데이터 전처리 오류: {e}")




with c2:
    with card("검색"):
        _icon(ICONS["search"])
        st.markdown('<div class="card-title">아래에서 검색 해주세요</div>', unsafe_allow_html=True)
        if st.session_state.get("org_df") is None:
            st.info("먼저 서베이 파일을 업로드하세요.")
        else:
            org_df = st.session_state.org_df
            companies = ["전체"] + sorted(org_df["회사명"].dropna().unique())
            sel_comp = st.selectbox("회사명 선택", companies, key="sel_comp")
            if sel_comp != "전체":
                org_df = org_df[org_df["회사명"] == sel_comp]




            hqs = ["전체"] + sorted(org_df["본부"].dropna().unique())
            sel_hq = st.selectbox("본부 선택", hqs, key="sel_hq")
            if sel_hq != "전체":
                org_df = org_df[org_df["본부"] == sel_hq]




            depts = ["전체"] + sorted(org_df["실"].dropna().unique())
            sel_dept = st.selectbox("실 선택", depts, key="sel_dept")
            if sel_dept != "전체":
                org_df = org_df[org_df["실"] == sel_dept]




            # 범위 변경 시 ZIP/개별 캐시 초기화
            new_scope = (st.session_state.get("sel_comp","전체"),
                         st.session_state.get("sel_hq","전체"),
                         st.session_state.get("sel_dept","전체"))
            if st.session_state.get("dl_scope") != new_scope:
                st.session_state["dl_scope"] = new_scope
                st.session_state.pop("bulk_zip_bytes", None)
                for k in list(st.session_state.keys()):
                    if k.startswith("pdf_bytes_"): st.session_state.pop(k)




            st.session_state.filtered_org_df = org_df.copy()




st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)




# ───────── 2행: (좌) 교육 프로그램 업로드 / (우) 해당 실 팀원 목록 ─────────
c3, c4 = st.columns([1,1], gap="large")




with c3:
    edu_icon = os.path.join(os.path.dirname(__file__), "edu_upload.png")
    right_icon  = os.path.join(os.path.dirname(__file__), "upload.png")
    edu_file = upload_card(
    title="교육 프로그램 업로드",
    subtitle="교육 프로그램 목록(Excel) 파일을 업로드하세요.",
    icon_path=edu_icon,
    right_icon_path=right_icon,      # ← 여기!
    key="edu_upload",
    types=["xlsx"]
    )
    if edu_file:
        st.session_state.edu_file = edu_file
        st.success("교육 프로그램 DB 업로드 완료")




with c4:
    with card("팀원 목록"):
        st.markdown('<div class="card-title">해당 실의 팀과 팀원 목록</div>', unsafe_allow_html=True)
        if st.session_state.get("filtered_org_df") is None:
            st.info("오른쪽 카드에서 조직을 먼저 선택하세요.")
        else:
            display_df = (
                st.session_state.filtered_org_df[[LEADER_KEY_COL, "팀","이름"]]
                .drop_duplicates(subset=[LEADER_KEY_COL])
                .sort_values("팀")
                .reset_index(drop=True)
            )
            st.dataframe(hide_ids(display_df[["팀","이름"]]), use_container_width=True)




# ───────── 업로드/검색 카드 바로 아래: ZIP & 팀장별 다운로드 ─────────
st.markdown("---")
st.header("대시보드")




if st.session_state.get("survey_df") is None:
    st.info("먼저 서베이 파일을 업로드하세요.")
else:
    # 전사(미필터) 원본은 따로 보관
    survey_df_global = st.session_state.survey_df.copy()
    raw_df_global    = st.session_state.raw_df.copy()

    # 화면/범위 표시용은 아래에서 필터링
    survey_df = survey_df_global.copy()
    raw_df    = raw_df_global.copy()



    if st.session_state.get("filtered_org_df") is not None:
        allowed_keys = st.session_state.filtered_org_df[LEADER_KEY_COL].dropna().unique()
        survey_df = survey_df[survey_df[LEADER_KEY_COL].isin(allowed_keys)]
        raw_df_disp = raw_df[raw_df[LEADER_KEY_COL].isin(allowed_keys)]
    else:
        raw_df_disp = raw_df




    # 팀장 목록(리더키 기준)
    leaders_df = (
        raw_df_disp[[LEADER_KEY_COL, LEADER_NAME_COL, "회사명","본부","실","팀"]]
        .dropna(subset=[LEADER_KEY_COL])
        .drop_duplicates(subset=[LEADER_KEY_COL])
        .sort_values([LEADER_NAME_COL, "회사명","본부","실","팀"])
    )




    # 공통 집계 (ZIP/개별 공용)
    # 전사(미필터) 기준
    survey_df2 = survey_df_global.copy()
    survey_df2["Composite"] = survey_df2[VISUAL_COLS].mean(axis=1)
    overall_trend_corp = survey_df2.groupby("연도")["Composite"].mean().sort_index()

    base_df_global = survey_df_global[survey_df_global["연도"] == FOCUS_YEAR]
    if base_df_global.empty: base_df_global = survey_df_global
    leader_avg_df_global = base_df_global.groupby(LEADER_KEY_COL)[VISUAL_COLS].mean()
    overall_mean_corp   = leader_avg_df_global.mean()
    overall_series_corp = {col: leader_avg_df_global[col] for col in VISUAL_COLS}




    # 리더별 최저영역 라벨(팀장 PDF 전달용) - 전사(미필터) 기준
    if leader_avg_df_global.empty:
        st.session_state['weakest_label_by_key'] = {}
    else:
        leader_min_col = leader_avg_df_global.idxmin(axis=1)   # 각 리더의 최저 카테고리 키(예: '팀원_상호배려')
        weakest_label_by_key = leader_min_col.map(LABEL_MAP)   # 표시 라벨(예: '상호배려')
        st.session_state['weakest_label_by_key'] = weakest_label_by_key.dropna().to_dict()




    left, right = st.columns([1,1], gap="large")




    # ── 좌: 실 단위 ZIP ──
    with left:
        st.subheader("실 단위 대시보드")

        # 준비 여부(회사/본부/실 선택)
        ready = st.session_state.get("sel_dept","전체") != "전체"
        if not ready:
            st.info("회사/본부/실을 모두 선택하면 활성화됩니다.")

        # 이미지 경로 → data URI
        base_dir = os.path.dirname(__file__)
        zip_src  = _img_data_uri(os.path.join(base_dir, "ZIP.png"))
        dash_src = _img_data_uri(os.path.join(base_dir, "dashboard.png"))
        dl_src   = _img_data_uri(os.path.join(base_dir, "download.png"))

        # 두 개 카드 나란히
        p1, p2 = st.columns(2, gap="large")

        # ----------------------- 만들기 카드 -----------------------
        with p1:
            st.markdown(
                f"""
                <div class="zip-card">
                <img src="{zip_src}"  class="zip-top-icon" />
                <div class="zip-title">전체 ZIP<br>만들기</div>
                <img src="{dash_src}" class="zip-btm-icon" />
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button("전체 ZIP 만들기", disabled=not ready, type="primary",
                        use_container_width=True, key="make_zip_btn"):
                with st.spinner("전체 대시보드 ZIP 생성 중..."):
                    org_scope = " / ".join([s for s in [
                        st.session_state.get("sel_comp"),
                        st.session_state.get("sel_hq"),
                        st.session_state.get("sel_dept"),
                    ] if s and s != "전체"]) or None

                    keys = leaders_df[LEADER_KEY_COL].tolist()
                    zip_bytes = make_zip_for_leaders(
                        keys, survey_df_global, overall_trend_corp,
                        overall_mean_corp, overall_series_corp,
                        include_llm_text=True,
                        edu_file=st.session_state.get("edu_file"),
                        org_scope=org_scope,
                        weakest_label_by_key=st.session_state.get('weakest_label_by_key')
                    )
                    st.session_state["bulk_zip_bytes"] = zip_bytes
                st.toast("ZIP 생성 완료! ✅")

        # ----------------------- 다운로드 카드 -----------------------
        with p2:
            st.markdown(
                f"""
                <div class="zip-card">
                <img src="{zip_src}" class="zip-top-icon" />
                <div class="zip-title">전체 ZIP<br>다운로드</div>
                <img src="{dl_src}"  class="zip-btm-icon" />
                </div>
                """,
                unsafe_allow_html=True
            )

            if "bulk_zip_bytes" in st.session_state and ready:
                st.download_button(
                    "전체 ZIP 다운로드",
                    data=st.session_state["bulk_zip_bytes"],
                    file_name="대시보드_일괄.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="zip_dl_btn",
                )
            else:
                # 자리 고정용 비활성 버튼
                st.button("전체 ZIP 다운로드", disabled=True,
                        use_container_width=True, key="zip_dl_btn_disabled")




    # ── 우: 팀장별 생성/다운로드 ──
    with right:
        _icon(ICONS["list"])
        st.subheader("팀장별 다운로드")
        if not ready:
            st.info("회사/본부/실을 모두 선택하면 표시됩니다.")
        elif leaders_df.empty:
            st.warning("선택된 조직에 팀장이 없습니다.")
        else:
            for _, row in leaders_df.iterrows():
                k   = row[LEADER_KEY_COL]
                nm  = row[LEADER_NAME_COL]
                team= str(row.get("팀",""))
                cL, cR = st.columns([3,2])
                with cL:
                    st.markdown(f"○ **{nm}** <span style='color:#6B7280'>/ {team}</span>", unsafe_allow_html=True)
                with cR:
                    # 두 개의 하위 컬럼: [생성] [다운로드]
                    cGen, cDl = st.columns([1, 1])

                    gen_key = f"gen_{k}"
                    dl_key  = f"dl_{k}"
                    state_key = f"pdf_bytes_{k}"

                    # 1) 생성 버튼 (왼쪽)
                    with cGen:
                        if st.button("PDF 생성", key=gen_key, disabled=not ready, use_container_width=True):
                            with st.spinner(f"{nm} PDF 생성 중..."):
                                org_scope = " / ".join([s for s in [
                                    st.session_state.get("sel_comp"),
                                    st.session_state.get("sel_hq"),
                                    st.session_state.get("sel_dept"),
                                ] if s and s != "전체"]) or None

                                wk = (st.session_state.get('weakest_label_by_key') or {}).get(k)
                                name, pdf = make_pdf_for_leader(
                                    k, survey_df_global, overall_trend_corp, overall_mean_corp, overall_series_corp,
                                    include_llm_text=True,
                                    edu_file=st.session_state.get("edu_file"),
                                    weak_text=(f"가장 낮은 영역: {wk}" if wk else None),
                                    org_scope=org_scope
                                )
                                st.session_state[state_key] = (name, pdf)

                            # 레이아웃 밀리지 않게 토스트만 띄우고 리렌더
                            st.toast("생성 완료! ✅")
                            st.rerun()

                    # 2) 다운로드 버튼 (오른쪽)
                    with cDl:
                        if state_key in st.session_state:
                            name, pdf = st.session_state[state_key]
                            st.download_button(
                                "PDF 다운로드",
                                data=pdf,
                                file_name=f"{_safe_filename(name)}_dashboard.pdf",
                                mime="application/pdf",
                                key=dl_key,
                                use_container_width=True
                            )
                        else:
                            # 아직 없을 때는 자리 고정용 비활성 버튼
                            st.button("PDF 다운로드", key=f"{dl_key}_disabled", disabled=True, use_container_width=True)








    # ============================================================================
    # 팀장 선택(표시는 '이름 - 조직경로', 값은 리더키)
    st.subheader("👤 팀장 선택")




    def _fmt(k):
        row = leaders_df.loc[leaders_df[LEADER_KEY_COL] == k].iloc[0]
        org_bits = [str(row.get(c, "")) for c in ["회사명","본부","실","팀"] if pd.notna(row.get(c, "")) and str(row.get(c, ""))]
        org_path = "/".join(org_bits)
        return f'{row[LEADER_NAME_COL]} - {org_path}' if org_path else f'{row[LEADER_NAME_COL]}'




    leader_keys = leaders_df[LEADER_KEY_COL].tolist()


    selected_key = st.selectbox(
        "리포트를 생성할 팀장을 선택하세요",
        options=leader_keys,
        format_func=_fmt,
        index=None,                         # ★ 기본 미선택
        placeholder="팀장을 선택하세요",      # ★ 자리표시자
        key="selected_leader",
        disabled=leaders_df.empty           # 선택지 없으면 비활성
    )


    # 아직 선택 전이면 아래 대시보드 생성 로직 중단
    if selected_key is None:
        st.info("팀장을 선택하면 대시보드가 표시됩니다.")
        st.stop()




    # 선택된 리더 데이터(지표 계산은 survey_df, 원본 표시는 raw_df_for_display)
    sel = survey_df[survey_df[LEADER_KEY_COL] == selected_key].copy()




    # 2024만 사용 (없으면 전체로 폴백)
    sel_y = sel[sel["연도"] == FOCUS_YEAR]
    if sel_y.empty:
        sel_y = sel
    
    selected_name = sel[LEADER_NAME_COL].iloc[0] if not sel.empty else ""

    meta_row = raw_df_disp[raw_df_disp[LEADER_KEY_COL] == selected_key].iloc[0]
    def _get(*cols):
        for c in cols:
            if c in meta_row and pd.notna(meta_row[c]) and str(meta_row[c]).strip():
                return str(meta_row[c])
        return ""


    # 합성 점수 및 3개년 추이
    sel['Composite'] = sel[VISUAL_COLS].mean(axis=1)
    trend = sel.groupby('연도')['Composite'].mean().sort_index()






    # 전체(분모)도 리더키 기준으로 집계 (동명이인 섞임 방지)
    # 전사 기준으로 다시 산출
    base_df = survey_df_global[survey_df_global["연도"] == FOCUS_YEAR]
    if base_df.empty: base_df = survey_df_global
    leader_avg_df = base_df.groupby(LEADER_KEY_COL)[VISUAL_COLS].mean()
    overall_mean   = leader_avg_df.mean()
    overall_series = {col: leader_avg_df[col] for col in VISUAL_COLS}




    # 평균/레이더 준비
    team_24, overall_24 = get_scores_for_radar(sel, survey_df, RADAR_COMPARE_YEAR)
    score_dict = { LABEL_MAP[c]: float(team_24[c]) for c in VISUAL_COLS }
    avg_scores = sel_y[VISUAL_COLS].mean()




    # 가장 낮은 카테고리 + 프로그램 추천 인트로
    weakest_col   = avg_scores.idxmin()
    weakest_label = LABEL_MAP[weakest_col]
    weakest_score = round(float(avg_scores[weakest_col]), 2)




    intro_sentence = (
    f"현재 가장 개선해야 할 리더십 영역은 **{weakest_label}**입니다. "
    "이 영역을 중심으로 개인 행동, 역량 강화, 조직 실행의 3단계 실천안을 제시합니다."
    )




    # HERO
    # 선택된 리더 메타
    _id  = _get(LEADER_ID_COL)
    _comp= _get('회사명')
    _team= _get('팀')
    _pos = _get('직위','직책','직급')

    st.markdown(
        f"""
        <div class="hero-wrap" style="text-align:center">
        <div class="hero-pill">
            {selected_name} 팀장의 <span style="color:#EA002C">Red Re:born</span> 
        </div>
        <div class="hero-meta">ID: {_id} · 회사: {_comp} · 팀: {_team} · 직위:{_pos}</div>
        <div class="hero-tip">좌측 정보를 모두 읽은 후 단계별 개선방안을 확인해주세요 :)</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("---")
    
    # === NEW: 리더십 7개 영역 평균 총점 (100점 기준) ===
    # 선택된 팀장 데이터(sel)가 이미 있음
    # 1~5 평균 → 0~100 환산: (x-1)/4*100
    sel_comp_mean_1to5 = sel_y[VISUAL_COLS].mean(axis=1).mean()
    total_100 = round(to_percent(float(sel_y[VISUAL_COLS].mean().mean())), 1)




    # === HERO (헤더 좌/우) === (여기 위쪽은 기존 그대로 두세요)
 
    st.markdown("---")

    # ▼▼▼ PATCH B: 선택 팀장 대시보드 레이아웃 교체 ▼▼▼

    # (보조) rr 카드 열고/닫기 (PATCH A 미적용 환경 대비)
    if 'rr_open' not in globals():
        def rr_open():  st.markdown("<div class='rr-card'>", unsafe_allow_html=True)
        def rr_close(): st.markdown("</div>", unsafe_allow_html=True)

    # 1) 공통 준비값
    # - 범위(현재 필터) 전사 평균 추이
    survey_df_with_comp = survey_df_global.copy()
    survey_df_with_comp["Composite"] = survey_df_with_comp[VISUAL_COLS].mean(axis=1)
    overall_trend_corp2 = survey_df_with_comp.groupby("연도")["Composite"].mean().sort_index()

    # - 2024 평균(1~5), 최저 영역/점수
    avg_scores      = sel_y[VISUAL_COLS].mean()
    year_avg_1to5   = float(sel_y[VISUAL_COLS].mean().mean()) if not sel_y.empty else float(sel[VISUAL_COLS].mean().mean())
    weakest_col     = avg_scores.idxmin()
    weakest_label   = LABEL_MAP[weakest_col]
    weakest_score   = float(avg_scores[weakest_col])

    # - STEP1 생성
    strengths_txt      = _format_sw_text(avg_scores)[0]
    weaknesses_txt     = _format_sw_text(avg_scores)[1]
    score_summary_txt  = _make_score_summary(avg_scores, overall_mean)
    subj_strength_txt, subj_weak_txt = get_subjectives_for_leader(st.session_state.raw_df, selected_key)
    step1_content = generate_step1_md(
        name=selected_name,
        strengths=strengths_txt,
        weaknesses=weaknesses_txt,
        score_summary=score_summary_txt,
        weak_text=weakest_label,
        subjective_weak=subj_weak_txt
    )
    # 헤더 제거(🌱… 라인 제거)
    step1_body = re.sub(r'^\s*#{1,6}\s*🌱[^\n]*\n+', '', step1_content, flags=re.M)

    # - STEP2 생성
    trend_str = "\n".join(f"{y}년: {s:.2f}점" for y, s in trend.items())
    step2_md  = build_step2_with_recos(
        trend_str, weakest_label, st.session_state.get("edu_file"),
        subjective_weak=subj_weak_txt
    )
    # 헤더 제거(🧭… 라인 제거)
    step2_body = re.sub(r'^\s*#{1,6}\s*🧭[^\n]*\n+', '', step2_md)

    # - STEP3 생성
    org_scope = " / ".join([s for s in [
        st.session_state.get("sel_comp"),
        st.session_state.get("sel_hq"),
        st.session_state.get("sel_dept"),
    ] if s and s != "전체"]) or None
    step3_md = generate_step3_md(
        weak_text=weakest_label, org_scope=org_scope, subjective_weak=subj_weak_txt
    )
    # 헤더 제거(🪄… 라인 제거)
    step3_body = re.sub(r'^\s*#{1,6}\s*🪄\s*[^\n]*\n+', '', step3_md)

    # ===== TOP KPIs (시안 타일 2개) =====
    k1, k2 = st.columns([1,1])
    # 5점 평균 → 100점 환산
    year_avg_100 = round(to_percent(year_avg_1to5), 1)

    # ① 왼쪽 KPI: "2024년 리더십 평균 점수" ← 왼쪽, 점수는 오른쪽
    with k1:
        # 이미 위에서 year_avg_100 = round(to_percent(year_avg_1to5), 1) 계산됨
        st.markdown(
            f"""
            <div class='rr-kpi'>
            <div class='row'>
                <div class='label'>2024년 리더십 평균 점수 (100점 기준)</div>
                <div class='num'>{year_avg_100:.1f}<small>/100</small></div>
            </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # ② 오른쪽 KPI: 학습 권장 영역 + 안내 문구
    with k2:
        weakest_100 = round(to_percent(weakest_score), 1)
        st.markdown(
            f"""
            <div class='rr-kpi'>
            <div class='row'>
                <div class='label'>학습 권장 영역</div>
                <div class='num'>
                    <span style="color:{PRIMARY}">{weakest_label}</span>
                </div>
            </div>
            <div class='sub'>단계별 개선방안 1→2→3 단계를 수행해볼까요?</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<div class='rr-rule'></div>", unsafe_allow_html=True)

    with st.spinner("3개년 추이 분석 중..."):
        trend_comment = make_trend_commentary_via_llm_from_series(
        trend, weakest_label=weakest_label
    ).strip()
    # ===== Row 1: 3개년 추이 | STEP 1 =====
    row1L, row1R = st.columns([1,1], gap="large")
    with row1L:
        st.markdown("<span class='rr-badge'>3개년 리더십 추이</span>", unsafe_allow_html=True)
        rr_open()

        # 안전 가드: Composite 없으면 만들어주기
        if "Composite" not in sel.columns:
            sel["Composite"] = sel[VISUAL_COLS].mean(axis=1)

        # ✅ 그래프 호출(이게 빠져서 안 보였던 거예요)
        _ = plot_yearly_bar_with_avg(
            sel_df=sel,
            overall_trend=overall_trend_corp2,   # 위에서 만든 전사 평균 추이 Series
            year_col="연도",
            score_col="Composite",
            all_years=(2022, 2023, 2024),
            y_tick_step=10
        )
        rr_close()

        # 코멘트
        st.markdown(
            "<div style='margin:10px 4px 4px; font-size:16px; color:#111827; line-height:1.7; font-weight:500;'>"
            + "".join([f"<div>• {ln.strip()}</div>" for ln in trend_comment.splitlines() if ln.strip()])
            + "</div>",
            unsafe_allow_html=True
        )
        
        # △/▽ 요약(안내문 제거 버전)
        rf_res   = calc_biggest_rise_fall(sel)
        rf_items = format_rise_fall_items(rf_res)         # note 항목 제거된 함수
        if rf_items:
            st.markdown(render_rise_fall_html(rf_items), unsafe_allow_html=True)

    with row1R:
        st.markdown("<div class='rr-step-head'><span class='rr-step-chip'>STEP 1</span> 🌱 즉각적인 개인 실천의 시작</div>", unsafe_allow_html=True)
        rr_open()
        st.markdown(step1_body) 
        rr_close()

    st.markdown("<div class='rr-rule'></div>", unsafe_allow_html=True)

    # ===== Row 2: 레이더 | STEP 2 =====
    row2L, row2R = st.columns([1,1], gap="large")
    with row2L:
        st.markdown("<span class='rr-badge'>종합 다이어그램</span>", unsafe_allow_html=True)
        rr_open()
        # 비교 기준 점수 재계산(안전)
        team_24, overall_24 = get_scores_for_radar(sel, survey_df_global, RADAR_COMPARE_YEAR)
        radar_fig = plot_radar_compare_7(
            sel_score={ LABEL_MAP[c]: float(team_24[c]) for c in VISUAL_COLS },
            ref_score={ LABEL_MAP[c]: float(overall_24[c]) for c in VISUAL_COLS },
            title=f"{RADAR_COMPARE_YEAR}년 기준",
            legend_labels=(selected_name, "전사 평균"),
            show_values=False
        )
        rr_close()

    with row2R:
        st.markdown("<div class='rr-step-head'><span class='rr-step-chip'>STEP 2</span> 🧭 체계적인 교육을 통한 역량 강화</div>", unsafe_allow_html=True)
        rr_open()
        st.markdown(textwrap.dedent(step2_body))
        rr_close()

    st.markdown("<div class='rr-rule'></div>", unsafe_allow_html=True)

    # ===== Row 3: 게이지 | STEP 3 =====
    row3L, row3R = st.columns([1,1], gap="large")
    with row3L:
        st.markdown("<span class='rr-badge'>강점 & 약점</span>", unsafe_allow_html=True)
        render_sw_legend_streamlit()
        rr_open()
        _figs, _meta = render_strength_weakness(
            avg_scores, overall_mean, overall_series,
            leader_key=selected_key, raw_df=st.session_state.raw_df,
            return_figs=True, return_meta=True
        )
        rr_close()

    with row3R:
        st.markdown("<div class='rr-step-head'><span class='rr-step-chip'>STEP 3</span> 🪄 팀/조직 차원의 변화</div>", unsafe_allow_html=True)
        rr_open()
        st.markdown(textwrap.dedent(step3_body)) 
        rr_close()



# 실행 안내
# streamlit run ax4_final.py
