# -*- coding: utf-8 -*-
"""
web_app.py — 무역 리스크 분석 + 제품 시뮬레이터 (통합 단일 파일)
실행: python web_app.py  →  http://127.0.0.1:5000 자동 오픈
필요: apikey.txt (Gemini), kotra_key.txt (KOTRA)
"""

import os, json, re, ssl, time, hashlib, threading, webbrowser, io, math
import requests
from datetime import datetime
from flask import Flask, request, Response, render_template_string, jsonify, stream_with_context

try:
    import openpyxl
    _HAS_OPENPYXL = True
except Exception:
    _HAS_OPENPYXL = False

try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except Exception:
    pass
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except Exception:
    pass

try:
    from bs4 import BeautifulSoup
    _HAS_BS4 = True
except Exception:
    _HAS_BS4 = False

try:
    import chromadb
    from chromadb.utils import embedding_functions
    _HAS_CHROMA = True
except Exception:
    _HAS_CHROMA = False

try:
    from google import genai
    from google.genai import types as genai_types
    _HAS_GEMINI = True
except Exception:
    genai = None; genai_types = None; _HAS_GEMINI = False

try:
    import pandas as pd
    from io import StringIO as _SIO
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = False

# 조원 엔진: 산업연관(레온티에프) 가격파급 + AHP/z-score 스코어링
try:
    import scoring as _scoring
    import io_analysis as _io_analysis
    _HAS_ENGINES = True
except Exception as _e:
    _scoring = None; _io_analysis = None; _HAS_ENGINES = False
    print(f"[engines] scoring/io_analysis 로드 실패: {_e}")

# 공공데이터 K-SURE 국별신용등급 + 백테스트(실제 사례 검증)
try:
    import data_sources as _data_sources
except Exception as _e:
    _data_sources = None; print(f"[engines] data_sources 로드 실패: {_e}")
try:
    import backtest as _backtest
except Exception as _e:
    _backtest = None; print(f"[engines] backtest 로드 실패: {_e}")


# ════════════════════════════════════════════════════════════
#  설정 & 상수
# ════════════════════════════════════════════════════════════
FX_BASELINE = 1330.0
HS_BASE     = "https://www.hs-tariff.com/main/hs_mti_ai_main"
UA          = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
               "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
KOTRA_NEWS_ENDPOINT    = "https://apis.data.go.kr/B410001/kotra_overseasMarketNews/ovseaMrktNews/ovseaMrktNews"
KOTRA_COUNTRY_ENDPOINT = "https://apis.data.go.kr/B410001/kotra_nationalInformation/natnInfo/natnInfo"
# 무료 할당량 넉넉한 모델 우선. 2.5-pro는 무료 0이라 제외.
# 각 모델은 일일 무료 quota가 따로라, 여러 개 두면 하루 총 처리량이 늘어남.
GEMINI_MODELS = ["gemini-2.0-flash", "gemini-2.0-flash-lite",
                 "gemini-2.5-flash", "gemini-1.5-flash"]

COUNTRY_CURRENCY = {
    "미국":       ("USD", "달러",           1330.0, 1),
    "중국":       ("CNY", "위안",            185.0, 1),
    "일본":       ("JPY", "엔",                9.0, 2),
    "독일":       ("EUR", "유로",            1450.0, 1),
    "프랑스":     ("EUR", "유로",            1450.0, 1),
    "이탈리아":   ("EUR", "유로",            1450.0, 1),
    "스페인":     ("EUR", "유로",            1450.0, 1),
    "네덜란드":   ("EUR", "유로",            1450.0, 1),
    "베트남":     ("VND", "동",               0.053, 4),
    "인도":       ("INR", "루피",              16.0, 2),
    "호주":       ("AUD", "호주달러",         870.0, 1),
    "캐나다":     ("CAD", "캐나다달러",       970.0, 1),
    "영국":       ("GBP", "파운드",          1680.0, 1),
    "브라질":     ("BRL", "헤알",             260.0, 1),
    "멕시코":     ("MXN", "페소",              75.0, 1),
    "러시아":     ("RUB", "루블",              14.5, 2),
    "태국":       ("THB", "바트",              37.0, 1),
    "인도네시아": ("IDR", "루피아",            0.085, 4),
    "말레이시아": ("MYR", "링깃",             300.0, 1),
    "싱가포르":   ("SGD", "싱가포르달러",     990.0, 1),
    "대만":       ("TWD", "신대만달러",        41.0, 1),
    "사우디":     ("SAR", "리얄",             354.0, 1),
    "UAE":        ("AED", "디르함",           362.0, 1),
    "터키":       ("TRY", "리라",              40.0, 2),
    "폴란드":     ("PLN", "즈워티",           330.0, 1),
}

OFFLINE_DATA = [
    ("Gold", 2320.5, 0.31), ("Silver", 27.8, -0.52), ("Copper", 4.52, 0.83),
    ("Aluminum", 2450.0, -0.88), ("Iron Ore", 105.0, -1.12), ("Nickel", 16800.0, 0.55),
    ("Zinc", 2780.0, -0.33), ("Platinum", 952.0, 0.37), ("Crude Oil", 78.2, 1.24),
    ("Natural Gas", 2.15, -2.10), ("Wheat", 545.0, -1.30), ("Corn", 430.0, 0.62),
    ("Steel", 480.0, 0.50), ("Coal", 145.0, -0.80), ("Rubber", 1.65, 0.40),
]

COMMODITY_PROFILES = {
    "copper":   {"linkage": 0.92, "tariff": 3.1, "logistics": 0.78, "concentration": 0.82},
    "aluminum": {"linkage": 0.88, "tariff": 5.2, "logistics": 0.70, "concentration": 0.75},
    "steel":    {"linkage": 0.90, "tariff": 6.8, "logistics": 0.72, "concentration": 0.70},
    "iron":     {"linkage": 0.89, "tariff": 2.5, "logistics": 0.80, "concentration": 0.85},
    "nickel":   {"linkage": 0.86, "tariff": 4.5, "logistics": 0.76, "concentration": 0.88},
    "gold":     {"linkage": 0.35, "tariff": 0.5, "logistics": 0.92, "concentration": 0.40},
    "silver":   {"linkage": 0.42, "tariff": 0.8, "logistics": 0.88, "concentration": 0.45},
    "oil":      {"linkage": 0.95, "tariff": 8.0, "logistics": 0.85, "concentration": 0.90},
    "gas":      {"linkage": 0.93, "tariff": 3.0, "logistics": 0.82, "concentration": 0.88},
}
DEFAULT_PROFILE = {"linkage": 0.55, "tariff": 4.0, "logistics": 0.60, "concentration": 0.50}

MATERIAL_ALIASES = {
    "철강": "steel", "철": "iron", "구리": "copper", "플라스틱": "plastic",
    "알루미늄": "aluminum", "금": "gold", "은": "silver", "니켈": "nickel",
    "아연": "zinc", "고무": "rubber", "강철": "steel", "동": "copper",
    "원유": "crude oil", "천연가스": "natural gas",
}

KO_EN = {
    "선풍기": "electric fan", "에어컨": "air conditioner", "냉장고": "refrigerator",
    "세탁기": "washing machine", "전기차": "electric vehicle", "자동차": "automobile",
    "제빙기": "ice maker", "냉동기": "freezer", "펌프": "pump", "모터": "motor",
    "컴프레서": "compressor", "밸브": "valve", "볼트": "bolt screw",
    "나사": "screw", "너트": "nut", "파이프": "steel pipe",
    "철판": "steel plate", "H빔": "H-beam steel", "각관": "square steel tube",
    "원형관": "circular steel tube", "스프링": "spring", "베어링": "bearing",
    "기어": "gear", "체인": "chain", "공구": "tool", "드릴": "drill",
    "플랜지": "flange", "엘보": "elbow fitting",
}

MANUFACTURING_CONTEXT = {
    "철광석":   "고로(BF) 생산비의 30~35%. 호주 53%/브라질 25% 편중, 해상운송 4~6주.",
    "유연탄":   "코크스용 필수, 고로 생산비 20~25%. 호주·캐나다 의존, EAF 전환 구조적 대안.",
    "철스크랩": "전기로(EAF) 생산비 60~70%. 국내 50%+수입 50%(일본·미국), 불순물 품질 리스크.",
    "후판":     "조선·건설용. 고로사 과점 → 전가력 높음. 조선 수주 사이클 연동.",
    "열연강판": "냉연·강관 중간소재. 중국 수출물량에 가격 민감, 반덤핑 빈발.",
    "니켈":     "STS 핵심 합금. LME 변동성 큼, 인니 수출정책 리스크.",
    "구리":     "전선·전장용. 경기 선행지표, LME 재고·중국 수요가 가격 좌우.",
}
_CTX_ALIAS = {
    "iron ore":"철광석","coal":"유연탄","코크스":"유연탄","scrap":"철스크랩",
    "고철":"철스크랩","plate":"후판","열연":"열연강판","hot rolled":"열연강판",
    "nickel":"니켈","copper":"구리","동":"구리",
    "철판":"후판","steel plate":"후판","steel":"후판","철강":"후판",
}
GYEONGSANG_CONTEXT = (
    "경상도(경남·경북·부산·울산)는 한국 철강·기계·자동차부품·조선·항공 수출 제조업 집적지. "
    "부산·울산항 의존도 높고 중소기업 비중 커 환헤지·무역금융 접근 한계. "
    "KOTRA 부산·경남·경북 지원단 공공망 활용."
)

# ── 경상도 철강·금속 제조 클러스터 (지역·주력항만·지원기관) ──
GYEONGSANG_REGIONS = {
    "포항": {"cluster":"포항철강산업단지(POSCO 연관)", "port":"포항 영일만항",
             "specialty":"철강 1차제품(열연·냉연·후판·선재·강관)",
             "support":"포항테크노파크 · KOTRA 대구경북지원단 · 경상북도 수출지원사업"},
    "울산": {"cluster":"울산 미포·온산국가산단", "port":"울산항",
             "specialty":"자동차·조선·석유화학·비철금속(구리·알루미늄)",
             "support":"울산테크노파크 · 울산경제진흥원 · KOTRA"},
    "창원": {"cluster":"창원국가산업단지", "port":"마산항·부산신항",
             "specialty":"기계·방산·자동차부품·밸브·플랜지",
             "support":"경남테크노파크 · KOTRA 경남지원단 · 경상남도 수출지원사업"},
    "김해": {"cluster":"김해 골든루트·테크노밸리", "port":"부산신항",
             "specialty":"중소 금속가공·정밀부품·주물단조",
             "support":"경남테크노파크 · 중소벤처기업진흥공단 경남"},
    "부산": {"cluster":"부산 녹산·신평장림국가산단", "port":"부산항(신항)",
             "specialty":"종합 금속·기계·조선기자재",
             "support":"부산경제진흥원 · KITA 부산본부 · KOTRA"},
    "거제": {"cluster":"거제 조선해양 클러스터", "port":"부산항·통영항",
             "specialty":"조선기자재·해양플랜트",
             "support":"경남테크노파크 · KOTRA 경남지원단"},
}
# 품목 키워드 → 경상도 대표 지역
_GS_PRODUCT_MAP = [
    (("열연","냉연","후판","철근","형강","선재","강판","철판","코일","철강","steel"), "포항"),
    (("강관","파이프","pipe","튜브","관"), "포항"),
    (("조선","선박","기자재","해양","ship","플랜트"), "거제"),
    (("자동차","차부품","엔진","변속","구동","베어링","샤프트"), "울산"),
    (("기계","밸브","펌프","플랜지","valve","공작기계","방산","감속기"), "창원"),
    (("주물","단조","볼트","너트","나사","금속가공","프레스","주조"), "김해"),
    (("구리","알루미늄","아연","니켈","비철","동","copper","aluminum"), "울산"),
]
# 경상도 주력 수출품 프리셋(원클릭)
GYEONGSANG_PRESETS = [
    "열연강판","냉연강판","후판","철근","강관","형강",
    "자동차부품","조선기자재","산업용밸브","플랜지","베어링","주단조품",
]

def gyeongsang_context(product):
    """품목 → 경상도 산단·주력항만·지원기관 매핑."""
    p = (product or "").lower()
    region = "포항"   # 철강 기본
    for kws, rg in _GS_PRODUCT_MAP:
        if any(k.lower() in p for k in kws):
            region = rg; break
    info = dict(GYEONGSANG_REGIONS[region]); info["region"] = region
    return info
_KOTRA_TITLE_KEYS = ["newsTitl","title","newsTitle","newsTtl","cntntsSj","sj","subject","titl"]
_KOTRA_URL_KEYS   = ["kotraNewsUrl","url","newsUrl","link","cntntsUrl","detailUrl","newsOrgnlUrl","orgnlUrl"]
_KOTRA_DATE_KEYS  = ["othbcDt","regDate","newsWrtDt","wrtDt","regDt","pubDate","crtDt","date","regYmd"]
_KOTRA_BODY_KEYS  = ["cntnt","contents","newsCn","cn","summary","description","txt","bdyText"]
_RELEVANCE_STOP   = frozenset("및 the of in to for a an".split())
_COUNTRY_MAP = {
    "베트남":{"vietnam","viet nam","vn"}, "미국":{"usa","u.s.","united states","america"},
    "중국":{"china","prc"}, "일본":{"japan","jp"}, "인도":{"india"},
    "독일":{"germany","deutschland"}, "태국":{"thailand"},
    "인도네시아":{"indonesia"}, "말레이시아":{"malaysia"},
}
_RI_GRADE = {"RI 1":"매우 낮음","RI 2":"낮음","RI 3":"보통","RI 4":"높음","RI 5":"매우 높음"}


# ════════════════════════════════════════════════════════════
#  유틸
# ════════════════════════════════════════════════════════════
def _match_manufacturing(material):
    if not material: return ""
    m = material.strip().lower()
    for key in MANUFACTURING_CONTEXT:
        if key in material: return MANUFACTURING_CONTEXT[key]
    for alias, key in _CTX_ALIAS.items():
        if alias in m: return MANUFACTURING_CONTEXT.get(key, "")
    return ""

def _material_keywords(material):
    if not material: return set()
    raw = material.strip()
    keys = {raw.lower(), raw.replace(" ","").lower()}
    m = raw.lower()
    for alias, canonical in _CTX_ALIAS.items():
        if alias in m or canonical in raw or raw in canonical:
            keys.add(alias.lower()); keys.add(canonical.lower())
    for key in MANUFACTURING_CONTEXT:
        if key in raw or raw in key: keys.add(key.lower())
    for tok in re.split(r"[\s,/·\-]+", raw):
        tok = tok.strip().lower()
        if len(tok) >= 2 and tok not in _RELEVANCE_STOP: keys.add(tok)
    steel_like = {"철판","후판","열연","열연강판","철강","강판","steel","plate"}
    if keys & steel_like or any(k in m for k in ("철","강","steel","plate")):
        keys.update({"steel","plate","철강","강판","철강판","hot rolled","flat steel"})
    return keys

def _country_keywords(country):
    if not country: return set()
    c = country.strip()
    keys = {c.lower(), c.replace(" ","").lower()}
    for alias in _COUNTRY_MAP.get(c, set()): keys.add(alias)
    return keys

def _is_relevant(text, material, country):
    if not text: return False
    t = text.lower()
    mat_keys = _material_keywords(material)
    ctr_keys = _country_keywords(country)
    if material: return any(k in t for k in mat_keys) if mat_keys else False
    if country:  return any(k in t for k in ctr_keys) if ctr_keys else False
    return True

def _search_queries(material, country):
    m, c = (material or "").strip(), (country or "").strip()
    queries = []
    if m and c:
        queries.extend([f"{m} {c} 수출", f"{c} {m} 무역", f"{c} {m} 시장"])
    elif m:
        queries.extend([f"{m} 수출", f"{m} 가격"])
    elif c:
        queries.extend([f"{c} 수출", f"{c} 산업"])
    else:
        queries.append("한국 수출 무역")
    seen, out = set(), []
    for q in queries:
        if q not in seen: seen.add(q); out.append(q)
    return out

def _read_txt(path):
    try:
        if os.path.exists(path): return open(path, encoding="utf-8").read().strip()
    except Exception: pass
    return ""

def _is_quota_error(err):
    s = str(err).lower()
    return "429" in s or "resource_exhausted" in s or "quota" in s

def _friendly_ai_error(err):
    """Gemini 예외 → 사용자용 한 줄. 키 문제와 할당량(429)을 구분(오해 방지)."""
    s = str(err).lower()
    if _is_quota_error(err):
        return ("⏳ Gemini 무료 일일 할당량 소진 (키는 정상입니다). "
                "내일 자정(태평양시 PT) 리셋되거나, 다른 구글계정으로 새 무료 키를 발급해 "
                "apikey.txt에 덮어쓰면 바로 됩니다.")
    if any(k in s for k in ("api key","unauthenticated","401","403","permission","invalid argument")):
        return "🔑 Gemini API 키 오류 — apikey.txt의 키를 확인하세요."
    return f"AI 분석 일시 실패: {type(err).__name__}"

# ── 데이터 저장 (OneDrive 밖 ~/.traderisk 에 보관 → 동기화 revert 방지) ──
DATA_DIR = os.path.join(os.path.expanduser("~"), ".traderisk")
try:
    os.makedirs(DATA_DIR, exist_ok=True)
except Exception:
    DATA_DIR = "."
PROFILES_FILE = os.path.join(DATA_DIR, "profiles.json")   # 회사 실제 BOM/관세
HISTORY_FILE  = os.path.join(DATA_DIR, "history.json")    # 분석 이력(추세용)

def _load_store(path, default):
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"[store] 로드 실패 {os.path.basename(path)}: {e}")
    return default

def _save_store(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"[store] 저장 실패 {os.path.basename(path)}: {e}")
        return False

def _norm_key(s):
    return (s or "").strip().lower().replace(" ", "")

# ── 공개 배포용: 결과 캐시 + IP 레이트리밋 (Gemini 할당량 보호) ──
IS_DEPLOY = bool(os.environ.get("PORT") or os.environ.get("RENDER"))
_CACHE = {}
def _cache_get(key, ttl):
    v = _CACHE.get(key)
    if v and (time.time() - v[0] < ttl): return v[1]
    return None
def _cache_set(key, val):
    _CACHE[key] = (time.time(), val)
    if len(_CACHE) > 600:
        for k in list(_CACHE)[:150]: _CACHE.pop(k, None)
_RL = {}
def _rate_ok(ip, limit=30, window=60):
    """IP당 window초에 limit회. 공개 사이트 도배·할당량 소진 방지."""
    if not IS_DEPLOY: return True          # 로컬은 제한 없음
    now = time.time(); h = [t for t in _RL.get(ip, []) if now - t < window]
    if len(h) >= limit:
        _RL[ip] = h; return False
    h.append(now); _RL[ip] = h; return True

def _make_gemini_http_client(timeout_ms=60000):
    try:
        import httpx
        return httpx.Client(verify=False, timeout=max(timeout_ms/1000.0, 30.0))
    except Exception:
        return None


# ============================================================
#  KOTRA 클라이언트
# ============================================================
class KotraClient:
    def __init__(self, service_key=None, timeout=15):
        self.key = service_key or _read_txt("kotra_key.txt") or os.environ.get("KOTRA_SERVICE_KEY","")
        self.timeout = timeout

    @property
    def enabled(self): return bool(self.key)

    def _call(self, endpoint, extra):
        params = {"serviceKey":self.key,"numOfRows":50,"pageNo":1,
                  "returnType":"JSON","type":"json","dataType":"JSON"}
        params.update(extra)
        try:
            r = requests.get(endpoint, params=params, timeout=self.timeout, verify=False)
        except Exception as e:
            print(f"  [KOTRA] {e}"); return []
        if r.status_code != 200:
            print(f"  [KOTRA] HTTP {r.status_code}"); return []
        return self._extract(r)

    @staticmethod
    def _extract(resp):
        try:
            node = resp.json()
            for k in ("response","body","items","itemList","item"):
                if isinstance(node, dict) and k in node: node = node[k]
            if isinstance(node, dict): node = [node]
            if isinstance(node, list):
                rows = [x for x in node if isinstance(x, dict)]
                if rows: return rows
        except Exception: pass
        if _HAS_BS4:
            try:
                soup = BeautifulSoup(resp.content, "xml")
                return [{c.name: c.get_text(" ",strip=True) for c in it.find_all(recursive=False)}
                        for it in soup.find_all("item")]
            except Exception: pass
        return []

    @staticmethod
    def _pick(d, keys):
        for k in keys:
            if d.get(k): return str(d[k]).strip()
        return ""

    def fetch_news(self, pages=4, rows=50):
        if not self.enabled: return []
        items, seen = [], set()
        for p in range(1, pages+1):
            rowlist = self._call(KOTRA_NEWS_ENDPOINT, {"pageNo":p,"numOfRows":rows})
            if not rowlist: break
            for d in rowlist:
                title = self._pick(d, _KOTRA_TITLE_KEYS)
                body  = self._pick(d, _KOTRA_BODY_KEYS)
                url   = self._pick(d, _KOTRA_URL_KEYS)
                if not title and not body: continue
                text = title if not body else (f"{title}. {body}" if title else body)
                uid  = url or (title + self._pick(d, _KOTRA_DATE_KEYS))
                _id  = "kotra_" + hashlib.md5(uid.encode("utf-8")).hexdigest()[:12]
                if _id in seen: continue
                seen.add(_id)
                items.append({"id":_id,"text":text,"metadata":{
                    "title":title or text[:40],"url":url,
                    "date":self._pick(d,_KOTRA_DATE_KEYS),"source":"KOTRA"}})
            time.sleep(0.3)
        return items

    def country_info(self, country, max_chars=1500):
        if not self.enabled: return ""
        rows = []
        for param in ("natnNm","nationNm","search1","cntryNm","countryNm"):
            rows = self._call(KOTRA_COUNTRY_ENDPOINT, {param:country,"numOfRows":5})
            if rows: break
        if not rows:
            allrows = self._call(KOTRA_COUNTRY_ENDPOINT, {"numOfRows":300})
            rows = [d for d in allrows if country in " ".join(str(v) for v in d.values())][:2]
        if not rows: return ""
        parts = [f"{k}: {str(v).strip()}" for d in rows[:2] for k,v in d.items()
                 if str(v).strip() and len(str(v).strip()) < 400]
        txt = " / ".join(parts)
        return ("[KOTRA 국가정보] " + txt)[:max_chars] if txt else ""


# ============================================================
#  벡터 코퍼스
# ============================================================
class NewsCorpus:
    COLLECTION = "kotra_news"
    EMBED_MODEL = "paraphrase-multilingual-MiniLM-L12-v2"

    def __init__(self, persist_dir="./chroma_db"):
        self.collection = None
        if not _HAS_CHROMA: return
        try:
            client = chromadb.PersistentClient(path=persist_dir)
            fn = None
            try:
                fn = embedding_functions.SentenceTransformerEmbeddingFunction(model_name=self.EMBED_MODEL)
            except Exception:
                print("  [Corpus] 임베딩 미설치 -> 기본")
            self.collection = client.get_or_create_collection(
                name=self.COLLECTION, metadata={"hnsw:space":"cosine"}, embedding_function=fn)
        except Exception as e:
            print(f"  [Corpus] 초기화 실패: {e}")

    @property
    def is_ready(self):
        if self.collection is None: return False
        try: return self.collection.count() > 0
        except Exception: return False

    def add(self, items):
        if self.collection is None or not items: return 0
        self.collection.upsert(
            ids=[i["id"] for i in items],
            documents=[i["text"] for i in items],
            metadatas=[i.get("metadata",{}) for i in items])
        return len(items)

    def search(self, query, top_k=3):
        if not self.is_ready: return []
        try:
            res = self.collection.query(query_texts=[query], n_results=top_k)
            docs  = (res.get("documents") or [[]])[0]
            metas = (res.get("metadatas") or [[]])[0]
            return [{"text":d,"metadata":(metas[i] if i < len(metas) else {})}
                    for i,d in enumerate(docs)]
        except Exception as e:
            print(f"  [Corpus] 검색 오류: {e}"); return []


# ============================================================
#  K-SURE 리스크 인덱스
# ============================================================
class KSureClient:
    BASE = "https://ksight.ksure.or.kr"
    _cache_data  = None
    _cache_label = ""

    def _fetch_latest_xlsx(self):
        h = {"User-Agent":"Mozilla/5.0","Accept":"application/json"}
        try:
            r = requests.get(self.BASE+"/api/board/riskidx/list?size=1",
                             verify=False, timeout=12, headers=h)
            items = r.json()["data"]["list"]
            if not items: return None, ""
            post_seq = items[0]["postSeq"]
            label    = items[0].get("postTitle","K-SURE Risk Index")
            r2 = requests.get(self.BASE+f"/api/board/riskidx/{post_seq}",
                              verify=False, timeout=12, headers=h)
            files = r2.json()["data"]["post"].get("attachFileList", [])
            xlsx_files = [f for f in files if f.get("fileExt","").upper() == "XLSX"]
            if not xlsx_files: return None, label
            file_id = xlsx_files[0]["fileId"]
            r3 = requests.get(self.BASE+f"/api/files/{file_id}",
                              verify=False, timeout=20, headers=h)
            return r3.content, label
        except Exception as e:
            print(f"  [K-SURE] {e}"); return None, ""

    def _parse_xlsx(self, content):
        if not _HAS_OPENPYXL or not content: return {}
        try:
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
            ws = wb.active
            data = {}
            for row in ws.iter_rows(min_row=3, values_only=True):
                cn  = str(row[0] or "").strip()
                ind = str(row[1] or "").strip()
                ri  = str(row[4] or "").strip()
                if not cn or not ri: continue
                data.setdefault(cn, []).append((ind, ri))
            return data
        except Exception as e:
            print(f"  [K-SURE] 파싱 오류: {e}"); return {}

    def _load_cache(self):
        if KSureClient._cache_data is None:
            print("  [K-SURE] 다운로드 중...")
            content, label = self._fetch_latest_xlsx()
            KSureClient._cache_data  = self._parse_xlsx(content)
            KSureClient._cache_label = label
            n = sum(len(v) for v in KSureClient._cache_data.values())
            print(f"  [K-SURE] {label} - {len(KSureClient._cache_data)}개국/{n}개 산업")

    def load(self, country, max_chars=1500):
        if not country: return ""
        try: self._load_cache()
        except Exception as e:
            print(f"  [K-SURE] 실패: {e}"); return ""
        data = KSureClient._cache_data or {}
        rows = data.get(country) or next(
            (v for k, v in data.items() if country in k or k in country), None)
        if not rows: return ""
        label = KSureClient._cache_label
        grade_map = {}
        for industry, ri in rows:
            grade_map.setdefault(ri, []).append(industry)
        lines = [f"[K-SURE {label}] {country} 리스크 인덱스"]
        for ri in sorted(grade_map.keys()):
            grade_str = _RI_GRADE.get(ri, ri)
            inds = ", ".join(grade_map[ri][:5])
            if len(grade_map[ri]) > 5: inds += f" 외 {len(grade_map[ri])-5}개"
            lines.append(f"  - {ri} ({grade_str}): {inds}")
        most_common_ri = max(grade_map, key=lambda k: len(grade_map[k]))
        lines.append(f"  -> 전체 {len(rows)}개 산업, 최다: {most_common_ri}")
        return "\n".join(lines)[:max_chars]

    def country_risk(self, country):
        """국가 리스크를 0~1 수치로 변환 (K-SURE RI 1~5 → (RI-1)/4).
        반환: (risk_0_1, 라벨문자열). 데이터 없으면 (None, 사유)."""
        if not country:
            return None, "국가 미지정"
        try:
            self._load_cache()
        except Exception:
            return None, "K-SURE 미연결"
        data = KSureClient._cache_data or {}
        rows = data.get(country) or next(
            (v for k, v in data.items() if country in k or k in country), None)
        if not rows:
            return None, f"{country} K-SURE 등급 없음"
        # 산업별 RI 중 최빈값(대표 등급) 사용
        from collections import Counter
        nums = []
        for _industry, ri in rows:
            m = re.search(r"(\d+)", str(ri))
            if m: nums.append(int(m.group(1)))
        if not nums:
            return None, f"{country} 등급 파싱 실패"
        rep = Counter(nums).most_common(1)[0][0]          # 최빈 RI (1~5)
        rep = max(1, min(5, rep))
        risk = (rep - 1) / 4.0                              # 1→0.0, 5→1.0
        grade_str = _RI_GRADE.get(f"RI {rep}", f"RI {rep}")
        return risk, f"K-SURE RI{rep}({grade_str})"


# ============================================================
#  RAG 분석기
# ============================================================
class RagAnalyzer:
    def __init__(self, corpus=None, kotra=None, ksure=None, api_key=None,
                 models=None, timeout_ms=60000):
        self.corpus = corpus
        self.kotra  = kotra
        self.ksure  = ksure
        self.models = models or GEMINI_MODELS
        self.client = None
        key = api_key or _read_txt("apikey.txt") or os.environ.get("GEMINI_API_KEY","")
        self.api_key = key
        if key and _HAS_GEMINI:
            try:
                http_client = _make_gemini_http_client(timeout_ms)
                opts = {"timeout": timeout_ms}
                if http_client: opts["httpx_client"] = http_client
                self.client = genai.Client(api_key=key,
                                           http_options=genai_types.HttpOptions(**opts))
                self._model = self.models[0]
                print(f"  [RAG] Gemini 연결됨 (키 길이 {len(key)})")
            except Exception as e:
                print(f"  [RAG] Gemini 초기화 실패: {e}")
        else:
            print("  [RAG] Gemini 키 없음")

    @property
    def mode(self):
        if self.corpus and self.corpus.is_ready: return "rag"
        if self.client: return "grounding"
        return "offline"

    def _search_news(self, material, country, top_k=3):
        if not self.corpus or not self.corpus.is_ready: return [], []
        seen_ids, hits = set(), []
        for q in _search_queries(material, country):
            for h in self.corpus.search(q, top_k=top_k):
                uid = h.get("metadata",{}).get("url") or h.get("text","")[:60]
                if uid in seen_ids: continue
                seen_ids.add(uid); hits.append(h)
        relevant = [h for h in hits if _is_relevant(h.get("text",""), material, country)]
        ordered  = relevant + [h for h in hits if h not in relevant]
        return ordered[:top_k], relevant

    def build_context(self, material, country=None, query=None):
        ctx = {"mode":self.mode,"news":[],"news_relevant":False,"sources":[],
               "manufacturing":_match_manufacturing(material),
               "gyeongsang":GYEONGSANG_CONTEXT,"country":"","events":[]}
        if self.corpus and self.corpus.is_ready:
            if query:
                ctx["news"] = self.corpus.search(query, top_k=3)
                ctx["news_relevant"] = any(_is_relevant(h.get("text",""), material, country)
                                           for h in ctx["news"])
            else:
                ctx["news"], relevant = self._search_news(material, country, top_k=3)
                ctx["news_relevant"] = bool(relevant)
            seen = set()
            for h in ctx["news"]:
                md = h.get("metadata",{}); u = (md.get("url") or "").strip()
                if u and u not in seen:
                    seen.add(u)
                    ctx["sources"].append({"title":md.get("title",""),"url":u,
                                           "source":md.get("source",""),"date":md.get("date","")})
        if country:
            c = self.ksure.load(country) if self.ksure else ""
            if not c and self.kotra:
                try: c = self.kotra.country_info(country)
                except Exception: pass
            ctx["country"] = c or ""
        return ctx

    @staticmethod
    def _prompt(material, country, ctx):
        ctr = country or "대상국"
        has_relevant = ctx.get("news_relevant", False)
        if ctx["news"] and has_relevant:
            lines = [f"  - [{h['metadata'].get('date','')}] {h['text']}"
                     for h in ctx["news"] if _is_relevant(h.get("text",""), material, country)]
            if not lines:
                lines = [f"  - [{h['metadata'].get('date','')}] {h['text']}" for h in ctx["news"]]
            kotra_block = "\n".join(lines)
            kotra_guide = f"위 KOTRA 뉴스에서 {material}·{ctr}와 직접 연관된 내용 우선 정리."
        elif ctx["news"]:
            kotra_block = "\n".join(
                f"  - [{h['metadata'].get('date','')}] {h['text']}" for h in ctx["news"])
            kotra_guide = (
                f"위 KOTRA 뉴스는 {material}·{ctr}와 직접 관련 적음. "
                f"간접 도출 가능한 {ctr} 산업 동향 1~2개 bullet 후 "
                f"Google 검색·공개 무역통계로 {ctr} 수출 전망·{material} 수출 전망 작성. "
                f"'미확인','관련성 낮음','확인되지 않음' 표현 금지."
            )
        else:
            kotra_block = "  (KOTRA 코퍼스에 직접 뉴스 없음)"
            kotra_guide = (
                f"Google 검색·공개 무역통계로 {ctr} 수출 동향과 {material} 수출 전망 작성."
            )
        ksure_block  = ctx.get("country","") or "  (K-SURE 미연결 - 공개 국가정보 활용)"
        other_block  = ""
        if ctx["manufacturing"]: other_block += f"  - 원자재 구조: {ctx['manufacturing']}\n"
        other_block += f"  - 지역 맥락: {ctx['gyeongsang']}"

        return f"""당신은 한국 경상도 제조·수출 중소기업 무역 리스크 분석관입니다.
**모든 출력은 반드시 한국어.** 품목({material})·수출국({ctr}) 맞춤 분석.
원가 변동·환율은 정량적 범위(%)로 추정하되 근거 없는 수치는 '추정' 명시.
현재 시점: {datetime.now().strftime('%Y년 %m월')}

[분석 대상] 품목: {material} / 수출국: {ctr}

[KOTRA 해외시장뉴스]
{kotra_block}

[K-SURE 국가보고서]
{ksure_block}

[기타 맥락]
{other_block}

[규칙] 1. '미확인','관련성 낮음','확인되지 않음' 금지. 2. KOTRA 직접 뉴스 없으면 {ctr} 수출 전망·{material} 전망으로 대체. 3. {kotra_guide}

[보고서 형식 - 마크다운 한국어]

## 📰 KOTRA 해외시장뉴스 분석
{material}·{ctr} 시장 동향 3~5개 bullet.

## 🏦 K-SURE 국가보고서 분석
{ctr} 국가 리스크(신용·정치·경제·통상) 3~5개 bullet.

## 🌐 시장 추가 정보
{material} 가격 흐름, 수급, 통상규제(관세·반덤핑·FTA) 3~5개 bullet.

## 📊 종합 분석
- **원가·수익성**: 자재 원가 변동(% 범위 추정)과 환율 효과가 마진에 미치는 영향
- **수출 전망**: {material}의 {ctr} 수출 전망(단기·중기)
- **리스크 요인**: 통상규제·지정학·물류 중 가장 시급한 리스크
- **기회 요인**: FTA 활용, 대체 시장, 가격 전가 등

## ✅ 의사결정 권고
> **진행 / 대기 / 회피** 중 하나를 굵게 표시하고 핵심 근거 2문장."""

    @staticmethod
    def _config(model_name, grounding):
        kwargs = {"temperature": 0}
        if grounding:
            try: kwargs["tools"] = [genai_types.Tool(google_search=genai_types.GoogleSearch())]
            except Exception: pass
        if model_name.startswith(("gemini-2.5","gemini-3")):
            try: kwargs["thinking_config"] = genai_types.ThinkingConfig(thinking_budget=0)
            except Exception: pass
        return genai_types.GenerateContentConfig(**kwargs)

    def stream_report(self, material, country=None, query=None):
        ctx = self.build_context(material, country=country, query=query)
        if not self.client:
            if self.api_key:
                yield "> Gemini 키 있으나 클라이언트 초기화 실패. 콘솔 확인.\n\n"
            yield from self._offline(material, country, ctx)
            return
        prompt    = self._prompt(material, country, ctx)
        grounding = not ctx.get("news_relevant", False)
        last_err  = None
        for name in self.models:
            try:
                stream = self.client.models.generate_content_stream(
                    model=name, contents=prompt, config=self._config(name, grounding))
                for chunk in stream:
                    if getattr(chunk,"text",""): yield chunk.text
                for line in self._sources_block(ctx["sources"]): yield line
                return
            except Exception as e:
                last_err = e
                if any(k in str(e).lower() for k in ("api key","unauthenticated","401","permission_denied")):
                    break
        yield from self._offline(material, country, ctx, reason=_friendly_ai_error(last_err))

    @staticmethod
    def _sources_block(sources):
        if not sources: return
        yield "\n\n-- 참고 뉴스 (KOTRA) --\n"
        for i, s in enumerate(sources, 1):
            meta = " · ".join(x for x in (s.get("source"), s.get("date")) if x)
            url  = s.get("url") or ""
            yield (f"{i}. {s.get('title','')}" + (f" - {meta}" if meta else "") +
                   (f"\n   {url}" if url else "") + "\n")

    def _offline(self, material, country, ctx, reason=None):
        ctr   = country or "대상국"
        cause = reason or "apikey.txt에 Gemini 키를 설정하면 AI 종합 분석이 활성화됩니다."
        yield (
            f"> ⚠️ **AI 종합 분석 일시 비활성** — {cause}\n"
            f"> (아래는 실시간 수집된 실데이터 기반 요약이며, 수치·등급은 정상입니다.)\n\n"
            f"## KOTRA 해외시장뉴스 분석\n- {material}·{ctr} 직접 보도는 아래 참고뉴스 참조\n\n"
            f"## K-SURE 국가보고서 분석\n"
            f"- {ctx.get('country','') or 'K-SURE 미연결'}\n\n"
            f"## 시장 추가 정보\n"
            f"- {ctx['manufacturing'] or f'{material} - 공개 시장 기반 분석'}\n\n"
            f"## 종합 분석\n"
            f"- **원가·수익성**: 좌측 카드의 BOM·환율·종합등급 참고 (정상 산출됨)\n"
            f"- **수출 전망**: {material} -> {ctr}\n"
            f"- **리스크 요인**: 통상규제·환율·물류 모니터링 필요\n"
            f"- **기회 요인**: FTA 활용, 단일 시장 의존도 점검 권장\n\n"
            f"## 의사결정 권고\n"
            f"> 점수·등급은 위 카드로 판단 가능. AI 서술형 종합은 {('할당량 리셋/새 키 후' if reason and '할당량' in reason else 'AI 활성화 후')} 재실행.\n"
        )
        for line in self._sources_block(ctx["sources"]): yield line


    def stream_unified_report(self, material, country, sim_data):
        """시뮬레이터 결과 + KOTRA + K-SURE 전부 통합해 단일 AI 보고서 스트리밍"""
        ctx = self.build_context(material, country=country)

        # KOTRA 뉴스 블록
        news = ctx.get("news", [])
        if news:
            kotra_block = "\n".join(
                f"  - [{h['metadata'].get('date','')}] {h['text'][:150]}" for h in news)
        else:
            kotra_block = "  (KOTRA 직접 뉴스 없음)"

        # K-SURE 블록
        ksure_block = ctx.get("country","") or "  (K-SURE 미연결 - 공개 국가정보 기반)"

        # BOM 원가 블록
        bom = sim_data.get("bom_details", [])
        bom_block = "\n".join(
            f"  - {b['material']}: 비중 {b['ratio']:.0f}%, "
            f"현재가 ${b['price']:,.2f}, 변동 {b['change']:+.2f}%, "
            f"원가영향 {b['impact']:+.3f}%"
            for b in bom) or "  (BOM 데이터 없음)"

        # HS코드/관세 블록
        hs = sim_data.get("hs_candidates", [])
        hs_block = "\n".join(
            f"  - {c['code']}: {c['desc']} | 관세: {c['tariff']}"
            for c in hs[:4]) or "  (HS코드 조회 실패)"

        mi  = sim_data.get("raw_impact", 0)
        fi  = sim_data.get("fx_impact", 0)
        ti  = sim_data.get("total_impact", 0)
        fr  = sim_data.get("fx_rate", 1330)
        fp  = sim_data.get("fx_pair", "USD/KRW")
        fn  = sim_data.get("fx_currency_name", "달러")
        ps  = sim_data.get("product_scores", {})

        # 산업연관(레온티에프) 가격파급 블록
        io = sim_data.get("io_analysis") or {}
        io_lines = []
        iop = io.get("product")
        if iop and iop.get("found"):
            ur = ", ".join(f"{r['name']} {r['effective_pct']}"
                           for r in iop.get("ultimate_raw_materials", []))
            io_lines.append(f"  제품구조: {iop.get('product_name','')} (공정 {iop.get('process','')})")
            if ur: io_lines.append(f"  궁극 원자재 분해: {ur}")
        imp = io.get("impact")
        if imp and imp.get("impacts"):
            io_lines.append(f"  {imp.get('raw_material','')} {imp.get('shock_pct',0):+.1f}% 변동 시 부문 파급:")
            for i in imp["impacts"][:4]:
                io_lines.append(f"    - {i['name']}(HS {i['hs_range']}): {i['price_change_pct']:+.2f}%")
        io_block = "\n".join(io_lines) or "  (산업연관 데이터 없음)"

        gs = sim_data.get("gyeongsang") or {}
        gs_block = (f"  지역 클러스터: {gs.get('cluster','')} ({gs.get('region','')})\n"
                    f"  주력 수출항: {gs.get('port','')} | 지역 주력: {gs.get('specialty','')}\n"
                    f"  활용 가능 지원: {gs.get('support','')}") if gs else "  (경상도 컨텍스트 없음)"

        grounding = not ctx.get("news_relevant", False)
        prompt = f"""당신은 한국 경상도 철강·금속 제조 중소기업의 무역 리스크 분석관입니다.
아래 수집된 모든 데이터를 통합 분석하여 **하나의 완결된 수출 의사결정 보고서**를 작성하세요.
현재: {datetime.now().strftime('%Y년 %m월')} | 품목: {material} | 수출국: {country}

=== 수집된 실제 데이터 ===

[KOTRA 해외시장뉴스]
{kotra_block}

[K-SURE 국가 리스크]
{ksure_block}

[BOM 원자재 원가 분석]
{bom_block}
  -> 자재 원가 총변동: {mi:+.2f}% | 환율({fp}) 효과: {fi:+.2f}% | 실질 수익성 변동: {ti:+.2f}%
  -> 현재 환율: {fr:.0f} 원/1{fn}

[HS코드 · 관세율]
{hs_block}

[공급망 리스크 점수 (0-100, 높을수록 안전) — AHP/z-score 엔진]
  국가/통관: {ps.get('country','N/A')} | 항만/물류: {ps.get('port','N/A')}
  품목/시장: {ps.get('item','N/A')} | 납기/지연: {ps.get('delivery','N/A')}
  ★ AHP 종합리스크: {ps.get('composite_risk','N/A')}/100 → 판정 등급: {ps.get('grade','N/A')} (정상<40<주의<60<경고<80<위험)

[산업연관(레온티에프) 가격파급]
{io_block}

[경상도 제조 현장 컨텍스트]
{gs_block}

=== 작성 규칙 ===
1. 위 수집 데이터를 그대로 인용하되 해석과 판단 추가
2. '미확인'·'확인되지 않음'·'관련성 낮음' 표현 금지
3. KOTRA 직접 뉴스 없으면 {country} 수출 전망·{material} 시장 현황으로 대체
4. 한국어만 사용. 수치는 위 데이터의 실제 숫자 그대로 활용
5. 물류·납기 분석엔 위 경상도 주력항만을 명시하고, 최종 권고엔 활용 가능한 경상도 지원기관을 1개 이상 구체적으로 연결할 것

=== 보고서 형식 (마크다운) ===

## 🚦 수출 신호: [GO / CAUTION / STOP]
> **[신호]** — 핵심 판단 이유 1-2문장

## 📰 KOTRA 시장 동향
{country} 시장 및 {material} 수출 관련 동향 3-5개 bullet

## 🏦 K-SURE 국가 리스크
{country} 신용·정치·경제·통상 리스크 3-5개 bullet

## 💰 원가 · 환율 분석
BOM 기반 원자재 원가 변동과 {fp} 환율 효과를 실제 수치로 분석. 실질 수익성에 미치는 영향.

## 📋 HS코드 · 관세 · FTA
조회된 HS코드와 실제 관세율. FTA 활용 가능 여부.

## 📊 공급망 리스크 진단
4개 리스크 점수를 해석하고 주의할 물류·통관·납기 리스크 구체 언급.

## ✅ 최종 결론 및 권고
> 위 모든 데이터를 종합한 **명확한 결론** 1개와 즉시 실행 가능한 액션 3가지."""

        if not self.client:
            yield from self._offline(material, country, ctx)
            return
        last_err = None
        for name in self.models:
            try:
                stream = self.client.models.generate_content_stream(
                    model=name, contents=prompt, config=self._config(name, grounding))
                for chunk in stream:
                    if getattr(chunk, "text", ""): yield chunk.text
                for line in self._sources_block(ctx["sources"]): yield line
                return
            except Exception as e:
                last_err = e
                if any(k in str(e).lower() for k in ("api key","unauthenticated","401","permission")):
                    break
                continue
        yield from self._offline(material, country, ctx, reason=_friendly_ai_error(last_err))


def build_corpus(kotra, corpus, pages=4):
    if not kotra.enabled:
        print("  [build_corpus] KOTRA 키 없음"); return corpus
    items = kotra.fetch_news(pages=pages)
    n = corpus.add(items)
    print(f"  [build_corpus] KOTRA {len(items)}건 수집 / {n}건 적재")
    return corpus


# ============================================================
#  공급망 엔진 (원자재 시세 / FX / BOM / HS코드 / 시뮬레이터)
# ============================================================
class SupplyChainEngine:

    def __init__(self, rag_client=None, models=None, ksure=None, io_engine=None,
                 country_grades=None):
        self._rag_client = rag_client
        self._models     = models or GEMINI_MODELS
        self._model      = self._models[0] if self._models else "gemini-2.0-flash"
        self.ksure       = ksure
        self.io_engine   = io_engine     # 레온티에프 산업연관 가격파급 엔진
        self.country_grades = country_grades or {}   # data.go.kr K-SURE 국별신용등급 1~7
        self.data_dict   = {}
        self._fx_cache   = {}
        self._bdi        = None      # 실시간 Baltic Dry Index (항만/물류 프록시)

    # ── 원자재 시세 ──────────────────────────────────────────
    def sync_data(self):
        print("원자재 시세 수집 중...")
        if _HAS_PANDAS:
            try:
                res = requests.get("https://tradingeconomics.com/commodities",
                                   headers={"User-Agent": UA}, timeout=12, verify=False)
                tables = pd.read_html(_SIO(res.text))
                count = 0
                for df in tables:
                    for _, row in df.iterrows():
                        try:
                            name   = str(row.iloc[0]).strip()
                            if not name or name == "nan": continue
                            price  = float(str(row.iloc[1]).replace(",", ""))
                            change = float(str(row.iloc[3]).replace("%", ""))
                            self._store(name, price, change)
                            count += 1
                        except (ValueError, IndexError):
                            continue
                if count > 0:
                    print(f"  원자재 {count}개 로드")
                    self._capture_bdi()
                    return
            except Exception as e:
                print(f"  시세 크롤링 실패: {e}")
        print("  오프라인 샘플 사용")
        for name, price, change in OFFLINE_DATA:
            self._store(name, price, change)
        self._capture_bdi()

    def _capture_bdi(self):
        """수집된 시세 중 Baltic Dry(건화물운임지수)를 항만 혼잡 프록시로 저장."""
        for k, v in self.data_dict.items():
            if "baltic" in k:
                try:
                    self._bdi = float(v["p"]);
                    print(f"  [BDI] {self._bdi:.0f} (항만/물류 지표)")
                    return
                except Exception:
                    pass
        # 별도 페이지에서 직접 시도
        if _HAS_PANDAS:
            try:
                res = requests.get("https://tradingeconomics.com/commodity/baltic",
                                   headers={"User-Agent": UA}, timeout=10, verify=False)
                tables = pd.read_html(_SIO(res.text))
                for df in tables:
                    for _, row in df.iterrows():
                        cell = str(row.iloc[0]).lower()
                        if "baltic" in cell:
                            self._bdi = float(str(row.iloc[1]).replace(",", ""))
                            print(f"  [BDI] {self._bdi:.0f} (직접 수집)")
                            return
            except Exception as e:
                print(f"  [BDI] 수집 실패: {e}")
        print("  [BDI] 미수집 - 항만 점수는 중립 기준 적용")

    def _store(self, name, price, change):
        e = {"name": name, "p": price, "c": change}
        self.data_dict[name.lower()] = e
        self.data_dict[name.lower().replace(" ", "")] = e

    # ── 실시간 환율 ──────────────────────────────────────────
    def get_fx_rate(self, country="미국"):
        currency_code, currency_name, baseline, decimals = COUNTRY_CURRENCY.get(
            country, ("USD", "달러", FX_BASELINE, 1))
        now = time.time()
        cached = self._fx_cache.get(currency_code)
        if cached and now - cached.get("ts", 0) < 300:
            return cached
        sources = [
            "https://open.er-api.com/v6/latest/USD",
            "https://api.exchangerate-api.com/v4/latest/USD",
        ]
        for url in sources:
            try:
                r = requests.get(url, timeout=6, verify=False)
                d = r.json()
                rates  = d.get("rates") or d.get("conversion_rates") or {}
                krw    = float(rates.get("KRW", 0))
                target = float(rates.get(currency_code, 0))
                if krw > 900 and target > 0:
                    rate = round(krw, decimals) if currency_code == "USD" else round(krw / target, decimals)
                    chg  = round((rate - baseline) / baseline * 100, 2)
                    result = {"rate":rate,"change_pct":chg,"favorable":chg>0,
                              "source":url.split("/")[2],"ts":now,
                              "currency":currency_code,"currency_name":currency_name,
                              "pair":f"{currency_code}/KRW","decimals":decimals}
                    self._fx_cache[currency_code] = result
                    return result
            except Exception:
                continue
        fallback = {"rate":baseline,"change_pct":0.0,"favorable":False,
                    "source":"offline","ts":now,"currency":currency_code,
                    "currency_name":currency_name,"pair":f"{currency_code}/KRW","decimals":decimals}
        self._fx_cache[currency_code] = fallback
        return fallback

    # ── 리스크 점수 (정규화 기반) ─────────────────────────────
    #  설계 원칙: 모든 점수 = 100 − Σ(가중치×리스크요소×100)
    #  · 리스크요소ᵢ = clamp((관측값−안전기준)/(위험기준−안전기준), 0, 1)
    #  · 가중치 합 = 1 (각 점수 내부)
    #  · 안전/위험 기준은 출처 있는 임계밴드, 가중치는 명시값. 떠다니는 상수 없음.
    @staticmethod
    def _risk_norm(x, safe, danger):
        """관측값 x를 안전기준(0)~위험기준(1) 리스크로 정규화."""
        if danger == safe: return 0.0
        return max(0.0, min(1.0, (x - safe) / (danger - safe)))

    def _profile_for(self, name):
        n = (name or "").lower()
        for key, meta in COMMODITY_PROFILES.items():
            if key in n: return dict(meta)
        return dict(DEFAULT_PROFILE)

    def _port_risk(self):
        """항만/물류 리스크 0~1. BDI(건화물운임지수) 실시간.
        밴드: 1000(평년 저점)~3000(혼잡 고점). 반환:(risk, 근거문자열)."""
        if self._bdi and self._bdi > 0:
            r = self._risk_norm(self._bdi, 1000.0, 3000.0)
            return r, f"BDI {self._bdi:.0f}"
        return 0.45, "BDI 미수집(중립)"

    def _price_risk(self, c):
        """가격 변동성 리스크 0~1. |변동률| 0%(안전)~10%(위험)."""
        return self._risk_norm(abs(c), 0.0, 10.0), f"변동 {c:+.2f}%"

    def _tariff_risk(self, tariff_pct):
        """관세 리스크 0~1. 0%(안전)~25%(고관세 위험)."""
        return self._risk_norm(float(tariff_pct or 0), 0.0, 25.0), f"관세 {float(tariff_pct or 0):.1f}%"

    def _country_risk(self, country, fallback_profile=None):
        """국가 리스크 0~1.
        ① data.go.kr K-SURE 국별신용등급(1~7) → GRADE_TO_SCORE/100 (표준, 우선)
        ② K-SURE 리스크인덱스(RI 1~5) → (RI-1)/4
        ③ 품목 공급 집중도(추정)"""
        if country and self.country_grades and _scoring is not None:
            g = self.country_grades.get(country) or next(
                (v for k, v in self.country_grades.items() if country in k or k in country), None)
            if g is not None:
                risk = _scoring.GRADE_TO_SCORE.get(int(g), 45.0) / 100.0
                return risk, f"K-SURE {int(g)}등급"
        if country and self.ksure:
            r, label = self.ksure.country_risk(country)
            if r is not None:
                return r, label
        if fallback_profile:
            return float(fallback_profile.get("concentration", 0.5)), "국가등급 미수집(추정)"
        return 0.5, "국가정보 없음(중립)"

    def forecast_price_change(self, name, current_change=None):
        """[가격예측 훅] 다음 기간 예상 변동률(%).

        현재는 실시간 변동률을 그대로 패스스루한다. modeling.ipynb의
        RandomForest 예측기(철광석 기준 MAPE 3.95%)를 학습용 CSV
        (산업통상부_철강원자재 가격동향)와 함께 이 자리에 연결하면,
        '관측 변동률' 대신 '예측 변동률'을 scoring 축1로 넘길 수 있다.
        """
        # TODO: RandomForest/SARIMA 예측 연동 (학습 CSV 확보 시 활성화)
        return current_change if current_change is not None else 0.0

    def _sigma_for(self, name):
        """품목명 → 조원 MATERIAL_SIGMA_FALLBACK(월간 변동성 σ) 매핑. 없으면 일반 8.0%."""
        n = (name or "").lower()
        table = {"철광석":"철광석","iron":"철광석","ore":"철광석",
                 "유연탄":"유연탄","coal":"유연탄","coking":"유연탄",
                 "스크랩":"철스크랩","scrap":"철스크랩","고철":"철스크랩",
                 "구리":"구리광","copper":"구리광",
                 "니켈":"니켈","nickel":"니켈"}
        for k, mat in table.items():
            if k in n and _scoring is not None:
                return _scoring.MATERIAL_SIGMA_FALLBACK.get(mat, 8.0), mat
        return 8.0, None

    def calculate_scores(self, name, c, price=0, country=None, tariff_pct=None):
        prof = self._profile_for(name)
        if tariff_pct is None:
            tariff_pct = prof["tariff"]               # HS 미조회 시 품목 기본 관세
        r_price,   b_price   = self._price_risk(c)
        r_port,    b_port    = self._port_risk()
        r_tariff,  b_tariff  = self._tariff_risk(tariff_pct)
        r_country, b_country = self._country_risk(country, prof)

        # ── 조원 scoring.py 엔진: z-score 정규화 + AHP 가중 종합 ──
        if _HAS_ENGINES and _scoring is not None:
            sigma, _mat = self._sigma_for(name)
            z     = (c / sigma) if sigma else 0.0
            axis1 = _scoring.zscore_to_axis_score(z)     # 가격 리스크 0~100 (z-score 매핑)
            axis2 = r_port    * 100.0                    # 항만 리스크 (실시간 BDI)
            axis3 = r_country * 100.0                    # 국가 리스크 (실시간 K-SURE RI)
            composite = _scoring.composite_score(axis1, axis2, axis3, _scoring.DEFAULT_WEIGHTS)
            grade     = _scoring.classify(composite)     # 정상/주의/경고/위험
            w         = _scoring.DEFAULT_WEIGHTS
            delivery_risk = 0.6 * axis2 + 0.4 * axis3    # 납기=항만(0.6)+국가(0.4) 국소블렌드
            s = {"country":  max(5, min(99, round(100 - axis3))),
                 "port":     max(5, min(99, round(100 - axis2))),
                 "item":     max(5, min(99, round(100 - axis1))),
                 "delivery": max(5, min(99, round(100 - delivery_risk)))}
            sig = _scoring.trade_signal(composite, c)
            s["export_ok"] = sig["export"] == "적절"
            s["import_ok"] = sig["import"] == "적절"
            s["composite_risk"] = round(composite, 1)
            s["grade"]   = grade
            s["axes"]    = {"price":round(axis1,1),"port":round(axis2,1),"country":round(axis3,1)}
            s["weights"] = {k: round(v,3) for k,v in w.items()}
            zlabel = f" z={z:+.2f}" if sigma else ""
            s["basis_lines"] = {
                "country":  b_country,
                "port":     b_port,
                "item":     f"{b_price}{zlabel}",
                "delivery": f"{b_port} · {b_country}",
            }
            s["basis"] = {"price_risk":round(r_price,2),"port_risk":round(r_port,2),
                          "tariff_risk":round(r_tariff,2),"country_risk":round(r_country,2),
                          "tariff":float(tariff_pct or 0),"volatility":round(abs(c),2),
                          "z_score":round(z,2),"composite":round(composite,1),
                          "grade":grade,"engine":"AHP/z-score(조원 scoring.py)"}
            return s

        # ── 폴백: 엔진 미로드 시 기존 정규화식 ──
        country_s  = 100 - 100 * (0.65 * r_country + 0.35 * r_tariff)
        port_s     = 100 - 100 * (0.70 * r_port    + 0.30 * r_price)
        item_s     = 100 - 100 * (0.70 * r_price   + 0.30 * r_tariff)
        delivery_s = 100 - 100 * (0.60 * r_port    + 0.40 * r_country)
        s = {"country":max(5,min(99,round(country_s))),
             "port":max(5,min(99,round(port_s))),
             "item":max(5,min(99,round(item_s))),
             "delivery":max(5,min(99,round(delivery_s)))}
        comp = sum(v for k, v in s.items()) / 4
        s["export_ok"] = comp >= 50 and c >= -2.5
        s["import_ok"] = r_port <= 0.6 and c <= 3.5
        s["basis_lines"] = {
            "country":  f"{b_country} · {b_tariff}",
            "port":     f"{b_port} · {b_price}",
            "item":     f"{b_price} · {b_tariff}",
            "delivery": f"{b_port} · {b_country}",
        }
        s["basis"] = {"price_risk":round(r_price,2),"port_risk":round(r_port,2),
                      "tariff_risk":round(r_tariff,2),"country_risk":round(r_country,2),
                      "tariff":float(tariff_pct or 0),"volatility":round(abs(c),2),
                      "composite":round(comp,1)}
        return s

    # ── 원자재 조회 ──────────────────────────────────────────
    def _norm(self, mat):
        raw = (mat or "").strip().lower().replace(" ", "")
        return MATERIAL_ALIASES.get(raw, raw)

    def _resolve(self, mat):
        key = self._norm(mat)
        if key in self.data_dict: return self.data_dict[key]
        for k, v in self.data_dict.items():
            if key and (key in k or k in key): return v
        fb = {"steel":(480,0.8),"copper":(4.5,1.2),"iron":(105,-0.5),
              "aluminum":(2500,2.1),"nickel":(16800,0.55),"rubber":(1.65,0.4),"zinc":(2780,-0.33)}
        if key in fb:
            p, c = fb[key]; return {"name":mat.upper(),"p":p,"c":c}
        return {"name":str(mat).upper(),"p":0.0,"c":0.0}

    # ── HS 코드 크롤링 ────────────────────────────────────────
    def _hs_headers(self):
        return {"User-Agent":UA,"Referer":HS_BASE+"/",
                "Accept-Language":"ko-KR,ko;q=0.9","Accept":"text/html,*/*"}

    def _hs_post(self, path, data):
        try:
            r = requests.post(f"{HS_BASE}/{path}", data=data,
                              headers=self._hs_headers(), timeout=12, verify=False)
            if r.status_code == 200: return r.text
        except Exception as e:
            print(f"  hs-tariff 오류: {e}")
        return ""

    def _hs_keyword_items(self, keyword):
        html = self._hs_post("keyword_hs-fta.php", {"keyword": keyword})
        if not html or not _HAS_BS4: return []
        soup = BeautifulSoup(html, "html.parser")
        items = []
        for li in soup.select("li.keyword_select"):
            oc = li.get("onclick", "")
            m  = re.search(r"keyword_select\('([^']*)','([^']*)','([^']*)'\)", oc)
            if m: items.append((m.group(1), m.group(2), m.group(3)))
        return items

    def _hs_score(self, product, wkr, wen, hs4):
        p = product.lower().strip(); w = (wkr or "").lower(); we = (wen or "").lower()
        sc = 10 if hs4 else 0
        if p == w: sc += 80
        elif p in w or w in p: sc += 50
        eng = KO_EN.get(product, "").lower()
        if eng and eng in we: sc += 40
        if eng and eng in w:  sc += 30
        if any(k in p for k in ("선풍기","에어컨","모터","펌프","냉장","세탁")):
            if hs4 and hs4[:2] in ("84","85"): sc += 60
        if any(k in p for k in ("철판","H빔","파이프","볼트","나사","스프링","플랜지")):
            if hs4 and hs4[:2] in ("72","73","74"): sc += 60
        return sc

    def _hs_parse_data2(self, html):
        codes = []
        for c in re.findall(r"hs_code6'\)\.val\('(\d{6})'\)", html):
            codes.append(f"{c[:4]}.{c[4:6]}")
        for c in re.findall(r"<i[^>]*>(\d{4}\.\d{2})</i>", html):
            if c not in codes: codes.append(c)
        return codes

    def _hs_parse_data4(self, html, country, product):
        seen, out = set(), []
        for code, desc in re.findall(r"trade_goods_name\('([^']+)','([^']*)'", html):
            code = code.strip()
            if code in seen: continue
            seen.add(code)
            out.append({"code":code, "desc":f"{desc.strip() or product}", "tariff":""})
        for code in re.findall(r"(\d{4}\.\d{2}-\d{4})", html):
            if code in seen: continue
            seen.add(code)
            out.append({"code":code, "desc":product, "tariff":""})
        return out

    def crawl_hs_codes(self, product, country):
        terms = [product]
        eng = KO_EN.get(product)
        if eng and eng not in terms: terms.append(eng)
        if len(product) >= 2: terms.append(product[:2])
        all_items = []
        for term in list(dict.fromkeys(terms)):
            all_items.extend(self._hs_keyword_items(term))
        if not all_items: return []
        ranked = sorted(all_items,
                        key=lambda x: self._hs_score(product, x[0], x[1], x[2]),
                        reverse=True)
        sh, sh2, hs4 = ranked[0]
        hs4 = (hs4 or "")[:4] or "8414"
        w2  = f"sh_type=KoE|sh_name={sh}|sh_name2={sh2 or sh}|hsCode={hs4}|hs_code4={hs4}|"
        html2  = self._hs_post("ai/index.data2.php", {"wheres":w2,"page":"1","viewnum":"200"})
        hs6s   = self._hs_parse_data2(html2)
        if not hs6s: hs6s = [f"{hs4}.10", f"{hs4}.90"]
        results = []
        for hs6 in hs6s[:5]:
            hs6c = hs6.replace(".", "")
            w4   = f"sh_type=KoE|sh_name={sh}|sh_name2={sh2 or sh}|hsCode={hs4}|hs_code4={hs4}|hs_code6={hs6c}|"
            html4 = self._hs_post("ai/index.data4.php", {"wheres":w4,"page":"1","viewnum":"200"})
            results.extend(self._hs_parse_data4(html4, country, product))
        seen2, unique = set(), []
        for c in results:
            if c["code"] not in seen2:
                seen2.add(c["code"]); unique.append(c)
        return unique[:10]

    # ── AI HS 코드 + 실제 관세율 (세금 수정) ─────────────────
    def _ai_json(self, prompt):
        """모델 폴백 루프로 Gemini 호출 → 첫 성공 응답 텍스트 반환(없으면 None).
        429(할당량)는 모델별로 따로 잡히므로, 한 모델이 막히면 다음 모델로 넘어간다."""
        if not self._rag_client: return None
        last = None
        for mdl in self._models:
            try:
                resp = self._rag_client.models.generate_content(
                    model=mdl, contents=prompt,
                    config=genai_types.GenerateContentConfig(temperature=0))
                t = (getattr(resp, "text", "") or "").strip()
                if t: return t
            except Exception as e:
                last = e
                if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
                    continue          # 다음 모델은 별도 quota → 시도 가치 있음
                continue
        if last:
            print(f"  AI 호출 실패(전 모델 소진): {type(last).__name__}: {str(last)[:90]}")
        return None

    def _ai_hs(self, product, country):
        prompt = (
            f"관세·무역 전문가. '{product}'을 한국에서 '{country}'로 수출 시 "
            f"적용 HS코드 후보 3개와 {country} 실제 관세율을 JSON만 출력.\n"
            f"반드시 실제 관세율(%) 또는 FTA 협정세율을 포함할 것. 예시:\n"
            f'{{"hs_candidates":['
            f'{{"code":"8414.51","desc":"가정용 선풍기","tariff":"{country} 일반 2.0%, 한-{country} FTA 0%"}},'
            f'{{"code":"8414.59","desc":"기타 팬","tariff":"{country} 일반 3.0%, FTA 0%"}}'
            f']}}'
        )
        raw = self._ai_json(prompt)
        if not raw: return []
        try:
            raw  = raw.replace("```json","").replace("```","").strip()
            m    = re.search(r"\{[\s\S]*\}", raw)
            data = json.loads(m.group(0) if m else raw)
            return data.get("hs_candidates", [])
        except Exception as e:
            print(f"  AI HS 파싱 오류: {e}"); return []

    # ── AI BOM 추론 ──────────────────────────────────────────
    def _mock_bom(self, product):
        if any(k in product for k in ("냉장","에어컨","냉동","제빙")):
            return {"steel":30,"copper":25,"aluminum":20,"plastic":25}
        if any(k in product for k in ("선풍기","팬","fan")):
            return {"steel":35,"aluminum":30,"copper":15,"plastic":20}
        if any(k in product for k in ("철판","H빔","각관","파이프")):
            return {"steel":85,"aluminum":10,"zinc":5}
        if any(k in product for k in ("자동차","전기차")):
            return {"steel":55,"aluminum":20,"copper":15,"plastic":10}
        return {"steel":45,"aluminum":25,"copper":20,"plastic":10}

    def _ai_bom(self, product, country):
        if not self._rag_client:
            return self._mock_bom(product), []
        prompt = (
            f"공급망·원가 전문가. 한국 경상도 제조업체가 '{product}'을 '{country}'에 수출.\n"
            f"[임무1] '{product}' 핵심 원자재(영문 소문자)와 비중(%). 합계 100%. 최대 5종.\n"
            f"[임무2] '{country}' 수출 필수 인증·통관 서류 5개.\n"
            f'JSON만: {{"bom":{{"steel":40,"copper":25,"plastic":35}},"checklist":["인증1","서류2"]}}'
        )
        raw = self._ai_json(prompt)
        if not raw:
            return self._mock_bom(product), []
        try:
            raw  = raw.replace("```json","").replace("```","").strip()
            m    = re.search(r"\{[\s\S]*\}", raw)
            data = json.loads(m.group(0) if m else raw)
            return (data.get("bom") or {}), data.get("checklist", [])
        except Exception as e:
            print(f"  AI BOM 파싱 오류: {e}"); return self._mock_bom(product), []

    # ── 시뮬레이터 ──────────────────────────────────────────
    def run_simulator(self, product, country, bom_override=None, tariff_override=None):
        print(f"\n[시뮬레이터] {product} -> {country}"
              + (" [회사 BOM]" if bom_override else ""))
        fx = self.get_fx_rate(country)

        # HS 코드: 크롤링 후 반드시 AI로 실제 관세율 보완
        crawled  = self.crawl_hs_codes(product, country)
        ai_cands = self._ai_hs(product, country)

        # AI 관세율로 크롤 데이터 보완 (세금 수정 핵심)
        if ai_cands:
            ai_map = {c["code"][:7]: c for c in ai_cands}
            for c in crawled:
                ai_match = ai_map.get(c["code"][:7])
                if ai_match and not c.get("tariff"):
                    c["tariff"] = ai_match.get("tariff","")
        cands = crawled or ai_cands or [{"code":"N/A","desc":"HS 검색 실패","tariff":"정보 없음"}]
        # 회사가 직접 입력한 관세율(override) 우선 — AI 추정보다 신뢰
        if tariff_override is not None:
            cands = [{"code": cands[0].get("code","N/A"), "desc": cands[0].get("desc", product),
                      "tariff": f"회사 입력 {float(tariff_override):.1f}%"}] + cands
        for c in cands:
            if not c.get("tariff"):
                c["tariff"] = f"{country} 관세율 - 회사 BOM에 직접 입력 또는 FTA 협정세율 확인"

        # BOM: 회사 실제 입력(override) 우선, 없으면 AI 추정
        if bom_override:
            bom, checklist, bom_source = dict(bom_override), [], "회사 실제"
        else:
            bom, checklist = self._ai_bom(product, country)
            if not bom: bom = self._mock_bom(product)
            bom_source = "AI 추정"

        bom_details, raw_impact = [], 0.0
        for mat, ratio in bom.items():
            ratio  = float(ratio)
            market = self._resolve(mat)
            chg    = float(market.get("c", 0))
            price  = float(market.get("p", 0))
            impact = chg * (ratio / 100.0)
            raw_impact += impact         # BOM 가중 원자재 가격변동(%)
            bom_details.append({"material":market.get("name",mat.upper()),
                                 "ratio":ratio,"price":round(price,2),
                                 "change":round(chg,2),"impact":round(impact,3)})

        fx_impact    = fx.get("change_pct", 0.0)
        total_impact = raw_impact - fx_impact

        # 관세율: 회사 입력 우선, 없으면 조회된 후보 중 첫 숫자%
        if tariff_override is not None:
            tariff_pct = float(tariff_override)
        else:
            tariff_pct = None
            for cand in cands:
                m = re.search(r"(\d+(?:\.\d+)?)\s*%", str(cand.get("tariff","")))
                if m:
                    tariff_pct = float(m.group(1)); break

        # 제품 단위 단일 산출: BOM가중 가격변동 + K-SURE 국가등급(live)
        #  + BDI(live) + 실제 관세. 모두 출처 있는 입력값.
        scores = self.calculate_scores(product, raw_impact, country=country,
                                       tariff_pct=tariff_pct)
        product_scores = {k: scores[k] for k in ("country","port","item","delivery")}
        product_scores["basis_lines"] = scores.get("basis_lines", {})
        # AHP 종합점수·등급(정상/주의/경고/위험)·3축·가중치 전달
        for k in ("composite_risk", "grade", "axes", "weights"):
            if k in scores: product_scores[k] = scores[k]

        # ── 산업연관(레온티에프) 가격파급 분석 ──
        io_result = None
        if self.io_engine is not None:
            try:
                io_result = self._io_propagation(cands, bom_details)
            except Exception as e:
                print(f"  [IO] 파급분석 실패: {e}")

        if not checklist:
            checklist = [
                f"{country} 제품안전 인증 (CE/UL/FCC 해당 시)",
                "원산지 증명서 (C/O) 발급", "상업송장·포장명세서",
                "HS CODE 품목분류 사전확인서",
                f"한-{country} FTA 협정세율 적용 검토",
            ]
        return {
            "hs_candidates":  cands,
            "bom_details":    bom_details,
            "checklist":      checklist,
            "raw_impact":     round(raw_impact, 3),
            "fx_impact":      round(fx_impact, 3),
            "total_impact":   round(total_impact, 3),
            "fx_rate":        fx.get("rate", FX_BASELINE),
            "fx_change":      fx_impact,
            "fx_source":      fx.get("source","offline"),
            "fx_pair":        fx.get("pair","USD/KRW"),
            "fx_currency_name": fx.get("currency_name","달러"),
            "product_scores": product_scores,
            "io_analysis":    io_result,
            "bom_source":     bom_source,
            "tariff_pct":     tariff_pct,
            "gyeongsang":     gyeongsang_context(product),
        }

    def _io_propagation(self, cands, bom_details):
        """레온티에프 산업연관 엔진으로 HS제품 구조 + 지배원자재 가격파급 산출."""
        eng = self.io_engine
        out = {"mode": getattr(eng, "mode", "fallback")}
        # 1) 대표 HS코드로 철강제품 구조(궁극 원자재 분해) 조회
        for c in cands:
            digits = "".join(ch for ch in str(c.get("code","")) if ch.isdigit())
            if len(digits) >= 4:
                lk = eng.lookup_by_hs(digits[:4])
                if lk.get("found"):
                    out["product"] = lk; break
        # 2) 지배 원자재 → 철강 부문별 가격파급
        io_raw_map = {"iron":"철광석","steel":"철광석","철":"철광석","ore":"철광석",
                      "scrap":"철스크랩","고철":"철스크랩","coal":"유연탄","유연탄":"유연탄",
                      "copper":"구리광","구리":"구리광","aluminum":"알루미늄","알루미늄":"알루미늄",
                      "nickel":"니켈","니켈":"니켈"}
        dom = max(bom_details, key=lambda b: b.get("ratio",0), default=None)
        if dom:
            name_l = str(dom.get("material","")).lower()
            raw = next((v for k, v in io_raw_map.items() if k in name_l), None)
            if raw:
                imp = eng.calculate_price_impact(raw, dom.get("change", 0))
                if imp.get("impacts"):
                    out["impact"] = imp
        return out if (out.get("product") or out.get("impact")) else None

    # ── 종합 SSE 분석 (신호등) ───────────────────────────────
    def stream_comprehensive(self, product, country, mat_impact, fx_impact,
                             fx_rate, fx_pair="USD/KRW", currency_name="달러"):
        total = mat_impact - fx_impact
        if not self._rag_client:
            sig = "go" if total < -0.3 else "stop" if total > 1.5 else "caution"
            result = {"signal":sig,"signal_reason":"원자재·환율 자동 판단 (AI 키 미설정)",
                      "lines":[
                          f"{product} -> {country} 수출 시나리오 분석 (오프라인 모드)",
                          f"자재 원가 변동 {mat_impact:+.2f}% / 환율({fx_pair}) 효과 {fx_impact:+.2f}%",
                          "글로벌 무역환경: 공개 정보 기반 추정",
                          "경상도 수출 여건: 창원·포항 산업단지 가동률 정상 추정",
                          "apikey.txt에 Gemini 키 입력 시 실시간 AI 분석 활성화",
                      ]}
            yield "data: " + json.dumps(result, ensure_ascii=False) + "\n\n"
            yield "data: [DONE]\n\n"; return

        # 프롬프트: JSON 형식 예시를 명확히 분리
        prompt = (
            f"당신은 한국 경상도 철강·금속 소기업 수출 컨설턴트입니다.\n"
            f"현재: {datetime.now().strftime('%Y년 %m월')}\n\n"
            f"[분석 데이터]\n"
            f"- 완제품: {product}\n"
            f"- 수출국: {country}\n"
            f"- 자재 원가 변동: {mat_impact:+.2f}%\n"
            f"- 환율효과({fx_pair}, 외화강세=양수): {fx_impact:+.2f}%\n"
            f"- 실질 수익성 변동: {total:+.2f}%\n"
            f"- 현재 {fx_pair}: {fx_rate:.0f}원 per 1 {currency_name}\n\n"
            f"[판단 기준]\n"
            f"1. 원자재 가격 추세 (철강·알루미늄·구리 국제 시세)\n"
            f"2. 환율 수혜 여부 ({currency_name} 강세/원화 약세 = 수출 유리)\n"
            f"3. 글로벌 무역환경 (관세, 중국 경쟁, FTA)\n"
            f"4. 비경제적 요소 (지정학, 물류, ESG)\n"
            f"5. 경상도 소기업 여건 (창원·포항·부산항)\n\n"
            f"아래 JSON 형식으로만 출력하세요. 코드블록이나 다른 텍스트 없이 JSON만.\n"
            f'{{"signal":"go","signal_reason":"한 문장 판단 이유",'
            f'"lines":["제품과 HS코드 관세 FTA 분석","원가와 환율 수익성 분석",'
            f'"글로벌 무역 지정학 리스크","경상도 소기업 수출 물류 여건","최종 수출 진행 여부 권고"]}}\n\n'
            f"signal은 go(유리)/caution(보통)/stop(위험) 중 하나. lines의 각 항목은 실제 분석 내용으로 채우세요."
        )
        import traceback as _tb
        last_err = None
        for try_model in self._models:
            try:
                print(f"  [신호등] 모델 시도: {try_model}")
                resp = self._rag_client.models.generate_content(
                    model=try_model, contents=prompt)
                text = ""
                try:
                    text = resp.text or ""
                except Exception as te:
                    print(f"  [신호등] resp.text 오류: {te}")
                    text = ""
                print(f"  [신호등] 응답 길이: {len(text)} / 앞부분: {repr(text[:80])}")
                if not text.strip():
                    last_err = ValueError(f"{try_model}: 빈 응답")
                    continue
                clean = text.replace("```json","").replace("```","").strip()
                m = re.search(r"\{[\s\S]*\}", clean)
                if m:
                    clean = m.group(0)
                yield "data: " + clean.replace("\n", "\\n") + "\n\n"
                yield "data: [DONE]\n\n"
                return
            except Exception as e:
                print(f"  [신호등] {try_model} 실패: {type(e).__name__}: {e}")
                _tb.print_exc()
                last_err = e
                continue
        # 모든 모델 실패 시
        print(f"  [신호등] 전체 실패: {last_err}")
        sig = "go" if total < -0.3 else "stop" if total > 1.5 else "caution"
        err = {"signal": sig,
               "signal_reason": f"AI 분석 실패 - {type(last_err).__name__}",
               "lines": [
                   f"{product} -> {country} 수출 분석 (원자재·환율 자동 판단)",
                   f"자재 원가 {mat_impact:+.2f}% / 환율({fx_pair}) {fx_impact:+.2f}% / 실질 {total:+.2f}%",
                   "글로벌 무역환경: AI 분석 불가 - CMD 창 오류 메시지 확인 필요",
                   "경상도 여건: 창원·포항 산업단지 기본 정상 가동 추정",
                   f"오류: {str(last_err)[:80]}",
               ]}
        yield "data: " + json.dumps(err, ensure_ascii=False) + "\n\n"
        yield "data: [DONE]\n\n"

    # ── 단일 원자재 스트림 분석 ──────────────────────────────
    def stream_analysis(self, name, c, scores, price):
        basis = scores.get("basis", {})
        if not self._rag_client:
            d = "상승" if c > 0 else "하락"
            r = "HIGH" if abs(c)>5 else "MED" if abs(c)>2 else "LOW"
            msg = (f"## {name} 리스크 브리핑\\n\\n"
                   f"현재가 **${price:,.2f}**, **{abs(c):.2f}% {d}**. 리스크: **{r}**\\n"
                   f"변동성리스크 {basis.get('price_risk',0):.2f} / 관세 {basis.get('tariff',0):.1f}% / 항만리스크 {basis.get('port_risk',0):.2f}\\n\\n"
                   f"### 권고\\n- apikey.txt 설정 시 AI 분석 활성화")
            yield "data: " + msg + "\n\n"; yield "data: [DONE]\n\n"; return
        prompt = (
            f"원자재: {name} | 현가: ${price:,.4f} | 변동: {c:+.2f}%\n"
            f"리스크(국가/항만/품목/납기): {scores['country']}/{scores['port']}/{scores['item']}/{scores['delivery']}\n"
            f"변동성리스크: {basis.get('price_risk',0):.2f} | 관세: {basis.get('tariff',0):.1f}% | 항만리스크: {basis.get('port_risk',0):.2f}\n"
            f"수출: {'유리' if scores['export_ok'] else '불리'} | 수입: {'원활' if scores['import_ok'] else '주의'}\n\n"
            "경상도 철강 소기업 관점 비경제 리스크 중심 실무 브리핑 한국어:\n"
            f"## {name} 수출입 브리핑\n"
            "### 시장·지정학 | ### 수출전략 | ### 수입조달 | ### 비경제 리스크 3개 | ### 권고 3개"
        )
        try:
            for name_m in self._models:
                try:
                    stream = self._rag_client.models.generate_content_stream(
                        model=name_m, contents=prompt,
                        config=genai_types.GenerateContentConfig(temperature=0))
                    for chunk in stream:
                        if getattr(chunk,"text",""):
                            yield "data: " + chunk.text.replace("\n","\\n") + "\n\n"
                    break
                except Exception:
                    continue
        except Exception as e:
            yield f"data: 분석 오류: {e}\\n\n"
        yield "data: [DONE]\n\n"


# ============================================================
#  Flask 앱 초기화
# ============================================================
app    = Flask(__name__)
kotra  = KotraClient()
corpus = NewsCorpus()
if kotra.enabled:
    try: build_corpus(kotra, corpus, pages=2)
    except Exception as e: print(f"[코퍼스] 적재 건너뜀: {e}")

ksure    = KSureClient()
analyzer = RagAnalyzer(corpus=corpus, kotra=kotra, ksure=ksure)
# 산업연관(레온티에프) 엔진 - 폴백 모드면 내장 업계평균 투입구조 사용
io_engine = None
if _HAS_ENGINES:
    try:
        io_engine = _io_analysis.IOAnalysisEngine()
    except Exception as e:
        print(f"[engines] IO 엔진 초기화 실패: {e}")

engine   = SupplyChainEngine(rag_client=analyzer.client, models=GEMINI_MODELS,
                             ksure=ksure, io_engine=io_engine, country_grades={})
threading.Thread(target=engine.sync_data, daemon=True).start()

# 공공데이터 K-SURE 국별신용등급(1~7) - 백그라운드 로드(기동 차단 방지)
def _load_country_grades_bg():
    if _data_sources is None: return
    _dg_key = _read_txt("data_go_kr_key.txt") or os.environ.get("DATA_GO_KR_KEY", "")
    if not _dg_key:
        print("[data_sources] data_go_kr_key.txt 없음 → K-SURE RI(1~5)로 국가리스크 산정"); return
    try:
        # 라이브 성공 시에만 1~7등급 사용. 실패하면 빈 채로 두고,
        # _country_risk가 K-SURE 리스크인덱스(RI 1~5, 162개국 라이브)로 대체한다.
        # (12개국 하드코딩 폴백표는 라이브 RI보다 못하므로 쓰지 않음)
        grades = _data_sources.fetch_kosure_grades(service_key=_dg_key)
        if grades:
            engine.country_grades = grades
            print(f"[data_sources] 국별신용등급(1~7) 라이브 {len(grades)}개국 로드")
    except Exception as e:
        print(f"[data_sources] 국별신용등급 라이브 실패 → K-SURE RI(1~5)로 대체 ({str(e)[:60]})")
threading.Thread(target=_load_country_grades_bg, daemon=True).start()

print(f"[web_app] 분석 모드: {analyzer.mode}")
_MODE_LABEL = {"rag":"RAG","grounding":"그라운딩","offline":"오프라인"}


# ============================================================
#  HTML 페이지
# ============================================================
PAGE = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TradeRisk AI - 경상도 수출 인텔리전스</title>
<script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{
  --bg:#080c14;--surface:#0e1422;--surface2:#131929;--surface3:#1a2236;
  --border:#1c2538;--border2:#243048;
  --txt:#d4ddf5;--muted:#5a6a8a;--muted2:#8494b8;
  --accent:#4f8ef7;--green:#2dd4a0;--yellow:#f5b942;--red:#f87171;--purple:#a78bfa;
}
body.light{
  --bg:#eaeef6;--surface:#ffffff;--surface2:#f3f6fb;--surface3:#e8edf5;
  --border:#e2e8f2;--border2:#d3dbe9;
  --txt:#1b2438;--muted:#6b7890;--muted2:#48526b;
}
body.light .fx-bar{background:#ffffff}
body.light #md-out p,body.light #md-out ul li,body.light #md-out ol li{color:#33405c}
body.light #md-out ul li{background:rgba(0,0,0,.025)}
body.light #md-out strong{color:#0c2a5c;background:rgba(79,142,247,.12)}
body.light #md-out blockquote{color:#27406e}
body.light #md-out blockquote strong{color:#0e1830}
body.light .hero-sub,body.light .hact,body.light .bt-table td,
body.light .io-meta,body.light .io-imp-name{color:#33405c}
body.light .hero-sub b,body.light .hact b,body.light .io-meta b{color:#0e1830}
body.light .hero-gs{color:#3a475f}
body.light .hero.go .hero-verdict{color:#0f6e56}
body.light .hero.caution .hero-verdict{color:#9a5b00}
body.light .hero.stop .hero-verdict{color:#a32d2d}
body.light .hm{background:rgba(0,0,0,.05)}
body.light .pos{color:#0f8a66}
body.light .neg{color:#cc4444}

/* 간단 카드 모드(사장님·모바일) — 핵심만, 나머지 숨김 */
body.simple .sidebar{display:none!important}
body.simple #res-wrap{display:none!important}
body.simple #evidence-wrap{display:none!important}
body.simple .act-row,body.simple .preset-wrap,body.simple #profile-badge,
body.simple #backtest-section,body.simple #profile-section,
body.simple #history-section,body.simple #compare-section{display:none!important}
body.simple .main{max-width:600px;margin:0 auto}
body.simple .hero{border-radius:18px}
body.simple .hero-verdict{font-size:30px}
body.simple .hero-detail-btn{width:100%;text-align:center;padding:11px;font-size:14px;margin-top:18px}

*{box-sizing:border-box;margin:0;padding:0}
html,body{overflow-x:hidden;width:100%}
body{background:var(--bg);color:var(--txt);
  font-family:'Inter','Malgun Gothic',system-ui,sans-serif;min-height:100vh;display:flex;flex-direction:column;}

/* FX BAR */
.fx-bar{background:#080c14;border-bottom:1px solid var(--border);
  padding:7px 20px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-size:12px;flex-shrink:0}
.fx-pill{display:flex;align-items:center;gap:5px;background:var(--surface);
  border:1px solid var(--border);border-radius:20px;padding:3px 10px}
.fx-pill.active{border-color:var(--accent)}
.fx-pair{color:var(--muted);font-weight:600;font-size:11px}
.fx-val{font-weight:700;font-size:13px;color:#f5b942}
.fx-usd .fx-val{color:var(--accent)}
.fx-chg{font-size:11px}
.pos{color:var(--green)}.neg{color:var(--red)}
.mode-badge{display:inline-flex;padding:2px 9px;border-radius:99px;font-size:11px;font-weight:600;
  background:rgba(79,142,247,.12);color:#7eb8ff;border:1px solid rgba(79,142,247,.2);margin-left:auto}
.btn-shutdown{background:var(--red);color:#fff;border:none;padding:4px 12px;border-radius:6px;
  cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s}
.btn-shutdown:hover{background:#dc2626}
.btn-verify{background:var(--surface3);color:#c4b5fd;border:1px solid rgba(167,139,250,.3);
  padding:4px 12px;border-radius:6px;cursor:pointer;font-size:11px;font-weight:700;
  font-family:inherit;transition:all .15s}
.btn-verify:hover{background:rgba(167,139,250,.18)}

/* 입력카드 액션 버튼 */
.act-row{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}
.btn-act{flex:1;min-width:120px;background:var(--surface2);color:var(--muted2);
  border:1px solid var(--border2);padding:9px;border-radius:8px;cursor:pointer;
  font-size:13px;font-weight:600;font-family:inherit;transition:all .15s}
.btn-act:hover{background:var(--surface3);color:var(--txt);border-color:var(--accent)}
.prof-badge{margin-top:10px;font-size:12px;color:#5eead4;background:rgba(45,212,160,.1);
  border:1px solid rgba(45,212,160,.25);border-radius:7px;padding:7px 11px}

/* 경상도 프리셋 칩 */
.preset-wrap{display:flex;align-items:center;gap:10px;margin:12px 0 4px;flex-wrap:wrap}
.preset-lbl{font-size:11px;font-weight:700;color:#fcd34d;background:rgba(245,185,66,.12);
  border:1px solid rgba(245,185,66,.25);border-radius:6px;padding:3px 9px;white-space:nowrap}
#preset-chips{display:flex;gap:6px;flex-wrap:wrap}
.chip{background:var(--surface2);border:1px solid var(--border2);color:var(--muted2);
  padding:5px 11px;border-radius:99px;cursor:pointer;font-size:12px;transition:all .12s}
.chip:hover{background:var(--surface3);color:var(--txt);border-color:#f5b942}

/* 수출 판정 히어로 카드 */
.hero{border-radius:18px;padding:22px 24px;border:1px solid;position:relative;overflow:hidden}
.hero.go{background:linear-gradient(135deg,rgba(45,212,160,.14),rgba(16,185,129,.06));border-color:rgba(45,212,160,.4)}
.hero.caution{background:linear-gradient(135deg,rgba(245,185,66,.14),rgba(245,158,11,.06));border-color:rgba(245,185,66,.4)}
.hero.stop{background:linear-gradient(135deg,rgba(248,113,113,.14),rgba(239,68,68,.06));border-color:rgba(248,113,113,.4)}
.hero-top{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.hero-verdict{font-size:26px;font-weight:800;line-height:1.2}
.hero.go .hero-verdict{color:#5eead4}
.hero.caution .hero-verdict{color:#fcd34d}
.hero.stop .hero-verdict{color:#fca5a5}
.hero-sub{font-size:14px;color:#c8d8f8;margin-top:8px;line-height:1.6}
.hero-sub b{color:#fff}
.hero-money{margin-top:14px;display:flex;gap:20px;flex-wrap:wrap}
.hm{background:rgba(0,0,0,.18);border-radius:10px;padding:10px 14px;min-width:130px}
.hm-l{font-size:11px;color:var(--muted2);margin-bottom:3px}
.hm-v{font-size:19px;font-weight:800}
.hero-acts{margin-top:16px}
.hero-acts-t{font-size:11px;font-weight:700;color:var(--muted2);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.hact{display:flex;gap:9px;align-items:flex-start;font-size:14px;color:#d4e0f8;
  padding:8px 12px;background:rgba(255,255,255,.04);border-radius:9px;margin-bottom:6px;line-height:1.5}
.hact b{color:#fff}
.hero-gs{margin-top:16px;padding-top:14px;border-top:1px solid rgba(255,255,255,.1);
  font-size:13px;color:#b8c8e8;line-height:1.7}
.hero-gs b{color:#fcd34d}
.hero-gs .gs-port{color:#7eb8ff;font-weight:700}
.hero-detail-btn{margin-top:14px;background:rgba(255,255,255,.06);border:1px solid var(--border2);
  color:var(--muted2);padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:inherit}
.hero-detail-btn:hover{color:var(--txt)}
.hero-adjust{margin-top:12px;padding:11px 14px;border-radius:10px;font-size:13px;line-height:1.6;
  background:rgba(248,113,113,.12);border:1px solid rgba(248,113,113,.3);color:#fecaca}
.hero-adjust b{color:#fff}

/* 근거 펼치기 토글 */
.ev-toggle{width:100%;display:flex;justify-content:space-between;align-items:center;
  background:var(--surface);border:1px solid var(--border);border-radius:12px;
  padding:13px 18px;cursor:pointer;font-size:14px;font-weight:600;color:var(--muted2);
  font-family:inherit;transition:all .15s}
.ev-toggle:hover{color:var(--txt);border-color:var(--accent)}
#evidence-body{display:flex;flex-direction:column;gap:14px;margin-top:14px}
.panel{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px 20px}
.panel h3{font-size:15px;font-weight:700;color:var(--txt);margin-bottom:4px}
.panel .psub{font-size:12px;color:var(--muted2);margin-bottom:14px;line-height:1.6}
/* BOM 편집 행 */
.bom-edit{display:flex;gap:8px;margin-bottom:7px;align-items:center}
.bom-edit input{background:var(--bg);border:1px solid var(--border2);color:var(--txt);
  padding:8px 11px;border-radius:7px;font-size:13px;font-family:inherit;outline:none}
.bom-edit input:focus{border-color:var(--accent)}
.bom-edit .be-mat{flex:1}
.bom-edit .be-ratio{width:90px}
.bom-edit .be-del{background:rgba(248,113,113,.12);color:#fca5a5;border:none;
  border-radius:7px;width:34px;cursor:pointer;font-size:15px}
.be-total{font-size:13px;font-weight:700;margin:10px 0}
.be-actions{display:flex;gap:8px;margin-top:14px}
.btn-save{background:linear-gradient(135deg,#2dd4a0,#10b981);color:#04221a;border:none;
  padding:10px 18px;border-radius:9px;cursor:pointer;font-size:14px;font-weight:700;font-family:inherit}
.btn-ghost{background:var(--surface2);color:var(--muted2);border:1px solid var(--border2);
  padding:10px 16px;border-radius:9px;cursor:pointer;font-size:14px;font-family:inherit}
/* 추세 미니바 */
.trend-row{display:flex;align-items:flex-end;gap:4px;height:90px;margin:10px 0;padding:0 2px}
.trend-bar{flex:1;border-radius:3px 3px 0 0;min-height:4px;position:relative;transition:opacity .15s}
.trend-bar:hover{opacity:.8}
.hist-item{display:flex;justify-content:space-between;align-items:center;font-size:13px;
  padding:8px 10px;border-bottom:1px solid var(--border)}
.hist-item:last-child{border:none}
.hg{font-weight:700;font-size:11px;padding:1px 8px;border-radius:99px}
/* 비교표 */
.cmp-table{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}
.cmp-table th{text-align:center;color:var(--muted2);font-size:11px;font-weight:700;
  padding:8px;border-bottom:1px solid var(--border2);text-transform:uppercase}
.cmp-table td{padding:10px 8px;border-bottom:1px solid var(--border);text-align:center}
.cmp-table td.cmp-ctr{text-align:left;font-weight:700;color:var(--txt)}
.cmp-best{background:rgba(45,212,160,.07)}
.cmp-inp{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}
.cmp-inp input{flex:1;min-width:160px;background:var(--bg);border:1px solid var(--border2);
  color:var(--txt);padding:9px 12px;border-radius:8px;font-size:13px;font-family:inherit;outline:none}

/* 백테스트 검증 패널 */
.bt-wrap{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px 20px}
.bt-head{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.bt-head h3{font-size:15px;font-weight:700;color:var(--txt)}
.bt-sub{font-size:12px;color:var(--muted2);margin-bottom:14px;line-height:1.6}
.bt-sub b{color:#c4b5fd}
.bt-table{width:100%;border-collapse:collapse;font-size:13px}
.bt-table th{text-align:left;color:var(--muted2);font-size:11px;font-weight:700;
  text-transform:uppercase;letter-spacing:.4px;padding:8px 10px;border-bottom:1px solid var(--border2)}
.bt-table td{padding:10px;border-bottom:1px solid var(--border);color:#b8c8e8;vertical-align:top}
.bt-case{font-weight:600;color:var(--txt)}
.bt-note{font-size:11px;color:var(--muted);margin-top:3px;line-height:1.5}
.bt-vchip{display:inline-block;font-size:10px;font-weight:600;padding:1px 7px;border-radius:99px}
.bt-v-ok{background:rgba(45,212,160,.15);color:#5eead4}
.bt-v-no{background:rgba(245,185,66,.15);color:#fcd34d}
.bt-grade{font-weight:700}
.bt-match{font-size:16px}
#overlay{display:none;position:fixed;inset:0;background:rgba(8,12,20,.96);z-index:9999;
  flex-direction:column;align-items:center;justify-content:center;text-align:center}
#overlay h1{color:var(--green);font-size:22px;margin-bottom:8px}
#overlay p{color:var(--muted2);font-size:14px}

/* LAYOUT */
.layout{display:flex;flex:1;min-height:0;height:calc(100vh - 36px)}
.sidebar{width:240px;background:var(--surface);border-right:1px solid var(--border);
  display:flex;flex-direction:column;flex-shrink:0;overflow:hidden}
.sb-hdr{padding:10px 12px;border-bottom:1px solid var(--border);font-size:11px;
  color:var(--muted2);font-weight:700;letter-spacing:.5px;text-transform:uppercase}
.sb-search{width:calc(100% - 16px);margin:8px;padding:7px 10px;background:var(--bg);
  border:1px solid var(--border2);color:var(--txt);border-radius:7px;font-size:12px;outline:none}
.sb-search:focus{border-color:var(--accent)}
.c-list{flex:1;overflow-y:auto}
.c-item{padding:7px 12px;cursor:pointer;border-bottom:1px solid var(--border);
  display:flex;justify-content:space-between;align-items:center;font-size:12px;transition:background .1s}
.c-item:hover{background:var(--surface2)}
.c-chg{font-weight:700;font-size:11px}

/* MAIN */
.main{flex:1;overflow-y:auto;padding:18px 22px;display:flex;flex-direction:column;gap:14px}

/* INPUT CARD */
.inp-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px}
.inp-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}
.field label{font-size:11px;color:var(--muted2);font-weight:700;letter-spacing:.4px;
  text-transform:uppercase;display:block;margin-bottom:5px}
.field input{width:100%;background:var(--bg);border:1px solid var(--border2);color:var(--txt);
  padding:10px 13px;border-radius:9px;font-size:14px;font-family:inherit;outline:none;transition:all .2s}
.field input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(79,142,247,.13)}
.field input::placeholder{color:var(--muted)}
.btn-main{width:100%;padding:12px;border:0;border-radius:10px;font-size:15px;font-weight:700;
  font-family:inherit;cursor:pointer;transition:all .2s;
  background:linear-gradient(135deg,#4f8ef7,#6c5cfc);color:#fff;
  box-shadow:0 4px 16px rgba(79,142,247,.3)}
.btn-main:hover{transform:translateY(-1px);box-shadow:0 6px 22px rgba(79,142,247,.4)}
.btn-main:disabled{opacity:.5;cursor:wait;transform:none!important}

/* DATA CARDS ROW — 2×2 배치로 카드 폭 2배 확보 */
.data-row{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}
.d-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px;overflow:hidden}
.d-card-title{font-size:12px;color:var(--muted2);font-weight:700;letter-spacing:.5px;
  text-transform:uppercase;margin-bottom:11px}
.d-card-val{font-size:24px;font-weight:700;color:var(--txt);margin-bottom:4px}
.d-card-sub{font-size:13px;color:var(--muted)}
.bom-row{display:flex;align-items:center;gap:9px;margin-bottom:8px;font-size:13px}
.bom-mat{width:96px;color:var(--muted2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bom-track{flex:1;height:7px;background:var(--surface3);border-radius:4px;overflow:hidden}
.bom-fill{height:100%;border-radius:4px}
.bom-pct{width:36px;text-align:right;color:var(--muted);font-size:12px}
.bom-chg{width:62px;text-align:right;font-weight:700;font-size:12px}
.hs-row{padding:7px 0;border-bottom:1px solid var(--border);font-size:14px}
.hs-row:last-child{border:none}
.hs-code{color:var(--accent);font-weight:700;font-size:15px}
.hs-tariff{color:var(--green);font-size:13px;margin-top:3px}
.score-mini{display:grid;grid-template-columns:1fr 1fr;gap:9px}
.sc{background:var(--surface2);border-radius:8px;padding:11px 9px;text-align:center}
.sc-label{font-size:12px;color:var(--muted);margin-bottom:5px}
.sc-val{font-size:24px;font-weight:700}
.sc-basis{font-size:11px;color:var(--muted2);margin-top:5px;line-height:1.4}
.c-green{color:var(--green)}.c-yellow{color:var(--yellow)}.c-red{color:var(--red)}

/* AHP 종합 판정 배너 */
.gb{display:flex;align-items:center;gap:16px;padding:16px 20px;border-radius:14px;
  border:1px solid var(--border);background:var(--surface)}
.gb-badge{font-size:22px;font-weight:800;padding:8px 18px;border-radius:10px;white-space:nowrap}
.gb-num{font-size:13px;color:var(--muted2);margin-top:3px}
.gb-axes{display:flex;gap:18px;margin-left:auto;flex-wrap:wrap}
.gb-ax{text-align:center}
.gb-ax-l{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}
.gb-ax-v{font-size:17px;font-weight:700;color:var(--txt)}
.gb-ax-w{font-size:10px;color:var(--muted)}
.g-정상{background:rgba(45,212,160,.15);color:#5eead4}
.g-주의{background:rgba(245,185,66,.15);color:#fcd34d}
.g-경고{background:rgba(248,140,80,.15);color:#fb923c}
.g-위험{background:rgba(248,113,113,.15);color:#fca5a5}

/* IO 가격파급 섹션 */
.io-wrap{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px 20px}
.io-title{font-size:12px;font-weight:700;color:var(--muted2);letter-spacing:.5px;
  text-transform:uppercase;margin-bottom:12px;display:flex;align-items:center;gap:8px}
.io-tag{font-size:10px;font-weight:600;padding:2px 8px;border-radius:99px;
  background:rgba(167,139,250,.15);color:#c4b5fd}
.io-meta{font-size:13px;color:var(--muted2);margin-bottom:12px;line-height:1.6}
.io-meta b{color:var(--txt)}
.io-imp-row{display:flex;align-items:center;gap:10px;margin-bottom:7px;font-size:13px}
.io-imp-name{flex:1;color:var(--muted2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.io-imp-hs{font-size:11px;color:var(--muted);font-family:monospace}
.io-imp-track{width:120px;height:6px;background:var(--surface3);border-radius:3px;overflow:hidden}
.io-imp-fill{height:100%;background:linear-gradient(90deg,#a78bfa,#f87171);border-radius:3px}
.io-imp-pct{width:62px;text-align:right;font-weight:700;color:#fca5a5}

/* AI REPORT */
.result-wrap{background:var(--surface);border:1px solid var(--border);border-radius:16px;overflow:hidden}
.r-header{display:flex;align-items:center;gap:8px;padding:13px 18px;
  background:var(--surface2);border-bottom:1px solid var(--border)}
.pulse{width:7px;height:7px;border-radius:50%;background:var(--green);flex-shrink:0}
.pulse.on{animation:pulse 1.2s ease-in-out infinite}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.3;transform:scale(.6)}}
.r-title{font-size:13px;font-weight:600}
.r-sub{font-size:11px;color:var(--muted);margin-left:auto}
#md-out{padding:22px}

/* MARKDOWN */
#md-out h2{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:700;
  letter-spacing:.5px;text-transform:uppercase;margin:26px 0 11px;
  padding:9px 13px;border-radius:8px;border-left:3px solid}
#md-out h2:first-child{margin-top:0}
#md-out h2:nth-of-type(1){background:rgba(248,113,113,.08);color:#fca5a5;border-color:#f87171}
#md-out h2:nth-of-type(2){background:rgba(79,142,247,.1);color:#7eb8ff;border-color:#4f8ef7}
#md-out h2:nth-of-type(3){background:rgba(45,212,160,.1);color:#5eead4;border-color:#2dd4a0}
#md-out h2:nth-of-type(4){background:rgba(245,185,66,.1);color:#fcd34d;border-color:#f5b942}
#md-out h2:nth-of-type(5){background:rgba(167,139,250,.1);color:#c4b5fd;border-color:#a78bfa}
#md-out h2:nth-of-type(6){background:rgba(79,142,247,.07);color:#7eb8ff;border-color:#4f8ef7}
#md-out h2:nth-of-type(7){background:rgba(45,212,160,.1);color:#5eead4;border-color:#2dd4a0}
#md-out p{font-size:14px;line-height:1.8;color:#b8c8e8;margin-bottom:10px}
#md-out ul{list-style:none;padding:0;margin-bottom:13px}
#md-out ul li{position:relative;padding:7px 11px 7px 30px;font-size:14px;line-height:1.7;
  color:#b0c0e0;border-radius:7px;margin-bottom:4px;background:rgba(255,255,255,.022)}
#md-out ul li:hover{background:rgba(255,255,255,.038)}
#md-out ul li::before{content:'>';position:absolute;left:10px;top:7px;
  font-size:15px;line-height:1.5;color:var(--accent);font-weight:700}
#md-out ol{padding-left:18px;margin-bottom:13px}
#md-out ol li{font-size:14px;line-height:1.7;color:#b0c0e0;margin-bottom:4px}
#md-out strong{color:#e8f0ff;font-weight:600;background:rgba(79,142,247,.12);
  padding:1px 4px;border-radius:4px}
#md-out blockquote{margin:14px 0;padding:14px 16px;
  background:linear-gradient(135deg,rgba(79,142,247,.08),rgba(108,92,252,.08));
  border:1px solid rgba(79,142,247,.2);border-radius:11px;border-left:4px solid var(--accent);
  font-size:14px;line-height:1.75;color:#c8d8f8}
#md-out blockquote strong{background:none;padding:0;color:#fff;font-size:15px}
#md-out hr{border:none;border-top:1px solid var(--border2);margin:18px 0}
#md-out a{color:#7eb8ff;text-decoration:none;border-bottom:1px solid rgba(126,184,255,.3)}
.src-block{margin:0 18px 18px;padding:12px 14px;background:var(--surface2);
  border:1px solid var(--border);border-radius:9px;font-size:12px;color:var(--muted2)}
.src-block a{color:#5a8dee;display:block;margin-top:3px}

/* SINGLE COMMODITY REPORT */
.cr-wrap{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px}
.ai-rep{font-size:12px;line-height:1.75;color:var(--muted2);margin-top:8px;
  border:1px solid var(--border);border-radius:8px;padding:10px;max-height:220px;overflow-y:auto}

/* SPINNER */
.spin{width:15px;height:15px;border:2px solid var(--border2);border-top-color:var(--accent);
  border-radius:50%;animation:sp .7s linear infinite;display:inline-block;vertical-align:middle;margin-right:6px}
@keyframes sp{to{transform:rotate(360deg)}}
</style>
</head>
<body>

<!-- 상단 환율 바 -->
<div class="fx-bar">
  <span style="font-size:10px;color:var(--muted);font-weight:700">실시간 환율</span>
  <div class="fx-pill fx-usd">
    <span class="fx-pair">USD/KRW</span>
    <span class="fx-val" id="fx-usd-v">--</span>
    <span class="fx-chg" id="fx-usd-c"></span>
  </div>
  <div class="fx-pill active" id="fx-tgt" style="display:none">
    <span class="fx-pair" id="fx-tgt-p">--</span>
    <span class="fx-val" id="fx-tgt-v">--</span>
    <span class="fx-chg" id="fx-tgt-c"></span>
  </div>
  <span style="font-size:10px;color:var(--muted)" id="fx-src"></span>
  <span class="mode-badge">{{mode}}</span>
  <button class="btn-verify" id="simple-btn" onclick="toggleSimple()">🪪 간단</button>
  <button class="btn-verify" id="theme-btn" onclick="toggleTheme()">🌙 다크</button>
  <button class="btn-verify" onclick="toggleBacktest()">📋 모델 검증</button>
  <button class="btn-shutdown" onclick="doShutdown()">⏹ 종료</button>
</div>

<!-- 종료 오버레이 -->
<div id="overlay">
  <h1>✅ 서버가 정상 종료되었습니다.</h1>
  <p>이 창을 닫으셔도 됩니다.</p>
</div>

<div class="layout">
  <!-- 사이드바: 원자재 시세 -->
  <div class="sidebar">
    <div class="sb-hdr">원자재 시세</div>
    <input class="sb-search" id="sb-q" placeholder="검색..." oninput="filterC(this.value)">
    <div class="c-list" id="c-list">
      <div style="padding:16px;color:var(--muted);font-size:12px;text-align:center">
        <span class="spin"></span>로딩 중...
      </div>
    </div>
  </div>

  <!-- 메인 컨텐츠 -->
  <div class="main">

    <!-- 모델 검증 패널 (백테스트) -->
    <div id="backtest-section" style="display:none"></div>

    <!-- 입력 카드 -->
    <div class="inp-card">
      <div class="inp-row">
        <div class="field">
          <label>품목 / 완제품</label>
          <input id="i-mat" placeholder="예: 선풍기, 철판, 구리배관">
        </div>
        <div class="field">
          <label>수출 대상국</label>
          <input id="i-ctr" placeholder="예: 베트남, 미국, 독일" oninput="onCtr(this.value)">
        </div>
      </div>
      <div class="preset-wrap">
        <span class="preset-lbl">경상도 주력품</span>
        <div id="preset-chips"></div>
      </div>
      <button class="btn-main" id="btn-go" onclick="runAll()">🔍 통합 분석 시작</button>
      <div class="act-row">
        <button class="btn-act" onclick="toggleProfile()">🏭 내 회사 BOM</button>
        <button class="btn-act" onclick="toggleHistory()">📈 분석 이력</button>
        <button class="btn-act" onclick="toggleCompare()">🌍 다국가 비교</button>
      </div>
      <div id="profile-badge" class="prof-badge" style="display:none"></div>
    </div>

    <!-- 회사 BOM 편집 패널 -->
    <div id="profile-section" style="display:none"></div>
    <!-- 분석 이력/추세 패널 -->
    <div id="history-section" style="display:none"></div>
    <!-- 다국가 비교 패널 -->
    <div id="compare-section" style="display:none"></div>

    <!-- ① 수출 판정 히어로 카드 (진단서 표지) -->
    <div id="hero-section" style="display:none"></div>

    <!-- ② AI 진단 보고서 (히어로 바로 밑 = 본문) -->
    <div class="result-wrap" id="res-wrap" style="display:none">
      <div class="r-header">
        <div class="pulse" id="pulse"></div>
        <span class="r-title" id="r-title">분석 준비 중...</span>
        <span class="r-sub" id="r-sub"></span>
      </div>
      <div id="md-out"></div>
      <div id="src-blk"></div>
    </div>

    <!-- ③ 근거 데이터 (펼치기) -->
    <div id="evidence-wrap" style="display:none">
      <button class="ev-toggle" id="ev-toggle" onclick="toggleEvidence()">
        <span>📊 수집 데이터 · 점수 근거</span>
        <span id="ev-arrow">▾</span>
      </button>
      <div id="evidence-body" style="display:none">
        <div id="grade-banner"></div>
        <div class="data-row" id="data-row">
          <div class="d-card">
            <div class="d-card-title">환율 현황</div>
            <div class="d-card-val" id="dc-fx-val">--</div>
            <div class="d-card-sub" id="dc-fx-sub">--</div>
            <div id="dc-fx-detail" style="margin-top:6px;font-size:12px"></div>
          </div>
          <div class="d-card">
            <div class="d-card-title">원가 영향 (BOM)</div>
            <div id="dc-bom"></div>
          </div>
          <div class="d-card">
            <div class="d-card-title">HS코드 · 관세</div>
            <div id="dc-hs"></div>
          </div>
          <div class="d-card">
            <div class="d-card-title">공급망 리스크 (AHP)</div>
            <div class="score-mini" id="dc-scores"></div>
          </div>
        </div>
        <div id="io-section"></div>
      </div>
    </div>

    <!-- 원자재 단일 진단 (사이드바 클릭 시) -->
    <div id="cr-section" style="display:none">
      <div class="cr-wrap">
        <div id="cr-content"></div>
      </div>
    </div>

  </div>
</div>

<script>
marked.use({gfm:true,breaks:true});
let _allC=[], _curCtr='미국';

/* ── 환율 ── */
function loadFx(country){
  fetch('/api/fx?country=' + encodeURIComponent('미국')).then(r=>r.json()).then(d=>{
    document.getElementById('fx-usd-v').textContent = d.rate.toLocaleString();
    let el = document.getElementById('fx-usd-c');
    if(d.change_pct != null){
      el.textContent = (d.change_pct>=0?'+':'')+d.change_pct.toFixed(2)+'%';
      el.className = 'fx-chg '+(d.change_pct>=0?'pos':'neg');
    }
    document.getElementById('fx-src').textContent = '출처: '+(d.source||'');
  }).catch(()=>{});

  if(!country || country==='미국') return;
  fetch('/api/fx?country=' + encodeURIComponent(country)).then(r=>r.json()).then(d=>{
    let pill = document.getElementById('fx-tgt');
    document.getElementById('fx-tgt-p').textContent = d.pair||'';
    document.getElementById('fx-tgt-v').textContent =
      (d.decimals>=3) ? d.rate.toFixed(d.decimals) : d.rate.toLocaleString();
    let el = document.getElementById('fx-tgt-c');
    if(d.change_pct != null){
      el.textContent = (d.change_pct>=0?'+':'')+d.change_pct.toFixed(2)+'%';
      el.className = 'fx-chg '+(d.change_pct>=0?'pos':'neg');
    }
    pill.style.display = 'flex';
  }).catch(()=>{});
}
function onCtr(v){ _curCtr = v.trim()||'미국'; if(v.trim()) loadFx(v.trim()); }

/* ── 원자재 사이드바 ── */
function loadC(){
  fetch('/api/search').then(r=>r.json()).then(items=>{
    _allC=items; renderC(items);
  }).catch(()=>{});
}
function renderC(items){
  let h='';
  items.forEach(v=>{
    let c=v.c||0, cl=c>=0?'neg':'pos';
    h+=`<div class="c-item" onclick="loadCReport('${esc(v.name)}',${c},${v.p})">
      <span>${esc(v.name)}</span>
      <span class="c-chg ${cl}">${c>=0?'+':''}${c.toFixed(2)}%</span>
    </div>`;
  });
  document.getElementById('c-list').innerHTML = h ||
    '<div style="padding:14px;color:var(--muted);font-size:12px">없음</div>';
}
function filterC(q){
  let f=q.trim().toLowerCase();
  renderC(f ? _allC.filter(v=>v.name.toLowerCase().includes(f)) : _allC);
}

/* ── 통합 분석 메인 흐름 ── */
async function runAll(){
  let mat = document.getElementById('i-mat').value.trim();
  let ctr = document.getElementById('i-ctr').value.trim() || '미국';
  if(!mat){ alert('품목을 입력하세요'); return; }

  let btn = document.getElementById('btn-go');
  btn.disabled=true; btn.textContent='데이터 수집 중...';
  document.getElementById('cr-section').style.display='none';

  /* 1단계: 시뮬레이터 (BOM + HS + FX + 리스크 점수) */
  let simData = {};
  try{
    let sr = await fetch('/api/simulate',{method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({product:mat, country:ctr})});
    simData = await sr.json();
    renderDataCards(simData);
    document.getElementById('evidence-wrap').style.display='block';
  }catch(e){
    document.getElementById('evidence-wrap').style.display='none';
  }

  /* 2단계: AI 통합 보고서 스트리밍 */
  btn.textContent='AI 분석 중...';
  let wrap = document.getElementById('res-wrap');
  wrap.style.display='block';
  document.getElementById('md-out').innerHTML =
    '<div style="padding:18px;color:var(--muted)"><span class="spin"></span>AI 통합 보고서 생성 중...</div>';
  document.getElementById('src-blk').innerHTML='';
  document.getElementById('pulse').classList.add('on');
  document.getElementById('r-title').textContent = mat+' → '+ctr+' 통합 분석';
  document.getElementById('r-sub').textContent='';

  let url = '/api/analyze_unified?' + new URLSearchParams({
    material: mat, country: ctr, sim: JSON.stringify(simData)
  });
  let buffer='';
  try{
    let res = await fetch(url);
    let reader = res.body.getReader(), dec = new TextDecoder();
    while(true){
      let {value, done} = await reader.read();
      if(done) break;
      buffer += dec.decode(value, {stream:true});
      renderMd(buffer);
    }
    document.getElementById('r-title').textContent = mat+' → '+ctr+' 분석 완료';
    document.getElementById('r-sub').textContent = new Date().toLocaleTimeString('ko-KR');
    reconcileHeroWithAI(buffer);   // AI 신호(반덤핑 등 뉴스 리스크)로 히어로 판정 보정
  }catch(e){
    document.getElementById('md-out').innerHTML =
      '<p style="color:var(--red);padding:18px">오류: '+esc(String(e))+'</p>';
  }
  document.getElementById('pulse').classList.remove('on');
  btn.disabled=false; btn.textContent='🔍 통합 분석 시작';
}

function renderDataCards(d){
  /* FX */
  let fxChg = d.fx_change||0;
  document.getElementById('dc-fx-val').textContent = (d.fx_rate||0).toLocaleString()+' 원';
  document.getElementById('dc-fx-sub').textContent = d.fx_pair||'USD/KRW';
  document.getElementById('dc-fx-detail').innerHTML =
    '<span class="'+(fxChg>=0?'pos':'neg')+'" style="font-size:15px;font-weight:700">'+
    (fxChg>=0?'+':'')+fxChg.toFixed(2)+'%</span> '+
    '<span style="color:var(--muted);font-size:13px">'+(d.fx_source||'offline')+'</span>';

  /* BOM */
  let bom=d.bom_details||[], colors=['#4f8ef7','#2dd4a0','#f5b942','#a78bfa','#f87171'];
  let total=bom.reduce((s,b)=>s+b.ratio,0)||1;
  let srcReal=d.bom_source==='회사 실제';
  let bh=`<div style="font-size:11px;font-weight:600;margin-bottom:8px;color:${srcReal?'#5eead4':'var(--muted)'}">`+
    (srcReal?'🏭 회사 실제 BOM':'🤖 AI 추정 BOM')+'</div>';
  bh+=bom.map((b,i)=>`
    <div class="bom-row">
      <div class="bom-mat">${esc(b.material.split(' ')[0])}</div>
      <div class="bom-track"><div class="bom-fill" style="width:${b.ratio/total*100}%;background:${colors[i%5]}"></div></div>
      <div class="bom-pct">${b.ratio.toFixed(0)}%</div>
      <div class="bom-chg ${b.change>=0?'neg':'pos'}">${b.change>=0?'+':''}${b.change.toFixed(1)}%</div>
    </div>`).join('');
  let mi=d.raw_impact||0, fi=d.fx_impact||0, ti=d.total_impact||0;
  bh+=`<div style="margin-top:10px;padding-top:9px;border-top:1px solid var(--border);font-size:14px">
    <div style="display:flex;justify-content:space-between;margin-bottom:4px">
      <span style="color:var(--muted)">자재 변동</span>
      <span class="${mi>=0?'neg':'pos'}" style="font-weight:700">${mi>=0?'+':''}${mi.toFixed(2)}%</span>
    </div>
    <div style="display:flex;justify-content:space-between;margin-bottom:4px">
      <span style="color:var(--muted)">환율 효과</span>
      <span class="${fi>=0?'pos':'neg'}" style="font-weight:700">${fi>=0?'+':''}${fi.toFixed(2)}%</span>
    </div>
    <div style="display:flex;justify-content:space-between;font-weight:700;font-size:15px;
      padding-top:6px;border-top:1px solid var(--border)">
      <span>실질 변동</span>
      <span class="${ti<=0?'pos':'neg'}">${ti<=0?'':'+'}${ti.toFixed(2)}%</span>
    </div>
  </div>`;
  document.getElementById('dc-bom').innerHTML=bh;

  /* HS 코드 */
  let hs=d.hs_candidates||[];
  let hh=hs.slice(0,3).map(c=>`
    <div class="hs-row">
      <div class="hs-code">${esc(c.code)}</div>
      <div style="font-size:13px;color:var(--muted2)">${esc(c.desc)}</div>
      <div class="hs-tariff">${esc(c.tariff)}</div>
    </div>`).join('');
  document.getElementById('dc-hs').innerHTML =
    hh||'<div style="color:var(--muted);font-size:12px">조회 실패</div>';

  /* 리스크 점수 + 근거 한 줄 */
  let ps=d.product_scores||{}, bl=ps.basis_lines||{};
  let sh=['country','port','item','delivery'].map(k=>{
    let v=ps[k]||0;
    let cl=v>=70?'c-green':v>=40?'c-yellow':'c-red';
    let lb={country:'국가',port:'항만',item:'품목',delivery:'납기'}[k];
    let basis=bl[k]?`<div class="sc-basis">${esc(bl[k])}</div>`:'';
    return `<div class="sc"><div class="sc-label">${lb}</div><div class="sc-val ${cl}">${v}</div>${basis}</div>`;
  }).join('');
  document.getElementById('dc-scores').innerHTML=sh;

  /* 수출 판정 히어로 카드 (사장님용) */
  renderHero(d);
  /* AHP 종합 판정 배너 (기술 근거 — 기본 숨김, 히어로에서 토글) */
  renderGradeBanner(ps);
  /* 산업연관 가격파급 */
  renderIO(d.io_analysis);
}

/* ── 수출 판정 히어로 카드 ── */
function toggleGradeBanner(){
  let el=document.getElementById('grade-banner');
  el.style.display = el.style.display==='block' ? 'none' : 'block';
}
function renderHero(d){
  let el=document.getElementById('hero-section');
  let ps=d.product_scores||{};
  let grade=ps.grade;
  if(grade==null){ el.style.display='none'; return; }
  let cls = grade==='정상'?'go' : grade==='주의'?'caution' : 'stop';
  let verdict = grade==='정상'?'✅ 수출하기 좋은 여건입니다'
              : grade==='주의'?'⚠️ 신중하게 접근하세요'
              : grade==='경고'?'🟠 리스크가 큽니다 — 대비 필수'
              : '🛑 지금은 보류를 권합니다';
  // 발목 잡는 요인 = 점수 가장 낮은 타일
  let tiles=[['country','국가·통관 리스크',ps.country],['port','항만·물류',ps.port],
             ['item','원자재·시장',ps.item],['delivery','납기·지연',ps.delivery]];
  let worst=tiles.filter(t=>t[2]!=null).sort((a,b)=>a[2]-b[2])[0]||['','',0];
  let drag = worst[1]?`<b>${esc(worst[1])}</b>가 가장 발목을 잡고 있어요 (${worst[2]}점)`:'';
  // 돈 관점
  let ti=d.total_impact, fi=d.fx_change, mi=d.raw_impact;
  let marginTxt = ti!=null ? (ti<=0?`마진 약 ${Math.abs(ti).toFixed(1)}% 개선 여력`:`마진 약 ${ti.toFixed(1)}% 잠식 위험`) : '-';
  // 액션 3개 (점수 기반)
  let acts=[];
  if(ps.country!=null && ps.country<50) acts.push(['🛡️','무역보험(K-SURE) 부보 검토 — 대금 미회수 리스크 대비']);
  if(fi!=null && fi>3) acts.push(['💱',`선물환 헤지 검토 — 환율로 마진 ${fi.toFixed(1)}% 출렁`]);
  if(ps.port!=null && ps.port<45) acts.push(['🚢',`복수 운송경로·납기 버퍼 확보`]);
  if(ps.item!=null && ps.item<50) acts.push(['📦','원자재 선구매·가격 고정(헤지) 검토']);
  if((d.tariff_pct||0)>3) acts.push(['📜',`FTA 원산지증명으로 관세 ${d.tariff_pct}% 절감 추진`]);
  if(!acts.length) acts.push(['👍','현재 리스크 낮음 — 계약·납기 조건만 표준 점검']);
  acts=acts.slice(0,3);
  // 경상도 컨텍스트
  let gs=d.gyeongsang||{};
  let gsHtml = gs.region?`<div class="hero-gs">📍 <b>${esc(gs.region)}</b> ${esc(gs.cluster||'')} 품목 ·
      주력 수출항 <span class="gs-port">${esc(gs.port||'')}</span><br>
      💡 활용 가능: ${esc(gs.support||'')}</div>`:'';
  el.className=`hero ${cls}`;
  el.innerHTML=`
    <div class="hero-top"><div class="hero-verdict">${verdict}</div></div>
    <div class="hero-sub">${esc(d.product||'')||'품목'} → <b>${esc(d.country||'')||'대상국'}</b> 수출 종합 판단. ${drag}</div>
    <div class="hero-money">
      <div class="hm"><div class="hm-l">실질 수익성 (원가+환율)</div>
        <div class="hm-v ${ti<=0?'pos':'neg'}">${marginTxt}</div></div>
      <div class="hm"><div class="hm-l">자재 변동 / 환율 효과</div>
        <div class="hm-v" style="font-size:15px;color:var(--muted2)">${mi>=0?'+':''}${(mi||0).toFixed(1)}% / ${fi>=0?'+':''}${(fi||0).toFixed(1)}%</div></div>
    </div>
    <div class="hero-acts"><div class="hero-acts-t">지금 할 일</div>
      ${acts.map(a=>`<div class="hact"><span>${a[0]}</span><span>${esc(a[1])}</span></div>`).join('')}
    </div>
    ${gsHtml}
    <button class="hero-detail-btn" onclick="heroDetail()">📄 상세 보고서·근거 보기 ▾</button>`;
  el.style.display='block';
}

function heroDetail(){
  if(document.body.classList.contains('simple')){
    applySimple(false);                       // 간단 → 상세 진단서로 전환
    setTimeout(()=>document.getElementById('res-wrap').scrollIntoView({behavior:'smooth',block:'start'}),60);
  } else {
    toggleEvidence(true);
  }
}

function toggleEvidence(scroll){
  let b=document.getElementById('evidence-body');
  let open = b.style.display!=='block';
  b.style.display = open?'block':'none';
  let ar=document.getElementById('ev-arrow'); if(ar) ar.textContent=open?'▴':'▾';
  if(open && scroll) b.scrollIntoView({behavior:'smooth',block:'nearest'});
}

/* AI 보고서의 수출신호(GO/CAUTION/STOP)로 히어로 판정 보정.
   AHP 점수는 가격·항만·국가만 보고 통상정책(반덤핑·관세장벽) 리스크는 못 보므로,
   AI가 뉴스 기반으로 더 보수적이면 그쪽을 따른다(안전 우선). */
function reconcileHeroWithAI(text){
  let el=document.getElementById('hero-section');
  if(el.style.display==='none') return;
  let m=text.match(/수출\s*신호[:：\s\]\[]*\**\s*(GO|CAUTION|STOP)/i);
  if(!m) return;
  let aiSig=m[1].toUpperCase();
  let aiSev={GO:0,CAUTION:1,STOP:2}[aiSig];
  let curSev = el.classList.contains('stop')?2 : el.classList.contains('caution')?1 : 0;
  if(aiSev<=curSev) return;   // AI가 더 보수적일 때만 보정(안전 방향)
  let cls=aiSev===2?'stop':'caution';
  let verdict=aiSig==='STOP'?'🛑 AI 판단: 보류 권고':'⚠️ AI 판단: 신중하게 접근';
  el.className='hero '+cls;
  let v=el.querySelector('.hero-verdict'); if(v) v.textContent=verdict;
  // 보정 사유 배너 삽입(중복 방지)
  if(!el.querySelector('.hero-adjust')){
    let note=document.createElement('div');
    note.className='hero-adjust';
    note.innerHTML='🤖 <b>AI가 최신 뉴스를 반영해 등급을 낮췄습니다</b> — 점수(가격·항만·국가)만으론 안전해 보이지만, AI 보고서가 통상정책 리스크(반덤핑·관세장벽 등)를 감지했습니다. 아래 보고서를 꼭 확인하세요.';
    let top=el.querySelector('.hero-top');
    if(top) top.insertAdjacentElement('afterend', note);
  }
}

function renderGradeBanner(ps){
  let el=document.getElementById('grade-banner');
  if(ps.grade==null||ps.composite_risk==null){ el.style.display='none'; return; }
  let ax=ps.axes||{}, w=ps.weights||{};
  let axHtml=[['price','가격','price'],['port','항만','port'],['country','국가','country']]
    .map(([k,lb])=>`<div class="gb-ax">
        <div class="gb-ax-l">${lb}</div>
        <div class="gb-ax-v">${(ax[k]!=null?ax[k]:'-')}</div>
        <div class="gb-ax-w">w ${w[k]!=null?w[k]:'-'}</div>
      </div>`).join('');
  el.innerHTML=`<div class="gb">
      <div>
        <div class="gb-badge g-${esc(ps.grade)}">${esc(ps.grade)}</div>
        <div class="gb-num">AHP 종합리스크 <b style="color:var(--txt)">${ps.composite_risk}</b>/100
          <span style="color:var(--muted)">(낮을수록 안전)</span></div>
      </div>
      <div class="gb-axes">${axHtml}</div>
    </div>`;
  el.style.display='block';   // evidence-body가 펼침 여부를 제어
}

function renderIO(io){
  let el=document.getElementById('io-section');
  if(!io){ el.style.display='none'; return; }
  let h='<div class="io-wrap"><div class="io-title">🔗 산업연관 가격파급 (레온티에프)'+
        `<span class="io-tag">${esc(io.mode||'')} 모드</span></div>`;
  let p=io.product;
  if(p&&p.found){
    let ur=(p.ultimate_raw_materials||[]).map(r=>`<b>${esc(r.name)}</b> ${esc(r.effective_pct)}`).join(', ');
    h+=`<div class="io-meta">제품: <b>${esc(p.product_name||'')}</b> · 공정: <b>${esc(p.process||'')}</b>`+
       (ur?`<br>궁극 원자재 분해: ${ur}`:'')+`</div>`;
  }
  let imp=io.impact;
  if(imp&&imp.impacts&&imp.impacts.length){
    let mx=Math.max(...imp.impacts.map(i=>Math.abs(i.price_change_pct)))||1;
    h+=`<div class="io-meta" style="margin-bottom:8px"><b>${esc(imp.raw_material)}</b> `+
       `${imp.shock_pct>=0?'+':''}${imp.shock_pct}% 변동 시 철강 부문별 파급:</div>`;
    h+=imp.impacts.slice(0,5).map(i=>`
      <div class="io-imp-row">
        <span class="io-imp-name">${esc(i.name)} <span class="io-imp-hs">${esc(i.hs_range)}</span></span>
        <span class="io-imp-track"><span class="io-imp-fill" style="width:${Math.abs(i.price_change_pct)/mx*100}%"></span></span>
        <span class="io-imp-pct">${i.price_change_pct>=0?'+':''}${i.price_change_pct.toFixed(2)}%</span>
      </div>`).join('');
  }
  h+='</div>';
  el.innerHTML=h; el.style.display='block';
}

function renderMd(text){
  let sep = text.indexOf('-- 참고 뉴스');
  let main = sep>=0 ? text.slice(0,sep) : text;
  let srcs = sep>=0 ? text.slice(sep) : '';
  main = main.replace(/~~(.*?)~~/g,'$1');
  document.getElementById('md-out').innerHTML = marked.parse(main);
  if(srcs){
    let h='<div class="src-block"><b style="font-size:11px;color:var(--muted2)">참고 뉴스 · KOTRA</b>';
    srcs.split('\n').filter(l=>l.trim()&&!l.startsWith('--')).forEach(l=>{
      let u=l.match(/https?:\/\/\S+/);
      if(u) h+=`<a href="${u[0]}" target="_blank">${u[0]}</a>`;
      else if(l.trim()) h+=`<div style="margin-top:2px">${l.replace(/^\d+\.\s*/,'')}</div>`;
    });
    document.getElementById('src-blk').innerHTML=h+'</div>';
  }
}

/* ── 원자재 단일 진단 (사이드바 클릭) ── */
function loadCReport(name, c, p){
  document.getElementById('cr-section').style.display='block';
  let cr=document.getElementById('cr-content');
  cr.innerHTML='<div style="color:var(--muted);font-size:13px"><span class="spin"></span>진단 중...</div>';
  fetch(`/api/diagnose?name=${encodeURIComponent(name)}&c=${c}&p=${p}`)
    .then(r=>r.json()).then(d=>{
      let s=d.scores, b=d.model_basis||{}, bl=d.basis_lines||{};
      let cC=s.country>=70?'c-green':s.country>=40?'c-yellow':'c-red';
      let pC=s.port>=70?'c-green':s.port>=40?'c-yellow':'c-red';
      let iC=s.item>=70?'c-green':s.item>=40?'c-yellow':'c-red';
      let dC=s.delivery>=70?'c-green':s.delivery>=40?'c-yellow':'c-red';
      let bz=k=>bl[k]?`<div class="sc-basis">${esc(bl[k])}</div>`:'';
      cr.innerHTML=`
        <div style="font-size:14px;font-weight:700;margin-bottom:10px">${esc(name)} 단일 리스크 진단</div>
        <div style="font-size:11px;color:var(--muted);background:var(--surface2);
          border-radius:6px;padding:5px 9px;margin-bottom:10px">
          변동성리스크 ${(b.price_risk||0).toFixed(2)} / 관세 ${(b.tariff||0).toFixed(1)}% / 항만리스크 ${(b.port_risk||0).toFixed(2)}
        </div>
        <div class="score-mini" style="margin-bottom:10px">
          <div class="sc"><div class="sc-label">국가/통관</div><div class="sc-val ${cC}">${s.country}</div>${bz('country')}</div>
          <div class="sc"><div class="sc-label">항만/물류</div><div class="sc-val ${pC}">${s.port}</div>${bz('port')}</div>
          <div class="sc"><div class="sc-label">품목/시장</div><div class="sc-val ${iC}">${s.item}</div>${bz('item')}</div>
          <div class="sc"><div class="sc-label">납기/지연</div><div class="sc-val ${dC}">${s.delivery}</div>${bz('delivery')}</div>
        </div>
        <div style="font-size:12px;font-weight:700;padding:7px 10px;border-radius:7px;margin-bottom:8px;
          background:${d.export_ok?'rgba(45,212,160,.1)':'rgba(248,113,113,.1)'};
          color:${d.export_ok?'var(--green)':'var(--red)'}">
          수출: ${d.export_ok?'✅ 유리':'❌ 불리'} &nbsp;|&nbsp; 수입: ${d.import_ok?'✅ 원활':'⚠️ 주의'}
        </div>
        <div class="ai-rep" id="ai-rep-inner">AI 분석 중...</div>`;
      let aiEl=document.getElementById('ai-rep-inner'), raw='';
      let es=new EventSource(`/api/stream?name=${encodeURIComponent(name)}&c=${c}&p=${p}`);
      es.onmessage=function(e){
        if(e.data==='[DONE]'){es.close();aiEl.innerHTML=raw.replace(/\n/g,'<br>');return;}
        raw+=e.data.replace(/\\n/g,'\n');
        aiEl.innerHTML=raw.replace(/\n/g,'<br>');
      };
      es.onerror=function(){es.close();};
    });
}

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

/* ── 서버 종료 ── */
function doShutdown(){
  if(!confirm('시스템을 완전히 종료하시겠습니까?')) return;
  document.getElementById('overlay').style.display='flex';
  fetch('/api/shutdown',{method:'POST'}).catch(()=>{});
}

/* ── 모델 검증 (백테스트) ── */
let _btLoaded=false;
function toggleBacktest(){
  let el=document.getElementById('backtest-section');
  if(el.style.display==='block'){ el.style.display='none'; return; }
  el.style.display='block';
  el.scrollIntoView({behavior:'smooth',block:'start'});
  if(_btLoaded) return;
  el.innerHTML='<div class="bt-wrap"><span class="spin"></span>검증 실행 중...</div>';
  fetch('/api/backtest').then(r=>r.json()).then(renderBacktest).catch(e=>{
    el.innerHTML='<div class="bt-wrap" style="color:var(--red)">검증 로드 실패: '+esc(String(e))+'</div>';
  });
}
function renderBacktest(d){
  let el=document.getElementById('backtest-section');
  if(!d.available){ el.innerHTML='<div class="bt-wrap" style="color:var(--muted)">검증 모듈 미로드: '+esc(d.msg||'')+'</div>'; return; }
  _btLoaded=true;
  let w=d.weights||{};
  let rows=(d.cases||[]).map(c=>{
    let gc=c.grade==='정상'?'c-green':c.grade==='주의'?'c-yellow':'c-red';
    let v=c.verified?'<span class="bt-vchip bt-v-ok">검증됨</span>':'<span class="bt-vchip bt-v-no">미검증</span>';
    return `<tr>
      <td><div class="bt-case">${esc(c.name)}</div><div class="bt-note">${esc(c.source_note)}</div></td>
      <td>${v}</td>
      <td style="font-weight:700;color:${c.price_change_pct>=0?'#fca5a5':'#5eead4'}">${c.price_change_pct>=0?'+':''}${c.price_change_pct}%</td>
      <td>${c.axis1}</td>
      <td style="font-weight:700">${c.composite}</td>
      <td><span class="bt-grade ${gc}">${esc(c.grade)}</span></td>
      <td style="color:var(--muted2)">${esc(c.expected)}</td>
      <td class="bt-match">${c.match?'✅':'⚠️'}</td>
    </tr>`;
  }).join('');
  let cr=d.consistency_ratio;
  el.innerHTML=`<div class="bt-wrap">
    <div class="bt-head"><h3>📋 모델 검증 — 실제 공급망 충격 사례</h3></div>
    <div class="bt-sub">
      scoring 엔진(AHP/z-score)이 실제 사례에서 산출하는 등급을 점검.
      AHP 가중치 <b>가격 ${w.price} · 항만 ${w.port} · 국가 ${w.country}</b>
      (일관성비율 CR=${cr!=null?cr:'-'}, ${d.consistent?'≤0.1 일관성 충족':'재검토 필요'}).
      <br>※ 종합등급은 3축(가격+항만+국가) 합산이라, 가격단독 심각도(예상)보다 보수적으로 나올 수 있음 — 가격 급등에만 과잉반응하지 않는다는 의미.
    </div>
    <table class="bt-table">
      <thead><tr>
        <th>사례 / 출처</th><th>검증</th><th>변동률</th><th>축1(가격)</th><th>종합</th><th>산출등급</th><th>예상</th><th>일치</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

/* ── 회사 BOM 프로필 ── */
function gradeColor(g){return g==='정상'?'c-green':g==='주의'?'c-yellow':g==='경고'||g==='위험'?'c-red':'';}
function gradeBg(g){return g==='정상'?'rgba(45,212,160,.15)':g==='주의'?'rgba(245,185,66,.15)':'rgba(248,113,113,.15)';}

function toggleProfile(){
  let el=document.getElementById('profile-section');
  if(el.style.display==='block'){el.style.display='none';return;}
  ['history-section','compare-section'].forEach(i=>document.getElementById(i).style.display='none');
  el.style.display='block'; el.scrollIntoView({behavior:'smooth',block:'nearest'});
  let prod=document.getElementById('i-mat').value.trim();
  fetch('/api/profiles').then(r=>r.json()).then(profs=>{
    let saved=prod?profs[prod.toLowerCase().replace(/ /g,'')]:null;
    renderProfileEditor(prod, saved, profs);
  });
}
function renderProfileEditor(prod, saved, profs){
  let rows = (saved&&saved.bom) ? Object.entries(saved.bom) : [['steel',45],['aluminum',25],['copper',20],['plastic',10]];
  let list=Object.values(profs||{});
  let savedList = list.length?('<div style="font-size:12px;color:var(--muted);margin-top:14px">저장된 품목: '+
    list.map(p=>`<a href="#" onclick="loadProf('${esc(p.product)}');return false" style="color:#7eb8ff;margin-right:8px">${esc(p.product)}(${p.bom_total||0}%)</a>`).join('')+'</div>'):'';
  document.getElementById('profile-section').innerHTML=`<div class="panel">
    <h3>🏭 내 회사 실제 BOM</h3>
    <div class="psub">AI 추정 대신 <b>우리 회사 진짜 원자재 구성</b>을 입력·저장하면, 분석이 그 값으로 돌아 정확해집니다. 합계 100% 권장.</div>
    <div class="bom-edit"><input class="be-mat" id="pf-prod" placeholder="품목명 (예: 각관)" value="${esc(prod||(saved?saved.product:''))}"></div>
    <div id="bom-rows">${rows.map(([m,r])=>bomRowHtml(m,r)).join('')}</div>
    <button class="btn-ghost" onclick="addBomRow()" style="margin-top:6px;padding:7px 14px;font-size:13px">+ 원자재 추가</button>
    <div class="be-total" id="bom-total"></div>
    <div class="bom-edit" style="margin-top:6px">
      <input class="be-ratio" id="pf-tariff" placeholder="관세%" value="${saved&&saved.tariff!=null?saved.tariff:''}" style="width:110px">
      <input class="be-mat" id="pf-note" placeholder="메모 (선택)" value="${esc(saved?saved.note||'':'')}">
    </div>
    <div class="be-actions">
      <button class="btn-save" onclick="saveProfile()">💾 저장</button>
      ${saved?`<button class="btn-ghost" onclick="deleteProfile('${esc(saved.product)}')">삭제</button>`:''}
      <button class="btn-ghost" onclick="document.getElementById('profile-section').style.display='none'">닫기</button>
    </div>
    ${savedList}
  </div>`;
  updateBomTotal();
}
function bomRowHtml(m,r){return `<div class="bom-edit">
  <input class="be-mat bom-m" placeholder="원자재(영문, 예: steel)" value="${esc(m)}">
  <input class="be-ratio bom-r" type="number" placeholder="비중%" value="${r}" oninput="updateBomTotal()">
  <button class="be-del" onclick="this.parentNode.remove();updateBomTotal()">×</button></div>`;}
function addBomRow(){document.getElementById('bom-rows').insertAdjacentHTML('beforeend',bomRowHtml('',''));}
function updateBomTotal(){
  let t=0;document.querySelectorAll('#bom-rows .bom-r').forEach(i=>t+=parseFloat(i.value)||0);
  let el=document.getElementById('bom-total');
  el.innerHTML='합계: <span class="'+(Math.abs(t-100)<0.5?'c-green':'c-yellow')+'">'+t.toFixed(0)+'%</span>'+(Math.abs(t-100)<0.5?'':' (100% 권장)');
}
function loadProf(p){document.getElementById('i-mat').value=p;toggleProfile();toggleProfile();}
function saveProfile(){
  let prod=document.getElementById('pf-prod').value.trim();
  if(!prod){alert('품목명을 입력하세요');return;}
  let bom={};document.querySelectorAll('#bom-rows .bom-edit').forEach(row=>{
    let m=row.querySelector('.bom-m').value.trim(), r=parseFloat(row.querySelector('.bom-r').value);
    if(m&&!isNaN(r)) bom[m]=r;
  });
  let tv=document.getElementById('pf-tariff').value;
  fetch('/api/profile',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({product:prod,bom:bom,tariff:tv!==''?parseFloat(tv):null,note:document.getElementById('pf-note').value})})
   .then(r=>r.json()).then(d=>{
     if(d.ok){document.getElementById('i-mat').value=prod;refreshProfileBadge();
       document.getElementById('profile-section').style.display='none';
       alert('저장됨: '+prod+' (분석 시 이 BOM이 자동 적용됩니다)');}
     else alert('저장 실패: '+(d.msg||''));
   });
}
function deleteProfile(p){
  if(!confirm(p+' 프로필을 삭제할까요?'))return;
  fetch('/api/profile/delete',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({product:p})}).then(()=>{refreshProfileBadge();
    document.getElementById('profile-section').style.display='none';});
}
function refreshProfileBadge(){
  let prod=document.getElementById('i-mat').value.trim();
  let b=document.getElementById('profile-badge');
  if(!prod){b.style.display='none';return;}
  fetch('/api/profiles').then(r=>r.json()).then(profs=>{
    let s=profs[prod.toLowerCase().replace(/ /g,'')];
    if(s){b.innerHTML=`🏭 <b>${esc(s.product)}</b> 회사 BOM 적용 중 (${(s.bom_total||0)}%${s.tariff!=null?', 관세 '+s.tariff+'%':''}) — 분석이 이 값으로 실행됩니다`;b.style.display='block';}
    else b.style.display='none';
  }).catch(()=>{});
}

/* ── 분석 이력 / 추세 ── */
function toggleHistory(){
  let el=document.getElementById('history-section');
  if(el.style.display==='block'){el.style.display='none';return;}
  ['profile-section','compare-section'].forEach(i=>document.getElementById(i).style.display='none');
  el.style.display='block'; el.scrollIntoView({behavior:'smooth',block:'nearest'});
  el.innerHTML='<div class="panel"><span class="spin"></span>이력 로딩...</div>';
  fetch('/api/history').then(r=>r.json()).then(renderHistory);
}
function renderHistory(hist){
  let el=document.getElementById('history-section');
  if(!hist.length){el.innerHTML='<div class="panel"><h3>📈 분석 이력</h3><div class="psub">아직 분석 기록이 없습니다. 통합 분석을 실행하면 자동 저장됩니다.</div></div>';return;}
  let last=hist.slice(-30);
  let mx=Math.max(...last.map(h=>h.composite_risk||0),60);
  let bars=last.map(h=>{
    let v=h.composite_risk||0;
    let col=v<40?'#2dd4a0':v<60?'#f5b942':'#f87171';
    return `<div class="trend-bar" style="height:${v/mx*100}%;background:${col}" title="${esc(h.ts)} ${esc(h.product)}→${esc(h.country)}: ${v} ${esc(h.grade||'')}"></div>`;
  }).join('');
  let items=hist.slice(-12).reverse().map(h=>`<div class="hist-item">
    <span><b>${esc(h.product)}</b> → ${esc(h.country)} <span style="color:var(--muted);font-size:11px">${esc(h.ts)}</span>
      ${h.bom_source==='회사 실제'?'<span style="color:#5eead4;font-size:10px">🏭회사BOM</span>':''}</span>
    <span><span style="color:var(--muted);font-size:11px">실질 ${h.total_impact!=null?(h.total_impact<=0?'':'+')+h.total_impact+'%':'-'}</span>
      &nbsp;<span class="hg ${gradeColor(h.grade)}" style="background:${gradeBg(h.grade)}">${esc(h.grade||'-')} ${h.composite_risk!=null?h.composite_risk:''}</span></span>
  </div>`).join('');
  el.innerHTML=`<div class="panel">
    <h3>📈 분석 이력 & 추세</h3>
    <div class="psub">최근 분석들의 <b>종합리스크 추세</b>. 막대가 낮을수록(초록) 안전. 같은 품목·국가를 반복 분석하면 변화를 추적할 수 있습니다.</div>
    <div class="trend-row">${bars}</div>
    <div style="margin-top:14px">${items}</div>
    <button class="btn-ghost" style="margin-top:14px;font-size:12px;padding:7px 14px"
      onclick="if(confirm('이력 전체 삭제?')){fetch('/api/history/clear',{method:'POST'}).then(()=>toggleHistory());toggleHistory();}">이력 비우기</button>
  </div>`;
}

/* ── 다국가 비교 ── */
function toggleCompare(){
  let el=document.getElementById('compare-section');
  if(el.style.display==='block'){el.style.display='none';return;}
  ['profile-section','history-section'].forEach(i=>document.getElementById(i).style.display='none');
  el.style.display='block'; el.scrollIntoView({behavior:'smooth',block:'nearest'});
  let prod=document.getElementById('i-mat').value.trim();
  el.innerHTML=`<div class="panel">
    <h3>🌍 다국가 비교</h3>
    <div class="psub">한 품목을 <b>여러 수출국</b>에 대해 동시 비교 → 어디로 수출하는 게 가장 안전한지 한눈에. (AI 호출 절약 위해 BOM은 1회만 계산)</div>
    <div class="cmp-inp">
      <input id="cmp-prod" placeholder="품목" value="${esc(prod)}">
      <input id="cmp-ctrs" placeholder="수출국들 (쉼표로, 최대 5개): 베트남,미국,독일">
      <button class="btn-save" onclick="runCompare()">비교</button>
    </div>
    <div id="cmp-result"></div>
  </div>`;
}
function runCompare(){
  let prod=document.getElementById('cmp-prod').value.trim();
  let ctrs=document.getElementById('cmp-ctrs').value.trim();
  if(!prod||!ctrs){alert('품목과 수출국들을 입력하세요');return;}
  let res=document.getElementById('cmp-result');
  res.innerHTML='<div style="color:var(--muted);padding:10px"><span class="spin"></span>비교 중... (국가 수만큼 걸립니다)</div>';
  fetch('/api/compare?'+new URLSearchParams({product:prod,countries:ctrs}))
    .then(r=>r.json()).then(d=>{
      if(!d.rows.length){res.innerHTML='<div style="color:var(--muted)">결과 없음</div>';return;}
      let rows=d.rows.map((x,i)=>{
        if(x.error)return `<tr><td class="cmp-ctr">${esc(x.country)}</td><td colspan="6" style="color:var(--red)">${esc(x.error)}</td></tr>`;
        return `<tr class="${i===0?'cmp-best':''}">
          <td class="cmp-ctr">${i===0?'🏆 ':''}${esc(x.country)}</td>
          <td><span class="hg ${gradeColor(x.grade)}" style="background:${gradeBg(x.grade)}">${esc(x.grade)}</span></td>
          <td style="font-weight:700">${x.composite_risk}</td>
          <td class="${x.country_s>=70?'c-green':x.country_s>=40?'c-yellow':'c-red'}">${x.country_s}</td>
          <td class="${x.port_s>=70?'c-green':x.port_s>=40?'c-yellow':'c-red'}">${x.port_s}</td>
          <td class="${x.delivery_s>=70?'c-green':x.delivery_s>=40?'c-yellow':'c-red'}">${x.delivery_s}</td>
          <td class="${x.fx_change>=0?'pos':'neg'}">${x.fx_change>=0?'+':''}${x.fx_change}%</td>
        </tr>`;}).join('');
      res.innerHTML=`<table class="cmp-table">
        <thead><tr><th style="text-align:left">수출국</th><th>등급</th><th>종합리스크</th><th>국가</th><th>항만</th><th>납기</th><th>환율</th></tr></thead>
        <tbody>${rows}</tbody></table>
        <div style="font-size:12px;color:var(--muted);margin-top:10px">🏆 = 종합리스크 가장 낮은(안전한) 수출국. 종합리스크는 낮을수록 안전.</div>`;
    }).catch(e=>{res.innerHTML='<div style="color:var(--red)">비교 실패: '+esc(String(e))+'</div>';});
}

/* ── 경상도 주력품 프리셋 ── */
const GS_PRESETS=["열연강판","냉연강판","후판","철근","강관","형강","자동차부품","조선기자재","산업용밸브","플랜지","베어링","주단조품"];
function pickPreset(name){
  document.getElementById('i-mat').value=name;
  refreshProfileBadge();
  document.getElementById('i-ctr').focus();
}
function renderPresets(){
  document.getElementById('preset-chips').innerHTML =
    GS_PRESETS.map(p=>`<span class="chip" onclick="pickPreset('${p}')">${p}</span>`).join('');
}

/* ── 라이트/다크 테마 ── */
function applyTheme(light){
  document.body.classList.toggle('light', light);
  let b=document.getElementById('theme-btn');
  if(b) b.textContent = light ? '☀️ 라이트' : '🌙 다크';
  try{ localStorage.setItem('tr_theme', light?'light':'dark'); }catch(e){}
}
function toggleTheme(){ applyTheme(!document.body.classList.contains('light')); }
(function(){ try{ applyTheme(localStorage.getItem('tr_theme')==='light'); }catch(e){} })();

/* ── 간단/상세 모드 ── */
function applySimple(simple){
  document.body.classList.toggle('simple', simple);
  let b=document.getElementById('simple-btn');
  if(b) b.textContent = simple ? '📋 상세' : '🪪 간단';
  try{ localStorage.setItem('tr_simple', simple?'1':'0'); }catch(e){}
}
function toggleSimple(){ applySimple(!document.body.classList.contains('simple')); }
(function(){ try{ applySimple(localStorage.getItem('tr_simple')==='1'); }catch(e){} })();

/* 초기화 */
loadFx('미국'); loadC(); refreshProfileBadge(); renderPresets();
document.getElementById('i-mat').addEventListener('input', refreshProfileBadge);
setInterval(()=>loadFx(_curCtr), 60000);
setInterval(loadC, 120000);
</script>
</body>
</html>"""


# ============================================================
#  Flask 라우트
# ============================================================
@app.route("/")
def index():
    return render_template_string(PAGE, mode=_MODE_LABEL.get(analyzer.mode, analyzer.mode))

@app.route("/analyze")
def analyze():
    material = request.args.get("material","").strip()
    country  = request.args.get("country","").strip() or None
    if not material:
        return Response(iter(["품목/원자재를 입력해 주세요."]),
                        mimetype="text/plain; charset=utf-8")
    def gen():
        try:
            for piece in analyzer.stream_report(material, country=country): yield piece
        except Exception as e:
            yield f"\n[분석 오류] {e}\n"
    return Response(gen(), mimetype="text/plain; charset=utf-8")

@app.route("/api/fx")
def api_fx():
    country = request.args.get("country", "미국")
    return jsonify(engine.get_fx_rate(country))

@app.route("/api/search")
def api_search():
    q = request.args.get("q","").lower().strip()
    seen, results = set(), []
    for v in engine.data_dict.values():
        k = v["name"]
        if k in seen: continue
        if not q or q in v["name"].lower():
            seen.add(k); results.append(v)
    return jsonify(results[:40])

@app.route("/api/diagnose")
def api_diagnose():
    name  = request.args.get("name","")
    c     = float(request.args.get("c", 0))
    price = float(request.args.get("p", 0))
    s     = engine.calculate_scores(name, c, price)
    b     = s.pop("basis", {})
    blines = s.pop("basis_lines", {})
    export_ok = s.pop("export_ok", False)
    import_ok = s.pop("import_ok", False)
    return jsonify({"scores":s,"model_basis":b,"basis_lines":blines,
                    "export_ok":export_ok,"import_ok":import_ok})

@app.route("/api/simulate", methods=["POST"])
def api_simulate():
    d       = request.json or {}
    product = d.get("product","")
    country = d.get("country","미국")
    if not _rate_ok(request.remote_addr):
        return jsonify({"error": "요청이 너무 잦습니다. 잠시 후 다시 시도하세요."}), 429
    use_profile = d.get("use_profile", True)   # 저장된 회사 BOM 자동 적용
    bom_override = tariff_override = None
    prof = None
    if use_profile:
        prof = _load_store(PROFILES_FILE, {}).get(_norm_key(product))
        if prof:
            bom_override    = prof.get("bom") or None
            tariff_override = prof.get("tariff")
    # 캐시: 회사 프로필 없는 일반 조회는 30분 재사용 (Gemini HS/BOM 호출 절약)
    ckey = f"sim|{_norm_key(product)}|{country}"
    if not prof:
        cached = _cache_get(ckey, 1800)
        if cached is not None:
            return jsonify(cached)
    result = engine.run_simulator(product, country,
                                  bom_override=bom_override, tariff_override=tariff_override)
    if not prof:
        _cache_set(ckey, result)
    # 분석 이력 자동 저장(추세용)
    try:
        ps = result.get("product_scores", {})
        hist = _load_store(HISTORY_FILE, [])
        hist.append({
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "product": product, "country": country,
            "composite_risk": ps.get("composite_risk"),
            "grade": ps.get("grade"),
            "raw_impact": result.get("raw_impact"),
            "fx_impact": result.get("fx_impact"),
            "total_impact": result.get("total_impact"),
            "bom_source": result.get("bom_source"),
        })
        _save_store(HISTORY_FILE, hist[-300:])   # 최근 300건 유지
    except Exception as e:
        print(f"[history] 저장 실패: {e}")
    return jsonify(result)

# ── 회사 BOM 프로필 ──────────────────────────────────────────
@app.route("/api/profiles")
def api_profiles():
    return jsonify(_load_store(PROFILES_FILE, {}))

@app.route("/api/profile", methods=["POST"])
def api_profile_save():
    d = request.json or {}
    product = (d.get("product") or "").strip()
    if not product:
        return jsonify({"ok": False, "msg": "품목명 필요"})
    bom = d.get("bom") or {}
    # 비중 숫자화 + 합계 검증
    try:
        bom = {str(k).strip(): float(v) for k, v in bom.items() if str(k).strip()}
    except Exception:
        return jsonify({"ok": False, "msg": "BOM 비중은 숫자여야 함"})
    profs = _load_store(PROFILES_FILE, {})
    profs[_norm_key(product)] = {
        "product": product, "bom": bom,
        "tariff": d.get("tariff"),
        "note": (d.get("note") or "").strip(),
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "bom_total": round(sum(bom.values()), 1),
    }
    ok = _save_store(PROFILES_FILE, profs)
    return jsonify({"ok": ok, "saved": profs[_norm_key(product)]})

@app.route("/api/profile/delete", methods=["POST"])
def api_profile_delete():
    d = request.json or {}
    profs = _load_store(PROFILES_FILE, {})
    key = _norm_key(d.get("product",""))
    if key in profs:
        del profs[key]; _save_store(PROFILES_FILE, profs)
    return jsonify({"ok": True})

# ── 분석 이력 / 추세 ─────────────────────────────────────────
@app.route("/api/history")
def api_history():
    hist = _load_store(HISTORY_FILE, [])
    prod = request.args.get("product","").strip()
    if prod:
        hist = [h for h in hist if h.get("product","") == prod]
    return jsonify(hist[-60:])   # 최근 60건

@app.route("/api/history/clear", methods=["POST"])
def api_history_clear():
    _save_store(HISTORY_FILE, [])
    return jsonify({"ok": True})

# ── 다국가 경량 비교 ─────────────────────────────────────────
@app.route("/api/compare")
def api_compare():
    """한 품목을 여러 수출국에 대해 비교(경량: BOM/HS는 1회 계산, 국가·환율만 가변).
    AI 호출을 아끼려고 BOM/관세는 첫 결과를 재사용한다."""
    product   = request.args.get("product","").strip()
    countries = [c.strip() for c in request.args.get("countries","").split(",") if c.strip()]
    if not product or not countries:
        return jsonify({"product": product, "rows": []})
    countries = countries[:5]
    prof = _load_store(PROFILES_FILE, {}).get(_norm_key(product)) if product else None
    bom_override    = (prof or {}).get("bom") or None
    tariff_override = (prof or {}).get("tariff")
    rows = []
    for ctr in countries:
        try:
            r  = engine.run_simulator(product, ctr,
                                      bom_override=bom_override, tariff_override=tariff_override)
            ps = r.get("product_scores", {})
            rows.append({
                "country": ctr, "grade": ps.get("grade"),
                "composite_risk": ps.get("composite_risk"),
                "country_s": ps.get("country"), "port_s": ps.get("port"),
                "item_s": ps.get("item"), "delivery_s": ps.get("delivery"),
                "fx_pair": r.get("fx_pair"), "fx_change": r.get("fx_change"),
                "total_impact": r.get("total_impact"),
            })
            # 첫 결과의 BOM/관세를 재사용해 이후 국가는 AI 호출 절약
            if bom_override is None:
                bom_override = {b["material"]: b["ratio"] for b in r.get("bom_details", [])}
            if tariff_override is None:
                tariff_override = r.get("tariff_pct")
        except Exception as e:
            rows.append({"country": ctr, "error": str(e)[:60]})
    # 종합리스크 낮은(안전한) 순 정렬
    rows.sort(key=lambda x: (x.get("composite_risk") is None, x.get("composite_risk", 999)))
    return jsonify({"product": product, "rows": rows})

@app.route("/api/backtest")
def api_backtest():
    """실제 공급망 충격 사례로 AHP/z-score 엔진 등급 산정 검증."""
    if _backtest is None or _scoring is None:
        return jsonify({"available": False, "msg": "backtest/scoring 모듈 미로드"})
    sm  = _scoring.score_material
    PS  = _scoring.PriceSignal
    cases = []
    for c in _backtest.CASES:
        r = sm(c.material, PS(material=c.material, price_change_pct=c.price_change_pct),
               c.ports, c.countries)
        cases.append({
            "name": c.name, "material": c.material,
            "price_change_pct": c.price_change_pct,
            "axis1": r.axis1_price, "axis2": r.axis2_port, "axis3": r.axis3_country,
            "composite": r.composite, "grade": r.grade,
            "expected": c.expected_severity, "match": r.grade == c.expected_severity,
            "verified": c.verified, "source_note": c.source_note,
        })
    ahp = getattr(_scoring, "_AHP_RESULT", {})
    return jsonify({"available": True,
                    "weights": _scoring.DEFAULT_WEIGHTS,
                    "consistency_ratio": ahp.get("consistency_ratio"),
                    "consistent": ahp.get("consistent"),
                    "cases": cases})

@app.route("/api/test_gemini")
def api_test_gemini():
    """Gemini 연결 직접 테스트 - 브라우저에서 /api/test_gemini 접속"""
    if not analyzer.client:
        return jsonify({"status":"error","msg":"Gemini 클라이언트 없음"})
    try:
        resp = analyzer.client.models.generate_content(
            model="gemini-2.0-flash",
            contents='{"test":"ok"} 라고만 JSON으로 출력하세요.')
        text = ""
        try:
            text = resp.text or ""
        except Exception as e:
            return jsonify({"status":"resp.text 오류","error":str(e)})
        return jsonify({"status":"ok","response_len":len(text),"response":text[:500]})
    except Exception as e:
        import traceback
        return jsonify({"status":"exception","error":str(e),
                        "type":type(e).__name__,"traceback":traceback.format_exc()[-800:]})

@app.route("/api/stream_comprehensive")
def api_stream_comprehensive():
    product       = request.args.get("product","")
    country       = request.args.get("country","")
    mat_impact    = float(request.args.get("mat_impact", 0))
    fx_impact     = float(request.args.get("fx_impact",  0))
    fx_rate       = float(request.args.get("fx_rate", FX_BASELINE))
    fx_pair       = request.args.get("fx_pair","USD/KRW")
    currency_name = request.args.get("currency_name","달러")
    return Response(
        stream_with_context(engine.stream_comprehensive(
            product, country, mat_impact, fx_impact, fx_rate, fx_pair, currency_name)),
        mimetype="text/event-stream",
        headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})

@app.route("/api/shutdown", methods=["POST"])
def api_shutdown():
    # 공개 서버 보호: 로컬(127.0.0.1)에서만 종료 허용
    if IS_DEPLOY or request.remote_addr not in ("127.0.0.1", "::1", "localhost"):
        return jsonify({"status": "forbidden", "msg": "원격 종료는 차단되어 있습니다."}), 403
    def kill():
        time.sleep(0.5); os._exit(0)
    threading.Thread(target=kill, daemon=True).start()
    return jsonify({"status": "shutdown"})

@app.route("/api/analyze_unified")
def api_analyze_unified():
    material = request.args.get("material","").strip()
    country  = request.args.get("country","").strip() or "미국"
    try:
        sim_data = json.loads(request.args.get("sim","{}"))
    except Exception:
        sim_data = {}
    if not material:
        def _err():
            yield "품목을 입력하세요."
        return Response(_err(), mimetype="text/plain; charset=utf-8")
    if not _rate_ok(request.remote_addr):
        def _rl():
            yield "> ⏳ 요청이 너무 잦습니다. 잠시 후 다시 시도하세요. (공개 서버 보호)\n"
        return Response(_rl(), mimetype="text/plain; charset=utf-8")
    def gen():
        try:
            for piece in analyzer.stream_unified_report(material, country, sim_data):
                yield piece
        except Exception as e:
            yield f"\n[오류] {e}\n"
    return Response(stream_with_context(gen()),
                    mimetype="text/plain; charset=utf-8",
                    headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})

@app.route("/api/stream")
def api_stream():
    name  = request.args.get("name","")
    c     = float(request.args.get("c", 0))
    price = float(request.args.get("p", 0))
    s     = engine.calculate_scores(name, c, price)
    return Response(
        stream_with_context(engine.stream_analysis(name, c, s, price)),
        mimetype="text/event-stream",
        headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})


# ============================================================
#  진입점
# ============================================================
if __name__ == "__main__":
    print("=" * 55)
    print("  TradeRisk AI - 경상도 수출 인텔리전스 (통합)")
    print("=" * 55)
    print(f"  분석 모드: {analyzer.mode}")
    print(f"  Gemini:   {'연결됨' if analyzer.client else '오프라인 (apikey.txt 확인)'}")
    print(f"  KOTRA:    {'연결됨' if kotra.enabled else '비활성 (kotra_key.txt 확인)'}")
    print(f"  통화 지원: {len(COUNTRY_CURRENCY)}개국")
    _port = int(os.environ.get("PORT", 5000))
    _host = "0.0.0.0" if IS_DEPLOY else "127.0.0.1"
    print(f"  서버 주소: http://{_host}:{_port}")
    print("=" * 55)
    if not IS_DEPLOY:   # 로컬에서만 브라우저 자동 오픈
        threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{_port}")).start()
    app.run(host=_host, port=_port, debug=False, threaded=True)
