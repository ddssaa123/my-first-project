import os
import sys
import webbrowser
import threading
import time
import requests
import pandas as pd
from datetime import datetime
import yfinance as yf
from io import StringIO
import xml.etree.ElementTree as ET
from google import genai
from google.genai import types
from flask import Flask, jsonify, request

app = Flask(__name__)

class SupplyChainEngine:
    def __init__(self):
        self.gemini_key = os.environ.get("GEMINI_API_KEY", "AIzaSyB6lbUdiEcqVyBHS2cNvb3Gw9pSQRxkl58")
        self.url = "https://tradingeconomics.com/commodities"
        self.headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'}
        self.data_dict = {}
        self.fx_p, self.fx_c = 0.0, 0.0
        
        if self.gemini_key and not self.gemini_key.startswith("🔍"):
            self.ai_client = genai.Client(api_key=self.gemini_key)
        else:
            self.ai_client = None

    def sync_market_data(self):
        try:
            res = requests.get(self.url, headers=self.headers, timeout=10)
            tables = pd.read_html(StringIO(res.text))
            self.data_dict = {}
            for df in tables:
                for _, row in df.iterrows():
                    try:
                        name = str(row.iloc[0]).strip()
                        if name == "nan" or "Price" in name: continue
                        self.data_dict[name.lower()] = {
                            'name': name,
                            'p': float(str(row.iloc[1]).replace(',', '')),
                            'c': float(str(row.iloc[3]).replace('%', ''))
                        }
                    except: continue
            
            fx = yf.download("USDKRW=X", period="2d", progress=False)
            if not fx.empty and len(fx) >= 2:
                close_series = fx['Close'].dropna()
                self.fx_p = float(close_series.iloc[-1].iloc[0]) if isinstance(close_series.iloc[-1], pd.Series) else float(close_series.iloc[-1])
                prev_p = float(close_series.iloc[-2].iloc[0]) if isinstance(close_series.iloc[-2], pd.Series) else float(close_series.iloc[-2])
                self.fx_c = ((self.fx_p - prev_p) / prev_p) * 100
            print("✅ [엔진] 모든 마켓 데이터 동기화 완료!")
        except Exception as e:
            print(f"❌ 데이터 수집 실패: {e}")

    def run_ai_analysis(self, target_name, p, c):
        if not self.ai_client:
            return "<h3>AI 분석 불가</h3>API Key 환경변수가 등록되지 않았습니다."

        scraped_context = f"[로컬 상태] 자재: {target_name}, 현재가: {p:,.2f} ({c:+.2f}%), 환율: {self.fx_p:,.2f}원 ({self.fx_c:+.2f}%)"
        prompt = """
        너는 글로벌 공급망 리스크 관리 최고 전문가이다. 
        제공된 툴 'Google Search'를 적극 활용하여 오늘 날짜(현재 2026년) 기준의 공급망 현황, 전 세계 지정학적 분쟁 속보, 미/중 규제 및 국내 정책 동향을 실시간 검색해라.
        결과는 다른 서론 없이 '### 📊 글로벌 리스크 5줄 요약'과 '### 💡 구매/조달 실무자를 위한 수출입 추천 의사결정' 두 파트로만 나누어 깔끔하게 격식 있는 한국어로 출력해라.
        """
        try:
            response = self.ai_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[scraped_context, f"[{target_name}] 자재에 대해 분석하라. {prompt}"],
                config=types.GenerateContentConfig(tools=[types.Tool(google_search=types.GoogleSearch())])
            )
            import re
            txt = response.text.replace("###", "<h3 style='color:#2c3e50; margin-top:25px; border-left:4px solid #3498db; padding-left:10px;'>").replace("---", "<hr style='border-color:#e0e0e0;'>")
            txt = re.sub(r'\*\*(.*?)\*\*', r'<strong style="color:#2c3e50;">\1</strong>', txt)
            txt = txt.replace("\n", "<br>")
            
            # AI가 준 백슬래시 이스케이프 자잘한 기호 원복 처리
            txt = txt.replace("&lt;", "<").replace("&gt;", ">")
            return txt
        except Exception as e:
            return f"❌ LLM 분석 중 예외 발생: {e}"

engine = SupplyChainEngine()

@app.route('/')
def index():
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>공급망 인텔리전스 대시보드 [MARK VI]</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;700&display=swap');
            body {{ font-family: 'Noto Sans KR', sans-serif; background-color: #f8f9fa; color: #333333; margin: 0; padding: 40px; }}
            .container {{ max-width: 950px; margin: 0 auto; background: #ffffff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 16px rgba(0,0,0,0.08); border: 1px solid #e9ecef; }}
            h1 {{ color: #212529; font-size: 26px; margin-top: 0; border-bottom: 2px solid #dee2e6; padding-bottom: 15px; font-weight: 700; }}
            .search-box {{ display: flex; gap: 10px; margin: 25px 0; }}
            input {{ flex: 1; padding: 12px; font-size: 16px; border: 2px solid #ced4da; border-radius: 6px; outline: none; }}
            input:focus {{ border-color: #3498db; }}
            button {{ padding: 12px 24px; font-size: 16px; font-weight: bold; background: #3498db; color: #fff; border: none; border-radius: 6px; cursor: pointer; }}
            button:hover {{ background: #2980b9; }}
            .main-layout {{ display: grid; grid-template-columns: 350px 1fr; gap: 30px; margin-top: 20px; }}
            .list-panel {{ background: #f1f3f5; padding: 20px; border-radius: 8px; border: 1px solid #dee2e6; height: 450px; overflow-y: auto; }}
            .list-item {{ padding: 12px; background: #fff; border: 1px solid #e9ecef; margin-bottom: 8px; border-radius: 6px; cursor: pointer; font-size: 14px; transition: 0.2s; }}
            .list-item:hover {{ background: #e3fafc; border-color: #99e9f2; }}
            .report-panel {{ background: #ffffff; padding: 25px; border-radius: 8px; border: 1px solid #dee2e6; height: 440px; overflow-y: auto; }}
            .status-box {{ padding: 12px 15px; border-radius: 6px; font-weight: bold; font-size: 14px; margin-bottom: 20px; }}
            .status-danger {{ background-color: #fff5f5; border: 1px solid #ffa8a8; color: #e64980; }}
            .status-success {{ background-color: #f4fce3; border: 1px solid #c0eb75; color: #2b8a3e; }}
            .status-normal {{ background-color: #f1f3f5; border: 1px solid #ced4da; color: #495057; }}
            
            /* 💡 화이트 테마 전용 가시성 쩌는 배지 스타일 시트 완벽 탑재 */
            span.highlight-run {{ background: #d3f9d8 !important; color: #2b8a3e !important; padding: 3px 10px; border-radius: 4px; font-weight: bold; border: 1px solid #b2f2bb; display: inline-block; }}
            span.highlight-hold {{ background: #ffe3e3 !important; color: #c92a2a !important; padding: 3px 10px; border-radius: 4px; font-weight: bold; border: 1px solid #ffc9c9; display: inline-block; }}
            
            .loading {{ color: #3498db; font-weight: bold; animation: blink 1.5s infinite; }}
            @keyframes blink {{ 0% {{ opacity: 0.3; }} 50% {{ opacity: 1; }} 100% {{ opacity: 0.3; }} }}
                </style>
                <script>
                    function searchCommodity() {{
                        let q = document.getElementById('search-input').value;
                        fetch('/api/search?q=' + encodeURIComponent(q))
                            .then(res => res.json())
                            .then(data => {{
                                let html = '';
                                data.forEach(item => {{
                                    html += `<div class="list-item" onclick="loadReport('${{item.name}}', ${{item.p}}, ${{item.c}})">
                                                <strong>${{item.name.toUpperCase()}}</strong><br>
                                                시세: ${{item.p.toLocaleString()}} | 변동: ${{item.c >= 0 ? '+' : ''}}${{item.c}}%
                                             </div>`;
                                }});
                                document.getElementById('list-container').innerHTML = html || '<div style="color:#868e96; text-align:center; padding-top:20px;">검색 결과 없음</div>';
                            }});
                    }}
                    function loadReport(name, p, c) {{
                        document.getElementById('report-container').innerHTML = '<div class="loading">🧠 Gemini가 실시간 구글 뉴스를 분석 중입니다... (약 5초 소요)</div>';
                        fetch(`/api/report?name=${{encodeURIComponent(name)}}&p=${{p}}&c=${{c}}`)
                            .then(res => res.json())
                            .then(data => {{
                                let total = c + {engine.fx_c};
                                let statusHtml = '';
                                if(total > 1.0) {{
                                    statusHtml = '<div class="status-box status-danger">🚨 경제 리스크: 자재 및 환율 동반 상승 주의 (원화 집행 부담 폭등)</div>';
                                }} else if(total < -1.0) {{
                                    statusHtml = '<div class="status-box status-success">✅ 경제 리스크: 자재 조달 및 수입 단가 매우 유리 (적기 매수 구간)</div>';
                                }} else {{
                                    statusHtml = '<div class="status-box status-normal">🔎 경제 리스크: 시장 보합 - 유의미한 거시 변동 없음</div>';
                                }}
                                
                                let infoHtml = `<div style="display:flex; gap:15px; margin-bottom:15px; font-size:14px; background:#f8f9fa; padding:12px; border-radius:6px; border:1px solid #e9ecef;">
                                                    <div><strong>원달러 환율:</strong> {engine.fx_p:,.2f}원 ({engine.fx_c:+.2f}%)</div>
                                                    <div><strong>선택 자재:</strong> ${{name.toUpperCase()}} (${{c >= 0 ? '+' : ''}}${{c}}%)</div>
                                                </div>`;
                                
                                // 🚨 핵심 수정: textContent 대신 innerHTML을 써서 글자 안에 섞인 배지 코드가 진짜 태그로 작동하게 만듦!
                                document.getElementById('report-container').innerHTML = statusHtml + infoHtml + data.report;
                            }});
                    }}
                </script>
            </head>
            <body>
                <div class="container">
                    <h1>⚙️ SUPPLY CHAIN RISK INTELLIGENCE [MARK VI]</h1>
                    <div class="search-box">
                        <input type="text" id="search-input" placeholder="원자재 키워드 입력 (예: tin, copper, iron ore)..." onkeypress="if(event.keyCode==13) searchCommodity()">
                        <button onclick="searchCommodity()">조회 및 검색</button>
                    </div>
                    <div class="main-layout">
                        <div class="list-panel" id="list-container">
                            <div style="color:#868e96; text-align:center; padding-top:180px;">원자재를 검색하면 목록이 나타납니다.</div>
                        </div>
                        <div class="report-panel" id="report-container">
                            <div style="color:#868e96; text-align:center; padding-top:180px;">좌측 리스트에서 자재를 선택하면 실시간 AI 리포트가 완성됩니다.</div>
                        </div>
                    </div>
                </div>
            </body>
            </html>
            """
    return html

@app.route('/api/search')
def api_search():
    q = request.args.get('q', '').lower().strip()
    results = []
    for k, v in engine.data_dict.items():
        words = k.replace('/', ' ').replace('(', ' ').replace(')', ' ').split()
        if q in words: results.append(v)
    return jsonify(results)

@app.route('/api/report')
def api_report():
    name = request.args.get('name', '')
    p = float(request.args.get('p', 0))
    c = float(request.args.get('c', 0))
    ai_report = engine.run_ai_analysis(name, p, c)
    return jsonify({'report': ai_report})

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8989")

if __name__ == "__main__":
    print("🔄 [1단계] 마켓 데이터 초기화 중...")
    engine.sync_market_data()
    
    print("🖥️ [2단계] 자동 브라우저 활성화 쓰레드 가동...")
    threading.Thread(target=open_browser, daemon=True).start()
    
    print("🚀 [3단계] 안정적인 Flask 웹 서비스 시작 (Port: 8989)...")
    app.run(host='127.0.0.1', port=8989, debug=False)