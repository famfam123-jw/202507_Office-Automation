import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import requests
import email.utils
from datetime import datetime, timedelta
import pytz
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import os
from urllib.parse import urlparse
import webbrowser

kst = pytz.timezone("Asia/Seoul")
#나의 네이버 API 키 아이디, 비밀번호
NAVER_CLIENT_ID = "MY_NAVER_CLIENT_ID"
NAVER_CLIENT_SECRET = "MY_NAVER_CLIENT_SECRET"

MEDIA_DOMAIN_MAP = {
    "newsis.com": "뉴시스",
    "redian.org": "래디앙",
    "yna.co.kr": "연합뉴스",
    "edaily.co.kr": "이데일리",
    "asiatime.co.kr": "아시아타임즈",
    "sisajournal.com": "시사저널",
    "pinpointnews.co.kr": "핀포인트뉴스",
    "news.naver.com": "네이버뉴스",
    "kyeonggi.com": "경기일보",
    "queen.co.kr": "이코노미퀸",
    "tjb.co.kr": "TJB뉴스",
    "naver.com": "네이버",
    "ohmynews.com": "오마이뉴스",
    "pressian.com": "프레시안",
    "mk.co.kr": "매일경제",
    "nocutnews.co.kr":"노컷뉴스",
    "weeklytoday.com": "위클리오늘",
    "news.bbsi.co.kr": "불교방송",
    "hankyung.com": "한국경제",
    "yonhapnewstv.co.kr": "연합뉴스TV",
    "hani.co.kr": "한겨레",
    "mbnmoney.mbn.co.kr": "매일경제TV",
    "seouleconews.com": "서울이코노미뉴스",
    "h21.hani.co.kr": "한겨레21",
    "chosun.com": "조선일보",
    "donga.com": "동아일보",
    "joongang.co.kr": "중앙일보",
    "news.kbs.co.kr": "KBS",
    "news.sbs.co.kr": "SBS",
    "kbiznews.co.kr": "중소기업뉴스",
    "mbn.co.kr": "MBN",
    "etoday.co.kr": "이투데이",
    "etnews.com": "전자신문",
    "news1.kr": "뉴스1",
    "khan.co.kr": "경향신문",
    "segye.com": "세계일보",
    "wowtv.co.kr": "한국경제TV",
    "seoul.co.kr": "서울신문",
    "biz.chosun.com": "조선비즈",
    "pulse.kr": "펄스뉴스",
    "mediatoday.co.kr": "미디어투데이",
    "hankookilbo.com": "한국일보",
    "kookje.co.kr": "국제신문",
    "kmib.co.kr": "국민일보",
    "koscaj.com": "대한전문건설신문",
    "safetynews.co.kr": "안전신문",
    "efnews.co.kr": "파이낸셜뉴스",
    "ksmnews.co.kr": "경상매일신문",
    "hapt.co.kr": "한국아파트신문",
    "thepowernews.co.kr": "더파워뉴스",
    "ekn.kr": "에너지경제",
    "view.asiae.co.kr": "아시아경제",
    "newstomato.com": "뉴스토마토",
    "kpanews.co.kr": "약사공론",
    "kukinews.com": "쿠키뉴스",
    "news.mt.co.kr": "머니투데이",
    "ytn.co.kr": "YTN",
    "worklaw.co.kr": "월간노동법률",
    "newscj.com": "천지일보",
    "asiatoday.co.kr": "아시아투데이",
    "ftoday.co.kr": "파이낸셜투데이",
    "biz.newdaily.co.kr": "뉴데일리",
    "todaykorea.co.kr": "투데이코리아",
    "ajunews.com": "아주경제",
    "newspim.com": "뉴스핌",
    "eroun.net": "이로운넷",
    "lec.co.kr": "법률저널",
    "00news.co.kr": "공공뉴스",
    "sisajournal-e.com": "시사저널e",
    "inews24.com": "아이뉴스24",
    "polinews.co.kr": "폴리뉴스",
    "viva100.com": "브릿지경제",
    "sedaily.com": "서울경제",
    "imaeil.com": "매일신문",
    "meconomynews.com": "시장경제",
    "jeonmae.co.kr": "전국매일신문",
    "biz.sbs.co.kr": "SBSbiz",
    "g-enews.com": "글로벌이코노믹",
    "dailian.co.kr": "데일리안",
    "niceeconomy.co.kr": "나이스경제",
    "businesspost.co.kr": "비즈니스포스트",
    "job-post.co.kr": "잡포스트",
    "fnnews.com": "파이낸셜",
    "obsnews.co.kr": "OBS뉴스",
    "ceoscoredaily.com": "CEO스코어데일리",
    "labortoday.co.kr": "매일노동뉴스",
    "electimes.com": "전기신문",
    "naeil.com": "내일신문",
    "dt.co.kr": "디지털타임스",
    "fntimes.com": "한국금융신문",
    "starin.edaily.co.kr": "이데일리",
    "biz.heraldcorp.com": "헤럴드경제",
    "newsworks.co.kr": "뉴스웍스",
    "news.tf.co.kr": "더팩트",
    "joongangenews.com": "중앙이코노미뉴스",
    "daily.hankooki.com": "데일리한국",
    "imnews.imbc.com": "MBC뉴스",
    "fetv.co.kr": "FETV",
    "m-i.kr": "매일일보",
    "hansbiz.co.kr": "한스경제",
    "munwha.com":"문화일보",
    "radio.ytn.co.kr": "YTN라디오",
    "energy-news.co.kr": "에너지뉴스",
    "impacton.net": "임팩트온",
    "enewstoday.co.kr": "뉴스투데이",
    "sisacast.kr": "시사캐스트",
    "kmecnews.co.kr": "기계설비신문",
    "insight.co.kr": "인사이트"
}

def get_media_name(url):
    domain = urlparse(url).netloc.lower().replace("www.", "")
    return MEDIA_DOMAIN_MAP.get(domain, domain or "언론사 정보 없음")

def clean_title(title):
    clean = re.sub('<[^>]+>', '', title)
    clean = re.sub(r"\s*-\s*[^-]+$", "", clean)
    clean = clean.replace("&quot;", "\"").replace("&#39;", "'")
    return clean

def is_within_date_range(pubdate, from_date, to_date):
    try:
        dt = email.utils.parsedate_to_datetime(pubdate)
        dt = dt.astimezone(kst)
        return from_date <= dt <= to_date
    except Exception as e:
        print("날짜 파싱 오류:", e)
        return False

def get_naver_news_entries(query, use_filter, start_dt=None, end_dt=None, limit=20):
    url = "https://openapi.naver.com/v1/search/news.json"
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }
    params = {
        "query": query,
        "display": min(limit, 100),
        "start": 1,
        "sort": "date"
    }
    response = requests.get(url, headers=headers, params=params)
    print(f"DEBUG: API 응답 코드: {response.status_code}")
    if response.status_code != 200:
        messagebox.showerror("API 오류", f"네이버 API 요청 실패: {response.status_code}")
        return []
    data = response.json()
    print(f"DEBUG: 응답 데이터 개수: {len(data.get('items', []))}")

    items = response.json().get("items", [])
    results = []

    for item in items:
        pub_str = item.get("pubDate", "")
        try:
            pub_dt = email.utils.parsedate_to_datetime(pub_str).astimezone(kst)
            print(f"🕒 뉴스 시간: {pub_dt}")
        except Exception as e:
            print(f"❌ 날짜 파싱 오류: {e} / pubDate: {pub_str}")
            pub_dt = None

        # 필터링 조건 처리
        if use_filter and pub_dt and start_dt and end_dt:
            if start_dt <= pub_dt <= end_dt:
                results.append((pub_dt, item))
        elif pub_dt:
            results.append((pub_dt, item))

    # 최신 뉴스가 위로 오도록 정렬
    results.sort(key=lambda x: x[0], reverse=True)

    return results[:limit]

def get_naver_news_entries_multi(queries, use_filter=False, start_dt=None, end_dt=None, limit=20):
    all_results = []
    seen_urls = set()
    for q in queries:
        entries = get_naver_news_entries(q, use_filter, start_dt, end_dt, limit)
        for entry in entries:
            pub_dt, item = entry
            url = item.get("link", "")
            if url not in seen_urls:
                seen_urls.add(url)
                all_results.append(entry)
    all_results.sort(key=lambda x: x[0], reverse=True)
    return all_results[:limit]

def generate_output(entries, queries, file_path, file_type, time_range_info=""):
    if file_type == "html":
        if not file_path.lower().endswith(".html"):
            file_path += ".html"
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("<html><head><meta charset='utf-8'><title>뉴스 검색 결과</title></head><body>")
            f.write(f"<h2>뉴스 검색 결과: '{', '.join(queries)}'</h2>")
            if time_range_info:
                f.write(f"<p><strong>검색 시간대:</strong> {time_range_info}</p>")
            f.write("<ul>")
            for pub_dt, entry in entries:
                url = entry.get("link", "#")
                title = clean_title(entry.get("title", ""))
                media = get_media_name(entry.get("originallink", ""))
                time = pub_dt.strftime("%Y.%m.%d. %H:%M")
                f.write(f"<li>({media}) <a href='{url}' target='_blank'>{title}</a> [{time}]</li>")
            f.write("</ul></body></html>")
    else:
        if not file_path.lower().endswith(".txt"):
            file_path += ".txt"
        with open(file_path, "w", encoding="utf-8") as f:
            if time_range_info:
                f.write(f"🔎 검색 시간대: {time_range_info}\n\n")
            for i, (pub_dt, entry) in enumerate(entries, 1):
                url = entry.get("link", "#")
                title = clean_title(entry.get("title", ""))
                media = get_media_name(entry.get("originallink", ""))
                time = pub_dt.strftime("%Y.%m.%d. %H:%M")
                f.write(f"{i}. ({media}) {title} [{time}]\n{url}\n\n")
    return os.path.abspath(file_path)

def run_gui():
    root = tk.Tk()
    root.title("뉴스 검색기")
    root.geometry("620x750")
    root.configure(bg="#4a90e2")

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure('TFrame', background="#4a90e2")
    style.configure('TLabel', background="#4a90e2", foreground="white", font=('맑은 고딕', 11))
    style.configure('Header.TLabel', font=('맑은 고딕', 14, 'bold'))
    style.configure('TButton', font=('맑은 고딕', 11, 'bold'), foreground='white', background='#357ae8')
    style.configure('TScale', troughcolor='#a9c1f7', background='#357ae8')

    keyword = tk.StringVar()
    use_filter = tk.IntVar(value=0)
    file_format = tk.StringVar(value="html")
    news_limit = tk.IntVar(value=20)
    preview_entries = []
    time_info = tk.StringVar()
    time_range = {}

    def show_frame(frame):
        for f in (frame1, frame2, frame3, frame4):
            f.pack_forget()
        frame.pack(fill="both", expand=True, padx=15, pady=10)

    # --- Frame 1: 검색어 입력 ---
    frame1 = ttk.Frame(root)
    ttk.Label(frame1, text="🔍 검색어를 쉼표(,)로 구분하여 여러 개 입력하세요", style="Header.TLabel").pack(anchor="w", pady=(10,5))
    entry_keyword = ttk.Entry(frame1, textvariable=keyword, font=("맑은 고딕", 11), width=50)
    entry_keyword.pack(anchor="w", pady=5)
    ttk.Button(frame1, text="다음 ▶", command=lambda: show_frame(frame2)).pack(anchor="e", pady=10)

    # --- Frame 2: 시간 필터 사용 여부 ---
    frame2 = ttk.Frame(root)
    ttk.Label(frame2, text="🕒 시간 필터 사용 여부", style='Header.TLabel').pack(anchor="w", pady=(10,5))
    ttk.Radiobutton(frame2, text="사용함", variable=use_filter, value=1).pack(anchor='w')
    ttk.Radiobutton(frame2, text="사용 안 함", variable=use_filter, value=0).pack(anchor='w')
    ttk.Button(frame2, text="◀ 이전", command=lambda: show_frame(frame1)).pack(side="left", pady=10, padx=5)
    ttk.Button(frame2, text="다음 ▶", command=lambda: show_frame(frame4) if use_filter.get() == 1 else show_frame(frame3)).pack(side="right", pady=10, padx=5)

    # --- Frame 3: 시간 필터 미사용시 뉴스 제한 및 저장 ---
    frame3 = ttk.Frame(root)
    ttk.Label(frame3, text="📊 뉴스 개수 제한 (최대 100)", style="Header.TLabel").pack(anchor='w')
    scale = ttk.Scale(frame3, from_=5, to=100, variable=news_limit, orient="horizontal")
    scale.pack(fill='x', pady=5)
    count_label = ttk.Label(frame3, text=str(news_limit.get()), font=('맑은 고딕', 10, 'bold'), background="#4a90e2", foreground="white")
    count_label.pack(anchor='e', pady=(0, 10))
    def update_count_label(value):
        count_label.config(text=str(int(float(value))))
    scale.config(command=update_count_label)

    ttk.Label(frame3, text="📁 저장 파일 형식 선택", style="Header.TLabel").pack(anchor='w', pady=(10,0))
    ttk.Radiobutton(frame3, text="HTML", variable=file_format, value="html").pack(anchor="w")
    ttk.Radiobutton(frame3, text="TXT", variable=file_format, value="txt").pack(anchor="w")

    preview_btn = ttk.Button(frame3, text="🔍 뉴스 미리보기")
    preview_btn.pack(pady=10)

    status_label = tk.Label(frame3, text="🖱️ 기사 더블클릭 시 새창으로 열립니다.", foreground="white", background="#4a90e2", font=("맑은 고딕", 10))
    status_label.pack(anchor="w", pady=(0, 5))

    listbox = tk.Listbox(frame3, height=12, font=('맑은 고딕', 10), activestyle='none', selectmode="extended")
    listbox.pack(fill='both', expand=True)

    ttk.Button(frame3, text="💾 모든 뉴스 저장", command=lambda: save_news()).pack(pady=(10, 5))
    ttk.Button(frame3, text="📥 선택 뉴스 저장(Ctrl버튼으로 중복선택 가능)", command=lambda: save_selected_news()).pack(pady=(0, 5))
    ttk.Button(frame3, text="◀ 이전", command=lambda: show_frame(frame2)).pack(pady=(0, 10))

    # --- Frame 4: 시간 필터 사용시 시작/종료 날짜 및 시간 입력 ---
    frame4 = ttk.Frame(root)
    ttk.Label(frame4, text="⏱ 시작 날짜 및 시간", style='Header.TLabel').grid(row=0, column=0, sticky='w')
    s_date = DateEntry(frame4, width=12, date_pattern='yyyy-MM-dd')
    ttk.Label(frame4, text="시", style='TLabel').grid(row=0, column=1, sticky='w', padx=(5,0))
    s_hour = ttk.Entry(frame4, width=3)
    ttk.Label(frame4, text="분", style='TLabel').grid(row=0, column=2, sticky='w', padx=(5,0))
    s_min = ttk.Entry(frame4, width=3)
    s_date.grid(row=1, column=0)
    s_hour.grid(row=1, column=1)
    s_min.grid(row=1, column=2)

    ttk.Label(frame4, text="⏱ 종료 날짜 및 시간", style='Header.TLabel').grid(row=2, column=0, sticky='w', pady=(10,0))
    e_date = DateEntry(frame4, width=12, date_pattern='yyyy-MM-dd')
    ttk.Label(frame4, text="시", style='TLabel').grid(row=2, column=1, sticky='w', padx=(5,0))
    e_hour = ttk.Entry(frame4, width=3)
    ttk.Label(frame4, text="분", style='TLabel').grid(row=2, column=2, sticky='w', padx=(5,0))
    e_min = ttk.Entry(frame4, width=3)
    e_date.grid(row=3, column=0)
    e_hour.grid(row=3, column=1)
    e_min.grid(row=3, column=2)

    def to_frame3():
        try:
            s_hour_val = int(s_hour.get()) if s_hour.get().isdigit() and 0 <= int(s_hour.get()) <= 23 else 0
            s_min_val = int(s_min.get()) if s_min.get().isdigit() and 0 <= int(s_min.get()) <= 59 else 0
            e_hour_val = int(e_hour.get()) if e_hour.get().isdigit() and 0 <= int(e_hour.get()) <= 23 else 23
            e_min_val = int(e_min.get()) if e_min.get().isdigit() and 0 <= int(e_min.get()) <= 59 else 59

            start_str = f"{s_date.get()} {s_hour_val:02d}:{s_min_val:02d}"
            end_str = f"{e_date.get()} {e_hour_val:02d}:{e_min_val:02d}"

            start_dt = kst.localize(datetime.strptime(start_str, "%Y-%m-%d %H:%M"))
            end_dt = kst.localize(datetime.strptime(end_str, "%Y-%m-%d %H:%M"))

            if start_dt > end_dt:
                messagebox.showerror("시간 오류", "시작 시간이 종료 시간보다 늦습니다.")
                return

            time_range["start"] = start_dt
            time_range["end"] = end_dt

            print("DEBUG: start_dt =", start_dt, "end_dt =", end_dt)  # 여기 추가

            time_info.set(f"{start_dt.strftime('%Y.%m.%d %H:%M')} ~ {end_dt.strftime('%Y.%m.%d %H:%M')}")
            show_frame(frame3)
        except Exception as e:
            messagebox.showerror("시간 형식 오류", str(e))

    ttk.Button(frame4, text="◀ 이전", command=lambda: show_frame(frame2)).grid(row=4, column=0, pady=10, sticky='w')
    ttk.Button(frame4, text="다음 ▶", command=to_frame3).grid(row=4, column=2, pady=10, sticky='e')

    def preview_news():
        listbox.delete(0, tk.END)
        status_label.config(text="뉴스 로딩 중...")
        root.update()
        queries = [q.strip() for q in keyword.get().split(",") if q.strip()]
        if not queries:
            messagebox.showerror("입력 오류", "검색어를 하나 이상 입력하세요.")
            status_label.config(text="")
            return
        entries = get_naver_news_entries_multi(
            queries,
            use_filter=use_filter.get() == 1,
            start_dt=time_range.get("start"),
            end_dt=time_range.get("end"),
            limit=news_limit.get()
        )

        print("미리보기용 뉴스:", entries)

        if not entries:
            status_label.config(text="뉴스가 없습니다.")
            return
        preview_entries.clear()
        preview_entries.extend(entries)
        # ---- 여기서 중복 삽입 라인 삭제 ----
        for i, (dt, entry) in enumerate(entries, 1):
            try:
                title = clean_title(entry.get("title", ""))
                media = get_media_name(entry.get("originallink", ""))
                time = dt.strftime("%Y.%m.%d %H:%M")
                listbox.insert(tk.END, f"{i}. ({media}) {title} [{time}]")
            except Exception as e:
                print(f"❌ listbox insert 오류: {e}, entry={entry}")
        status_label.config(text=f"{len(entries)}건 뉴스 로드 완료, 더블클릭 시 새창으로 열립니다.")
        print(f"DEBUG: preview_news() - start_dt={time_range.get('start')}, end_dt={time_range.get('end')}")

    def on_listbox_click(event):
        idx = listbox.curselection()
        if not idx:
            return
        index = idx[0]
        if index >= len(preview_entries):
            return
        url = preview_entries[index][1].get("link", "")
        if url:
            webbrowser.open(url)

    def save_news():
        if not preview_entries:
            messagebox.showerror("오류", "먼저 미리보기를 실행하세요.")
            return
        filetype = file_format.get()
        path = filedialog.asksaveasfilename(
            defaultextension=f".{filetype}",
            filetypes=[("HTML 파일", "*.html"), ("텍스트 파일", "*.txt")],
            title="저장할 파일 선택"
        )
        if not path:
            return
        saved_path = generate_output(preview_entries, [q.strip() for q in keyword.get().split(",")], path, filetype, time_info.get())
        messagebox.showinfo("저장 완료", f"{len(preview_entries)}건 저장 완료\n📁 {saved_path}")

    def save_selected_news():
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("선택 없음", "저장할 기사를 선택하세요.")
            return
        selected_entries = [preview_entries[i] for i in selected_indices if i < len(preview_entries)]
        if not selected_entries:
            messagebox.showerror("오류", "유효한 뉴스 항목이 없습니다.")
            return
        filetype = file_format.get()
        path = filedialog.asksaveasfilename(
            defaultextension=f".{filetype}",
            filetypes=[("HTML 파일", "*.html"), ("텍스트 파일", "*.txt")],
            title="선택 뉴스 저장"
        )
        if not path:
            return
        saved_path = generate_output(selected_entries, [q.strip() for q in keyword.get().split(",")], path, filetype, time_info.get())
        messagebox.showinfo("저장 완료", f"{len(selected_entries)}건 저장 완료\n📁 {saved_path}")

    listbox.bind("<Double-Button-1>", on_listbox_click)
    preview_btn.config(command=preview_news)

    show_frame(frame1)
    root.mainloop()

if __name__ == "__main__":
    run_gui()
