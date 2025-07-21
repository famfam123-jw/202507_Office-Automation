import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import pytesseract
import cv2
import numpy as np
import os
import re

# Tesseract 경로 (설치 위치에 따라 맞게 수정)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def preprocess_image_cv(img_path):
    # 한글 경로도 OK
    img_array = np.fromfile(img_path, np.uint8)
    img_cv = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    if img_cv is None:
        raise Exception(f"이미지 파일을 읽지 못했습니다: {img_path}")
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    proc = cv2.medianBlur(thresh, 3)
    return proc

def clean_text(text):
    text = re.sub(r'[■◆□·�※▶★●◎▪️]', ' ', text)
    text = re.sub(r'\b[a-zA-Z]{1,3}\b', '', text)      # 3자 이하 영단어
    text = re.sub(r'\b[0-9]{1,4}\b', '', text)         # 1~4자리 숫자
    text = re.sub(r'\b[ㄱ-ㅎㅏ-ㅣ]{1,2}\b', '', text)  # 한글 낱자
    text = re.sub(r'[^\w\s가-힣.,?!·]', '', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

def extract_and_save_text(image_paths, save_target, mode):
    all_text = ""
    for img_path in image_paths:
        try:
            img_cv = preprocess_image_cv(img_path)
            text = pytesseract.image_to_string(img_cv, lang='kor+eng', config='--psm 6')
            cleaned = clean_text(text)
            if mode == "individual":
                base_name = os.path.splitext(os.path.basename(img_path))[0]
                save_path = os.path.join(save_target, f"{base_name}.txt")
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(cleaned)
            else:
                all_text += f"\n\n===== {os.path.basename(img_path)} =====\n\n{cleaned}"
        except Exception as e:
            messagebox.showerror("오류 발생", f"{img_path} 처리 중 문제: {e}")
    if mode == "merged":
        with open(save_target, 'w', encoding='utf-8') as f:
            f.write(all_text)
    messagebox.showinfo("완료", "✅ 텍스트 추출 및 저장이 완료되었습니다!")
    reset_ui()

def add_image():
    files = filedialog.askopenfilenames(
        title="이미지 선택",
        filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp")]
    )
    for file in files:
        if file not in selected_images:
            selected_images.append(file)
            listbox.insert(tk.END, os.path.basename(file))

def remove_selected():
    selected = listbox.curselection()
    for i in reversed(selected):
        listbox.delete(i)
        del selected_images[i]

def move_up():
    idx = listbox.curselection()
    for i in idx:
        if i > 0:
            selected_images[i], selected_images[i-1] = selected_images[i-1], selected_images[i]
            update_listbox()
            listbox.select_set(i-1)

def move_down():
    idx = listbox.curselection()
    for i in reversed(idx):
        if i < len(selected_images)-1:
            selected_images[i], selected_images[i+1] = selected_images[i+1], selected_images[i]
            update_listbox()
            listbox.select_set(i+1)

def update_listbox():
    listbox.delete(0, tk.END)
    for path in selected_images:
        listbox.insert(tk.END, os.path.basename(path))

def select_path():
    if save_mode.get() == "merged":
        folder = filedialog.askdirectory(title="병합 저장할 폴더 선택")
        if folder:
            output_target.set(folder)
            path_label.config(text=f"📁 저장 폴더: {folder}")
    else:
        folder = filedialog.askdirectory(title="개별 저장할 폴더 선택")
        if folder:
            output_target.set(folder)
            path_label.config(text=f"📁 저장 폴더: {folder}")

def check_save_ready():
    if not selected_images:
        messagebox.showwarning("⚠️ 이미지 없음", "이미지를 추가해주세요.")
        return False
    if not output_target.get():
        messagebox.showwarning("⚠️ 저장 위치 없음", "저장할 위치를 지정해주세요.")
        return False
    if save_mode.get() == "merged":
        filename = filename_var.get().strip()
        if not filename:
            messagebox.showwarning("⚠️ 파일명 없음", "병합 저장할 파일명을 입력해주세요.")
            return False
        if not filename.lower().endswith('.txt'):
            filename += ".txt"
        merged = os.path.join(output_target.get(), filename)
        try:
            with open(merged, 'w', encoding='utf-8') as f:
                pass
            os.remove(merged)
        except Exception as e:
            messagebox.showwarning("⚠️ 잘못된 파일명/경로", f"저장 파일을 만들 수 없습니다:\n{e}")
            return False
        output_file.set(merged)
    else:
        output_file.set(output_target.get())
    return True

def start_processing():
    if check_save_ready():
        if save_mode.get() == "merged":
            extract_and_save_text(selected_images.copy(), output_file.get(), "merged")
        else:
            extract_and_save_text(selected_images.copy(), output_file.get(), "individual")

def reset_ui():
    selected_images.clear()
    listbox.delete(0, tk.END)
    output_target.set("")
    output_file.set("")
    filename_var.set("")
    path_label.config(text="❌ 저장 위치 미지정")

def on_mode_change():
    reset_file_widgets()
    reset_path_label()

def reset_file_widgets():
    if save_mode.get() == "merged":
        filename_frame.pack(pady=(2,2))
    else:
        filename_frame.forget()
        filename_var.set("")

def reset_path_label():
    path_label.config(text="❌ 저장 위치 미지정")

# --- UI ----
app = tk.Tk()
app.title("📸 카드뉴스 OCR 추출기")
app.geometry("540x600")
app.resizable(False, False)

selected_images = []
output_target = tk.StringVar()
output_file = tk.StringVar()
filename_var = tk.StringVar()
save_mode = tk.StringVar(value="individual")

tk.Label(app, text="🖼️ 이미지 추가 및 순서 조절").pack(pady=(10, 3))
tk.Button(app, text="이미지 추가", command=add_image, bg="#ddeeff").pack()
listbox = tk.Listbox(app, selectmode=tk.SINGLE, height=8, width=60)
listbox.pack(pady=5)

btn_frame = tk.Frame(app)
btn_frame.pack()
tk.Button(btn_frame, text="⏫ 위로", command=move_up).grid(row=0, column=0, padx=5)
tk.Button(btn_frame, text="⏬ 아래로", command=move_down).grid(row=0, column=1, padx=5)
tk.Button(btn_frame, text="🗑️ 삭제", command=remove_selected).grid(row=0, column=2, padx=5)

tk.Label(app, text="📝 저장 방식 선택").pack(pady=(15, 3))
mode_frame = tk.Frame(app)
mode_frame.pack()
tk.Radiobutton(mode_frame, text="이미지별 개별 저장", variable=save_mode, value="individual", command=on_mode_change).pack(side=tk.LEFT, padx=10)
tk.Radiobutton(mode_frame, text="모두 병합하여 저장", variable=save_mode, value="merged", command=on_mode_change).pack(side=tk.LEFT, padx=10)

filename_frame = tk.Frame(app)
tk.Label(filename_frame, text="병합 파일명(.txt):").pack(side=tk.LEFT)
tk.Entry(filename_frame, textvariable=filename_var, width=30).pack(side=tk.LEFT)
reset_file_widgets()

tk.Label(app, text="💾 저장 위치 선택").pack(pady=(15, 3))
tk.Button(app, text="저장 위치 선택", command=select_path, bg="#ddeeff").pack()
path_label = tk.Label(app, text="❌ 저장 위치 미지정", fg="gray")
path_label.pack()

tk.Button(app, text="🔍 텍스트 추출 시작", command=start_processing, bg="#cdeac0", width=25).pack(pady=25)

app.mainloop()
