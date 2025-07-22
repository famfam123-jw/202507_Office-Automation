# Office Automation Scripts

A collection of Python GUI tools for news monitoring using the Naver API and extracting text from card news images (OCR).

## 📁 Included Files

- `monitoring_5_github.py`: News monitoring/search tool using Naver News API (with GUI)
- `texter_github.py`: Card news image OCR text extractor and saver (with GUI)

---

## 1️⃣ monitoring_5_github.py (Naver News Monitoring Tool)

### Features
- Search and filter news articles using the Naver News API
- Support for multiple keywords search at once
- Date/time range filtering (to the minute)
- Preview news items in-app, open article links in browser
- Save search results as HTML or TXT files
- Auto-display media sources for major Korean press

### How to Run
1. Requires Python 3.x
2. Install required packages:
    ```
    pip install requests tkcalendar pytz
    ```
    *(tkinter is usually included with standard Python installations)*
3. Set up your Naver API credentials in the code
4. Run the program:
    ```
    python monitoring_5_github.py
    ```

---

## 2️⃣ texter_github.py (Card News Image OCR Extractor)

### Features
- Extract text from images (supports Korean & English, via Tesseract OCR)
- Process multiple images at once; save as individual .txt files or merge into one file
- Reorder and delete images within the list UI
- Automatic preprocessing (binarization, denoising) for better OCR
- Supports various image formats (.jpg, .png, etc.)

### How to Run
1. Requires Python 3.x
2. Install required packages:
    ```
    pip install pillow pytesseract opencv-python numpy
    ```
3. Install [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
   (On Windows, the default path is usually: `C:\Program Files\Tesseract-OCR\tesseract.exe`)
4. Make sure `pytesseract.pytesseract.tesseract_cmd` in the code matches your Tesseract installation path.
5. Run the program:
    ```
    python texter_github.py
    ```

---

## Notes

- Both programs use a graphical user interface (GUI). When you run the scripts, a window should open.
- **For Naver News Monitoring**, your own Naver API credentials are required.
- **For OCR extraction**, make sure Tesseract is installed and the Korean language pack (`kor.traineddata`) is present for accurate Hangul recognition.
- If sharing the project publicly, **do not upload your personal API keys or secrets to GitHub!**

---

## License

MIT License

Copyright (c) famfam123-jw.github

---

## Contact / Contributing

Pull requests and issues are welcome!

