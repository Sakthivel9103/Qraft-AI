# KMG College - Question Paper Assistant (Python / Flask)

## Setup

```bash
pip install -r requirements.txt
```

`.env` file create pannunga (project root-la, app.py irukura idathulaye):
```
GEMINI_API_KEY=your-gemini-key-here
```
(No quotes, no spaces around the =)

Run:
```bash
python app.py
```

Browser: http://localhost:5000

## What's new in this version

1. **Real error messages while debugging** - `/api/generate` ippo actual
   error-ah JSON-la anupidum (chat bubble-la screenshot edukka mudiyum,
   terminal thevai illa). Ella vela seyyum-nu confirm aana, `app.py`-la
   `except Exception as e:` block-ah generic message-ku maathikonga.

2. **PDF or DOCX download** - Settings panel-la "Download format" dropdown,
   Word (.docx) illa PDF (.pdf) select pannalam.

3. **Syllabus upload (+ button)** - Composer-la "+" click pannitu PDF,
   DOCX, illa TXT file upload pannalam, adhu automatic-a message box-la
   text-a fill agum.

4. **KMG College of Arts and Science** - Home screen-la college name
   update pannirukom. Logo replace panna:
   `templates/index.html`-la:
   ```html
   <span id="logoInitials">KMG</span>
   ```
   idha:
   ```html
   <img src="/static/logo.png" alt="College logo">
   ```
   nu maathi, logo file-ah `static/logo.png` nu save pannunga.

## Flow

1. Home screen-la CIA 1 / CIA 2 / Model Exam / Semester Exam - edho
   ondrai click pannunga.
2. Settings (gear icon) - K-Level range, total questions, marks,
   download format edit pannalam.
3. Syllabus: type pannalam OR "+" click panni file upload pannalam.
4. Send pannunga - question paper chat-la varum, download link kidaikum.

## Files

```
final_bot/
├── app.py
├── requirements.txt
├── .env                    (create this yourself, don't share/commit it)
├── templates/
│   └── index.html
├── static/
│   ├── style.css
│   └── script.js
└── generated/               (auto-created)
```
