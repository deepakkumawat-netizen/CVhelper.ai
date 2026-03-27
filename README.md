# CVhelper.ai — AI-Powered Full Course Generator

> Generate complete, structured courses on **any skill** — with modules, classes, quizzes, downloadable PPT slides, and PDF notes — powered by Google Gemini or Groq AI.

---

## 🚀 Features

- **AI Course Generation** — Type any skill and instantly get a full 4-module course with unique class content
- **Rich Class Content** — Each class includes a video script, key points, learning objectives, real-world examples, and key stats
- **MCQ Quizzes** — Auto-generated 10-question multiple choice quiz at the end of every module
- **Downloadable PPT** — Export any class or full course as a PowerPoint presentation
- **Downloadable PDF Notes** — Save class notes as a formatted PDF
- **AI Chat Assistant** — Ask questions about any class topic in a built-in chat panel
- **Voice Input** — Use your microphone to search for skills or ask the AI assistant
- **Progress Tracking** — Track completed classes with a visual progress bar
- **All content in English** — Consistent, high-quality English content for every course

---

## 🛠️ Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML, CSS, Vanilla JavaScript |
| Backend | Python 3, Flask, Flask-CORS |
| AI Engine | Google Gemini 2.0 Flash (primary) / Groq LLaMA 3.3 70B (fallback) |
| PDF Generation | fpdf2 |
| PPTX Generation | pptxgenjs (Node.js) |
| Web Search | Wikipedia API + DuckDuckGo (free, no key needed) |

---

## 📁 Project Structure

```
cvhelper.ai/
├── app.py               # Flask backend — all API routes and AI logic
├── gen.js               # Node.js script for PowerPoint generation (pptxgenjs)
├── index.html           # Frontend — single-page app (HTML + CSS + JS)
├── requirements.txt     # Python dependencies
├── package.json         # Node.js dependencies
├── .env                 # API keys (not committed to git)
└── README.md            # This file
```

---

## ⚙️ Setup & Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/cvhelper.ai.git
cd cvhelper.ai
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. Install Node.js dependencies

```bash
npm install
```

### 4. Configure API Keys

Create a `.env` file in the project root:

```env
# Get a free key at: https://aistudio.google.com
GEMINI_API_KEY=your_gemini_api_key_here

# Optional fallback — get a free key at: https://console.groq.com
GROQ_API_KEY=your_groq_api_key_here

# Optional: force a specific engine ("gemini" | "groq" | "auto")
AI_ENGINE=auto

# Optional: change the port (default: 5001)
PORT=5001
```

> **Note:** You only need one key — Gemini OR Groq. If both are set, Gemini is used first with Groq as automatic fallback.

### 5. Run the app

```bash
python app.py
```

Then open your browser and go to: **http://localhost:5001**

---

## 🔑 Getting Free API Keys

| Provider | Link | Free Tier |
|----------|------|-----------|
| Google Gemini | https://aistudio.google.com | ✅ Free |
| Groq | https://console.groq.com | ✅ Free |

---

## 📡 API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/api/health` | Check backend status and AI engine |
| `GET` | `/api/test` | Test AI connectivity |
| `GET` | `/api/categories` | Get course categories list |
| `POST` | `/api/generate-course` | Generate a full course structure |
| `POST` | `/api/generate-class-content` | Generate content for a specific class |
| `POST` | `/api/generate-mcq` | Generate a 10-question quiz |
| `POST` | `/api/chat` | AI chat assistant |
| `POST` | `/api/generate-pptx` | Generate PowerPoint for a class |
| `POST` | `/api/generate-course-pptx` | Generate PowerPoint for full course |
| `POST` | `/api/generate-pdf` | Generate PDF notes for a class |

### Example: Generate a Course

```bash
curl -X POST http://localhost:5001/api/generate-course \
  -H "Content-Type: application/json" \
  -d '{
    "skill": "Python Programming",
    "level": "Beginner",
    "duration": "4 weeks"
  }'
```

---

## 🎓 How It Works

1. **User enters a skill** (e.g. "Guitar Playing", "Python Programming", "Stock Market Investing")
2. **AI generates a unique course structure** — 4 modules × 2 teaching classes + 1 quiz each
3. **User clicks a class** — AI generates rich content: key points, objectives, real-world examples, video script
4. **User can download** class notes as PDF or slides as PPTX
5. **Module quiz** tests knowledge with 10 auto-generated MCQs
6. **AI Chat** answers any questions about the course topic

---

## 📋 Requirements

- Python 3.8+
- Node.js 16+
- A free Gemini API key **or** Groq API key

---

## 🧠 AI Engine Logic

The backend uses a smart auto-selection strategy:

```
AI_ENGINE=auto (default)
  → Try Gemini first
  → If Gemini fails, automatically fall back to Groq
  → If both fail, return a structured fallback response

AI_ENGINE=gemini → Use only Gemini
AI_ENGINE=groq   → Use only Groq
```

---

## 📦 Python Dependencies

```
flask
flask-cors
requests
python-dotenv
fpdf2
```

---

## 📦 Node.js Dependencies

```
pptxgenjs ^4.0.1
```

---

## 🐛 Troubleshooting

**Backend not starting?**
- Make sure Python 3.8+ is installed
- Run `pip install -r requirements.txt`
- Check your `.env` file has a valid API key

**PPT download not working?**
- Make sure Node.js 16+ is installed
- Run `npm install` to install pptxgenjs
- Check that `gen.js` is in the same directory as `app.py`

**Content generating in wrong language?**
- All content is now locked to English
- If you see non-English content, restart the backend: `python app.py`

**AI returning errors?**
- Visit http://localhost:5001/api/test to check connectivity
- Verify your API key is valid and has not expired
- Try switching `AI_ENGINE` to `groq` or `gemini` in `.env`

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 🙏 Credits

Built with:
- [Google Gemini](https://aistudio.google.com) — AI content generation
- [Groq](https://console.groq.com) — Fast LLM fallback
- [pptxgenjs](https://gitbrent.github.io/PptxGenJS/) — PowerPoint generation
- [fpdf2](https://pyfpdf.github.io/fpdf2/) — PDF generation
- [Wikipedia API](https://www.mediawiki.org/wiki/API:Main_page) — Free knowledge base
