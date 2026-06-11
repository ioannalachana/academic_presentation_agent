#  Academic Presentation Agent — Web App

A web tool anyone can use.  
Upload a research paper → get a slide outline, poster layout, and downloadable PowerPoint.

**Powered by Google Gemini (free API).**

---



### Step 1 — Push to GitHub
1. Create a new repo at [github.com](https://github.com)
2. Upload all files in this folder

### Step 2 — Deploy on Render (free hosting)
1. Go to [render.com](https://render.com) and sign up (free)
2. Click **New** → **Web Service**
3. Connect your GitHub repo
4. Render auto-detects `render.yaml` — just click **Deploy**
5. In the dashboard, go to **Environment** and add:
   - Key: `GEMINI_API_KEY`
   - Value: your key from [aistudio.google.com](https://aistudio.google.com)

Your app will be live at `https://your-app-name.onrender.com` 🎉

---

##  Run locally (for testing)

```bash
pip install -r requirements.txt
export GEMINI_API_KEY=your_key_here
uvicorn main:app --reload
```
Then open [http://localhost:8000](http://localhost:8000)

---

##  Project structure

```
academic-agent-web/
├── main.py            ← FastAPI backend (Python)
├── requirements.txt   ← Python dependencies
├── render.yaml        ← Render.com deployment config
└── static/
    └── index.html     ← Frontend (HTML + CSS + JS)
```

---


##  Important

- The `GEMINI_API_KEY` stays on the server — users never see it
- The free Gemini tier allows ~1,500 requests/day — enough for hundreds of papers
- Generated `.pptx` files are stored temporarily in memory and cleared on restart
