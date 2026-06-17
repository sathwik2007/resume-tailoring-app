# ATS Resume Tailor (Gemini Edition)

A small CLI tool that tailors a `.docx` resume to a job description using Google Gemini AI. It extracts the document XML, sends the resume + job description to Gemini to generate suggested text replacements (as JSON), applies those replacements directly into the OOXML while preserving run-level formatting, and repacks the edited document back into a `.docx`.

Files
- [tailor_resume.py](tailor_resume.py): Main script that performs extraction, AI tailoring, and repacking.
- [job_description.txt](job_description.txt): Example job description file (or can be used as input).

Why this exists
- Helps optimize resumes for Applicant Tracking Systems (ATS) by incorporating JD keywords and tone without inventing experience.
- Preserves original Word formatting by applying changes at the XML run level.

Requirements
- Python 3.10+ (3.12 recommended)
- Google Gemini API key (free at https://aistudio.google.com/app/apikey)

Python packages
- google-genai
- lxml
- python-docx

Quickstart (local)

1) Create and activate a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2) Install dependencies

```bash
pip install --upgrade pip
pip install google-genai lxml python-docx
```

3) Export your Gemini API key

```bash
export GEMINI_API_KEY=your-key-here
```

4) Run the script (example)

```bash
python3 tailor_resume.py --resume resume.docx --jd job_description.txt --output tailored_resume.docx
```

Notes
- The `--jd` argument accepts either a path to a `.txt` JD file or raw JD text. For inline text, use `--jd "<paste job description here>"`.
- The script requires the resume to be a `.docx` file and will preserve formatting by editing `word/document.xml` directly.
- If Gemini returns JSON wrapped in markdown fences, the script strips them before parsing.

Privacy & cost
- Text is sent to Google Gemini; ensure you are comfortable sending resume text to the API.
- Gemini usage may incur quotas; check your account limits.

Troubleshooting
- If you see a deprecation warning for `google.generativeai`, install `google-genai` and ensure imports in `tailor_resume.py` use `from google import genai`.
- If `GEMINI_API_KEY` is not set, the script will exit with a helpful message.


