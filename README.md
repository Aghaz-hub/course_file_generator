# Course File Generator (Minimal vs Full)

Streamlit app that generates a structured **course file (.docx)** aligned with the Green Computing / MRU School of Engineering sample.

## Modes

| Mode | Use when | What you enter |
|------|----------|----------------|
| **Minimal** | Need a quick skeleton | Course identity + optional syllabus & COs |
| **Full** | Want most design sections filled | Syllabus, COs, targets, mappings, CAP, lesson plan, experiments, optional teaching/attainment |

## Shared defaults (same for everyone)

These are **built into the app** — faculty do **not** re-upload per course:

- University Vision & Mission  
- School of Engineering Vision & Mission  
- Department Vision & Mission (CST and/or ME)  
- Graduate Attributes, POs (summaries)  
- PEOs / PSOs by program family (CSE / ME)  
- KSA profile summary  
- Academic calendar note (attach current session calendar once if needed)

Select **Department(s)** and **Program family** on the form so the correct blocks are inserted.

## Run locally

```bash
cd course_file_generator
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Push this folder (or repo) to GitHub  
2. [share.streamlit.io](https://share.streamlit.io) → New app  
3. Main file: `app.py`  
4. Deploy  

## Output

- Downloadable `.docx` named `CourseFile_{CODE}_{Name}_{Minimal|Full}.docx`  
- Full checklist table (sections 1–31 + Appendices A–E)  
- Placeholders `[To be filled by faculty]` wherever content was not provided  

## Tech

- Python · Streamlit · python-docx · pdfplumber (syllabus PDF) · pandas/openpyxl (optional Excel later)
