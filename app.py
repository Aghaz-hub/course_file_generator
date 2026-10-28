"""
Course File Generator (Minimal vs Full)
Generates a structured course-file DOCX aligned with the Green Computing sample.
Shared institutional content is built-in; faculty only enter course-specific data.
"""

import re
import traceback
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# Optional: syllabus extraction
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

st.set_page_config(
    page_title="Course File Generator",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# SHARED INSTITUTIONAL DEFAULTS (same for everyone — edit here or via sidebar)
# =============================================================================

DEFAULTS = {
    "university_vision": (
        "To educate students in frontier areas of knowledge enabling them to take up "
        "challenges as ethical and responsible global citizens."
    ),
    "university_mission": (
        "To impart outcome based holistic education; to disseminate education in frontier areas; "
        "to produce globally competitive, ethical and socially responsible human resources; "
        "to produce human resources sensitive to issues of Environment and Sustainable Development; "
        "to develop Environment and Sustainable development as a thrust area of research and development."
    ),
    "school_vision": (
        "To build a future where education and innovation, empowered by collaboration, "
        "create a sustainable global impact."
    ),
    "school_mission": (
        "SM1 Quality Education: Provide a rigorous, industry-aligned education that equips students "
        "with the foundational knowledge and technical expertise to excel in the academic and professional world.\n"
        "SM2 Skill Development: Enhance skills by integrating hands-on learning experiences to promote "
        "critical thinking, problem-solving, and creativity.\n"
        "SM3 Research and Innovation: Foster a culture of research and innovation that leads to new "
        "technologies and solutions to real-world problems.\n"
        "SM4 Industry Collaboration: Build strong, strategic partnerships with industry leaders to bridge "
        "the gap between academic learning and real-world application.\n"
        "SM5 Ethics and Sustainability: Nurture engineers to prioritize ethical decision-making for "
        "developing responsible solutions with a focus on a sustainable future."
    ),
    "dept_cst_vision": (
        "To be a role model computer science department which imparts research based multidisciplinary "
        "competencies to foster innovation, creativity and ethical responsibility among students for "
        "transformative change in society and industry with sustainable solutions."
    ),
    "dept_cst_mission": (
        "DM1 Quality Education: Build student competency through industry-aligned curriculum delivered "
        "through relevant pedagogy and assessment.\n"
        "DM2 Skill Development: Deploy employability skills in students by integrating hands-on learning "
        "labs experiences, interdisciplinary projects, and internships to imbibe critical thinking, "
        "problem solving, and creativity.\n"
        "DM3 Research and Innovation: Collaborate with NI institutions and industry to foster a culture "
        "of research and innovation by supporting interdisciplinary projects.\n"
        "DM4 Industry Collaboration: Build strong, strategic partnerships with industry for capacity "
        "building in innovation and real time knowledge.\n"
        "DM5 Ethics and Sustainability: To produce ethical and socially responsible engineers who are "
        "allied with sustainable solutions for the industry."
    ),
    "dept_me_vision": (
        "To be a center to nurture competent mechanical engineers equipped with strong technical knowledge, "
        "innovative thinking and ethical values, who contribute globally to industry, research and society "
        "through sustainable and future ready solutions."
    ),
    "dept_me_mission": (
        "To provide quality education in mechanical engineering through an industry-aligned curriculum, "
        "hands-on learning, and modern pedagogical practices that foster technical excellence and lifelong learning.\n"
        "To equip students with globally relevant technical and professional skills through hands-on training, "
        "internships, industrial exposure, and problem to project based learning.\n"
        "To cultivate a research-driven mindset by fostering innovation, multidisciplinary projects, and collaboration.\n"
        "To nurture socially responsible engineers by instilling ethical values and environmental consciousness.\n"
        "To strengthen industry-institute partnerships and community engagement for solving real-world problems."
    ),
    "gas_summary": (
        "GA1 Engineering Knowledge · GA2 Problem Analysis · GA3 Design/Development of Solutions · "
        "GA4 Investigation · GA5 Tool Usage · GA6 The Engineer and the World · GA7 Ethics · "
        "GA8 Individual and Collaborative Teamwork · GA9 Communication · GA10 Project Management and Finance · "
        "GA11 Lifelong Learning\n\n"
        "(Full WA statements as per GAPC 4.0 to be attached from the standard program booklet if required.)"
    ),
    "pos_summary": (
        "PO1 Engineering Knowledge · PO2 Problem Analysis · PO3 Design/Development of Solutions · "
        "PO4 Conduct Investigations · PO5 Engineering Tool Usage · PO6 The Engineer and The World · "
        "PO7 Ethics · PO8 Individual and Collaborative Teamwork · PO9 Communication · "
        "PO10 Project Management and Finance · PO11 Life-Long Learning\n\n"
        "(Full PO statements as per NBA to be attached from the standard program booklet if required.)"
    ),
    "peos_cse": (
        "PEO1 Core Competency: To inculcate analytical, design and implementation skills required to "
        "provide solutions to computing and interdisciplinary problems.\n"
        "PEO2 Breadth: To groom graduates to innovate, design and develop products to provide sustainable "
        "solutions to contemporary societal, environmental and business problems.\n"
        "PEO3 Professionalism: To instill the ability to work in teams, communicate effectively and lead "
        "as ethically and socially responsible professionals, researchers and entrepreneurs.\n"
        "PEO4 Life Long Learning: To foster the ability of lifelong learning to constantly adapt to "
        "emerging technologies, to pursue higher education and instill life skills for a successful career."
    ),
    "peos_me": (
        "PEO1: Graduates will demonstrate strong technical knowledge and problem-solving skills in "
        "mechanical engineering, smart manufacturing, and electric vehicle technologies.\n"
        "PEO2: Graduates will pursue successful careers in global industries or as entrepreneurs by "
        "continuously upgrading their skills through modern tools and technologies.\n"
        "PEO3: Graduates will engage in research, innovation, and higher studies by applying modern tools "
        "and interdisciplinary knowledge.\n"
        "PEO4: Graduates will practice engineering with a commitment to ethical values, sustainability, "
        "and societal well-being.\n"
        "PEO5: Graduates will exhibit effective leadership, teamwork, and communication skills in "
        "multidisciplinary environments."
    ),
    "psos_cse": (
        "PSO1: Attain the ability to design and develop computer programs and possess acquaintance with "
        "emerging technologies and open-source platforms in mobile app development, AI/ML, generative AI, "
        "web development, data analytics, cloud, networking, cyber security, quantum computing, gaming "
        "and animation.\n"
        "PSO2: Acquire technical competency to deliver computer-based innovative and effective solutions "
        "to tackle business and societal challenges, for a successful career, entrepreneurship, research "
        "and higher studies."
    ),
    "psos_me": (
        "PSO1: Apply mechanical engineering tools and sustainable practices to analyze and solve engineering "
        "problems considering technical, economic, environmental, and ethical aspects.\n"
        "PSO2: Apply automation, CAD/CAM, sensors, industrial IoT, and data analytics to monitor, control "
        "and improve manufacturing processes in line with Industry 4.0.\n"
        "PSO3 (ME-EV): Design, analyze, and implement key components of electric vehicle systems including "
        "electric drives, battery management systems, power electronics, and thermal systems."
    ),
    "ksa_summary": (
        "Knowledge Profile (WK1–WK9): Natural Sciences, Mathematics, Engineering Fundamentals, Specialist "
        "Knowledge, Engineering Design, Engineering Practice, Comprehension, Research Literature, Ethics.\n"
        "Skill Profile (SK1–SK8): Problem Solving, Design and Innovation, Use of Modern Tools, Project "
        "Management, Communication, Collaboration and Teamwork, Lifelong Learning, Research and Experimentation.\n"
        "Attitude Profile (AK1–AK8): Ethical Responsibility, Equity/Diversity/Inclusion, Sustainability "
        "Awareness, Global Awareness, Accountability, Leadership, Adaptability, Social Responsibility.\n\n"
        "(Full WK/SK/AK statements to be attached from the standard program booklet if required.)"
    ),
    "academic_calendar_note": (
        "[Attach / paste the Academic Calendar for the current session — same for all courses in this term.]"
    ),
}

PLACEHOLDER = "[To be filled by faculty]"

# =============================================================================
# SESSION STATE
# =============================================================================

def init_state():
    defaults = {
        "mode": "Minimal",
        "generated": None,
        "gen_name": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()

# =============================================================================
# TEXT / FILE HELPERS
# =============================================================================

def extract_text_from_upload(uploaded):
    """Extract plain text from PDF / DOCX / TXT upload."""
    if uploaded is None:
        return ""
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    try:
        if name.endswith(".txt"):
            return raw.decode("utf-8", errors="ignore")
        if name.endswith(".docx"):
            doc = Document(BytesIO(raw))
            parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        parts.append(" | ".join(cells))
            return "\n".join(parts)
        if name.endswith(".pdf") and pdfplumber:
            text = []
            with pdfplumber.open(BytesIO(raw)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text.append(t)
            return "\n".join(text)
    except Exception as e:
        st.warning(f"Could not fully extract text from {uploaded.name}: {e}")
    return ""


def parse_lines(text: str):
    """Split non-empty lines from a text area."""
    if not text:
        return []
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


# =============================================================================
# DOCX BUILDERS
# =============================================================================

def set_run_font(run, size=11, bold=False, italic=False):
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    except Exception:
        pass
    run.bold = bold
    run.italic = italic


def add_center(doc, text, size=16, bold=True, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, size, bold=bold)
    p.space_after = Pt(space_after)
    return p


def add_h(doc, text, size=14, bold=True):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size, bold=bold)
    p.space_after = Pt(6)
    p.space_before = Pt(10)
    return p


def add_body(doc, text, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size)
    p.space_after = Pt(4)
    return p


def add_placeholder(doc, text=PLACEHOLDER):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, 11, italic=True)
    try:
        run.font.color.rgb = RGBColor(128, 128, 128)
    except Exception:
        pass
    p.space_after = Pt(4)
    return p


def add_filled_or_placeholder(doc, text: str):
    if text and str(text).strip() and str(text).strip() != PLACEHOLDER:
        for line in str(text).strip().split("\n"):
            if line.strip():
                add_body(doc, line.strip())
    else:
        add_placeholder(doc)


def add_checklist_table(doc):
    """COURSE FILE CONTENTS checklist matching the sample structure."""
    add_h(doc, "COURSE FILE CONTENTS", 14)

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "S. No."
    hdr[1].text = "Description"
    hdr[2].text = "Y/N"

    items = [
        ("", "OUTCOME BASED CURRICULUM", True),
        ("1", "Vision and Mission of the University", False),
        ("2", "Vision and Mission of the School", False),
        ("3", "Vision and Mission of the Department, if applicable", False),
        ("4", "GA's, PEOs, POs and PSOs of the Program", False),
        ("5", "List of KSA Profiles of the Program", False),
        ("6", "Academic Calendar", False),
        ("", "OUTCOME BASED COURSE DESIGN", True),
        ("7", "Setting CO Targets", False),
        ("8", "Course Design of Theory component", False),
        ("", "   a. KSA Profile Elements of the Theory component", False),
        ("", "   b. List of SDGs/Ethics/Skill Levels & NSQF levels", False),
        ("", "   c. Course Objective and List of Course Outcomes", False),
        ("", "   d. Mapping to Skill/Employment/Entrepreneurship", False),
        ("", "   e. Relevance of the Theory component", False),
        ("", "   f. Syllabus with Text books, Reference books, Online Resources, MOOCs", False),
        ("", "   g. Lesson Plan including innovative pedagogies, ICT Tools and AI Integration", False),
        ("", "   h. CO-PO Mapping", False),
        ("", "   i. CO-KSA Mapping", False),
        ("", "   j. Course Assessment Plan", False),
        ("9", "Course Design of Lab component", False),
        ("", "   a–l. Lab design items (experiments, miniprojects, CAP, etc.)", False),
        ("", "OUTCOME BASED TEACHING & LEARNING", True),
        ("10", "List of faculty members teaching the course", False),
        ("11", "Time table of concerned faculty members", False),
        ("12", "Lecture Notes and PowerPoint handouts", False),
        ("13", "Tutorials/Assignments", False),
        ("14", "Solution of Tutorials/Assignments", False),
        ("15", "Mid-Sem Question Paper", False),
        ("16", "End-Term Question Paper", False),
        ("17", "Solution of Mid-Sem and End-Term Question Papers", False),
        ("18", "Question Papers / Continuous Assessment", False),
        ("19", "Rubrics for Lab assessments, miniprojects and case studies", False),
        ("20", "Note on how the course can contribute to the ePortfolio", False),
        ("21", "Attendance Record for theory and labs", False),
        ("", "OUTCOME-BASED ASSESSMENT & ATTAINMENT", True),
        ("22", "Continuous Evaluation for all Assessments from ERP", False),
        ("23", "Assessment-wise Marks Uploaded on ERP with CO-PO mapping", False),
        ("24", "Records of Rubrics based Evaluation", False),
        ("25", "Records of Assessments based on Student ePortfolios", False),
        ("26", "Minutes of Course Coordinator Level Meetings", False),
        ("27", "CO attainment analysis and action taken", False),
        ("28", "List of Slow / Moderate / Advanced Learners", False),
        ("29", "Action taken to support slow, moderate and advanced learners", False),
        ("30", "Impact Analysis Report and Action Taken", False),
        ("31", "Alternative Plan of Action for the Next Academic Cycle", False),
        ("", "APPENDICES", True),
        ("A", "Sample hardcopies of Tutorials and Assignments", False),
        ("B", "Sample hardcopies of answer sheets", False),
        ("C", "Evidence for pedagogy used", False),
        ("D", "Evidence for ICT Tools used", False),
        ("E", "Evidence for AI integration/AI Tools used", False),
    ]

    for num, desc, is_header in items:
        row = table.add_row().cells
        row[0].text = num
        row[1].text = desc
        row[2].text = ""
        if is_header:
            for c in row:
                for p in c.paragraphs:
                    for r in p.runs:
                        r.bold = True


def build_course_file(meta: dict, content: dict, mode: str) -> BytesIO:
    """
    meta: course identity fields
    content: optional filled sections (syllabus, cos, experiments, ...)
    mode: Minimal | Full
    """
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # ----- Cover -----
    add_center(doc, "SCHOOL OF ENGINEERING", 14)
    add_center(doc, "DEPARTMENT OF COMPUTER SCIENCE & TECHNOLOGY", 12)
    doc.add_paragraph()
    add_center(doc, "COURSE FILE", 16)
    doc.add_paragraph()
    add_center(doc, meta["course_name"].upper(), 14)
    add_center(doc, f"({meta['course_code']})", 12)
    doc.add_paragraph()
    add_body(doc, f"Program(s) and semester: {meta['program']}")
    add_body(doc, f"NCrF Level of Program: {meta.get('ncrf', 'Level 6')}")
    add_body(doc, f"Course Coordinator: {meta['coordinator']}")
    add_body(doc, f"Session: {meta['session']}")
    if meta.get("theory_faculty"):
        add_body(doc, f"Theory Faculty: {meta['theory_faculty']}")
    if meta.get("lab_faculty"):
        add_body(doc, f"Lab Faculty: {meta['lab_faculty']}")
    add_body(doc, f"L-T-P-C: {meta['ltpc']}")
    doc.add_paragraph()

    # Info table
    info = doc.add_table(rows=4, cols=2)
    info.style = "Table Grid"
    pairs = [
        ("PROGRAM & SEMESTER", meta["program"]),
        ("COURSE CODE / NAME", f"{meta['course_code']} / {meta['course_name']}"),
        ("COURSE COORDINATOR", meta["coordinator"]),
        ("LTPC / SESSION", f"{meta['ltpc']} / {meta['session']}"),
    ]
    for i, (k, v) in enumerate(pairs):
        info.rows[i].cells[0].text = k
        info.rows[i].cells[1].text = str(v)

    doc.add_paragraph()
    add_checklist_table(doc)
    doc.add_paragraph()

    # ----- OUTCOME-BASED CURRICULUM (shared defaults) -----
    add_h(doc, "OUTCOME-BASED CURRICULUM", 14)

    add_h(doc, "1. Vision and Mission of the University", 12)
    add_body(doc, "Vision of MRU")
    add_body(doc, DEFAULTS["university_vision"])
    add_body(doc, "Mission of MRU")
    add_filled_or_placeholder(doc, DEFAULTS["university_mission"])

    add_h(doc, "2. Vision and Mission of the School", 12)
    add_body(doc, "Vision of SoE")
    add_body(doc, DEFAULTS["school_vision"])
    add_body(doc, "Mission of SoE")
    add_filled_or_placeholder(doc, DEFAULTS["school_mission"])

    add_h(doc, "3. Vision and Mission of the Department", 12)
    depts = meta.get("departments") or ["CST"]
    if "CST" in depts:
        add_body(doc, "Vision of DoCST")
        add_body(doc, DEFAULTS["dept_cst_vision"])
        add_body(doc, "Mission of DoCST")
        add_filled_or_placeholder(doc, DEFAULTS["dept_cst_mission"])
    if "ME" in depts:
        add_body(doc, "Vision of DoME")
        add_body(doc, DEFAULTS["dept_me_vision"])
        add_body(doc, "Mission of DoME")
        add_filled_or_placeholder(doc, DEFAULTS["dept_me_mission"])

    add_h(doc, "4. GA's, PEOs, POs and PSOs of the Program", 12)
    add_body(doc, "Graduate Attributes (summary)")
    add_filled_or_placeholder(doc, DEFAULTS["gas_summary"])
    add_body(doc, "Program Outcomes (summary)")
    add_filled_or_placeholder(doc, DEFAULTS["pos_summary"])

    programs = meta.get("program_family") or ["CSE"]
    if "CSE" in programs or "AIML" in programs or "CDFD" in programs or "CSTI" in programs:
        add_body(doc, "PEOs — B Tech CSE & specializations")
        add_filled_or_placeholder(doc, DEFAULTS["peos_cse"])
        add_body(doc, "PSOs — B Tech CSE & specializations")
        add_filled_or_placeholder(doc, DEFAULTS["psos_cse"])
    if "ME" in programs:
        add_body(doc, "PEOs — B Tech ME")
        add_filled_or_placeholder(doc, DEFAULTS["peos_me"])
        add_body(doc, "PSOs — B Tech ME")
        add_filled_or_placeholder(doc, DEFAULTS["psos_me"])

    add_h(doc, "5. List of KSA Profiles of the Program", 12)
    add_filled_or_placeholder(doc, DEFAULTS["ksa_summary"])

    add_h(doc, "6. Academic Calendar", 12)
    add_filled_or_placeholder(doc, content.get("academic_calendar") or DEFAULTS["academic_calendar_note"])

    # ----- OUTCOME BASED COURSE DESIGN -----
    add_h(doc, "OUTCOME BASED COURSE DESIGN", 14)
    add_body(doc, f"Course Code: {meta['course_code']}    Course Title: {meta['course_name']}")
    add_body(doc, f"L-T-P-C: {meta['ltpc']}    Session: {meta['session']}")

    add_h(doc, "7. Setting CO Targets", 12)
    if content.get("co_targets"):
        add_filled_or_placeholder(doc, content["co_targets"])
    else:
        add_placeholder(doc, "[Set CO targets for this semester — e.g. CO1: 65%, CO2: 65%, ...]")

    add_h(doc, "8. Course Design of Theory component", 12)

    add_h(doc, "Course Objective", 11)
    add_filled_or_placeholder(doc, content.get("course_objective"))

    add_h(doc, "Course Outcomes (COs)", 11)
    if content.get("cos"):
        for i, co in enumerate(content["cos"], 1):
            add_body(doc, f"CO{i}: {co}")
    else:
        add_placeholder(doc, "[List Course Outcomes with Bloom's levels]")

    add_h(doc, "Syllabus / Course Contents", 11)
    add_filled_or_placeholder(doc, content.get("syllabus"))

    add_h(doc, "Textbooks / Reference Books / Online Resources / MOOCs", 11)
    add_filled_or_placeholder(doc, content.get("resources"))

    add_h(doc, "Lesson Plan (pedagogy, ICT, AI Integration)", 11)
    add_filled_or_placeholder(doc, content.get("lesson_plan"))

    add_h(doc, "CO-PO Mapping", 11)
    add_filled_or_placeholder(doc, content.get("co_po"))

    add_h(doc, "CO-KSA Mapping (SDGs, Ethics, Skill/NSQF levels)", 11)
    add_filled_or_placeholder(doc, content.get("co_ksa"))

    add_h(doc, "Course Assessment Plan (Theory)", 11)
    add_filled_or_placeholder(doc, content.get("cap_theory"))

    add_h(doc, "Mapping to Skill/Employment/Entrepreneurship", 11)
    add_filled_or_placeholder(doc, content.get("skill_mapping"))

    add_h(doc, "Relevance of the Theory component", 11)
    add_filled_or_placeholder(doc, content.get("relevance_theory"))

    # Lab
    add_h(doc, "9. Course Design of Lab component", 12)
    add_h(doc, "List of Experiments", 11)
    if content.get("experiments"):
        for i, exp in enumerate(content["experiments"], 1):
            add_body(doc, f"{i}. {exp}")
    else:
        add_placeholder(doc)

    add_h(doc, "List of Miniprojects / Case studies", 11)
    if content.get("miniprojects"):
        for i, mp in enumerate(content["miniprojects"], 1):
            add_body(doc, f"{i}. {mp}")
    else:
        add_placeholder(doc)

    add_h(doc, "Course Assessment Plan (Lab)", 11)
    add_filled_or_placeholder(doc, content.get("cap_lab"))

    add_h(doc, "Relevance of the Lab component", 11)
    add_filled_or_placeholder(doc, content.get("relevance_lab"))

    # ----- TEACHING & LEARNING -----
    add_h(doc, "OUTCOME BASED TEACHING & LEARNING", 14)
    teaching_sections = [
        ("10. List of faculty members teaching the course", content.get("faculty_list") or meta.get("theory_faculty")),
        ("11. Time table of concerned faculty members", content.get("timetable")),
        ("12. Lecture Notes and PowerPoint handouts", content.get("lecture_notes")),
        ("13. Tutorials/Assignments", content.get("tutorials")),
        ("14. Solution of Tutorials/Assignments", content.get("tutorial_solutions")),
        ("15. Mid-Sem Question Paper", content.get("midsem_qp") or ("NA" if content.get("midsem_na") else None)),
        ("16. End-Term Question Paper", content.get("endterm_qp")),
        ("17. Solution of Mid-Sem and End-Term Question Papers", content.get("qp_solutions")),
        ("18. Question Papers / Continuous Assessment", content.get("ca_papers")),
        ("19. Rubrics for Lab assessments, miniprojects and case studies", content.get("rubrics")),
        ("20. Note on how the course can contribute to the ePortfolio", content.get("eportfolio")),
        ("21. Attendance Record for theory and labs", content.get("attendance")),
    ]
    for title, val in teaching_sections:
        add_h(doc, title, 12)
        add_filled_or_placeholder(doc, val)

    # ----- ASSESSMENT & ATTAINMENT -----
    add_h(doc, "OUTCOME-BASED ASSESSMENT & ATTAINMENT", 14)
    attainment_sections = [
        ("22. Continuous Evaluation for all Assessments from ERP", content.get("continuous_eval")),
        ("23. Assessment-wise Marks Uploaded on ERP with CO-PO mapping", content.get("erp_marks")),
        ("24. Records of Rubrics based Evaluation", content.get("rubrics_records")),
        ("25. Records of Assessments based on Student ePortfolios", content.get("eportfolio_assess")),
        ("26. Minutes of Course Coordinator Level Meetings", content.get("minutes")),
        ("27. CO attainment analysis and action taken", content.get("co_attainment")),
        ("28. List of Slow / Moderate / Advanced Learners", content.get("slow_advanced")),
        ("29. Action taken to support slow, moderate and advanced learners", content.get("learner_actions")),
        ("30. Impact Analysis Report and Action Taken", content.get("impact")),
        ("31. Alternative Plan of Action for the Next Academic Cycle", content.get("alt_plan")),
    ]
    for title, val in attainment_sections:
        add_h(doc, title, 12)
        add_filled_or_placeholder(doc, val)

    # ----- APPENDICES -----
    add_h(doc, "APPENDICES", 14)
    for label, note in [
        ("A. Sample hardcopies of Tutorials and Assignments", "Attach 6 samples (2 Slow, 2 Moderate, 2 Advanced) each."),
        ("B. Sample hardcopies of answer sheets", "Attach Mid-Sem and End-Term samples (Slow / Moderate / Advanced)."),
        ("C. Evidence for pedagogy used", "Geotagged photographs / evidence per pedagogy."),
        ("D. Evidence for ICT Tools used", "Snapshots, ERP reports, simulation tools, online resource links."),
        ("E. Evidence for AI integration/AI Tools used", "Geotagged photographs / snapshots of AI tools used."),
    ]:
        add_h(doc, label, 12)
        add_placeholder(doc, note)

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run(
        f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')} | Mode: {mode} | Course File Generator"
    )
    set_run_font(run, 9, italic=True)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# =============================================================================
# UI
# =============================================================================

st.title("📚 Course File Generator")
st.caption("Minimal vs Full mode · Shared institutional defaults · Download structured .docx")

with st.sidebar:
    st.markdown("### Mode")
    mode = st.radio(
        "Generation mode",
        ["Minimal", "Full"],
        horizontal=True,
        help="Minimal = fast skeleton. Full = fill more sections from your inputs.",
    )
    st.session_state.mode = mode

    st.markdown("---")
    st.markdown("### About defaults")
    st.info(
        "University / School / Department Vision–Mission, GAs, PEOs, POs, PSOs and KSA "
        "profiles are **shared** and pre-loaded. You do not re-upload them per course."
    )
    st.markdown("---")
    st.markdown("Developed for faculty course-file workflow")

# ----- Core identity (both modes) -----
st.subheader("1. Course identity")
c1, c2, c3 = st.columns(3)
with c1:
    course_code = st.text_input("Course Code *", placeholder="e.g. CSS325B")
    course_name = st.text_input("Course Name *", placeholder="e.g. Green Computing")
    coordinator = st.text_input("Course Coordinator *", placeholder="e.g. Dr. Sachin Lakra")
with c2:
    program = st.text_input(
        "Program(s) & Semester *",
        placeholder="e.g. B Tech ME5 (SMA), B Tech CSE7/AIML7/...",
    )
    session = st.text_input("Session *", value="Jul-Dec, 2025")
    ltpc = st.text_input("L-T-P-C *", placeholder="e.g. 1-0-2-2")
with c3:
    theory_faculty = st.text_input("Theory Faculty", placeholder="Same as coordinator if blank")
    lab_faculty = st.text_input("Lab Faculty", placeholder="Same as theory if blank")
    ncrf = st.text_input("NCrF Level", value="Level 6")

d1, d2 = st.columns(2)
with d1:
    departments = st.multiselect(
        "Department(s) for Vision/Mission",
        ["CST", "ME"],
        default=["CST"],
        help="Select both for dual-program courses (e.g. CSE + ME).",
    )
with d2:
    program_family = st.multiselect(
        "Program family (for PEOs / PSOs)",
        ["CSE", "AIML", "CDFD", "CSTI", "ME"],
        default=["CSE"],
        help="Controls which PEOs/PSOs blocks are inserted.",
    )

# ----- Mode-specific inputs -----
content = {}

st.subheader("2. Course content")

if mode == "Minimal":
    st.markdown("**Minimal mode** — optional syllabus and COs only. Everything else uses defaults or placeholders.")
    syl = st.file_uploader("Upload syllabus (PDF / DOCX / TXT)", type=["pdf", "docx", "txt"], key="syl_min")
    course_objective = st.text_area("Course objective (optional)", height=80)
    cos_text = st.text_area(
        "Course Outcomes — one per line (optional)",
        placeholder="Explain the concept of carbon footprint...\nApply green computing techniques...",
        height=120,
    )
    content["syllabus"] = extract_text_from_upload(syl) if syl else ""
    content["course_objective"] = course_objective
    content["cos"] = parse_lines(cos_text)

else:
    st.markdown("**Full mode** — add as many course-specific sections as you have. Missing ones stay as placeholders.")

    with st.expander("A. Syllabus, objective, COs & targets", expanded=True):
        syl = st.file_uploader("Syllabus (PDF / DOCX / TXT)", type=["pdf", "docx", "txt"], key="syl_full")
        content["syllabus"] = extract_text_from_upload(syl) if syl else ""
        content["course_objective"] = st.text_area("Course objective", height=80)
        cos_text = st.text_area("Course Outcomes — one per line", height=120)
        content["cos"] = parse_lines(cos_text)
        content["co_targets"] = st.text_area(
            "CO Targets (e.g. CO1: 65%, CO2: 65%)",
            height=80,
            placeholder="CSS325B.1 → 65%\nCSS325B.2 → 65%",
        )
        content["resources"] = st.text_area("Textbooks / References / MOOCs", height=80)

    with st.expander("B. Mappings, CAP, Lesson plan, Relevance"):
        content["co_po"] = st.text_area("CO–PO Mapping (paste table or summary)", height=100)
        content["co_ksa"] = st.text_area("CO–KSA Mapping (SDGs, Ethics, Skill/NSQF)", height=100)
        content["cap_theory"] = st.text_area("Course Assessment Plan — Theory", height=80)
        content["cap_lab"] = st.text_area("Course Assessment Plan — Lab", height=80)
        content["lesson_plan"] = st.text_area(
            "Lesson Plan (pedagogy, ICT, AI integration)",
            height=100,
        )
        content["skill_mapping"] = st.text_area("Mapping to Skill / Employment / Entrepreneurship", height=60)
        content["relevance_theory"] = st.text_area("Relevance of Theory component", height=60)
        content["relevance_lab"] = st.text_area("Relevance of Lab component", height=60)

    with st.expander("C. Lab — Experiments & Miniprojects"):
        exp_text = st.text_area("List of Experiments — one per line", height=120)
        mp_text = st.text_area("List of Miniprojects / Case topics — one per line", height=80)
        content["experiments"] = parse_lines(exp_text)
        content["miniprojects"] = parse_lines(mp_text)

    with st.expander("D. Teaching & learning (optional)"):
        content["faculty_list"] = st.text_input("Faculty list (if different from identity)")
        content["timetable"] = st.text_area("Timetable note / paste", height=60)
        content["eportfolio"] = st.text_area("ePortfolio contribution note", height=60)
        content["midsem_na"] = st.checkbox("Mid-Sem Question Paper is NA", value=False)
        content["rubrics"] = st.text_area("Rubrics summary (or paste)", height=60)

    with st.expander("E. Assessment & attainment (optional — often filled later)"):
        content["co_attainment"] = st.text_area("CO attainment summary", height=60)
        content["slow_advanced"] = st.text_area("Slow / Advanced learners summary", height=60)
        content["impact"] = st.text_area("Impact analysis summary", height=60)
        content["alt_plan"] = st.text_area("Alternative plan for next cycle", height=60)

# ----- Completeness hint -----
filled = sum(
    1
    for k, v in content.items()
    if v and v is not False and (not isinstance(v, list) or len(v) > 0)
)
st.caption(f"Mode: **{mode}** · Optional content fields with data: **{filled}**")

# ----- Generate -----
st.subheader("3. Generate")
if st.button("Generate Course File", type="primary", use_container_width=True):
    missing = []
    if not course_code.strip():
        missing.append("Course Code")
    if not course_name.strip():
        missing.append("Course Name")
    if not program.strip():
        missing.append("Program(s) & Semester")
    if not session.strip():
        missing.append("Session")
    if not ltpc.strip():
        missing.append("L-T-P-C")
    if not coordinator.strip():
        missing.append("Course Coordinator")

    if missing:
        st.error("Please fill required fields: " + ", ".join(missing))
    else:
        meta = {
            "course_code": course_code.strip(),
            "course_name": course_name.strip(),
            "program": program.strip(),
            "session": session.strip(),
            "ltpc": ltpc.strip(),
            "coordinator": coordinator.strip(),
            "theory_faculty": (theory_faculty or coordinator).strip(),
            "lab_faculty": (lab_faculty or theory_faculty or coordinator).strip(),
            "ncrf": ncrf.strip() or "Level 6",
            "departments": departments or ["CST"],
            "program_family": program_family or ["CSE"],
        }
        try:
            with st.spinner("Building course file..."):
                buf = build_course_file(meta, content, mode)
                safe = re.sub(r"[^\w\s-]", "", course_name).strip().replace(" ", "_")
                fname = f"CourseFile_{course_code.strip()}_{safe}_{mode}.docx"
                st.session_state.generated = buf.getvalue()
                st.session_state.gen_name = fname
            st.success(f"Course file generated ({mode} mode).")
        except Exception as e:
            st.error("Generation failed.")
            st.code(str(e))
            with st.expander("Details"):
                st.code(traceback.format_exc())

if st.session_state.generated:
    st.download_button(
        label=f"Download {st.session_state.gen_name}",
        data=st.session_state.generated,
        file_name=st.session_state.gen_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    if mode == "Minimal":
        st.info(
            "Minimal file includes: cover, checklist, institutional sections (1–6), "
            "and any syllabus/COs you entered. Remaining sections are placeholders."
        )
    else:
        st.info(
            "Full file merges every section you filled. Empty optional sections remain placeholders "
            "for later (ERP marks, minutes, appendix evidence, etc.)."
        )

st.markdown("---")
st.caption("Course File Generator · Shared institutional defaults · Minimal / Full modes")
