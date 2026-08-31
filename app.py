import os
import time
import uuid
from datetime import datetime

from flask import (
    Flask,
    request,
    jsonify,
    render_template,
    send_from_directory,
)

from google import genai
from google.genai import types

from docx import Document
from pypdf import PdfReader
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
)
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)
from reportlab.lib.enums import TA_CENTER

from dotenv import load_dotenv


# ============================================================
# ENVIRONMENT
# ============================================================

load_dotenv()


# ============================================================
# FLASK APP
# ============================================================

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Vercel allows writing inside /tmp
GENERATED_DIR = os.path.join("/tmp", "qraft_generated")
os.makedirs(GENERATED_DIR, exist_ok=True)


# ============================================================
# GEMINI SETTINGS
# ============================================================

# Fast model.
# You can change this from Vercel Environment Variables using
# GEMINI_MODEL if required.
MODEL = os.environ.get(
    "GEMINI_MODEL",
    "gemini-3.5-flash-lite"
)


def get_gemini_client():
    """
    Create Gemini client only when generation is requested.
    This prevents the whole Flask app from crashing if the
    environment variable is missing.
    """

    api_key = os.environ.get("GEMINI_API_KEY")

    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY is not configured."
        )

    return genai.Client(
        api_key=api_key
    )


# ============================================================
# FAST GENERATION CONFIG
# ============================================================

# IMPORTANT:
# Do NOT use thinking_budget here.
# Your installed google-genai version rejected it.

FAST_CONFIG = types.GenerateContentConfig(
    max_output_tokens=2500,
)


# ============================================================
# K-LEVEL DEFINITIONS
# ============================================================

K_LEVEL_DEFS = {

    "K1":
        "Remember - recall facts, terms, basic concepts "
        "(define, list, state, name)",

    "K2":
        "Understand - explain ideas or concepts "
        "(explain, describe, summarize, classify)",

    "K3":
        "Apply - use information in new situations "
        "(solve, demonstrate, apply, calculate)",

    "K4":
        "Analyze - draw connections among ideas "
        "(compare, differentiate, examine, categorize)",

    "K5":
        "Evaluate - justify a stand or decision "
        "(justify, critique, assess, argue)",

    "K6":
        "Create - produce new or original work "
        "(design, develop, formulate, propose)",
}


# ============================================================
# EXAM DEFAULTS
# ============================================================

EXAM_DEFAULTS = {

    "CIA 1": {
        "start": "K1",
        "end": "K3",
        "total_questions": 10,
        "marks": 5,
    },

    "CIA 2": {
        "start": "K1",
        "end": "K4",
        "total_questions": 10,
        "marks": 5,
    },

    "Model Exam": {
        "start": "K1",
        "end": "K5",
        "total_questions": 20,
        "marks": 5,
    },

    "Semester Exam": {
        "start": "K1",
        "end": "K6",
        "total_questions": 20,
        "marks": 5,
    },
}


# ============================================================
# QUESTION DISTRIBUTION
# ============================================================

def distribute_questions(
    start_level,
    end_level,
    total_questions
):

    levels_order = list(
        K_LEVEL_DEFS.keys()
    )

    start_idx = levels_order.index(
        start_level
    )

    end_idx = levels_order.index(
        end_level
    )

    selected_levels = levels_order[
        start_idx:end_idx + 1
    ]

    number_of_levels = len(
        selected_levels
    )

    base = (
        total_questions //
        number_of_levels
    )

    remainder = (
        total_questions %
        number_of_levels
    )

    distribution = {
        level: 0
        for level in levels_order
    }

    for i, level in enumerate(
        selected_levels
    ):

        distribution[level] = (
            base +
            (
                1
                if i < remainder
                else 0
            )
        )

    return distribution


# ============================================================
# SECTION DISTRIBUTION
# ============================================================

def build_section_distribution(
    start_level,
    end_level,
    total_questions
):

    levels_order = list(
        K_LEVEL_DEFS.keys()
    )

    start_idx = levels_order.index(
        start_level
    )

    end_idx = levels_order.index(
        end_level
    )

    selected_levels = levels_order[
        start_idx:end_idx + 1
    ]

    number_of_levels = len(
        selected_levels
    )

    base = (
        total_questions //
        number_of_levels
    )

    remainder = (
        total_questions %
        number_of_levels
    )

    distribution = {}

    for i, level in enumerate(
        selected_levels
    ):

        distribution[level] = (
            base +
            (
                1
                if i < remainder
                else 0
            )
        )

    return distribution


# ============================================================
# NORMAL QUESTION PAPER PROMPT
# ============================================================

def build_prompt(
    core_content,
    subject,
    k_distribution,
    marks_per_question,
    exam_title
):

    level_lines = []

    total_questions = 0

    for level, count in (
        k_distribution.items()
    ):

        if count > 0:

            level_lines.append(
                f"{level} - "
                f"{K_LEVEL_DEFS[level]}: "
                f"{count} question(s)"
            )

            total_questions += count

    total_marks = (
        total_questions *
        marks_per_question
    )

    prompt = f"""
You are an expert university examination
question paper setter.

Create a formal question paper.

IMPORTANT:
Use ONLY the CORE CONTENT provided below.
Do not invent topics outside the CORE CONTENT.

EXAM TITLE:
{exam_title}

SUBJECT:
{subject}

CORE CONTENT:
{core_content}

K-LEVEL DISTRIBUTION:
{chr(10).join(level_lines)}

REQUIREMENTS:

1. Generate exactly {total_questions} questions.

2. Every question carries exactly
{marks_per_question} marks.

3. Total marks must be {total_marks}.

4. Every question must have a K-Level tag.

5. Number questions continuously.

6. Group questions according to K-Level.

7. Do not repeat questions.

8. Questions must be answerable using
only the CORE CONTENT.

9. Use clear university examination language.

10. Do not add explanations.

OUTPUT FORMAT:

Section A - K1 (Remember)

1. Question text [K1] ({marks_per_question} marks)

Section B - K2 (Understand)

2. Question text [K2] ({marks_per_question} marks)

Continue only for K-levels that contain
questions.

At the end write:

TOTAL MARKS: {total_marks}

Output ONLY the question paper.
"""

    return (
        prompt,
        total_marks,
        total_questions
    )


# ============================================================
# CUSTOM PATTERN PROMPT
# ============================================================

def build_pattern_prompt(
    core_content,
    subject,
    pattern,
    exam_title
):

    total_marks = 0
    total_questions = 0

    pattern_details = []

    for section in pattern:

        section_name = (
            section.get("section")
            or "Part"
        )

        question_count = int(
            section.get(
                "question_count",
                0
            )
        )

        marks = int(
            section.get(
                "marks",
                0
            )
        )

        has_choice = bool(
            section.get(
                "has_choice",
                False
            )
        )

        k_levels = section.get(
            "k_levels",
            {}
        )

        total_questions += (
            question_count
        )

        total_marks += (
            question_count *
            marks
        )

        k_level_text = []

        for level in [
            "K1",
            "K2",
            "K3",
            "K4",
            "K5",
            "K6",
        ]:

            count = int(
                k_levels.get(
                    level,
                    0
                )
            )

            if count > 0:

                k_level_text.append(
                    f"{level}: "
                    f"{count} question(s)"
                )

        if has_choice:

            choice_text = (
                "YES - Every question "
                "must have A OR B choice."
            )

        else:

            choice_text = "NO"

        pattern_details.append(
            f"""
SECTION:
{section_name}

QUESTIONS:
{question_count}

MARKS:
{marks}

INTERNAL CHOICE:
{choice_text}

K-LEVEL:
{chr(10).join(k_level_text)}
"""
        )

    prompt = f"""
You are an expert university examination
question paper setter.

Create ONE COMPLETE QUESTION PAPER.

EXAM TITLE:
{exam_title}

SUBJECT:
{subject}

SYLLABUS:
{core_content}

QUESTION PAPER PATTERN:

{chr(10).join(pattern_details)}

STRICT RULES:

1. Follow the exact pattern.

2. Generate the exact number of questions.

3. Follow the exact marks.

4. Follow the specified K-Level distribution.

5. Add K-Level to every question.

6. Do not repeat questions.

7. Generate questions ONLY from the syllabus.

8. Number questions continuously.

9. If INTERNAL CHOICE is YES:

(A) Question [K-Level] (marks)

OR

(B) Alternative Question [K-Level] (marks)

10. A and B must have similar difficulty.

11. Use clear university examination language.

12. Output ONLY the question paper.

At the end write:

TOTAL QUESTIONS: {total_questions}

TOTAL MARKS: {total_marks}
"""

    return (
        prompt,
        total_marks,
        total_questions
    )


# ============================================================
# GEMINI GENERATION
# ============================================================

def generate_with_gemini(prompt):

    last_error = None

    # Only 3 attempts.
    # This prevents the website from hanging forever.
    for attempt in range(1, 4):

        try:

            client = get_gemini_client()

            response = (
                client.models.generate_content(
                    model=MODEL,
                    contents=prompt,
                    config=FAST_CONFIG,
                )
            )

            text = (
                response.text or ""
            ).strip()

            if text:

                return (
                    text,
                    None
                )

            return (
                "",
                "AI returned an empty response."
            )

        except Exception as e:

            last_error = e

            error_str = (
                str(e).lower()
            )

            temporary_error = (
                "503" in error_str
                or
                "unavailable" in error_str
                or
                "overloaded" in error_str
                or
                "429" in error_str
                or
                "rate limit" in error_str
                or
                "resource exhausted" in error_str
            )

            if (
                temporary_error
                and attempt < 3
            ):

                # Short delay
                time.sleep(1)

                continue

            break

    return (
        None,
        str(last_error)
    )


# ============================================================
# HOME
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html",
        k_defs=K_LEVEL_DEFS,
        exam_defaults=EXAM_DEFAULTS,
    )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():

    return jsonify({
        "status": "ok",
        "service": "Qraft AI",
        "model": MODEL,
    })


# ============================================================
# UPLOAD SYLLABUS
# ============================================================

@app.route(
    "/api/upload",
    methods=["POST"]
)
def upload_syllabus():

    if "file" not in request.files:

        return jsonify({
            "error":
            "No file uploaded"
        }), 400

    file = request.files["file"]

    filename = (
        file.filename or ""
    ).lower()

    try:

        if filename.endswith(".pdf"):

            reader = PdfReader(file)

            text = "\n".join(
                page.extract_text() or ""
                for page in reader.pages
            )

        elif filename.endswith(".docx"):

            doc = Document(file)

            text = "\n".join(
                p.text
                for p in doc.paragraphs
            )

        elif filename.endswith(".txt"):

            text = (
                file.read()
                .decode(
                    "utf-8",
                    errors="ignore"
                )
            )

        else:

            return jsonify({
                "error":
                "Only PDF, DOCX or TXT files are supported"
            }), 400

        text = text.strip()

        if not text:

            return jsonify({
                "error":
                "Could not read any text from this file"
            }), 400

        return jsonify({
            "text": text
        })

    except Exception as e:

        return jsonify({
            "error":
            f"Could not read this file: {str(e)}"
        }), 500


# ============================================================
# NORMAL GENERATION
# ============================================================

@app.route(
    "/api/generate",
    methods=["POST"]
)
def generate_question_paper():

    try:

        data = request.get_json(
            force=True
        )

    except Exception:

        return jsonify({
            "error":
            "Invalid JSON request"
        }), 400

    core_content = (
        data.get("core_content")
        or ""
    ).strip()

    subject = (
        data.get("subject")
        or "General"
    ).strip()

    exam_title = (
        data.get("exam_title")
        or "Internal Assessment Test"
    ).strip()

    try:

        marks_per_question = int(
            data.get(
                "marks_per_question",
                5
            )
        )

        total_questions = int(
            data.get(
                "total_questions",
                0
            )
        )

    except Exception:

        return jsonify({
            "error":
            "Marks and question count must be numbers"
        }), 400

    start_level = (
        data.get("start_level")
        or "K1"
    )

    end_level = (
        data.get("end_level")
        or "K4"
    )

    file_format = (
        data.get("format")
        or "docx"
    ).lower()

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    if not core_content:

        return jsonify({
            "error":
            "core_content (topic/notes) is required"
        }), 400

    if total_questions <= 0:

        return jsonify({
            "error":
            "total_questions must be at least 1"
        }), 400

    if marks_per_question <= 0:

        return jsonify({
            "error":
            "marks_per_question must be greater than 0"
        }), 400

    levels_order = list(
        K_LEVEL_DEFS.keys()
    )

    if (
        start_level not in levels_order
        or
        end_level not in levels_order
    ):

        return jsonify({
            "error":
            "Invalid K-level range"
        }), 400

    if (
        levels_order.index(start_level)
        >
        levels_order.index(end_level)
    ):

        return jsonify({
            "error":
            "Start level must come before end level"
        }), 400

    if file_format not in (
        "docx",
        "pdf"
    ):

        file_format = "docx"

    # --------------------------------------------------------
    # DISTRIBUTION
    # --------------------------------------------------------

    k_distribution = (
        distribute_questions(
            start_level,
            end_level,
            total_questions
        )
    )

    prompt, total_marks, total_q = (
        build_prompt(
            core_content,
            subject,
            k_distribution,
            marks_per_question,
            exam_title
        )
    )

    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------

    paper_text, error = (
        generate_with_gemini(
            prompt
        )
    )

    if error:

        return jsonify({
            "error":
            f"Generation failed: {error}"
        }), 500

    if not paper_text:

        return jsonify({
            "error":
            "AI returned an empty response"
        }), 500

    # --------------------------------------------------------
    # FILE CREATION
    # --------------------------------------------------------

    filename_base = (
        "question_paper_"
        +
        uuid.uuid4().hex[:8]
    )

    try:

        if file_format == "pdf":

            filename = (
                filename_base
                + ".pdf"
            )

            filepath = os.path.join(
                GENERATED_DIR,
                filename
            )

            save_as_pdf(
                paper_text,
                exam_title,
                subject,
                total_marks,
                filepath
            )

        else:

            filename = (
                filename_base
                + ".docx"
            )

            filepath = os.path.join(
                GENERATED_DIR,
                filename
            )

            save_as_docx(
                paper_text,
                exam_title,
                subject,
                total_marks,
                filepath
            )

    except Exception as e:

        return jsonify({
            "error":
            f"File creation failed: {str(e)}"
        }), 500

    return jsonify({

        "paper_text":
        paper_text,

        "total_marks":
        total_marks,

        "total_questions":
        total_q,

        "k_distribution":
        k_distribution,

        "download_url":
        f"/download/{filename}"

    })


# ============================================================
# CUSTOM PATTERN GENERATION
# ============================================================

@app.route(
    "/api/generate-pattern",
    methods=["POST"]
)
def generate_pattern_question_paper():

    try:

        data = request.get_json(
            force=True
        )

    except Exception:

        return jsonify({
            "error":
            "Invalid JSON request"
        }), 400

    core_content = (
        data.get("core_content")
        or ""
    ).strip()

    subject = (
        data.get("subject")
        or "General"
    ).strip()

    exam_title = (
        data.get("exam_title")
        or "Custom Question Paper"
    ).strip()

    pattern = (
        data.get("pattern")
        or []
    )

    file_format = (
        data.get("format")
        or "docx"
    ).lower()

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    if not core_content:

        return jsonify({
            "error":
            "Please provide syllabus/content"
        }), 400

    if not pattern:

        return jsonify({
            "error":
            "Please add at least one question pattern"
        }), 400

    for section in pattern:

        try:

            question_count = int(
                section.get(
                    "question_count",
                    0
                )
            )

            marks = int(
                section.get(
                    "marks",
                    0
                )
            )

        except Exception:

            return jsonify({
                "error":
                "Question count and marks must be numbers"
            }), 400

        if question_count <= 0:

            return jsonify({
                "error":
                f"{section.get('section', 'Section')} "
                "must have at least 1 question"
            }), 400

        if marks <= 0:

            return jsonify({
                "error":
                f"{section.get('section', 'Section')} "
                "marks must be greater than 0"
            }), 400

    if file_format not in (
        "pdf",
        "docx"
    ):

        file_format = "docx"

    # --------------------------------------------------------
    # BUILD PROMPT
    # --------------------------------------------------------

    (
        prompt,
        total_marks,
        total_questions
    ) = build_pattern_prompt(
        core_content,
        subject,
        pattern,
        exam_title
    )

    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------

    paper_text, error = (
        generate_with_gemini(
            prompt
        )
    )

    if error:

        return jsonify({
            "error":
            f"Pattern generation failed: {error}"
        }), 500

    if not paper_text:

        return jsonify({
            "error":
            "AI returned an empty response"
        }), 500

    # --------------------------------------------------------
    # CREATE FILE
    # --------------------------------------------------------

    filename_base = (
        "pattern_question_paper_"
        +
        uuid.uuid4().hex[:8]
    )

    try:

        if file_format == "pdf":

            filename = (
                filename_base
                + ".pdf"
            )

            filepath = os.path.join(
                GENERATED_DIR,
                filename
            )

            save_as_pdf(
                paper_text,
                exam_title,
                subject,
                total_marks,
                filepath
            )

        else:

            filename = (
                filename_base
                + ".docx"
            )

            filepath = os.path.join(
                GENERATED_DIR,
                filename
            )

            save_as_docx(
                paper_text,
                exam_title,
                subject,
                total_marks,
                filepath
            )

    except Exception as e:

        return jsonify({
            "error":
            f"File creation failed: {str(e)}"
        }), 500

    return jsonify({

        "paper_text":
        paper_text,

        "total_marks":
        total_marks,

        "total_questions":
        total_questions,

        "download_url":
        f"/download/{filename}"

    })


# ============================================================
# DOWNLOAD
# ============================================================

@app.route(
    "/download/<path:filename>"
)
def download(filename):

    return send_from_directory(
        GENERATED_DIR,
        filename,
        as_attachment=True
    )


# ============================================================
# DOCX WRITER
# ============================================================

def save_as_docx(
    paper_text,
    exam_title,
    subject,
    total_marks,
    filepath
):

    doc = Document()

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    title = doc.add_paragraph()

    title.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = title.add_run(
        exam_title
    )

    run.bold = True
    run.font.size = Pt(16)

    # --------------------------------------------------------
    # SUBTITLE
    # --------------------------------------------------------

    sub = doc.add_paragraph()

    sub.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    sub_run = sub.add_run(

        f"Subject: {subject}    |    "
        f"Total Marks: {total_marks}    |    "
        f"Date: "
        f"{datetime.now().strftime('%d-%m-%Y')}"

    )

    sub_run.font.size = Pt(11)

    doc.add_paragraph()

    # --------------------------------------------------------
    # PAPER CONTENT
    # --------------------------------------------------------

    for line in paper_text.split("\n"):

        line = line.strip()

        if not line:

            doc.add_paragraph()

            continue

        lower_line = (
            line.lower()
        )

        upper_line = (
            line.upper()
        )

        if (
            lower_line.startswith(
                "section"
            )
            or
            lower_line.startswith(
                "part"
            )
        ):

            heading = (
                doc.add_paragraph()
            )

            heading_run = (
                heading.add_run(line)
            )

            heading_run.bold = True
            heading_run.font.size = Pt(13)

        elif (
            upper_line.startswith(
                "TOTAL MARKS"
            )
            or
            upper_line.startswith(
                "TOTAL QUESTIONS"
            )
        ):

            paragraph = (
                doc.add_paragraph()
            )

            bold_run = (
                paragraph.add_run(line)
            )

            bold_run.bold = True

        else:

            doc.add_paragraph(line)

    doc.save(filepath)


# ============================================================
# PDF WRITER
# ============================================================

def save_as_pdf(
    paper_text,
    exam_title,
    subject,
    total_marks,
    filepath
):

    doc = SimpleDocTemplate(

        filepath,

        pagesize=A4,

        topMargin=40,

        bottomMargin=40,

        leftMargin=40,

        rightMargin=40,
    )

    styles = (
        getSampleStyleSheet()
    )

    title_style = ParagraphStyle(

        "TitleCenter",

        parent=styles["Title"],

        alignment=TA_CENTER,

        fontSize=16,
    )

    sub_style = ParagraphStyle(

        "SubCenter",

        parent=styles["Normal"],

        alignment=TA_CENTER,

        fontSize=10.5,
    )

    section_style = ParagraphStyle(

        "Section",

        parent=styles["Heading2"],

        fontSize=12.5,

        spaceBefore=12,

        spaceAfter=8,
    )

    body_style = ParagraphStyle(

        "Body",

        parent=styles["Normal"],

        fontSize=10.5,

        spaceAfter=6,

        leading=15,
    )

    story = [

        Paragraph(
            exam_title,
            title_style
        ),

        Paragraph(

            f"Subject: {subject} "
            f"&nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Total Marks: {total_marks} "
            f"&nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Date: "
            f"{datetime.now().strftime('%d-%m-%Y')}",

            sub_style
        ),

        Spacer(1, 16),

    ]

    # --------------------------------------------------------
    # PAPER CONTENT
    # --------------------------------------------------------

    for line in paper_text.split("\n"):

        line = line.strip()

        if not line:

            story.append(
                Spacer(1, 6)
            )

            continue

        # Escape ReportLab XML
        safe_line = (
            line
            .replace(
                "&",
                "&amp;"
            )
            .replace(
                "<",
                "&lt;"
            )
            .replace(
                ">",
                "&gt;"
            )
        )

        lower_line = (
            line.lower()
        )

        upper_line = (
            line.upper()
        )

        if (
            lower_line.startswith(
                "section"
            )
            or
            lower_line.startswith(
                "part"
            )
        ):

            story.append(

                Paragraph(
                    safe_line,
                    section_style
                )

            )

        elif (
            upper_line.startswith(
                "TOTAL MARKS"
            )
            or
            upper_line.startswith(
                "TOTAL QUESTIONS"
            )
        ):

            story.append(

                Paragraph(
                    f"<b>{safe_line}</b>",
                    body_style
                )

            )

        else:

            story.append(

                Paragraph(
                    safe_line,
                    body_style
                )

            )

    doc.build(story)


# ============================================================
# LOCAL SERVER
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )