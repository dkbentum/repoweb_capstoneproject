import io
import os
import re
import shutil
import time
import uuid
import zipfile
from pathlib import Path

import pandas as pd
from flask import Flask, after_this_request, jsonify, request, send_file, send_from_directory
from PIL import Image
from pypdf import PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
default_generated_dir = "/tmp/generated" if os.environ.get("VERCEL") else str(BASE_DIR / "generated")
GENERATED_DIR = Path(os.environ.get("GENERATED_DIR", default_generated_dir))
GENERATED_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_PRIMARY = "#1638B7"
TEMPLATE_LABELS = {
    "classic": "Classic Board",
    "modern": "Modern Split",
    "minimal": "Minimal Ledger",
}
LOGO_STYLE_LABELS = {
    "header_badge": "Header Badge",
    "watermark": "Watermark",
    "corner_stamp": "Corner Stamp",
    "side_mark": "Side Mark",
}

app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/static")


def normalize_column(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(name).strip().lower()).strip("_")


def safe_stem(filename: str) -> str:
    stem = Path(filename).stem.strip()
    stem = re.sub(r"[_-]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem)
    return stem.title() or "Subject"


def coerce_float(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def pick_column(column_map, options):
    for option in options:
        if option in column_map:
            return column_map[option]
    return None


def read_marks_dataframe(file_storage):
    suffix = Path(file_storage.filename).suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(file_storage)
    else:
        df = pd.read_excel(file_storage)
    return df.dropna(how="all")


def normalize_hex_color(value, fallback=DEFAULT_PRIMARY):
    text = str(value or "").strip()
    if re.fullmatch(r"#?[0-9a-fA-F]{6}", text):
        if not text.startswith("#"):
            text = f"#{text}"
        return text.upper()
    return fallback


def hex_to_rgb_tuple(hex_color):
    color = normalize_hex_color(hex_color)
    return tuple(int(color[index : index + 2], 16) for index in (1, 3, 5))


def rgb_tuple_to_hex(rgb):
    r, g, b = (max(0, min(255, int(v))) for v in rgb)
    return f"#{r:02X}{g:02X}{b:02X}"


def blend_colors(hex_a, hex_b, factor):
    factor = max(0.0, min(1.0, float(factor)))
    a = hex_to_rgb_tuple(hex_a)
    b = hex_to_rgb_tuple(hex_b)
    mixed = tuple(round((1.0 - factor) * a[index] + factor * b[index]) for index in range(3))
    return rgb_tuple_to_hex(mixed)


def lighten(hex_color, factor):
    return blend_colors(hex_color, "#FFFFFF", factor)


def darken(hex_color, factor):
    return blend_colors(hex_color, "#000000", factor)


def ordinal(n: int) -> str:
    if 10 <= (n % 100) <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def assign_competition_positions(scored_rows):
    sorted_rows = sorted(scored_rows, key=lambda row: (-row[1], row[2].lower()))
    positions = {}
    last_score = None
    last_rank = 0

    for index, (student_key, score, _) in enumerate(sorted_rows, start=1):
        if last_score is None or abs(score - last_score) > 1e-9:
            last_rank = index
            last_score = score
        positions[student_key] = last_rank

    return positions


def grade_from_score(score: float) -> str:
    if score >= 75:
        return "A"
    if score >= 65:
        return "B"
    if score >= 55:
        return "C"
    if score >= 45:
        return "D"
    return "F"


def remark_from_average(avg: float) -> str:
    if avg >= 75:
        return "Excellent performance"
    if avg >= 65:
        return "Very good performance"
    if avg >= 55:
        return "Good performance"
    if avg >= 45:
        return "Fair performance"
    return "Needs improvement"


def register_fonts():
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "DejaVuSans.ttf"))
        return "DejaVu"
    except Exception:
        return "Helvetica"


def _resample_lanczos():
    if hasattr(Image, "Resampling"):
        return Image.Resampling.LANCZOS
    return Image.LANCZOS


def load_logo_assets(file_storage, watermark_opacity_factor):
    if not file_storage or not file_storage.filename:
        return None, None

    try:
        raw_bytes = file_storage.read()
        if not raw_bytes:
            return None, "logo file is empty"

        image = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
        if image.width == 0 or image.height == 0:
            return None, "logo image has invalid dimensions"

        max_side = 1200
        if max(image.size) > max_side:
            image.thumbnail((max_side, max_side), _resample_lanczos())

        watermark = image.copy()
        alpha_channel = watermark.getchannel("A")
        strength = max(0.05, min(0.45, float(watermark_opacity_factor)))
        alpha_channel = alpha_channel.point(lambda alpha: int(alpha * strength))
        watermark.putalpha(alpha_channel)

        return (
            {
                "normal": ImageReader(image),
                "watermark": ImageReader(watermark),
                "ratio": image.width / image.height,
            },
            None,
        )
    except Exception as exc:
        return None, f"could not process logo ({exc})"


def _fit_by_height(ratio, target_h, max_w):
    ratio = max(0.2, min(5.0, float(ratio)))
    width = target_h * ratio
    height = target_h
    if width > max_w:
        width = max_w
        height = width / ratio
    return width, height


def draw_logo(c, school, width, height, area):
    logo_assets = school.get("logo_assets")
    if not logo_assets:
        return

    style = school.get("logo_style", "header_badge")
    ratio = logo_assets.get("ratio", 1.0)

    if style == "watermark":
        reader = logo_assets["watermark"]
        target_h = 80 * mm if area != "modern" else 70 * mm
        max_w = 120 * mm
        logo_w, logo_h = _fit_by_height(ratio, target_h, max_w)
        x = (width - logo_w) / 2
        y = (height - logo_h) / 2
        c.drawImage(reader, x, y, width=logo_w, height=logo_h, preserveAspectRatio=True, mask="auto")
        return

    reader = logo_assets["normal"]
    if style == "side_mark":
        target_h = 24 * mm if area == "modern" else 28 * mm
        max_w = 35 * mm
        logo_w, logo_h = _fit_by_height(ratio, target_h, max_w)
        if area == "modern":
            x = 10 * mm
            y = height * 0.4
        else:
            x = 9 * mm
            y = (height - logo_h) / 2
    elif style == "corner_stamp":
        target_h = 20 * mm
        max_w = 42 * mm
        logo_w, logo_h = _fit_by_height(ratio, target_h, max_w)
        x = width - logo_w - 18 * mm
        y = 24 * mm
    else:
        target_h = 24 * mm
        max_w = 46 * mm
        logo_w, logo_h = _fit_by_height(ratio, target_h, max_w)
        x = width - logo_w - 18 * mm
        if area == "modern":
            y = height - 42 * mm
        elif area == "minimal":
            y = height - 34 * mm
        else:
            y = height - 49 * mm

    x = max(6 * mm, min(x, width - logo_w - 6 * mm))
    y = max(14 * mm, min(y, height - logo_h - 14 * mm))

    border_color = colors.HexColor(lighten(school["theme_primary"], 0.58))
    frame_padding = 2.1 * mm

    c.setFillColor(colors.white)
    c.roundRect(
        x - frame_padding,
        y - frame_padding,
        logo_w + 2 * frame_padding,
        logo_h + 2 * frame_padding,
        3 * mm,
        stroke=0,
        fill=1,
    )

    c.setStrokeColor(border_color)
    c.setLineWidth(0.75)
    c.roundRect(
        x - frame_padding,
        y - frame_padding,
        logo_w + 2 * frame_padding,
        logo_h + 2 * frame_padding,
        3 * mm,
        stroke=1,
        fill=0,
    )

    c.drawImage(reader, x, y, width=logo_w, height=logo_h, preserveAspectRatio=True, mask="auto")


def build_subject_rows(student):
    rows = [["Subject", "ClassScore", "ExamScore", "Final", "Position", "Grade"]]
    for subject, details in sorted(student["subjects"].items()):
        rows.append(
            [
                subject,
                f"{details['class_score']:.2f}",
                f"{details['exam_score']:.2f}",
                f"{details['final_score']:.2f}",
                ordinal(details["subject_position"]),
                grade_from_score(details["final_score"]),
            ]
        )
    return rows


def draw_footer(c, width, font_name, school):
    c.setFont(font_name, 10)
    c.setFillColor(colors.HexColor("#1C1D21"))
    c.drawString(18 * mm, 18 * mm, f"Principal: {school['principal_name']}")
    c.drawRightString(width - 18 * mm, 18 * mm, "Generated by Report Card System")


def draw_summary_lines(c, x, start_y, font_name, student):
    c.setFont(font_name, 10)
    c.setFillColor(colors.HexColor("#1C1D21"))
    c.drawString(x, start_y, f"Overall Total: {student['overall_total']:.2f}")
    c.drawString(x, start_y - 6 * mm, f"Overall Average: {student['overall_average']:.2f}")
    c.drawString(x, start_y - 12 * mm, f"Overall Position: {ordinal(student['overall_position'])}")
    c.drawString(x, start_y - 18 * mm, f"Remark: {remark_from_average(student['overall_average'])}")

def render_classic_template(c, width, height, font_name, school, student, rows):
    primary = colors.HexColor(school["theme_primary"])
    secondary = colors.HexColor(school["theme_secondary"])
    light_bg = colors.HexColor(school["theme_surface"])
    grid_color = colors.HexColor(darken(school["theme_primary"], 0.45))

    c.setFillColor(primary)
    c.rect(0, height - 56 * mm, width, 56 * mm, stroke=0, fill=1)

    if school["logo_style"] == "watermark":
        draw_logo(c, school, width, height, "classic")

    c.setFillColor(colors.white)
    c.setFont(font_name, 22)
    c.drawString(18 * mm, height - 20 * mm, school["school_name"])
    c.setFont(font_name, 11)
    c.drawString(18 * mm, height - 28 * mm, school["school_address"])
    c.drawString(18 * mm, height - 34 * mm, f"Session: {school['session']}   Term: {school['term']}")

    c.setFillColor(colors.HexColor("#111317"))
    c.setFont(font_name, 12)
    c.drawString(18 * mm, height - 64 * mm, f"Name: {student['student_name']}")
    c.drawString(18 * mm, height - 71 * mm, f"Student ID: {student['student_id'] or 'N/A'}")
    c.drawString(18 * mm, height - 78 * mm, f"Class: {school['default_class'] or 'N/A'}")
    c.drawString(18 * mm, height - 85 * mm, f"Overall Position: {ordinal(student['overall_position'])}")
    c.drawString(
        18 * mm,
        height - 92 * mm,
        f"Template: {TEMPLATE_LABELS[school['template_id']]} | Theme: {school['theme_primary']}",
    )

    if school["logo_style"] != "watermark":
        draw_logo(c, school, width, height, "classic")

    table = Table(rows, colWidths=[56 * mm, 23 * mm, 23 * mm, 22 * mm, 24 * mm, 18 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), secondary),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 8.8),
                ("GRID", (0, 0), (-1, -1), 0.45, grid_color),
                ("BACKGROUND", (0, 1), (-1, -1), light_bg),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    _, table_height = table.wrap(0, 0)
    table_x = 18 * mm
    table_y = max(72 * mm, height - 104 * mm - table_height)
    table.drawOn(c, table_x, table_y)

    summary_y = max(42 * mm, table_y - 8 * mm)
    draw_summary_lines(c, 18 * mm, summary_y, font_name, student)
    draw_footer(c, width, font_name, school)


def render_modern_template(c, width, height, font_name, school, student, rows):
    primary = colors.HexColor(school["theme_primary"])
    secondary = colors.HexColor(school["theme_secondary"])
    light_bg = colors.HexColor(lighten(school["theme_surface"], 0.35))
    light_card = colors.HexColor(lighten(school["theme_surface"], 0.12))
    grid_color = colors.HexColor(darken(school["theme_primary"], 0.42))

    sidebar_w = 56 * mm

    c.setFillColor(light_bg)
    c.rect(0, 0, width, height, stroke=0, fill=1)

    if school["logo_style"] == "watermark":
        draw_logo(c, school, width, height, "modern")

    c.setFillColor(primary)
    c.rect(0, 0, sidebar_w, height, stroke=0, fill=1)

    c.setFillColor(colors.white)
    c.setFont(font_name, 17)
    c.drawString(8 * mm, height - 20 * mm, school["school_name"])
    c.setFont(font_name, 9.4)
    c.drawString(8 * mm, height - 27 * mm, school["school_address"])
    c.drawString(8 * mm, height - 34 * mm, f"{school['session']} | {school['term']}")

    c.setFont(font_name, 10)
    c.drawString(8 * mm, height - 52 * mm, f"Student: {student['student_name']}")
    c.drawString(8 * mm, height - 59 * mm, f"ID: {student['student_id'] or 'N/A'}")
    c.drawString(8 * mm, height - 66 * mm, f"Class: {school['default_class'] or 'N/A'}")
    c.drawString(8 * mm, height - 73 * mm, f"Overall: {ordinal(student['overall_position'])}")

    c.setFillColor(colors.HexColor(lighten(school["theme_primary"], 0.1)))
    c.roundRect(7 * mm, 28 * mm, 42 * mm, 30 * mm, 4 * mm, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(font_name, 9.2)
    c.drawString(10 * mm, 50 * mm, "Theme")
    c.drawString(10 * mm, 44 * mm, school["theme_primary"])
    c.drawString(10 * mm, 38 * mm, TEMPLATE_LABELS[school["template_id"]])

    main_x = sidebar_w + 10 * mm
    c.setFillColor(colors.HexColor("#151A28"))
    c.setFont(font_name, 18)
    c.drawString(main_x, height - 20 * mm, "Report Card")
    c.setFont(font_name, 10)
    c.drawString(
        main_x,
        height - 27 * mm,
        f"Weighting: ClassScore {school['class_weight']:.2f}% | ExamScore {school['exam_weight']:.2f}%",
    )

    c.setFillColor(light_card)
    c.roundRect(main_x, height - 44 * mm, width - main_x - 14 * mm, 13 * mm, 3 * mm, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#0E1320"))
    c.setFont(font_name, 9.6)
    c.drawString(main_x + 4 * mm, height - 36 * mm, f"Name: {student['student_name']}")
    c.drawRightString(width - 16 * mm, height - 36 * mm, f"Position: {ordinal(student['overall_position'])}")

    if school["logo_style"] != "watermark":
        draw_logo(c, school, width, height, "modern")

    table = Table(rows, colWidths=[38 * mm, 17 * mm, 17 * mm, 17 * mm, 18 * mm, 13 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), secondary),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 8.2),
                ("GRID", (0, 0), (-1, -1), 0.45, grid_color),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    _, table_height = table.wrap(0, 0)
    table_y = max(70 * mm, height - 90 * mm - table_height)
    table.drawOn(c, main_x, table_y)

    summary_y = max(40 * mm, table_y - 8 * mm)
    draw_summary_lines(c, main_x, summary_y, font_name, student)
    draw_footer(c, width, font_name, school)


def render_minimal_template(c, width, height, font_name, school, student, rows):
    primary = colors.HexColor(school["theme_primary"])
    secondary = colors.HexColor(school["theme_secondary"])
    line_color = colors.HexColor(darken(school["theme_primary"], 0.35))
    light_bg = colors.HexColor(school["theme_surface"])

    c.setFillColor(colors.white)
    c.rect(0, 0, width, height, stroke=0, fill=1)

    if school["logo_style"] == "watermark":
        draw_logo(c, school, width, height, "minimal")

    c.setFillColor(primary)
    c.rect(0, height - 8 * mm, width, 8 * mm, stroke=0, fill=1)

    c.setFillColor(colors.HexColor("#121419"))
    c.setFont(font_name, 20)
    c.drawString(18 * mm, height - 20 * mm, school["school_name"])
    c.setFont(font_name, 10)
    c.drawString(18 * mm, height - 26 * mm, school["school_address"])
    c.drawRightString(width - 18 * mm, height - 26 * mm, f"{school['session']} | {school['term']}")

    c.setStrokeColor(line_color)
    c.setLineWidth(0.8)
    c.line(18 * mm, height - 30 * mm, width - 18 * mm, height - 30 * mm)

    c.setFont(font_name, 10)
    c.drawString(18 * mm, height - 38 * mm, f"Student: {student['student_name']}")
    c.drawString(18 * mm, height - 44 * mm, f"Student ID: {student['student_id'] or 'N/A'}")
    c.drawString(18 * mm, height - 50 * mm, f"Class: {school['default_class'] or 'N/A'}")
    c.drawRightString(width - 18 * mm, height - 38 * mm, f"Template: {TEMPLATE_LABELS[school['template_id']]}")
    c.drawRightString(width - 18 * mm, height - 44 * mm, f"Theme: {school['theme_primary']}")
    c.drawRightString(width - 18 * mm, height - 50 * mm, f"Position: {ordinal(student['overall_position'])}")

    if school["logo_style"] != "watermark":
        draw_logo(c, school, width, height, "minimal")

    table = Table(rows, colWidths=[56 * mm, 23 * mm, 23 * mm, 22 * mm, 24 * mm, 18 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), secondary),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 8.8),
                ("GRID", (0, 0), (-1, -1), 0.4, line_color),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light_bg]),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    _, table_height = table.wrap(0, 0)
    table_y = max(76 * mm, height - 96 * mm - table_height)
    table.drawOn(c, 18 * mm, table_y)

    summary_y = max(40 * mm, table_y - 8 * mm)
    c.setFillColor(light_bg)
    c.roundRect(16 * mm, summary_y - 22 * mm, width - 32 * mm, 28 * mm, 3 * mm, stroke=0, fill=1)
    draw_summary_lines(c, 20 * mm, summary_y, font_name, student)
    draw_footer(c, width, font_name, school)


def render_report_card(pdf_path: Path, school, student):
    font_name = register_fonts()
    width, height = A4
    c = canvas.Canvas(str(pdf_path), pagesize=A4)

    rows = build_subject_rows(student)
    template_id = school["template_id"]

    if template_id == "modern":
        render_modern_template(c, width, height, font_name, school, student, rows)
    elif template_id == "minimal":
        render_minimal_template(c, width, height, font_name, school, student, rows)
    else:
        render_classic_template(c, width, height, font_name, school, student, rows)

    c.showPage()
    c.save()

def aggregate_students(subject_files, class_weight: float, exam_weight: float):
    students = {}
    warnings = []
    subject_names = set()

    for file_storage in subject_files:
        if not file_storage or not file_storage.filename:
            continue

        subject_name = safe_stem(file_storage.filename)

        try:
            df = read_marks_dataframe(file_storage)
        except Exception as exc:
            warnings.append(f"{file_storage.filename}: could not read file ({exc})")
            continue

        if df.empty:
            warnings.append(f"{file_storage.filename}: file is empty")
            continue

        column_map = {normalize_column(col): col for col in df.columns}

        id_col = pick_column(
            column_map,
            ["student_id", "studentid", "id", "admission_no", "admission_number", "reg_no", "registration_no"],
        )
        name_col = pick_column(column_map, ["name", "student_name", "student", "full_name"])
        class_score_col = pick_column(column_map, ["classscore", "class_score", "class_mark", "ca"])
        exam_score_col = pick_column(column_map, ["examscore", "exam_score", "exam", "exam_mark"])

        missing = []
        if not id_col:
            missing.append("student ID")
        if not name_col:
            missing.append("Name")
        if not class_score_col:
            missing.append("ClassScore")
        if not exam_score_col:
            missing.append("ExamScore")

        if missing:
            warnings.append(f"{file_storage.filename}: missing required columns ({', '.join(missing)})")
            continue

        skipped_rows = 0
        added_rows = 0
        for _, row in df.iterrows():
            raw_name = str(row.get(name_col, "")).strip()
            if not raw_name or raw_name.lower() == "nan":
                skipped_rows += 1
                continue

            raw_id = str(row.get(id_col, "")).strip()
            student_id = "" if raw_id.lower() == "nan" else raw_id

            class_score = coerce_float(row.get(class_score_col))
            exam_score = coerce_float(row.get(exam_score_col))

            if class_score is None and exam_score is None:
                skipped_rows += 1
                continue

            class_score = 0.0 if class_score is None else class_score
            exam_score = 0.0 if exam_score is None else exam_score

            final_score = round((class_score * class_weight + exam_score * exam_weight) / 100.0, 2)

            key = student_id.lower() if student_id else normalize_column(raw_name)
            if not key:
                skipped_rows += 1
                continue

            if key not in students:
                students[key] = {
                    "student_name": raw_name,
                    "student_id": student_id,
                    "subjects": {},
                }

            if not students[key]["student_id"] and student_id:
                students[key]["student_id"] = student_id

            students[key]["subjects"][subject_name] = {
                "class_score": round(class_score, 2),
                "exam_score": round(exam_score, 2),
                "final_score": final_score,
                "subject_position": None,
            }
            added_rows += 1

        if skipped_rows:
            warnings.append(f"{file_storage.filename}: skipped {skipped_rows} row(s) with missing/invalid values")
        if added_rows:
            subject_names.add(subject_name)

    students = {student_key: student for student_key, student in students.items() if student["subjects"]}

    for subject_name in sorted(subject_names):
        scored_rows = []
        for student_key, student in students.items():
            if subject_name in student["subjects"]:
                score = student["subjects"][subject_name]["final_score"]
                scored_rows.append((student_key, score, student["student_name"]))

        subject_positions = assign_competition_positions(scored_rows)
        for student_key, position in subject_positions.items():
            students[student_key]["subjects"][subject_name]["subject_position"] = position

    overall_rows = []
    for student_key, student in students.items():
        finals = [details["final_score"] for details in student["subjects"].values()]
        total = round(sum(finals), 2)
        average = round(total / len(finals), 2) if finals else 0.0
        student["overall_total"] = total
        student["overall_average"] = average
        student["overall_position"] = None
        overall_rows.append((student_key, average, student["student_name"]))

    overall_positions = assign_competition_positions(overall_rows)
    for student_key, position in overall_positions.items():
        students[student_key]["overall_position"] = position

    return students, warnings, sorted(subject_names)


def create_excel_template():
    sample_rows = [
        {
            "student ID": "STD-001",
            "Name": "Ada Johnson",
            "ClassScore": 38,
            "ExamScore": 62,
        },
        {
            "student ID": "STD-002",
            "Name": "Musa Ibrahim",
            "ClassScore": 34,
            "ExamScore": 58,
        },
    ]
    instructions = [
        {
            "Instruction": "Use one file per subject and name the file after the subject (e.g., Chemistry.xlsx)."
        },
        {"Instruction": "Required columns per file: student ID, Name, ClassScore, ExamScore."},
        {
            "Instruction": "Set ClassScore% and ExamScore% in the frontend (must add up to 100). Backend computes weighted final score."
        },
        {
            "Instruction": "Choose one PDF template and custom colors in the frontend. Backend applies that design to all cards."
        },
        {
            "Instruction": "Optional: upload a school logo and choose how it appears (header badge, watermark, corner stamp, or side mark)."
        },
        {"Instruction": "Backend computes each student's overall position and per-subject position."},
    ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(sample_rows).to_excel(writer, index=False, sheet_name="Marks Template")
        pd.DataFrame(instructions).to_excel(writer, index=False, sheet_name="Instructions")
    output.seek(0)
    return output


def cleanup_generated_jobs(max_age_seconds: int = 3600):
    if not GENERATED_DIR.exists():
        return

    now = time.time()
    for entry in GENERATED_DIR.iterdir():
        if not entry.is_dir():
            continue
        try:
            age = now - entry.stat().st_mtime
            if age > max_age_seconds:
                shutil.rmtree(entry, ignore_errors=True)
        except Exception:
            continue


def prune_empty_parents(path: Path, stop_at: Path):
    current = path
    while True:
        if not current.exists() or current == stop_at:
            break
        try:
            current.rmdir()
        except OSError:
            break
        current = current.parent


@app.get("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.get("/api/template")
def download_template():
    template_stream = create_excel_template()
    return send_file(
        template_stream,
        as_attachment=True,
        download_name="subject_marks_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/generate")
def generate_reports():
    cleanup_generated_jobs(max_age_seconds=3600)

    subject_files = request.files.getlist("subject_files")
    if not subject_files:
        return jsonify({"error": "Upload at least one subject file."}), 400

    class_weight = coerce_float(request.form.get("class_weight", ""))
    exam_weight = coerce_float(request.form.get("exam_weight", ""))

    if class_weight is None or exam_weight is None:
        return jsonify({"error": "Provide both ClassScore and ExamScore percentages."}), 400
    if class_weight < 0 or exam_weight < 0:
        return jsonify({"error": "Percentages cannot be negative."}), 400
    if abs((class_weight + exam_weight) - 100.0) > 0.001:
        return jsonify({"error": "ClassScore% and ExamScore% must add up to exactly 100."}), 400

    template_id = str(request.form.get("template_id", "classic")).strip().lower()
    if template_id not in TEMPLATE_LABELS:
        template_id = "classic"

    logo_style = str(request.form.get("logo_style", "header_badge")).strip().lower()
    if logo_style not in LOGO_STYLE_LABELS:
        logo_style = "header_badge"

    logo_opacity = coerce_float(request.form.get("logo_opacity", "14"))
    if logo_opacity is None:
        logo_opacity = 14.0
    logo_opacity = max(5.0, min(45.0, logo_opacity))

    primary_color = normalize_hex_color(request.form.get("primary_color", DEFAULT_PRIMARY), DEFAULT_PRIMARY)
    secondary_default = darken(primary_color, 0.18)
    surface_default = lighten(primary_color, 0.90)
    secondary_color = normalize_hex_color(request.form.get("secondary_color", secondary_default), secondary_default)
    surface_color = normalize_hex_color(request.form.get("surface_color", surface_default), surface_default)

    students, warnings, subject_names = aggregate_students(subject_files, class_weight, exam_weight)
    if not students:
        return jsonify({"error": "No valid student records found in uploaded files.", "warnings": warnings}), 400

    logo_file = request.files.get("school_logo")
    logo_assets, logo_error = load_logo_assets(logo_file, logo_opacity / 100.0)
    if logo_error:
        warnings.append(f"School logo: {logo_error}")

    school = {
        "school_name": request.form.get("school_name", "My School").strip() or "My School",
        "school_address": request.form.get("school_address", "No address provided").strip() or "No address provided",
        "principal_name": request.form.get("principal_name", "Principal").strip() or "Principal",
        "session": request.form.get("session", "2025/2026").strip() or "2025/2026",
        "term": request.form.get("term", "First Term").strip() or "First Term",
        "default_class": request.form.get("default_class", "").strip(),
        "class_weight": round(class_weight, 2),
        "exam_weight": round(exam_weight, 2),
        "template_id": template_id,
        "theme_primary": primary_color,
        "theme_secondary": secondary_color,
        "theme_surface": surface_color,
        "logo_style": logo_style,
        "logo_opacity": round(logo_opacity, 2),
        "logo_assets": logo_assets,
    }

    job_id = secure_filename(uuid.uuid4().hex[:12])
    job_dir = GENERATED_DIR / job_id
    students_dir = job_dir / "students"
    students_dir.mkdir(parents=True, exist_ok=True)

    student_payload = []
    pdf_paths = []

    ordered_students = sorted(students.values(), key=lambda student: (student["overall_position"], student["student_name"].lower()))

    for idx, student in enumerate(ordered_students, start=1):
        safe_name = normalize_column(student["student_name"])[:50] or f"student_{idx}"
        pdf_name = f"{idx:03d}_{safe_name}.pdf"
        pdf_path = students_dir / pdf_name
        render_report_card(pdf_path, school, student)
        pdf_paths.append(pdf_path)

        student_payload.append(
            {
                "student_name": student["student_name"],
                "student_id": student["student_id"],
                "overall_average": student["overall_average"],
                "overall_position": student["overall_position"],
                "overall_position_label": ordinal(student["overall_position"]),
                "subject_positions": {
                    subject: details["subject_position"] for subject, details in sorted(student["subjects"].items())
                },
                "pdf_url": f"/api/download/{job_id}/students/{pdf_name}?delete=1",
            }
        )

    compiled_pdf_name = "compiled_report_cards.pdf"
    compiled_pdf_path = job_dir / compiled_pdf_name
    writer = PdfWriter()
    for path in pdf_paths:
        writer.append(str(path))
    with compiled_pdf_path.open("wb") as compiled_file:
        writer.write(compiled_file)

    zip_name = "all_report_cards.zip"
    zip_path = job_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(compiled_pdf_path, arcname=compiled_pdf_name)
        for path in pdf_paths:
            archive.write(path, arcname=f"students/{path.name}")

    return jsonify(
        {
            "job_id": job_id,
            "student_count": len(student_payload),
            "subject_count": len(subject_names),
            "subjects": subject_names,
            "class_weight": school["class_weight"],
            "exam_weight": school["exam_weight"],
            "template_id": template_id,
            "template_label": TEMPLATE_LABELS[template_id],
            "primary_color": school["theme_primary"],
            "secondary_color": school["theme_secondary"],
            "surface_color": school["theme_surface"],
            "has_logo": bool(logo_assets),
            "logo_style": school["logo_style"],
            "logo_style_label": LOGO_STYLE_LABELS[school["logo_style"]],
            "logo_opacity": school["logo_opacity"],
            "warnings": warnings,
            "compiled_pdf_url": f"/api/download/{job_id}/{compiled_pdf_name}?delete=1",
            "zip_url": f"/api/download/{job_id}/{zip_name}?delete=1",
            "students": student_payload,
        }
    )

@app.get("/api/download/<job_id>/<path:filename>")
def download_generated(job_id, filename):
    safe_job_id = secure_filename(job_id)
    base = (GENERATED_DIR / safe_job_id).resolve()
    target = (base / filename).resolve()
    delete_after = str(request.args.get("delete", "")).strip().lower() in {"1", "true", "yes"}

    if base not in target.parents and target != base:
        return jsonify({"error": "Invalid file path."}), 400
    if not target.exists() or not target.is_file():
        return jsonify({"error": "File not found."}), 404

    if delete_after:
        @after_this_request
        def remove_generated_files(response):
            try:
                # If ZIP is downloaded, remove the whole job folder at once.
                if target.name == "all_report_cards.zip":
                    shutil.rmtree(base, ignore_errors=True)
                    return response

                # Otherwise delete only the downloaded file and prune empty folders.
                if target.exists():
                    target.unlink(missing_ok=True)
                prune_empty_parents(target.parent, base)
            except Exception:
                pass
            return response

    return send_file(target, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
