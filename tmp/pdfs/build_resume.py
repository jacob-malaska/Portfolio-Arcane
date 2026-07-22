from pathlib import Path

from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from pypdf import PdfReader


ROOT = Path(r"C:\Users\redma\Documents\Portfolio Arcane")
OUTPUT = ROOT / "output" / "pdf" / "Malaska_Resume_2026.pdf"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

pdfmetrics.registerFont(TTFont("Segoe", r"C:\Windows\Fonts\segoeui.ttf"))
pdfmetrics.registerFont(TTFont("Segoe-Bold", r"C:\Windows\Fonts\segoeuib.ttf"))
pdfmetrics.registerFont(TTFont("Segoe-Italic", r"C:\Windows\Fonts\segoeuii.ttf"))
pdfmetrics.registerFont(
    TTFont("Inconsolata", r"C:\Windows\Fonts\Inconsolata-VariableFont_wdth,wght.ttf")
)

PAGE_W, PAGE_H = letter
BG = HexColor("#E7E6E5")
INK = HexColor("#111936")
BODY = HexColor("#34385F")
ACCENT = HexColor("#FF3150")
RULE = HexColor("#C9C8CA")
FAINT = HexColor("#D8D7D8")

LEFT_X = 30
RIGHT_X = 30
CONTENT_X = 202
CONTENT_W = PAGE_W - CONTENT_X - RIGHT_X
LABEL_RIGHT_X = CONTENT_X - 22
DETAIL_X = CONTENT_X + 56
DATE_RIGHT_X = PAGE_W - RIGHT_X


def style(name, font, size, leading, color=BODY, space_after=0):
    return ParagraphStyle(
        name,
        fontName=font,
        fontSize=size,
        leading=leading,
        textColor=color,
        alignment=TA_LEFT,
        spaceAfter=space_after,
        allowWidows=0,
        allowOrphans=0,
    )


PROFILE = style("profile", "Segoe", 9.1, 12.0)
BODY_TEXT = style("body", "Segoe", 8.55, 10.45)
COMPACT = style("compact", "Segoe", 8.45, 10.2)
META = style("meta", "Segoe", 8.15, 10.0)


def draw_paragraph(c, text, x, top, width, paragraph_style):
    paragraph = Paragraph(text, paragraph_style)
    _, height = paragraph.wrap(width, PAGE_H)
    paragraph.drawOn(c, x, top - height)
    return top - height


def section_label(c, text, baseline):
    c.setFillColor(INK)
    c.setFont("Inconsolata", 14.5)
    c.drawRightString(LABEL_RIGHT_X, baseline, text.upper())


def rule(c, y):
    c.setStrokeColor(RULE)
    c.setLineWidth(0.75)
    c.line(LEFT_X, y, PAGE_W - RIGHT_X, y)


def role_header(c, company, role, dates, y):
    c.setFillColor(INK)
    c.setFont("Segoe-Bold", 10.1)
    c.drawString(CONTENT_X, y, company.upper())
    company_width = pdfmetrics.stringWidth(company.upper(), "Segoe-Bold", 10.1)
    c.setFillColor(BODY)
    c.setFont("Segoe-Italic", 9.2)
    c.drawString(CONTENT_X + company_width + 6, y, role.upper())
    c.setFont("Segoe", 8.8)
    c.drawRightString(DATE_RIGHT_X, y, dates)


def bullet(c, text, top):
    return draw_paragraph(
        c,
        f'<font color="#FF3150">+</font>&nbsp;{text}',
        CONTENT_X,
        top,
        CONTENT_W,
        BODY_TEXT,
    )


c = canvas.Canvas(str(OUTPUT), pagesize=letter, pageCompression=1)
c.setTitle("Jacob Malaska - Senior Industrial Designer Resume")
c.setAuthor("Jacob Malaska")
c.setSubject("Senior Industrial Designer resume")
c.setFillColor(BG)
c.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)

# Header
c.setFillColor(INK)
c.setFont("Segoe", 23.5)
c.drawRightString(LABEL_RIGHT_X, 746, "JACOB")
c.drawRightString(LABEL_RIGHT_X, 716, "MALASKA")

c.setFont("Segoe-Bold", 11.2)
c.drawString(CONTENT_X, 749, "SENIOR INDUSTRIAL DESIGNER")

contact_rows = [
    ("EMAIL", "redmalaska@gmail.com", "mailto:redmalaska@gmail.com"),
    (
        "WEB",
        "jacob-malaska.github.io/Portfolio-Arcane/",
        "https://jacob-malaska.github.io/Portfolio-Arcane/",
    ),
    ("PHONE", "435.216.2040", "tel:+14352162040"),
]
for index, (label, value, url) in enumerate(contact_rows):
    y = 729 - index * 16
    c.setFillColor(ACCENT)
    c.setFont("Inconsolata", 7.8)
    c.drawString(CONTENT_X, y, label)
    c.setFillColor(BODY)
    c.setFont("Segoe", 9.4)
    value_x = DETAIL_X
    c.drawString(value_x, y, value)
    value_w = pdfmetrics.stringWidth(value, "Segoe", 9.4)
    c.linkURL(url, (value_x, y - 2, value_x + value_w, y + 10), relative=0)

rule(c, 681)

# Profile
profile_top = 657
section_label(c, "Profile", 647)
draw_paragraph(
    c,
    "Senior industrial designer with 9+ years of experience developing consumer products from early research and strategy through mass production. Focused on expressive, approachable design across consumer electronics, hardgoods, softgoods, and jewelry, with deep experience in audio products, CAD, visualization, ergonomic validation, and DFM.",
    CONTENT_X,
    profile_top,
    CONTENT_W,
    PROFILE,
)

profile_rule_y = 591
rule(c, profile_rule_y)

# Work experience
experience_header_y = 566
section_label(c, "Experience", experience_header_y)
role_header(c, "Skullcandy Inc.", "Senior Industrial Designer", "2019 - PRESENT", experience_header_y)
y = experience_header_y - 15
y = bullet(
    c,
    "Lead industrial design across over-ear headphones, true wireless earbuds, gaming headsets, and accessories for a global consumer audio brand.",
    y,
) - 3
y = bullet(
    c,
    "Own projects from research and strategy through sketching, CAD, prototyping, visualization, fit and acoustic validation, DFM, and mass production.",
    y,
) - 3
y = bullet(
    c,
    "Collaborate with CMF, mechanical, electrical, firmware, acoustics, program management, and CG teams to ship multiple product generations on lean teams and fast timelines.",
    y,
) - 3
y = bullet(
    c,
    "Translate complex technology into approachable product experiences through tactile controls, ergonomic fit systems, visible structures, and expressive material stories.",
    y,
) - 14

role_header(c, "Freelance", "Industrial Designer / Consultant", "2020 - PRESENT", y)
y = draw_paragraph(
    c,
    "Design and development across softgoods, jewelry, automotive accessories, and visualization for clients including Shyft Global, Chums, Kizik, and Ikaika.",
    CONTENT_X,
    y - 14,
    CONTENT_W,
    COMPACT,
) - 13

role_header(c, "Ooblec LLC", "Industrial Designer", "2017 - 2018", y)
y = draw_paragraph(
    c,
    "Developed consumer products from concept through prototyping and manufacturer handoff, spanning inflatables, furniture, and audio equipment.",
    CONTENT_X,
    y - 14,
    CONTENT_W,
    COMPACT,
)

experience_rule_y = 342
rule(c, experience_rule_y)

# Expertise
expertise_top = 317
section_label(c, "Expertise", expertise_top - 8)
expertise_rows = [
    ("DESIGN", "Research and methodology  /  Sketching  /  Hardgoods  /  Softgoods development"),
    ("PRODUCTION", "DFM  /  Prototyping  /  Pattern making  /  Woodworking"),
    ("TOOLS", "SolidWorks  /  Blender  /  KeyShot  /  Adobe Creative Suite"),
    ("MEDIA", "3D visualization  /  Animation  /  Photography  /  Video editing  /  Audio and sound design"),
]
y = expertise_top - 1
for label, value in expertise_rows:
    c.setFillColor(ACCENT)
    c.setFont("Inconsolata", 7.7)
    c.drawString(CONTENT_X, y - 8, label)
    y = draw_paragraph(c, value, DETAIL_X, y, PAGE_W - RIGHT_X - DETAIL_X, META) - 5

expertise_rule_y = 229
rule(c, expertise_rule_y)

# Education
education_header_y = 204
section_label(c, "Education", education_header_y)
c.setFillColor(INK)
c.setFont("Segoe-Bold", 10.1)
c.drawString(CONTENT_X, education_header_y, "BFA, INDUSTRIAL DESIGN")
c.setFont("Segoe", 8.8)
c.drawRightString(DATE_RIGHT_X, education_header_y, "2015 - 2019")
c.setFillColor(BODY)
c.setFont("Segoe", 9.0)
c.drawString(CONTENT_X, education_header_y - 16, "Brigham Young University")
draw_paragraph(
    c,
    "Sponsored projects for Black Diamond and Native Shoes. Multidisciplinary and entrepreneurial training through the Crocker Innovation Fellowship.",
    CONTENT_X,
    education_header_y - 29,
    CONTENT_W,
    META,
)

rule(c, 137)

# Footer
c.setFillColor(FAINT)
c.setFont("Inconsolata", 8.0)
c.drawRightString(DATE_RIGHT_X, 28, "UPDATED 07.2026")

c.showPage()
c.save()

reader = PdfReader(str(OUTPUT))
assert len(reader.pages) == 1, "Resume must remain one page"
text = reader.pages[0].extract_text() or ""
for required in [
    "SENIOR INDUSTRIAL DESIGNER",
    "SKULLCANDY INC.",
    "FREELANCE",
    "OOBLEC LLC",
    "BFA, INDUSTRIAL DESIGN",
]:
    assert required in text, f"Missing expected text: {required}"

print(OUTPUT)
