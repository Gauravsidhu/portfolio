import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_docx_cv():
    doc = docx.Document()

    # Set narrow margins (0.4 inch)
    for s in doc.sections:
        s.top_margin = Inches(0.4)
        s.bottom_margin = Inches(0.4)
        s.left_margin = Inches(0.4)
        s.right_margin = Inches(0.4)

    # Header Table with dark navy background
    header_table = doc.add_table(rows=1, cols=1)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    header_table.columns[0].width = Inches(7.5)

    cell = header_table.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="0E141E"/>')
    tcPr.append(shd)

    # Name
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run('PUNEET OHRI\n')
    r.font.name = 'Arial'
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)

    # Subtitle
    r2 = p.add_run('AI VIDEO EDITOR | PROMPT ENGINEER\n')
    r2.font.name = 'Arial'
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(245, 158, 11)

    # Tagline
    r3 = p.add_run('Photorealistic AI Filmmaking • JSON Prompt Systems • Premiere Pro & DaVinci Resolve')
    r3.font.name = 'Arial'
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(203, 213, 225)

    # Spacing
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(4)
    sp.paragraph_format.space_after = Pt(2)

    # 2-Column Layout
    body_table = doc.add_table(rows=1, cols=2)
    body_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    body_table.columns[0].width = Inches(2.3)
    body_table.columns[1].width = Inches(5.2)

    left_cell = body_table.cell(0, 0)
    right_cell = body_table.cell(0, 1)

    def add_heading_left(cell, text):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text.upper())
        r.font.name = 'Arial'
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)

    def add_text_left(cell, title, value=''):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        if title:
            r1 = p.add_run(title + ('\n' if value else ''))
            r1.font.name = 'Arial'
            r1.font.size = Pt(8)
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(71, 85, 105)
        if value:
            r2 = p.add_run(value)
            r2.font.name = 'Arial'
            r2.font.size = Pt(8.5)
            r2.font.color.rgb = RGBColor(30, 41, 59)

    # LEFT COLUMN
    add_heading_left(left_cell, 'CONTACT')
    add_text_left(left_cell, 'LOCATION', 'Ludhiana, Punjab, India')
    add_text_left(left_cell, 'PHONE', '+91 75080 00719')
    add_text_left(left_cell, 'EMAIL', 'puneetohri615@gmail.com')
    add_text_left(left_cell, 'PORTFOLIO', 'puneetedits.netlify.app')

    add_heading_left(left_cell, 'AI FILM PRODUCTION')
    add_text_left(left_cell, 'GENERATION', 'Runway ML Gen-3\nHiggsfield AI\nGoogle Flow • Seedance\nWan 3.0')
    add_text_left(left_cell, 'VOICE & AUDIO', 'ElevenLabs • Ambient Vedic')
    add_text_left(left_cell, 'PROMPTING', 'JSON-based prompting\nCustom GPT systems\nClaude AI Prompting\nRealistic image generation')

    add_heading_left(left_cell, 'POST-PRODUCTION')
    add_text_left(left_cell, '', 'Adobe Premiere Pro\nAdobe After Effects\nDaVinci Resolve\nAdobe Photoshop\nAI-based video editing\nColour grading • Transitions')

    add_heading_left(left_cell, 'WEB & DESIGN')
    add_text_left(left_cell, '', 'HTML5 • CSS3\nJavaScript (DOM)\nBootstrap • Tailwind CSS')

    add_heading_left(left_cell, 'LANGUAGES')
    add_text_left(left_cell, '', 'English • Hindi • Punjabi')

    add_heading_left(left_cell, 'INTERESTS')
    add_text_left(left_cell, '', 'AI video production\nCommercial Video Editing\nUI/UX design\nDigital media & web tech')

    # RIGHT COLUMN
    def add_heading_right(cell, text):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(text.upper())
        r.font.name = 'Arial'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)

    add_heading_right(right_cell, 'PROFESSIONAL PROFILE')
    p_prof = right_cell.add_paragraph()
    p_prof.paragraph_format.line_spacing = 1.15
    p_prof.paragraph_format.space_after = Pt(4)
    r_prof = p_prof.add_run(
        'AI Video Editor and Prompt Engineer specializing in photorealistic AI-generated video and image content. '
        'Delivers natural facial detail, authentic human emotion and convincing on-screen performance, moving beyond synthetic-looking '
        'output to footage that feels genuinely human. Builds refined AI generation workflows and custom JSON-based prompt systems that standardise '
        'output quality, then finishes every piece with hands-on post-production in Adobe Premiere Pro, DaVinci Resolve, and Photoshop for production-ready delivery.'
    )
    r_prof.font.name = 'Arial'
    r_prof.font.size = Pt(8.5)
    r_prof.font.color.rgb = RGBColor(51, 65, 85)

    add_heading_right(right_cell, 'CORE EXPERTISE')
    core_p = right_cell.add_paragraph()
    core_p.paragraph_format.line_spacing = 1.15
    core_p.paragraph_format.space_after = Pt(4)
    r_core = core_p.add_run(
        '• Photorealistic AI Video: Lifelike faces, natural emotion and actor-grade performance\n'
        '• Prompt Engineering: Structured JSON systems, Claude AI, and custom GPT workflows\n'
        '• AI Generation Pipeline: Runway ML Gen-3, Higgsfield, Google Flow, Seedance, Wan 3.0\n'
        '• Edit & Post-Production: Cutting, color grading, pacing in Premiere Pro, DaVinci Resolve & After Effects'
    )
    r_core.font.name = 'Arial'
    r_core.font.size = Pt(8.3)
    r_core.font.color.rgb = RGBColor(51, 65, 85)

    add_heading_right(right_cell, 'PROFESSIONAL EXPERIENCE')

    # ROLE 1: MORETASKS (PRESENT)
    p_exp1 = right_cell.add_paragraph()
    p_exp1.paragraph_format.space_before = Pt(3)
    p_exp1.paragraph_format.space_after = Pt(1)
    r_t1 = p_exp1.add_run('AI Video Editor & Prompt Engineer                    ')
    r_t1.bold = True
    r_t1.font.name = 'Arial'
    r_t1.font.size = Pt(9)
    r_t1.font.color.rgb = RGBColor(15, 23, 42)

    r_d1 = p_exp1.add_run('2026 — Present (Current)')
    r_d1.bold = True
    r_d1.font.name = 'Arial'
    r_d1.font.size = Pt(8.5)
    r_d1.font.color.rgb = RGBColor(16, 185, 129)

    p_c1 = right_cell.add_paragraph()
    p_c1.paragraph_format.space_after = Pt(2)
    r_co1 = p_c1.add_run('MoreTasks — AI Video & Creative Production')
    r_co1.font.name = 'Arial'
    r_co1.font.size = Pt(8.5)
    r_co1.font.italic = True
    r_co1.font.color.rgb = RGBColor(37, 99, 235)

    bullets_moretasks = [
        'Direct commercial AI video production, prompt architecture, and high-retention video pipelines for client digital campaigns.',
        'Achieve zero-drift character consistency and lifelike facial portrayal across complex scenes using Runway ML Gen-3 & Higgsfield.',
        'Engineer structured prompt architectures utilizing Claude AI and Custom GPTs to standardize high-retention video output.',
        'Execute master-grade editing, color grading, sound design, and motion timing in Adobe Premiere Pro, DaVinci Resolve, and After Effects.',
        'Collaborate with production teams to deliver client-ready 4K video assets and high-converting marketing reels.'
    ]
    for b in bullets_moretasks:
        bp = right_cell.add_paragraph()
        bp.paragraph_format.left_indent = Inches(0.15)
        bp.paragraph_format.space_after = Pt(1)
        bp.paragraph_format.line_spacing = 1.12
        rb = bp.add_run('• ' + b)
        rb.font.name = 'Arial'
        rb.font.size = Pt(8.2)
        rb.font.color.rgb = RGBColor(51, 65, 85)

    # ROLE 2: HOLYINDIA
    p_exp2 = right_cell.add_paragraph()
    p_exp2.paragraph_format.space_before = Pt(4)
    p_exp2.paragraph_format.space_after = Pt(1)
    r_t2 = p_exp2.add_run('AI Video Editor & Prompt Engineer                    ')
    r_t2.bold = True
    r_t2.font.name = 'Arial'
    r_t2.font.size = Pt(9)
    r_t2.font.color.rgb = RGBColor(15, 23, 42)

    r_d2 = p_exp2.add_run('3 Months Project')
    r_d2.font.name = 'Arial'
    r_d2.font.size = Pt(8.5)
    r_d2.font.color.rgb = RGBColor(100, 116, 139)

    p_c2 = right_cell.add_paragraph()
    p_c2.paragraph_format.space_after = Pt(2)
    r_co2 = p_c2.add_run('Holyindia — AI Filmmaking Project')
    r_co2.font.name = 'Arial'
    r_co2.font.size = Pt(8.5)
    r_co2.font.italic = True
    r_co2.font.color.rgb = RGBColor(37, 99, 235)

    bullets_holyindia = [
        'Followed the complete AI filmmaking pipeline end to end: character design, storyboarding, video generation and post-production.',
        'Produced high-quality mythological video content using Runway ML and Higgsfield with volumetric sacred lighting.',
        'Achieved realistic human portrayal: natural facial nuance, authentic emotion, and convincing actor-grade presence.',
        'Designed and built custom GPTs on a structured JSON prompt system to streamline multi-episode video generation.',
        'Edited and post-produced AI footage in Adobe Premiere Pro: color grading, pacing, audio synch, and master 4K finish.'
    ]
    for b in bullets_holyindia:
        bp = right_cell.add_paragraph()
        bp.paragraph_format.left_indent = Inches(0.15)
        bp.paragraph_format.space_after = Pt(1)
        bp.paragraph_format.line_spacing = 1.12
        rb = bp.add_run('• ' + b)
        rb.font.name = 'Arial'
        rb.font.size = Pt(8.2)
        rb.font.color.rgb = RGBColor(51, 65, 85)

    # CERTIFICATION
    add_heading_right(right_cell, 'CERTIFICATION')
    p_cert = right_cell.add_paragraph()
    p_cert.paragraph_format.space_after = Pt(1)
    r_cert_title = p_cert.add_run('Video Editing — Adobe Premiere Pro, Photoshop & AI Video\n')
    r_cert_title.bold = True
    r_cert_title.font.name = 'Arial'
    r_cert_title.font.size = Pt(8.5)
    r_cert_title.font.color.rgb = RGBColor(15, 23, 42)

    r_cert_sub = p_cert.add_run('DigitalPenduSkills | 2 Months | 2026 | Certificate Awarded')
    r_cert_sub.font.name = 'Arial'
    r_cert_sub.font.size = Pt(8)
    r_cert_sub.font.color.rgb = RGBColor(100, 116, 139)

    # EDUCATION
    add_heading_right(right_cell, 'EDUCATION')
    p_edu = right_cell.add_paragraph()
    p_edu.paragraph_format.space_after = Pt(1)
    r_edu_t = p_edu.add_run('Bachelor of Computer Applications (BCA)                    2023 – 2026\n')
    r_edu_t.bold = True
    r_edu_t.font.name = 'Arial'
    r_edu_t.font.size = Pt(8.5)
    r_edu_t.font.color.rgb = RGBColor(15, 23, 42)

    r_edu_sub = p_edu.add_run('Kamla Lohtia Sanatan Dharam College, Ludhiana\n')
    r_edu_sub.font.name = 'Arial'
    r_edu_sub.font.size = Pt(8)
    r_edu_sub.font.color.rgb = RGBColor(100, 116, 139)

    r_edu_10 = p_edu.add_run('Senior Secondary (10+2) 86% — Sanatan Vidya Mandir Sen. Sec. School, Ludhiana (PSEB)')
    r_edu_10.font.name = 'Arial'
    r_edu_10.font.size = Pt(8)
    r_edu_10.font.color.rgb = RGBColor(100, 116, 139)

    doc.save('Puneet_Ohri_CV.docx')
    print('Puneet_Ohri_CV.docx saved successfully!')

if __name__ == '__main__':
    create_docx_cv()
