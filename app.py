import streamlit as st
import PyPDF2
from groq import Groq
import json
import textwrap
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="AnswerThis: Journal Club Auto-Deck", layout="wide", initial_sidebar_state="collapsed")

# Presentation Canvas CSS
st.markdown("""
    <style>
    .stApp { background-color: #0b0f19; }
    .deck-viewport {
        background-color: #1e293b; padding: 30px; border-radius: 16px;
        max-width: 1050px; margin: 20px auto; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    }
    .slide-canvas {
        background: #ffffff; border-radius: 10px; padding: 40px 50px;
        aspect-ratio: 16 / 9; display: flex; flex-direction: column;
        justify-content: space-between; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.2);
    }
    .slide-header { border-bottom: 2px solid #e2e8f0; padding-bottom: 14px; }
    .slide-badge {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 1.5px;
        color: #059669; text-transform: uppercase; margin-bottom: 6px;
    }
    .slide-title { font-size: 1.85rem; font-weight: 800; color: #0f172a; line-height: 1.2; }
    .slide-body {
        display: grid; grid-template-columns: 1fr 1.5fr; gap: 30px;
        margin: 20px 0; align-items: start;
    }
    .highlight-column { display: flex; flex-direction: column; gap: 16px; }
    .highlight-card {
        background: #f0fdf4; border: 1px solid #d1fae5; border-left: 5px solid #059669;
        border-radius: 8px; padding: 16px 20px;
    }
    .highlight-val { font-size: 1.5rem; font-weight: 800; color: #064e3b; line-height: 1.2; }
    .highlight-lbl {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 1px;
        color: #047857; text-transform: uppercase; margin-top: 6px;
    }
    .narrative-column { display: flex; flex-direction: column; gap: 12px; }
    .takeaway-row {
        display: flex; align-items: flex-start; gap: 12px; font-size: 1.05rem;
        color: #334155; line-height: 1.5; background: #f8fafc; padding: 12px 16px;
        border-radius: 6px; border: 1px solid #edf2f7;
    }
    .bullet-dot {
        background-color: #059669; color: white; min-width: 20px; height: 20px;
        border-radius: 50%; display: flex; align-items: center; justify-content: center;
        font-size: 0.7rem; font-weight: bold; margin-top: 4px;
    }
    .slide-footer {
        display: flex; justify-content: space-between; align-items: center;
        font-size: 0.75rem; color: #94a3b8; border-top: 1px solid #f1f5f9; padding-top: 12px;
    }
    </style>
""", unsafe_allow_html=True)

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

if "deck" not in st.session_state:
    st.session_state.deck = []
if "current_slide" not in st.session_state:
    st.session_state.current_slide = 0

def create_pptx(deck_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for idx, slide_data in enumerate(deck_data):
        slide = prs.slides.add_slide(blank_layout)

        # Header
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.2))
        tf = title_box.text_frame
        p_badge = tf.paragraphs[0]
        p_badge.text = f"JOURNAL CLUB | SLIDE {idx + 1}"
        p_badge.font.size = Pt(11)
        p_badge.font.bold = True
        p_badge.font.color.rgb = RGBColor(5, 150, 105)

        p_title = tf.add_paragraph()
        p_title.text = slide_data.get('slide_title', 'Scientific Pillar')
        p_title.font.size = Pt(26)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(15, 23, 42)

        # Highlight 1
        h1_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(3.5), Inches(1.8))
        tf1 = h1_box.text_frame
        p1_val = tf1.paragraphs[0]
        p1_val.text = str(slide_data.get('highlight_1_value', '-'))
        p1_val.font.size = Pt(22)
        p1_val.font.bold = True
        p1_val.font.color.rgb = RGBColor(6, 78, 59)
        p1_lbl = tf1.add_paragraph()
        p1_lbl.text = str(slide_data.get('highlight_1_label', 'Focus Area')).upper()
        p1_lbl.font.size = Pt(11)
        p1_lbl.font.bold = True
        p1_lbl.font.color.rgb = RGBColor(4, 120, 87)

        # Highlight 2
        h2_box = slide.shapes.add_textbox(Inches(0.8), Inches(4.2), Inches(3.5), Inches(1.8))
        tf2 = h2_box.text_frame
        p2_val = tf2.paragraphs[0]
        p2_val.text = str(slide_data.get('highlight_2_value', '-'))
        p2_val.font.size = Pt(22)
        p2_val.font.bold = True
        p2_val.font.color.rgb = RGBColor(6, 78, 59)
        p2_lbl = tf2.add_paragraph()
        p2_lbl.text = str(slide_data.get('highlight_2_label', 'Core Metric')).upper()
        p2_lbl.font.size = Pt(11)
        p2_lbl.font.bold = True
        p2_lbl.font.color.rgb = RGBColor(4, 120, 87)

        # Narrative / Bullets
        body_box = slide.shapes.add_textbox(Inches(4.8), Inches(2.2), Inches(7.5), Inches(4.2))
        tf_body = body_box.text_frame
        tf_body.word_wrap = True

        for i, point in enumerate(slide_data.get('bullet_points', [])):
            p = tf_body.paragraphs[0] if i == 0 else tf_body.add_paragraph()
            p.text = f"•  {point}"
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(51, 65, 85)
            p.space_after = Pt(20)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

st.markdown("<h2 style='color: white;'>AnswerThis: Journal Club Auto-Presenter</h2>", unsafe_allow_html=True)
st.markdown("<p style='color: #94a3b8;'>Upload a Research Paper PDF to instantly generate a 5-slide literature review deck.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Scientific Paper (PDF)", type=["pdf"])

if uploaded_file and st.button("Synthesize Journal Club Deck", type="primary"):
    with st.spinner("Extracting text and identifying research gaps..."):
        try:
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            extracted_text = ""
            # Extract first 6 pages to stay within context limits while capturing core science
            for i in range(min(6, len(pdf_reader.pages))):
                extracted_text += pdf_reader.pages[i].extract_text() + "\n"

            prompt = f"""
            You are a Post-Doctoral Researcher preparing a Journal Club presentation.
            Analyze the following text extracted from a scientific paper.
            
            Extract the 5 core pillars of the paper and create exactly 5 slides:
            Slide 1: Background & Literature Context
            Slide 2: Core Hypothesis & Objectives
            Slide 3: Methodology & Experimental Design
            Slide 4: Key Results & Findings
            Slide 5: Research Gaps, Limitations, & Next Steps

            Return strictly this JSON array format containing exactly 5 objects:
            {{
                "deck": [
                    {{
                        "slide_title": "Slide Headline (max 8 words)",
                        "highlight_1_label": "Category/Metric Label",
                        "highlight_1_value": "Short descriptive phrase or data point",
                        "highlight_2_label": "Second Category Label",
                        "highlight_2_value": "Short descriptive phrase or data point",
                        "bullet_points": [
                            "Detailed scientific insight 1",
                            "Detailed scientific insight 2",
                            "Detailed scientific insight 3"
                        ]
                    }}
                ]
            }}

            Paper Text Extract:
            {extracted_text[:12000]}
            """

            response = client.chat.completions.create(
                model="openai/gpt-oss-20b",
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": "Generate the 5-slide JSON presentation deck."}
                ]
            )
            
            parsed = json.loads(response.choices[0].message.content)
            st.session_state.deck = parsed.get("deck", [])
            st.session_state.current_slide = 0
            st.rerun()

        except Exception as e:
            st.error(f"Failed to process PDF: {e}")

# Presentation Stage Viewport
if st.session_state.deck:
    deck = st.session_state.deck
    total_slides = len(deck)
    idx = st.session_state.current_slide
    current = deck[idx]

    ctrl_col1, ctrl_col2, ctrl_col3, ctrl_col4 = st.columns([1, 1, 2, 2])
    with ctrl_col1:
        if st.button("⬅️ Previous", disabled=(idx == 0), use_container_width=True):
            st.session_state.current_slide -= 1
            st.rerun()
    with ctrl_col2:
        if st.button("Next ➡️", disabled=(idx == total_slides - 1), use_container_width=True):
            st.session_state.current_slide += 1
            st.rerun()
    with ctrl_col3:
        st.markdown(f"<p style='color: white; text-align: center; margin-top: 8px;'>Slide <b>{idx + 1}</b> of <b>{total_slides}</b></p>", unsafe_allow_html=True)
    with ctrl_col4:
        pptx_file = create_pptx(deck)
        st.download_button(
            label="📥 Export to PowerPoint (.pptx)",
            data=pptx_file,
            file_name="journal_club_deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

    rows_html = "".join([
f"""<div class="takeaway-row">
<div class="bullet-dot">{i+1}</div>
<div>{b}</div>
</div>"""
        for i, b in enumerate(current.get("bullet_points", []))
    ])

    slide_html = f"""
<div class="deck-viewport">
<div class="slide-canvas">
<div class="slide-header">
<div class="slide-badge">LITERATURE REVIEW | SLIDE {idx + 1}</div>
<div class="slide-title">{current.get('slide_title', 'Scientific Pillar')}</div>
</div>
<div class="slide-body">
<div class="highlight-column">
<div class="highlight-card">
<div class="highlight-val">{current.get('highlight_1_value', '-')}</div>
<div class="highlight-lbl">{current.get('highlight_1_label', 'Focus Area')}</div>
</div>
<div class="highlight-card">
<div class="highlight-val">{current.get('highlight_2_value', '-')}</div>
<div class="highlight-lbl">{current.get('highlight_2_label', 'Core Metric')}</div>
</div>
</div>
<div class="narrative-column">
{rows_html}
</div>
</div>
<div class="slide-footer">
<span>AnswerThis AI • Academic Presentation Engine</span>
<span>CONFIDENTIAL • INTERNAL LAB MEETING USE ONLY</span>
<span>{idx + 1} / {total_slides}</span>
</div>
</div>
</div>
"""

    st.markdown(slide_html, unsafe_allow_html=True)