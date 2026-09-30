# AnswerThis: Journal Club Auto-Presenter

A proof-of-work prototype demonstrating how AnswerThis can expand beyond literature reviews into academic presentation generation. 

## How it Works
1. Ingests a scientific research paper (PDF).
2. Uses PyPDF2 and an LLM (Llama 3 via Groq) to extract the 5 core scientific pillars (Background, Hypothesis, Methods, Results, Gaps).
3. Synthesizes a structured 5-slide JSON representation.
4. Renders an interactive digital stage using Streamlit.
5. Compiles a native `.pptx` file via `python-pptx` for lab meeting presentations.