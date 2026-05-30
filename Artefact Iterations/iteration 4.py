from nicegui import classes, events, ui
from pydantic import BaseModel 
from openai import OpenAI
from pypdf import PdfReader
from io import BytesIO
from typing import Optional
from typing import Literal
import pdfplumber

client = OpenAI(api_key='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')

input_area = ui.textarea(label='PDF text').classes('w-full')
output_area = ui.textarea(label='Extracted text').classes('w-full')

class CandidateInfo(BaseModel):
    gpa: list[str] = []
    gpa_snippets: list[str] = []
    education: list[str] = []
    education_snippets: list[str] = []
    persona: str | None = None
    persona_reason: str | None = None
    persona_snippets: list[str] = []

persona_prompt = """
Choose exactly one of these five personas based only on explicit evidence in the CV, cover letter and recommendation (if one of these have been uploaded).

Persona 1 - Tech Strategist:
[Key Words: Strategic problem solving, Digital transformation, Technology-enabled business models, C-level advisory, Operating model design, Business case development, Technology strategy.
Primary Target Universities: CBS, Aarhus University (AU), University of Copenhagen (KU).
Secondary / Feeder Universities: DTU (tech management tracks), Lund University.
Key Programs to Target: Business Administration & Digital Business (CBS), Strategy, Organization & Leadership (CBS), Management of Innovation & Business Development (CBS), Economics (KU / AU), International Business (CBS / AU), Digital Business Management (AU).]

Persona 2 - Transformation Orchestrator (PMO):
[Key Words: Program management, Execution & value realization, Governance & operating models, Change enablement, Stakeholder coordination, Implementation roadmap design.
Primary Target Universities: CBS, Aarhus University (AU), DTU.
Secondary / Feeder Universities: ITU, AAU.
Key Programs to Target: Digital Business Management (AU), Business Administration & Information Systems (CBS), Management of Innovation & Business Development (CBS), Operations & Supply Chain programs, Engineering with business specialization (DTU).]

Persona 3 - Tech Translator:
[Key Words: Data literacy, AI enablement, Business-IT alignment, Systems & architecture understanding, Analytics interpretation, Use-case development, GenAI applications. 
Primary Target Universities: ITU, DTU, AAU.
Secondary / Feeder Universities: CBS (Data / Digital tracks), KU (Quantitative / Data tracks).
Key Programs to Target: Business Intelligence (CBS / AU), Business Administration & Data Science (CBS), Computer Science (ITU / DTU / AAU), Data Science programs, Analytics & Information Systems programs.]

Persona 4 - Applied Technical Specialist:
[Key Words: Advanced quantitative methods, Machine learning & modelling, Algorithmic thinking, Statistical analysis, Optimization & simulation, Technical solution development, Systems engineering.
Primary Target Universities: DTU, ITU, AAU, KU (Science Faculty), Lund University (Engineering).
Secondary / Feeder Universities: SDU, International technical programs.
Key Programs to Target: Computer Science, Applied Mathematics, Physics, Statistics, Engineering (software, systems, industrial, energy), Machine Learning & AI programs.]

Persona 5 - Industry / Functional Specialist:
[Key Words: Industry expertise, Regulatory knowledge, Value chain insight, Commercial awareness, Sector strategy, Domain-specific analytics, Strategy execution.
Primary Target Universities: KU (Life Science, Public Sector), DTU (Energy, Engineering), CBS (Financial Services, Commercial), Aarhus University (AU), Lund University.
Secondary / Feeder Universities: RUC, SDU, AAU.
Key Programs to Target: Applied Economics & Finance (KU), Finance & Strategic Management (CBS), Life Science / Biomed programs (KU / DTU / Lund), Energy & Engineering programs (DTU / AAU), Supply Chain & Operations programs, Public Policy / Political Science (KU / AU).]

Return only the single best fitting persona.
Also return a short reason and exact snippets from the CV supporting the choice.
If evidence is insufficient, return null.
"""

def send_input():
    try:
        t = input_area.value or ''

        print("TEXT LENGTH:", len(t)) 
        print("FULL TEXT:", repr(t))

        completion = client.beta.chat.completions.parse(
            model="gpt-4.1",
            messages=[
                {"role": "system", "content": (
                    "Extract all education entries and all GPA values from the text. "
                    "GPA may be written as 'grade average', 'average grade', or similar. "
                    "The GPA may appear on the same line or the next line after the label. "
                    "Return the GPA exactly as written. "
                    "For each extracted value, also return the exact snippet from the text that supports it. "
                    "Return all entries you find, not just the most recent one. "
                    "education_snippets should be short exact substrings from the text. "
                    "If a value is not explicitly present, return null. Do not guess. "
                    + persona_prompt
                )},
                {"role": "user", "content": t},
            ],
            response_format=CandidateInfo,
        )

        npa = completion.choices[0].message.parsed
        print(npa)
        output_area.value = f"gpa={npa.gpa}\neducation={npa.education}\npersona={npa.persona}\npersona_reason={npa.persona_reason}"

        highlighted_output.set_content(
            highlight_all(t, npa.gpa_snippets + npa.education_snippets + npa.persona_snippets)
        )

    except Exception as e:
        print("ERROR:", e)
        output_area.value = str(e)

ui.button("Extract", on_click=send_input)


#Upload pdf file to Nice Gui
async def uploads(e: events.UploadEventArguments):
    pdf_bytes = await e.file.read()
    text = ''
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ''
    
    text = text or 'No text found in PDF.'
    
    if input_area.value:
        input_area.value += '\n\n----- Next document -----\n\n' + text
    else:
        input_area.value = text

ui.upload(on_upload=uploads, auto_upload=True, multiple=True).props('accept=.pdf')

markdown = ui.markdown('Choose a PDF file!')

#Add second area for highlighting nice gui
# create UI element
highlighted_output = ui.html().classes('w-full')

def highlight_all(t, snippets):
    for s in snippets:
        if s:
            t = t.replace(s, f'<mark>{s}</mark>')
    return t.replace('\n', '<br>')

ui.run()