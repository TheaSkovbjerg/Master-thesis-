from nicegui import classes, events, ui
from pydantic import BaseModel 
from openai import OpenAI
from pypdf import PdfReader
from io import BytesIO
from typing import Optional
import pdfplumber

client = OpenAI(api_key='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')

input_area = ui.textarea(label='PDF text').classes('w-full')
output_area = ui.textarea(label='Extracted text').classes('w-full')

class CandidateInfo(BaseModel):
    gpa: list[str] = []
    gpa_snippets: list[str] = []
    education: list[str] = []
    education_snippets: list[str] = []

def send_input():
    t = input_area.value

    print("TEXT LENGTH:", len(t)) 
    print("FULL TEXT:", repr(t))
    
    completion = client.chat.completions.parse(
    model="gpt-4.1",
    messages=[
        {"role": "system", "content": (
            "Extract all education entries and all GPA values from the text. "
            "GPA may be written as 'grade average', 'average grade', or similar. "
            "The GPA may appear on the same line or the next line after the label. "
            "Return the GPA exactly as written (e.g., 10.4 or 9.7). "
            "For each extracted value, also return the exact snippet from the text that supports it. "
            "Return all entries you find, not just the most recent one. "
            "education_snippets should be short exact substrings from the text (e.g. degree name only, not full sentence."
            "If a value is not explicitly present, return null. Do not guess."),},
        {"role": "system", "content": t},
    ],
    response_format=CandidateInfo,
    )

    npa = completion.choices[0].message.parsed
    print(npa)
    output_area.value = f"gpa={npa.gpa}\neducation={npa.education}"

    #highlight
    highlighted_output.set_content(
    highlight_all(t, npa.gpa_snippets + npa.education_snippets)
)

ui.button("Extract", on_click=send_input)


#Upload pdf file to Nice Gui
async def uploads(e: events.UploadEventArguments):
    pdf_bytes = await e.file.read()
    text = ''
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ''
    input_area.value = text or 'No text found in PDF.'

ui.upload(on_upload=uploads, auto_upload=True).props('accept=.pdf')

markdown = ui.markdown('Choose a PDF file!')

#Add second area for highlighting nice gui
# create UI element
highlighted_output = ui.html().classes('w-full')

def highlight_all(t, snippets):
    for s in snippets:
        if s:
            t = t.replace(s, f'<mark>{s}</mark>')
    return t.replace('\n', '<br>')

def send_input():
    t = input_area.value

    completion = client.chat.completions.parse(...)
    result = completion.choices[0].message.parsed

    highlighted_output.set_content(
        highlight(highlight(t, result.gpa_snippet), result.education_snippet)
    )

ui.run()