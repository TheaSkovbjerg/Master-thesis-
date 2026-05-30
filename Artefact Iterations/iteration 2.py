##Importer pakker
from nicegui import ui
from pydantic import BaseModel
from openai import AsyncOpenAI
from pypdf import PdfReader
from pdf2image import convert_from_bytes
import pytesseract
import html
import json
import re

# Optional OCR fallback for scanned PDFs
try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

##API key indsat 
client = AsyncOpenAI(api_key='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')
ui.label('CV parser')

# Gemmer teksten fra den uploadede PDF
uploaded_data = {
    'text': '',
    'filename': '',
}

# Outputfelter
file_status = ui.label('Ingen PDF uploadet endnu.')
output_area = ui.textarea(label='Output').classes('w-full')
highlighted_area = ui.html('').classes('w-full border p-4')

class Education(BaseModel):
    university: str | None = None
    programme_name: str | None = None
    evidence_text: str | None = None
    gpa: str | None = None
    bachelor: str | None = None

class CVExtraction(BaseModel):
    educations: list[Education]

def extract_text_from_pdf(file_bytes: bytes) -> str:
    """Læs tekst fra en PDF. Først text extraction, så OCR fallback for scannede PDF."""
    pages_text = []

    try:
        reader = PdfReader(BytesIO(file_bytes))
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                pages_text.append(extracted)

        if pages_text:
            return "\n".join(pages_text).strip()
    except Exception as e:
        print('PdfReader extraction fejl:', e)

    # OCR fallback for scannede PDF'er
    if OCR_AVAILABLE:
        try:
            images = convert_from_bytes(file_bytes)
            ocr_pages = []
            for i, image in enumerate(images, start=1):
                text = pytesseract.image_to_string(image, lang='eng')
                if text.strip():
                    ocr_pages.append(text)
            if ocr_pages:
                return "\n".join(ocr_pages).strip()
        except Exception as e:
            print('OCR fallback fejl:', e)

    return ''


def parse_educations_from_text(text: str) -> list[Education]:
    """Fallback parser: find education entries in text using simple rules."""
    educations: list[Education] = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    current = Education()
    for i, line in enumerate(lines):
        lower = line.lower()
        if 'university' in lower or 'college' in lower or 'institution' in lower:
            value = line
            if ':' in line:
                value = line.split(':', 1)[1].strip()
            if current.university or current.programme_name:
                if not current.evidence_text:
                    current.evidence_text = ''
                educations.append(current)
                current = Education()
            current.university = value
            current.evidence_text = line
        elif 'programme' in lower or 'program' in lower or 'major' in lower:
            value = line
            if ':' in line:
                value = line.split(':', 1)[1].strip()
            current.programme_name = value
            if current.evidence_text:
                current.evidence_text += ' ' + line
            else:
                current.evidence_text = line
        elif 'gpa' in lower or 'grade point average' in lower:
            value = line
            if ':' in line:
                value = line.split(':', 1)[1].strip()
            current.gpa = value
            if current.evidence_text:
                current.evidence_text += ' ' + line
            else:
                current.evidence_text = line
        elif re.search(r"\b(bachelor(?:'s)?|b(?:\.\s*?sc|\.\s*?a)|bachelor of)\b", lower):
            line_value = line.strip()
            # normalize spacing/punctuation
            normalized_line = re.sub(r"\s+", " ", line_value).strip()
            normalized_line = normalized_line.replace("\u2019", "'")
            normalized_line = normalized_line.replace("..", ".")

            # capture full bachelor phrase possibly with following line(s)
            match = re.search(r"\b(bachelor(?:'s)?|b(?:\.\s*?sc|\.\s*?a)|bachelor of)(.*)", normalized_line, re.IGNORECASE)
            if match:
                value = match.group(0).strip()
            else:
                value = normalized_line

            # include next line(s) that look like continuation & not a new section heading
            j = i + 1
            while j < len(lines):
                look = lines[j].strip()
                if look == '':
                    break
                if re.match(r"^(gpa|\d{4}|university|college|institution|master|experience|work|skills)\b", look.lower()):
                    break
                if look.lower().startswith('in ') or look.lower().startswith('of ') or look.lower().startswith('for '):
                    value = f"{value} {re.sub(r'\s+', ' ', look).strip()}"
                    j += 1
                    continue
                break

            # if line is like 'BSc: ...', prefer the part after ':'
            if ':' in normalized_line:
                extracted = normalized_line.split(':', 1)[1].strip()
                if extracted:
                    value = extracted

            current.bachelor = value
            if current.evidence_text:
                current.evidence_text += ' ' + normalized_line
            else:
                current.evidence_text = normalized_line
        elif current.university and not current.programme_name:
            # if we're already inside an education block, capture next line as programme
            current.programme_name = line
            if current.evidence_text:
                current.evidence_text += ' ' + line
            else:
                current.evidence_text = line

    if current.university or current.programme_name or current.gpa or current.bachelor:
        if not current.evidence_text:
            current.evidence_text = ''
        educations.append(current)

    return educations


async def handle_upload(e):
    try:
        print('--- handle_upload called ---')
        print('event', type(e), e)

        files = []
        if e is None:
            raise ValueError('Upload event is None')

        if isinstance(e, dict):
            if e.get('files'):
                files = e.get('files')
            else:
                files = [e]
        elif hasattr(e, 'files') and e.files:
            files = e.files
        elif isinstance(e, (list, tuple)) and e:
            files = list(e)
        elif hasattr(e, 'content') or hasattr(e, 'filename') or hasattr(e, 'name'):
            files = [e]
        else:
            files = [e]  # final attempt

        if not files:
            raise ValueError('No file found in upload event')

        file_obj = files[0]

        if isinstance(file_obj, dict):
            filename = file_obj.get('name') or file_obj.get('filename') or 'unknown.pdf'
            content = file_obj.get('content') or file_obj.get('file') or None
        else:
            filename = getattr(file_obj, 'name', None) or getattr(file_obj, 'filename', None) or 'unknown.pdf'
            content = getattr(file_obj, 'content', None)

        if content is None and hasattr(file_obj, 'file'):
            content = file_obj.file

        if content is None and hasattr(file_obj, 'read'):
            content = file_obj

        if content is None:
            raise ValueError('Uploaded file has no content')

        if isinstance(content, (bytes, bytearray)):
            file_bytes = content
        elif hasattr(content, 'read'):
            file_bytes = content.read()
            if hasattr(file_bytes, '__await__'):
                file_bytes = await file_bytes
        elif isinstance(content, str):
            file_bytes = content.encode('utf-8')
        else:
            raise ValueError('Unsupported content type: %s' % type(content))

        print('has name', filename)
        print('content type', type(content), 'bytes' if isinstance(file_bytes, (bytes, bytearray)) else 'stream')

        print('bytes len', len(file_bytes))
        text = extract_text_from_pdf(file_bytes)
        print('extracted text len', len(text))

        uploaded_data['text'] = text
        uploaded_data['filename'] = filename

        if text.strip():
            file_status.text = f'Uploadet: {uploaded_data["filename"]}'
            output_area.value = 'PDF uploadet. Tryk Run.'
            # Hide raw PDF text in the interface; we only show extracted Education results.
            highlighted_area.content = ''
        else:
            file_status.text = f'Uploadet: {uploaded_data["filename"]}'
            output_area.value = 'PDF uploadet, men ingen tekst fundet (scannet PDF?).'
            highlighted_area.content = ''
    except Exception as ex:
        print('upload Exception', repr(ex))
        uploaded_data['text'] = ''
        uploaded_data['filename'] = ''
        file_status.text = 'Upload fejlede.'
        output_area.value = f'Fejl ved upload: {type(ex).__name__}: {ex}'
        highlighted_area.content = ''

ui.upload(
    label='Upload CV som PDF',
    on_upload=handle_upload,
    auto_upload=True,
    multiple=False,
).props('accept=.pdf').classes('w-full')

async def send_input():
    t = uploaded_data['text']

    if not t.strip():
        output_area.value = 'Der er ikke uploadet en PDF med læsbar tekst.'
        highlighted_area.content = ''
        return

    try:
        completion = await client.beta.chat.completions.parse(
            model='gpt-4.1',
            messages=[
                {
                    'role': 'system',
                    'content': (
                        'Extract all educations mentioned in the CV text. '
                        'For each education, return the university, programme name, GPA, bachelor status, and evidence text. '
                        'Only return structured output as JSON with field `educations`.'
                    ),
                },
                {
                    'role': 'user',
                    'content': t,
                },
            ],
            response_format=CVExtraction,
        )

        result = completion.choices[0].message.parsed
        educations = []

        if result:
            raw = None
            if isinstance(result, dict) and 'educations' in result:
                raw = result['educations']
            elif hasattr(result, 'educations'):
                raw = result.educations
            else:
                raw = result

            def build_education_from_item(item):
                if isinstance(item, Education):
                    return item
                if isinstance(item, dict):
                    clean = {
                        'university': item.get('university'),
                        'programme_name': item.get('programme_name'),
                        'evidence_text': item.get('evidence_text'),
                        'gpa': item.get('gpa'),
                        'bachelor': item.get('bachelor'),
                    }
                    return Education(**{k: v for k, v in clean.items() if v is not None})
                return None

            if raw:
                if isinstance(raw, list):
                    for item in raw:
                        edu = build_education_from_item(item)
                        if edu:
                            educations.append(edu)
                        elif isinstance(item, str):
                            parts = [p.strip() for p in item.split(',') if p.strip()]
                            if parts:
                                educations.append(Education(university=parts[0], programme_name=parts[1] if len(parts) > 1 else None))
                elif isinstance(raw, str):
                    # try to parse as JSON string fallback
                    try:
                        parsed_raw = json.loads(raw)
                        if isinstance(parsed_raw, dict) and 'educations' in parsed_raw:
                            for ed in parsed_raw['educations']:
                                edu = build_education_from_item(ed)
                                if edu:
                                    educations.append(edu)
                    except Exception:
                        pass

        if not educations:
            educations = parse_educations_from_text(t)

        if not educations:
            output_area.value = 'Ingen uddannelser fundet. Prøv at ændre CV prompt eller tjek tekstindholdet.'
            highlighted_area.content = html.escape(t).replace('\n', '<br>')
            return

        def mark_phrase(raw_text, phrase):
            if not phrase:
                return raw_text

            phrase_clean = ' '.join(phrase.replace('\u2019', "'").split())
            if not phrase_clean:
                return raw_text

            # flexible whitespace matching plus case-insensitive
            regex_pattern = re.escape(phrase_clean).replace('\\ ', r'\\s+')
            try:
                pattern = re.compile(regex_pattern, re.IGNORECASE)
                if pattern.search(raw_text):
                    return pattern.sub(lambda m: f'@@MARK_START@@{m.group(0)}@@MARK_END@@', raw_text, count=1)
            except re.error:
                pass

            # fallback to common bachelor tokens
            for token in ['Bachelor', "Bachelor's", 'B.Sc', 'BSc', 'BA', 'B.A']:
                token_pattern = re.compile(re.escape(token), re.IGNORECASE)
                if token_pattern.search(raw_text):
                    return token_pattern.sub(lambda m: f'@@MARK_START@@{m.group(0)}@@MARK_END@@', raw_text, count=1)

            return raw_text

        # Keep full CV text in output, but highlight only education fields.
        education_output = []
        highlighted_raw = t

        for edu in educations:
            education_output.append({
                'university': edu.university,
                'programme_name': edu.programme_name,
                'evidence_text': edu.evidence_text,
                'gpa': edu.gpa,
                'bachelor': edu.bachelor,
            })

            if edu.university:
                highlighted_raw = mark_phrase(highlighted_raw, edu.university)
            if edu.programme_name:
                highlighted_raw = mark_phrase(highlighted_raw, edu.programme_name)
            if edu.evidence_text:
                highlighted_raw = mark_phrase(highlighted_raw, edu.evidence_text)
            if edu.gpa:
                highlighted_raw = mark_phrase(highlighted_raw, edu.gpa)
            if edu.bachelor:
                highlighted_raw = mark_phrase(highlighted_raw, edu.bachelor)

        highlighted_text = html.escape(highlighted_raw).replace('@@MARK_START@@', '<mark>').replace('@@MARK_END@@', '</mark>')
        output_area.value = t + '\n\n' + json.dumps({'educations': education_output}, ensure_ascii=False, indent=2)
        highlighted_area.content = highlighted_text.replace('\n', '<br>')

        # Keep full CV text in output, but highlight only education fields.
        education_output = []
        highlighted_text = html.escape(t)

        for edu in educations:
            education_output.append({
                'university': edu.university,
                'programme_name': edu.programme_name,
                'evidence_text': edu.evidence_text,
                'gpa': edu.gpa,
                'bachelor': edu.bachelor,
            })

            if edu.university:
                highlighted_text = highlight_phrase(highlighted_text, t, edu.university)
            if edu.programme_name:
                highlighted_text = highlight_phrase(highlighted_text, t, edu.programme_name)
            if edu.evidence_text:
                highlighted_text = highlight_phrase(highlighted_text, t, edu.evidence_text)
            if edu.gpa:
                highlighted_text = highlight_phrase(highlighted_text, t, edu.gpa)
            if edu.bachelor:
                highlighted_text = highlight_phrase(highlighted_text, t, edu.bachelor)

        output_area.value = t + '\n\n' + json.dumps({'educations': education_output}, ensure_ascii=False, indent=2)
        highlighted_area.content = highlighted_text.replace('\n', '<br>')

    except Exception as e:
        output_area.value = f'Fejl: {type(e).__name__}: {e}'
        highlighted_area.content = ''

ui.button('Run', on_click=send_input)
ui.run()