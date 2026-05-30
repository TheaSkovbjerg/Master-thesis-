from nicegui import app, classes, events, ui
import fitz
import uuid
from pydantic import BaseModel 
from openai import AsyncClient, OpenAI
from pypdf import PdfReader
from io import BytesIO
from typing import Optional
from typing import Literal
import pdfplumber
from openpyxl import Workbook, load_workbook
import os
import Levenshtein, re
import pytesseract
from pdf2image import convert_from_bytes, convert_from_path
from typing import Literal
import asyncio
import gspread
from google.oauth2.service_account import Credentials

PDF_DIR = '/tmp/nicegui_pdfs'
os.makedirs(PDF_DIR, exist_ok=True)
app.add_static_files('/pdfs', PDF_DIR)

all_candidates = []

cv_pdf_path = None
cover_pdf_path = None

cv_tokens = []
cover_tokens = []

cv_raw_text = ''
cover_raw_text = ''

client = AsyncClient(api_key='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')
input_area = ui.textarea(label='PDF text').classes('w-full') 
output_area = ui.textarea(label='Extracted text').classes('w-full').style(
    'min-width: 500px; border: 2px solid white;'
)
ui.add_head_html('''
                 <style>
                 textarea {
                 color: white !important;
                 }
                 </style>
                 ''')

class CandidateInfo(BaseModel):
    name: Optional[str] = None
    bachelor_gpa: Optional[list[str]] = None
    bachelor_gpa_snippets: Optional[list[str]] = None
    bachelor_education: Optional[list[str]] = None
    bachelor_education_snippets: Optional[list[str]] = None
    master_gpa: Optional[list[str]] = None
    master_gpa_snippets: Optional[list[str]] = None
    master_education: Optional[list[str]] = None
    master_education_snippets: Optional[list[str]] = None
    exchange_gpa: Optional[list[str]] = None
    exchange_gpa_snippets: Optional[list[str]] = None
    exchange_education: Optional[list[str]] = None
    exchange_education_snippets: Optional[list[str]] = None
    student_job: Optional[list[str]] = None
    student_job_snippets: Optional[list[str]] = None
    persona: Optional[str] = None
    persona_reason: Optional[str] = None
    persona_snippets: Optional[list[str]] = None

system_instructions = """
You will receive a CV. Classify whether the following are true, and extract the text that relates to it.
"""

class TechStrategist(BaseModel):
    # UNIVERSITIES
    extract_evidence_of_attending_Copenhagen_Business_School: list[str]
    boolean_classify_evidence_of_attending_Copenhagen_Business_School: bool
    extract_evidence_of_attending_Aarhus_university: list[str]
    boolean_classify_evidence_of_attending_Aarhus_university: bool
    extract_evidence_of_attending_University_of_Copenhagen: list[str]
    boolean_classify_evidence_of_attending_University_of_Copenhagen: bool
    extract_evidence_of_affending_Danmarks_Tekniske_Universitet_on_a_tech_management_track: list[str]
    boolean_classify_evidence_of_affending_Danmarks_Tekniske_Universitet_on_a_tech_management_track: bool
    extract_evidence_of_attending_Lund_University: list[str]
    boolean_classify_evidence_of_attending_Lund_University: bool

    # PROGRAMS
    extract_evidence_of_business_admin_digital_business_cbs: list[str]
    boolean_classify_business_admin_digital_business_cbs: bool
    extract_evidence_of_strategy_org_leadership_cbs: list[str]
    boolean_classify_strategy_org_leadership_cbs: bool
    extract_evidence_of_innovation_business_dev_cbs: list[str]
    boolean_classify_innovation_business_dev_cbs: bool
    extract_evidence_of_economics_ku_au: list[str]
    boolean_classify_economics_ku_au: bool
    extract_evidence_of_international_business_cbs_au: list[str]
    boolean_classify_international_business_cbs_au: bool
    extract_evidence_of_digital_business_management_au: list[str]
    boolean_classify_digital_business_management_au: bool

    # SIMILAR PROGRAMS
    boolean_classify_similar_relevant_program: bool

    # GPA
    extract_evidence_of_grade_point_average: list[str]
    boolean_classify_evidence_of_grade_point_average: bool
    grade_point_average: float | None

    # KEYWORDS
    extract_evidence_of_strategic_problem_solving: list[str]
    boolean_classify_strategic_problem_solving: bool
    extract_evidence_of_digital_transformation: list[str]
    boolean_classify_digital_transformation: bool
    extract_evidence_of_technology_enabled_business_models: list[str]
    boolean_classify_technology_enabled_business_models: bool
    extract_evidence_of_c_level_advisory: list[str]
    boolean_classify_c_level_advisory: bool
    extract_evidence_of_operating_model_design: list[str]
    boolean_classify_operating_model_design: bool
    extract_evidence_of_business_case_development: list[str]
    boolean_classify_business_case_development: bool
    extract_evidence_of_technology_strategy: list[str]
    boolean_classify_technology_strategy: bool

class TransformationOrchestratorPMO(BaseModel):
    # UNIVERSITIES
    extract_evidence_of_attending_Copenhagen_Business_School: list[str]
    boolean_classify_evidence_of_attending_Copenhagen_Business_School: bool
    extract_evidence_of_attending_Aarhus_university: list[str]
    boolean_classify_evidence_of_attending_Aarhus_university: bool
    extract_evidence_of_attending_Danmarks_Tekniske_Universitet: list[str]
    boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet: bool
    extract_evidence_of_attending_IT_University_of_Copenhagen: list[str]
    boolean_classify_evidence_of_attending_IT_University_of_Copenhagen: bool
    extract_evidence_of_attending_Aalborg_University: list[str]
    boolean_classify_evidence_of_attending_Aalborg_University: bool

    # PROGRAMS
    extract_evidence_of_digital_business_management_au: list[str]
    boolean_classify_digital_business_management_au: bool
    extract_evidence_of_business_administration_and_information_systems_cbs: list[str]
    boolean_classify_business_administration_and_information_systems_cbs: bool
    extract_evidence_of_management_of_innovation_and_business_development_cbs: list[str]
    boolean_classify_management_of_innovation_and_business_development_cbs: bool
    extract_evidence_of_operations_and_supply_chain_programs: list[str]
    boolean_classify_operations_and_supply_chain_programs: bool
    extract_evidence_of_engineering_with_business_specialization_dtu: list[str]
    boolean_classify_engineering_with_business_specialization_dtu: bool

    # SIMILAR PROGRAMS
    boolean_classify_similar_relevant_program: bool

    # GPA
    extract_evidence_of_grade_point_average: list[str]
    boolean_classify_evidence_of_grade_point_average: bool
    grade_point_average: float | None

    # KEYWORDS
    extract_evidence_of_program_management: list[str]
    boolean_classify_program_management: bool
    extract_evidence_of_execution_and_value_realization: list[str]
    boolean_classify_execution_and_value_realization: bool
    extract_evidence_of_governance_and_operating_models: list[str]
    boolean_classify_governance_and_operating_models: bool
    extract_evidence_of_change_enablement: list[str]
    boolean_classify_change_enablement: bool
    extract_evidence_of_stakeholder_coordination: list[str]
    boolean_classify_stakeholder_coordination: bool
    extract_evidence_of_implementation_roadmap_design: list[str]
    boolean_classify_implementation_roadmap_design: bool


class TechTranslator(BaseModel):
    # UNIVERSITIES
    extract_evidence_of_attending_IT_University_of_Copenhagen: list[str]
    boolean_classify_evidence_of_attending_IT_University_of_Copenhagen: bool
    extract_evidence_of_attending_Danmarks_Tekniske_Universitet: list[str]
    boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet: bool
    extract_evidence_of_attending_Aalborg_University: list[str]
    boolean_classify_evidence_of_attending_Aalborg_University: bool
    extract_evidence_of_attending_Copenhagen_Business_School_on_data_or_digital_tracks: list[str]
    boolean_classify_evidence_of_attending_Copenhagen_Business_School_on_data_or_digital_tracks: bool
    extract_evidence_of_attending_University_of_Copenhagen_on_quantitative_or_data_tracks: list[str]
    boolean_classify_evidence_of_attending_University_of_Copenhagen_on_quantitative_or_data_tracks: bool

    # PROGRAMS
    extract_evidence_of_business_intelligence_cbs_or_au: list[str]
    boolean_classify_business_intelligence_cbs_or_au: bool
    extract_evidence_of_business_administration_and_data_science_cbs: list[str]
    boolean_classify_business_administration_and_data_science_cbs: bool
    extract_evidence_of_computer_science_itu_dtu_aau: list[str]
    boolean_classify_computer_science_itu_dtu_aau: bool
    extract_evidence_of_data_science_programs: list[str]
    boolean_classify_data_science_programs: bool
    extract_evidence_of_analytics_and_information_systems_programs: list[str]
    boolean_classify_analytics_and_information_systems_programs: bool

    # SIMILAR PROGRAMS
    boolean_classify_similar_relevant_program: bool

    # GPA
    extract_evidence_of_grade_point_average: list[str]
    boolean_classify_evidence_of_grade_point_average: bool
    grade_point_average: float | None

    # KEYWORDS
    extract_evidence_of_data_literacy: list[str]
    boolean_classify_data_literacy: bool
    extract_evidence_of_ai_enablement: list[str]
    boolean_classify_ai_enablement: bool
    extract_evidence_of_business_it_alignment: list[str]
    boolean_classify_business_it_alignment: bool
    extract_evidence_of_systems_and_architecture_understanding: list[str]
    boolean_classify_systems_and_architecture_understanding: bool
    extract_evidence_of_analytics_interpretation: list[str]
    boolean_classify_analytics_interpretation: bool
    extract_evidence_of_use_case_development: list[str]
    boolean_classify_use_case_development: bool
    extract_evidence_of_genai_applications: list[str]
    boolean_classify_genai_applications: bool


class AppliedTechnicalSpecialist(BaseModel):
    # UNIVERSITIES
    extract_evidence_of_attending_Danmarks_Tekniske_Universitet: list[str]
    boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet: bool
    extract_evidence_of_attending_IT_University_of_Copenhagen: list[str]
    boolean_classify_evidence_of_attending_IT_University_of_Copenhagen: bool
    extract_evidence_of_attending_Aalborg_University: list[str]
    boolean_classify_evidence_of_attending_Aalborg_University: bool
    extract_evidence_of_attending_University_of_Copenhagen_on_science_faculty: list[str]
    boolean_classify_evidence_of_attending_University_of_Copenhagen_on_science_faculty: bool
    extract_evidence_of_attending_Lund_University_on_engineering_tracks: list[str]
    boolean_classify_evidence_of_attending_Lund_University_on_engineering_tracks: bool
    extract_evidence_of_attending_Southern_Denmark_University: list[str]
    boolean_classify_evidence_of_attending_Southern_Denmark_University: bool
    extract_evidence_of_attending_international_technical_programs: list[str]
    boolean_classify_evidence_of_attending_international_technical_programs: bool

    # PROGRAMS
    extract_evidence_of_computer_science: list[str]
    boolean_classify_computer_science: bool
    extract_evidence_of_applied_mathematics: list[str]
    boolean_classify_applied_mathematics: bool
    extract_evidence_of_physics: list[str]
    boolean_classify_physics: bool
    extract_evidence_of_statistics: list[str]
    boolean_classify_statistics: bool
    extract_evidence_of_engineering_software_systems_industrial_or_energy: list[str]
    boolean_classify_engineering_software_systems_industrial_or_energy: bool
    extract_evidence_of_machine_learning_and_ai_programs: list[str]
    boolean_classify_machine_learning_and_ai_programs: bool

    # SIMILAR PROGRAMS
    boolean_classify_similar_relevant_program: bool

    # GPA
    extract_evidence_of_bachelor_grade_point_average: list[str]
    boolean_classify_evidence_of_grade_point_average: bool
    extract_evidence_of_master_grade_point_average: list[str]
    boolean_classify_master_grade_point_average: bool
    bachelor_grade_point_average: float | None
    master_grade_point_average: float | None

    # KEYWORDS
    extract_evidence_of_advanced_quantitative_methods: list[str]
    boolean_classify_advanced_quantitative_methods: bool
    extract_evidence_of_machine_learning_and_modelling: list[str]
    boolean_classify_machine_learning_and_modelling: bool
    extract_evidence_of_algorithmic_thinking: list[str]
    boolean_classify_algorithmic_thinking: bool
    extract_evidence_of_statistical_analysis: list[str]
    boolean_classify_statistical_analysis: bool
    extract_evidence_of_optimization_and_simulation: list[str]
    boolean_classify_optimization_and_simulation: bool
    extract_evidence_of_technical_solution_development: list[str]
    boolean_classify_technical_solution_development: bool
    extract_evidence_of_systems_engineering: list[str]
    boolean_classify_systems_engineering: bool

    # SCORE: UNIVERSITIES
    def university_score(self):
        if any([
            self.boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet,
            self.boolean_classify_evidence_of_attending_IT_University_of_Copenhagen,
            self.boolean_classify_evidence_of_attending_Aalborg_University,
            self.boolean_classify_evidence_of_attending_University_of_Copenhagen_on_science_faculty,
            self.boolean_classify_evidence_of_attending_Lund_University_on_engineering_tracks,
        ]):
            return 5

        elif any([
            self.boolean_classify_evidence_of_attending_Southern_Denmark_University,
            self.boolean_classify_evidence_of_attending_international_technical_programs,
        ]):
            return 4

        else:
            return 0


    # SCORE: GPA
    # OBS: scorecardet skelner mellem master og bachelor grades,
    # men din klasse har kun ét samlet grade_point_average-felt.
    def gpa_score(self):
        if self.grade_point_average is None:
            return 0

        gpa = self.grade_point_average

        if gpa >= 8.5:
            return 5
        elif gpa >= 7.5:
            return 4
        elif gpa >= 6:
            return 3
        else:
            return 1


    # SCORE PROGRAM
    def program_score(self):
        if any([
            self.boolean_classify_computer_science,
            self.boolean_classify_applied_mathematics,
            self.boolean_classify_physics,
            self.boolean_classify_statistics,
            self.boolean_classify_engineering_software_systems_industrial_or_energy,
            self.boolean_classify_machine_learning_and_ai_programs,
        ]):
            return 5

        elif self.boolean_classify_similar_relevant_program:
            return 4

        else:
            return 0


    # SCORE KEYWORDS
    def keyword_score(self):
        if any([
            self.boolean_classify_advanced_quantitative_methods,
            self.boolean_classify_machine_learning_and_modelling,
            self.boolean_classify_algorithmic_thinking,
            self.boolean_classify_statistical_analysis,
            self.boolean_classify_optimization_and_simulation,
            self.boolean_classify_technical_solution_development,
            self.boolean_classify_systems_engineering,
        ]):
            return 5
        else:
            return 0


    # FINAL SCORE
    def final_score(self):
        return (
            self.university_score() +
            self.gpa_score() +
            self.program_score() +
            self.keyword_score()
        )

class IndustryFunctionalSpecialist(BaseModel):
    # UNIVERSITIES
    extract_evidence_of_attending_University_of_Copenhagen_on_life_science_or_public_sector_tracks: list[str]
    boolean_classify_evidence_of_attending_University_of_Copenhagen_on_life_science_or_public_sector_tracks: bool
    extract_evidence_of_attending_Danmarks_Tekniske_Universitet_on_energy_or_engineering_tracks: list[str]
    boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet_on_energy_or_engineering_tracks: bool
    extract_evidence_of_attending_Copenhagen_Business_School_on_financial_services_or_commercial_tracks: list[str]
    boolean_classify_evidence_of_attending_Copenhagen_Business_School_on_financial_services_or_commercial_tracks: bool
    extract_evidence_of_attending_Aarhus_university: list[str]
    boolean_classify_evidence_of_attending_Aarhus_university: bool
    extract_evidence_of_attending_Lund_University: list[str]
    boolean_classify_evidence_of_attending_Lund_University: bool
    extract_evidence_of_attending_Roskilde_University: list[str]
    boolean_classify_evidence_of_attending_Roskilde_University: bool
    extract_evidence_of_attending_Southern_Denmark_University: list[str]
    boolean_classify_evidence_of_attending_Southern_Denmark_University: bool
    extract_evidence_of_attending_Aalborg_University: list[str]
    boolean_classify_evidence_of_attending_Aalborg_University: bool

    # PROGRAMS
    extract_evidence_of_applied_economics_and_finance_ku: list[str]
    boolean_classify_applied_economics_and_finance_ku: bool
    extract_evidence_of_finance_and_strategic_management_cbs: list[str]
    boolean_classify_finance_and_strategic_management_cbs: bool
    extract_evidence_of_life_science_or_biomed_programs_ku_dtu_lund: list[str]
    boolean_classify_life_science_or_biomed_programs_ku_dtu_lund: bool
    extract_evidence_of_energy_and_engineering_programs_dtu_aau: list[str]
    boolean_classify_energy_and_engineering_programs_dtu_aau: bool
    extract_evidence_of_supply_chain_and_operations_programs: list[str]
    boolean_classify_supply_chain_and_operations_programs: bool
    extract_evidence_of_public_policy_or_political_science_ku_au: list[str]
    boolean_classify_public_policy_or_political_science_ku_au: bool

    # SIMILAR PROGRAMS
    boolean_classify_similar_relevant_program: bool

    # GPA
    extract_evidence_of_grade_point_average: list[str]
    boolean_classify_evidence_of_grade_point_average: bool
    grade_point_average: float | None

    # KEYWORDS
    extract_evidence_of_industry_expertise: list[str]
    boolean_classify_industry_expertise: bool
    extract_evidence_of_regulatory_knowledge: list[str]
    boolean_classify_regulatory_knowledge: bool
    extract_evidence_of_value_chain_insight: list[str]
    boolean_classify_value_chain_insight: bool
    extract_evidence_of_commercial_awareness: list[str]
    boolean_classify_commercial_awareness: bool
    extract_evidence_of_sector_strategy: list[str]
    boolean_classify_sector_strategy: bool
    extract_evidence_of_domain_specific_analytics: list[str]
    boolean_classify_domain_specific_analytics: bool
    extract_evidence_of_strategy_execution: list[str]
    boolean_classify_strategy_execution: bool

#definer funktioner for at læse pdf
norm = lambda t: re.sub(r"[^\w\s:/().,+-]", "", re.sub(r"\s+", " ", t.lower())).strip()

def extract_word_tokens(b):
    out = []
    with pdfplumber.open(BytesIO(b)) as pdf:
        for p, page in enumerate(pdf.pages):
            words = page.extract_words()
            mid = page.width * 0.4

            left = [w for w in words if w["x0"] < mid]
            right = [w for w in words if w["x0"] >= mid]

            if len(left) > len(words) * 0.2 and len(right) > len(words) * 0.2:
                words = sorted(left, key=lambda w: (w["top"], w["x0"])) + sorted(right, key=lambda w: (w["top"], w["x0"]))
            else:
                words = sorted(words, key=lambda w: (w["top"], w["x0"]))

            out += [{
                "text": w["text"],
                "norm": norm(w["text"]),
                "page": p,
                "x0": w["x0"],
                "x1": w["x1"],
                "top": w["top"],
                "bottom": w["bottom"],
            } for w in words]
    return out

def best_window_match(f, t, s=4):
    trg = " ".join(norm(w) for w in str(f).split() if norm(w))
    if not trg:
        return None

    n = len(trg.split())
    best = None

    for z in range(max(1, n - s), n + s + 1):
        for i in range(len(t) - z + 1):
            w = t[i:i+z]

            if len({x["page"] for x in w}) > 1:
                continue

            c = " ".join(x["norm"] for x in w if x["norm"])
            sc = Levenshtein.distance(c, trg) / max(len(trg), 1)

            if best is None or sc < best["score"]:
                best = {"score": sc, "tokens": w}

    return best if best and best["score"] <= 0.30 else None

def split_field_for_matching(f):
    if not f:
        return []

    if isinstance(f, list):
        parts = []
        for item in f:
            parts.extend(split_field_for_matching(item))
        return parts

    # split på linjeskift
    lines = [line.strip() for line in str(f).split('\n') if line.strip()]

    # behold kun linjer med lidt substans
    return [line for line in lines if len(line) >= 4]

def collect_evidence_matches(npa, tokens):
    raw_fields = [
        npa.name,
        *(npa.master_education or []),
        *(npa.master_gpa or []),
        *(npa.bachelor_education or []),
        *(npa.bachelor_gpa or []),
        *(npa.exchange_education or []),
        *(npa.exchange_gpa or []),
        *(npa.student_job or []),
    ]

    fields = []
    for f in raw_fields:
        fields.extend(split_field_for_matching(f))

    matches = []

    for f in fields:
        m = best_window_match(f, tokens)

        print("TARGET:", repr(f))
        print("MATCH:", " ".join(tok["text"] for tok in m["tokens"]) if m else None)
        print("----")

        if m:
            matches.append(m)

    return matches

def clean_field(f):
    if isinstance(f, list):
        f = " ".join(f)
    return str(f).strip()
   
async def send_input():
    
    progress.visible = True
    progress.value = 0.3

    for v in [0.2, 0.4, 0.6, 0.8]:
        progress.value = v
        label.text = f'{int(v*100)}%'
        await asyncio.sleep(0.1)

        # your real logic here
        progress.value = 1
        label.text = '100%'

        await asyncio.sleep(0.2)
        progress.visible = False
        label.visible = False

    try:
        await asyncio.sleep(2)

        t = f"CV:\n{cv_raw_text}\n\nCOVER LETTER:\n{cover_raw_text}"

        print("TEXT LENGTH:", len(t)) 
        print("FULL TEXT:", repr(t))

        completion = await client.beta.chat.completions.parse(
            model="gpt-4.1",
            messages=[
                {"role": "system", "content": (
                    "Extract all bachelor entries and bachelor GPA values from the text. "
                    "Extract all master entries and master GPA values from the text. "
                    "Extract all exchange entries and exchange GPA values from the text (if any). "
                    "Exchange may be written as 'exchange', 'exchange semester', 'exchange program', 'study abroad', or similar. "
                    "Student job refers to the most recent relevant student job experience mentioned in the text, and may be written as 'student job', 'student assistant', 'internship', or similar. "
                    "GPA may be written as 'grade average', 'average grade', or similar. "
                    "The GPA may appear on the same line or the next line after the label. "
                    "Return the GPA's exactly as written. "
                    "For each extracted value, also return the exact snippet from the text that supports it. "
                    "Return all entries you find, not just the most recent one. "
                    "education_snippets should be short exact substrings from the text."
                    "If a value is not explicitly present, return null. Do not guess. "
                    + system_instructions
                )},
                {"role": "user", "content": t},
            ],
            response_format=CandidateInfo,
        )

        npa = completion.choices[0].message.parsed

        persona, reason = system_instructions(npa)
        npa.persona = persona
        npa.persona_reason = reason

        global all_candidates
        all_candidates.append(npa)

        print(npa)

        output_area.value = (
            f"name={npa.name}\n"
            f"master_education={npa.master_education}\n"
            f"master_gpa={npa.master_gpa}\n"
            f"bachelor_education={npa.bachelor_education}\n"
            f"bachelor_gpa={npa.bachelor_gpa}\n"
            f"exchange_education={npa.exchange_education}\n"
            f"student_job={npa.student_job}\n"
            f"persona={npa.persona}\n"
            f"persona_reason={npa.persona_reason}"
            )
        
        cv_matches = collect_evidence_matches(npa, cv_tokens) if cv_tokens else []
        cover_matches = collect_evidence_matches(npa, cover_tokens) if cover_tokens else []

        print("CV TOKENS:", len(cv_tokens))
        print("COVER TOKENS:", len(cover_tokens))
        print("CV MATCHES:", cv_matches)
        print("COVER MATCHES:", cover_matches)

        words = []
        for value in [
            npa.name,
            *(npa.bachelor_gpa or []),
            *(npa.bachelor_education or []),
            *(npa.master_gpa or []),
            *(npa.master_education or []),
            *(npa.exchange_gpa or []),
            *(npa.exchange_education or []),
            *(npa.student_job or []),

        ]:
            if value:
                words += [w.strip(" ,.;:()[]{}") for w in str(value).split() if len(w.strip(" ,.;:()[]{}")) > 2]

        words = list(set(words))

        if cv_pdf_path and cv_matches:
            highlighted_cv_url = make_highlighted_pdf(cv_pdf_path, cv_matches)
            cv_viewer.props(f'src={highlighted_cv_url}')

        elif cv_pdf_path and words:
            highlighted_cv_url = make_highlighted_pdf_ocr(cv_pdf_path, words)
            cv_viewer.props(f'src={highlighted_cv_url}')

        if cover_pdf_path and cover_matches:
            highlighted_cover_url = make_highlighted_pdf(cover_pdf_path, cover_matches)
            cover_viewer.props(f'src={highlighted_cover_url}')
        elif cover_pdf_path and words:
            highlighted_cover_url = make_highlighted_pdf_ocr(cover_pdf_path, words)
            cover_viewer.props(f'src={highlighted_cover_url}')

    except Exception as e:
        print("ERROR:", e)
        output_area.value = str(e)

    finally:
        progress.visible = False
        progress.value = 0

progress = ui.linear_progress(value=0).classes('w-full')
label = ui.label('0%')

progress.visible = False
label.visible = False

ui.button("Extract", on_click=send_input)


def get_column_split(page):
    words = page.extract_words()
    if len(words) < 20:
        return None

    centers = sorted((w["x0"] + w["x1"]) / 2 for w in words)
    gaps = [(centers[i+1] - centers[i], centers[i], centers[i+1]) for i in range(len(centers)-1)]
    if not gaps:
        return None

    biggest_gap, left_edge, right_edge = max(gaps, key=lambda x: x[0])
    split = (left_edge + right_edge) / 2

    # Kun acceptér split hvis det ligner en reel midter-gutter
    if (
        biggest_gap > page.width * 0.12
        and page.width * 0.3 < split < page.width * 0.7
    ):
        left_count = sum(1 for c in centers if c < split)
        right_count = sum(1 for c in centers if c >= split)

        if left_count > len(centers) * 0.2 and right_count > len(centers) * 0.2:
            return split

    return None


def extract_page_text(page):
    split = get_column_split(page)

    if split is None:
        return page.extract_text() or ''

    left_text = page.crop((0, 0, split, page.height)).extract_text() or ''
    right_text = page.crop((split, 0, page.width, page.height)).extract_text() or ''

    return left_text + '\n' + right_text

#Upload pdf file to Nice Gui
def extract_text_from_pdf_bytes(pdf_bytes):
    text = ''

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text(x_tolerance=2, y_tolerance=3) or ''

            if not page_text.strip():
                page_text = page.extract_text() or ''

            text += page_text + '\n'

    if not text.strip():
        print("Falling back to OCR (pytesseract)...")
        images = convert_from_bytes(pdf_bytes)
        text = ''

        for img in images:
            text += pytesseract.image_to_string(img) + '\n'
    else:
        print("Used pdfplumber (text-based PDF)")

    return text

markdown = ui.markdown('Choose a PDF file!')

async def upload_cv(e: events.UploadEventArguments):
    global cv_pdf_path, cv_tokens, cv_raw_text

    pdf_bytes = await e.file.read()
    safe_name = (e.file.name or "cv.pdf").replace(" ", "_")
    filename = f'{uuid.uuid4().hex}_{safe_name}'
    cv_pdf_path = os.path.join(PDF_DIR, filename)

    with open(cv_pdf_path, 'wb') as f:
        f.write(pdf_bytes)

    cv_viewer.props(f'src=/pdfs/{filename}')
    cv_viewer.update()

    cv_raw_text = extract_text_from_pdf_bytes(pdf_bytes)
    cv_tokens = extract_word_tokens(pdf_bytes)
    print("FIRST CV TOKEN:", cv_tokens[0] if cv_tokens else "NO TOKENS")
    print("FIRST COVER TOKEN:", cover_tokens[0] if cover_tokens else "NO TOKENS")

async def upload_cover(e: events.UploadEventArguments):
    global cover_pdf_path, cover_tokens, cover_raw_text

    print('upload_cover fired')

    pdf_bytes = await e.file.read()
    safe_name = (e.file.name or "cover.pdf").replace(" ", "_")
    filename = f'{uuid.uuid4().hex}_{safe_name}'
    cover_pdf_path = os.path.join(PDF_DIR, filename)

    with open(cover_pdf_path, 'wb') as f:
        f.write(pdf_bytes)

    print('cover filename:', filename)
    print('cover path:', cover_pdf_path)

    cover_viewer.props(f'src=/pdfs/{filename}')
    cover_viewer.update()

    cover_raw_text = extract_text_from_pdf_bytes(pdf_bytes)
    cover_tokens = extract_word_tokens(pdf_bytes)

ui.label('Upload CV')
ui.upload(on_upload=upload_cv, auto_upload=True).props('accept=.pdf')

ui.label('Upload Cover Letter')
ui.upload(on_upload=upload_cover, auto_upload=True).props('accept=.pdf')

with ui.row().classes('w-full'):
    cv_viewer = ui.element('iframe').classes('w-1/2').style('height:900px; border:none;')
    cover_viewer = ui.element('iframe').classes('w-1/2').style('height:900px; border:none;')

#Add second area for highlighting nice gui
# create UI element
highlighted_output = ui.html().classes('w-full')

def make_highlighted_pdf(src_path, matches):
    doc = fitz.open(src_path)

    for m in matches:
        for tok in m["tokens"]:
            page = doc[tok["page"]]
            rect = fitz.Rect(tok["x0"], tok["top"], tok["x1"], tok["bottom"])
            page.add_highlight_annot(rect)

    out_name = f'highlighted_{uuid.uuid4().hex}.pdf'
    out_path = os.path.join(PDF_DIR, out_name)
    doc.save(out_path, garbage=4, deflate=True)
    doc.close()
    return f'/pdfs/{out_name}'

#Highlight pdf'er scannet som billeder
def make_highlighted_pdf_ocr(src_path, target_words):
    doc = fitz.open(src_path)
    images = convert_from_path(src_path)

    target_words_norm = [norm(w) for w in target_words]

    for page_num, img in enumerate(images):
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        page = doc[page_num]
        pw, ph = page.rect.width, page.rect.height

        for i, text in enumerate(data["text"]):
            word = text.strip()
            if not word:
                continue
        
            if norm(word) in target_words_norm:
                x0 = data["left"][i] * pw / img.width
                y0 = data["top"][i] * ph / img.height
                x1 = (data["left"][i] + data["width"][i]) * pw / img.width
                y1 = (data["top"][i] + data["height"][i]) * ph / img.height
                page.add_highlight_annot(fitz.Rect(x0, y0, x1, y1))
        
    out_name = f'highlighted_{uuid.uuid4().hex}.pdf'
    out_path = os.path.join(PDF_DIR, out_name)
    doc.save(out_path, garbage=4, deflate=True)
    doc.close()
    return f'/pdfs/{out_name}'


def normalize_persona_tab(persona: str | None) -> str: #denne matcher persona med fane i google sheets
    if not persona:
        return "Unknown"

    p = persona.lower()

    if "persona 1" in p or "tech strategist" in p:
        return "Persona 1 - tech strategist"
    if "persona 2" in p or "transformation orchestrator" in p or "pmo" in p:
        return "Persona 2 - transformation orchestrator (PMO)"
    if "persona 3" in p or "tech translator" in p:
        return "Persona 3 - tech translator"
    if "persona 4" in p or "applied technical specialist" in p:
        return "Persona 4 - applied technical specialist"
    if "persona 5" in p or "industry / functional specialist" in p or "industry specialist" in p:
        return "Persona 5 - industry / functional specialist"

    return "Unknown"

def save_to_google_sheets(): #denne laver kolonnerne i google sheets og sørger for at dataen kommer derind
    print("SAVE BUTTON CLICKED")
    print("all_candidates:", len(all_candidates))

    if not all_candidates:
        ui.notify("No data")
        return

    try:
        client = get_google_client()
        print("Google client created")

        spreadsheet = client.open("Accenture_prototype_sheet")
        print("Spreadsheet opened")

        print("Existing tabs before save:", [ws.title for ws in spreadsheet.worksheets()])

        headers = [
            'name', 'master_education', 'master_gpa',
            'bachelor_education', 'bachelor_gpa',
            'exchange_education', 'exchange_gpa',
            'student_job', 'persona', 'persona_reason'
        ]

        for c in all_candidates:
            sheet_name = normalize_persona_tab(c.persona)
            print("Working on sheet:", sheet_name, "| original persona:", c.persona)

            try:
                ws = spreadsheet.worksheet(sheet_name)
                print("Using existing worksheet:", sheet_name)
            except gspread.WorksheetNotFound:
                ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
                print("Created new worksheet:", sheet_name)

                existing_values = ws.get_all_values()
                if not existing_values:
                    ws.append_row(headers)
                    print("Headers added to:", sheet_name)

            row = [
                c.name[0] if isinstance(c.name, list) and c.name else c.name or '',
                ', '.join(c.master_education or []),
                ', '.join(c.master_gpa or []),
                ', '.join(c.bachelor_education or []),
                ', '.join(c.bachelor_gpa or []),
                ', '.join(c.exchange_education or []),
                ', '.join(c.exchange_gpa or []),
                ', '.join(c.student_job or []),
                c.persona or '',
                (c.persona_reason or '').replace('. ', '.\n')
            ]

            ws.append_row(row)
            print("Row appended to:", sheet_name)

        print("Existing tabs after save:", [ws.title for ws in spreadsheet.worksheets()])
        ui.notify(f"Saved {len(all_candidates)} candidates to Google Sheets")
        all_candidates.clear()

    except Exception as e:
        print("SAVE ERROR:", repr(e))
        ui.notify(f"Error: {e}")

ui.button('Save to Sheets', on_click=save_to_google_sheets)

def set_background(color: str = '#000000'):
    ui.query('body').style(f'background-color: {color}') 

set_background('#000000')

ui.run()
