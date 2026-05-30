from nicegui import app, classes, events, ui
import fitz
import uuid
from pydantic import BaseModel 
from openai import AsyncClient, OpenAI
from pypdf import PdfReader
from io import BytesIO
from datetime import datetime
from dateutil.relativedelta import relativedelta
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
import random
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from openai import AsyncClient

PDF_DIR = '/tmp/nicegui_pdfs'
os.makedirs(PDF_DIR, exist_ok=True)
app.add_static_files('/pdfs', PDF_DIR)
app.add_static_files('/assets', '/Users/theafrost/Artefakt')

all_candidates = []  # Global list to store candidate Excel row dictionaries

cv_pdf_path = None
cover_pdf_path = None
transcript_pdf_path = None

cv_tokens = []
cover_tokens = []
transcript_tokens = []

cv_raw_text = ''
cover_raw_text = ''
transcript_raw_text = ''

GRADING_SYSTEMS = [
    "Danish 7-point",
    "Italian 18-30",
    "German 1-5",
    "French 0-20",
    "Dutch 1-10",
    "Spanish 0-10",
    "US GPA 0-4",
    "ECTS A-F",
]

load_dotenv()
client = AsyncClient(api_key=os.getenv("API_KEY"))
# UI widgets are all created together at the bottom of the file (after their
# handler functions are defined) so the page reads top-to-bottom in the order
# the user actually moves through it: upload -> analyse -> review -> save.
# Only global styles are registered up here.

ui.add_head_html('''
<style>
    @font-face {
        font-family: 'Inter';
        src: url('/assets/Inter.woff2') format('woff2');
    }

    /* ---------- Accenture-flavoured dark theme ---------- */
    :root {
        --bg:           #0d0d10;
        --bg-elev:      #16161c;
        --bg-card:      #1c1c24;
        --border:       #2a2a35;
        --border-soft:  #22222b;
        --text:         #e9e9ec;
        --text-muted:   #a8a8b3;
        --text-dim:     #7a7a85;
        --accent:       #a100ff;   /* Accenture purple */
        --accent-soft:  #c266ff;
        --accent-bg:    rgba(161, 0, 255, 0.12);
        --success:      #5fd1a0;
        --warn:         #ffc857;
        --danger:       #ff6b6b;
        --hl-yellow:    #ffe24a;
        --hl-blue:      #87cfff;
    }

    html, body {
        background: var(--bg) !important;
        color: var(--text);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, Arial, sans-serif;
        font-size: 15px;
        line-height: 1.55;
    }

    .nicegui-content { padding: 0 !important; }

    /* ---------- Top header bar ---------- */
    .app-header {
        position: sticky;
        top: 0;
        z-index: 1000;
        background: rgba(13, 13, 16, 0.92);
        backdrop-filter: blur(8px);
        border-bottom: 1px solid var(--border-soft);
        padding: 14px 32px;
        display: flex;
        align-items: center;
        gap: 16px;
    }
    .app-header .app-title {
        font-size: 16px;
        font-weight: 600;
        letter-spacing: 0.2px;
    }
    .app-header .app-subtitle {
        margin-left: auto;
        color: var(--text-muted);
        font-size: 13px;
    }

    /* ---------- Page container ---------- */
    .page-shell {
        max-width: 1240px;
        margin: 0 auto;
        padding: 32px 32px 80px 32px;
    }

    .hero {
        margin-bottom: 28px;
    }
    .hero h1 {
        font-size: 28px;
        font-weight: 700;
        margin: 0 0 6px 0;
        letter-spacing: -0.2px;
    }
    .hero p {
        color: var(--text-muted);
        margin: 0;
        max-width: 720px;
    }

    /* ---------- Step cards ---------- */
    .step-card {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 24px 26px;
        margin-bottom: 20px;
    }
    .step-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 18px;
    }
    .step-number {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: var(--accent-bg);
        color: var(--accent-soft);
        font-weight: 700;
        font-size: 13px;
    }
    .step-title {
        font-size: 18px;
        font-weight: 600;
        color: var(--text);
    }
    .step-help {
        color: var(--text-muted);
        font-size: 14px;
        margin: -8px 0 16px 40px;
    }

    /* ---------- Upload tiles ---------- */
    .upload-tile-label {
        font-weight: 500;
        font-size: 14px;
        color: var(--text);
        margin-bottom: 6px;
    }
    .upload-tile-help {
        color: var(--text-muted);
        font-size: 13px;
        margin-bottom: 8px;
    }

    /* Quasar uploader overrides (file picker) */
    .q-uploader {
        background: var(--bg-elev) !important;
        border: 1px dashed var(--border) !important;
        border-radius: 10px !important;
        color: var(--text) !important;
        max-width: 100%;
        width: 100%;
    }
    .q-uploader__header {
        background: var(--accent) !important;
        color: #fff !important;
    }
    .q-uploader__title { color: #fff !important; font-weight: 500; }
    .q-uploader__subtitle { color: rgba(255,255,255,0.85) !important; }
    .q-uploader__list { background: var(--bg-elev) !important; }
    .q-uploader__file { background: var(--bg-card) !important; color: var(--text) !important; }

    /* Quasar select (grading dropdown) */
    .q-select,
    .q-select .q-field__native,
    .q-select .q-field__input,
    .q-select .q-field__append,
    .q-select .q-field__label {
        color: var(--text) !important;
    }
    .q-select .q-field__control:before,
    .q-select .q-field__control:after {
        border-color: var(--border) !important;
    }
    .q-select .q-field__control:hover:before {
        border-color: var(--accent-soft) !important;
    }

    /* Buttons */
    .q-btn {
        border-radius: 10px !important;
        font-weight: 500;
        text-transform: none !important;
        letter-spacing: 0 !important;
    }
    .q-btn .q-btn__content { padding: 4px 6px; }

    /* Textareas */
    textarea {
        background: var(--bg-elev) !important;
        color: var(--text) !important;
        border-radius: 8px;
    }

    /* Markdown / html result areas */
    .result-area {
        color: var(--text);
        font-size: 14.5px;
    }
    .result-area h1, .result-area h2, .result-area h3 {
        color: var(--text);
        margin-top: 0;
    }
    .result-area b, .result-area strong { color: var(--text); }
    .result-area p { color: var(--text); }
    .result-area code {
        background: var(--bg-elev);
        padding: 1px 6px;
        border-radius: 4px;
        font-size: 13px;
    }

    /* Empty state */
    .empty-state {
        color: var(--text-muted);
        font-style: italic;
        font-size: 14px;
        padding: 8px 0;
    }

    /* Highlight legend (now inline, not floating) */
    .legend {
        display: inline-flex;
        gap: 18px;
        align-items: center;
        background: var(--bg-elev);
        border: 1px solid var(--border-soft);
        border-radius: 8px;
        padding: 8px 14px;
        font-size: 13px;
        color: var(--text-muted);
    }
    .legend .swatch {
        display: inline-block;
        width: 14px;
        height: 10px;
        border-radius: 2px;
        margin-right: 6px;
        vertical-align: middle;
    }
    .legend .swatch.yellow { background: var(--hl-yellow); }
    .legend .swatch.blue   { background: var(--hl-blue); }

    /* Progress card shown while analysing */
    .progress-card {
        display: flex;
        align-items: center;
        gap: 14px;
        background: var(--accent-bg);
        border: 1px solid rgba(161, 0, 255, 0.35);
        border-radius: 10px;
        padding: 14px 18px;
        color: var(--text);
        font-size: 14px;
    }

    /* PDF viewer wrapper */
    .pdf-viewer {
        background: var(--bg-elev);
        border: 1px solid var(--border-soft);
        border-radius: 10px;
        overflow: hidden;
    }
    .pdf-viewer iframe { display: block; width: 100%; height: 760px; border: 0; }

    /* Small utility */
    .row-gap-16 > * + * { margin-left: 16px; }

    /* Quasar notify polish */
    .q-notification { font-family: inherit; }

    /* Pre-save confirmation dialog */
    .q-dialog__backdrop { background: rgba(0, 0, 0, 0.62) !important; }
    .save-confirm-card {
        background: var(--bg-card) !important;
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 26px 28px !important;
        color: var(--text);
        min-width: 360px;
        max-width: 460px;
        box-shadow: 0 18px 40px rgba(0, 0, 0, 0.55);
    }
    .save-confirm-title {
        font-size: 17px;
        font-weight: 600;
        color: var(--text);
        margin-bottom: 8px;
    }
    .save-confirm-body {
        color: var(--text-muted);
        font-size: 14px;
        line-height: 1.55;
        max-width: 380px;
    }
</style>
''')

class GradeEntry(BaseModel):
    course: Optional[str] = None
    grade: Optional[float] = None
    result: Optional[str] = None
    ects: Optional[float] = None
    assessment_date: Optional[str] = None

class CandidateInfo(BaseModel):
    name: Optional[str] = None
    bachelor_gpa: Optional[list[str]] = None
    bachelor_gpa_snippets: Optional[list[str]] = None
    bachelor_grades: Optional[list[GradeEntry]] = None
    bachelor_education: Optional[list[str]] = None
    bachelor_education_snippets: Optional[list[str]] = None
    master_gpa: Optional[list[str]] = None
    master_gpa_snippets: Optional[list[str]] = None
    master_grades: Optional[list[GradeEntry]] = None
    transcript_grades: Optional[list[GradeEntry]] = None
    master_education: Optional[list[str]] = None
    master_education_snippets: Optional[list[str]] = None
    master_university: Optional[list[str]] = None
    master_program: Optional[list[str]] = None
    bachelor_university: Optional[list[str]] = None
    bachelor_program: Optional[list[str]] = None
    exchange_gpa: Optional[list[str]] = None
    exchange_gpa_snippets: Optional[list[str]] = None
    exchange_education: Optional[list[str]] = None
    exchange_education_snippets: Optional[list[str]] = None
    student_job: Optional[list[str]] = None
    student_job_snippets: Optional[list[str]] = None
    extracurricular: Optional[list[str]] = None
    international_experience: Optional[list[str]] = None
    nordic_language: Optional[str] = None
    nordic_language_snippet: Optional[str] = None
    gender: Optional[str] = None
    gender_snippet: Optional[str] = None
    persona: Optional[str] = None
    persona_reason: Optional[str] = None
    persona_snippets: Optional[list[str]] = None

class SimilarProgramScore(BaseModel):
    score: int
    reason: str

class StudentJobScore(BaseModel):
    score: int

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
    extract_evidence_of_attending_Danmarks_Tekniske_Universitet_on_a_tech_management_track: list[str]
    boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet_on_a_tech_management_track: bool
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

    def university_score(self) -> int:
        if any([
            self.boolean_classify_evidence_of_attending_Copenhagen_Business_School,
            self.boolean_classify_evidence_of_attending_Aarhus_university,
            self.boolean_classify_evidence_of_attending_University_of_Copenhagen,
            self.boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet_on_a_tech_management_track,
            self.boolean_classify_evidence_of_attending_Lund_University,
        ]):
            return 5

        return 0

    def program_score(self) -> int:
        if any([
            self.boolean_classify_business_admin_digital_business_cbs,
            self.boolean_classify_strategy_org_leadership_cbs,
            self.boolean_classify_innovation_business_dev_cbs,
            self.boolean_classify_economics_ku_au,
            self.boolean_classify_international_business_cbs_au,
            self.boolean_classify_digital_business_management_au,
        ]):
            return 5
        elif self.boolean_classify_similar_relevant_program:
            return 4

        return 0

    def keyword_score(self) -> int:
        if any([
            self.boolean_classify_strategic_problem_solving,
            self.boolean_classify_digital_transformation,
            self.boolean_classify_technology_enabled_business_models,
            self.boolean_classify_c_level_advisory,
            self.boolean_classify_operating_model_design,
            self.boolean_classify_business_case_development,
            self.boolean_classify_technology_strategy,
        ]):
            return 5

        return 0

    def final_score(self) -> int:
        return (
            self.university_score() +
            self.program_score() +
            self.keyword_score()
        )

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

    def university_score(self) -> int:
        if any([
            self.boolean_classify_evidence_of_attending_Copenhagen_Business_School,
            self.boolean_classify_evidence_of_attending_Aarhus_university,
            self.boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet,
            self.boolean_classify_evidence_of_attending_IT_University_of_Copenhagen,
            self.boolean_classify_evidence_of_attending_Aalborg_University,
        ]):
            return 5

        return 0

    def program_score(self) -> int:
        if any([
            self.boolean_classify_digital_business_management_au,
            self.boolean_classify_business_administration_and_information_systems_cbs,
            self.boolean_classify_management_of_innovation_and_business_development_cbs,
            self.boolean_classify_operations_and_supply_chain_programs,
            self.boolean_classify_engineering_with_business_specialization_dtu,
        ]):
            return 5
        elif self.boolean_classify_similar_relevant_program:
            return 4

        return 0

    def keyword_score(self) -> int:
        if any([
            self.boolean_classify_program_management,
            self.boolean_classify_execution_and_value_realization,
            self.boolean_classify_governance_and_operating_models,
            self.boolean_classify_change_enablement,
            self.boolean_classify_stakeholder_coordination,
            self.boolean_classify_implementation_roadmap_design,
        ]):
            return 5

        return 0

    def final_score(self) -> int:
        return (
            self.university_score() +
            self.program_score() +
            self.keyword_score()
        )


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

    def university_score(self) -> int:
        if any([
            self.boolean_classify_evidence_of_attending_IT_University_of_Copenhagen,
            self.boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet,
            self.boolean_classify_evidence_of_attending_Aalborg_University,
            self.boolean_classify_evidence_of_attending_Copenhagen_Business_School_on_data_or_digital_tracks,
            self.boolean_classify_evidence_of_attending_University_of_Copenhagen_on_quantitative_or_data_tracks,
        ]):
            return 5

        return 0

    def program_score(self) -> int:
        if any([
            self.boolean_classify_business_intelligence_cbs_or_au,
            self.boolean_classify_business_administration_and_data_science_cbs,
            self.boolean_classify_computer_science_itu_dtu_aau,
            self.boolean_classify_data_science_programs,
            self.boolean_classify_analytics_and_information_systems_programs,
        ]):
            return 5
        elif self.boolean_classify_similar_relevant_program:
            return 4

        return 0

    def keyword_score(self) -> int:
        if any([
            self.boolean_classify_data_literacy,
            self.boolean_classify_ai_enablement,
            self.boolean_classify_business_it_alignment,
            self.boolean_classify_systems_and_architecture_understanding,
            self.boolean_classify_analytics_interpretation,
            self.boolean_classify_use_case_development,
            self.boolean_classify_genai_applications,
        ]):
            return 5

        return 0

    def final_score(self) -> int:
        return (
            self.university_score() +
            self.program_score() +
            self.keyword_score()
        )


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

    def university_score(self) -> int:
        if any([
            self.boolean_classify_evidence_of_attending_University_of_Copenhagen_on_life_science_or_public_sector_tracks,
            self.boolean_classify_evidence_of_attending_Danmarks_Tekniske_Universitet_on_energy_or_engineering_tracks,
            self.boolean_classify_evidence_of_attending_Copenhagen_Business_School_on_financial_services_or_commercial_tracks,
            self.boolean_classify_evidence_of_attending_Aarhus_university,
            self.boolean_classify_evidence_of_attending_Lund_University,
        ]):
            return 5
        elif any([
            self.boolean_classify_evidence_of_attending_Roskilde_University,
            self.boolean_classify_evidence_of_attending_Southern_Denmark_University,
            self.boolean_classify_evidence_of_attending_Aalborg_University,
        ]):
            return 4

        return 0

    def program_score(self) -> int:
        if any([
            self.boolean_classify_applied_economics_and_finance_ku,
            self.boolean_classify_finance_and_strategic_management_cbs,
            self.boolean_classify_life_science_or_biomed_programs_ku_dtu_lund,
            self.boolean_classify_energy_and_engineering_programs_dtu_aau,
            self.boolean_classify_supply_chain_and_operations_programs,
            self.boolean_classify_public_policy_or_political_science_ku_au,
        ]):
            return 5
        elif self.boolean_classify_similar_relevant_program:
            return 4

        return 0

    def keyword_score(self) -> int:
        if any([
            self.boolean_classify_industry_expertise,
            self.boolean_classify_regulatory_knowledge,
            self.boolean_classify_value_chain_insight,
            self.boolean_classify_commercial_awareness,
            self.boolean_classify_sector_strategy,
            self.boolean_classify_domain_specific_analytics,
            self.boolean_classify_strategy_execution,
        ]):
            return 5

        return 0

    def final_score(self) -> int:
        return (
            self.university_score() +
            self.program_score() +
            self.keyword_score()
        )

#definer funktioner for at læse pdf
norm = lambda t: re.sub(r"[^\w\s:/().,+-]", "", re.sub(r"\s+", " ", t.lower())).strip()

def extract_word_tokens(b):
    column_order = []
    line_order = []
    with pdfplumber.open(BytesIO(b)) as pdf:
        for p, page in enumerate(pdf.pages):
            words = page.extract_words()
            line_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
            mid = page.width * 0.4

            left = [w for w in words if w["x0"] < mid]
            right = [w for w in words if w["x0"] >= mid]

            # treat as 2 columns
            if len(left) > len(words) * 0.2 and len(right) > len(words) * 0.2:
                words = sorted(left, key=lambda w: (w["top"], w["x0"])) + sorted(right, key=lambda w: (w["top"], w["x0"]))
            
            # treat as 1 column
            else:
                words = sorted(words, key=lambda w: (w["top"], w["x0"]))

            column_order += [
                {
                    "text": w["text"],
                    "norm": norm(w["text"]),
                    "page": p,
                    "x0": w["x0"],
                    "top": w["top"],
                    "x1": w["x1"],
                    "bottom": w["bottom"],
                }
                for w in words
            ]
            line_order += [
                {
                    "text": w["text"],
                    "norm": norm(w["text"]),
                    "page": p,
                    "x0": w["x0"],
                    "top": w["top"],
                    "x1": w["x1"],
                    "bottom": w["bottom"],
                }
                for w in line_words
            ]
    return [column_order, line_order]

def best_window_match(f, t, s=2):
    if isinstance(f, list):
        f = " ".join(f)

    trg = " ".join(norm(w) for w in f.split() if norm(w))
    if not trg: 
        return None

    n, best = len(trg.split()), None

    for z in range(max(1, n-s), n+s+1):
        for i in range(len(t)-z+1):
            w = t[i:i+z]

            if len({x["page"] for x in w}) > 1: 
                continue

            c = " ".join(x["norm"] for x in w if x["norm"])
            sc = Levenshtein.distance(c, trg) / max(len(trg), 1)

            if not best or sc < best["score"]:
                best = {"score": sc, "tokens": w}

    return best if best and best["score"] <= 0.25 else None

def collect_text_matches(fields, tokens, clean_targets=True):
    matches = []
    token_sets = tokens if tokens and isinstance(tokens[0], list) else [tokens]
    for f in fields:
        if not f:
            continue

        target = clean_field(f) if clean_targets else f
        candidates = [best_window_match(target, token_set) for token_set in token_sets]
        candidates = [candidate for candidate in candidates if candidate]
        m = min(candidates, key=lambda candidate: candidate["score"]) if candidates else None

        print("TARGET:", target)
        print("MATCH:", " ".join(tok["text"] for tok in m["tokens"]) if m else None)
        print("----")

        if m:
            matches.append(m)
        
    return matches

def collect_evidence_matches(npa, tokens):
    exchange_lines = []
    for snippet in npa.exchange_education_snippets or []:
        exchange_lines.extend(snippet.splitlines())

    fields = [
        *(npa.master_gpa or []),
        *(npa.bachelor_gpa or []),
        *(npa.student_job or []),
        *(npa.extracurricular or []),
        *exchange_lines,
        *(npa.exchange_education or []),
        *(npa.international_experience or []),
        npa.nordic_language_snippet,
    ]

    return collect_text_matches(fields, tokens)

def is_usable_evidence(value):
    if not isinstance(value, str):
        return False

    cleaned = norm(value)
    return bool(cleaned) and cleaned not in {"n/a", "na", "none", "null", "not mentioned", "not applicable"}

def model_field_names(model):
    if hasattr(model, "model_fields"):
        return model.model_fields.keys()
    return model.__fields__.keys()

def matching_boolean_field(evidence_field, field_names):
    suffix = evidence_field.removeprefix("extract_evidence_of_")
    candidates = [
        f"boolean_classify_{suffix}",
        f"boolean_classify_evidence_of_{suffix}",
    ]

    if suffix.startswith("bachelor_"):
        candidates.append(f"boolean_classify_evidence_of_{suffix.removeprefix('bachelor_')}")

    for candidate in candidates:
        if candidate in field_names:
            return candidate

    return None

def collect_persona_evidence(persona, npa=None):
    field_names = set(model_field_names(persona))
    master_text = norm(join_values(npa.master_education)) if npa else ""
    evidence = []

    for field_name in field_names:
        if not field_name.startswith("extract_evidence_of_"):
            continue

        boolean_field = matching_boolean_field(field_name, field_names)
        if not boolean_field or not getattr(persona, boolean_field, False):
            continue

        education_field = (
            boolean_field.startswith("boolean_classify_evidence_of_attending")
            or boolean_field in PROGRAM_BOOLEAN_FIELDS
        )

        value = getattr(persona, field_name, None)
        values = value if isinstance(value, list) else [value]

        for v in values:
            if not is_usable_evidence(v):
                continue
            v_text = norm(v)
            if len(v_text.split()) < 2:
                continue
            if education_field and master_text and v_text not in master_text and master_text not in v_text:
                continue
            evidence.append(v)

    return evidence

PROGRAM_FIELDS = {
    "TechStrategist": [
        "boolean_classify_business_admin_digital_business_cbs",
        "boolean_classify_strategy_org_leadership_cbs",
        "boolean_classify_innovation_business_dev_cbs",
        "boolean_classify_economics_ku_au",
        "boolean_classify_international_business_cbs_au",
        "boolean_classify_digital_business_management_au",
    ],
    "TransformationOrchestratorPMO": [
        "boolean_classify_digital_business_management_au",
        "boolean_classify_business_administration_and_information_systems_cbs",
        "boolean_classify_management_of_innovation_and_business_development_cbs",
        "boolean_classify_operations_and_supply_chain_programs",
        "boolean_classify_engineering_with_business_specialization_dtu",
    ],
    "TechTranslator": [
        "boolean_classify_business_intelligence_cbs_or_au",
        "boolean_classify_business_administration_and_data_science_cbs",
        "boolean_classify_computer_science_itu_dtu_aau",
        "boolean_classify_data_science_programs",
        "boolean_classify_analytics_and_information_systems_programs",
    ],
    "AppliedTechnicalSpecialist": [
        "boolean_classify_computer_science",
        "boolean_classify_applied_mathematics",
        "boolean_classify_physics",
        "boolean_classify_statistics",
        "boolean_classify_engineering_software_systems_industrial_or_energy",
        "boolean_classify_machine_learning_and_ai_programs",
    ],
    "IndustryFunctionalSpecialist": [
        "boolean_classify_applied_economics_and_finance_ku",
        "boolean_classify_finance_and_strategic_management_cbs",
        "boolean_classify_life_science_or_biomed_programs_ku_dtu_lund",
        "boolean_classify_energy_and_engineering_programs_dtu_aau",
        "boolean_classify_supply_chain_and_operations_programs",
        "boolean_classify_public_policy_or_political_science_ku_au",
    ],
}

PROGRAM_BOOLEAN_FIELDS = set().union(*PROGRAM_FIELDS.values())

EXCLUDED_PERSONA_BOOLEAN_PREFIXES = (
    "boolean_classify_evidence_of_attending",
)

EXCLUDED_PERSONA_BOOLEANS = PROGRAM_BOOLEAN_FIELDS | {
    "boolean_classify_similar_relevant_program",
}

def is_persona_reasoning_boolean(name):
    return (
        name.startswith("boolean_")
        and not name.startswith(EXCLUDED_PERSONA_BOOLEAN_PREFIXES)
        and name not in EXCLUDED_PERSONA_BOOLEANS
    )

def true_boolean_count(model):
    return sum(
        1
        for field_name in model_field_names(model)
        if is_persona_reasoning_boolean(field_name) and getattr(model, field_name, False) is True
    )

MASTER_UNIVERSITY_TIERS = {
    "TechStrategist": {
        "primary": [
            "cbs",
            "copenhagen business school",
            "aarhus university",
            "university of copenhagen",
            "københavns universitet",
        ],
        "secondary": [
            "dtu",
            "danmarks tekniske universitet",
            "technical university of denmark",
            "lund university",
        ],
    },
    "TransformationOrchestratorPMO": {
        "primary": [
            "cbs",
            "copenhagen business school",
            "aarhus university",
            "dtu",
            "danmarks tekniske universitet",
            "technical university of denmark",
        ],
        "secondary": [
            "itu",
            "it university of copenhagen",
            "aalborg university",
            "aau",
        ],
    },
    "TechTranslator": {
        "primary": [
            "itu",
            "it university of copenhagen",
            "dtu",
            "danmarks tekniske universitet",
            "technical university of denmark",
            "aalborg university",
            "aau",
        ],
        "secondary": [
            "cbs",
            "copenhagen business school",
            "university of copenhagen",
            "københavns universitet",
        ],
    },
    "AppliedTechnicalSpecialist": {
        "primary": [
            "dtu",
            "danmarks tekniske universitet",
            "technical university of denmark",
            "itu",
            "it university of copenhagen",
            "aalborg university",
            "aau",
            "university of copenhagen",
            "københavns universitet",
            "lund university",
        ],
        "secondary": [
            "sdu",
            "southern denmark university",
            "university of southern denmark",
            "international technical",
        ],
    },
    "IndustryFunctionalSpecialist": {
        "primary": [
            "university of copenhagen",
            "københavns universitet",
            "dtu",
            "danmarks tekniske universitet",
            "technical university of denmark",
            "cbs",
            "copenhagen business school",
            "aarhus university",
            "lund university",
        ],
        "secondary": [
            "ruc",
            "roskilde university",
            "sdu",
            "southern denmark university",
            "university of southern denmark",
            "aalborg university",
            "aau",
        ],
    },
}

MASTER_PROGRAM_TERMS = {
    "TechStrategist": {
        "key": [
            "business administration and digital business",
            "strategy organization leadership",
            "management of innovation and business development",
            "economics",
            "international business",
            "digital business management",
        ],
        "similar": [
            "digital business",
            "business development",
            "innovation",
            "strategy",
        ],
    },
    "TransformationOrchestratorPMO": {
        "key": [
            "digital business management",
            "business administration and information systems",
            "management of innovation and business development",
            "operations",
            "supply chain",
            "engineering with business",
        ],
        "similar": [
            "program management",
            "project management",
            "operations management",
        ],
    },
    "TechTranslator": {
        "key": [
            "business intelligence",
            "business administration and data science",
            "computer science",
            "data science",
            "analytics",
            "information systems",
        ],
        "similar": [
            "business analytics",
            "information management",
            "digital business",
        ],
    },
    "AppliedTechnicalSpecialist": {
        "key": [
            "computer science",
            "applied mathematics",
            "physics",
            "statistics",
            "engineering",
            "machine learning",
            "artificial intelligence",
        ],
        "similar": [
            "mathematics",
            "quantitative",
            "technical",
            "data science",
        ],
    },
    "IndustryFunctionalSpecialist": {
        "key": [
            "applied economics and finance",
            "finance and strategic management",
            "life science",
            "biomed",
            "energy",
            "supply chain",
            "operations",
            "public policy",
            "political science",
        ],
        "similar": [
            "economics",
            "finance",
            "public administration",
        ],
    },
}

def score_master_university(master_university, persona_name):
    master_text = norm(join_values(master_university))
    tiers = MASTER_UNIVERSITY_TIERS.get(persona_name, {})

    if any(university in master_text for university in tiers.get("primary", [])):
        return 2
    if any(university in master_text for university in tiers.get("secondary", [])):
        return 1

    return 0

async def score_master_program(master_program, persona_name):
    master_text = norm(join_values(master_program))
    terms = MASTER_PROGRAM_TERMS.get(persona_name, {})

    if any(term in master_text for term in terms.get("key", [])):
        return 2

    if not master_text:
        return 0

    result = await client.beta.chat.completions.parse(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": (
                "You score whether a candidate's master program is similar to a persona's target master programs. "
                "Return score 1 only if the program is clearly close in academic subject matter. "
                "Return score 0 if it is generic, unclear, or only loosely related."
            )},
            {"role": "user", "content": (
                f"Persona: {persona_name}\n"
                f"Candidate master program: {join_values(master_program)}\n"
                f"Target programs: {', '.join(terms.get('key', []))}"
            )},
        ],
        response_format=SimilarProgramScore,
        temperature=0,
    )

    return 1 if result.choices[0].message.parsed.score == 1 else 0

def true_boolean_labels(persona):
    return [
        name.replace("boolean_classify_evidence_of_", "")
            .replace("boolean_classify_", "")
            .replace("_", " ")
        for name in model_field_names(persona)
        if is_persona_reasoning_boolean(name) and getattr(persona, name) is True
    ]

def format_persona_fit_report(persona_name, persona, true_count, master_university_weight, program_weight, persona_score):
    labels = true_boolean_labels(persona)
    label_lines = "\n".join(f"- {label}" for label in labels) or "- None"

    return (
        f"#### {persona_name}\n"
        f"**Persona score: {persona_score}** "
        f"({true_count} true booleans + {master_university_weight} master university + {program_weight} program)\n\n"
        f"**True booleans**\n"
        f"{label_lines}"
    )

def first_float(values):
    if isinstance(values, str):
        values = [values]

    for value in values or []:
        match = re.search(r"\d+(?:[,.]\d+)?", str(value))
        if match:
            return float(match.group(0).replace(",", "."))

    return None

def convert_to_danish_grade(grade, grading_system="Danish 7-point", result=None):
    if grading_system == "ECTS A-F":
        letter = str(result or "").strip().upper()
        return {
            "A": 12,
            "B": 10,
            "C": 7,
            "D": 4,
            "E": 2,
            "FX": 0,
            "F": -3,
        }.get(letter)

    if grade is None:
        return None

    conversion_ranges = {
        "Danish 7-point": [
            (12, 12, 12),
            (10, 10, 10),
            (7, 7, 7),
            (4, 4, 4),
            (2, 2, 2),
            (0, 0, 0),
            (-3, -3, -3),
        ],
        "Italian 18-30": [
            (30, 30, 12),
            (27, 29.99, 10),
            (24, 26.99, 7),
            (21, 23.99, 4),
            (18, 20.99, 2),
            (0, 17.99, 0),
        ],
        "German 1-5": [
            (1, 1.5, 12),
            (1.51, 2, 10),
            (2.01, 2.5, 7),
            (2.51, 3.5, 4),
            (3.51, 4, 2),
            (4.01, 5, 0),
        ],
        "French 0-20": [
            (18, 20, 12),
            (16, 17.99, 10),
            (14, 15.99, 7),
            (12, 13.99, 4),
            (10, 11.99, 2),
            (0, 9.99, 0),
        ],
        "Dutch 1-10": [
            (9, 10, 12),
            (8, 8.99, 10),
            (7, 7.99, 7),
            (6, 6.99, 4),
            (5.5, 5.99, 2),
            (1, 5.49, 0),
        ],
        "Spanish 0-10": [
            (9, 10, 12),
            (8, 8.99, 10),
            (7, 7.99, 7),
            (6, 6.99, 4),
            (5, 5.99, 2),
            (0, 4.99, 0),
        ],
        "US GPA 0-4": [
            (3.7, 4, 12),
            (3.3, 3.69, 10),
            (2.7, 3.29, 7),
            (2, 2.69, 4),
            (1, 1.99, 2),
            (0, 0.99, 0),
        ],
    }

    for min_grade, max_grade, danish_grade in conversion_ranges.get(grading_system, []):
        if min_grade <= grade <= max_grade:
            return danish_grade

    return None

def calculate_gpa(grades, grading_system="Danish 7-point"):
    valid_grades = [
        (g, convert_to_danish_grade(g.grade, grading_system, g.result))
        for g in (grades or [])
    ]
    valid_grades = [(g, danish_grade) for g, danish_grade in valid_grades if danish_grade is not None]

    if not valid_grades:
        return None

    weighted_grades = [(g, danish_grade) for g, danish_grade in valid_grades if g.ects is not None]

    if weighted_grades:
        total_ects = sum(g.ects for g, _ in weighted_grades)
        if total_ects:
            return round(sum(danish_grade * g.ects for g, danish_grade in weighted_grades) / total_ects, 2)

    return round(sum(danish_grade for _, danish_grade in valid_grades) / len(valid_grades), 2)

def format_grade_value(grade):
    if grade is None:
        return ""

    if isinstance(grade, float) and grade.is_integer():
        return str(int(grade))

    return str(grade)

def format_conversion_report_section(label, grades, grading_system, gpa):
    if not grades:
        return f"#### {label}\nNo transcript grades extracted."

    lines = [f"#### {label}"]

    for grade in grades:
        original = format_grade_value(grade.grade) or (grade.result or "No numeric grade")
        converted = convert_to_danish_grade(grade.grade, grading_system, grade.result)
        converted_text = "not included in GPA" if converted is None else f"{converted} Danish"
        ects = format_grade_value(grade.ects) or "no ECTS"
        course = grade.course or "Unnamed course"
        lines.append(f"- {course}: {original} -> {converted_text}, {ects} ECTS")

    if gpa is not None:
        lines.append(f"**{label} GPA: {gpa}**")

    return "\n".join(lines)

def format_conversion_report(grading_system, bachelor_grades, master_grades, bachelor_gpa, master_gpa):
    if grading_system == "Danish 7-point" or (not bachelor_grades and not master_grades):
        return ""

    bachelor_converted_grades = [
        convert_to_danish_grade(grade.grade, grading_system, grade.result)
        for grade in bachelor_grades or []
    ]
    bachelor_converted_grades = [grade for grade in bachelor_converted_grades if grade is not None]
    bachelor_total_ects = sum(grade.ects for grade in bachelor_grades or [] if grade.ects is not None)
    master_converted_grades = [
        convert_to_danish_grade(grade.grade, grading_system, grade.result)
        for grade in master_grades or []
    ]
    master_converted_grades = [grade for grade in master_converted_grades if grade is not None]
    master_total_ects = sum(grade.ects for grade in master_grades or [] if grade.ects is not None)

    intro = [
        "### Grade conversion",
        f"**Selected system:** {grading_system}",
        "**Converted to:** Danish 7-point scale",
    ]

    if master_grades:
        intro.extend([
            f"**Master GPA after conversion:** {master_gpa}",
            f"**Master grades converted:** {len(master_converted_grades)}",
        ])
        if master_total_ects:
            intro.append(f"**Master ECTS included:** {format_grade_value(master_total_ects)}")

    if bachelor_grades:
        intro.extend([
            f"**Bachelor GPA after conversion:** {bachelor_gpa}",
            f"**Bachelor grades converted:** {len(bachelor_converted_grades)}",
        ])
        if bachelor_total_ects:
            intro.append(f"**Bachelor ECTS included:** {format_grade_value(bachelor_total_ects)}")

    return "\n".join([line + "  " for line in intro])

def is_au_transcript(text):
    return "aarhus university" in norm(text or "")

def uses_au_extended_bachelor_cutoff(text):
    text = norm(text or "")
    return any([
        "beng" in text,
        "bachelor of engineering" in text,
        "diplomingenior" in text,
        "diplomingeniør" in text,
        "bachelor in software development" in text,
        "bachelors degree programme in software development" in text,
        "bachelor's degree programme in software development" in text,
    ])

def parse_transcript_date(value):
    if not value:
        return None

    value = str(value).strip()

    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass

    return None

def split_au_transcript_grades(grades, text):
    dated_grades = [
        (grade, parse_transcript_date(grade.assessment_date))
        for grade in grades or []
    ]
    dated_grades = [(grade, date) for grade, date in dated_grades if date is not None]

    if not dated_grades:
        return grades or [], []

    start_date = min(date for _, date in dated_grades)
    cutoff_date = start_date + relativedelta(months=42 if uses_au_extended_bachelor_cutoff(text) else 36)
    bachelor_grades = []
    master_grades = []

    for grade in grades or []:
        assessment_date = parse_transcript_date(grade.assessment_date)

        if assessment_date is None or assessment_date <= cutoff_date:
            bachelor_grades.append(grade)
        else:
            master_grades.append(grade)

    print("AU transcript start date:", start_date.strftime("%d.%m.%Y"))
    print("AU transcript cutoff date:", cutoff_date.strftime("%d.%m.%Y"))
    return bachelor_grades, master_grades

def print_grade_debug(label, grades, gpa):
    print(f"{label} GPA:", gpa)
    for grade in grades or []:
        print(
            f"{label} grade row:",
            "course=", grade.course,
            "| grade=", grade.grade,
            "| result=", grade.result,
            "| ects=", grade.ects,
        )

def is_applied_technical(persona_name):
    return persona_name == "AppliedTechnicalSpecialist"

def score_master_grade(master_gpa, persona_name=None):
    gpa = first_float(master_gpa)
    if gpa is None:
        return 2
    if not is_applied_technical(persona_name):
        if gpa >= 10.5:
            return 5
        if gpa >= 9:
            return 4
        if gpa >= 7.5:
            return 3
        return 1
    if gpa >= 8.5:
        return 5
    if gpa >= 7.5:
        return 4
    if gpa >= 6:
        return 3
    return 1

def score_bachelor_grade(bachelor_gpa, persona_name=None):
    gpa = first_float(bachelor_gpa)
    if gpa is None:
        return 2
    if not is_applied_technical(persona_name):
        if gpa >= 10:
            return 5
        if gpa >= 9:
            return 4
        if gpa >= 7:
            return 3
        return 1
    if gpa >= 8:
        return 5
    if gpa > 7:
        return 4
    if gpa >= 5.5:
        return 3
    return 1

def score_student_job(student_job, persona_name=None):
    text = norm(" ".join(student_job or []))
    if not text:
        return 0 if is_applied_technical(persona_name) else 1

    if is_applied_technical(persona_name):
        strong_terms = [
            "software development",
            "software developer",
            "backend",
            "back end",
            "operations",
            "system development",
            "systems development",
        ]
        if any(term in text for term in strong_terms):
            return 4
        return 1

    score_5_terms = [
        "consulting",
        "internship",
        "internal strategy",
        "strategy consulting",
        "advisory",
        "ministry",
        "styrelse",
        "styrelser",
        "accenture",
        "mckinsey",
        "bain",
        "bcg",
        "deloitte",
        "pwc",
        "ey",
        "kpmg",
    ]
    score_4_terms = [
        "multinational",
        "business and technology",
        "business technology",
        "digital transformation",
        "strategy analyst",
        "student assistant",
        "student analyst",
        "student consultant",
        "project student",
        "project assistant",
        "business development",
        "pmo",
        "commercial excellence",
        "finance",
        "fp&a",
        "sourcing",
        "operations",
        "semco maritime",
        "nordea",
        "arla",
        "vestas",
        "novo nordisk",
        "topsoe",
        "jysk",
        "columbus",
        "netcompany",
        "ørsted",
        "orsted",
    ]
    score_1_terms = [
        "cafe",
        "restaurant",
        "supermarket",
        "service assistant",
    ]

    if any(term in text for term in score_5_terms):
        return 5
    if any(term in text for term in score_4_terms):
        return 4
    if any(term in text for term in score_1_terms):
        return 1
    return 1

async def score_student_job_llm(student_job, persona_name=None):
    if not is_applied_technical(persona_name):
        return score_student_job(student_job, persona_name)

    text = join_values(student_job)
    if not text:
        return 0

    result = await client.beta.chat.completions.parse(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": (
                "Score the candidate's student job/internship for Applied Technical Specialist. "
                "Return only the integer score. "
                "4 = hands-on technical work such as software development, backend operations, "
                "system development, applied AI, data science, machine learning, programming, "
                "engineering, quantitative modelling, technical research, or similar. "
                "1 = other student jobs. "
                "0 = no student job/internship mentioned."
            )},
            {"role": "user", "content": text},
        ],
        response_format=StudentJobScore,
        temperature=0,
    )

    score = result.choices[0].message.parsed.score
    return score if score in {0, 1, 4} else 1

def score_extracurricular(extracurricular, persona_name=None):
    if is_applied_technical(persona_name):
        return binary_score(extracurricular)

    text = norm(" ".join(extracurricular or []))
    if not text:
        return 0

    lead_terms = [
        "lead",
        "leader",
        "chair",
        "president",
        "board",
        "founder",
    ]
    return 2 if any(term in text for term in lead_terms) else 1

def binary_score(value):
    if isinstance(value, list):
        return 1 if any(is_usable_evidence(v) for v in value) else 0

    return 1 if is_usable_evidence(value) else 0

def score_nordic_language(value):
    text = norm(value or "")
    return 1 if any(language in text for language in ["danish", "dansk", "swedish", "svensk", "norwegian", "norsk"]) else 0

def score_gender(value):
    text = norm(value or "")
    return 1 if text in {"f", "female", "woman", "kvinde"} else 0

async def build_scorecard(npa, persona_score, persona_name=None):
    scorecard = {
        "master_grade_score": score_master_grade(npa.master_gpa, persona_name),
        "bachelor_grade_score": score_bachelor_grade(npa.bachelor_gpa, persona_name),
        "student_job_score": await score_student_job_llm(npa.student_job, persona_name),
        "extracurricular_score": score_extracurricular(npa.extracurricular, persona_name),
        "international_experience_score": binary_score(npa.international_experience or npa.exchange_education),
        "nordic_language_score": score_nordic_language(npa.nordic_language),
        "gender_score": 0,
    }

    scorecard["final_score"] = persona_score + sum(scorecard.values())
    return scorecard

def join_values(values):
    if isinstance(values, list):
        return ", ".join(str(value) for value in values if value)
    return values or ""

def combine_values(*values):
    combined = []

    for value in values:
        if isinstance(value, list):
            combined.extend(item for item in value if item)
        elif value:
            combined.append(value)

    return combined

def build_candidate_row(npa, selected_persona, persona_score, scorecard, persona_reason="",
                        second_selected_persona="", second_persona_score=""):
    international_experience = npa.international_experience or npa.exchange_education

    return {
        "name": npa.name or "",
        "master_education": join_values(npa.master_education),
        "master_university": join_values(npa.master_university),
        "master_program": join_values(npa.master_program),
        "master_gpa": join_values(npa.master_gpa),
        "master_grade_score": scorecard["master_grade_score"],
        "bachelor_education": join_values(npa.bachelor_education),
        "bachelor_university": join_values(npa.bachelor_university),
        "bachelor_program": join_values(npa.bachelor_program),
        "bachelor_gpa": join_values(npa.bachelor_gpa),
        "bachelor_grade_score": scorecard["bachelor_grade_score"],
        "international_experience": join_values(international_experience),
        "student_job": join_values(npa.student_job),
        "student_job_score": scorecard["student_job_score"],
        "extracurricular": join_values(npa.extracurricular),
        "extracurricular_score": scorecard["extracurricular_score"],
        "international_experience_score": scorecard["international_experience_score"],
        "nordic_language": npa.nordic_language or "",
        "nordic_language_score": scorecard["nordic_language_score"],
        "gender": npa.gender or "unknown",
        "gender_score": scorecard["gender_score"],
        "selected_persona": selected_persona,
        "persona_reason": persona_reason,
        "persona_score": persona_score,
        "final_score": scorecard["final_score"],
        "second_selected_persona": second_selected_persona,
        "second_persona_score": second_persona_score,
        "verification_initial": "",
        "comment": ""
    }

def format_scorecard_report(row):
    return (
        f"### Selected persona\n"
        f"**{row['selected_persona']}** (persona score: {row['persona_score']})\n\n"
        f"### Candidate scorecard\n"
        f"- Master grade ({row['master_gpa'] or 'N/A'}): {'N/A' if row['master_grade_score'] == 2 else row['master_grade_score']}\n"
        f"- Bachelor grade ({row['bachelor_gpa'] or 'N/A'}): {'N/A' if row['bachelor_grade_score'] == 2 else row['bachelor_grade_score']}\n"
        f"- Student job/internship: {row['student_job_score']}\n"
        f"- Extracurricular: {row['extracurricular_score']}\n"
        f"- International experience: {row['international_experience_score']}\n"
        f"- Danish/Swedish/Norwegian: {row['nordic_language_score']}\n"
        f"- Gender: {row['gender_score']}\n"
        f"- Persona score: {row['persona_score']}\n"
        f"- **Final score: {row['final_score']}**"
    )

SHEET_HEADERS = [
    "name",
    "master_university",
    "master_program",
    "master_gpa",
    "master_grade_score",
    "bachelor_university",
    "bachelor_program",
    "bachelor_gpa",
    "bachelor_grade_score",
    "international_experience",
    "international_experience_score",
    "student_job",
    "student_job_score",
    "extracurricular",
    "extracurricular_score",
    "nordic_language",
    "nordic_language_score",
    "gender",
    "selected_persona",
    "persona_reason",
    "persona_score",
    "final_score",
    "second_selected_persona",   # column W
    "second_persona_score",      # column X
    "verification_initial",
    "comment"
    # column Y (verification) and column Z (Comment) are left out on purpose:
    # they are filled in by the human reviewer, never written by the script.
]

def clean_field(f):
    if isinstance(f, list):
        f = " ".join(f)

    f = f.strip()

    if len(f.split()) <= 3 and not re.search(r"\d", f):
        return f.upper()

    # handle short fields like names
    if len(f.split()) <= 3:
        return f.upper()

    # keep GPA numbers
    m = re.search(r"\d+(?:\.\d+)?/12", f)
    if m:
        return m.group(0)

    # remove brackets (dates etc.)
    f = re.sub(r"\([^)]*\)", "", f)

    lower_f = f.lower()
    if "bachelor" in lower_f or lower_f.startswith("ba "):
        return "BA"
    if "master" in lower_f or "master’s" in lower_f or "master of" in lower_f:
        return "Master"
    
    return f.strip()

   
async def send_input():
    #analysis_clases = [IndustryFunctionalSpecialist]
    analysis_clases = [TechStrategist, TransformationOrchestratorPMO, TechTranslator, AppliedTechnicalSpecialist, IndustryFunctionalSpecialist]
    
    progress.visible = True
    label.visible = True
    label.text = "Extracting..."

    try:
        await asyncio.sleep(2)

        t = f"CV:\n{cv_raw_text}\n\nCOVER LETTER:\n{cover_raw_text}\n\nGRADE TRANSCRIPT:\n{transcript_raw_text}"

        print("TEXT LENGTH:", len(t)) 
        print("FULL TEXT:", repr(t))

        candidate_task = client.beta.chat.completions.parse(
            model="gpt-4.1",
            messages=[
                {"role": "system", "content": (
                    "Extract all bachelor entries and bachelor GPA values from the text. "
                    "Extract all master entries and master GPA values from the text. "
                    "Extract master_university, master_program, bachelor_university, and bachelor_program separately. Do not mix bachelor and master fields. "
                    "Extract all international experience entries from the text as separate entries (if any). "
                    "International experience may be written as 'exchange', 'exchange semester', 'exchange program', 'study abroad', lived abroad, studied abroad, worked abroad, or similar. "
                    "Student job refers to the most recent relevant student job experience mentioned in the text, and may be written as 'student job', 'student assistant', 'internship', or similar. "
                    "For student_job, include the role and company if both are shown. "
                    "Extract extracurricular activities, including student organizations, voluntary work, elite athletics, board roles, clubs, or similar. "
                    "Extract every separate international experience entry, including exchange semester, study abroad, living abroad, studying abroad, working abroad, international delegation work, and international market work. "
                    "Extract whether the candidate speaks Danish, Swedish, or Norwegian only if explicitly mentioned. Do not infer it from nationality, location, education, name, or country. Also return nordic_language_snippet as the exact supporting text. "
                    "Extract gender only if it is explicitly stated in the text. Do not infer gender from pronouns or from the candidate's name. Also return gender_snippet as the exact supporting text. Otherwise return null. "
                    "GPA may be written as 'grade average', 'average grade', or similar. "
                    "The GPA may appear on the same line or the next line after the label. "
                    "Return the GPA's exactly as written. "
                    "From the grade transcript, extract every ECTS-bearing course row into transcript_grades in the same chronological order as the transcript. "
                    "For transcript_grades, include every row with ECTS points, including rows where the result is Passed, Failed, pass/fail, or any non-numeric result. "
                    "For transcript_grades, include course name, grade if numeric, result if non-numeric, ECTS points, and assessment date if present. Use null grade for pass/fail rows, but never omit pass/fail rows with ECTS. "
                    "Extract transcript grades exactly as written in the transcript's own grading system; do not convert grades yourself. "
                    "For non-Aarhus University transcripts, also extract individual transcript grades into bachelor_grades and master_grades when the level is clear. "
                    "If the transcript uses Danish grades, valid Danish grades are 12, 10, 7, 4, 02, 00, and -3. "
                    "For each transcript grade, include course name and ECTS points if present. "
                    "Classify transcript grades as bachelor or master only when the transcript clearly indicates the level. "
                    "Do not invent missing transcript grades, ECTS points, or education levels. "
                    "For each extracted value, also return the exact snippet from the text that supports it. "
                    "Return all entries you find, not just the most recent one. "
                    "education_snippets should be short exact substrings from the text."
                    "If a value is not explicitly present, return null. Do not guess. "
                    + system_instructions
                )},
                {"role": "user", "content": t},
            ],
            response_format=CandidateInfo,
            temperature=0,
        )


        for c in analysis_clases:
            print(f"Starting persona analysis for {c.__name__}...")

        persona_tasks = [
            client.beta.chat.completions.parse(
                model="gpt-4.1",
                messages=[
                    {"role": "system", "content": (
                        system_instructions
                        + f"Classify the candidate according to the {c.__name__} role, and extract evidence for each classification. "
                        "Use a strict evidence standard. Only mark a boolean true when the CV or cover letter contains explicit evidence for that exact field. "
                        "Do not mark broad or generic words like strategy, digital, data, analysis, IT, project, process, stakeholder, business, transformation, or consulting as true unless they directly prove the specific field. "
                        "For every true boolean, provide a short exact quote from the text. If unsure, mark false."
                    )},
                    {"role": "user", "content": t},
                ],
                response_format=c,
                temperature=0,
            )
            for c in analysis_clases
        ]

        completion, *persona_analyses = await asyncio.gather(
            candidate_task,
            *persona_tasks,
        )

        for analysis in persona_analyses:
            print(analysis)

        persona_results = []
        persona_score_lines = []
        npa = completion.choices[0].message.parsed

        if is_au_transcript(transcript_raw_text) and npa.transcript_grades:
            npa.bachelor_grades, npa.master_grades = split_au_transcript_grades(
                npa.transcript_grades,
                transcript_raw_text,
            )

        grading_system = transcript_grading_system.value
        print("Transcript grading system:", grading_system)
        calculated_bachelor_gpa = calculate_gpa(npa.bachelor_grades, grading_system)
        calculated_master_gpa = calculate_gpa(npa.master_grades, grading_system)
        print_grade_debug("Bachelor", npa.bachelor_grades, calculated_bachelor_gpa)
        print_grade_debug("Master", npa.master_grades, calculated_master_gpa)
        conversion_area.content = format_conversion_report(
            grading_system,
            npa.bachelor_grades,
            npa.master_grades,
            calculated_bachelor_gpa,
            calculated_master_gpa,
        )
        conversion_area.update()

        if calculated_bachelor_gpa is not None:
            npa.bachelor_gpa = [str(calculated_bachelor_gpa)]
            npa.bachelor_gpa_snippets = ["Calculated from uploaded grade transcript"]

        if calculated_master_gpa is not None:
            npa.master_gpa = [str(calculated_master_gpa)]
            npa.master_gpa_snippets = ["Calculated from uploaded grade transcript"]

        for persona_class, analysis in zip(analysis_clases, persona_analyses):
            parsed_persona = analysis.choices[0].message.parsed
            true_count = true_boolean_count(parsed_persona)
            master_university_weight = score_master_university(npa.master_university, persona_class.__name__)
            program_weight = await score_master_program(npa.master_program, persona_class.__name__)
            persona_score = true_count + master_university_weight + program_weight
            persona_results.append((persona_class, parsed_persona, persona_score, program_weight, master_university_weight, true_count))
            persona_score_lines.append(
                format_persona_fit_report(
                    persona_class.__name__,
                    parsed_persona,
                    true_count,
                    master_university_weight,
                    program_weight,
                    persona_score,
                )
            )

        # Rank all personas with the same key the original max() used.
        # Index [0] is the winner; index [1] is the runner-up that we save
        # to columns W and X of the sheet for the human reviewer.
        ranked_personas = sorted(
            persona_results,
            key=lambda result: (result[2], result[3], result[4], result[5]),
            reverse=True,
        )
        winning_persona_class, winning_persona, highest_score, _, _, winning_true_count = ranked_personas[0]
        selected_persona_name = "Unknown" if winning_true_count < 4 else winning_persona_class.__name__
        winning_persona_evidence = collect_persona_evidence(winning_persona, npa)
        winning_persona_evidence.extend(npa.master_university or [])
        winning_persona_evidence.extend(npa.master_program or [])

        # Runner-up persona (same Unknown-threshold as the winner so an
        # under-evidenced fallback doesn't pretend to be a real fit).
        if len(ranked_personas) > 1:
            runner_up_class, _, runner_up_score, _, _, runner_up_true_count = ranked_personas[1]
            runner_up_name = (
                "Unknown" if runner_up_true_count < 4 else runner_up_class.__name__
            )
        else:
            runner_up_name = ""
            runner_up_score = ""

        global all_candidates
        scorecard = await build_scorecard(npa, highest_score, selected_persona_name)
        persona_reason = "" if winning_true_count < 4 else ", ".join(true_boolean_labels(winning_persona))
        candidate_row = build_candidate_row(
            npa, selected_persona_name, highest_score, scorecard, persona_reason,
            second_selected_persona=runner_up_name,
            second_persona_score=runner_up_score,
        )
        all_candidates.append(candidate_row)

        persona_score_lines.append(format_scorecard_report(candidate_row))
        score_area.content = "### Persona fit\n" + "\n".join(persona_score_lines)
        score_area.update()

        print(npa)

        summary_area.content = f"""
        <h2 style="font-size: 32px;">Candidate summary</h2>
        <p><b>Name:</b> {npa.name}</p>
        <p><b>Master:</b> {join_values(npa.master_education)}</p>
        <p><b>Bachelor:</b> {join_values(npa.bachelor_education)}</p>
        <p><b>Student job/internship:</b> {join_values(npa.student_job)}</p>
        <p><b>International experience:</b> {join_values(npa.international_experience)}</p>
        <p><b>Selected persona:</b> {candidate_row['selected_persona']}</p>
        <p><b>Final score:</b> {candidate_row['final_score']}</p>
        """

        cv_matches = collect_evidence_matches(npa, cv_tokens) if cv_tokens else []
        cover_matches = collect_evidence_matches(npa, cover_tokens) if cover_tokens else []
        cv_persona_matches = collect_text_matches(winning_persona_evidence, cv_tokens, clean_targets=False) if cv_tokens else []
        cover_persona_matches = collect_text_matches(winning_persona_evidence, cover_tokens, clean_targets=False) if cover_tokens else []

        print("CV MATCHES FOUND:", len(cv_matches))
        print("COVER MATCHES FOUND:", len(cover_matches))
        print("CV PERSONA MATCHES FOUND:", len(cv_persona_matches))
        print("COVER PERSONA MATCHES FOUND:", len(cover_persona_matches))

        words = []
        for value in [
            *(npa.bachelor_gpa or []),
            *(npa.master_gpa or []),
            *(npa.student_job or []),
            *(npa.extracurricular or []),
            *(npa.exchange_education_snippets or []),
            *(npa.exchange_education or []),
            *(npa.international_experience or []),
            npa.nordic_language_snippet,

        ]:
            if value:
                words += [w.strip(" ,.;:()[]{}") for w in str(value).split() if len(w.strip(" ,.;:()[]{}")) > 2]

        words = list(set(words))

        if cv_pdf_path and (cv_matches or cv_persona_matches):
            highlighted_cv_url = make_highlighted_pdf(cv_pdf_path, cv_matches, cv_persona_matches)
            cv_viewer.props(f'src={highlighted_cv_url}')

        elif cv_pdf_path and words:
            highlighted_cv_url = make_highlighted_pdf_ocr(cv_pdf_path, words)
            cv_viewer.props(f'src={highlighted_cv_url}')

        if cover_pdf_path and (cover_matches or cover_persona_matches):
            highlighted_cover_url = make_highlighted_pdf(cover_pdf_path, cover_matches, cover_persona_matches)
            cover_viewer.props(f'src={highlighted_cover_url}')
        elif cover_pdf_path and words:
            highlighted_cover_url = make_highlighted_pdf_ocr(cover_pdf_path, words)
            cover_viewer.props(f'src={highlighted_cover_url}')

    except Exception as e:
        print("ERROR:", e)
        ui.notify(f"Error: {e}")

    finally:
        progress.visible = False
        label.visible = False

# progress + label widgets are created together with the rest of the UI
# at the bottom of the file, so they appear inside the Step 2 card.

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

# `markdown` element is created together with the rest of the UI at the
# bottom of the file. Keeping the variable so any external reference still works.

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

async def upload_transcript(e: events.UploadEventArguments):
    global transcript_pdf_path, transcript_tokens, transcript_raw_text

    pdf_bytes = await e.file.read()
    safe_name = (e.file.name or "transcript.pdf").replace(" ", "_")
    filename = f'{uuid.uuid4().hex}_{safe_name}'
    transcript_pdf_path = os.path.join(PDF_DIR, filename)

    with open(transcript_pdf_path, 'wb') as f:
        f.write(pdf_bytes)

    transcript_raw_text = extract_text_from_pdf_bytes(pdf_bytes)
    transcript_tokens = extract_word_tokens(pdf_bytes)

# The full UI layout is constructed at the bottom of this file, after every
# handler function (upload_cv, send_input, save_to_google_sheets, ...) is
# defined. See the "Page layout" section below.

def make_highlighted_pdf(src_path, candidate_matches, persona_matches=None):
    doc = fitz.open(src_path)
    persona_matches = persona_matches or []

    for matches, color in [
        (candidate_matches, (1, 1, 0)),
        (persona_matches, (0.53, 0.81, 1)),
    ]:
        for m in matches:
            for token in m["tokens"]:
                page = doc[token["page"]]
                rect = fitz.Rect(token["x0"], token["top"], token["x1"], token["bottom"])
                annot = page.add_highlight_annot(rect)
                annot.set_colors(stroke=color)
                annot.update(opacity=0.35)

    out = f'/tmp/nicegui_pdfs/h_{uuid.uuid4().hex}.pdf'
    doc.save(out)
    doc.close()
    return out.replace('/tmp/nicegui_pdfs', '/pdfs')

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

#Gem til Excel
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GOOGLE_SHEETS_CREDENTIALS_FILE = os.path.join(BASE_DIR, "master-thesis-493710-9125f7300e40.json")

def get_google_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_file(
        GOOGLE_SHEETS_CREDENTIALS_FILE,
        scopes=scopes,
    )
    return gspread.authorize(credentials)

def normalize_persona_tab(persona: str | None) -> str: #denne matcher persona med fane i google sheets
    if not persona:
        return "Unknown"

    p = re.sub(r"[^a-z0-9]", "", persona.lower())

    if "persona1" in p or "techstrategist" in p:
        return "Persona 1 - tech strategist"
    if "persona2" in p or "transformationorchestrator" in p or "pmo" in p:
        return "Persona 2 - transformation orchestrator (PMO)"
    if "persona3" in p or "techtranslator" in p:
        return "Persona 3 - tech translator"
    if "persona4" in p or "appliedtechnicalspecialist" in p:
        return "Persona 4 - applied technical specialist"
    if "persona5" in p or "industryfunctionalspecialist" in p or "industryspecialist" in p:
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

        headers = SHEET_HEADERS

        for c in all_candidates:
            sheet_name = normalize_persona_tab(c.get("selected_persona"))
            print("Working on sheet:", sheet_name, "| original persona:", c.get("selected_persona"))

            try:
                ws = spreadsheet.worksheet(sheet_name)
                print("Using existing worksheet:", sheet_name)
            except gspread.WorksheetNotFound:
                ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(headers))
                print("Created new worksheet:", sheet_name)

            first_row = ws.row_values(1)
            if not first_row:
                ws.append_row(headers)
                print("Headers added to:", sheet_name)
            elif first_row != headers:
                ws.update([headers], range_name="A1")
                print("Headers updated on:", sheet_name)

            ws.append_row([c.get(header, "") for header in headers])
            print("Row appended to:", sheet_name)

        print("Existing tabs after save:", [ws.title for ws in spreadsheet.worksheets()])
        ui.notify(f"Saved {len(all_candidates)} candidates to Google Sheets")
        all_candidates.clear()

    except Exception as e:
        print("SAVE ERROR:", repr(e))
        ui.notify(f"Error: {e}")

def save_to_google_sheets_then_remind():
    save_to_google_sheets()
    save_confirm_dialog.open()

# =============================================================================
# Page layout
# -----------------------------------------------------------------------------
# All ui.* widgets live inside this block so the page reads in the same order
# the user moves through the tool: upload -> analyse -> review & save.
# Every variable that handlers reference (summary_area, score_area,
# conversion_area, cv_viewer, cover_viewer, progress, label, markdown,
# transcript_grading_system, highlighted_output, input_area) is declared here
# at module scope, so the existing logic is unchanged.
# =============================================================================

# --- Header bar (logo + app title) -------------------------------------------
ui.html('''
<div class="app-header">
    <span class="app-title">Candidate Screener</span>
    <span class="app-subtitle">First-pass review of CV, cover letter and transcript</span>
</div>
''')

# --- Page shell --------------------------------------------------------------
with ui.element('div').classes('page-shell'):

    # Hero / intro ------------------------------------------------------------
    ui.html('''
    <div class="hero">
        <h1>Review a candidate</h1>
        <p>Upload the candidate's documents, run the analysis, and save the
        result to the team spreadsheet. The whole flow takes about a minute.</p>
    </div>
    ''')

    # =========================================================================
    # STEP 1 — Upload candidate documents
    # =========================================================================
    with ui.element('div').classes('step-card'):
        ui.html('''
            <div class="step-header">
                <span class="step-number">1</span>
                <span class="step-title">Upload candidate documents</span>
            </div>
            <div class="step-help">Drop in the candidate's PDFs. The CV is required; cover letter and transcript are optional but improve scoring.</div>
        ''')

        with ui.row().classes('w-full').style('gap: 18px; flex-wrap: wrap;'):

            with ui.column().classes('grow').style('min-width: 280px; flex: 1;'):
                ui.html('<div class="upload-tile-label">CV (required)</div>'
                        '<div class="upload-tile-help">PDF of the candidate\'s CV / resume.</div>')
                ui.upload(on_upload=upload_cv, auto_upload=True,
                          label='Drop CV here or click to browse').props('accept=.pdf')

            with ui.column().classes('grow').style('min-width: 280px; flex: 1;'):
                ui.html('<div class="upload-tile-label">Cover letter</div>'
                        '<div class="upload-tile-help">Optional — sharpens persona scoring.</div>')
                ui.upload(on_upload=upload_cover, auto_upload=True,
                          label='Drop cover letter here or click to browse').props('accept=.pdf')

            with ui.column().classes('grow').style('min-width: 280px; flex: 1;'):
                ui.html('<div class="upload-tile-label">Grade transcript</div>'
                        '<div class="upload-tile-help">Optional — used to compute GPA.</div>')
                ui.upload(on_upload=upload_transcript, auto_upload=True,
                          label='Drop transcript here or click to browse').props('accept=.pdf')

        ui.html('<div style="height: 18px;"></div>')

        with ui.column().style('max-width: 360px;'):
            ui.html('<div class="upload-tile-label">Grading system on the transcript</div>'
                    '<div class="upload-tile-help">Which grading system did the candidate study under? (danish is selected as default)</div>')
            transcript_grading_system = ui.select(
                GRADING_SYSTEMS,
                value="Danish 7-point",
            ).props('outlined dense').classes('w-full')

        # Hidden helper widget kept so any reference elsewhere still resolves.
        input_area = ui.textarea(label='PDF text').classes('w-full')
        input_area.visible = False
        # Friendly hint widget (replaces the old floating "Choose a PDF file!" line).
        markdown = ui.markdown('').classes('result-area')
        markdown.visible = False

    # =========================================================================
    # STEP 2 — Run analysis
    # =========================================================================
    with ui.element('div').classes('step-card'):
        ui.html('''
            <div class="step-header">
                <span class="step-number">2</span>
                <span class="step-title">Run the analysis</span>
            </div>
            <div class="step-help">Reads the documents, calculates the GPA, and scores the candidate against five Accenture personas. This may take a minute.</div>
        ''')

        with ui.row().classes('items-center').style('gap: 16px;'):
            ui.button("Analyse candidate",
                      on_click=send_input,
                      color='#a100ff',
                      icon='auto_awesome').classes('text-white').props('unelevated')

            # Progress card — the wrapper's visibility is bound to the
            # spinner's `.visible`, so the existing handler's
            # `progress.visible = True/False` toggles the whole card too,
            # without touching the analysis logic.
            with ui.row().classes('items-center progress-card') as _progress_box:
                progress = ui.spinner(size='md', color='#c266ff')
                label = ui.label('0%')
            progress.visible = False
            label.visible = False
            _progress_box.bind_visibility_from(progress, 'visible')

        ui.html('<div style="height: 14px;"></div>')

        # GPA conversion report — shown after analysis runs.
        conversion_area = ui.markdown('').classes('w-full result-area')

    # =========================================================================
    # STEP 3 — Review & save
    # =========================================================================
    with ui.element('div').classes('step-card'):
        ui.html('''
            <div class="step-header">
                <span class="step-number">3</span>
                <span class="step-title">Review &amp; save</span>
            </div>
            <div class="step-help">Yellow highlights show candidate facts; blue highlights show evidence for the suggested persona. Both the best fitted persona and the second best will be saved to the spreadsheet. Save to the spreadsheet when you're happy.</div>
        ''')

        # Candidate summary + persona scores -----------------------------------
        with ui.row().classes('w-full').style('gap: 24px; flex-wrap: wrap;'):
            with ui.column().style('flex: 1; min-width: 320px;'):
                ui.html('<div class="upload-tile-label" style="margin-bottom: 10px;">Candidate summary</div>')
                summary_area = ui.html('<div class="empty-state">Run the analysis to see the candidate summary here.</div>') \
                    .classes('w-full result-area')

            with ui.column().style('flex: 1; min-width: 320px;'):
                ui.html('<div class="upload-tile-label" style="margin-bottom: 10px;">Persona fit</div>')
                score_area = ui.markdown('No persona scores yet — upload documents and click Analyse candidate.') \
                    .classes('w-full result-area empty-state')

        ui.html('<div style="height: 18px;"></div>')

        # Inline highlight legend (replaces the old floating box) -------------
        ui.html('''
            <div class="legend">
                <span><span class="swatch yellow"></span>Candidate facts</span>
                <span><span class="swatch blue"></span>Persona evidence</span>
            </div>
        ''')

        ui.html('<div style="height: 14px;"></div>')

        # PDF viewers side-by-side --------------------------------------------
        with ui.row().classes('w-full').style('gap: 16px; flex-wrap: wrap;'):
            with ui.element('div').classes('pdf-viewer').style('flex: 1; min-width: 360px;'):
                cv_viewer = ui.element('iframe')
            with ui.element('div').classes('pdf-viewer').style('flex: 1; min-width: 360px;'):
                cover_viewer = ui.element('iframe')

        # Hidden helper kept so any reference elsewhere still resolves.
        highlighted_output = ui.html().classes('w-full')
        highlighted_output.visible = False

        ui.html('<div style="height: 22px;"></div>')

        # Pre-save reminder dialog -------------------------------------------
        # Augmented review: nudges the human to add initials + comment in the
        # spreadsheet. Clicking the Save button saves first, then opens this
        # dialog; "Okay" closes it.
        with ui.dialog() as save_confirm_dialog, ui.card().classes('save-confirm-card'):
            ui.html('<div class="save-confirm-title">Don\'t forget</div>')
            ui.html('<div class="save-confirm-body">Write your initials in the sheet for verification and add a comment if anything stood out to you.</div>')
            with ui.row().classes('w-full').style('justify-content: flex-end; margin-top: 18px;'):
                ui.button('Okay',
                          on_click=save_confirm_dialog.close,
                          color='#a100ff').classes('text-white').props('unelevated')

        # Save action ---------------------------------------------------------
        with ui.row().classes('w-full items-center').style('gap: 12px; justify-content: flex-end;'):
            ui.html('<span style="color: var(--text-muted); font-size: 13px;">Saves to the Accenture prototype spreadsheet.</span>')
            ui.button('Save candidate to spreadsheet',
                      on_click=save_to_google_sheets_then_remind,
                      color='#a100ff',
                      icon='save').classes('text-white').props('unelevated')


def set_background(color: str = '#0d0d10'):
    ui.query('body').style(f'background-color: {color}')

set_background('#0d0d10')

ui.run()
