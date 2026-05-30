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
import random
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from openai import AsyncClient

PDF_DIR = '/tmp/nicegui_pdfs'
os.makedirs(PDF_DIR, exist_ok=True)
app.add_static_files('/pdfs', PDF_DIR)
app.add_static_files('/assets', '/Users/danicahennelly/Speciale UDENFOR iCloud')

all_candidates = []  # Global list to store candidate Excel row dictionaries

cv_pdf_path = None
cover_pdf_path = None

cv_tokens = []
cover_tokens = []

cv_raw_text = ''
cover_raw_text = ''

load_dotenv()
client = AsyncClient(api_key=os.getenv("API_KEY"))
input_area = ui.textarea(label='PDF text').classes('w-full') 
input_area.visible = False
summary_area = ui.html().classes('w-full text-white').style('margin-top: 90px;')
output_area = ui.textarea(label='Extracted text').classes('w-full').style('height: 180px;') 
score_area = ui.markdown('No persona scores yet.').classes('w-full text-white')
ui.image('/assets/logo.png').style('position: fixed; top: 24px; right: 24px; width: 180px; z-index: 1000;')
ui.add_head_html('''
	                 <style>
		                 @font-face {
		                 font-family: 'Inter';
		                 src: url('/assets/Inter.woff2') format('woff2');
		                 }
		                 body {
		                 font-family: 'Inter', Arial, sans-serif;
		                 }
		                 textarea {
		                 color: white !important;
		                 }
	                 .q-uploader__header {
	                 background: #a301ff !important;
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

def build_candidate_row(npa, selected_persona, persona_score, scorecard, persona_reason=""):
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

        t = f"CV:\n{cv_raw_text}\n\nCOVER LETTER:\n{cover_raw_text}"

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

        winning_persona_class, winning_persona, highest_score, _, _, winning_true_count = max(
            persona_results,
            key=lambda result: (result[2], result[3], result[4], result[5])
        )
        selected_persona_name = "Unknown" if winning_true_count < 4 else winning_persona_class.__name__
        winning_persona_evidence = collect_persona_evidence(winning_persona, npa)
        winning_persona_evidence.extend(npa.master_university or [])
        winning_persona_evidence.extend(npa.master_program or [])

        global all_candidates
        scorecard = await build_scorecard(npa, highest_score, selected_persona_name)
        persona_reason = "" if winning_true_count < 4 else ", ".join(true_boolean_labels(winning_persona))
        candidate_row = build_candidate_row(npa, selected_persona_name, highest_score, scorecard, persona_reason)
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

        output_area.value = str(persona_analyses[0])
        
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
        output_area.value = str(e)

    finally:
        progress.visible = False
        label.visible = False

progress = ui.spinner(size='lg', color='#a301ff')
label = ui.label('0%')

progress.visible = False
label.visible = False

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

ui.label('Upload CV').classes('text-white text-bold').style('font-size: 18px;')
ui.upload(on_upload=upload_cv, auto_upload=True).props('accept=.pdf color=#a301ff')

ui.label('Upload Cover Letter').classes('text-white text-bold').style('font-size: 18px;')
ui.upload(on_upload=upload_cover, auto_upload=True).props('accept=.pdf color=#a301ff')

ui.button("Extract", on_click=send_input, color='#a301ff').classes('text-white')

with ui.row().classes('w-full'):
    cv_viewer = ui.element('iframe').classes('w-1/2').style('height:900px; border:none;')
    cover_viewer = ui.element('iframe').classes('w-1/2').style('height:900px; border:none;')

#Add second area for highlighting nice gui
# create UI element
highlighted_output = ui.html().classes('w-full')

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
GOOGLE_SHEETS_CREDENTIALS_FILE = os.path.join(BASE_DIR, "TOP SECRET GOOGLE SHEET.json")

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
            if first_row != headers:
                ws.insert_row(headers, 1)
                print("Headers added to:", sheet_name)

            ws.append_row([c.get(header, "") for header in headers])
            print("Row appended to:", sheet_name)

        print("Existing tabs after save:", [ws.title for ws in spreadsheet.worksheets()])
        ui.notify(f"Saved {len(all_candidates)} candidates to Google Sheets")
        all_candidates.clear()

    except Exception as e:
        print("SAVE ERROR:", repr(e))
        ui.notify(f"Error: {e}")

ui.button('Save to Sheets', on_click=save_to_google_sheets, color='#a301ff').classes('text-white')

def set_background(color: str = '#000000'):
    ui.query('body').style(f'background-color: {color}') 

set_background('#000000')

ui.run()
