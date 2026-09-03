# Assessment Import Script - Run Guide (`new_convert_assessment_update_07.py`)

## Purpose

`new_convert_assessment_update_07.py` converts student Google Forms assessment answers into PostgreSQL/Supabase seed SQL files.

### Key Rules
- **Learners are never created:** The SQL resolves existing learner records by email. The learner must already exist in `public.learners`.
- **Question UUIDs are never generated:** Existing question UUIDs from the question bank are strictly reused.
- **Target tables only:**
  1. `public.adaptive_aptitude_sessions`
  2. `public.personal_assessment_attempts`
  3. `public.adaptive_aptitude_responses`
  4. `public.adaptive_aptitude_results`
- **Does not insert into `personal_assessment_results`** (those are handled by separate result seeds).

---

## Command Syntax

```powershell
python new_convert_assessment_update_07.py "Assessment profiling answers.xlsx" --question-bank "question_bank.xlsx" --production-data "PRODUCTION_ASSESSMENT.xlsx" --allow-partial
```

---

## Required Input Files & Where They Are Used

### 1. File 1: Student Answers Workbook (Positional Argument: `answers`)
* **Default:** `"Assessment Answers.xlsx"` (or specified filename e.g. `"Assessment profiling answers.xlsx"`)
* **Loaded in:** `load_form_students(path)` (lines 593–618)
* **What it is:** Raw student submissions exported directly from Google Forms / Google Sheets.
* **Expected Sheets (`FORM_SHEETS`):**
  * `Big5` $\to$ mapped to `bigfive`
  * `Riasec` $\to$ mapped to `riasec`
  * `Employability Assessment ` $\to$ mapped to `employability`
  * `Work Values` $\to$ mapped to `values`
  * `Genearl Apt` $\to$ mapped to `adaptive_aptitude`
  * `MBA Domain knowledge` $\to$ mapped to `mba_knowledge` (MBA only)
  * `mba apt` $\to$ mapped to `mba_aptitude` (MBA only)
  * `mca domain ` $\to$ mapped to `mca_knowledge` (MCA only)
  * `mca apt` $\to$ mapped to `mca_aptitude` (MCA only)
* **Where it is used in the code:**
  * Extracts each student's **Email**, **Full Name**, **Program/Stream** (`MBA` vs `MCA`), and submission timestamps.
  * Used in `convert_student(...)` (lines 773–995) to parse raw responses:
    * Likert text (e.g., *"Very Accurate"*, *"Strongly Like"*) $\to$ scores $1\text{--}5$.
    * Situational Judgment Tests (SJT Best/Worst choices) $\to$ paired JSON objects.
    * Multiple-choice options $\to$ option keys (`A`, `B`, `C`, `D`).

---

### 2. File 2: Master Question Bank (`--question-bank "question_bank.xlsx"`)
* **Default:** `"question_bank.xlsx"`
* **Loaded in:** `load_reference_bank(path)` (lines 320–367)
* **What it is:** The master database export containing verified question UUIDs, question texts, options, and correct answers.
* **Expected Sheets:**
  1. **`personal_assessment_questions`**:
     * **Columns:** `id` (UUID), `question_text`, `question_type`, `options`, `correct_answer`, `metadata`, `description`.
     * **Where used:** In `resolve_question(...)` (line 412) to match headers in `Big5`, `Riasec`, `Work Values`, `Employability`, and `Genearl Apt` to static question database UUIDs.
  2. **`career_assessment_ai_questions`**:
     * **Columns:** `question_type`, `stream_id`, and a JSON column `questions` containing `uuid`/`id`, `question`, `options`, `correct_answer`, `category`, `skill_tag`, `difficulty`.
     * **Where used:** In `resolve_ai_question(...)` (line 631) to map MBA and MCA domain knowledge and domain aptitude questions to their existing UUIDs and correct answer keys.
* **Why it is strictly required:** The script never creates new question UUIDs. It strictly matches questions against this bank to ensure all generated seed SQL references valid database question records.

---

### 3. File 3: Curated Production Data (`--production-data "PRODUCTION_ASSESSMENT.xlsx"`)
* **Default:** `None` (typically `FINAL_PRODUCTION_ASSESSMENT_DATA_v3.xlsx` or `PRODUCTION_ASSESSMENT.xlsx`)
* **Loaded in:** `load_production_data(path)` (lines 705–750) and `apply_cleaned_rows(...)` (lines 751–764)
* **What it is:** The curated master file that controls eligibility policy and supplies cleaned/deduplicated student responses.
* **Expected Sheets:**
  1. **`Eligible Students`**:
     * **Columns:** `Email`, `Program`.
     * Marks student status as `ELIGIBLE` and sets their verified program (`MBA` or `MCA`).
  2. **`Not Eligible Students`**:
     * **Columns:** `Email`, `Primary Reason`, `Program`.
     * If reason is `"Conflicting Duplicate Answer"`, the script sets status to `ALLOW_FIRST_DUPLICATE` (retaining the first answer and allowing seed generation).
     * If reason is missing sections or identity issues, generation is blocked (`BLOCK`).
  3. **`Eligible - <Stage>` Sheets** (`Eligible - Big5`, `Eligible - RIASEC`, `Eligible - Employability`, `Eligible - Work Values`, `Eligible - General Aptitude`, `Eligible - MBA Domain`, `Eligible - MBA Aptitude`, `Eligible - MCA Domain`, `Eligible - MCA Aptitude`):
     * **Where used:** For students marked `ELIGIBLE`, these cleaned rows override the raw form data from the answers file to resolve conflicts and data errors.

---

### CLI Flags

* `--allow-partial`:
  * By default, if any student has blocking validation errors, the script exits without writing SQL.
  * When `--allow-partial` is passed, the script writes SQL seed files for all valid students while logging issues for invalid ones.
* `--output-dir <path>`:
  * Defaults to `output/`. Specifies where `validation_report.csv` and the `seeds/` folder are saved.

---

## Output Files

```text
output/
|-- validation_report.csv
`-- seeds/
    |-- <email_local_part>_assessment_seed.sql
    `-- ...
```

1. **`output/validation_report.csv`**:
   * CSV report detailing validation `ERROR`s and `WARNING`s for every student.
2. **`output/seeds/<email_local_part>_assessment_seed.sql`**:
   * PostgreSQL seed transaction for each valid student inserting into:
     * `public.adaptive_aptitude_sessions`
     * `public.personal_assessment_attempts`
     * `public.adaptive_aptitude_responses`
     * `public.adaptive_aptitude_results`

---

## Step-by-Step Instructions

### 1. Environment Setup

```powershell
python --version
pip install openpyxl
```

### 2. Run Migration Script

```powershell
python new_convert_assessment_update_07.py "Assessment profiling answers.xlsx" --question-bank "question_bank.xlsx" --production-data "PRODUCTION_ASSESSMENT.xlsx" --allow-partial
```

### 3. Select Student Count

When prompted:
```text
How many valid students do you want to generate seed files for? Enter 0 for all <N>:
```
* Enter `0` to generate seed files for all valid students.
* Or enter a specific number (e.g. `1`, `5`) to generate a test batch.

### 4. Review Reports & Run Seeds in Supabase

1. Inspect `output/validation_report.csv` for any warnings or rejected students.
2. Copy the generated `.sql` file(s) from `output/seeds/` to `skillpassport/supabase/seed/college/` or execute directly in the Supabase SQL editor.
