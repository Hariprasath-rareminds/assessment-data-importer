# Assessment Import Script - Run Guide

## Purpose

`convert_assessment.py` converts Google Forms assessment answers into
PostgreSQL/Supabase seed SQL files.

It: - Reads `Assessment Answers.xlsx`. - Uses `question_bank.xlsx` to
map answers to existing question UUIDs. - Uses
`career_assessment_ai_questions_rows.sql` for AI/MBA question
mappings. - Validates student responses. - Creates one SQL seed file per
valid student. - Does not insert into `personal_assessment_results`.

Generated seeds insert into: 1. `adaptive_aptitude_sessions` 2.
`personal_assessment_attempts` 3. `adaptive_aptitude_responses` 4.
`adaptive_aptitude_results`

## Required Files

``` text
assessment-import/
|-- convert_assessment.py
|-- Assessment Answers.xlsx
|-- question_bank.xlsx
`-- career_assessment_ai_questions_rows.sql
```

## 1. Check Python

``` powershell
python --version
```

If needed:

``` powershell
py --version
```

## 2. Install Dependency

``` powershell
pip install openpyxl
```

## 3. Run the Script

``` powershell
python convert_assessment.py "Assessment Answers.xlsx" --question-bank "question_bank.xlsx" --ai-sql "career_assessment_ai_questions_rows.sql" --allow-partial
```

## 4. Choose Student Count

After validation, the script asks how many valid students to process.

Example:

``` text
Students found: 54
Valid students: 33
Validation issues: 27
Report: output\validation_report.csv

How many valid students do you want to generate seed files for?
Enter 0 for all 33:
```

Enter `0` to process all valid students.

Enter a number such as `5` to process only the first 5 valid students.

## 5. Output

``` text
output/
|-- validation_report.csv
`-- seeds/
    |-- abhikdabhi_assessment_seed.sql
    |-- abhisheknd267_assessment_seed.sql
    `-- ...
```

The filename uses only the part before `@`.

``` text
abhikdabhi@gmail.com
        |
        v
abhikdabhi_assessment_seed.sql
```

## 6. Check Validation

Review:

``` text
output/validation_report.csv
```

With `--allow-partial`, valid students can still be processed even when
other students have validation issues.

## 7. Run Seed in Supabase

Open the required `.sql` file from `output/seeds/`, copy it into the
Supabase SQL Editor, and run it.

The SQL resolves the existing learner by email:

``` sql
SELECT id
INTO v_learner_id
FROM public.learners
WHERE lower(email) = lower('student@gmail.com')
LIMIT 1;
```

The learner must already exist in `public.learners`.

## Complete Process

``` text
Google Forms Answers
        |
        v
Assessment Answers.xlsx
        |
        v
convert_assessment.py
        |
        v
Load Question Bank + AI Questions
        |
        v
Map Answers to Existing Question UUIDs
        |
        v
Validate Students
        |
        v
Choose Number of Students
        |
        v
Generate One Seed File Per Student
        |
        v
Review validation_report.csv
        |
        v
Run Seed in Supabase
```

## Important

-   Do not change existing question UUIDs.
-   The script does not create learners.
-   Student email must already exist in `public.learners`.
-   Review `validation_report.csv` before importing.
-   Test one student seed first before importing all students.
