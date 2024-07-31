**Input: Excel file containing tasks**

- rows: represent tasks
- required columns:
- Status-- Blocked, Working, Committed, Soft Commit, Completed - Partial, Completed, Deprioritzed, Canceled, or Review
- Timeline-- target dates for subitems of the project
- Name-- task/ project name

**Output: Creates an executive PPP**

- Progress: tasks completed with a target date within the last week - next two months
- Plan: tasks planned for the next two months
- Problems: tasks that are blocked or are overdue

**Additional Instructions**

- to run locally:
     - windows: python app.py
     - macOS: python3 app.py
- to run with openai (ISSUE: 429 status code-- exeeding current quota):
     1. must activate virtual environment:
        - windows: venv\Scripts\activate
        - macOS: source venv/bin/activate
     2. set environment variable:
        - windows: set OPENAI_API_KEY=<your api key (does not need to be enclosed in quotations)> and run as you would locally
        - macOS: export OPENAI_API_KEY='your_api_key_here' and run as you would locally


           
