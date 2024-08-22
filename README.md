**Input: Monday.com Excel export (.xls, .xlsx)**

- rows: represent tasks (separated by Status)
- required columns:
   - Status-- Blocked, Working, Committed, Soft Commit, Completed - Partial, Completed, Deprioritzed, Canceled, or Review
   - Timeline-- target dates for subitems of the project
   - Completed Date-- if task is complete, time stamp of completion
   - Audience-- flags tasks should be prioritized for a certain audience
     - current options-- Natalie or blank
  
**Output: Creates an executive PPP**

- Progress: Status is 'Completed' and target date is within last week
- Plan: Status is ['Working', 'Commited', 'Completed - Partial'] and target date is within next two months 
- Problems: Status is blocked or overdue

**How it works**

- traverses through tasks
- performs necessary filtering to distribute tasks to their respective sections of the PPP
- uses openai to to summarize tasks and their subitems
    - formatting: <target date> <task title>: <task summmary> [<assignee>]

**Additional Instructions**

- to run locally:
     - windows: python app.py
     - macOS: python3 app.py
- Configuration:
     1. must activate virtual environment and install dependencies (requirements.txt):
        - windows: venv\Scripts\activate
        - macOS: source venv/bin/activate
     2. in config.py, write this line:
        - OPENAI_API_KEY=<your api key (does not need to be enclosed in quotations)>
     3. in mondayPPP.py, configure openAI access
        - make sure that this line is defined:
            - openai.api_key = OPENAI_API_KEY
     5. set environment variable in virtual environment:
        - windows: set OPENAI_API_KEY=<your api key (does not need to be enclosed in quotations)> and run as you would locally
        - macOS: export OPENAI_API_KEY='your_api_key_here' and run as you would locally
           

**Tech Stack**

Frontend (Interface): 
- Programming Language: HTML/CSS (for structure) and JS (handling form submission)
  
Backend (API endpoint):
- Programming Language: Python
- Web Framework: Flask (handles HTTP requests, render templates, manages routing)
- Libraries: 
   - Data manipulation/ analysis: Pandas (handles Excel files and processes data)
   - File uploads: Werkzeug (securely handles file uploads)
   - Summary generation: OpenAI model GPT-3.5 (generates/ formats PPP report content)
     
