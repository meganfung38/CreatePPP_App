import pandas as pd
from datetime import datetime, timedelta
import openai
import requests
from config import OPENAI_API_KEY


# debugging connection errors
print(f"OpenAI library version: {openai.__version__}")
print(f"API Key (first 5 chars): {OPENAI_API_KEY[:5]}...")
try:
    response = requests.get("https://api.openai.com/v1/engines",
                            headers={"Authorization": f"Bearer {OPENAI_API_KEY}"})
    print(f"OpenAI API response status: {response.status_code}")
    print(f"OpenAI API response: {response.text[:100]}...")  # Print first 100 chars
except requests.RequestException as request_error:
    print(f"Error reaching OpenAI API: {request_error}")


# configuring openAI access
openai.api_key = OPENAI_API_KEY
client = openai.OpenAI()  # creating an OpenAI client instance


def ask_openai(openai_client, system_prompt, user_prompt):
    """calls openai"""
    try:
        completion = openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            temperature=0,
            messages=[
                {
                    "role": "system",
                    "content": system_prompt
                },
                {
                    "role": "user",
                    "content": f"Here is the task information: {user_prompt}"
                }
            ]
        )
        return completion.choices[0].message.content
    # debugging
    except Exception as openai_error:
        return f"Unexpected error: {openai_error}"


def format_task(row, has_og_target_date, has_comments, is_blocked=False, is_overdue=False):
    """formats the progress, plan, and overdue tasks using AI"""
    try:
        target_date = row['Timeline'].strftime('%-m/%-d')
        if has_og_target_date and pd.notna(row['Original Target Date']):  # task has an original target date to include
            og_target_date = f"<span class='og-date'>({row['Original Target Date'].strftime('%-m/%-d')})</span> "
        else:
            og_target_date = ""
        if is_blocked:  # task is blocked
            if has_comments and pd.notna(row['Comments']):
                blocked = f"<span class='red-text'>({row['Comments']})</span> "
            else:
                blocked = "<span class='red-text'>(blocked)</span> "
        else:
            blocked = ""
        if is_overdue:  # task is overdue
            overdue = "<span class='red-text'>(overdue)</span>"
        else:
            overdue = ""

        # setting AI prompts up
        user_prompt = row.to_dict()
        system_prompt = (
            "You are an expert in summarizing and restructuring tasks.\n"
            "Your purpose is to analyze the provided details for a task, and "
            "extract critical information that is important and relevant for an executive summary.\n"
            "NOTE: keep GOAL and SUMMARY simple, concise, and easy to understand."
            "1. Identify the name of this task and what department it belongs to."
            " Come up with an overall business goal/objective/ corporate initiative for this task"
            " using the name and department of the task. "
            "This field is called GOAL.\n"
            "2. Identify the subitems/ sub-tasks and the name of this task."
            " Summarize in a sentence what the task is using the subitems/ sub-tasks and name of this task."
            " Only cover what would be relevant for an executive. "
            "This field is called SUMMARY.\n"
            "3. Identify which individual(s) is/ are in charge of driving/ completing this task (DRI). "
            "This field is called ASSIGNEE.\n"
            "4. FORMAT YOUR RESPONSE: format the fields identified in steps 1-3 in your response. Use this syntax:\n"
            "<b>{GOAL}</b>: {SUMMARY} <span class='assignee'>[{ASSIGNEE}]</span>\n"
            "NOTE: fields are enclosed in curly brackets {}. "
            "Replace the fields enclosed in curly brackets with the information you've identified.\n"
            "\n"
            "Example of how information for a task may be given: \n"
            "- Name: Internship Training Program Mission Control Access\n"
            "- Department: Internship Program\n"
            "- Subitems: Configure Mission Control access,"
            " give phase one interns access to Mission Control trainings,"
            " give phase two interns access to Mission Control trainings,"
            " give phase three interns access to Mission Control trainings\n"
            "- DRI: Program Lead\n"
            "Here is an example of how a task should be formatted in your response: \n"
            "<b>Internship Training Program</b>:"
            " Give all interns access to Mission Control trainings <span class='assignee'>[Program Lead]</span>\n"
        )

        # asking AI
        task = ask_openai(client, system_prompt, user_prompt)
        print(task)

        return f"<span class='date'>{target_date}</span> {og_target_date}{task} {blocked}{overdue}", row['Timeline']

    except Exception as formatting_error:
        return f"Error formatting task: {formatting_error}"


def to_datetime(timeline):
    """returns the latest date in a cell as a datetime"""
    # check if cell is empty
    if pd.isna(timeline) or timeline.strip() == "":
        return pd.NaT  # skip empty or NaN cells
    dates = [date.strip() for date in timeline.split(',')]  # get list of dates in cell
    converted_dates = pd.to_datetime(dates, errors='coerce')  # convert to datetime
    return converted_dates.max()  # get the latest date


def create_ppp(file_path, pg=None):
    """takes a file path and a page to an Excel sheet (provided optionally)
    and generates a PPP for it"""
    try:
        # open file and skip first four rows:
        # - row 1: '24Q3 Review Portfolio'-- board name
        # - row 2: 'A high level overview of all your upcoming, current and completed projects.'-- board description
        # - row 3: blank spacer
        # - row 4: 'Committed'-- data frame is sorted by column 'Status'
        if pg:  # page to Excel sheet provided
            df = pd.read_excel(file_path, sheet_name=pg, skiprows=4)
        else:  # no page to Excel sheet provided
            df = pd.read_excel(file_path, skiprows=4)

        # checking for required columns for PPP report
        required = ['Status', 'Timeline']
        missing_columns = [col for col in required if col not in df.columns]
        if missing_columns:  # missing columns
            raise Exception(f"Missing required columns: {', '.join(missing_columns)}")

        # Filter rows that aren't project tasks
        df = df[~df['Name'].isna() &
                (df['Name'].str.strip() != '') &
                (df['Name'].str.strip() != 'Subitems') &
                (df['Name'].str.strip() != 'Name') &
                (df['Name'].str.strip() != 'Review') &
                (df['Name'].str.strip() != 'Closed')]

        df['Timeline'] = df['Timeline'].apply(to_datetime)  # find latest date in timeline column

        # get range of dates for progress and plan sections
        today = datetime.now()  # today's date
        last_week = today - timedelta(days=7)  # 7 days ago
        two_months = today + timedelta(days=60)  # 60 days from now

        # getting tasks for each section
        # progress -- ‘Completed’ (target date is within last week - next two months)
        # plan -- ‘Working’, ‘Committed’, ‘Completed - Partial’ (target date is in next two months)
        # problems -- 'Blocked' or overdue

        progress_section = df[(df['Timeline'] >= last_week) &
                              (df['Timeline'] <= two_months) &
                              (df['Status'] == 'Completed')]
        plan_section = df[((df['Status'] == 'Working') |
                          (df['Status'] == 'Committed') |
                          (df['Status'] == 'Completed - Partial')) &
                          ((df['Timeline'] > today) &
                          (df['Timeline'] <= two_months))]
        blocked_section = df[df['Status'] == 'Blocked']
        overdue_section = df[(df['Timeline'] <= today) &
                             (df['Status'] != 'Completed') &
                             (df['Status'] != 'Soft Commit') &
                             (df['Status'] != 'Deprioritized') &
                             (df['Status'] != 'Canceled') &
                             (df['Status'] != 'Review') &
                             (df['Status'] != 'Blocked')]

        # check for optionally existing columns: comments or original target date
        has_og_target_date = 'Original Target Date' in df.columns
        has_comments = 'Comments' in df.columns

        # formatting tasks for PPP
        progress_tasks = [
            format_task(row, has_og_target_date, has_comments)
            for index, row in progress_section.iterrows()
        ]
        plan_tasks = [
            format_task(row, has_og_target_date, has_comments)
            for index, row in plan_section.iterrows()
        ]
        problem_tasks = [
            format_task(row, has_og_target_date, has_comments, is_blocked=True)
            for index, row in blocked_section.iterrows()
        ] + [
            format_task(row, has_og_target_date, has_comments, is_overdue=True)
            for index, row in overdue_section.iterrows()
        ]

        # sorting tasks for each section by target date
        progress = sorted(progress_tasks, key=lambda target_date: target_date[1])
        plan = sorted(plan_tasks, key=lambda target_date: target_date[1])
        problems = sorted(problem_tasks, key=lambda target_date: target_date[1])

        # formatting PPP
        progress_output = "<br>".join(f"  •  {task[0]}" for task in progress)
        plans_output = "<br>".join(f"  • {task[0]}" for task in plan)
        problems_output = "<br>".join(f"  • {task[0]}" for task in problems)

        # Check if sections are empty and set default message if so
        if not progress_output:
            progress_output = "No tasks completed within the last week."
        else:
            progress_output += "<br><br>"

        if not plans_output:
            plans_output = "Nothing planned for the next two months."
        else:
            plans_output += "<br><br>"

        if not problems_output:
            problems_output = "No blocked or overdue projects."
        else:
            problems_output += "<br><br>"

        # calling openai
        return progress_output, plans_output, problems_output

    except Exception as ppp_error:
        print(f"Error: {ppp_error}")
        return str(ppp_error)
