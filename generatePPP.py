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
        target_date = row['Target Date'].strftime('%-m/%-d')
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
            "NOTE: keep GOAL and NAME simple, concise, and easy to understand."
            "1. Identify the executive business goal/objective/ corporate initiative this task addresses. "
            "This field is called GOAL.\n"
            "2. Name this task. This should briefly describe in summary what the task is specifically. "
            "This field is called NAME.\n"
            "3. Identify which individual(s) is in charge of driving/ completing this task. "
            "This field is called ASSIGNEE.\n"
            "4. FORMAT YOUR RESPONSE: format the fields identified in steps 1-3 in your response. Use this syntax:\n"
            "<b>{GOAL}</b>: {NAME} [{ASSIGNEE}]\n"
            "NOTE: fields are enclosed in curly brackets {}. "
            "Replace the fields enclosed in curly brackets with the information you've identified.\n"
            "\n"
            "Here is an example of how a task should be formatted in your response: \n"
            "<b>Internship Training Program: Give all interns access to Mission Control trainings [Program Lead]\n"
        )

        # asking AI
        task = ask_openai(client, system_prompt, user_prompt)
        print(task)

        return f"<span class='date'>{target_date}</span> {og_target_date}{task} {blocked}{overdue}", row['Target Date']

    except Exception as formatting_error:
        return f"Error formatting task: {formatting_error}"


def create_ppp(file_path, pg=None):
    """takes a file path and a page to an Excel sheet (provided optionally)
    and generates a PPP for it"""
    try:
        if pg:  # page to Excel sheet provided
            df = pd.read_excel(file_path, sheet_name=pg)
        else:  # no page to Excel sheet provided
            df = pd.read_excel(file_path)

        # converting dates to datetime
        df['Target Date'] = pd.to_datetime(df['Target Date'], format='%m/%d/%Y', errors='coerce')
        df['Complete Date'] = pd.to_datetime(df['Complete Date'], format='%m/%d/%Y', errors='coerce')

        # check for optionally existing columns: comments or original target date
        has_og_target_date = 'Original Target Date' in df.columns
        has_comments = 'Comments' in df.columns

        # get range of dates for progress and plan sections
        today = datetime.now()  # today's date
        last_week = today - timedelta(days=7)  # 7 days ago
        two_months = today + timedelta(days=60)  # 60 days from now

        # getting tasks for each section
        # progress -- completed within last week
        # plan -- new or being worked on for next two months
        # problems -- blocked or overdue

        progress_section = df[(df['Complete Date'] >= last_week) & (df['Complete Date'] <= today)]
        plan_section = df[((df['Status'] != 'Completed') & (df['Status'] != 'Canceled') & (df['Status'] != 'Blocked')) &
                          ((df['Target Date'] > today) & (df['Target Date'] <= two_months))]
        blocked_section = df[df['Status'] == 'Blocked']
        overdue_section = df[(df['Target Date'] <= today) * (df['Status'] != 'Completed')]

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
        progress_output = "<br>".join(f"  •  {task[0]}" for task in progress) + "<br><br>"
        plans_output = "<br>".join(f"  • {task[0]}" for task in plan) + "<br><br>"
        problems_output = "<br>".join(f"  • {task[0]}" for task in problems) + "<br><br>"

        print(problems_output)

        # calling openai
        return progress_output, plans_output, problems_output

    except Exception as ppp_error:
        return str(ppp_error)
