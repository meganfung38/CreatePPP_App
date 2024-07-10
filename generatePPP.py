import pandas as pd
from datetime import datetime, timedelta
# import openai
# import socket
# import requests
# from config import OPENAI_API_KEY
#
#
# # debugging connection errors
# print(f"OpenAI library version: {openai.__version__}")
# print(f"API Key (first 5 chars): {OPENAI_API_KEY[:5]}...")
# try:
#     response = requests.get("https://api.openai.com/v1/engines",
#                             headers={"Authorization": f"Bearer {OPENAI_API_KEY}"})
#     print(f"OpenAI API response status: {response.status_code}")
#     print(f"OpenAI API response: {response.text[:100]}...")  # Print first 100 chars
# except requests.RequestException as e:
#     print(f"Error reaching OpenAI API: {e}")
#
#
# # configuring openAI access
# openai.api_key = OPENAI_API_KEY
# client = openai.OpenAI()  # creating an OpenAI client instance
#
#
# def ask_openai(openai_client, system_prompt, user_prompt):
#     """calls openai"""
#     try:
#         completion = openai_client.chat.completions.create(
#             model="gpt-3.5-turbo",
#             temperature=0,
#             messages=[
#                 {
#                     "role": "system",
#                     "content": system_prompt
#                 },
#                 {
#                     "role": "user",
#                     "content": user_prompt
#                 }
#             ]
#         )
#         return completion.choices[0].message.content
#     # debugging
#     except openai.APIConnectionError as e:
#         error_details = f"Connection error: {e}\n"
#         error_details += f"API Key (first 5 chars): {OPENAI_API_KEY[:5]}...\n"
#         try:
#             response = requests.get("https://api.openai.com/v1/engines", timeout=5)
#             error_details += f"OpenAI API reachable: {response.status_code == 200}\n"
#         except requests.RequestException as req_e:
#             error_details += f"Error reaching OpenAI API: {req_e}\n"
#         try:
#             socket.create_connection(("www.google.com", 80))
#             error_details += "Internet connection: Available\n"
#         except OSError:
#             error_details += "Internet connection: Not available\n"
#         return error_details
#     except openai.APIError as e:
#         return f"OpenAI API returned an API Error: {e}"
#     except openai.RateLimitError as e:
#         return f"OpenAI API request exceeded rate limit: {e}"
#     except Exception as e:
#         return f"Unexpected error: {e}"


def format_task(target, og_target, corporate_initiative, name, dri, is_overdue=False):
    """formats the progress, plan, and overdue tasks"""
    if is_overdue:  # task is overdue
        overdue = "OVERDUE: "
    else:
        overdue = ""
    bolded_initiative = f"<b>{corporate_initiative}</b>"
    if og_target:  # if an original target date exists, include it
        return f"{overdue}{target} ({og_target}) {bolded_initiative}: {name} [{dri}]"
    else:
        return f"{overdue}{target} {bolded_initiative}: {name} [{dri}]"


def format_problem_task(name, comment):
    """formats problem tasks"""
    bolded_name = f"<b>{name}</b>"
    if comment:  # comment for why problem task is blocked, exists
        return f"{bolded_name}- {comment}"
    else:
        return f"{bolded_name}- no comments"


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
        plan_section = df[((df['Status'] != 'Completed') & (df['Status'] != 'Canceled')) &
                          ((df['Target Date'] > today) & (df['Target Date'] <= two_months))]
        blocked_section = df[df['Status'] == 'Blocked']
        overdue_section = df[(df['Target Date'] <= today) * (df['Status'] != 'Completed')]

        # sorting tasks by completion date or target date
        progress = progress_section.sort_values(by='Complete Date')
        plan = plan_section.sort_values(by='Target Date')
        blocked = blocked_section.sort_values(by='Target Date')
        overdue = overdue_section.sort_values(by='Target Date')

        # formatting tasks for PPP
        progress_tasks = [
            format_task(row['Complete Date'].strftime('%m/%d'),
                        row['Target Date'].strftime('%m/%d') if pd.notna(row['Target Date']) else None,
                        row['Corporate Initiative'],
                        row['Project Name'],
                        row['Project DRI'])
            for index, row in progress.iterrows()
        ]
        plan_tasks = [
            format_task(row['Target Date'].strftime('%m/%d'),
                        row['Original Target Date'].strftime('%m/%d')
                        if has_og_target_date and pd.notna(row['Original Target Date']) else None,
                        row['Corporate Initiative'],
                        row['Project Name'],
                        row['Project DRI'])
            for index, row in plan.iterrows()
        ]
        problem_tasks = [
            format_problem_task(row['Project Name'],
                                row['Comments']
                                if has_comments and pd.notna(row['Comments']) else None)
            for index, row in blocked.iterrows()
        ] + [
            format_task(row['Target Date'].strftime('%m/%d'),
                        row['Original Target Date'].strftime('%m/%d')
                        if has_og_target_date and pd.notna(row['Original Target Date']) else None,
                        row['Corporate Initiative'],
                        row['Project Name'],
                        row['Project DRI'],
                        is_overdue=True)
            for index, row in overdue.iterrows()
        ]

        # formatting PPP
        ppp = (
            "<b>Progress [Last Week]</b> <br><br>" +
            "<br>".join(f"- {task}" for task in progress_tasks) + "<br><br>" +
            "<b>Plans [Next Two Months]</b> <br><br>" +
            "<br>".join(f"- {task}" for task in plan_tasks) + "<br><br>" +
            "<b>Progress [Ongoing]</b> <br><br>" +
            "<br>".join(f"- {task}" for task in problem_tasks) + "<br><br>"
        )

        # setting up openai prompts
        system_prompt = (
            "A PPP covers three sections: \n"
            "- Progress: tasks completed within the last week \n"
            "- Plans: tasks to be completed within the next two months \n"
            "- Problems: ongoing blocked tasks or tasks that are overdue \n"
            "**Summarize the PPP report by: \n"
            "- summarizing similar tasks into one task \n"
            "- describing tasks at an executive level \n"
            "**Maintain the current PP formatting. Only rephrase."
        )
        user_prompt = ppp

        # calling openai
        return ppp

    except Exception as e:
        return str(e)
