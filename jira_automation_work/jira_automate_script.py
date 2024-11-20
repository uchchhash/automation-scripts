import argparse
import pandas as pd
from jira import JIRA
from time import sleep
import random
from openpyxl import load_workbook

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Automate Jira issue creation from an Excel file.")
parser.add_argument('-u', '--username', type=str, required=True, help="Jira username (email)")
parser.add_argument('-t', '--token', type=str, required=True, help="Jira API token")
parser.add_argument('-f', '--file', type=str, required=True, help="Path to the input Excel file")
args = parser.parse_args()

# Assign arguments to variables
USERNAME  = args.username
API_TOKEN = args.token
file_path = args.file

# Provide the Jira server URL (replace with actual URL if needed)
JIRA_SERVER = "provide-jira-server-url-here"

# Connect to Jira
try:
    jira = JIRA(server=JIRA_SERVER, basic_auth=(USERNAME, API_TOKEN))
    print("Connected to Jira successfully.")
except Exception as e:
    print(f"Failed to connect to Jira: {e}")
    exit(1)

# Read the existing Excel file
try:
    data = pd.read_excel(file_path)
    print("Excel file read successfully.")
except FileNotFoundError:
    print("Excel file not found. Please check the file path.")
    exit(1)
except Exception as e:
    print(f"Failed to read Excel file: {e}")
    exit(1)

# Validate input data
required_columns = ["Hierarchy", "Title", "Project", "Key", "Checklist"]
if not all(column in data.columns for column in required_columns):
    print("Excel file is missing one or more required columns.")
    exit(1)

# Initialize 'Key' column to store issue keys and 'Checklist' if not already present
if "Key" not in data.columns:
    data["Key"] = ""
if "Checklist" not in data.columns:
    data["Checklist"] = ""

# Ensure correct data types for 'Key' and 'Checklist' columns
data["Key"] = data["Key"].astype(str)
data["Checklist"] = data["Checklist"].astype(str)

# Store created issue keys to handle hierarchy
issue_keys = {}
failed_issues = []

def create_jira_issue(issue_type, title, project_key, labels=None, parent_key=None, description=None, retries=3):
    for attempt in range(retries):
        try:
            issue_dict = {
                "project": {"key": project_key},
                "summary": title,
                "issuetype": {"name": issue_type},
                "labels": labels.split(", ") if pd.notna(labels) else []
            }

            if description:
                issue_dict["description"] = description
            if parent_key and issue_type in ["Task", "Subtask"]:
                issue_dict["parent"] = {"key": parent_key}

            new_issue = jira.create_issue(fields=issue_dict)
            return new_issue.key
        except Exception as e:
            print(f"Attempt {attempt + 1} failed to create {issue_type} '{title}': {e}")
            if attempt < retries - 1:
                sleep_time = random.uniform(2, 5)
                print(f"Retrying in {sleep_time:.1f} seconds...")
                sleep(sleep_time)
            else:
                failed_issues.append((title, issue_type, str(e)))
                return None

# Iterate through the data and create issues
for index, row in data.iterrows():
    # Skip rows that already have 'Checked' in the Checklist column
    if row["Checklist"] == "Checked":
        print(f"Skipping already checked issue: {row['Title']}")
        # Store the already checked issue key (so that it's available for subtasks)
        issue_keys[row["Title"]] = row["Key"]
        continue

    issue_type = row["Hierarchy"]
    title = row["Title"]
    project_key = row["Project"]
    labels = row.get("Labels", "")
    parent_title = row.get("Parent", "")
    description = row.get("Description", "")

    new_issue_key = None

    if issue_type == "Epic":
        new_issue_key = create_jira_issue("Epic", title, project_key, labels, description=description)
        if new_issue_key:
            issue_keys[title] = new_issue_key
            print(f"Created Epic: {title} (Key: {new_issue_key})")

    elif issue_type == "Assignment":
        epic_key = issue_keys.get(parent_title)
        if not epic_key:
            print(f"Warning: Parent epic '{parent_title}' not found for Assignment '{title}'")
        else:
            new_issue_key = create_jira_issue("Task", title, project_key, labels, parent_key=epic_key, description=description)
            if new_issue_key:
                issue_keys[title] = new_issue_key
                print(f"Created Task: {title} (Key: {new_issue_key}) under Epic: {parent_title} (Key: {epic_key})")

    elif issue_type == "Task":
        parent_key = issue_keys.get(parent_title)
        if parent_key:
            new_issue_key = create_jira_issue("Task", title, project_key, labels, parent_key=parent_key, description=description)
            if new_issue_key:
                issue_keys[title] = new_issue_key
                print(f"Created Task: {title} (Key: {new_issue_key}) under Task: {parent_title} (Key: {parent_key})")
        else:
            print(f"Warning: Parent '{parent_title}' not found for Task '{title}'")

    elif issue_type == "Subtask":
        parent_key = issue_keys.get(parent_title)
        if parent_key:
            new_issue_key = create_jira_issue("Subtask", title, project_key, labels, parent_key=parent_key, description=description)
            if new_issue_key:
                issue_keys[title] = new_issue_key
                print(f"Created Subtask: {title} (Key: {new_issue_key}) under Task: {parent_title} (Key: {parent_key})")
        else:
            print(f"Warning: Parent '{parent_title}' not found for Subtask '{title}'")

    # Save the generated issue key to the DataFrame and mark the checklist as checked
    if new_issue_key:
        # Only update 'Key' and 'Checklist' if the issue was newly created
        if row["Key"] == "" or row["Checklist"] != "Checked":
            data.at[index, "Key"] = new_issue_key
            data.at[index, "Checklist"] = "Checked"
            print(f"Key for {issue_type} '{title}' saved: {new_issue_key}")

# Save updated DataFrame to the existing Excel file
try:
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        data.to_excel(writer, index=False, sheet_name="Sheet1")
    print(f"Updated Excel file '{file_path}' saved successfully.")
except Exception as e:
    print(f"Failed to save the updated Excel file: {e}")

# Report any failed issues
if failed_issues:
    print("Some issues failed to create. Details:")
    for failed in failed_issues:
        print(f"Title: {failed[0]}, Issue Type: {failed[1]}, Error: {failed[2]}")
