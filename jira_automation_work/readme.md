# ReadMe: Automate Jira Issue Creation from an Excel File

## Overview

This script automates the process of creating issues in Jira based on data from an Excel file. It supports creating Epics, Assignments, Tasks, and Subtasks while maintaining hierarchical relationships between them. The script uses the `JIRA` library, `pandas`, and `openpyxl` to interact with Jira and read/write Excel files.

## Prerequisites

1. Python 3.x installed.
2. Required Python libraries: `pandas`, `openpyxl`, `jira`, `argparse`.
3. A valid Jira account with API token access.
4. An input Excel file with a specific structure.

## Features

- Connects to Jira using a username and API token.
- Reads an Excel file and validates data.
- Creates Epics, Assignments, Tasks, and Subtasks with hierarchy support.
- Retries on failure to create issues with a randomized back-off.
- Updates the Excel file with generated Jira issue keys and marks completed entries.

## Setup and Installation

1. **Clone the repository or copy the script** to your project directory.
2. **Install required Python packages**:
   ```bash
   pip install pandas openpyxl jira argparse
   ```

3. **Ensure your Excel file has the following columns**:
   - `Hierarchy`: Specifies the type of issue (e.g., Epic, Assignment, Task, Subtask).
   - `Title`: The title of the issue.
   - `Project`: The project key in Jira.
   - `Key`: (Optional) Pre-existing Jira issue key (used for updating).
   - `Checklist`: (Optional) Marks completion status.

## Usage

### Command-Line Arguments

- `-u` or `--username`: Your Jira username (email).
- `-t` or `--token`: Your Jira API token.
- `-f` or `--file`: Path to the Excel file with issue data.

### Example Command

```bash
python automate_jira_creation.py -u your_email@example.com -t your_api_token -f path/to/excel_file.xlsx
```

## Script Workflow

1. **Connects to Jira**:
   - Authenticates using provided credentials and connects to the Jira server.
2. **Reads the Excel file**:
   - Validates the presence of required columns and reads data into a DataFrame.
3. **Processes each row**:
   - Skips rows marked as "Checked".
   - Creates issues based on the `Hierarchy` type:
     - **Epic**: Top-level issues.
     - **Assignment**: Sub-level issues under Epics.
     - **Task** and **Subtask**: Issues nested under parents.
   - Handles errors during creation with retries.
4. **Saves updates**:
   - Writes the updated DataFrame back to the Excel file, replacing the existing sheet.

## Excel File Requirements

Ensure your input Excel file has the following structure:

| Hierarchy   | Title             | Project | Key  | Checklist | Labels     | Parent       | Description          |
|-------------|-------------------|---------|------|-----------|------------|--------------|----------------------|
| Epic        | Epic Title        | PROJ    |      |           | epic-label |              | Description of epic  |
| Assignment  | Assignment Title  | PROJ    |      |           |            | Epic Title   | Assignment details   |
| Task        | Task Title        | PROJ    |      |           | task-label | Assignment Title | Task description |
| Subtask     | Subtask Title     | PROJ    |      |           | subtask-label | Task Title | Subtask details |

### Columns Explanation

- **Hierarchy**: Defines the type of issue.
- **Title**: Issue title in Jira.
- **Project**: Jira project key.
- **Key**: Jira issue key for existing issues (updated during execution).
- **Checklist**: Marks completion status. Set to "Checked" after creation.
- **Labels**: (Optional) Comma-separated labels for the issue.
- **Parent**: Title of the parent issue for hierarchical relationships.
- **Description**: (Optional) Details of the issue.

## Error Handling

- If an issue fails to be created after three retries, it is logged with its title, type, and error message.
- The script exits gracefully if there are errors in reading the Excel file or connecting to Jira.

## Output

- The script prints logs to the console indicating the progress of issue creation.
- Updates the input Excel file to mark completed issues and store new Jira keys.
- Provides a summary of any issues that failed to be created.

## Notes

- Customize the `JIRA_SERVER` variable with your Jira server URL.
- The script includes a back-off mechanism (`sleep`) to handle rate limits and transient errors.

## Troubleshooting

- **Connection Issues**: Ensure the `JIRA_SERVER` URL is correct and reachable.
- **Excel File Not Found**: Verify the `file_path` points to an existing file.
- **Missing Columns**: Ensure the Excel file has all required columns (`Hierarchy`, `Title`, etc.).

## Example Output

```text
Connected to Jira successfully.
Excel file read successfully.
Created Epic: Epic Title (Key: EPIC-123)
Created Task: Task Title (Key: TASK-456) under Epic: Epic Title (Key: EPIC-123)
Updated Excel file 'path/to/excel_file.xlsx' saved successfully.
Some issues failed to create. Details:
Title: Subtask Example, Issue Type: Subtask, Error: Error message here
```

## Future Enhancements

- Add support for custom field mappings.
- Implement logging to a file for better traceability.
- Integrate with additional project management tools for extended functionality.

## License

This script is available under the MIT License.