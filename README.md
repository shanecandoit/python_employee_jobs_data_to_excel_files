# Automating Resume Screening with Python

## Project Overview

This Python script automates the process of converting employee JSON data into individual Excel files, tailored for specific job roles. It efficiently extracts relevant information from the JSON data and organizes it into structured Excel sheets.

## Prerequisites

- Python: Ensure you have Python installed on your system.
- openpyxl: Install the openpyxl library using pip: pip install openpyxl

## Usage
Prepare JSON Data:

Create a JSON file containing employee data. The data should follow a specific structure, such as:
```JSON
{
  "employees": [
    {
      "name": "John Doe",
      "title": "Software Engineer",
      "experience": [
        // ... experience data
      ],
      "education": [
        // ... education data
      ]
    },
    // ... more employees
  ]
}
```


1. Run the Script:

Execute the Python script: python fake_employee_data.py

The script will process the JSON data and create individual Excel files for each employee, based on their top 3 job roles.

2. Script Structure

The script typically consists of the following steps:

  1. Load JSON Data: Reads the JSON file and parses its contents.

  2. Iterate Over Employees: Loops through each employee in the JSON data.

  3. Determine Top Job Roles: Identifies the top 3 job roles for the employee based on their experience or other criteria.

  4. Create Excel Files: For each top job role:

     - Creates a new Excel file.
     - Populates the file with relevant employee information, such as name, title, experience, and education.
     - Formats the Excel sheet as needed.

### Customization

- Job Role Prioritization: Modify the logic for determining top job roles to match your specific requirements.

- Excel Sheet Structure: Customize the layout and content of the Excel files to suit your needs.
- Data Validation: Implement data validation checks to ensure the integrity of the JSON data and Excel output.
- Error Handling?: DIY. Consider adding error handling mechanisms to gracefully handle potential exceptions, such as invalid JSON data or file I/O errors.
