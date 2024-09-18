
import json

# excel
import openpyxl

fake_data = """
[
{
    "name": "John Doe",
    "title": "Software Engineer",
    "email": "johnd@example.com",
    "experience": [
    {
        "company": "Goosoft Corporation",
        "role": "Senior Developer",
        "start_date": "2022-12-31",
        "end_date": "2022-12-31",
        "skills": ["Python", "JavaScript", "Django", "GCP", "Kubernetes"]
    },
    {
        "company": "Acme Corporation",
        "role": "Junior Developer",
        "start_date": "2020-01-01",
        "end_date": "2022-12-31",
        "skills": ["Python", "JavaScript", "Django"]
    },
    {
        "company": "TechCo",
        "role": "Full Stack Developer",
        "start_date": "2023-01-01",
        "end_date": "present",
        "skills": ["React", "Node.js", "SQL"]
    }
    ],
    "education": [
    {
        "degree": "Bachelor of Science",
        "major": "Computer Science",
        "university": "State University",
        "graduation_date": "2019-05-01"
    }
    ],
    "skills": ["Git", "Agile", "AWS"]
},
{
    "name": "Jane Smith",
    "title": "Data Analyst",
    "email": "jane_s@example.com",
    "experience": [
    {
        "company": "Goosoft Corporation",
        "role": "Senior Developer",
        "start_date": "2022-12-31",
        "end_date": "2022-12-31",
        "skills": ["Python", "JavaScript", "Django", "GCP", "Kubernetes"]
    },
    {
        "company": "Acme Corporation",
        "role": "Junior Developer",
        "start_date": "2020-01-01",
        "end_date": "2022-12-31",
        "skills": ["Python", "JavaScript", "Django"]
    },
    {
        "company": "TechCo",
        "role": "Full Stack Developer",
        "start_date": "2023-01-01",
        "end_date": "present",
        "skills": ["React", "Node.js", "SQL"]
    }
    ]
},
{
    "name": "Michael Johnson",
    "title": "Project Manager",
    "email": "Michael_j@example.com"
},
{
    "name": "Emily Davis",
    "title": "UI/UX Designer",
    "email": "emily_davis@example.com"
},
{
    "name": "Christopher Brown",
    "title": "Sales Representative",
    "email": "christopher_brown@example.com"
}
]
"""

# read from the json in this file
data = json.loads(fake_data)
# or read from a json file
# data = json.load(open("fake_data.json"))

def clean(text, make_lower=True):
    text = text.strip().replace(" ", "_")
    text = text.replace(",", "_")
    text = text.replace("@", "_at_")
    if make_lower:
        text = text.lower()
    return text



def employee_data_to_excel(data):
    # for each employee, create an excel file for each job

    def job_to_excel(employee, offset=0):
        # create an excel file for each job
        job = {}
        try:
            job['role'] = clean(employee["experience"][offset]["company"]) + "--" + \
                            clean(employee["experience"][offset]["role"])
            job['details'] = clean(employee["experience"][offset]["start_date"]) + "--" + \
                                clean(employee["experience"][offset]["end_date"])
            job['skills'] = clean("_".join(employee["experience"][offset]["skills"]), make_lower=False)

            filename = name_email + "-" + str(offset+1) + "--" + job['role'] + ".xlsx"

            # Create a new workbook
            workbook = openpyxl.Workbook()
            # Select the active worksheet
            sheet = workbook.active

            # Write the header row
            sheet['A1'] = 'Name'
            sheet['A2'] = 'Email'
            sheet['A3'] = 'Company'
            sheet['A4'] = 'Role'
            sheet['A5'] = 'Details'
            sheet['A6'] = 'Skills'
            # write the data
            sheet['B1'] = employee["name"]
            sheet['B2'] = employee["email"]
            sheet['B3'] = clean(employee["experience"][offset]["company"])
            sheet['B4'] = clean(employee["experience"][offset]["role"])
            sheet['B5'] = clean(employee["experience"][offset]["start_date"]) + " to " + \
                            clean(employee["experience"][offset]["end_date"])
            sheet['B6'] = clean("\n".join(employee["experience"][offset]["skills"]), make_lower=False)

            # Save the workbook
            workbook.save(filename)
            print(f"Created {filename}")

        except Exception as e:
            print(e)

    for employee in data:
        # print(employee["name"])
        # print(employee["title"])
        # print(employee["experience"])
        # print(employee["education"])
        
        name_email = clean(employee["name"])
        name_email += "--" + clean(employee["email"])
        
        # just the last 3 jobs
        for i in range(3):
            job_to_excel(employee, offset=i)

        # break


if __name__ == "__main__":

    employee_data_to_excel(data)

    # filenames look like this:
    # john_doe--johnd_at_example.com-1--goosoft_corporation--senior_developer.xlsx
    # john_doe--johnd_at_example.com-2--acme_corporation--junior_developer.xlsx
    # john_doe--johnd_at_example.com-3--techco--full_stack_developer.xlsx
