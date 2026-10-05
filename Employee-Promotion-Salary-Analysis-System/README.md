# Employee Promotion & Salary Analysis System

A Python business automation application that processes employee records from Excel, evaluates promotion eligibility, performs salary-position analysis, and generates professional reports.

## Why I built this

This project demonstrates practical Python skills by automating a business workflow involving employee data validation, business rules, reporting, and Excel document processing.

## Features

- Load employee records from Excel
- Validate employee IDs and email addresses
- Search employee records
- Evaluate promotion eligibility
- Compare current salary against a configurable benchmark
- Generate Excel results
- Generate TXT summary reports
- Generate formatted Word reports
- Handle invalid data without stopping the entire process
- Use a clean project structure and relative file paths

## Promotion Rules

An employee is marked **Eligible** when all of the following are true:

- At least 3 years of service
- Performance score of at least 65
- Attendance of at least 80
- No more than 3 warnings
- Training status is `Completed`

## Technology

- Python
- OpenPyXL
- python-docx
- Regular expressions
- File and path management
- Excel automation
- Report generation

## Project Structure

```text
Employee-Promotion-Salary-Analysis-System/
├── data/
│   └── EmployeeData.xlsx
├── docs/
├── reports/
├── src/
│   └── main.py
├── .gitignore
├── README.md
└── requirements.txt
```

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/Rowland112/Employee-Promotion-Salary-Analysis-System.git
cd Employee-Promotion-Salary-Analysis-System
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the application

```bash
python src/main.py
```

## Generated Reports

When option 3 is selected, the application creates:

- `reports/PromotionResults.xlsx`
- `reports/PromotionReport.txt`
- `reports/PromotionReport.docx`

Generated reports are excluded from version control so the repository stays clean.

## Salary Benchmark Note

The salary benchmark in this portfolio version is a configurable demonstration value. It is intentionally not presented as authoritative market salary data. A future version can connect to a verified salary dataset or approved API.

## Future Improvements

- Add a graphical user interface
- Add SQLite database support
- Add automated tests with pytest
- Add charts and dashboard summaries
- Add role-based access
- Add configurable promotion rules
- Connect to a verified salary data source
- Add CI with GitHub Actions

## Author

**Rowland Osei Frimpong**

Computer Programming Graduate  
Georgian College
