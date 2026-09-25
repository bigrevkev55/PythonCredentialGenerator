# Python Credential Generator

A Python-based credential and diploma generation application developed to automate the creation of student diploma documents.

The application uses student credential data to generate formatted diploma documents from Microsoft Word templates, reducing the amount of manual processing required when producing diplomas and replacement credentials.

## Features

- Generates diploma documents using Python
- Uses Microsoft Word templates to maintain consistent diploma formatting
- Supports production and testing environments
- Generates individual student diploma documents
- Uses SQLite for local credential data storage
- Includes SQL scripts for maintaining and updating credential data
- Provides a graphical user interface for credential generation
- Supports multiple credential and degree types

## Repository Structure

```text
PythonCredentialGenerator/
│
├── NSCC Credential Generator PROD/
│   ├── AR_LOGO.jpg
│   ├── NSCCDiplomaTemplate.docx
│   ├── main.py
│   └── treebase.ini
│
├── NSCC Credential Generator TEST/
│   ├── AR_LOGO.jpg
│   ├── TESTmain.py
│   ├── TestNsccDiploma.docx
│   └── treebase.ini
│
├── Sample Files and Templates/
│   ├── DoD user Interface.PNG
│   └── PythonCredentialGenerator.py
│
├── Scripts for SQLite DB/
│   └── Graduates Update Script for SQLite DB.sql
│
├── .gitignore
└── README.md
```

## Technologies Used

- Python
- SQLite
- Microsoft Word document templates
- SQL
- Git / GitHub

## Production and Test Versions

The repository includes separate production and testing versions of the credential generator.

### Production

The production application and associated templates are located in:

```text
NSCC Credential Generator PROD/
```

### Testing

Testing files and templates are located in:

```text
NSCC Credential Generator TEST/
```

This allows changes to be tested without modifying the production version of the application.

## Data Privacy

The application is designed to work with local student credential data.

Production databases and other files containing protected or sensitive student information are intentionally excluded from this repository through `.gitignore`.

Generated credential output and other locally generated data files may also be excluded from source control.

Users implementing this application should ensure that student information is handled in accordance with applicable institutional policies and data-security requirements.

## Purpose

This project was developed to improve the efficiency and consistency of diploma and credential production by replacing portions of a manual document-generation process with a Python-based workflow.

It also serves as an example of using Python, SQL, SQLite, and Microsoft Office document automation to support higher-education administrative processes.

## Notes

This repository contains application source code, templates, sample files, and supporting SQL scripts.

Local databases and other sensitive operational data required by a production environment are not included and must be configured separately.

## Author

Kevin Thomas