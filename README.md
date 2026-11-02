# 💼 Career Scraper — Internshala Job Listings

A Python web scraper that extracts job listings from [Internshala](https://internshala.com/jobs/) and exports structured data into a formatted Excel workbook.

Developed for **Task 1: Scrape Company Career Page for Job Listings (200 Karma Points)** by **Supe AI & µLearn**.

## ✨ Features

- Scrapes job titles, locations, experience, skills, and salaries.
- Extracts job URLs and descriptions.
- Supports pagination and duplicate removal.
- Handles missing information gracefully.
- Exports data into a professionally formatted Excel file.
- Supports customizable scraping options.

## 🛠️ Tech Stack

- Python
- Requests
- BeautifulSoup4
- Pandas
- OpenPyXL

## 📁 Project Structure

```text
├── scraper.py
├── Internshala_Jobs.xlsx
├── requirements.txt
├── .gitignore
└── README.md
```

## 🚀 Getting Started

Clone the repository:

```bash
git clone <your-github-repo-url>
cd <repo-folder>
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the scraper:

```bash
python scraper.py
```

**Optional commands:**

```bash
python scraper.py --pages 3
python scraper.py --pages 2 --output Jobs.xlsx
python scraper.py --pages 2 --delay 2.0
python scraper.py --pages 1 --fetch-details
```

## 📊 Output

Generated Excel file: `Internshala_Jobs.xlsx`

Includes:

- Job Title
- Location
- Experience Required
- Skills Required
- Salary
- Job URL
- Job Description Summary

## 👨‍💻 Author

**Febin Jacob**

**Task:** µLearn – Supe AI Career Scraper (200 Karma Points)

*Built with Python 🐍*
