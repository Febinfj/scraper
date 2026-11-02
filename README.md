# 💼 Career Scraper - Task 1: Scrape Company Career Page for Job Listings

[![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![Status](https://img.shields.io/badge/Status-Completed-success.svg)]()
[![Powered by](https://img.shields.io/badge/Powered%20by-Supe%20AI-orange.svg)]()
[![MuLearn](https://img.shields.io/badge/%C2%B5Learn-200%20Karma%20Points-purple.svg)](https://mulearn.org/)

A robust Python web scraper that systematically extracts job postings from a public career portal (**Internshala**), parses key listing details (title, location, experience, salary, skills, URL, and summary), and exports the structured data into a professionally formatted Excel (`.xlsx`) workbook.

---

## 📌 Project Overview

This project was built to fulfill **Task 1: Scrape Company Career Page for Job Listings** (200 Karma Points) powered by **Supe AI** & **µLearn**.

The script programmatically traverses live career listings, inspects the DOM hierarchy using `BeautifulSoup4`, gracefully handles missing data fields (e.g. undisclosed salary or experience), navigates across paginated result sets, and styles the output workbook using `openpyxl`.

---

## 🏗️ Architecture & HTML Inspection

The target careers portal for this scraper is **Internshala Jobs** (`https://internshala.com/jobs/`). Using browser Developer Tools (F12), the DOM structure was analyzed to target the following container and item elements:

| Target Field | HTML Selector / Tag Structure | Extraction Logic / Fallback |
| :--- | :--- | :--- |
| **Job Card Container** | `div.individual_internship` | Main loop target for each listing card |
| **JobTitle** | `a.job-title-href`, `a.view_detail_button` | Extracts stripped text representing job title |
| **Location** | `p.locations`, `span.location_link` | Cleaned city/state or remote status |
| **ExperienceRequired** | `div.row-1-item` containing `i.ic-16-briefcase` | Extracts years/level; blank if unspecified |
| **SkillsRequired** | `div.job_skills div.job_skill` | Aggregated as comma-separated string |
| **Salary** | `div.row-1-item` containing `i.ic-16-money` | Parses salary range in CTC/per year; blank if absent |
| **JobURL** | `a.job-title-href[href]` | Joined with `https://internshala.com` |
| **JobDescriptionSummary** | `div.about_job div.text` / detail page `div.text-container` | First summary snippet or key responsibilities |
| **Pagination** | `div.page_number a[data-page]` / `input#isLastPage` | Detects next page and stops at the final page |

---

## 📊 Output Schema (Excel Columns)

The output Excel file (`Internshala_Jobs.xlsx`) contains the following required columns:

1. `JobTitle` — The official title of the job position.
2. `Location` — City, state, or work-from-home location.
3. `ExperienceRequired` — Required years or level of experience (e.g., "1 year(s)", "No experience required", or blank if not specified).
4. `SkillsRequired` — Comma-separated list of required technical or soft skills.
5. `Salary` — Compensation or salary range (e.g., "₹ 4,00,000 - 4,80,000 /year"), left blank if not specified.
6. `JobURL` — Absolute URL to the full job listing on the platform.
7. `JobDescriptionSummary` — Concise summary or excerpt of the job responsibilities.

### 🎨 Excel Styling Features
- **Header formatting:** Dark navy blue (`#1E3A8A`) background with bold white text.
- **Auto-Filter:** Enabled across all columns on Row 1.
- **Freeze Panes:** Top header row stays fixed while scrolling down.
- **Auto Column Width:** Dynamic column dimensions preventing text clipping.
- **Borders & Alignment:** Clean gridlines and vertical alignment for enhanced readability.

---

## 📁 Repository Structure

```text
├── scraper.py             # Complete Python scraping & export script
├── Internshala_Jobs.xlsx   # Output Excel file with scraped job records
├── requirements.txt       # Dependencies (requests, bs4, pandas, openpyxl)
├── .gitignore             # Standard Python gitignore
└── README.md              # Project documentation
```

---

## 🚀 Getting Started

### 1. Clone the Repository
```bash
git clone <your-github-repo-url>
cd <repo-folder>
```

### 2. Set Up a Virtual Environment (Recommended)
```bash
python -m venv venv

# On Windows:
venv\Scripts\activate

# On Linux/macOS:
source venv/bin/activate
```

### 3. Install Required Libraries
```bash
pip install -r requirements.txt
```

---

## 💻 Running the Scraper

Run the script using Python:

```bash
python scraper.py
```

### Command-Line Arguments

The script supports flexible CLI options:

```bash
# Scrape specific number of pages (e.g., 3 pages)
python scraper.py --pages 3

# Custom output Excel filename
python scraper.py --pages 2 --output My_Custom_Jobs.xlsx

# Adjust delay between page requests (politeness rate-limiting)
python scraper.py --pages 2 --delay 2.0

# Enable deep inspection of individual job detail pages
python scraper.py --pages 1 --fetch-details
```

| Argument | Shorthand | Default | Description |
| :--- | :--- | :--- | :--- |
| `--pages` | `-p` | `3` | Number of pagination pages to scrape |
| `--output` | `-o` | `Internshala_Jobs.xlsx` | Output Excel workbook filename |
| `--delay` | `-d` | `1.5` | Delay in seconds between requests |
| `--fetch-details` | | `False` | Deep fetch individual job pages for extended descriptions |

---

## 🧪 Sample Output Preview

| JobTitle | Location | ExperienceRequired | SkillsRequired | Salary |
| :--- | :--- | :--- | :--- | :--- |
| **Business Development Associate** | Mumbai | 1 year(s) | MS-Word, MS-PowerPoint, English Proficiency (Spoken) | ₹ 4,00,000 - 4,80,000 /year |
| **Sales Associate - Urgent Hiring** | Delhi, Gurgaon, Madurai | No experience required | Inside Sales, Business Development, Sales | ₹ 6,50,000 - 7,50,000 /year |
| **Marketing Executive** | Bangalore | No experience required | Lead Generation, B2B Sales, Market research | ₹ 3,60,000 - 5,00,000 /year |
| **Associate CAD Designer** | Gurgaon | 1 year(s) | Embroidery making, Self-learning, Troubleshooting | ₹ 3,60,005 - 4,50,010 /year |

---

## 🛡️ Error Handling & Robustness

- **Graceful Null Handling:** Missing salary, experience, or skills fields are left as empty strings (`""`), reflecting blank cells in Excel as specified by the task criteria.
- **Session & Headers:** Utilizes custom browser `User-Agent` headers and connection timeouts to prevent network stalls and 403 bot blocks.
- **Deduplication:** Keeps track of visited `JobURL` links to ensure duplicate listings across page boundaries are filtered out.
- **Pagination Boundary Check:** Stops automatically when the `isLastPage` indicator is reached or when no additional job cards are detected.
