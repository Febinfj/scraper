"""
=============================================================================
Career Scraper - Task 1: Scrape Company Career Page for Job Listings
Powered by Supe AI & MuLearn
=============================================================================
Author: Febin Jacob
Objective:
    Extract job postings from a public company/career portal (Internshala),
    parse key listing details, handle pagination and missing data gracefully,
    and export the structured data into a professionally styled Excel file (.xlsx).

Required Output Columns:
    - JobTitle: Title of the job position
    - Location: City, state, or work-from-home status
    - ExperienceRequired: Years or level of experience required (blank if not specified)
    - SkillsRequired: Comma-separated list of required skills
    - Salary: Offered salary / CTC range (blank if not specified)
    - JobURL: Direct link to the full job listing
    - JobDescriptionSummary: Summary or key responsibilities of the role

Libraries:
    - requests (HTTP requests to download web pages)
    - beautifulsoup4 (HTML parsing and DOM extraction)
    - pandas (Data organization and DataFrame management)
    - openpyxl (Excel export and professional styling)
=============================================================================
"""

import argparse
import logging
import sys
import time
from urllib.parse import urljoin
from typing import List, Dict, Optional, Tuple

# Ensure stdout and stderr support UTF-8 on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("CareerScraper")

# Base URLs and default configurations
BASE_URL = "https://internshala.com"
JOBS_URL = "https://internshala.com/jobs/"
DEFAULT_OUTPUT_FILE = "Internshala_Jobs.xlsx"
DEFAULT_MAX_PAGES = 3
REQUEST_TIMEOUT = 15
DEFAULT_DELAY = 1.5

# Standard browser headers to avoid anti-scraping / bot blocks
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://internshala.com/",
}


def create_session() -> requests.Session:
    """
    Initializes and returns a requests Session pre-configured with headers.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def clean_text(text: Optional[str]) -> str:
    """
    Normalizes whitespace and removes excessive linebreaks from extracted text.
    """
    if not text:
        return ""
    # Normalize multiple whitespace / newline characters into single spaces
    return " ".join(text.split()).strip()


def extract_job_card(card: BeautifulSoup, session: requests.Session, fetch_details: bool = False) -> Optional[Dict[str, str]]:
    """
    Extracts required fields from an individual job card element.
    Handles missing fields gracefully with try-except blocks, ensuring
    blank values are set when data points are absent.
    """
    try:
        # 1. Job Title & Detail URL
        title = ""
        job_url = ""
        try:
            title_elem = card.find("a", class_=["job-title-href", "view_detail_button"])
            if title_elem:
                title = clean_text(title_elem.get_text())
                href = title_elem.get("href", "")
                if href:
                    job_url = urljoin(BASE_URL, href)
        except Exception as e:
            logger.debug(f"Error extracting title/url: {e}")

        # If there is no title, this is likely not a valid job card
        if not title:
            return None

        # 2. Location
        location = ""
        try:
            loc_elem = card.find("p", class_="locations")
            if loc_elem:
                location = clean_text(loc_elem.get_text())
            elif card.find("span", class_="location_link"):
                location = clean_text(card.find("span", class_="location_link").get_text())
        except Exception as e:
            logger.debug(f"Error extracting location: {e}")

        # 3. Salary (Leave blank if not available)
        salary = ""
        try:
            sal_icon = card.find("i", class_="ic-16-money")
            if sal_icon and sal_icon.parent:
                # Internshala has desktop and mobile spans for stipend/salary
                mob_span = sal_icon.parent.find("span", class_="mobile")
                desk_span = sal_icon.parent.find("span", class_="desktop")
                if mob_span and clean_text(mob_span.get_text()):
                    salary = clean_text(mob_span.get_text())
                elif desk_span and clean_text(desk_span.get_text()):
                    salary = clean_text(desk_span.get_text())
                else:
                    salary = clean_text(sal_icon.parent.get_text())
            elif card.find("span", class_="stipend"):
                salary = clean_text(card.find("span", class_="stipend").get_text())
        except Exception as e:
            logger.debug(f"Error extracting salary: {e}")

        # 4. Experience Required (Leave blank if not specified)
        experience = ""
        try:
            exp_icon = card.find("i", class_="ic-16-briefcase")
            if exp_icon and exp_icon.parent:
                experience = clean_text(exp_icon.parent.get_text())
            elif card.find("div", class_="job-experience-item"):
                exp_elem = card.find("div", class_="job-experience-item")
                experience = clean_text(exp_elem.get_text())
        except Exception as e:
            logger.debug(f"Error extracting experience: {e}")

        # 5. Skills Required (Comma-separated string of skills)
        skills_str = ""
        try:
            skill_elems = card.find_all("div", class_="job_skill")
            if skill_elems:
                skills_list = [clean_text(s.get_text()) for s in skill_elems if clean_text(s.get_text())]
                skills_str = ", ".join(skills_list)
        except Exception as e:
            logger.debug(f"Error extracting skills: {e}")

        # 6. Job Description Summary
        # Card contains a teaser summary under 'about_job'. If not present or if
        # fetch_details is explicitly requested, we can fetch from the detail page.
        summary = ""
        try:
            about_job = card.find("div", class_="about_job")
            if about_job:
                txt_div = about_job.find("div", class_="text")
                raw_summary = clean_text(txt_div.get_text() if txt_div else about_job.get_text())
                # Clean redundant prefix if present
                if raw_summary.lower().startswith("key responsibilities:"):
                    raw_summary = raw_summary[len("key responsibilities:"):].strip()
                summary = raw_summary
        except Exception as e:
            logger.debug(f"Error extracting card summary: {e}")

        # Deep fetch fallback if summary is missing on the card
        if not summary and job_url and fetch_details:
            try:
                resp = session.get(job_url, timeout=REQUEST_TIMEOUT)
                if resp.status_code == 200:
                    detail_soup = BeautifulSoup(resp.text, "html.parser")
                    desc_elem = (
                        detail_soup.find("div", class_="text-container") or
                        detail_soup.find("div", class_="job_details")
                    )
                    if desc_elem:
                        full_desc = clean_text(desc_elem.get_text())
                        # Keep a concise summary (first 300 characters)
                        summary = full_desc[:300] + ("..." if len(full_desc) > 300 else "")
            except Exception as e:
                logger.debug(f"Could not fetch detail page {job_url}: {e}")

        return {
            "JobTitle": title,
            "Location": location,
            "ExperienceRequired": experience,
            "SkillsRequired": skills_str,
            "Salary": salary,
            "JobURL": job_url,
            "JobDescriptionSummary": summary,
        }

    except Exception as e:
        logger.warning(f"Unexpected error parsing job card: {e}")
        return None


def scrape_page(page_url: str, session: requests.Session, fetch_details: bool = False) -> Tuple[List[Dict[str, str]], Optional[str]]:
    """
    Scrapes a single page of job listings.
    Returns:
        (jobs_list, next_page_url)
    """
    logger.info(f"Fetching URL: {page_url}")
    try:
        response = session.get(page_url, timeout=REQUEST_TIMEOUT)
        if response.status_code != 200:
            logger.error(f"Failed to fetch page (HTTP {response.status_code}): {page_url}")
            return [], None
    except requests.exceptions.RequestException as e:
        logger.error(f"Network error while requesting {page_url}: {e}")
        return [], None

    soup = BeautifulSoup(response.text, "html.parser")
    cards = soup.find_all("div", class_="individual_internship")
    logger.info(f"Found {len(cards)} job cards on this page.")

    jobs = []
    for card in cards:
        job_data = extract_job_card(card, session, fetch_details=fetch_details)
        if job_data:
            jobs.append(job_data)

    # Detect Next Page URL for pagination
    next_page_url = None
    try:
        # Check isLastPage input
        last_page_input = soup.find("input", id="isLastPage")
        is_last = last_page_input and last_page_input.get("value") == "1"

        if not is_last:
            # Check pagination div
            page_div = soup.find("div", class_="page_number")
            current_page_span = soup.find("span", id="pageNumber")
            current_page_num = int(current_page_span.get_text(strip=True)) if current_page_span and current_page_span.get_text(strip=True).isdigit() else 1

            target_page_num = current_page_num + 1

            if page_div:
                # Look for link with data-page equal to target_page_num
                next_link = page_div.find("a", attrs={"data-page": str(target_page_num)})
                if next_link and next_link.get("href"):
                    next_page_url = urljoin(BASE_URL, next_link["href"])

            # Fallback pagination link constructor
            if not next_page_url and len(cards) > 0:
                next_page_url = f"{JOBS_URL}page-{target_page_num}/"

    except Exception as e:
        logger.debug(f"Error resolving next page URL: {e}")

    return jobs, next_page_url


def scrape_jobs(max_pages: int = DEFAULT_MAX_PAGES, delay: float = DEFAULT_DELAY, fetch_details: bool = False) -> List[Dict[str, str]]:
    """
    Orchestrates the scraping process across multiple pages.
    Handles pagination, rate limiting delays, and deduplication.
    """
    session = create_session()
    all_jobs: List[Dict[str, str]] = []
    seen_urls = set()
    current_url = JOBS_URL
    page_count = 0

    print("=" * 80)
    print(" [*] Starting Career Scraper (Target: Internshala Public Career Portal)")
    print(f" [*] Target Pages: {max_pages} | Politeness Delay: {delay}s")
    print("=" * 80)

    while current_url and page_count < max_pages:
        page_count += 1
        print(f"\n[Page {page_count}/{max_pages}] Accessing {current_url} ...")

        jobs_on_page, next_url = scrape_page(current_url, session, fetch_details=fetch_details)

        if not jobs_on_page:
            logger.warning(f"No jobs extracted from page {page_count}. Stopping pagination.")
            break

        # Deduplicate
        added_count = 0
        for job in jobs_on_page:
            url_key = job.get("JobURL") or (job.get("JobTitle"), job.get("Location"))
            if url_key not in seen_urls:
                seen_urls.add(url_key)
                all_jobs.append(job)
                added_count += 1

        print(f"  [+] Extracted {len(jobs_on_page)} jobs ({added_count} unique new additions).")
        print(f"  [+] Running Total: {len(all_jobs)} jobs collected.")

        if not next_url:
            print("  [-] No next page detected (end of listings reached).")
            break

        current_url = next_url

        # Polite delay to avoid overloading server
        if page_count < max_pages:
            time.sleep(delay)

    return all_jobs


def export_to_excel(jobs: List[Dict[str, str]], filename: str = DEFAULT_OUTPUT_FILE) -> bool:
    """
    Exports scraped job data into a structured and professionally styled Excel file (.xlsx).
    Ensures:
      - Required column names and ordering match task requirements.
      - Pandas index column is excluded.
      - Styled header row with deep navy background and white bold text.
      - Auto-fit column widths with padding for optimal readability.
      - Text alignment and thin cell borders.
      - Freeze pane on the header row and auto-filter enabled.
    """
    if not jobs:
        logger.error("No jobs data to export.")
        return False

    # Define exact required column headers in order
    columns_order = [
        "JobTitle",
        "Location",
        "ExperienceRequired",
        "SkillsRequired",
        "Salary",
        "JobURL",
        "JobDescriptionSummary",
    ]

    df = pd.DataFrame(jobs)

    # Ensure all required columns exist in the DataFrame
    for col in columns_order:
        if col not in df.columns:
            df[col] = ""

    # Reorder columns explicitly
    df = df[columns_order]

    # Save initial DataFrame to Excel without index
    try:
        df.to_excel(filename, index=False, engine="openpyxl")
    except Exception as e:
        logger.error(f"Failed to write initial Excel file with pandas: {e}")
        return False

    # Apply openpyxl professional styling
    try:
        wb = load_workbook(filename)
        ws = wb.active
        ws.title = "Job Listings"

        # Styling definitions
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # Navy blue
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Calibri", size=10)
        link_font = Font(name="Calibri", size=10, color="0000EE", underline="single")

        thin_side = Side(border_style="thin", color="E2E8F0")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        data_alignment_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
        data_alignment_center = Alignment(horizontal="center", vertical="top")

        # Format header row
        ws.row_dimensions[1].height = 28
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = cell_border

        # Enable Auto-Filter on header row
        last_col_letter = get_column_letter(len(columns_order))
        ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

        # Freeze the top header row so it stays visible on scroll
        ws.freeze_panes = "A2"

        # Format data rows
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            ws.row_dimensions[row_idx].height = 22
            for col_idx, cell in enumerate(row, start=1):
                cell.border = cell_border
                cell.font = regular_font

                col_name = columns_order[col_idx - 1]
                if col_name in ["ExperienceRequired", "Location"]:
                    cell.alignment = data_alignment_center
                elif col_name == "JobURL":
                    cell.alignment = data_alignment_left
                    cell.font = link_font
                else:
                    cell.alignment = data_alignment_left

        # Adjust column widths automatically based on maximum content length
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            col_name = col[0].value

            for cell in col:
                val = str(cell.value or "")
                if len(val) > max_len:
                    max_len = len(val)

            # Assign sensible min/max constraints for aesthetics
            if col_name == "JobDescriptionSummary":
                assigned_width = 45
            elif col_name == "SkillsRequired":
                assigned_width = 35
            elif col_name == "JobURL":
                assigned_width = 35
            elif col_name == "JobTitle":
                assigned_width = max(max_len + 4, 25)
            elif col_name == "Salary":
                assigned_width = max(max_len + 4, 22)
            else:
                assigned_width = max(max_len + 4, 16)

            ws.column_dimensions[col_letter].width = min(assigned_width, 60)

        wb.save(filename)
        logger.info(f"Excel file successfully formatted and saved to '{filename}'.")
        return True

    except Exception as e:
        logger.error(f"Error styling Excel workbook: {e}")
        # Even if styling throws an issue, pandas already exported the valid raw xlsx file
        return True


def parse_arguments() -> argparse.Namespace:
    """
    Parses command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description="Scrape company career job postings into a structured Excel workbook."
    )
    parser.add_argument(
        "--pages",
        "-p",
        type=int,
        default=DEFAULT_MAX_PAGES,
        help=f"Number of pages to scrape (default: {DEFAULT_MAX_PAGES})",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        default=DEFAULT_OUTPUT_FILE,
        help=f"Output Excel file name (default: {DEFAULT_OUTPUT_FILE})",
    )
    parser.add_argument(
        "--delay",
        "-d",
        type=float,
        default=DEFAULT_DELAY,
        help=f"Delay in seconds between requests (default: {DEFAULT_DELAY}s)",
    )
    parser.add_argument(
        "--fetch-details",
        action="store_true",
        help="Fetch individual job detail pages to enrich job description summaries",
    )
    return parser.parse_args()


def main():
    args = parse_arguments()

    print("\n" + "=" * 80)
    print("   CAREER SCRAPER - TASK 1: SCRAPE COMPANY CAREER PAGE FOR JOB LISTINGS")
    print("=" * 80)

    start_time = time.time()

    # Step 1: Scrape job listings
    jobs = scrape_jobs(max_pages=args.pages, delay=args.delay, fetch_details=args.fetch_details)

    if not jobs:
        print("\n[-] Scraping completed but no job records were collected.")
        sys.exit(1)

    # Step 2: Export to Excel
    print("\n" + "-" * 80)
    print(f"[*] Exporting {len(jobs)} jobs to Excel: '{args.output}' ...")
    success = export_to_excel(jobs, filename=args.output)
    elapsed = time.time() - start_time

    if success:
        print("\n" + "=" * 80)
        print("[+] SUCCESS: Web scraping and Excel generation completed!")
        print("=" * 80)
        print(f" * Output File   : {args.output}")
        print(f" * Total Jobs    : {len(jobs)}")
        print(f" * Pages Scraped : {args.pages}")
        print(f" * Time Elapsed  : {elapsed:.2f} seconds")
        print("\nPreview of first 5 jobs:")
        df_preview = pd.DataFrame(jobs)[["JobTitle", "Location", "Salary", "ExperienceRequired"]]
        print(df_preview.head(5).to_string(index=False))
        print("=" * 80 + "\n")
    else:
        print("\n[-] Error occurred while saving the Excel file.")
        sys.exit(1)


if __name__ == "__main__":
    main()
