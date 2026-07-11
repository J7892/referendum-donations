# Alberta Referendum Political Donations Tracker

An automated scraper and database tracker for political contributions related to Alberta's forthcoming referendum on separation. 

This system automatically checks the Elections Alberta Financial Disclosure portal for updates, downloads weekly contribution Excel reports for all registered groups, merges them into a unified, searchable CSV file, and emails a weekly report highlighting overall totals, weekly changes, new advertiser registrations, and notable donations ($1,000+).

## Data Sources
- **Main Portal**: [Elections Alberta Referendum Advertisers](https://efpublic.elections.ab.ca/efRTPAs.cfm?MID=TPAS_TR&TPATYPE=R)
- **Weekly Excel Reports**: Dynamically discovered and downloaded for each group.

## Features
- **Historical Report-Date Matching**: Elections Alberta's weekly Excel reports only show cumulative contributions *to date* without individual contribution dates. This script uses a multiset matching algorithm to compare snapshots week-to-week, assigning and preserving the weekly report date (e.g. `Jul 2, 2026`) when a contribution first appears.
- **Deduplication**: Ensures that contributions are not duplicated when cumulative reports are merged, while accurately preserving multiple legitimate contributions of the same amount from the same entity.
- **Premium Email Alerts**: Formats HTML emails with clear tables, statistics, and highlight badges for newly registered groups and large contributions (>= $5,000).

---

## Project Structure
```
.github/
  workflows/
    tracker.yml              # GitHub Actions automation schedule
data/
  state.json                 # Caches portal date and advertiser metadata
  referendum_donations.csv   # Unified database of all referendum donations
utils/
  __init__.py
  email_sender.py            # Secure SMTP email sender
scraper.py                   # Main scraping and diffing script
requirements.txt             # Python dependencies
README.md                    # Setup and usage guide
```

---

## GitHub Setup & Automation

To host this on GitHub and automate it:

1. **Create a GitHub Repository** and push all files in this project directory to it.
2. **Enable Write Permissions for GitHub Actions**:
   In your repository:
   - Go to **Settings** > **Actions** > **General**.
   - Under **Workflow permissions**, select **Read and write permissions**.
   - Click **Save**.
   *(This permits the GitHub Action bot to commit and push database updates back to your repository.)*

3. **Configure Email Secrets**:
   To receive weekly email updates, go to **Settings** > **Secrets and variables** > **Actions** and add the following **Repository secrets**:
   
   - `SMTP_SERVER`: The SMTP host of your email provider (e.g., `smtp.gmail.com` or `smtp.sendgrid.net`).
   - `SMTP_PORT`: The SMTP port (usually `587` for TLS or `465` for SSL).
   - `SMTP_USERNAME`: The email account/username to send through.
   - `SMTP_PASSWORD`: The password or app-specific password (strongly recommended for Gmail/Yahoo).
   - `EMAIL_TO`: The email address where report alerts should be sent.
   - `EMAIL_FROM` (Optional): The sender address (defaults to `SMTP_USERNAME` if not provided).

---

## Automation Schedule

The GitHub Action is pre-configured in [.github/workflows/tracker.yml](.github/workflows/tracker.yml) to run on the following schedule (adjusted to Mountain Time for Alberta):
- **Thursdays (Afternoon & Evening)**: Runs **every hour** (from 12:00 PM MDT/MST onwards) to check for early releases.
- **Fridays**: Runs **every hour** in the early morning (until 12:00 AM MDT/MST Friday), and then **every 2 hours** for the remainder of the day.
- **Saturdays**: Runs **every 2 hours** to monitor any late portal updates.
- **Sundays – Wednesdays**: Runs **once daily** (12:00 PM UTC) to check for mid-week registrations.
- **Thursday Mornings**: Runs **once daily** (12:00 PM UTC / 6:00 AM MDT) to check for mid-week registrations before the afternoon hourly checks begin.
- **Manual Trigger**: You can run it manually at any time by going to the **Actions** tab in your GitHub repository, selecting the workflow, and clicking **Run workflow**. Checking "Force scraper run" will bypass the date check.

---

## Configuration

You can customize the project behavior by editing constants directly at the top of `scraper.py`:
- `NOTABLE_THRESHOLD`: The dollar threshold for highlighting individual donations in the alert email. (Default: `1000.0` or $1,000).
- `STATE_FILE` and `CSV_FILE`: Paths to save data files.

---

## Local Setup & Development

If you want to run or test the project locally:

1. Create a virtual environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the scraper:
   - **Dry Run** (downloads reports, parses them, and prints email details to terminal without writing database files or sending email):
     ```bash
     python scraper.py --dry-run
     ```
   - **Force Update** (forces files to update and alerts to send even if the portal date has not advanced):
     ```bash
     python scraper.py --force
     ```
   - **Standard Run** (only updates and alerts if new data is posted or a new group is found):
     ```bash
     python scraper.py
     ```
