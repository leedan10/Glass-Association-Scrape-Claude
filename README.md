# Glass-Association-Scrape-Claude

Scrapes the NGA (National Glass Association) member directory from
[members.glass.org](https://members.glass.org/cvweb/cgi-bin/utilities.dll/OpenPage?wrp=ngaSearch.htm)
and exports the data to CSV and XLSX.

## Data collected

| Field          | Description                                      |
|----------------|--------------------------------------------------|
| Name           | Company / organization name                      |
| Classification | Business type (e.g. Retailer, Contract Glazier)  |
| Address        | Street address                                   |
| City           | City                                             |
| State          | State                                            |
| Phone          | Phone number                                     |
| Web Address    | Company website URL                              |

## Run via GitHub Actions (recommended)

1. Go to the **Actions** tab in this repository.
2. Select the **Scrape NGA Members** workflow.
3. Click **Run workflow**.
4. When the run finishes, download the artifact (`nga-members-<run_id>`) which
   contains `nga_members.csv` and `nga_members.xlsx`.

The workflow also runs automatically on the **1st of every month**.

## Run locally

```bash
pip install -r requirements.txt
playwright install chromium --with-deps
python scraper.py
```

Output files are written to the `output/` directory.
