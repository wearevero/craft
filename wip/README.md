## Routine/WIP

A lightweight scraper that automates extracting table data from local or remote HTML pages, and dumps it into nicely formatted `.xlsx` files. No more copy-pasting from a browser tab all afternoon.

### Source

All main script files live in the `src` directory — `filling.py`, `komponen.py`, and `loss.py`. Each one scrapes a different report, but they all follow the same flow: log in, hit a bunch of generated URLs, scrape the table, save to Excel. Feel free to poke around or tweak them for your own report.

### Driver

You'll need a ChromeDriver build that matches your installed Chrome version, or Selenium will refuse to talk to it. Grab a compatible version here:

- [Chrome for Testing on GitHub](https://github.com/GoogleChromeLabs/chrome-for-testing)
- [Chrome for Testing website](https://googlechromelabs.github.io/chrome-for-testing/)

Once downloaded, point `CHROMEDRIVER_PATH` in your `.env` at the `.exe`.

### Data

Generated files aren't hardcoded to a folder anymore — where they land is controlled entirely by the `OUTPUT_DIR` variable in your `.env`. If it's not set, the scripts fall back to a sane default, but you'll almost certainly want to set your own path so files end up somewhere you actually check.

### Usage

1. Clone the repository.
2. Copy the example env file and fill it in:

```bash
   cp .env.example .env
```

3. Install dependencies:

```bash
   pip install -r requirements.txt
```

4. Open `.env` and configure the following:
   | Variable | What it's for |
   |---------------------|--------------------------------------------------------------|
   | `WEB_URL` | Base URL of the site you're scraping |
   | `EMAIL` | Login email/username |
   | `PASSWORD` | Login password |
   | `CHROMEDRIVER_PATH` | Path to your ChromeDriver executable |
   | `OUTPUT_DIR` | Where the generated `.xlsx` files should be saved |
   On Windows, wrap paths in quotes if they contain spaces or backslashes, e.g.:

```dotenv
   CHROMEDRIVER_PATH="driver\151.0.7922.76\chromedriver.exe"
   OUTPUT_DIR="D:\Work\CORE\WIP\KRIDO\WIP TEMPLATE"
```

5. Run whichever scraper you need:

```bash
   python src/filling.py
   python src/loss.py
   python src/komponen.py
```

Each script will prompt for a date (defaults to today if you just hit enter), then leave the browser open at the end so you can eyeball the results before closing it yourself.
