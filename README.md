# Tech Efficiency Analyzer v2.4

A Streamlit-based dashboard for analyzing technician performance, team efficiency, and workforce trends from Daily Tech Performance data.

## Features

- **Tech Leaderboard** - Individual technician rankings with efficiency scores
- **Diagnostic Summary** - Team quadrant analysis (Strong Team, Slow Techs, Turnover Risk, Needs Attention)
- **Capacity Analysis** - Understaffed/Overstaffed/Right-sized assessment per team
- **Seasonality Analysis** - Volume vs efficiency correlation over time
- **Weekly Staffing** - Headcount trends by team
- **Tenure Analysis** - Core tech % and workforce stability
- **Mileage Analysis** - Cost impact by team
- **Full Excel Export** - Multi-sheet report with charts

## Quick Start

### Option 1: Zero-Install Cloud Hosting (Easiest for Sharing)

If you want to share this tool with someone who isn't technical and doesn't want to download or install anything, the easiest way is to host it online for free using Streamlit Community Cloud:

1. Upload this entire code folder to a public or private GitHub repository.
2. Go to [share.streamlit.io](https://share.streamlit.io/) and log in.
3. Click **"New app"**, select your GitHub repository, and type `dashboard.py` in the "Main file path" box.
4. Click **Deploy!**

You'll instantly get a public URL (e.g., `https://your-tech-analyzer.streamlit.app/`) that anyone can visit to upload their data and get insights from any device.

### Option 2: Local Download (No Terminal Required)

You can zip this entire folder and send it to a friend or coworker. Assuming they have Python installed, they can just unzip the folder and double-click the setup script for their operating system:

- **Windows:** Double-click `run_windows.bat`
- **Mac/Linux:** Double-click `run_mac.command`

These scripts will automatically handle setting up the environment, installing required packages, and launching the dashboard in the browser.

### Option 3: Developer Setup (Terminal)

```bash
# Clone the repo
git clone https://github.com/brynr-builds/Tech-Analyzer.git
cd Tech-Analyzer

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the dashboard
streamlit run dashboard.py
```

## Usage

1. Open the dashboard at `http://localhost:8501`
2. Upload one or more Daily Tech Performance Excel files (.xlsx)
3. Use sidebar filters for date range and teams
4. View analysis across all tabs
5. Click "Generate Full Excel Report" to download

## Input Requirements

Your Excel files should contain columns like:
- **Technician** / Tech / Name
- **Technician Team** / Team / Region
- **Hours** / Hrs
- **Units** / Jobs / Builds
- **Amount** / Revenue
- **Date** (optional, enables trends)
- **Mileage** / Miles Paid (optional)

Column names are auto-normalized - the analyzer handles common variations.

## Configuration

Edit `config.json` to adjust scoring weights:

```json
{
  "w_units_per_hr": 0.4,
  "w_revenue_per_job": 0.3,
  "w_revenue_per_hr": 0.3,
  "tier_elite": 0.9,
  "tier_strong": 0.7,
  "tier_average": 0.3
}
```

## Files

| File | Purpose |
|------|---------|
| `dashboard.py` | Full visual dashboard with filtering |
| `streamlit_app.py` | Simpler upload-and-analyze interface |
| `analyzer_core.py` | Core analysis engine |
| `config.json` | Scoring weights and tier cutoffs |
| `requirements.txt` | Python dependencies |

## License

MIT License - free to use and modify.
