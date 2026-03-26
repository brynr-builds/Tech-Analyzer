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

### Option 1: Run Locally

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

### Option 2: One-Line Install (Mac/Linux)

```bash
git clone https://github.com/brynr-builds/Tech-Analyzer.git && cd Tech-Analyzer && python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt && streamlit run dashboard.py
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
