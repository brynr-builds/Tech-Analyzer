import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from analyzer_core import AnalyzerConfig, run_analysis

# Create dummy data
dates = [datetime(2023, 1, 1) + timedelta(days=i) for i in range(30)]
data = {
    "Date": np.random.choice(dates, 100),
    "Technician": np.random.choice(["Tech A", "Tech B", "Tech C"], 100),
    "Technician Team": np.random.choice(["Team 1", "Team 2"], 100),
    "Role": ["Technician"] * 100,
    "Hours": np.random.uniform(2, 8, 100),
    "Units": np.random.randint(1, 5, 100),
    "Amount": np.random.uniform(100, 500, 100),
    "Mileage": np.random.uniform(0, 50, 100),
}
df = pd.DataFrame(data)
df.to_excel("dummy_data.xlsx", index=False)

# Run analysis
cfg = AnalyzerConfig()
results = run_analysis(["dummy_data.xlsx"], cfg)

# Verify expected columns exist
assert "Consistency Score" in results["individual"].columns
assert "Flight Risk" in results["individual"].columns
assert "Coaching Profile" in results["individual"].columns
assert "Target Techs" in results["capacity_analysis"].columns
assert "Techs to Hire/Transfer" in results["capacity_analysis"].columns

print("Tests passed successfully!")
