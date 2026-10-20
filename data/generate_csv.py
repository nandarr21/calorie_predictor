"""
Script pembantu untuk generate dataset calories.csv
Dijalankan sekali untuk membuat file data/calories.csv
"""
import pandas as pd
import numpy as np

np.random.seed(42)
n = 500

duration = np.random.uniform(10, 120, n)
# Hubungan linear: Calories = intercept + slope * Duration + noise
intercept = 5.0
slope = 5.2
noise = np.random.normal(0, 15, n)
calories = intercept + slope * duration + noise
calories = np.clip(calories, 10, None)  # tidak boleh negatif

df = pd.DataFrame({
    'Duration': np.round(duration, 1),
    'Calories': np.round(calories, 2)
})

df.to_csv('calories.csv', index=False)
print(f"Dataset berhasil dibuat: {len(df)} baris")
print(df.head())
