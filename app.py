"""
app.py — Calorie Burn Predictor
Implementasi Regresi Linear Sederhana menggunakan Flask

Dataset : Calories Burnt Prediction Dataset (Kaggle)
          exercise.csv + calories.csv → digabung via User_ID → disimpan sebagai data/calories.csv
Variabel: X = Duration (menit), Y = Calories (kcal)
Model   : sklearn LinearRegression (Simple Linear Regression)
"""

import os
import pickle
import math

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')   # backend non-interaktif, wajib untuk server tanpa GUI
import matplotlib.pyplot as plt

from flask import Flask, render_template, request, jsonify
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error

# ─── Inisialisasi Flask ───────────────────────────────────────────────────────
app = Flask(__name__)

# ─── Path Konstanta ───────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_PATH  = os.path.join(BASE_DIR, 'data',   'calories.csv')
MODEL_PATH = os.path.join(BASE_DIR, 'model',  'linear_regression_model.pkl')
PLOT_PATH  = os.path.join(BASE_DIR, 'static', 'images', 'regression_plot.png')

# ─── Variabel Global ──────────────────────────────────────────────────────────
model       = None   # objek LinearRegression yang sudah dilatih
model_stats = {}     # dict berisi semua statistik & metrik model


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — PREPROCESSING DATA
# ═════════════════════════════════════════════════════════════════════════════

def load_and_prepare_data():
    """
    Membaca dataset, memilih kolom Duration & Calories,
    menangani missing value, dan memastikan tipe numerik.

    Returns:
        X         (np.ndarray): array 2D kolom Duration, shape (n, 1)
        y         (np.ndarray): array 1D kolom Calories, shape (n,)
        n_total   (int)       : jumlah baris setelah pembersihan
        dur_min   (float)     : nilai minimum Duration dalam dataset
        dur_max   (float)     : nilai maksimum Duration dalam dataset
        cal_min   (float)     : nilai minimum Calories dalam dataset
        cal_max   (float)     : nilai maksimum Calories dalam dataset
    """
    df = pd.read_csv(DATA_PATH)

    # Pilih hanya kolom yang diperlukan
    df = df[['Duration', 'Calories']].copy()

    # Pastikan tipe numerik
    df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')
    df['Calories'] = pd.to_numeric(df['Calories'], errors='coerce')

    # Hapus baris yang mengandung missing value
    before = len(df)
    df.dropna(subset=['Duration', 'Calories'], inplace=True)
    after = len(df)

    if before != after:
        print(f"[INFO] Dihapus {before - after} baris missing value.")

    # Hapus baris dengan nilai tidak logis
    df = df[(df['Duration'] > 0) & (df['Calories'] > 0)]

    dur_min = float(df['Duration'].min())
    dur_max = float(df['Duration'].max())
    cal_min = float(df['Calories'].min())
    cal_max = float(df['Calories'].max())

    X = df[['Duration']].values   # 2D array untuk scikit-learn
    y = df['Calories'].values

    return X, y, len(df), dur_min, dur_max, cal_min, cal_max


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — TRAINING MODEL
# ═════════════════════════════════════════════════════════════════════════════

def train_model(X, y):
    """
    Melakukan train-test split (80:20, random_state=42),
    melatih LinearRegression, dan menghitung 4 metrik evaluasi
    pada DATA TESTING.

    Returns:
        reg_model                   : model LinearRegression yang sudah dilatih
        metrics (dict)              : R², MAE, MSE, RMSE
        X_train, X_test, y_train, y_test
    """
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    reg_model = LinearRegression()
    reg_model.fit(X_train, y_train)

    # Prediksi hanya pada data TESTING
    y_pred = reg_model.predict(X_test)

    r2   = float(r2_score(y_test, y_pred))
    mae  = float(mean_absolute_error(y_test, y_pred))
    mse  = float(mean_squared_error(y_test, y_pred))
    rmse = float(np.sqrt(mse))

    metrics = {
        'r2'  : round(r2,   4),
        'mae' : round(mae,  4),
        'mse' : round(mse,  4),
        'rmse': round(rmse, 4),
    }

    return reg_model, metrics, X_train, X_test, y_train, y_test


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — SIMPAN MODEL
# ═════════════════════════════════════════════════════════════════════════════

def save_model(reg_model):
    """Menyimpan model ke file .pkl menggunakan pickle."""
    os.makedirs(os.path.dirname(MODEL_PATH), exist_ok=True)
    with open(MODEL_PATH, 'wb') as f:
        pickle.dump(reg_model, f)


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — GENERATE GRAFIK REGRESI
# ═════════════════════════════════════════════════════════════════════════════

def generate_regression_plot(reg_model, X, y, dur_min, dur_max):
    """
    Membuat scatter plot data aktual + garis regresi dari model,
    lalu menyimpan ke static/images/regression_plot.png.
    """
    os.makedirs(os.path.dirname(PLOT_PATH), exist_ok=True)

    fig, ax = plt.subplots(figsize=(10, 6))

    # Scatter plot — data aktual
    ax.scatter(
        X, y,
        alpha=0.35,
        color='#3b82f6',
        edgecolors='white',
        linewidths=0.4,
        s=35,
        label='Data Aktual',
        zorder=2
    )

    # Garis regresi — dari model yang sudah dilatih
    x_line = np.linspace(dur_min, dur_max, 300).reshape(-1, 1)
    y_line = reg_model.predict(x_line)
    ax.plot(
        x_line, y_line,
        color='#ef4444',
        linewidth=2.5,
        label='Garis Regresi Linear',
        zorder=3
    )

    # Styling
    ax.set_xlabel('Duration (Minutes)', fontsize=13, labelpad=8)
    ax.set_ylabel('Calories Burned (kcal)', fontsize=13, labelpad=8)
    ax.set_title(
        'Hubungan Durasi Olahraga dengan Kalori Terbakar\n(Regresi Linear Sederhana)',
        fontsize=14, fontweight='bold', pad=16
    )
    ax.legend(fontsize=11, loc='upper left')
    ax.grid(True, linestyle='--', alpha=0.4, zorder=1)
    ax.set_facecolor('#f8fafc')
    fig.patch.set_facecolor('#ffffff')

    plt.tight_layout()
    plt.savefig(PLOT_PATH, dpi=150, bbox_inches='tight')
    plt.close(fig)


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — INISIALISASI UTAMA
# ═════════════════════════════════════════════════════════════════════════════

def initialize_model():
    """
    Menjalankan seluruh pipeline:
      1. Load & preprocessing dataset
      2. Train-test split + training model
      3. Hitung R², MAE, MSE, RMSE (dari data test)
      4. Simpan model ke .pkl
      5. Generate grafik regresi
      6. Simpan semua statistik ke variabel global model_stats
    """
    global model, model_stats

    print("[INFO] Memuat dataset...")
    X, y, n_total, dur_min, dur_max, cal_min, cal_max = load_and_prepare_data()
    print(f"[INFO] Dataset dimuat: {n_total} baris data bersih.")

    print("[INFO] Melatih model Linear Regression...")
    reg_model, metrics, X_train, X_test, y_train, y_test = train_model(X, y)

    slope     = round(float(reg_model.coef_[0]), 4)
    intercept = round(float(reg_model.intercept_), 4)

    print(f"[INFO] Model dilatih:")
    print(f"       R²={metrics['r2']}  MAE={metrics['mae']}  MSE={metrics['mse']}  RMSE={metrics['rmse']}")
    print(f"       Persamaan: Calories = {intercept} + {slope} × Duration")

    save_model(reg_model)
    print(f"[INFO] Model disimpan ke: {MODEL_PATH}")

    generate_regression_plot(reg_model, X, y, dur_min, dur_max)
    print(f"[INFO] Grafik disimpan ke: {PLOT_PATH}")

    model = reg_model
    model_stats = {
        # Ukuran data
        'n_total' : n_total,
        'n_train' : len(X_train),
        'n_test'  : len(X_test),
        # Statistik deskriptif dataset
        'dur_min' : dur_min,
        'dur_max' : dur_max,
        'cal_min' : cal_min,
        'cal_max' : cal_max,
        # Metrik evaluasi (dari data test)
        'r2'      : metrics['r2'],
        'mae'     : metrics['mae'],
        'mse'     : metrics['mse'],
        'rmse'    : metrics['rmse'],
        # Koefisien model
        'slope'    : slope,
        'intercept': intercept,
    }


# ═════════════════════════════════════════════════════════════════════════════
# HELPER — FORMAT PERSAMAAN REGRESI
# ═════════════════════════════════════════════════════════════════════════════

def format_equation(intercept, slope, decimals=4):
    """
    Menghasilkan string persamaan regresi yang benar untuk ditampilkan.

    Contoh: "Calories = -21.6668 + 7.1534 × Duration"
            "Calories = 5.2000 + 4.8000 × Duration"
    """
    fmt = f".{decimals}f"
    return f"Calories = {intercept:{fmt}} + {slope:{fmt}} × Duration"


# ═════════════════════════════════════════════════════════════════════════════
# ROUTE — HALAMAN UTAMA
# ═════════════════════════════════════════════════════════════════════════════

@app.route('/')
def index():
    """Halaman utama dengan form prediksi dan statistik ringkas."""
    return render_template('index.html', stats=model_stats)


# ═════════════════════════════════════════════════════════════════════════════
# ROUTE — ENDPOINT PREDIKSI
# ═════════════════════════════════════════════════════════════════════════════

@app.route('/predict', methods=['POST'])
def predict():
    """
    Menerima JSON: { "duration": <angka> }
    Mengembalikan JSON: { "success": bool, "prediction": float, "equation": str, ... }

    Validasi:
      - Input wajib ada dan berupa angka
      - Nilai harus > 0
      - Nilai harus dalam rentang [dur_min, dur_max] dari dataset
        (mencegah ekstrapolasi tanpa peringatan)
    """
    data = request.get_json(silent=True)

    # Validasi keberadaan field
    if not data or 'duration' not in data:
        return jsonify({'success': False, 'error': 'Input durasi tidak ditemukan.'}), 400

    # Validasi tipe angka
    try:
        duration = float(data['duration'])
    except (ValueError, TypeError):
        return jsonify({'success': False, 'error': 'Durasi harus berupa angka.'}), 400

    dur_min = model_stats['dur_min']
    dur_max = model_stats['dur_max']

    # Validasi > 0
    if duration <= 0:
        return jsonify({
            'success': False,
            'error': 'Durasi olahraga harus lebih dari 0 menit.'
        }), 400

    # Validasi rentang dataset — cegah ekstrapolasi
    if duration < dur_min or duration > dur_max:
        return jsonify({
            'success': False,
            'error': (
                f"Durasi berada di luar rentang data yang digunakan untuk melatih model. "
                f"Masukkan durasi antara {int(dur_min)}–{int(dur_max)} menit."
            )
        }), 400

    # Prediksi
    prediction = model.predict(np.array([[duration]]))[0]
    prediction = max(0.0, round(float(prediction), 2))

    equation = format_equation(model_stats['intercept'], model_stats['slope'])

    return jsonify({
        'success'   : True,
        'duration'  : duration,
        'prediction': prediction,
        'equation'  : equation,
        'dur_min'   : dur_min,
        'dur_max'   : dur_max,
    })


# ═════════════════════════════════════════════════════════════════════════════
# ROUTE — HALAMAN ANALISIS MODEL
# ═════════════════════════════════════════════════════════════════════════════

@app.route('/analysis')
def analysis():
    """Halaman analisis lengkap: dataset info, metrik, grafik, interpretasi."""
    equation = format_equation(model_stats['intercept'], model_stats['slope'])
    return render_template('analysis.html', stats=model_stats, equation=equation)


# ═════════════════════════════════════════════════════════════════════════════
# ROUTE — HALAMAN ABOUT
# ═════════════════════════════════════════════════════════════════════════════

@app.route('/about')
def about():
    """Halaman tentang aplikasi."""
    return render_template('about.html', stats=model_stats)


# ═════════════════════════════════════════════════════════════════════════════
# INISIALISASI OTOMATIS SAAT APLIKASI DIMUAT
# ═════════════════════════════════════════════════════════════════════════════

initialize_model()

# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
