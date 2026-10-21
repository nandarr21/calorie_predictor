"""
app.py — Calorie Burn Predictor
Implementasi Regresi Linear Sederhana menggunakan Flask

Dataset : Calories Burnt Prediction Dataset (Kaggle)
          exercise.csv + calories.csv → digabung via User_ID → data/calories.csv
Variabel: X = Duration (menit), Y = Calories (kcal)
Model   : sklearn LinearRegression (Simple Linear Regression)
"""

import os
import sys

# Konfigurasi cache Matplotlib ke /tmp untuk lingkungan serverless (Vercel/AWS Lambda)
os.environ.setdefault('MPLCONFIGDIR', '/tmp')

import pickle
import math
import numpy as np
import pandas as pd

import matplotlib
matplotlib.use('Agg')   # backend non-interaktif untuk server
import matplotlib.pyplot as plt

from flask import Flask, render_template, request, jsonify
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error

# ─── Path Konstanta ───────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_PATH  = os.path.join(BASE_DIR, 'data',   'calories.csv')
MODEL_PATH = os.path.join(BASE_DIR, 'model',  'linear_regression_model.pkl')
PLOT_PATH  = os.path.join(BASE_DIR, 'static', 'images', 'regression_plot.png')

# ─── Inisialisasi Flask dengan Path Eksplisit ─────────────────────────────────
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)

# ─── Variabel Global ──────────────────────────────────────────────────────────
model       = None
model_stats = {
    'n_total'  : 15000,
    'n_train'  : 12000,
    'n_test'   : 3000,
    'dur_min'  : 1.0,
    'dur_max'  : 30.0,
    'cal_min'  : 1.0,
    'cal_max'  : 314.0,
    'r2'       : 0.9145,
    'mae'      : 13.6135,
    'mse'      : 345.0410,
    'rmse'     : 18.5753,
    'slope'    : 7.1534,
    'intercept': -21.6668,
}


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — PREPROCESSING DATA
# ═════════════════════════════════════════════════════════════════════════════

def load_and_prepare_data():
    df = pd.read_csv(DATA_PATH)
    df = df[['Duration', 'Calories']].copy()
    df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')
    df['Calories'] = pd.to_numeric(df['Calories'], errors='coerce')
    df.dropna(subset=['Duration', 'Calories'], inplace=True)
    df = df[(df['Duration'] > 0) & (df['Calories'] > 0)]

    dur_min = float(df['Duration'].min())
    dur_max = float(df['Duration'].max())
    cal_min = float(df['Calories'].min())
    cal_max = float(df['Calories'].max())

    X = df[['Duration']].values
    y = df['Calories'].values

    return X, y, len(df), dur_min, dur_max, cal_min, cal_max


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — TRAINING MODEL
# ═════════════════════════════════════════════════════════════════════════════

def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    reg_model = LinearRegression()
    reg_model.fit(X_train, y_train)

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
# FUNGSI — SIMPAN MODEL (AMAN SERVERLESS / READ-ONLY)
# ═════════════════════════════════════════════════════════════════════════════

def save_model(reg_model):
    try:
        os.makedirs(os.path.dirname(MODEL_PATH), exist_ok=True)
        with open(MODEL_PATH, 'wb') as f:
            pickle.dump(reg_model, f)
    except (OSError, PermissionError) as e:
        print(f"[INFO] Filesystem read-only (Serverless): Lewati penyimpanan pickle.")


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — GENERATE GRAFIK REGRESI (AMAN SERVERLESS)
# ═════════════════════════════════════════════════════════════════════════════

def generate_regression_plot(reg_model, X, y, dur_min, dur_max):
    if os.path.exists(PLOT_PATH):
        return

    try:
        os.makedirs(os.path.dirname(PLOT_PATH), exist_ok=True)

        fig, ax = plt.subplots(figsize=(10, 6))

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

        x_line = np.linspace(dur_min, dur_max, 300).reshape(-1, 1)
        y_line = reg_model.predict(x_line)
        ax.plot(
            x_line, y_line,
            color='#ef4444',
            linewidth=2.5,
            label='Garis Regresi Linear',
            zorder=3
        )

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
    except (OSError, PermissionError) as e:
        print(f"[INFO] Filesystem read-only (Serverless): Lewati penyimpanan grafik.")


# ═════════════════════════════════════════════════════════════════════════════
# FUNGSI — INISIALISASI UTAMA
# ═════════════════════════════════════════════════════════════════════════════

def initialize_model():
    global model, model_stats

    # 1. Coba muat model tersimpan (.pkl) terlebih dahulu
    if os.path.exists(MODEL_PATH):
        try:
            with open(MODEL_PATH, 'rb') as f:
                model = pickle.load(f)
            print(f"[INFO] Model dimuat dari: {MODEL_PATH}")
            return
        except Exception as e:
            print(f"[WARN] Gagal membaca pickle model: {e}")

    # 2. Jika pickle belum ada dan dataset tersedia, lakukan training
    if os.path.exists(DATA_PATH):
        try:
            print("[INFO] Memuat dataset...")
            X, y, n_total, dur_min, dur_max, cal_min, cal_max = load_and_prepare_data()

            print("[INFO] Melatih model Linear Regression...")
            reg_model, metrics, X_train, X_test, y_train, y_test = train_model(X, y)

            slope     = round(float(reg_model.coef_[0]), 4)
            intercept = round(float(reg_model.intercept_), 4)

            model = reg_model
            model_stats.update({
                'n_total'  : n_total,
                'n_train'  : len(X_train),
                'n_test'   : len(X_test),
                'dur_min'  : dur_min,
                'dur_max'  : dur_max,
                'cal_min'  : cal_min,
                'cal_max'  : cal_max,
                'r2'       : metrics['r2'],
                'mae'      : metrics['mae'],
                'mse'      : metrics['mse'],
                'rmse'     : metrics['rmse'],
                'slope'    : slope,
                'intercept': intercept,
            })

            save_model(reg_model)
            generate_regression_plot(reg_model, X, y, dur_min, dur_max)
            return
        except Exception as e:
            print(f"[ERROR] Gagal saat melatih model: {e}")

    # 3. Fallback instan jika model maupun data tidak dapat diakses
    if model is None:
        fallback_model = LinearRegression()
        fallback_model.coef_ = np.array([model_stats['slope']])
        fallback_model.intercept_ = model_stats['intercept']
        model = fallback_model
        print("[INFO] Menggunakan model fallback LinearRegression terkonfigurasi.")


# ═════════════════════════════════════════════════════════════════════════════
# HELPER — FORMAT PERSAMAAN REGRESI
# ═════════════════════════════════════════════════════════════════════════════

def format_equation(intercept, slope, decimals=4):
    fmt = f".{decimals}f"
    return f"Calories = {intercept:{fmt}} + {slope:{fmt}} × Duration"


# ═════════════════════════════════════════════════════════════════════════════
# ROUTE — ENDPOINT (Mendukung path langsung maupun rewrite Vercel)
# ═════════════════════════════════════════════════════════════════════════════

@app.route('/')
@app.route('/api')
@app.route('/api/index', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        return predict()
    return render_template('index.html', stats=model_stats)


@app.route('/predict', methods=['POST'])
@app.route('/api/predict', methods=['POST'])
def predict():
    data = request.get_json(silent=True)

    if not data or 'duration' not in data:
        return jsonify({'success': False, 'error': 'Input durasi tidak ditemukan.'}), 400

    try:
        duration = float(data['duration'])
    except (ValueError, TypeError):
        return jsonify({'success': False, 'error': 'Durasi harus berupa angka.'}), 400

    dur_min = model_stats['dur_min']
    dur_max = model_stats['dur_max']

    if duration <= 0:
        return jsonify({
            'success': False,
            'error': 'Durasi olahraga harus lebih dari 0 menit.'
        }), 400

    if duration < dur_min or duration > dur_max:
        return jsonify({
            'success': False,
            'error': (
                f"Durasi berada di luar rentang data yang digunakan untuk melatih model. "
                f"Masukkan durasi antara {int(dur_min)}–{int(dur_max)} menit."
            )
        }), 400

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


@app.route('/analysis')
@app.route('/api/analysis')
def analysis():
    equation = format_equation(model_stats['intercept'], model_stats['slope'])
    return render_template('analysis.html', stats=model_stats, equation=equation)


@app.route('/about')
@app.route('/api/about')
def about():
    return render_template('about.html', stats=model_stats)


@app.errorhandler(404)
def not_found(e):
    # Fallback jika URL dialihkan ke route root
    return render_template('index.html', stats=model_stats)


# ═════════════════════════════════════════════════════════════════════════════
# INISIALISASI AMAN
# ═════════════════════════════════════════════════════════════════════════════

try:
    initialize_model()
except Exception as e:
    print(f"[CRITICAL] Error initialize_model: {e}")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
