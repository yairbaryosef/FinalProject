
# -*- coding: utf-8 -*-

import torch
import torch.nn as nn
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import yfinance as yf
import json
import os
from sklearn.preprocessing import MinMaxScaler
from sklearn.metrics import mean_squared_error, r2_score
from torch.utils.data import DataLoader, TensorDataset
from datetime import datetime
import time
from .Calc_Ratios import *


device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
print(f"Using device: {device}")


class PctChangeLSTM(nn.Module):
    def __init__(self, input_size=1, hidden_size=64):
        super(PctChangeLSTM, self).__init__()
        self.lstm = nn.LSTM(input_size, hidden_size, batch_first=True)
        self.fc_pct = nn.Linear(hidden_size, 1)

    def forward(self, x):
        _, (h_n, _) = self.lstm(x)
        pct_change = self.fc_pct(h_n[-1])
        return pct_change.squeeze()


def load_data_with_ratios(ticker='AMZN'):
    print("Downloading close price data and financial ratios...")
    stock = yf.Ticker(ticker)

    # Load historical close prices
    end_date = datetime.today()
    start_date = end_date - pd.DateOffset(years=5)
    df_prices = yf.download(ticker, start=start_date, end=end_date)
    df_prices = df_prices[['Close']].copy()
    df_prices.index = pd.to_datetime(df_prices.index)
    df_prices.sort_index(inplace=True)
    df_prices.index.name = 'Date'

    # Get last two quarters
    if len(stock.quarterly_financials.columns) < 2:
        print("Not enough quarterly financial data.")
        return df_prices, None

    current_period = stock.quarterly_financials.columns[0]
    previous_period = stock.quarterly_financials.columns[1]

    ratios = calc_ratios_for_quarter(stock, current_period, previous_period)
    ratios_df = pd.DataFrame([ratios], index=[current_period]) if ratios else None

    print(f"Downloaded {len(df_prices)} price records for {ticker} and financial ratios.")
    return df_prices, ratios_df


def preprocess_data_with_ratios(data, ratios, input_len=90, target_len=90):
    print("Preprocessing data with ratios...")
    data_log = np.log1p(data)
    scaler = MinMaxScaler()
    scaled_prices = scaler.fit_transform(data_log.values.reshape(-1, 1))

    # Convert ratios to a feature vector (same length as input sequence)
    ratio_features = np.array(list(ratios.iloc[0])) if ratios is not None else np.zeros(16)
    ratio_features = ratio_features.astype(np.float32)
    ratio_features = np.nan_to_num(ratio_features)  # Handle NaNs

    X, y_price, y_pct = [], [], []
    for i in range(len(scaled_prices) - input_len - target_len):
        price_seq = scaled_prices[i:i + input_len]
        price_seq = np.hstack([price_seq, np.tile(ratio_features, (input_len, 1))])
        target_seq = scaled_prices[i + input_len:i + input_len + target_len].flatten()

        start_price = data.values[i + input_len - 1]
        end_price = data.values[i + input_len + target_len - 1]
        pct_change = (end_price - start_price) / start_price

        X.append(price_seq)
        y_price.append(target_seq)
        y_pct.append(pct_change)

    X = torch.tensor(np.array(X, dtype=np.float32), dtype=torch.float32)
    y_price = torch.tensor(np.array(y_price, dtype=np.float32), dtype=torch.float32)
    y_pct = torch.tensor(np.array(y_pct, dtype=np.float32), dtype=torch.float32)

    loader = DataLoader(TensorDataset(X, y_price, y_pct), batch_size=32, shuffle=True)

    print(f"Generated {len(X)} sequences with {X.shape[2]} features. Input shape: {X.shape}")
    return X, y_price, y_pct, loader, scaler

def inverse_log_transform(arr):
    return np.expm1(arr)


def plot_prediction(pred, pct_change, title, save_path=None, recent_prices=None):
    if recent_prices is None:
        raise ValueError("`recent_prices` must be provided.")
    future_dates = pd.date_range(start=recent_prices.index[-1] + pd.Timedelta(days=1), periods=len(pred), freq='B')

    full_dates = recent_prices.index.append(future_dates)
    full_prices = np.concatenate([recent_prices['Close'].values.reshape(-1), pred])

    plt.figure(figsize=(12, 6))
    plt.plot(full_dates, full_prices, label='Close Price + Forecast', color='purple')
    plt.axvline(recent_prices.index[-1], linestyle='--', color='gray', label='Forecast Start')
    plt.title(f"{title}\nPredicted 90-day Change: {pct_change*100:.2f}%")
    plt.xlabel('Date')
    plt.ylabel('Close Price')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    if save_path:
        plt.savefig(save_path)
    else:
        plt.show()
    plt.close()


def train_pct_model(model, loader, criterion, optimizer, epochs):
    model.train()
    X_all, _, y_pct_all = loader.dataset.tensors
    split_idx = int(0.8 * len(X_all))
    X_train, X_val = X_all[:split_idx], X_all[split_idx:]
    y_train, y_val = y_pct_all[:split_idx], y_pct_all[split_idx:]

    train_loader = DataLoader(TensorDataset(X_train, y_train), batch_size=32, shuffle=True)
    for epoch in range(epochs):
        total_loss = 0
        for X_batch, y_batch in train_loader:
            X_batch, y_batch = X_batch.to(device), y_batch.to(device)
            optimizer.zero_grad()
            out = model(X_batch)
            loss = criterion(out, y_batch)
            if torch.isnan(loss):
                continue
            loss.backward()
            optimizer.step()
            total_loss += loss.item()
        print(f"[PctChange Epoch {epoch+1}/{epochs}] Loss: {total_loss/len(train_loader):.6f}")

    model.eval()
    with torch.no_grad():
        val_preds = model(X_val.to(device)).cpu().numpy()
        y_true = y_val.cpu().numpy()
        mse = mean_squared_error(y_true, val_preds)
        r2 = r2_score(y_true, val_preds)
        direction_accuracy = np.mean((val_preds > 0) == (y_true > 0))

    return model, val_preds, mse, r2, direction_accuracy


def train_price_model(model, sequences, targets, criterion, optimizer, epochs):
    model.train()
    dataset = TensorDataset(sequences, targets)
    loader = DataLoader(dataset, batch_size=32, shuffle=True)

    for epoch in range(epochs):
        total_loss = 0
        for X_batch, y_batch in loader:
            X_batch, y_batch = X_batch.to(device), y_batch.to(device)
            optimizer.zero_grad()
            out = model(X_batch)
            loss = criterion(out, y_batch)
            if torch.isnan(loss):
                continue
            loss.backward()
            optimizer.step()
            total_loss += loss.item()
        print(f"[Price Epoch {epoch+1}/{epochs}] Loss: {total_loss/len(loader):.6f}")


def main(ticker):
    start_time = time.time()

    input_len, target_len, epochs = 90, 90, 50

    df_prices, ratios_df = load_data_with_ratios(ticker)
    df = df_prices.copy()

    X, y_price, y_pct, loader, scaler = preprocess_data_with_ratios(df['Close'], ratios_df, input_len, target_len)

    pct_model = PctChangeLSTM(input_size=X.shape[2], hidden_size=64).to(device)
    optimizer_pct = torch.optim.Adam(pct_model.parameters(), lr=1e-3)
    criterion_pct = nn.MSELoss()

    pct_model, pct_val_preds, mse, r2, direction_acc = train_pct_model(
        pct_model, loader, criterion_pct, optimizer_pct, epochs
    )

    last_seq = X[-1:].to(device)
    with torch.no_grad():
        predicted_pct = pct_model(last_seq).item()

    print(f"Predicted 90-day pct_change: {predicted_pct:.4f} ({predicted_pct * 100:.2f}%)")


    return f"Predicted 90-day pct_change: {predicted_pct:.4f} ({predicted_pct * 100:.2f}%)"


