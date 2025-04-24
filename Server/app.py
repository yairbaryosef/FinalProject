from flask import Flask, render_template, request, send_file
import os
import yfinance as yf
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from Presentation.PresentLatex import process
from SSH.Actions.Predict_Stock import submit_lstm_job

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    symbol = request.form.get("symbol")
    pdf_link = request.form.get("pdf_file")
    output_path = process(symbol, pdf_link)
    return send_file(output_path, as_attachment=True)

import os
import base64
from io import BytesIO
from flask import render_template, request
import matplotlib.pyplot as plt
from PIL import Image

@app.route("/predict", methods=["GET", "POST"])
def predict():
    if request.method == "POST":
        stock_symbol = request.form.get("stock_symbol").upper()  # e.g., "AMZN"
        model_type = request.form.get("model_type")
        features = request.form.getlist("features")
        local_stock_file = "stock.txt"
        with open(local_stock_file, "w") as file:
            file.write(stock_symbol)
        print("stock.txt created locally.")
        hostname = 'slurm.bgu.ac.il'
        port = 22
        username = 'yairbary'
        password = 'yairYAIR0_0'

        submit_lstm_job(hostname=hostname, port=port, username=username, password=password,model=model_type)
        # Set paths
        output_dir = os.path.join(os.getcwd(), "output")
        result_path = os.path.join(output_dir, f"results.txt")
        image_path = os.path.join(output_dir, f"simulated_forecast.png")

        # Load result text
        try:
            with open(result_path, "r") as file:
                prediction_text = file.read()
        except FileNotFoundError:
            prediction_text = f"Result file not found for {stock_symbol}"

        # Load and encode image
        try:
            with open(image_path, "rb") as img_file:
                encoded_img = base64.b64encode(img_file.read()).decode('utf-8')
                plot_url = f"data:image/png;base64,{encoded_img}"
        except FileNotFoundError:
            plot_url = None
        try:
            with open('C:/Users/yairb/PycharmProjects/Projectduhot/Server/output/results.txt', 'r') as file:
                results_text = file.read()
        except Exception as e:
            results_text = f"⚠️ Error reading results.txt: {e}"
        return render_template(
            "predict.html",
            prediction=prediction_text,
            results_text=results_text,
            stock_symbol=stock_symbol,
            model_type=model_type,
            features=features
        )


    return render_template("predict.html")

@app.route("/about")
def about():
    return render_template("about.html")

if __name__ == "__main__":
    app.run(debug=True)
