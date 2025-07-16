import traceback

from Presentation.UpdatePre.design.run import extract_lines_dict_from_textfile
from SSH.connect_to_slurm import run_it_all
import os
import yfinance as yf

# תיקיית הקובץ הנוכחי
base_dir = os.path.dirname(os.path.abspath(__file__))

# נתיב לתיקיית ה-PDFים
pdf_folder = os.path.join(base_dir, "pdf_files")
local_output = os.path.join(base_dir, "output_auto")
base_dir_for_llm = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
llm_file = os.path.join(base_dir_for_llm, "Files", 'llm_lines.txt')
# מציאת קבצי PDF
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]

for pdf_file in pdf_files:
    symbol = str(pdf_file).split('_')[0]
    stock = yf.Ticker(symbol)
    raw_cap = 'fill manually later'#stock.info.get('marketCap')
    run_it_all(str(pdf_file), base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..")))

    try:
        lines = extract_lines_dict_from_textfile(llm_file, symbol=symbol, market_cap=raw_cap)
        output_path = os.path.join(local_output, f"{symbol}.txt")


        with open(output_path, "w", encoding="utf-8") as f:
            for key, value in lines.items():
                if isinstance(value, list):
                    f.write(f"{key}:\n")
                    for item in value:
                        f.write(f"  {item}\n")
                else:
                    f.write(f"{key}: {value}\n")

        print(f"✅ Saved extracted lines to: {output_path}")
    except Exception as e:
        error_str = traceback.format_exc()
        print(error_str)
        print(f"❌ Failed to extract lines from LLM output for {symbol}: {e}")

