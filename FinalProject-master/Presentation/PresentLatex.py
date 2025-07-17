import json
import os
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import yfinance as yf
from pptx.dml.color import RGBColor
from Presentation.UpdatePre.design.run import Present


from Anlysis.Finnancial_Ratios.scraperyahoo import calc_Ratios_with_growth
from SSH.connect_to_slurm import run_it_all
from SSH.Actions.Predict_Stock import submit_lstm_job


import os
def main(stock, symbol, pdf):
    run_it_all(str(pdf), base_dir=os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
    hostname = 'slurm.bgu.ac.il'
    port = 22
    username = 'yairbary'
    password = 'yairYAIR0_00'

    #submit_lstm_job(hostname=hostname, port=port, username=username, password=password,symbol=symbol)
    # Set paths
    try:
        ratios = calc_Ratios_with_growth(stock, symbol)
    except:
        ratios = {}


    from pathlib import Path

    # Path to the current script
    project_root = Path(__file__).resolve().parent

    # Full path to the file
    base_dir = project_root

    print("Resolved path:", base_dir)

    sentiment_file = os.path.join(base_dir, "Files", "sentiment_results.json")
    print(sentiment_file)
    swot_output = os.path.join(base_dir, "Files", "swot_output.txt")


    llm_lines_path = os.path.join(base_dir,  "Files", "llm_lines.txt")  # ✅ חדש
    print()

    output_pptx = "Presentation/Stock_Financial_Analysis_Final_With_PctChangeChart.pptx"


    # תקין עם 3 ארגומנטים
    Present(symbol,sentiment_file, swot_output, llm_lines_path,output_pptx)

    return output_pptx




import os

def process(symbol, local_file_path):
    # ננרמל את הנתיב שיהיה חוקי בכל מערכת (Windows/Unix)
    local_file_path = os.path.normpath(local_file_path.strip().replace('"', '').replace("'", ''))

    # הסרה כפולה של פרוטוקולים כפולים (כמו C:/C:/ או C:/C:\\)
    drive, tail = os.path.splitdrive(local_file_path)
    if '\\\\' in tail or 'C:\\' in tail or 'C:/' in tail:
        tail = os.path.basename(tail)
        local_file_path = os.path.join(drive, tail)

    print("📂 Clean local_file_path:", local_file_path)

    if not os.path.isfile(local_file_path):
        raise FileNotFoundError(f"❌ File does not exist: {local_file_path}")

    stock = yf.Ticker(symbol)
    return main(stock, symbol, local_file_path)

if __name__ == '__main__':
    process('ESLT.TA',"")
