# מפעיל את הקוד המתוקן על הקובץ שהמשתמש העלה

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor

# טוען את הקובץ שהמשתמש העלה
pptx_path = 'Stock_Financial_Analysis_Fixed_Banner.pptx'
prs = Presentation(pptx_path)

# ערכים מעודכנים
new_ratios = {
    'Net Profit Margin': 0.23,
    'Operating Profit Margin': 0.21,
    'EBITDA Margin': 0.31,
    'ROE': 0.40,
    'Current Ratio': 1.6,
    'Quick Ratio': 1.3,
    'Immediate Liquidity Ratio': 1.2,
    'Cashflow to Sales': 0.25,
    'Leverage Ratio': 0.43,
    'Equity to Assets': 0.57,
    'Receivables Ratio': 0.16,
    'Inventory Turnover': 5.5,
    'DSO (Days Sales Outstanding)': 48,
    'Payables Days': 62
}

# פילוח
small_scale_ratios = {k: v for k, v in new_ratios.items() if v <= 2}
large_scale_ratios = {k: v for k, v in new_ratios.items() if v > 2}

# מחיקת כל השקפים שקשורים ל-"Financial Ratios"
slides_to_remove = []
for idx, slide in enumerate(prs.slides):
    if slide.shapes.title and "Financial Ratios" in slide.shapes.title.text:
        slides_to_remove.append(idx)

for idx in sorted(slides_to_remove, reverse=True):
    xml_slides = prs.slides._sldIdLst
    xml_slides.remove(xml_slides[idx])
class PresentationBuilder:
    def __init__(self,prs):
        self.prs = prs
        pass

   # פונקציה ליצירת שקף גרף חדש
    def create_column_chart_slide_simple(self,prs, data_dict, max_scale=None):
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)

        # רקע לבן
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # כותרת באנר
        title_shape = slide.shapes.title
        title_shape.text = "Financial Ratios"
        p = title_shape.text_frame.paragraphs[0]
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # גרף עמודות (לא אופקי)
        chart_data = CategoryChartData()
        chart_data.categories = list(data_dict.keys())
        chart_data.add_series('', list(data_dict.values()))

        x, y, cx, cy = Inches(0.8), Inches(2.2), Inches(8), Inches(4.5)
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart

        # עיצוב גרף
        chart.has_legend = False
        if max_scale:
            chart.value_axis.maximum_scale = max_scale

        chart.category_axis.tick_labels.font.size = Pt(12)
        chart.category_axis.tick_labels.font.color.rgb = RGBColor(10, 10, 10)
        chart.value_axis.tick_labels.font.size = Pt(12)
        chart.value_axis.tick_labels.font.color.rgb = RGBColor(10, 10, 10)

        if chart.value_axis.has_major_gridlines:
            gridlines = chart.value_axis.major_gridlines
            gridlines.format.line.color.rgb = RGBColor(200, 200, 200)
            gridlines.format.line.width = Pt(0.75)

        return slide
    def create_CAGR_slide(self):

        # חישוב שינוי באחוזים משנה קודמת (2020-2023 => 2021-2023)
        # חישוב CAGR כל שנה (בפועל: שינוי שנתי מדורג)
        def calc_cagr_from_start(values):
            base = values[0]
            return [0.0] + [
                round(((v / base) ** (1 / i) - 1) * 100, 2)
                for i, v in enumerate(values[1:], start=1)
            ]

        pct_change_data = {k: calc_cagr_from_start(v) for k, v in timeseries_data.items()}

        years = ["2020", "2021", "2022", "2023"]  # שינוי באחוזים מהשנה הקודמת

        # יצירת השקופית
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # כותרת בדיוק כמו בשקופיות הכחולות
        title_shape = slide.shapes.title
        title_shape.text = "Financial Change % Over Time"
        p = title_shape.text_frame.paragraphs[0]
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # יצירת נתוני גרף
        chart_data = CategoryChartData()
        chart_data.categories = years
        for name, values in pct_change_data.items():
            chart_data.add_series(name, values)

        # גרף מוסט למטה
        x, y, cx, cy = Inches(0.6), Inches(2.5), Inches(9), Inches(4.7)
        chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data).chart

        # עיצוב
        chart.has_legend = True
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
        chart.category_axis.tick_labels.font.size = Pt(12)
        chart.value_axis.tick_labels.font.size = Pt(12)
        chart.value_axis.has_major_gridlines = True
        chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(200, 200, 200)

    def create_sentiment_donut_chart(self, sentiment_data: dict):
        from pptx.enum.chart import XL_LEGEND_POSITION

        slide_layout = self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # כותרת השקופית
        title_shape = slide.shapes.title
        title_shape.text = "Sentiment Analysis Distribution"
        p = title_shape.text_frame.paragraphs[0]
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # גרף עוגה טבעתי — ממורכז ומתחת לכותרת
        chart_data = CategoryChartData()
        chart_data.categories = list(sentiment_data.keys())
        chart_data.add_series('Sentiment', list(sentiment_data.values()))

        chart_width = Inches(5.5)
        chart_height = Inches(4.0)
        slide_width = self.prs.slide_width

        x = Inches(3.5)  # מרכז בציר X
        y = Inches(2.5)  # מספיק נמוך כדי לא לגעת בכותרת

        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.DOUGHNUT, x, y, chart_width, chart_height, chart_data
        ).chart

        # ביטול כותרת הגרף הפנימית
        chart.has_title = False

        # עיצוב צבעים
        slice_colors = {
            "Positive": RGBColor(0, 176, 80),  # ירוק
            "Neutral": RGBColor(128, 128, 128),  # אפור
            "Negative": RGBColor(255, 0, 0)  # אדום
        }

        for i, category in enumerate(sentiment_data.keys()):
            point = chart.plots[0].series[0].points[i]
            if category in slice_colors:
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = slice_colors[category]

        # עיצוב אגדה
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(12)

    def create_swot_slide(self, title: str, section_data: dict):
        from pptx.enum.text import PP_ALIGN

        slide_layout = self.prs.slide_layouts[1]
        slide = self.prs.slides.add_slide(slide_layout)

        # כותרת השקופית
        slide.shapes.title.text = title

        # מיקום חדש — נמוך יותר
        left = Inches(1)
        top = Inches(2.5)  # 👈 תיבת הטקסט מוסטת למטה
        width = Inches(8)
        height = Inches(5)

        # יצירת תיבת טקסט חדשה
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.clear()

        for section_title, bullet_points in section_data.items():
            # כותרת הסעיף
            p = text_frame.add_paragraph()
            p.text = section_title
            p.font.bold = True
            p.font.size = Pt(20)
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(8)

            # פסקה אחת עם כל הבולטים
            bullet_text = "\n".join(f"• {bp}" for bp in bullet_points)
            para = text_frame.add_paragraph()
            para.text = bullet_text
            para.font.size = Pt(16)
            para.alignment = PP_ALIGN.LEFT
            para.line_spacing = Pt(24)  # 👈 שורה וחצי עבור טקסט בגודל 16pt

    def create_stock_forecast_slide(self, forecast_data: dict):
        from pptx.enum.chart import XL_LEGEND_POSITION

        slide_layout = self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # כותרת השקופית
        title_shape = slide.shapes.title
        title_shape.text = "Stock Price Forecast"
        p = title_shape.text_frame.paragraphs[0]
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # יצירת גרף עמודות
        chart_data = CategoryChartData()
        chart_data.categories = list(forecast_data.keys())
        chart_data.add_series("Forecast (%)", list(forecast_data.values()))

        x, y, cx, cy = Inches(1), Inches(2.3), Inches(8), Inches(4.5)
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart

        # עיצוב
        chart.has_legend = False
        chart.category_axis.tick_labels.font.size = Pt(12)
        chart.value_axis.tick_labels.font.size = Pt(12)

        # צבעים לעמודות: ירוק אם חיובי, אדום אם שלילי
        for i, val in enumerate(forecast_data.values()):
            point = chart.plots[0].series[0].points[i]
            point.format.fill.solid()
            if val >= 0:
                point.format.fill.fore_color.rgb = RGBColor(0, 176, 80)  # ירוק
            else:
                point.format.fill.fore_color.rgb = RGBColor(255, 0, 0)  # אדום

    def create_summary_slide(self, recommendation: str, reasons: list):
        from pptx.enum.text import PP_ALIGN

        slide_layout = self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

        # כותרת שקופית
        title_shape = slide.shapes.title
        title_shape.text = "Investment Summary & Recommendation"
        p = title_shape.text_frame.paragraphs[0]
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # אימוג’י וצבע
        if recommendation.lower() == "buy":
            rec_text = "Recommendation: BUY 📈"
            rec_color = RGBColor(0, 176, 80)
        elif recommendation.lower() == "sell":
            rec_text = "Recommendation: SELL 📉"
            rec_color = RGBColor(255, 0, 0)
        else:
            rec_text = f"Recommendation: {recommendation}"
            rec_color = RGBColor(0, 0, 0)

        # ✅ תיבת המלצה – מתחת לכותרת אבל מעל התוכן
        rec_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1))
        tf = rec_box.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.text = rec_text
        p.font.bold = True
        p.font.size = Pt(36)
        p.font.color.rgb = rec_color
        p.alignment = PP_ALIGN.LEFT

        # ✅ תיבת בולטים – מתחילה נמוך יותר כדי לא להתנגש
        reasons_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(3.5))
        tf2 = reasons_box.text_frame
        tf2.word_wrap = True
        tf2.clear()

        bullet_text = "\n".join(f"• {reason}" for reason in reasons)
        para = tf2.add_paragraph()
        para.text = bullet_text
        para.font.size = Pt(16)
        para.alignment = PP_ALIGN.LEFT
        para.line_spacing = Pt(30)


# יצירה בפועל
builder = PresentationBuilder(prs)
builder.create_column_chart_slide_simple(prs, small_scale_ratios, max_scale=2)
builder.create_column_chart_slide_simple(prs, large_scale_ratios)
# יצירת שקופית עם גרף קווי של שינוי באחוזים לאורך 4 שנים

# ערכים מקוריים לדוגמה:
timeseries_data = {
    'Revenue': [100, 110, 125, 140],
    'Net Income': [20, 25, 28, 32],
    'Operating Income': [18, 22, 26, 28],


    'Total Assets': [300, 320, 280, 400],
    'Total Liabilities': [160, 170, 180, 190],

    'Free Cash Flow': [16, 18, 22, 25]
}
builder.create_CAGR_slide()
sentiment = {
    "Positive": 45,
    "Neutral": 35,
    "Negative": 20
}
builder.create_sentiment_donut_chart(sentiment)
builder.create_swot_slide("Strengths and Opportunities", {
    "Strengths": [
        "Strong brand reputation and customer loyalty",
        "High R&D investment driving innovation",
        "Diversified product portfolio"
    ],
    "Opportunities": [
        "Expansion into emerging markets",
        "Growing demand for AI and automation",
        "Strategic partnerships and acquisitions"
    ]
})

builder.create_swot_slide("Weaknesses and Threats", {
    "Weaknesses": [
        "Dependence on a limited number of suppliers",
        "High production costs in core segments",
        "Limited presence in specific geographic areas"
    ],
    "Threats": [
        "Intensifying competition in global markets",
        "Regulatory pressures and compliance risks",
        "Economic downturn affecting consumer spending"
    ]
})
forecast = {
    "3 Days": 0.8,
    "1 Month": 2.4,
    "3 Months": -1.1,
    "1 Year": 5.7
}
builder.create_stock_forecast_slide(forecast)
builder.create_summary_slide("Buy", [
    "Strong financial ratios and profitability margins",
    "Positive sentiment dominates the market tone",
    "Consistent CAGR growth across major KPIs",
    "1-Year forecast shows strong potential upside",
    "SWOT analysis indicates strengths outweigh risks"
])

# שמירה
prs.save("Stock_Financial_Analysis_Final_With_PctChangeChart.pptx")
