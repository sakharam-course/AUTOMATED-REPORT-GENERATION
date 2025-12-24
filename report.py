import pandas as pd
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors


# READ EXCEL FILE
def read_data(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()  # Fix column spacing issues
    return df


# CREATE QUANTITY BAR CHART (NO SUMMARY ON IMAGE)
def create_quantity_bar_chart(df):

    plt.figure(figsize=(8, 5))
    plt.bar(df["Items"], df["Quantity"])
    plt.xlabel("Items")
    plt.ylabel("Quantity")
    plt.title("Quantity-wise Bar Chart")
    plt.xticks(rotation=30)

    chart_path = "quantity_bar_chart.png"
    plt.tight_layout()
    plt.savefig(chart_path)
    plt.close()

    return chart_path


# CALCULATE SALES SUMMARY
def calculate_summary(df):

    df["Date"] = pd.to_datetime(df["Date"])
    df["Sales"] = df["Quantity"] * df["Price"]

    monthly_sales = df.groupby(df["Date"].dt.month_name())["Sales"].sum()

    total_sales = df["Sales"].sum()
    avg_monthly_sales = monthly_sales.mean()
    highest_month = monthly_sales.idxmax()
    highest_value = monthly_sales.max()

    return total_sales, avg_monthly_sales, highest_month, highest_value, df


# GENERATE PDF REPORT
def generate_pdf(df, chart_path, summary, output_file):

    total_sales, avg_sales, highest_month, highest_value, df = summary

    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(output_file, pagesize=A4)
    story = []

    # Title
    story.append(Paragraph("<b>Excel Data Visualization Report</b>", styles["Title"]))
    story.append(Spacer(1, 20))

    # SUMMARY SECTION (ABOVE RAW DATA)
    story.append(Paragraph("<b>Summary of Sales Data</b>", styles["Heading2"]))
    story.append(Spacer(1, 8))

    summary_text = (
        f"Total Sales: {int(total_sales)}<br/>"
        f"Average Monthly Sales: {avg_sales:.2f}<br/>"
        f"Highest Sales Month: {highest_month} ({int(highest_value)})"
    )

    story.append(Paragraph(summary_text, styles["Normal"]))
    story.append(Spacer(1, 20))

    # RAW DATA TABLE
    story.append(Paragraph("<b>Raw Data</b>", styles["Heading2"]))
    story.append(Spacer(1, 10))

    raw_table_data = [df.columns.tolist()] + df.astype(str).values.tolist()
    table = Table(raw_table_data, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER')
    ]))

    story.append(table)
    story.append(Spacer(1, 20))

    # BAR CHART
    story.append(Paragraph("<b>Quantity-wise Bar Chart</b>", styles["Heading2"]))
    story.append(Spacer(1, 10))
    story.append(Image(chart_path, width=400, height=250))

    doc.build(story)
    print(f"PDF report generated: {output_file}")


# MAIN
if __name__ == "__main__":
    file_path = r"C:\Users\sakharam\OneDrive\Desktop\Task-2\data - Copy.xlsx"
    output_pdf = "Excel_Report.pdf"

    df = read_data(file_path)
    chart_path = create_quantity_bar_chart(df)
    summary = calculate_summary(df)
    generate_pdf(df, chart_path, summary, output_pdf)
