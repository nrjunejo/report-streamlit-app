import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import requests
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Function to parse and print tables from the document
def parse_and_print_tables(doc):
    for idx, table in enumerate(doc.tables, start=1):
        st.write(f"Table {idx}:")
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            st.write(row_data)

# Function to format cell value based on the column index and table index
def format_cell_value(value, col_index, table_index=None):
    if pd.isna(value) or value == '':
        return ''
    if table_index == 2 and col_index == 3:
        return str(value)
    if col_index == 0:
        return pd.to_datetime(value).strftime('%b-%y') if not isinstance(value, str) else value
    elif col_index == 2:
        try:
            return f"{int(value)}"
        except ValueError:
            return value
    elif col_index == 3:
        return f"${float(value):,.2f}" if isinstance(value, (int, float)) else value
    elif col_index == 4 or col_index == 5:
        return f"${float(value):,.2f}" if isinstance(value, (int, float)) else value
    elif col_index == 6:
        return f"{int(value)}%" if isinstance(value, (int, float)) else value
    else:
        return str(value)

# Function to update table with new values
def update_table_with_values(table, new_values, start_row=1):
    num_rows_to_add = len(new_values) - (len(table.rows) - start_row)
    if num_rows_to_add > 0:
        for _ in range(num_rows_to_add):
            table.add_row()
    for i, row in enumerate(new_values):
        for j, cell_value in enumerate(row):
            if (i + start_row) < len(table.rows) and j < len(table.columns):
                cell = table.cell(i + start_row, j)
                formatted_value = format_cell_value(cell_value, j)
                cell.text = formatted_value
                cell.width = cell.width
                cell.paragraphs[0].alignment = cell.paragraphs[0].alignment
            else:
                st.write(f"Skipping value at row {i + start_row}, column {j} - out of bounds")

# Function to update specific column of the table with new values
def update_table_column_with_values(table, column_index, new_values, start_row=1, table_index=None):
    num_rows_to_add = len(new_values) - (len(table.rows) - start_row)
    if num_rows_to_add > 0:
        for _ in range(num_rows_to_add):
            table.add_row()
    for i, value in enumerate(new_values):
        row_index = i + start_row
        if row_index < len(table.rows):
            cell = table.cell(row_index, column_index)
            formatted_value = format_cell_value(value, column_index, table_index)
            cell.text = formatted_value
            cell.width = cell.width
            cell.paragraphs[0].alignment = cell.paragraphs[0].alignment
        else:
            st.write(f"Skipping value at row {row_index}, column {column_index} - out of bounds")

    if len(new_values) > 1 and len(table.rows) == start_row + 1:
        for i, value in enumerate(new_values[1:], start=start_row + 1):
            cell = table.cell(i, column_index)
            formatted_value = format_cell_value(value, column_index, table_index)
            cell.text = formatted_value
            cell.width = cell.width
            cell.paragraphs[0].alignment = cell.paragraphs[0].alignment

# Function to set cell border
def set_cell_border(cell, **kwargs):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    for edge in ('top', 'left', 'bottom', 'right'):
        if edge in kwargs:
            edge_data = kwargs[edge]
            tag = f'w:{edge}'
            element = tcPr.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcPr.append(element)
            for key in edge_data:
                element.set(qn(key), str(edge_data[key]))

# Main function to create the Streamlit app
def main():
    st.set_page_config(page_title="Monthly Report Updater", page_icon=":bar_chart:", layout="wide")

    st.markdown(
        """
        <style>
        .main {
            background-color: #121212;
        }
        .report-uploader {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1);
        }
        .header {
            text-align: center;
            font-size: 2em;
            font-weight: bold;
            margin-bottom: 30px;
        }
        .subheader {
            font-size: 1.2em;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .instructions {
            background-color: #4CAF50;
            font-size: 1em;
            margin-bottom: 20px;
        }
        .button {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 10px 20px;
            text-align: center;
            font-size: 1em;
            border-radius: 5px;
            cursor: pointer;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<div class='header'>Monthly Report Generator</div>", unsafe_allow_html=True)

    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])

        with col2:
            st.markdown("<div class='report-uploader'>", unsafe_allow_html=True)
            st.markdown("<div class='subheader'>Upload Files</div>", unsafe_allow_html=True)
            st.markdown("<div class='instructions'>Please upload the required files in the specified order.</div>", unsafe_allow_html=True)

            timesheet_file = st.file_uploader("1. Upload Sample Timesheet Report.xlsx", type="xlsx", key="timesheet")
            funding_file = st.file_uploader("2. Upload Sample Funding Status.xlsx", type="xlsx", key="funding")

            # Download the report file from the GitHub repository
            report_file_url = "https://github.com/nrjunejo/report-streamlit-app/raw/main/SAMPLE%20REPORT.docx"
            report_response = requests.get(report_file_url)
            report_bytes = BytesIO(report_response.content)
            report_file = report_bytes if report_response.status_code == 200 else None

            month = st.text_input("Enter the month:", key="month")
            year = st.text_input("Enter the year:", key="year")
            if st.button("Update Report"):
                if timesheet_file and funding_file and report_response.status_code == 200 and month and year:
                    excel_sheet1 = pd.read_excel(timesheet_file)
                    excel_sheet2 = pd.read_excel(funding_file)
                    doc = Document(report_file)

                    new_month = month
                    new_year = year
                    new_date = "1st"

                    for paragraph in doc.paragraphs:
                        if "For the month of" in paragraph.text and "Precision commenced services on" in paragraph.text:
                            updated_sentence = f"For the month of {new_month} {new_year} – Precision commenced services on {new_month} {new_date} {new_year}. Primary PAS attendant continued government security clearance process."
                            paragraph.text = updated_sentence

                    table2_column4_values = excel_sheet1.iloc[26:30, 1].tolist()
                    table = doc.tables[1]
                    update_table_column_with_values(table, 3, table2_column4_values, table_index=2)

                    # Add your table updating code here

                    doc.save("Updated_SAMPLE_REPORT.docx")
                    st.success("Report updated successfully!")

                    with open("Updated_SAMPLE_REPORT.docx", "rb") as f:
                        st.download_button("Download Updated Report", f, file_name="Updated_SAMPLE_REPORT.docx")
                else:
                    st.error("Please upload all required files and enter the month and year.")

            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        """
        <style>
        footer {visibility: hidden;}
        .css-1aumxhk {padding: 20px;}
        .css-12oz5g7 {padding-top: 3.5rem;}
        </style>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
