import os
import win32com.client

def convert_to_pdf(filepath):
    """
    Convert an Excel file to PDF
    """
    filename = os.path.basename(filepath)
    filename_without_ext, _ = os.path.splitext(filename)
    pdf_filename = os.path.join(os.path.dirname(filepath), f"{filename_without_ext}.pdf")


    excel_app = None
    workbook = None

    try:

        # Create an instance of Excel Application
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Interactive = False
        excel_app.Visible = False

        # Open the Excel file
        workbook = excel_app.Workbooks.Open(filepath)

        # Select all sheets
        for worksheet in workbook.Worksheets:
            worksheet.Select()

        # Export to PDF
        workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_filename)

    except Exception as e:
        print(f"Failed to convert {filename} to PDF format.")
        print(str(e))
    finally:
        # Close the Excel file
        if workbook:
            workbook.Close()

        # Quit Excel
        if excel_app:
            excel_app.Quit()


if __name__ == '__main__':
    directory = os.getcwd()
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(directory, filename)
            print(filepath)
            convert_to_pdf(filepath)
