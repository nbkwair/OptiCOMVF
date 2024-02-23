import pandas as pd
from tkinter import Tk, Button, Label, filedialog
from reportlab.lib import colors
from reportlab.lib.pagesizes import legal, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import subprocess
import time

# Function to compare the selected files
def compare_files():
    start_time = time.time()  # Record the start time of the comparison process

    control_file = control_file_label.cget("text")
    test_file = test_file_label.cget("text")

    # Check if both control and test files are selected
    if control_file and test_file:
        output_pdf = 'comparison_report.pdf'
        try:
            # Generate comparison report
            compare_excel_files(control_file, test_file, output_pdf)
            # Calculate total runtime
            end_time = time.time()
            total_runtime = end_time - start_time
            # Update information label with total runtime
            info_label.config(
                text=f"Comparison report generated successfully. Total runtime: {total_runtime:.2f} seconds")
            # Automatically open the report
            subprocess.Popen(['open', output_pdf])
        except Exception as e:
            info_label.config(text=f"Error: {str(e)}")

# Function to select the control file
def select_control_file():
    control_file_path = filedialog.askopenfilename()
    control_file_label.config(text=control_file_path)

# Function to select the test file
def select_test_file():
    test_file_path = filedialog.askopenfilename()
    test_file_label.config(text=test_file_path)

# Function to compare excel files
def compare_excel_files(control_file, test_file, output_pdf):
    try:
        # Read the Excel files into pandas DataFrames
        df_control = pd.read_excel(control_file, engine='openpyxl')
        df_test = pd.read_excel(test_file, engine='openpyxl')

        # Check if column names are different
        if not df_control.columns.equals(df_test.columns):
            raise ValueError("Column names are different in the two files.")

        # Find mismatched values
        comparison_results = df_control.ne(df_test)

        # Write comparison results to a PDF report
        doc = SimpleDocTemplate(output_pdf, pagesize=landscape(legal))
        styles = getSampleStyleSheet()
        elements = []

        # Create title for the report
        title = Paragraph("Comparison Report", styles['Title'])
        elements.append(title)
        elements.append(Paragraph("", styles['Title']))

        # Create table for the comparison results
        data = [df_control.columns.tolist()] + df_control.values.tolist()
        table = Table(data)

        # Apply style to the entire table
        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)])

        # Highlight mismatched cells
        for i, row in enumerate(df_control.index):
            for j, val in enumerate(df_control.columns):
                if comparison_results.iat[i, j]:
                    style.add('BACKGROUND', (j, i + 1), (j, i + 1), colors.lightcoral)

        table.setStyle(style)
        elements.append(table)

        # Add summary of mismatched cells
        mismatched_values = []
        for i, row in enumerate(comparison_results.index):
            for j, val in enumerate(comparison_results.columns):
                if comparison_results.iat[i, j]:
                    control_value = df_control.iat[i, j]
                    test_value = df_test.iat[i, j]
                    mismatched_values.append([row, val, control_value, test_value])

        if mismatched_values:
            elements.append(Paragraph("Mismatched Values:", styles['Heading2']))

            # Create table for mismatched values
            mismatch_table_data = [['Row', 'Column', 'Control Value', 'Test Value']] + mismatched_values
            mismatch_table = Table(mismatch_table_data)

            # Apply styles to the mismatch table
            mismatch_table_style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                               ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                               ('GRID', (0, 0), (-1, -1), 1, colors.black)])

            mismatch_table.setStyle(mismatch_table_style)
            elements.append(mismatch_table)

        # Build PDF
        doc.build(elements)
        print(f"Comparison report generated and saved to {output_pdf}.")
    except Exception as e:
        raise e

# Function to center the window
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x_coordinate = (screen_width / 2) - (width / 2)
    y_coordinate = (screen_height / 2) - (height / 2)

    window.geometry(f"{width}x{height}+{int(x_coordinate)}+{int(y_coordinate)}")

# Create Tkinter window
root = Tk()
root.title("Excel File Comparison")

# Set window dimensions
window_width = 600
window_height = 250

# Center the window
center_window(root, window_width, window_height)

# Button to select control file
control_button = Button(root, text="Select Control File", command=select_control_file, width=20)
control_button.pack(pady=10)

# Label to display control file path
control_file_label = Label(root, text="", width=50, anchor='w')
control_file_label.pack()

# Button to select test file
test_button = Button(root, text="Select Test File", command=select_test_file, width=20)
test_button.pack(pady=10)

# Label to display test file path
test_file_label = Label(root, text="", width=50, anchor='w')
test_file_label.pack()

# Button to compare files
compare_button = Button(root, text="Compare Files", command=compare_files, width=20)
compare_button.pack(pady=10)

# Label to display accuracy percentage
accuracy_label = Label(root, text="", width=50, anchor='w')
accuracy_label.pack()

# Label to display information/errors
info_label = Label(root, text="", width=50, anchor='w')
info_label.pack()

# Run the Tkinter event loop
root.mainloop()
