import json
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog

import pandas as pd
import PyPDF2
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class PDFFetcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Fetcher")

        self.setup_ui()

    def setup_ui(self):
        # Create and place UI elements here

        # Label for instructions
        instructions_label = tk.Label(self.root, text="Fill in the details and click 'Fetch' to start fetching PDFs.")
        instructions_label.pack(pady=(10, 0))

        # Browse button to select xlsx file
        self.browse_button = tk.Button(self.root, text="Browse XLSX File", command=self.browse_xlsx)
        self.browse_button.pack(pady=(10, 0))

        # Entry field for xlsx file
        self.xlsx_file_entry = tk.Entry(self.root)
        self.xlsx_file_entry.pack(pady=(5, 0))

        # Entry field for URL
        self.url_entry = tk.Entry(self.root)
        self.url_entry.pack(pady=(5, 0))
        self.url_entry.insert(0, "https://onlineresults.unipune.ac.in/Result/Dashboard/ViewResult1")

        # Entry fields for PatternID and PatternName
        self.pattern_id_label = tk.Label(self.root, text="PatternID:")
        self.pattern_id_label.pack(pady=(5, 0))
        self.pattern_id_entry = tk.Entry(self.root)
        self.pattern_id_entry.pack(pady=(5, 0))
        self.pattern_id_entry.insert(0, "GxTZTSYcOVy18dCZIascgA==")

        self.pattern_name_label = tk.Label(self.root, text="PatternName:")
        self.pattern_name_label.pack(pady=(5, 0))
        self.pattern_name_entry = tk.Entry(self.root)
        self.pattern_name_entry.pack(pady=(5, 0))
        self.pattern_name_entry.insert(0, "5Zb5Cz8e8AKy7NyhnK8K9/usTIbNzt8PEhfGjr5QpMI=")

        # Fetch button
        self.fetch_button = tk.Button(self.root, text="Fetch", command=self.start_fetching)
        self.fetch_button.pack(pady=(10, 0))

        # Button to extract PDFs into Excel
        self.extract_button = tk.Button(self.root, text="Extract PDFs into Excel", command=self.extract_pdfs_to_excel)
        self.extract_button.pack(pady=(10, 0))

        # Text widget for logs
        logs_label = tk.Label(self.root, text="Logs:")
        logs_label.pack(pady=(10, 0))
        
        self.logs_text = tk.Text(self.root, height=10, width=50)
        self.logs_text.pack(pady=(10, 0))
    def extract_text_from_pdf(self, pdf_path):
        text = ""
        with open(pdf_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for page in pdf_reader.pages:
                text += page.extract_text()
        return text
            
    def extract_pdfs_to_excel(self):
        # List all PDF files in the current directory
        pdf_files = [file for file in os.listdir() if file.endswith('.pdf')]

        # Create lists to store data
        pdf_names = []
        total_credits_earned_list = []
        sgpa_list = []

        # Loop through each PDF file
        for pdf_filename in pdf_files:
            pdf_path = os.path.join(os.getcwd(), pdf_filename)
            pdf_text = self.extract_text_from_pdf(pdf_path)

            # Find and store "TOTAL CREDITS EARNED" value
            search_string_total_credits = "TOTAL CREDITS EARNED :"
            start_index_total_credits = pdf_text.find(search_string_total_credits)
            if start_index_total_credits != -1:
                start_index_total_credits += len(search_string_total_credits)
                end_index_total_credits = pdf_text.find("\n", start_index_total_credits)
                if end_index_total_credits != -1:
                    total_credits_earned = pdf_text[start_index_total_credits:end_index_total_credits].strip()
                    total_credits_earned = re.search(r'\d+', total_credits_earned).group()
                else:
                    total_credits_earned = "Value not found"
            else:
                total_credits_earned = "Value not found"

            # Find and store "SGPA" value
            search_string_sgpa = "SGPA :-"
            start_index_sgpa = pdf_text.find(search_string_sgpa)
            if start_index_sgpa != -1:
                start_index_sgpa += len(search_string_sgpa)
                end_index_sgpa = pdf_text.find("\n", start_index_sgpa)
                if end_index_sgpa != -1:
                    sgpa = pdf_text[start_index_sgpa:end_index_sgpa].strip()
                else:
                    sgpa = "Value not found"
            else:
                sgpa = "Value not found"

            # Store data in lists
            pdf_names.append(pdf_filename)
            total_credits_earned_list.append(total_credits_earned)
            sgpa_list.append(sgpa)

        # Create a DataFrame from the lists
        df = pd.DataFrame({
            "PDF Name": pdf_names,
            "Total Credits Earned": total_credits_earned_list,
            "SGPA": sgpa_list
        })

        # Save the DataFrame to an Excel file
        xlsx_file = 'pdf_data.xlsx'
        df.to_excel(xlsx_file, index=False)

        self.update_logs(f"Data saved to {xlsx_file}")

    def browse_xlsx(self):
        xlsx_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.xlsx_file_entry.delete(0, tk.END)
        self.xlsx_file_entry.insert(0, xlsx_file)

    def start_fetching(self):
        xlsx_file = self.xlsx_file_entry.get()
        url = self.url_entry.get()
        pattern_id = self.pattern_id_entry.get()
        pattern_name = self.pattern_name_entry.get()

        if not xlsx_file or not url or not pattern_id or not pattern_name:
            self.update_logs("Please fill in all fields.")
            return

        data_list = self.load_data_from_xlsx(xlsx_file)

        if not data_list:
            self.update_logs("No data found in the selected XLSX file.")
            return

        self.logs_text.delete("1.0", tk.END)
        fetch_thread = threading.Thread(target=self.fetch_pdfs, args=(data_list, url, pattern_id, pattern_name))
        fetch_thread.start()

    def load_data_from_xlsx(self, xlsx_file):
        try:
            df = pd.read_excel(xlsx_file)
            return df.to_dict(orient='records')
        except Exception as e:
            self.update_logs(f"Error loading data from XLSX file: {str(e)}")
            return []

    def fetch_pdfs(self, data_list, url, pattern_id, pattern_name):
        # Disable SSL certificate verification (not recommended for production use)
        verify_ssl = False

        retry_count = 4  # Number of times to retry fetching PDF
        retry_list = []  # List to store students with fetch errors
        
        for data in data_list:
            for _ in range(retry_count):
                payload = {
                    "PatternID": pattern_id,
                    "PatternName": pattern_name,
                    "SeatNo": data["SeatNo"],
                    "MotherName": data["MotherName"],
                }

                response = requests.post(url, data=payload, verify=verify_ssl)

                if response.status_code == 200 and response.headers["Content-Type"] == "application/pdf":
                    pdf_filename = f"{data['RollNo']} {data['StudentName']}.pdf"
                    with open(pdf_filename, "wb") as f:
                        f.write(response.content)
                    self.update_logs(f"PDF saved for SeatNo: {data['RollNo']} {data['StudentName']}")
                    break  # Successfully fetched PDF, exit retry loop
                else:
                    self.update_logs(f"Error fetching PDF for SeatNo: {data['RollNo']} {data['StudentName']}")
                    if _ == retry_count - 1:
                        retry_list.append(data['SeatNo'])

        if retry_list:
            self.update_logs("\nPDF fetching retries failed for the following students:")
            self.update_logs(", ".join(retry_list))

        self.update_logs("\nPDF fetching completed.")

    def update_logs(self, log_message):
        self.logs_text.insert(tk.END, log_message + "\n")
        self.logs_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFFetcherApp(root)
    root.mainloop()

# C:\Users\suyog\Desktop\RESULT>python gui.py
# 2023-08-08T22:08:23.154ZE [2156:NonCelloThread] crash.cc:84:HandleCrashpadLog [10884:2156:20230809,033823.154:ERROR crash_report_database_win.cc:614] CreateDirectory C:\Users\91940\AppData\Local\Google\DriveFS\Crashpad: The system cannot find the path specified. (3)

# Could not initialize crash reporting DB
# Can not init crashpad with status: UNKNOWN: Could not initialize crash reporting DB [type.googleapis.com/drive.ds.Status='CRASHPAD_DB_INIT_ERROR']
# C:\Users\suyog\Desktop\RESULT>python gui.py
# 2023-08-08T22:13:21.129ZE [16832:NonCelloThread] crash.cc:84:HandleCrashpadLog [6792:16832:20230809,034321.128:ERROR crash_report_database_win.cc:614] CreateDirectory C:\Users\91940\AppData\Local\Google\DriveFS\Crashpad: The system cannot find the path specified. (3)

# Could not initialize crash reporting DB
# Can not init crashpad with status: UNKNOWN: Could not initialize crash reporting DB [type.googleapis.com/drive.ds.Status='CRASHPAD_DB_INIT_ERROR']