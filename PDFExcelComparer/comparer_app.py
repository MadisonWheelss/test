import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from PIL import Image 
import pytesseract
import re
import os
import sys 
import fitz 
import io
from collections import defaultdict 

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # PyInstaller bundle path 
    tesseract_bundle_dir = os.path.join(sys._MEIPASS, "Tesseract-OCR")
    pytesseract.pytesseract.tesseract_cmd = os.path.join(tesseract_bundle_dir, "tesseract.exe")
    os.environ['TESSDATA_PREFIX'] = os.path.join(tesseract_bundle_dir, "tessdata")
else:
    # development path
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


class PDFExcelComparerApp:
    def __init__(self, master):
        # Initialize the main application window
        self.master = master
        master.title("PDF & Excel Comparer")
        master.geometry("800x600") 
        master.resizable(True, True) 
        master.configure(bg="#F0F4F8") 

        # Configure root grid to make the main_frame expand
        master.rowconfigure(0, weight=1)
        master.columnconfigure(0, weight=1)

        # Create a main frame to contain all other widgets
        self.main_frame = tk.Frame(master, bg="#F0F4F8", padx=10, pady=10)
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # Configure grid for responsive layout within the main_frame
        self.main_frame.grid_rowconfigure(0, weight=0) 
        self.main_frame.grid_rowconfigure(1, weight=0) 
        self.main_frame.grid_rowconfigure(2, weight=0) 
        self.main_frame.grid_rowconfigure(3, weight=1) 
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1) 

        # Variables to store file paths
        self.pdf_file_path = tk.StringVar()
        self.excel_file_path = tk.StringVar()

        # --- User Interface Elements ---

        # Frame for file selection inputs
        input_frame = tk.Frame(self.main_frame, bg="#F0F4F8", padx=15, pady=15)
        input_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0, pady=0) 
        input_frame.grid_columnconfigure(0, weight=1)
        input_frame.grid_columnconfigure(1, weight=4) 
        input_frame.grid_columnconfigure(2, weight=1) 

        # PDF File Selection
        tk.Label(input_frame, text="PDF File:", font=("Inter", 12, "bold"), bg="#F0F4F8", fg="#334155").grid(row=0, column=0, sticky="w", pady=5)
        self.pdf_entry = tk.Entry(input_frame, textvariable=self.pdf_file_path, width=70, font=("Inter", 10), bd=1, highlightbackground="#CBD5E1", highlightthickness=1, borderwidth=1, relief="flat", highlightcolor="#6366F1", insertbackground="#6366F1", fg="#334155")
        self.pdf_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.pdf_entry.configure(justify="left", borderwidth=1, highlightbackground="#D1D5DB", highlightcolor="#4F46E5", highlightthickness=1, insertbackground="#4F46E5")

        pdf_button = tk.Button(input_frame, text="Browse PDF", command=self.browse_pdf_file, font=("Inter", 10, "bold"), bg="#6366F1", fg="white", activebackground="#4338CA", activeforeground="white", relief="raised", bd=0, padx=10, pady=5)
        pdf_button.grid(row=0, column=2, sticky="ew", padx=5, pady=5)
        self.apply_button_style(pdf_button)

        # Excel File Selection
        tk.Label(input_frame, text="Excel File:", font=("Inter", 12, "bold"), bg="#F0F4F8", fg="#334155").grid(row=1, column=0, sticky="w", pady=5)
        self.excel_entry = tk.Entry(input_frame, textvariable=self.excel_file_path, width=70, font=("Inter", 10), bd=1, highlightbackground="#CBD5E1", highlightthickness=1, borderwidth=1, relief="flat", highlightcolor="#6366F1", insertbackground="#6366F1", fg="#334155")
        self.excel_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.excel_entry.configure(justify="left", borderwidth=1, highlightbackground="#D1D5DB", highlightcolor="#4F46E5", highlightthickness=1, insertbackground="#4F46E5")

        excel_button = tk.Button(input_frame, text="Browse Excel", command=self.browse_excel_file, font=("Inter", 10, "bold"), bg="#6366F1", fg="white", activebackground="#4338CA", activeforeground="white", relief="raised", bd=0, padx=10, pady=5)
        excel_button.grid(row=1, column=2, sticky="ew", padx=5, pady=5)
        self.apply_button_style(excel_button)

        # Run Comparison Button
        run_button = tk.Button(self.main_frame, text="Run Comparison", command=self.run_comparison, font=("Inter", 14, "bold"), bg="#22C55E", fg="white", activebackground="#16A34A", activeforeground="white", relief="raised", bd=0, padx=20, pady=10)
        run_button.grid(row=1, column=0, columnspan=2, pady=15) 
        self.apply_button_style(run_button)

        # Results Display Area
        tk.Label(self.main_frame, text="Comparison Results:", font=("Inter", 12, "bold"), bg="#F0F4F8", fg="#334155").grid(row=2, column=0, columnspan=2, sticky="nw", padx=0, pady=5) 
        self.results_text = scrolledtext.ScrolledText(self.main_frame, wrap=tk.WORD, font=("Inter", 10), bg="white", fg="#334155", bd=1, relief="solid", highlightbackground="#CBD5E1", highlightthickness=1, borderwidth=1, padx=10, pady=10)
        self.results_text.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=0, pady=10) 


    def apply_button_style(self, button):
        button.config(
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            cursor="hand2",
            padx=15,
            pady=8,
            activebackground=button["bg"] 
        )
        button.bind("<Enter>", lambda e: e.widget.config(relief="raised", borderwidth=1, highlightbackground="#9CA3AF"))
        button.bind("<Leave>", lambda e: e.widget.config(relief="flat", borderwidth=0, highlightbackground=button["bg"]))

    def browse_pdf_file(self):
        """Opens a file dialog for PDF selection."""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")],
            title="Select PDF File"
        )
        if file_path:
            self.pdf_file_path.set(file_path)

    def browse_excel_file(self):
        """Opens a file dialog for Excel selection."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Select Excel File"
        )
        if file_path:
            self.excel_file_path.set(file_path)

    # This method takes the input PDF document and turns it into a text file 
    def _get_pdf_full_text(self, input_pdf_path):
        
        full_text = ""
        try:
            self.results_text.insert(tk.END, "Opening PDF for OCR...\n")
            self.master.update_idletasks()
            doc = fitz.open(input_pdf_path)
            self.results_text.insert(tk.END, "PDF opened. Starting OCR...\n")
            self.master.update_idletasks()
            tesseract_config = '--oem 1 --psm 6' # '--oem 1 --psm 6' most accurate

            for page_number in range(len(doc)):
                self.results_text.insert(tk.END, f"Processing page {page_number + 1}...\n")
                self.master.update_idletasks()
                page = doc.load_page(page_number) 
                pix = page.get_pixmap(dpi=200)  # 200 (best range)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                text = pytesseract.image_to_string(img, config=tesseract_config) # Run OCR with the pre-processed image

                # Apply your common OCR fixes: 
                text = text.replace('@', '0').replace('e', '0').replace('Q', '0').replace('O', '0')
                text = text.replace('I', '1').replace('l', '1').replace('B', '8').replace('S', '5')
                text = text.replace('*', '').replace('$', '').replace('\f', '').replace(':', '').replace('%','')

                full_text += f"\n--- Page {page_number + 1} ---\n{text}\n"
            
            #DEBUG TEXT
            #print(full_text)
            self.results_text.insert(tk.END, "OCR complete. Text extracted from PDF.\n")
            self.master.update_idletasks()
            return full_text
        except pytesseract.TesseractNotFoundError:
            raise RuntimeError("Tesseract OCR engine not found. Ensure it's correctly bundled with the application.")
        except Exception as e:
            self.results_text.insert(tk.END, f"Error during PDF OCR: {e}\n")
            raise
    
    # This method takes the input pdf.txt file and finds all account numbers with their associated payments 
    def _get_info_from_pdf_text(self, pdf_full_text):
        account_totals = defaultdict(float)
        account_counts = defaultdict(int)
        current_account = None

        # Regex patterns 
        account_pattern = re.compile(r'\bW\d{6,8}\b')
    
        # Specific payment patterns
        elf_pay_pattern = re.compile(r'ELF PAY AU\s*(-?\s*\d+\s*\.\s*\d{2})')
        delete_bat_pattern = re.compile(r'XXDELETE FR[O0]M [B8]AT\s*(-?\s*\d+\s*\.\s*\d{2})')
        credit_card_pattern = re.compile(r'CARD PAYME\s*(-?\s*\d+\s*\.\s*\d{2})')

        

        self.results_text.insert(tk.END, "Extracting account and payment info from PDF text...\n")
        self.master.update_idletasks()

        lines = pdf_full_text.splitlines()

        # (NEED TO FIX) payments from the same account can be on different pages 
        for line in lines:
            # --- DEBUGGING START ---
            #self.results_text.insert(tk.END, f"\nProcessing line: '{line.strip()}'\n")
            #self.results_text.insert(tk.END, f"Current Account before processing line: {current_account}\n")
            # --- DEBUGGING END ---
            # Look for an account number
            account_match = account_pattern.search(line)
            if account_match:
                current_account = account_match.group()
                # --- DEBUGGING START ---
                #self.results_text.insert(tk.END, f"Account found: {current_account}\n")
                # --- DEBUGGING END ---
            
             # Attempt to find payments using the specific patterns first
            payments = []

            match = elf_pay_pattern.search(line)
            if match:
                payments.append(match.group(1))
                # --- DEBUGGING START ---
                #self.results_text.insert(tk.END, f"ELF PAY match found: {match.group(1)}\n")
                # --- DEBUGGING END ---

            match = delete_bat_pattern.search(line)
            if match:
                payments.append(match.group(1))
                # --- DEBUGGING START ---
                #self.results_text.insert(tk.END, f"DELETE BAT match found: {match.group(1)}\n")
                # --- DEBUGGING END ---

            match = credit_card_pattern.search(line)
            if match:
                payments.append(match.group(1))
                # --- DEBUGGING START ---
                #self.results_text.insert(tk.END, f"CREDIT CARD match found: {match.group(1)}\n")
                # --- DEBUGGING END ---
            
            # --- DEBUGGING START ---
            #self.results_text.insert(tk.END, f"Payments collected for line: {payments}\n")
            #self.results_text.insert(tk.END, f"Current Account before adding to total: {current_account}\n")
            # --- DEBUGGING END ---

            if payments and current_account:
                for payment_str in payments:
                    try:
                        amount = float(payment_str.replace(" ", ""))
                        account_totals[current_account] += amount
                        account_counts[current_account] += 1
                    except ValueError:
                        self.results_text.insert(tk.END, f"Warning: Could not parse payment amount from '{payment_str}' for account {current_account}.\n")
                        self.master.update_idletasks()

        self.results_text.insert(tk.END, f"Extracted {len(account_totals)} unique accounts from PDF text.\n")
        self.master.update_idletasks()

        return account_totals

    # This method takes the input .xlsx file and finds all account numbers with their associated payments 
    def _get_info_from_xlsx_data(self, input_xlsx_path):
        
        self.results_text.insert(tk.END, "Reading data from Excel...\n")
        self.master.update_idletasks()
        
        try:
            df = pd.read_excel(input_xlsx_path, engine="openpyxl", dtype=str)
        except Exception as e:
            self.results_text.insert(tk.END, f"Error reading Excel file: {e}\n")
            self.master.update_idletasks()
            raise

        # Validate required columns 
        account_col = 'merchant_defined_field_1'
        amount_col = 'amount'

        if amount_col not in df.columns:
            raise ValueError(f"Required column '{amount_col}' not found in Excel file '{input_xlsx_path}'.")
        if account_col not in df.columns:
            raise ValueError(f"Required column '{account_col}' not found in Excel file '{input_xlsx_path}'.")

        account_totals = defaultdict(float)
        account_counts = defaultdict(int)

        # Account pattern 
        account_pattern = re.compile(r'\bW\d{6,7}\b')

        for index, row in df.iterrows():
            amount_cell = row[amount_col]
            account_cell = row[account_col]

            # Try to convert amount
            try:
                amount = float(str(amount_cell).strip())
            except (ValueError, TypeError):
                self.results_text.insert(tk.END, f"Warning: Skipping row {index+2} due to invalid amount: '{amount_cell}'.\n")
                continue # Skip rows with invalid amounts

            # Extract account number from messy string
            if isinstance(account_cell, str):
                account_match = account_pattern.search(account_cell)
                if account_match:
                    account = account_match.group()
                    account_totals[account] += amount
                    account_counts[account] += 1
                else:
                    self.results_text.insert(tk.END, f"Warning: No valid account number (Wxxxxxx/Wxxxxxxx) found in '{account_cell}' for row {index+2}.\n")
            else:
                self.results_text.insert(tk.END, f"Warning: Account cell content is not a string for row {index+2}: '{account_cell}'.\n")

        self.results_text.insert(tk.END, f"Extracted {len(account_totals)} unique accounts from Excel.\n")
        self.master.update_idletasks()
        return account_totals

    # This method takes all records from _get_info_from_pdf_text and _get_info_from_xlsx_data and compares the data and prints the differences 
    def _compare_data(self, pdf_data, excel_data):
        
        self.results_text.insert(tk.END, "Comparing data...\n\n")
        self.master.update_idletasks()

        output_lines = []
        matched_accounts = set()

        import difflib 

        # Compare PDF data against Excel data
        for pdf_acc, pdf_amt in pdf_data.items():
            if pdf_acc in excel_data:
                excel_amt = excel_data[pdf_acc]
                if abs(pdf_amt - excel_amt) < 0.01: # Check if difference is less than 1 cent
                    output_lines.append(f"✅ Account {pdf_acc} matches: ${pdf_amt:.2f} in both files.\n")
                else:
                    output_lines.append(f"⚠️ Mismatch payment for account {pdf_acc}: PDF = ${pdf_amt:.2f}, Excel = ${excel_amt:.2f}\n")
                    output_lines.append(f"    Check PDF and Excel with account number...\n")
                matched_accounts.add(pdf_acc)
            else:
                # Try fuzzy matching only on unmatched Excel accounts for same amount
                possible_matches = [
                    acc for acc in excel_data if acc not in matched_accounts and abs(excel_data[acc] - pdf_amt) < 0.01
                ]
                # Using a higher cutoff for closer matches
                close_matches = difflib.get_close_matches(pdf_acc, possible_matches, n=1, cutoff=0.83)
                if close_matches:
                    match = close_matches[0]
                    output_lines.append(f"✅ Approximate match: PDF account {pdf_acc} ≈ Excel account {match}, both have amount ${pdf_amt:.2f}\n")
                    output_lines.append(f"    Check PDF with Excel account number...\n")
                    matched_accounts.add(match) # Mark the Excel match as handled
                else:
                    output_lines.append(f"❌ Account {pdf_acc} found in PDF but missing in Excel. Amount found = ${pdf_amt:.2f}\n")

        # Check Excel accounts not yet matched (those only in Excel or not found by fuzzy match)
        for excel_acc, excel_amt in excel_data.items():
            if excel_acc not in matched_accounts:
                output_lines.append(f"❌ Account {excel_acc} found in Excel but missing in PDF. Amount found = ${excel_amt:.2f}\n")

        if not output_lines:
            output_lines.append("No accounts found in either file for comparison or all matched perfectly.")
        
        self.results_text.insert(tk.END, "\n".join(output_lines))
        self.results_text.insert(tk.END, "\nComparison complete.\n")
        self.master.update_idletasks()

    # This method orchestrates the entire comparison 
    def run_comparison(self):

        pdf_path = self.pdf_file_path.get()
        excel_path = self.excel_file_path.get()

        self.results_text.delete(1.0, tk.END) # Clear previous results
        self.results_text.insert(tk.END, "Starting comparison...\n")
        self.results_text.insert(tk.END, f"PDF: {os.path.basename(pdf_path) if pdf_path else 'Not selected'}\n")
        self.results_text.insert(tk.END, f"Excel: {os.path.basename(excel_path) if excel_path else 'Not selected'}\n\n")
        self.master.update_idletasks() # Update GUI to show messages immediately

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file.")
            self.results_text.insert(tk.END, "Error: PDF file not found or not selected.\n")
            return
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            self.results_text.insert(tk.END, "Error: Excel file not found or not selected.\n")
            return

        try:
            pdf_full_text = self._get_pdf_full_text(pdf_path)
            pdf_data = self._get_info_from_pdf_text(pdf_full_text)
            excel_data = self._get_info_from_xlsx_data(excel_path)
            self._compare_data(pdf_data, excel_data)
            messagebox.showinfo("Comparison Complete", "Comparison finished successfully! Check the results area.")

        except pytesseract.TesseractNotFoundError:
            # This specific error handling is for when Tesseract isn't found at all,
            # even after attempting to resolve its path, which means bundling failed or path is wrong.
            error_msg = ("Tesseract OCR engine could not be found. "
                         "If running from a bundled application, ensure Tesseract was included "
                         "correctly during packaging. If running from script, check the "
                         "pytesseract.pytesseract.tesseract_cmd path and Tesseract installation.")
            messagebox.showerror("OCR Engine Error", error_msg)
            self.results_text.insert(tk.END, f"Error: {error_msg}\n")
        except FileNotFoundError as e:
            messagebox.showerror("File Error", f"File not found: {e}\n"
                                               "Please ensure the file exists and is accessible.")
            self.results_text.insert(tk.END, f"Error: File not found - {e}\n")
        except pd.errors.EmptyDataError:
            messagebox.showerror("Excel Error", "The Excel file is empty or contains no data.")
            self.results_text.insert(tk.END, "Error: Excel file is empty or contains no data.\n")
        except pd.errors.ParserError:
            messagebox.showerror("Excel Error", "Could not parse Excel file. Is it a valid format and not corrupted?")
            self.results_text.insert(tk.END, "Error: Could not parse Excel file. Invalid format or corrupted?\n")
        except ValueError as e:
            messagebox.showerror("Data Error", f"Data processing error: {e}")
            self.results_text.insert(tk.END, f"Data processing error: {e}\n")
        except Exception as e:
            # Catch any other unexpected errors
            messagebox.showerror("An unexpected error occurred", f"An unexpected error occurred: {e}")
            self.results_text.insert(tk.END, f"An unexpected error occurred: {e}\n")


def main():
    root = tk.Tk()
    app = PDFExcelComparerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()