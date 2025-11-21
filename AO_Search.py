#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Nov 21 10:07:38 2025

@author: kseeks61
"""


import pandas as pd
import re
import rapidfuzz as rf
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import logging

class CompanyMatcher:
    def __init__(self):
        self.df = None
        self.data_cleaned = False

    def name_cleaning(self, name):
        cleaned_name = re.sub(r'[^A-Za-z0-9\s]+', '', str(name)).upper()
        cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
        return cleaned_name

    def is_word_subsequence(self, search_words, target_words):
        if not search_words or not target_words:
            return False
        i = 0
        for word in target_words:
            if i < len(search_words) and search_words[i] == word:
                i += 1
        return i == len(search_words)

    def find_duplicates(self, company_name, name_threshold, data, text_output):
        cleaned_company_name = self.name_cleaning(company_name)
        matches = {}
        for index, row in data.iterrows():
            name_str = str(row['cleaned_name']) if pd.notnull(row['cleaned_name']) else ""
            try:
                name_ratio = rf.fuzz.ratio(cleaned_company_name, name_str) / 100.0
            except TypeError:
                text_output.insert(tk.END, f"Warning: Invalid name data at index {index}: {row['cleaned_name']}\n")
                text_output.see(tk.END)
                continue
            company_words = cleaned_company_name.split()
            name_words = name_str.split()
            is_subsequence = self.is_word_subsequence(company_words, name_words)
           # simple cleaned substring check
            normalized_input = cleaned_company_name.replace(' ', '')
            normalized_candidate = name_str.replace(' ', '')
            substring_match = (
                cleaned_company_name in name_str or
                name_str in cleaned_company_name or
                normalized_input in normalized_candidate or
                normalized_candidate in normalized_input
                )
            if (name_ratio >= name_threshold or is_subsequence or substring_match):
                matches[row['CUCC Code']] = [row['Account Name'], row['Address'], row['State'], row['Responsible User']]
        return matches

    def save_matches_to_excel(self, matches, root):
        try:
            if not matches:
                return False
            df_matches = pd.DataFrame.from_dict(matches, orient='index', columns=["Account Name", "Address", "State", "Responsible User"])
            df_matches.reset_index(inplace=True)
            df_matches.rename(columns={'index': 'CUCC Code'}, inplace=True)

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save AO Search Report",
                initialfile="AO_Search.xlsx"
            )
            if save_path:
                df_matches.to_excel(save_path, index=False, engine='openpyxl')
                return save_path
            return False
        except Exception as e:
            logging.error(f"Error saving Excel file: {e}")
            return False

    def process_csv(self, file_path, company_name, name_threshold, text_output, root):
        try:
            if not self.data_cleaned:
                self.df = pd.read_csv(file_path)
                self.df.columns = self.df.columns.str.strip()
                required_columns = ["Account Name", "CUCC Code"]
                if not all(col in self.df.columns for col in required_columns):
                    raise KeyError(f"Missing one or more required columns: {required_columns}")
                self.df['Account Name'] = self.df['Account Name'].astype(str).replace('nan', '')
                text_output.insert(tk.END, "Precomputing cleaned company names...\n")
                text_output.see(tk.END)
                root.update()
                self.df['cleaned_name'] = self.df['Account Name'].apply(self.name_cleaning)
                self.data_cleaned = True
                text_output.insert(tk.END, "Data cleaning complete.\n")
            else:
                text_output.insert(tk.END, "Using previously cleaned data.\n")

            matches = self.find_duplicates(company_name, name_threshold, self.df, text_output)
            if matches:
                text_output.insert(tk.END, "\nPotential customers found:\n")
                for cucc_code, details in matches.items():
                    text_output.insert(tk.END, f"CUCC Code: {cucc_code}\n")
                    text_output.insert(tk.END, f"Account Name: {details[0]}\n")
                    text_output.insert(tk.END, f"Address: {details[1]}\n")
                    text_output.insert(tk.END, f"State: {details[2]}\n")
                    text_output.insert(tk.END, f"Responsible User: {details[3]}\n")

                # ✅ Save to Excel
                save_result = self.save_matches_to_excel(matches, root)
                if save_result:
                    text_output.insert(tk.END, f"\nResults successfully saved to: {save_result}\n")
                else:
                    text_output.insert(tk.END, "\nUser canceled save or an error occurred during export.\n")
            else:
                text_output.insert(tk.END, "\nNo customers found. Consider adjusting the thresholds or checking the input data.\n")

            text_output.see(tk.END)
            return True
        except FileNotFoundError:
            text_output.insert(tk.END, "Error: The CSV file was not found. Please check the file path.\n")
            text_output.see(tk.END)
            return False
        except KeyError as e:
            text_output.insert(tk.END, f"Error: {e}\n")
            text_output.see(tk.END)
            return False
        except Exception as e:
            text_output.insert(tk.END, f"An unexpected error occurred: {e}\n")
            text_output.see(tk.END)
            return False

    def browse_file(self, entry_file, text_output):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            entry_file.delete(0, tk.END)
            entry_file.insert(0, file_path)
            text_output.insert(tk.END, f"Selected file: {file_path}\n")
            text_output.see(tk.END)

    def validate_threshold(self, value):
        try:
            val = float(value)
            return 0 <= val <= 1
        except ValueError:
            return False

    def process_action(self, entry_file, entry_company, entry_name_thresh, text_output, process_button, copy_button, root):
        file_path = entry_file.get().strip()
        company_name = entry_company.get().strip()
        name_threshold = entry_name_thresh.get().strip()

        if not file_path:
            messagebox.showerror("Error", "Please select a CSV file.")
            return
        if not os.path.isfile(file_path):
            messagebox.showerror("Error", "File does not exist or path is invalid.")
            return
        if not company_name:
            messagebox.showerror("Error", "Company name cannot be empty.")
            return
        if not self.validate_threshold(name_threshold):
            messagebox.showerror("Error", "Name threshold must be a number between 0 and 1.")
            return

        process_button.config(state='disabled')
        copy_button.config(state='disabled')
        text_output.insert(tk.END, f"Processing file: {file_path} (this may take a while for large datasets)...\n")
        text_output.see(tk.END)
        root.update()

        def run_process():
            success = self.process_csv(file_path, company_name, float(name_threshold), text_output, root)
            root.after(0, lambda: self.post_process(success, entry_file, entry_company, entry_name_thresh, text_output, process_button, copy_button, root))

        thread = threading.Thread(target=run_process)
        thread.start()

    def post_process(self, success, entry_file, entry_company, entry_name_thresh, text_output, process_button, copy_button, root):
        process_button.config(state='normal')
        copy_button.config(state='normal')
        if success:
            entry_company.delete(0, tk.END)
            entry_name_thresh.delete(0, tk.END)
            entry_name_thresh.insert(0, "0.7")
            text_output.insert(tk.END, "Ready for next check.\n")
        else:
            text_output.insert(tk.END, "Processing failed. Please check the input and try again.\n")
        text_output.see(tk.END)

    def main(self):
        root = tk.Tk()
        root.title("AO Search")
        root.geometry("600x500")

        tk.Label(root, text="AO Search", font=("Arial", 14, "bold")).pack(pady=10)

        frame = tk.Frame(root)
        frame.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(frame, text="CSV File Path:").grid(row=0, column=0, sticky="w", pady=5)
        entry_file = tk.Entry(frame, width=50)
        entry_file.grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Browse", command=lambda: self.browse_file(entry_file, text_output)).grid(row=0, column=2, padx=5)

        tk.Label(frame, text="Company Name:").grid(row=2, column=0, sticky="w", pady=5)
        entry_company = tk.Entry(frame, width=50)
        entry_company.grid(row=2, column=1, columnspan=2, padx=5)

        tk.Label(frame, text="Name Threshold (0-1):").grid(row=4, column=0, sticky="w", pady=5)
        entry_name_thresh = tk.Entry(frame, width=10)
        entry_name_thresh.insert(0, "0.7")
        entry_name_thresh.grid(row=4, column=1, sticky="w", padx=5)

        text_output = scrolledtext.ScrolledText(root, height=15, width=70, wrap=tk.WORD)
        text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_output.insert(tk.END, "Welcome to the Customer Account Owner Search Program!\n")

        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        process_button = tk.Button(button_frame, text="Process", command=lambda: self.process_action(entry_file, entry_company, entry_name_thresh, text_output, process_button, copy_button, root))
        process_button.pack(side=tk.LEFT, padx=5)
        copy_button = tk.Button(button_frame, text="Copy Results", command=lambda: root.clipboard_append(text_output.get(1.0, tk.END)))
        copy_button.pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Exit", command=root.quit).pack(side=tk.LEFT, padx=5)

        root.mainloop()

if __name__ == "__main__":
    app = CompanyMatcher()
    app.main()
