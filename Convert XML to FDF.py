# === FORM TEMPLATES ===
# FinalSheet2024
# SWDISC-NoCalculations

# === IMPORTS ===
import os
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext

def sanitize_for_fdf(text):
    """Replace special characters with their closest ASCII equivalents."""
    replacements = {
        '\u2018': "'",  # Left single quote
        '\u2019': "'",  # Right single quote
        '\u201C': '"',  # Left double quote
        '\u201D': '"',  # Right double quote
        '\u2013': '-',  # En dash
        '\u2014': '--', # Em dash
        '\u2026': '...',# Ellipsis
        '\u00A0': ' ',  # Non-breaking space
    }
    
    # Replace known special characters
    for special, replacement in replacements.items():
        text = text.replace(special, replacement)
    
    # Replace any remaining non-latin1 characters with '?'
    return text.encode('latin-1', errors='replace').decode('latin-1')

# === NAMESPACES FROM XML FILE===
namespaces = {
    'dfs': 'http://schemas.microsoft.com/office/infopath/2003/dataFormSolution',
    'q': 'http://schemas.microsoft.com/office/infopath/2003/ado/queryFields',
    'd': 'http://schemas.microsoft.com/office/infopath/2003/ado/dataFields',
    'my': 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2005-04-28T12:10:52'
}

# === FUNCTION TO SELECT INPUT FOLDER ===
def select_input_folder():
    tk.Tk().withdraw()
    folder_path = filedialog.askdirectory(title="Select the folder containing XML files")
    if not folder_path:
        messagebox.showwarning("Error", "No folder selected. The program will exit.")
        exit()
    return folder_path

# === DATE FORMATTING ===
date_patterns = [
    r"^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$", # YYYY-MM-DDTHH:MM:SS
    r"(\d{4})-(\d{2})-(\d{2})",   # YYYY-MM-DD
    r"(\d{2})/(\d{2})/(\d{4})",   # MM/DD/YYYY
    r"(\d{4})/(\d{2})/(\d{2})",   # YYYY/MM/DD
]
def format_date(value, file_name, progress_dialog):
    value = value.strip()  # Remove leading/trailing whitespace
    for pattern in date_patterns:
        match = re.match(pattern, value)
        if match:
            try:
                if len(match.groups()) == 3:  # Handle basic date formats
                    if len(match.group(1)) == 4:  # YYYY-MM-DD or YYYY/MM/DD
                        date = datetime.strptime(value, "%Y-%m-%d") if "-" in value else datetime.strptime(value, "%Y/%m/%d")
                    else:  # MM/DD/YYYY
                        date = datetime.strptime(value, "%m/%d/%Y")
                elif len(match.groups()) == 6:  # Handle date with time (YYYY-MM-DDTHH:MM:SS)
                    date = datetime.strptime(value, "%Y-%m-%dT%H:%M:%S")
            
                return date.strftime("%m/%d/%y")
            except Exception as e:
                progress_dialog.increment_date_errors()
                progress_dialog.log(f"⚠️ Date format error in file '{file_name}' for value: '{value}'")
                return value
    return value


def main():
    input_folder = select_input_folder()
    progress_dialog = ProgressDialog()

    # === WALK THROUGH INPUT FOLDER AND SUBFOLDERS SCANNING FOR XML FILES===
    progress_dialog.log("\n=== START ===\n")
    for root_dir, _, files in os.walk(input_folder):
        xml_files = [f for f in files if f.lower().endswith('.xml')]
        if not xml_files:
            continue

        folder_name = os.path.basename(root_dir.rstrip("\\/"))
        parent_dir = os.path.dirname(root_dir.rstrip("\\/"))
        output_subfolder = os.path.join(parent_dir, f"{folder_name} - CONVERTED")
        
        if not os.path.exists(output_subfolder):
            os.makedirs(output_subfolder)

        for xml_file in xml_files:
            input_file = os.path.join(root_dir, xml_file)
            output_file = os.path.join(output_subfolder, f"{os.path.splitext(xml_file)[0]}.fdf")

            try:
                tree = ET.parse(input_file)
                root = tree.getroot()

                # === FIND MASTER_PART1 ELEMENT UNDER d AND q NAMESPACES ===
                master_d = root.findall('.//d:MASTER_PART1', namespaces)
                master_q = root.findall('.//q:MASTER_PART1', namespaces)

                field_counter = defaultdict(int)
                fdf_fields = []

                # === EXTRACT FIELDS FROM q:MASTER_PART1 ===
                for master in master_q:
                    for attr, value in master.attrib.items():
                        if value is None or value.strip() == "":
                            continue
                        value = format_date(value, xml_file, progress_dialog)
                        value = sanitize_for_fdf(value)
                        field_counter[attr] += 1
                        field_name = attr if field_counter[attr] == 1 else f"{attr}_{field_counter[attr]}"
                        value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")

                # === EXTRACT FIELDS FROM d:MASTER_PART1 ===
                for master in master_d:
                    for attr, value in master.attrib.items():
                        if value is None or value.strip() == "":
                            continue
                        value = format_date(value, xml_file, progress_dialog)
                        value = sanitize_for_fdf(value)
                        field_counter[attr] += 1
                        field_name = attr if field_counter[attr] == 1 else f"{attr}_{field_counter[attr]}"
                        value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")
                
                # Process all children with my: namespace directly from root
                for my_field in root:
                    if my_field.tag.startswith('{' + namespaces['my'] + '}'):
                        field_name = my_field.tag.split('}')[1]
                        value = my_field.text if my_field.text else ''
                        
                        # Skip if empty or nil
                        if not value or value.strip() == '' or my_field.get('{http://www.w3.org/2001/XMLSchema-instance}nil') == 'true':
                            continue
                            
                        # Process the value
                        value = format_date(value, xml_file, progress_dialog)
                        value = sanitize_for_fdf(value)
                        value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")

                # === CREATE FDF CONTENT ===
                fdf_content = f"""%FDF-1.2
                1 0 obj
                << /FDF <<
                /F (C:/Users/rmarcum/Desktop/FinalSheet2024.pdf)
                /Fields [
                {chr(10).join(fdf_fields)}
                ] >> >>
                endobj
                trailer
                << /Root 1 0 R >>
                %%EOF
                """

                # === WRITE TO FDF FILE ===
                with open(output_file, "wb") as f:
                    f.write(fdf_content.encode("latin-1"))

                progress_dialog.increment_success()
                progress_dialog.log(f"✅ FDF created: {output_file}")

            except Exception as e:
                progress_dialog.increment_failure(xml_file, str(e))
                progress_dialog.log(f"❌ Error processing file {xml_file}: {e}")

    progress_dialog.show_final_summary()
    progress_dialog.window.wait_window(progress_dialog.window)

class ProgressDialog:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("XML to FDF Conversion Progress")
        self.window.geometry("900x600")
        self.error_details = []
        
        # Create and pack the text area
        self.text_area = scrolledtext.ScrolledText(self.window, wrap=tk.WORD, width=70, height=20)
        self.text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Create progress bar
        self.progress_var = tk.StringVar(value="Processing...")
        self.progress_label = ttk.Label(self.window, textvariable=self.progress_var)
        self.progress_label.pack(pady=5)
        
        # Create close button (initially disabled)
        self.close_button = ttk.Button(self.window, text="Close", command=self.window.destroy)
        self.close_button.pack(pady=5)
        self.close_button.config(state='disabled')
        
        # Initialize counts
        self.success_count = 0
        self.failure_count = 0
        self.date_format_errors = 0
        self.my_count = 0
        
    def log(self, message):
        self.text_area.insert(tk.END, message + "\n")
        self.text_area.see(tk.END)
        self.window.update()
        
    def show_final_summary(self):
        summary = "\n=== Conversion Summary ===\n"
        summary += f"Successfully converted {self.success_count} file(s).\n"
        summary += f"Failed to convert {self.failure_count} file(s).\n"
        summary += f"Date formatting errors encountered in {self.date_format_errors} field(s).\n"
        summary += f"MYCOUNT: {self.my_count}\n"

        if self.error_details:
            summary += "\n=== Error Details ===\n"
            for file_name, error_msg in self.error_details:
                summary += f"File: {file_name}\n"
                summary += f"Error: {error_msg}\n"
                summary += "-" * 50 + "\n"

        self.log(summary)
        self.progress_var.set("Conversion Complete")
        self.close_button.config(state='normal')
        
    def increment_success(self):
        self.success_count += 1
        
    def increment_failure(self, file_name, error_msg):
        self.failure_count += 1
        self.error_details.append((file_name, error_msg))

        
    def increment_date_errors(self):
        self.date_format_errors += 1
        
    def set_my_count(self, count):
        self.my_count = count

if __name__ == "__main__":
    main()
