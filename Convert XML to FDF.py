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
    return text.encode('latin-1', errors="ignore").decode('latin-1')

#=== NAMESPACES FROM XML FILE===
namespaces = {
    'dfs': 'http://schemas.microsoft.com/office/infopath/2003/dataFormSolution',
    'q': 'http://schemas.microsoft.com/office/infopath/2003/ado/queryFields',
    'd': 'http://schemas.microsoft.com/office/infopath/2003/ado/dataFields',
    'my': 'http://schemas.microsoft.com/office/infopath/2003/myXSD/',
}

# === SELECT INPUT FOLDER ===
def select_input_folder():
    """Select the folder containing XML files to convert."""
    tk.Tk().withdraw()
    folder_path = filedialog.askdirectory(title="Select the folder containing XML files")
    if not folder_path:
        messagebox.showwarning("Error", "No folder selected. The program will exit.")
        exit()
    return folder_path

def select_template_pdf():
    """Select the PDF template file to use for the conversion."""
    tk.Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="Select the PDF template file",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showwarning("Error", "No PDF template selected. The program will exit.")
        exit()
    return file_path

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
                progress_dialog.add_date_error(file_name, value)
                progress_dialog.log(f"⚠️ Date format error in file '{file_name}' for value: '{value}'")
                return value
    return value

def main():
    # Show initial setup dialog
    setup_dialog = InitialSetupDialog()
    result = setup_dialog.run()
    
    if result is None:
        return
        
    input_folder, template_pdf = result
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
                        progress_dialog.add_unique_field(field_name)

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
                        progress_dialog.add_unique_field(field_name)
                
                # Process all children with my: namespace directly from root
                for my_field in root:
                    if my_field.tag.startswith('{' + namespaces['my']):
                        # Get the main field name without namespace
                        field_name = my_field.tag.split('}')[1]
                        value = my_field.text if my_field.text else ''
                        
                        # Process the main field value if not empty or nil
                        if value and value.strip() and my_field.get('{http://www.w3.org/2001/XMLSchema-instance}nil') != 'true':
                            value = format_date(value, xml_file, progress_dialog)
                            value = sanitize_for_fdf(value)
                            value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                            fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")
                            progress_dialog.add_unique_field(field_name)

                        # Process attributes (sub-fields)
                        for attr_name, attr_value in my_field.attrib.items():
                            # Check if the attribute has the my: prefix
                            if attr_name.startswith('{' + namespaces['my']):
                                # Get the sub-field name without namespace
                                sub_field_name = attr_name.split('}')[1]
                                
                                if attr_value and attr_value.strip():
                                    # Process the sub-field value
                                    attr_value = format_date(attr_value, xml_file, progress_dialog)
                                    attr_value = sanitize_for_fdf(attr_value)
                                    attr_value = attr_value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                                    fdf_fields.append(f"<< /T ({sub_field_name}) /V ({attr_value}) >>")
                                    progress_dialog.add_unique_field(sub_field_name)

                # === CREATE FDF CONTENT ===
                fdf_content = f"""%FDF-1.2
                1 0 obj
                << /FDF <<
                /F ({template_pdf})
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
        self.date_error_details = []  # New list to store date formatting errors
        
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
        self.unique_fields = []
        
    def log(self, message):
        self.text_area.insert(tk.END, message + "\n")
        self.text_area.see(tk.END)
        self.window.update()
        
    def show_final_summary(self):
        summary = "\n=== Conversion Summary ===\n"
        summary += f"Successfully converted {self.success_count} file(s).\n"
        summary += f"Failed to convert {self.failure_count} file(s).\n"
        summary += f"Date formatting errors encountered in {self.date_format_errors} field(s).\n"

        fields_found = "\n=== Unique Fields Found ===\n"
        self.unique_fields.sort()
        for field in self.unique_fields:
            fields_found += f"{field}\n"

        if self.error_details:
            summary += "\n=== Error Details ===\n"
            for file_name, error_msg in self.error_details:
                summary += f"File: {file_name}\n"
                summary += f"Error: {error_msg}\n"
                summary += "-" * 50 + "\n"

        if self.date_error_details:
            summary += "\n=== Date Formatting Error Details ===\n"
            for file_name, field_value in self.date_error_details:
                summary += f"File: {file_name}\n"
                summary += f"Invalid date value: {field_value}\n"
                summary += "-" * 50 + "\n"

        self.log(summary)
        self.log(fields_found)
        self.progress_var.set("Conversion Complete")
        self.close_button.config(state='normal')
        
    def increment_success(self):
        self.success_count += 1
        
    def increment_failure(self, file_name, error_msg):
        self.failure_count += 1
        self.error_details.append((file_name, error_msg))

    def increment_date_errors(self):
        self.date_format_errors += 1
        
    def add_date_error(self, file_name, value):
        self.date_format_errors += 1
        self.date_error_details.append((file_name, value))

    def add_unique_field(self, field):
        if field not in self.unique_fields:
            self.unique_fields.append(field)

class InitialSetupDialog:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("XML to FDF Converter Setup")
        self.window.geometry("900x600")
        
        # Input folder selection
        self.input_folder = tk.StringVar()
        input_frame = ttk.LabelFrame(self.window, text="Input Folder", padding="10")
        input_frame.pack(fill="x", padx=10, pady=5)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_folder, width=50)
        self.input_entry.pack(side="left", padx=5)
        
        input_button = ttk.Button(input_frame, text="Choose Folder", command=self.choose_input_folder)
        input_button.pack(side="left", padx=5)
        
        # Template PDF selection
        self.template_pdf = tk.StringVar()
        template_frame = ttk.LabelFrame(self.window, text="PDF Template", padding="10")
        template_frame.pack(fill="x", padx=10, pady=5)
        
        self.template_entry = ttk.Entry(template_frame, textvariable=self.template_pdf, width=50)
        self.template_entry.pack(side="left", padx=5)
        
        template_button = ttk.Button(template_frame, text="Choose PDF", command=self.choose_template_pdf)
        template_button.pack(side="left", padx=5)
        
        # Start button
        self.start_button = ttk.Button(self.window, text="Start Conversion", command=self.start_conversion)
        self.start_button.pack(pady=20)
        
        # Status label
        self.status_var = tk.StringVar(value="Please select input folder and PDF template")
        self.status_label = ttk.Label(self.window, textvariable=self.status_var, wraplength=550)
        self.status_label.pack(pady=5)
        
        self.result = None
        
    def choose_input_folder(self):
        folder_path = filedialog.askdirectory(title="Select the folder containing XML files")
        if folder_path:
            self.input_folder.set(folder_path)
            self.update_status()
    
    def choose_template_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select the PDF template file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.template_pdf.set(file_path)
            self.update_status()
    
    def update_status(self):
        if not self.input_folder.get() and not self.template_pdf.get():
            self.status_var.set("Please select input folder and PDF template")
        elif not self.input_folder.get():
            self.status_var.set("Please select input folder")
        elif not self.template_pdf.get():
            self.status_var.set("Please select PDF template")
        else:
            self.status_var.set("Ready to start conversion")
    
    def start_conversion(self):
        if not self.input_folder.get():
            messagebox.showwarning("Error", "Please select an input folder")
            return
        if not self.template_pdf.get():
            messagebox.showwarning("Error", "Please select a PDF template")
            return
            
        self.result = (self.input_folder.get(), self.template_pdf.get())
        self.window.destroy()
    
    def run(self):
        self.window.mainloop()
        return self.result

if __name__ == "__main__":
    main()