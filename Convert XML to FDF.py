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
import re

# === DATE FORMATTING ===
date_patterns = [
    r"^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$", # YYYY-MM-DDTHH:MM:SS
    r"(\d{4})-(\d{2})-(\d{2})",   # YYYY-MM-DD
    r"(\d{2})/(\d{2})/(\d{4})",   # MM/DD/YYYY
    r"(\d{4})/(\d{2})/(\d{2})",   # YYYY/MM/DD
]

# === EXTRACT NAMESPACES ===
def extract_namespaces(xml_file, progress_dialog):
    try:
        with open(xml_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Find all namespace declarations like xmlns:prefix="uri"
        ns_dict = {}
        ns_matches = re.findall(r'xmlns:([a-zA-Z0-9]+)="([^"]+)"', content)
        for prefix, uri in ns_matches:
            ns_dict[prefix] = uri
            
        progress_dialog.log(f"✓ Found {len(ns_dict)} namespaces in {os.path.basename(xml_file)}")
        return ns_dict
    except Exception as e:
        progress_dialog.log(f"⚠️ Error extracting namespaces from {os.path.basename(xml_file)}: {str(e)}")
        # Return default namespaces if extraction fails
        return {
            'dfs': 'http://schemas.microsoft.com/office/infopath/2003/dataFormSolution',
            'q': 'http://schemas.microsoft.com/office/infopath/2003/ado/queryFields',
            'd': 'http://schemas.microsoft.com/office/infopath/2003/ado/dataFields',
            'my': 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2005-05-04T13:10:26'
        }

# === SELECT INPUT FOLDER ===
def select_input_folder():
    tk.Tk().withdraw()
    folder_path = filedialog.askdirectory(title="Select the folder containing XML files")
    if not folder_path:
        messagebox.showwarning("Error", "No folder selected. The program will exit.")
        exit()
    return folder_path

# === FORMAT DATE ===
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

# === PROCESS XML DATA WITH NAMESPACES ===
def process_xml(xml_file, output_file, namespaces, progress_dialog):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        field_counter = defaultdict(int)
        fdf_fields = []
        
        # Try to find and process MASTER_PART1 elements from all relevant namespaces
        for prefix, uri in namespaces.items():
            # Look for MASTER_PART1 elements in this namespace
            masters = root.findall(f'.//{prefix}:MASTER_PART1', namespaces)
            for master in masters:
                for attr, value in master.attrib.items():
                    if value is None or value.strip() == "":
                        continue
                    value = format_date(value, os.path.basename(xml_file), progress_dialog)
                    field_counter[attr] += 1
                    field_name = attr if field_counter[attr] == 1 else f"{attr}_{field_counter[attr]}"
                    value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")
        
        # Look for myFields in any namespace (typically dfs:)
        my_prefix = next((p for p in namespaces if p == 'dfs'), None)
        my_count = 0
        
        if my_prefix:
            my_fields = root.find(f'.//{my_prefix}:myFields', namespaces)
            if my_fields is not None:
                # Look for namespace prefixes that could contain field data
                field_prefixes = [p for p in namespaces if p not in ('xsi', 'xsd', 'xhtml')]
                
                for my_field in my_fields:
                    my_count += 1
                    
                    # Check if tag matches any of our known namespaces
                    namespace_match = False
                    for prefix in field_prefixes:
                        ns_uri = namespaces.get(prefix, '')
                        if ns_uri and my_field.tag.startswith('{' + ns_uri + '}'):
                            namespace_match = True
                            field_name = my_field.tag.split('}')[1]
                            value = my_field.text if my_field.text else ''
                            
                            # Process attributes
                            for attr_name, attr_value in my_field.attrib.items():
                                if attr_name.startswith('{' + ns_uri + '}'):
                                    nested_field_name = attr_name.split('}')[1]
                                    if attr_value and attr_value.strip():
                                        fdf_fields.append(f"<< /T ({field_name}_{nested_field_name}) /V ({attr_value}) >>")
                            
                            # Process field value
                            if value and value.strip():
                                value = format_date(value, os.path.basename(xml_file), progress_dialog)
                                value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                                fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")
                            
                            break
                    
                    # If no namespace match found, try to extract field name from tag directly
                    if not namespace_match and '}' in my_field.tag:
                        field_name = my_field.tag.split('}')[1]
                        value = my_field.text if my_field.text else ''
                        
                        if value and value.strip():
                            value = format_date(value, os.path.basename(xml_file), progress_dialog)
                            value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                            fdf_fields.append(f"<< /T ({field_name}) /V ({value}) >>")
        
        progress_dialog.set_my_count(my_count)
        
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
        return True
        
    except Exception as e:
        progress_dialog.increment_failure()
        progress_dialog.log(f"❌ Error processing file {os.path.basename(xml_file)}: {e}")
        return False

def main():
    input_folder = select_input_folder()
    progress_dialog = ProgressDialog()

    # === WALK THROUGH INPUT FOLDER AND SUBFOLDERS SCANNING FOR XML FILES===
    progress_dialog.log("----------------START -------------------")
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
            
            # Extract namespaces dynamically for this XML file
            namespaces = extract_namespaces(input_file, progress_dialog)
            
            # Process the XML file with the extracted namespaces
            process_xml(input_file, output_file, namespaces, progress_dialog)

    progress_dialog.show_final_summary()
    progress_dialog.window.mainloop()

# === PROGRESS DIALOG CLASS ===
class ProgressDialog:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("XML to FDF Conversion Progress")
        self.window.geometry("900x600")
        
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
        self.log(summary)
        self.progress_var.set("Conversion Complete")
        self.close_button.config(state='normal')
        
    def increment_success(self):
        self.success_count += 1
        
    def increment_failure(self):
        self.failure_count += 1
        
    def increment_date_errors(self):
        self.date_format_errors += 1
        
    def set_my_count(self, count):
        self.my_count = count

if __name__ == "__main__":
    main()
