import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import time
import win32com.client
from docxtpl import DocxTemplate
from docx2pdf import convert

class FileManager:
    """Handles file and directory operations for the report card generator."""
    
    @staticmethod
    def ensure_directory_exists(directory):
        """Create directory if it doesn't exist."""
        if not os.path.exists(directory):
            os.makedirs(directory)
    
    @staticmethod
    def get_absolute_path(base_dir, *path_parts):
        """Construct an absolute path from base directory and path parts."""
        return os.path.join(base_dir, *path_parts)

class DataProcessor:
    """Processes Excel data for report card generation."""
    
    def __init__(self, excel_path, mapping_path):
        """
        Initialize data processor with Excel file and mapping.
        
        :param excel_path: Path to the Excel file
        :param mapping_path: Path to the JSON mapping file
        """
        self.df = self._load_dataframe(excel_path)
        self.column_map = self._load_mapping(mapping_path)
    
    def _load_dataframe(self, excel_path):
        """
        Load and preprocess Excel dataframe.
        
        :param excel_path: Path to the Excel file
        :return: Processed DataFrame
        """
        try:
            df = pd.read_excel(excel_path)
            df = df.drop([0])
            df.columns = df.iloc[0]
            return df[1:]
        except Exception as e:
            raise ValueError(f"Error loading Excel file: {e}")
    
    def _load_mapping(self, mapping_path):
        """
        Load column mapping from JSON file.
        
        :param mapping_path: Path to the mapping JSON
        :return: Column mapping dictionary
        """
        with open(mapping_path, "r") as file:
            return dict(json.load(file))
    
    def process_student_data(self, row, class_name):
        """
        Process individual student data for report card generation.
        
        :param row: DataFrame row for a student
        :param class_name: Name of the class
        :return: Processed student data dictionary
        """
        field_dict = {}
        for key in self.column_map.keys():
            if key not in ['percentage', 'remark', 'class']:
                field_dict[key] = row[self.column_map[key]]
            elif key == 'percentage':
                field_dict['percentage'] = f"{float(row[self.column_map['percentage']]):.2f}%"
            elif key == 'remark':
                field_dict['remark'] = row[self.column_map['remark']] + "!"
        
        field_dict['class'] = class_name.replace("_", " ")
        
        # Handle null/empty values
        nan_values = ['NAN', 'NONE', 'NA']
        field_dict = {
            key: value if value is not None and 
            str(value).upper().replace(" ", "") not in nan_values 
            else "---" 
            for key, value in field_dict.items()
        }
        
        return field_dict

class ReportCardGenerator:
    """Manages the generation of report cards."""
    
    def __init__(self, base_directory):
        """
        Initialize report card generator.
        
        :param base_directory: Base directory for project files
        """
        self.base_directory = base_directory
        self.input_folder = 'input_files'
        self.mappings_folder = 'mappings'
    
    def generate_report_cards(self, excel_filename, class_name):
        """
        Generate report cards for a given class.
        
        :param excel_filename: Name of the Excel file
        :param class_name: Name of the class
        """
        # Setup directories
        FileManager.ensure_directory_exists('word')
        report_cards_dir = f'{class_name} report_cards'
        FileManager.ensure_directory_exists(report_cards_dir)
        
        # Paths
        excel_path = FileManager.get_absolute_path(
            self.base_directory, self.input_folder, excel_filename
        )
        mapping_path = FileManager.get_absolute_path(
            self.base_directory, 
            self.mappings_folder, 
            f"{class_name.replace(' ', '_')}_mapping.json"
        )
        template_path = FileManager.get_absolute_path(
            os.getcwd(), self.input_folder, 'template-word1A.docx'
        )
        
        # Process data
        data_processor = DataProcessor(excel_path, mapping_path)
        
        # Generate Word documents
        for row in range(len(data_processor.df)):
            student_data = data_processor.process_student_data(
                data_processor.df.iloc[row], class_name
            )
            
            if student_data['name'] == '---':
                break
            
            self._create_word_document(
                template_path, student_data, 'word'
            )
        
        # Convert to PDF
        self._convert_to_pdf(data_processor.df, data_processor.column_map, 
                              class_name, report_cards_dir)
    
    def _create_word_document(self, template_path, student_data, output_dir):
        """
        Create a Word document for a student.
        
        :param template_path: Path to the template document
        :param student_data: Processed student data
        :param output_dir: Output directory for Word files
        """
        try:
            template = DocxTemplate(template_path)
            template.render(student_data)
            
            filename = f"{student_data['name']}.docx"
            filled_path = os.path.join(os.getcwd(), output_dir, filename)
            template.save(filled_path)
        except Exception as e:
            print(f"Error creating document: {e}")
            template.save("dummy.docx")
    
    def _convert_to_pdf(self, df, column_map, class_name, output_dir):
        """
        Convert Word documents to PDF.
        
        :param df: DataFrame with student data
        :param column_map: Column mapping dictionary
        :param class_name: Name of the class
        :param output_dir: Output directory for PDF files
        """
        for row in range(len(df)):
            student = df.iloc[row][column_map['name']]
            word_filename = f'{student}.docx'
            word_path = os.path.join(os.getcwd(), 'word', word_filename)
            pdf_path = os.path.join(os.getcwd(), output_dir, f'{student}.pdf')
            
            print(f'Generating Report card for {student}')
            convert(word_path, pdf_path)

class ReportCardGeneratorApp:
    """Tkinter GUI for Report Card Generator."""
    
    def __init__(self, root):
        """
        Initialize the application GUI.
        
        :param root: Tkinter root window
        """
        self.root = root
        self.root.title("Report Card Generator")
        
        # Base project directory (modify as needed)
        self.base_directory = os.path.dirname(__file__)
        
        # Variables
        self.excel_file_path = tk.StringVar()
        self.class_name = tk.StringVar()
        
        # Create UI
        self._create_ui()
    
    def _create_ui(self):
        """Create the user interface components."""
        notebook = ttk.Notebook(self.root)
        tab_run_script = ttk.Frame(notebook)
        notebook.add(tab_run_script, text="Run Script")
        notebook.pack(expand=1, fill="both")
        
        # Excel File Selection
        ttk.Label(tab_run_script, text="Choose Excel File:").grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )
        ttk.Entry(
            tab_run_script, 
            textvariable=self.excel_file_path, 
            state="readonly"
        ).grid(row=0, column=1, padx=10, pady=10, sticky="we")
        ttk.Button(
            tab_run_script, 
            text="Choose File", 
            command=self._choose_file
        ).grid(row=0, column=2, padx=10, pady=10)
        
        # Class Name Entry
        ttk.Label(tab_run_script, text="Enter Class Name:").grid(
            row=2, column=0, padx=10, pady=10, sticky="w"
        )
        ttk.Entry(
            tab_run_script, 
            textvariable=self.class_name
        ).grid(row=2, column=1, padx=10, pady=10, sticky="we")
        
        # Run Script Button
        ttk.Button(
            tab_run_script, 
            text="Run Script", 
            command=self._run_script
        ).grid(row=3, column=0, columnspan=3, pady=10)
    
    def _choose_file(self):
        """Open file dialog to choose Excel file."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.excel_file_path.set(file_path)
    
    def _run_script(self):
        """Execute the report card generation script."""
        excel_filename = os.path.basename(self.excel_file_path.get())
        class_name = self.class_name.get()
        
        if not excel_filename or not class_name:
            messagebox.showerror("Error", "Please select Excel file and enter class name")
            return
        
        try:
            generator = ReportCardGenerator(self.base_directory)
            generator.generate_report_cards(excel_filename, class_name)
            messagebox.showinfo("Success", "Report cards generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def quit_word_application():
    """Quit Microsoft Word application if open."""
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Quit()
        print("Microsoft Word application quit successfully.")
    except Exception as e:
        print(f"Error: {e}")

def main():
    """Main entry point for the application."""
    root = tk.Tk()
    ReportCardGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()