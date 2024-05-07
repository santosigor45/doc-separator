# Doc Separator
This project provides a tool to process PDF files and separate pages based on city information linked to employee names. It uses the PyMuPDF (fitz) library for handling PDFs, fuzzywuzzy for fuzzy string matching, and tkinter for a graphical user interface.

# Features
- Extract text from specified regions in PDF files.
- Identify and sort PDF pages based on the city (and neighbour region, if applicable) information associated with employee names.
- Generate new PDF files organized.
- Simple GUI for selecting files and output directories.

# Installation
To run this project, you will need Python installed on your system. You can then install the required dependencies via pip:

    pip install -r requirements.txt

# Usage
- **Start the Application:** Run the script to open the GUI.

        python script_name.py

- **Select PDF File:** Use the 'Select File' button to choose the PDF you want to process.

- **Select Output Directory:** Choose where the sorted PDFs will be saved.

- **Process PDF:** Click 'Process PDF' to start processing. The application will notify you when processing is complete and if any names were not found.

- **Review Results:** The output PDFs will be organized in the selected directory.

# Configuration
Modify the pdfregion tuple in the process_pdf function to change the text extraction region according to your PDF layout. Also, ensure the funcionarios_cidade.txt file is correctly formatted with each line containing an employee name and city, separated by a comma.

# Contributing
Contributions are welcome! Please fork the repository and submit a pull request with your features or fixes.

# License
This project is licensed under the MIT License - see the LICENSE file for details.

# Contact

For any queries or further assistance, please contact santosigor45@gmail.com.
