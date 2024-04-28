# OST Placement Search Private Limited Assignment

# CV Data Extraction Program
## Overview
This program is designed to extract email IDs, contact numbers, and overall text from CVs (resumes) and store the information in an Excel file (.xls). It provides a convenient solution for parsing multiple CVs simultaneously and organizing the extracted data in a structured format.

## Features
- Supports both PDF and DOCX formats for CV files
- Extracts email IDs and contact numbers using regular expressions
- Stores the extracted information in an Excel file with separate columns for email IDs, contact numbers, and the overall text from the CV
- Utilizes popular libraries such as PyPDF2, textract, pandas, and docx

## Instructions
1. Clone this repository to your local machine.
    ```bash
    https://github.com/atulguptag/OST-Placement-Search-Private-Limited-Assignment
    ```

2. Ensure you have Python installed.
    ```bash
    Python 3.10 or above installed on your system
    ```

3. Install the required dependencies by running:
   ```bash
   pip install -r requirements.txt
   ```
4. Place your CV files (in PDF or DOCX format) inside the "Sample2" directory.

5. Run the script `task.py`:
   ```bash
   python task.py
   ```
6. Once the script completes execution, you will find the extracted data saved in an Excel file named `output.xls`.

## Sample Output
| File Name          | Emails                      | Phone Numbers      | Overall Text |
|--------------------|-----------------------------|--------------------|--------------|
| AarushiRohatgi.pdf | aarushi.9999218543@gmail.com | 999-921-8543       | [Overall Text from AarushiRohatgi.pdf] |
| AkashGoel.docx     | akashg2494@gmail.com        | 9310631244         | [Overall Text from AkashGoel.docx]     |
| AkashSharma.pdf    | akashsharma1894@gmail.com   | 8072458559         | [Overall Text from AkashSharma.pdf]    |
| ...                | ...                         | ...                | ...          |

## Requirements
- Python 3.x
- PyPDF2
- textract
- pandas
- docx

## Note
- Ensure that your CV files contain accurate and accessible contact information for proper extraction.
- In case of any issues or errors, please feel free to reach out for assistance.

## License
[MIT License](LICENSE)