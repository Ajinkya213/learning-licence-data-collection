# Data Collection from XLSX spreadsheet

## Introduction
This project aims to collect data from Learning Licence question bank [PDF](https://parivahan.gov.in/parivahan/sites/default/files/DownloadForm/STALL_QB_ENGLISH_NEW.pdf), which contain multiple-choice questions (MCQs). Converted the data pdf file to xlsx(spreadsheet) format for convenience. This helped to develop the project faster. Used [iLovePDF](https://www.ilovepdf.com/) to convert the PDF file to XLSX format.

## Project Overview
**Data Collection**: XLSX scraping techniques will be employed to extract MCQs from the Learning Licence question bank XLSX. The extracted data will be processed and structured for further use. 

## Technologies Used
- **Python**: Utilized for XLSX scraping, data processing. [link](https://www.python.org/)
- **OpenPyXL**: Python library for extracting text and metadata from excel sheet. [link](https://openpyxl.readthedocs.io/en/stable/)

## Project Structure
- **`main.py`**: Python script for scraping data from Xls(spreadsheet) and processing it.
- **`mcq.py`**: Class with attributes question, answer, options, image and serial_no.
- **`question-bank.pdf`**: Main data file, contains all the MCQs.
- **`question-bank.xlsx`**: Data file to scrape data from.
- **`/resource`**: Scraped images directory (JPEG format).
- **`mcq_data.json`**: Scraped data JSON file

## Usage
**Data Collection**: Run the `main.py` script to extract MCQs from the Learning Licence question bank spreadsheet.

## Contributors
[Ajinkya Bhushan](https://github.com/ajinkya213)

## Acknowledgements
- Special thanks to [Parivahan Sewa](https://parivahan.gov.in/parivahan/) for providing the question-bank in public domain.