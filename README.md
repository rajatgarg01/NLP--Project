Textual Data Analysis Script
Overview
This script automates the extraction, analysis, and scoring of textual content from multiple URLs. It uses NLP techniques to perform sentiment analysis, readability scoring, and word complexity evaluations. It also generates structured Excel reports for the analyzed data.

Features
Data Extraction:

Scrapes content (title, date, and text) from URLs provided in an Excel file.
Saves extracted content into individual .txt files.
Text Analysis:

Positive/Negative Word Counts: Identifies words based on predefined positive and negative word lists.
Polarity & Subjectivity: Sentiment analysis using TextBlob.
Readability Metrics: Calculates average sentence length, percentage of complex words, and the Fog Index.
Word Metrics: Computes word count, syllables per word, personal pronoun count, and average word length.
Output:

Processes results into an intermediate analysis.xlsx file.
Merges results with the original input structure into a final report Final Output.xlsx.
Dependencies
Install the required libraries using pip:

bash
pip install requests beautifulsoup4 pandas spacy syllapy textblob openpyxl
Spacy: Ensure the English model is installed:
bash
python -m spacy download en_core_web_sm
Folder Structure
Ensure your working directory contains the following files:

Input Files:

Input.xlsx: Contains the URL_ID and URL fields.
Output Data Structure.xlsx: Reference for final report structure.
Output Directory:

Extracted text files saved as <URL_ID>.txt.
Generated Excel outputs: analysis.xlsx and Final Output.xlsx.
How to Run
Update the folder_path and file_path variables with appropriate file paths.
Execute the script:
bash
Copy code
python texual-data-analysis.py
Output files (Final Output.xlsx) will be generated in the same directory.
File Outputs
Text Files: Extracted content from URLs.
analysis.xlsx: Intermediate analysis results.
Final Output.xlsx: Comprehensive analysis report combining input and output data.
Notes
Modify positive_words and negative_words lists to customize sentiment analysis.
Ensure internet access to scrape URL content.
Author
Rajat garg

Email: rgargr150@gmail.com
