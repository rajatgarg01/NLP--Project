import requests
from bs4 import BeautifulSoup
import pandas as pd
import spacy
import syllapy
from textblob import TextBlob
import os
import time


# Directory containing files
folder_path = 'D:\OneDrive\Desktop\html\env'  # Update this path


# Load the Excel file
file_path = 'D:\OneDrive\Desktop\html\env\Input.xlsx'  # Replace with your actual file path
#data = pd.read_excel(file_path)

# Loop through each row in the Excel file
 for index, row in data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    # Send a GET request to the URL
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extract title
    title_tag = soup.find('h1', class_='entry-title')
    title = title_tag.get_text() if title_tag else 'Title not found'

    # Extract date

    date_tag = soup.find('time', class_='entry-date')
    date = date_tag.get_text() if date_tag else 'Date not found'

    # Extract content
    content_tag = soup.find('div', class_='td-post-content')
    content = content_tag.get_text() if content_tag else 'Content not found'

    # Combine all the extracted information
    extracted_data = f"Title: {title}\nDate: {date}\nContent: {content}"

    # Define the file path and name for the notepad file
    notepad_path = f"{url_id}.txt"
    
    # Save the extracted data to a notepad file
    with open(notepad_path, 'w', encoding='utf-8') as file:
        file.write(extracted_data)
    
    print(f"Data for {url_id} has been saved to {notepad_path}")

print("All data has been successfully extracted and saved.")






# Load spacy model
nlp = spacy.load('en_core_web_sm')

# Define functions to calculate the required variables

def calculate_positive_negative_scores(text, positive_words, negative_words):
    # Split the text into words and count positive and negative words
    words = text.split()
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    return positive_score, negative_score

def calculate_polarity_subjectivity_scores(text):
    blob = TextBlob(text)
    polarity = blob.sentiment.polarity
    subjectivity = blob.sentiment.subjectivity
    return polarity, subjectivity

def calculate_avg_sentence_length(text):
    doc = nlp(text)
    sentences = list(doc.sents)
    if len(sentences) == 0:
        return 0
    total_words = sum(len(sentence) for sentence in sentences)
    return total_words / len(sentences)

def calculate_percentage_of_complex_words(text):
    words = text.split()
    complex_word_count = sum(1 for word in words if syllapy.count(word) > 2)
    return (complex_word_count / len(words)) * 100 if words else 0

def calculate_fog_index(avg_sentence_length, percentage_complex_words):
    return 0.4 * (avg_sentence_length + percentage_complex_words)

def calculate_complex_word_count(text):
    words = text.split()
    return sum(1 for word in words if syllapy.count(word) > 2)

def calculate_word_count(text):
    return len(text.split())

def calculate_syllable_per_word(text):
    words = text.split()
    if not words:
        return 0
    syllables = sum(syllapy.count(word) for word in words)
    return syllables / len(words)

def calculate_personal_pronouns(text):
    doc = nlp(text)
    personal_pronouns = [token.text for token in doc if token.pos_ == 'PRON' and token.tag_ in ['PRP', 'PRP$']]
    return len(personal_pronouns)

def calculate_avg_word_length(text):
    words = text.split()
    if not words:
        return 0
    return sum(len(word) for word in words) / len(words)

# Load positive and negative word lists
positive_words = ["good", "happy", "excellent", "positive", "fortunate", "correct", "superior"]
negative_words = ["bad", "sad", "terrible", "negative", "unfortunate", "wrong", "inferior"]



# Initialize an empty list to store data
data = []

# Iterate through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        file_path = os.path.join(folder_path, filename)
        
        # Read the file content
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
        
        # Calculate the variables
        positive_score, negative_score = calculate_positive_negative_scores(text, positive_words, negative_words)
        polarity_score, subjectivity_score = calculate_polarity_subjectivity_scores(text)
        avg_sentence_length = calculate_avg_sentence_length(text)
        percentage_complex_words = calculate_percentage_of_complex_words(text)
        fog_index = calculate_fog_index(avg_sentence_length, percentage_complex_words)
        complex_word_count = calculate_complex_word_count(text)
        word_count = calculate_word_count(text)
        syllable_per_word = calculate_syllable_per_word(text)
        personal_pronouns = calculate_personal_pronouns(text)
        avg_word_length = calculate_avg_word_length(text)
        
        # Append results to the list
        data.append({
            "URL_ID": filename.split(".")[0],
            "Positive Score": positive_score,
            "Negative Score": negative_score,
            "Polarity Score": polarity_score,
            "Subjectivity Score": subjectivity_score,
            "Avg Sentence Length": avg_sentence_length,
            "Percentage of Complex Words": percentage_complex_words,
            "Fog Index": fog_index,
            "Complex Word Count": complex_word_count,
            "Word Count": word_count,
            "Syllable Per Word": syllable_per_word,
            "Personal Pronouns": personal_pronouns,
            "Avg Word Length": avg_word_length
        })

# Convert the list of dictionaries to a pandas DataFrame
df = pd.DataFrame(data)
df.set_index('URL_ID', inplace=True)
df.to_excel('analysis.xlsx')


df_base = pd.read_excel('Output Data Structure.xlsx')
df_base = df_base[['URL_ID', 'URL']]

df_base.set_index('URL_ID', inplace=True)

time.sleep(3)

df_new = pd.read_excel('analysis.xlsx')

df_new.set_index('URL_ID', inplace=True)

# Merge the DataFrames on their index
merged_df = pd.merge(df_base, df_new, left_index=True, right_index=True, how='left')

merged_df.to_excel('Final Output.xlsx')
