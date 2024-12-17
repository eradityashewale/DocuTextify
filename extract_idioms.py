import pandas as pd
from docx import Document
import re

# Path to your Word file
file_path = 'Idioms - 31 ( 27 June ).docx'

# Open and read the Word file
doc = Document(file_path)

# Initialize the list to store idioms and their definitions
idioms_definitions = []

# Join all paragraphs to process all text as one block
full_text = ' '.join(para.text for para in doc.paragraphs)

# Regular expression to match idiom name and definition
# Modified to capture only the definition without extra text (e.g., "Example...")
idioms = re.findall(r'(\d+)\.\s*([A-Za-z\s’‘]+?)\s*[-–—\u2013]\s*([^\n]+?)(?=\s*(?:Example|$))', full_text)

# Add found idioms and their definitions to the list
for idiom in idioms:
    number, name, definition = idiom
    idioms_definitions.append({'idiom': name.strip(), 'definition': definition.strip()})

# Create a DataFrame from the scraped idioms and definitions
df = pd.DataFrame(idioms_definitions)

# Display the DataFrame
print(df)
