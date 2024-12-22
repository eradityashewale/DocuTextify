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
full_text = '\n'.join(para.text for para in doc.paragraphs)

# Debugging: Print the full text to verify
print("Full Text:\n", full_text)

# This pattern is to capture each idiom with its number, name, definition, Example and Quiz
pattern = r'(\d+)\.\s*([A-Za-z\s’‘]+?)\s*[-–—\u2013]\s*([^\n]+?)(?=\s*(?:Example|Quiz|$))\s*(?:Example\s*[-–—\u2013]*\s*([^\n]+))?\s*(?:Quiz\s*[-–—\u2013]*\s*([^\n]+))?'

# Extract all matches using the pattern
matches = re.findall(pattern, full_text)

# Debugging: Print the matches to verify
print("\nMatches:\n", matches)

# Loop through the matches and structure the idioms
for match in matches:
    number, name, definition, example, quiz = match
    idioms_definitions.append({
        'idiom': name.strip(),
        'definition': definition.strip(),
        'example': example.strip() if example else 'N/A',  # Use 'N/A' if no example is found
        'quiz': quiz.strip() if quiz else 'N/A'  # Use 'N/A' if no quiz is found
    })

# Create a DataFrame from the idioms
df = pd.DataFrame(idioms_definitions)

# Debugging: Print the final DataFrame content
print("\nDataFrame:\n", df)

# Save the DataFrame to a CSV file
csv_file_path = 'idioms_definitions.csv'
df.to_csv(csv_file_path, index=False)

print(f"CSV file saved as {csv_file_path}")
