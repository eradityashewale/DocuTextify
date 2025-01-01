import pandas as pd
from docx import Document
import re
import os

# ---------------------------- User Configurations ----------------------------

# Path to your Word file containing new idioms
new_word_file_path = 'Idioms - 60 ( 27 June ).docx'  # Update this path as needed

# Path to the existing CSV file with idioms (if it exists)
existing_csv_path = 'idioms_definitions.csv'  # Update this to your existing CSV file path

# Starting number for new idioms
# If you want to continue numbering from the existing CSV, set this to None
# Otherwise, set it to your desired starting number (e.g., 500)
specified_start_number = None  # Change to an integer if you want to specify a start number

# ---------------------------- End of Configurations ----------------------------

# Function to determine the starting number
# def determine_start_number(existing_csv, specified_start):
#     if existing_csv and os.path.isfile(existing_csv):
#         df_existing = pd.read_csv(existing_csv)
#         if not df_existing.empty and 'number' in df_existing.columns:
#             last_number = df_existing['number'].max()
#             return last_number + 1
#         else:
#             print("Existing CSV is empty or missing 'number' column. Starting from 1.")
#             return 1
#     elif specified_start is not None:
#         return specified_start
#     else:
#         return 1  # Default starting number if no existing CSV and no specified start

# Determine the starting number
start_number = 600

# Open and read the new Word file
doc = Document(new_word_file_path)

# Initialize the list to store idioms and their definitions
idioms_definitions = []

# Join all paragraphs to process all text as one block
full_text = '\n'.join(para.text for para in doc.paragraphs)

# Debugging: Print the full text to verify
print("Full Text:\n", full_text)

# Updated regex pattern with VERBOSE flag and greedy match for definition
pattern = re.compile(r'''
    (\d+)\.\s*                                # Group 1: Number followed by a dot
    ([A-Za-z\s’‘]+?)\s*[-–—\u2013]\s*         # Group 2: Idiom Name followed by a dash
    ([^\n]+)\s*                               # Group 3: Definition (greedy)
    (?:\n\s*Example\s*[-–—\u2013]*\s*([^\n]+))?\s*  # Group 4: Optional Example
    \n\s*Quiz\s*[-–—\u2013]*\s*([^\n]+)\s*    # Group 5: Quiz question
    \n\s*([^\n]+)\s*                          # Group 6: Option a
    \n\s*([^\n]+)\s*                          # Group 7: Option b
    \n\s*([^\n]+)\s*                          # Group 8: Option c
    \n\s*([^\n]+)                             # Group 9: Option d
''', re.VERBOSE | re.DOTALL)

# Extract all matches using the updated pattern
matches = pattern.findall(full_text)

# Debugging: Print the matches to verify
print("\nMatches:\n", matches)

# Loop through the matches and structure the idioms
for match in matches:
    if len(match) != 9:
        print(f"Warning: Unexpected number of groups in match: {match}")
        continue  # Skip this match if it doesn't have all required groups
    number_doc, name, definition, example, quiz, a, b, c, d = match
    idioms_definitions.append({
        'idiom': name.strip(),
        'definition': definition.strip(),
        'example': example.strip() if example else 'N/A',  # Use 'N/A' if no example is found
        'quiz': quiz.strip() if quiz else 'N/A',          # Use 'N/A' if no quiz is found
        'option_a': a.strip() if a else 'N/A',
        'option_b': b.strip() if b else 'N/A',
        'option_c': c.strip() if c else 'N/A',
        'option_d': d.strip() if d else 'N/A'
    })

# Create a DataFrame from the idioms
new_df = pd.DataFrame(idioms_definitions)

# Assign a 'number' column starting from the determined start_number
new_df.insert(0, 'number', range(start_number, start_number + len(new_df)))

# Debugging: Print the new DataFrame content
print("\nNew DataFrame:\n", new_df)

# Check if the existing CSV exists
if os.path.isfile(existing_csv_path):
    try:
        # Read the existing CSV into a DataFrame
        existing_df = pd.read_csv(existing_csv_path)
        print("\nExisting DataFrame Loaded Successfully.")
        
        # Ensure the existing DataFrame has the required columns
        required_columns = ['number', 'idiom', 'definition', 'example', 'quiz',
                            'option_a', 'option_b', 'option_c', 'option_d']
        if not all(column in existing_df.columns for column in required_columns):
            print(f"Error: Existing CSV does not contain all required columns: {required_columns}")
            print("Please ensure the existing CSV has the correct format.")
            exit(1)
        
        # Concatenate the existing DataFrame with the new DataFrame
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        print("\nNew Data has been appended to the existing DataFrame.")
    except Exception as e:
        print(f"Error reading the existing CSV file: {e}")
        exit(1)
else:
    print("\nNo existing CSV file found. A new CSV file will be created.")
    combined_df = new_df

# Debugging: Print the combined DataFrame content
print("\nCombined DataFrame:\n", combined_df)

# Save the combined DataFrame back to the existing CSV file
try:
    combined_df.to_csv(existing_csv_path, index=False)
    print(f"\nData successfully appended. CSV file saved as '{existing_csv_path}'.")
except Exception as e:
    print(f"Error saving the combined CSV file: {e}")
    exit(1)
