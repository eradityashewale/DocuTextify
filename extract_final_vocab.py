import mammoth
import re
import pandas as pd
import os

def extract_vocabulary(file_path):
    # Step 1: Extract raw text from the Word file
    with open(file_path, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        text = result.value

    # Normalize text to handle inconsistent spacing
    text = re.sub(r"Examples\s*[-\u2013]?\s*", "Examples: ", text)  # Normalize "Examples"

    # Step 2: Split text into individual entries
    entries = re.split(r"(?<!\w)\d+\.\s*", text)  # Split by patterns like "1.", "2.", etc.

    vocab_list = []  # To store the extracted vocabulary as dictionaries

    # Step 3: Process each entry to extract details
    for i, entry in enumerate(entries):
        if not entry.strip():  # Skip empty entries
            continue

        print(f"Processing Entry {i}: {entry[:50]}...")  # Debugging output

        lines = entry.strip().split("\n")  # Split entry into lines
        word_line = lines[0]  # First line should contain the word and meaning

        # Updated regex to handle missing spaces and optional translations
        word_match = re.match(r"^(\w+)\s\((.*?)\)\s[-\u2013]\s(.+?)(?:\s*\(.*?\))?$", word_line)
        if word_match:
            vocab_name = word_match.group(1).strip()
            vocab_type = word_match.group(2).strip()
            vocab_meaning = word_match.group(3).strip()
        else:
            print(f"Skipping Entry {i}: Failed to match word format.\nRaw Entry: {word_line}")
            continue

        # Initialize placeholders for other details
        examples = []
        synonyms = []
        hint = ""

        current_section = None
        for line in lines[1:]:
            line = line.strip()

            if line.startswith("Examples:"):
                current_section = "examples"
                continue
            elif line.startswith("Synonyms"):
                current_section = "synonyms"
                synonyms = line.replace("Synonyms -", "").strip().split(", ")
            elif line.startswith("Hint"):
                current_section = "hint"
                hint = re.sub(r"^Hint\s[-\u2013]\s*", "", line).strip()
            elif current_section == "examples":
                if line and not line.startswith(("Synonyms", "Hint")):
                    examples.append(line)

        # Fallback for missing examples
        if not examples:
            examples = ["No example provided."]

        # Add structured data to the list
        vocab_list.append({
            "name": vocab_name,
            "type": vocab_type,
            "meaning": vocab_meaning,
            "examples": " | ".join(examples),  # Join examples into a single string
            "synonyms": ", ".join(synonyms),
            "hint": hint
        })

    return vocab_list

def save_to_excel(vocabulary, output_file):
    # Convert the vocabulary list to a DataFrame
    df = pd.DataFrame(vocabulary)

    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False, sheet_name="Vocabulary")
    print(f"Excel file created at: {output_file}")

# Example usage
file_path = "Vocab - 186 with photos.docx"
vocabulary = extract_vocabulary(file_path)

# Define output file path
output_file = os.path.join(os.getcwd(), "Extracted_Vocabulary.xlsx")

# Save the vocabulary to an Excel file
save_to_excel(vocabulary, output_file)
