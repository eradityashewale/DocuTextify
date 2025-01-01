import mammoth
import re
import pandas as pd

def extract_raw_text(file_path):
    """
    Extracts raw text from a .docx file using Mammoth.
    """
    with open(file_path, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        text = result.value
    return text

def parse_quizzes(text, max_options=6):
    """
    Parses the raw text to extract quizzes, options, and their descriptions.

    Args:
        text (str): The raw text extracted from the Word document.
        max_options (int): Maximum number of options expected per quiz.

    Returns:
        list of dict: A list where each dict represents a quiz with options and descriptions.
    """
    lines = text.splitlines()
    quizzes = []
    current_quiz = {}
    option_labels = ['A', 'B', 'C', 'D']  # Extend if needed

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # Identify the start of a quiz
        if line.startswith('Quiz -'):
            # If there's an existing quiz being processed, save it
            if current_quiz:
                quizzes.append(current_quiz)
                current_quiz = {}

            # Initialize a new quiz
            quiz_question = line
            current_quiz['Quiz'] = quiz_question
            current_quiz['Options'] = {}
            current_quiz['Descriptions'] = {}
            i += 1

            # Extract options
            option_count = 0
            while i < len(lines) and option_count < max_options:
                option_line = lines[i].strip()
                if option_line and not option_line.startswith(('Examples –', 'Synonyms -', 'Hint -', 'Quiz -')):
                    label = option_labels[option_count] if option_count < len(option_labels) else f"Option_{option_count+1}"
                    current_quiz['Options'][label] = option_line
                    option_count += 1
                elif not option_line:
                    # Skip empty lines
                    pass
                else:
                    # Encountered a new section, stop collecting options
                    break
                i += 1

            # After options, extract descriptions
            description_pattern = re.compile(r'^(.+?)\s*\((.+?)\)\s*–\s*(.+)')
            while i < len(lines):
                desc_line = lines[i].strip()
                if desc_line.startswith('Quiz -'):
                    break  # Next quiz starts
                desc_match = description_pattern.match(desc_line)
                if desc_match:
                    option_text = desc_match.group(1).strip()
                    part_of_speech = desc_match.group(2).strip()
                    description = desc_match.group(3).strip()
                    # Find which option this description belongs to
                    matched_label = None
                    for label, opt in current_quiz['Options'].items():
                        if opt.lower() == option_text.lower():
                            matched_label = label
                            break
                    if matched_label:
                        current_quiz['Descriptions'][matched_label] = f"{option_text} ({part_of_speech}) – {description}"
                i += 1
        else:
            i += 1

    # Append the last quiz if exists
    if current_quiz:
        quizzes.append(current_quiz)

    return quizzes

def create_dataframe(quizzes, max_options=6):
    """
    Creates a Pandas DataFrame from the list of quizzes.

    Args:
        quizzes (list of dict): The parsed quizzes.
        max_options (int): Maximum number of options across all quizzes.

    Returns:
        pd.DataFrame: The structured DataFrame.
    """
    # Define DataFrame columns
    option_labels = ['A', 'B', 'C', 'D'][:max_options]
    columns = ['Quiz']
    for label in option_labels:
        columns.extend([f"Option_{label}", f"Option_{label}_description"])

    # Collect row data
    rows = []
    for quiz in quizzes:
        row = {'Quiz': quiz.get('Quiz', '')}
        for idx, label in enumerate(option_labels):
            option = quiz['Options'].get(label, '')
            description = quiz['Descriptions'].get(label, '')
            row[f"Option_{label}"] = option
            row[f"Option_{label}_description"] = description
        rows.append(row)

    # Create DataFrame
    df = pd.DataFrame(rows, columns=columns)
    return df

# Main Execution
if __name__ == "__main__":
    # Specify the path to your Word document
    file_path = "Vocab - 62 with photos.docx"  # Update with your actual file path

    # Step 1: Extract raw text
    raw_text = extract_raw_text(file_path)

    # Step 2: Parse quizzes
    quizzes = parse_quizzes(raw_text, max_options=6)  # Adjust max_options if needed

    # Step 3: Create DataFrame
    df_quizzes = create_dataframe(quizzes, max_options=6)  # Adjust max_options if needed

    # Display the DataFrame
    print(df_quizzes)

    # Optional: Export to CSV and Excel
    df_quizzes.to_csv("extracted_quizzes.csv", index=False)
