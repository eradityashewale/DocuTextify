import mammoth
import re
import pandas as pd

# Function to extract vocabulary data
def extract_quiz_data(text):
    quiz_data = []

    # Split the text into individual quizzes using the Quiz keyword
    quizzes = re.split(r'\n\s*Quiz\s*-\s*', text)
    
    # Iterate through quizzes and extract details
    for quiz in quizzes[1:]:  # Skip the first split as it is before the first Quiz
        # Extract question (first sentence before options)
        question_match = re.search(r'^(.*?)\n', quiz)
        question = question_match.group(1).strip() if question_match else ''

        # Extract options
        options = re.findall(r'\n([A-Z][a-z]*)', quiz)
        options += [''] * (4 - len(options))  # Ensure 4 options are present

        # Extract the answer
        answer_match = re.search(r'\n\s*(.*?)\s*\(\s*', quiz)
        answer = answer_match.group(1).strip() if answer_match else ''

        # Extract descriptions for options
        descriptions = re.findall(r'\n([A-Z][a-z]*).*?((\([a-z]\))\s*–\s*.*?)\n', quiz)
        description_dict = {desc[0]: desc[1] for desc in descriptions}
        description_a = description_dict.get(options[0], '')
        description_b = description_dict.get(options[1], '')
        description_c = description_dict.get(options[2], '')
        description_d = description_dict.get(options[3], '')

        # Append the quiz details to the list
        quiz_data.append({
            'Question': question,
            'Option A': options[0],
            'Option B': options[1],
            'Option C': options[2],
            'Option D': options[3],
            'Answer': answer,
            'Option A Desc': description_a,
            'Option B Desc': description_b,
            'Option C Desc': description_c,
            'Option D Desc': description_d
        })

    # Convert to a pandas DataFrame
    quiz_df = pd.DataFrame(quiz_data)
    return quiz_df

# Function to extract raw text from a Word file
def extract_text_from_word(file_path):
    with open(file_path, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        return result.value

# File path to the Word document
file_path = "Vocab - 183 with photos.docx"  # Replace with the actual file path

# Extract text from the Word file
text = extract_text_from_word(file_path)

# Extract quiz data from the text
quiz_df = extract_quiz_data(text)

# Save to Excel
quiz_df.to_excel("quiz_data.xlsx", index=False)

