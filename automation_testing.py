import pandas as pd
import requests
import json
import re
import time
import logging
from datetime import datetime
from typing import List, Dict, Tuple
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("automation_testing.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# API Configuration
BASE_URL = "https://aiapidev.pandai.org"
API_ENDPOINT = f"{BASE_URL}/grader/marking-scheme/evaluate"

"""# PREPROCESSING FUNCTIONS"""

def is_list_format(text: str) -> tuple:
    """
    Check if the text is in list format.

    Args:
        text (str): The text to analyze

    Returns:
        tuple: (is_list_format: bool, max_sentences: float)
    """
    lines = text.strip().splitlines()
    list_like = 0

    for line in lines:
        line = line.strip()
        if re.match(r"^[-*•]\s", line) or re.match(r"^\d+[\.\)]\s", line):
            list_like += 1
        elif len(line.split()) <= 8 and not has_connecting_words(line):
            list_like += 1

    # threshold of minimum kata hubung in a sentences
    items = [item.strip() for item in lines[0].split(".")]
    max_sentences = 0.25 * len(items)

    if len(lines) == 1:
        total_words = len(lines[0].split())
        total_sentences = len(re.findall(r"[.!?]", lines[0]))
        if total_sentences <= 1 and total_words < 30:
            max_sentences = 5

    if len(lines) == 1 and "," in lines[0]:
        short_items = [item for item in items if len(item.split()) <= 8]
        if len(short_items) >= max(2, len(items) // 2) and not has_connecting_words(lines[0], max_sentences):
            return True, max_sentences

    return list_like >= max(1, len(lines) // 2), max_sentences


def has_connecting_words(text: str, max_sentences: float = 5) -> bool:
    """
    Check if the text has connecting words.

    Args:
        text (str): The text to analyze
        max_sentences (float): Maximum sentence threshold

    Returns:
        bool: True if text has connecting words above threshold
    """
    KATA_HUBUNG = [
        "seterusnya", "berikutnya", "karena", "oleh kerana", "akhirnya", "contohnya", "demi",
        "agar", "bahawa", "untuk", "supaya", "kesimpulan", "yang terakhir", "oleh itu", "contohnya",
        "di samping itu", "juga adalah", "sekaligus", "selain itu", "akhir sekali", "namun", "selanjutnya,",
        "tambahan pula", "pendapat saya"
    ]

    text_lower = text.lower()
    flags = [word in text_lower for word in KATA_HUBUNG]
    return sum(flags) > max_sentences


def classify_answer(text: str) -> dict:
    """
    Classify the answer type based on format and connections.

    Args:
        text (str): The answer text to classify

    Returns:
        dict: Classification result with type, note, and weight
    """
    is_list, max_sentences = is_list_format(text)
    has_connections = has_connecting_words(text, max_sentences)

    if is_list:
        return {
            "type": "list",
            "note": "Jawaban berbentuk poin-poin",
            "weight": 0.5
        }

    if has_connections:
        return {
            "type": "essay connected",
            "style": "connected",
            "note": "Essay dengan kalimat yang terhubung",
            "weight": 1
        }

    return {
        "type": "essay disconnected",
        "style": "disconnected",
        "note": "Essay berupa poin-poin tanpa koneksi antar kalimat",
        "weight": 0.5
    }


def process_exam_schema(schema_text):
    """
    Process exam schema to split at codes F, H, and C.

    Args:
        schema_text (str): The exam schema text

    Returns:
        list: List of processed criteria
    """
    # Pattern for F, H, C codes and Roman numerals in parentheses
    pattern = r'([FHC]\d+[a-z]?|\([ivxlcdm]+\))'

    # Find all occurrences of the pattern
    code_positions = [(m.group(0), m.start()) for m in re.finditer(pattern, schema_text)]

    # Process the schema into separate criteria
    criteria = []
    for i, (code, pos) in enumerate(code_positions):
        end_pos = code_positions[i+1][1] if i < len(code_positions)-1 else len(schema_text)
        criterion_text = schema_text[pos:end_pos].strip()
        criteria.append(f'{criterion_text}')

    return criteria


def preprocess_user_input(user_prompt):
    """
    Preprocess user input to extract and format components.

    Args:
        user_prompt (str): The raw user input

    Returns:
        tuple: (formatted_user_string, weight, max_score)
    """
    text = user_prompt.replace("\n", " ")
    text = text.strip('"')

    pattern = r"(student_answer|exam_schema|max_score|keyword):\s*(.*?)(?=\s+(?:student_answer|exam_schema|max_score|keyword):|$)"
    matches = re.findall(pattern, text)
    result = {key: value.strip() for key, value in matches}

    # process the exam schema to split at codes F, H, and C
    exam_schema = result['exam_schema']
    processed_schema = "; ".join([schema for schema in process_exam_schema(exam_schema)])
    if len(processed_schema) < 1:
        processed_schema = exam_schema

    classified = classify_answer(result['student_answer'])
    logger.info(f"Classification: {classified}")
    answer_type = classified['type']
    weight = classified['weight']
    if result.get("keyword", "").lower() == "senaraikan":
        weight = 1
    answer_type_note = classified['note']

    user_string = (f"student_answer_input: {result['student_answer']}; "
        f"answer_type: {answer_type}; "
        f"weight: {weight}; "
        f"max_score: {result['max_score']}; "
        f"exam_scheme_input: {processed_schema}; ")

    return user_string, weight, int(result['max_score'])

"""# FILE READING FUNCTIONS"""

def read_excel_file(file_path: str, sheet_name: str = None) -> List[List[str]]:
    """
    Read Excel file and return data as list of lists.

    Args:
        file_path (str): Path to Excel file
        sheet_name (str): Sheet name to read (optional)

    Returns:
        List[List[str]]: Data from Excel file
    """
    try:
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Convert to list of lists including headers
        data = [df.columns.tolist()] + df.astype(str).values.tolist()
        logger.info(f"Successfully read Excel file: {file_path}")
        return data
    except Exception as e:
        logger.error(f"Error reading Excel file {file_path}: {e}")
        return []


def read_csv_file(file_path: str) -> List[List[str]]:
    """
    Read CSV file and return data as list of lists.

    Args:
        file_path (str): Path to CSV file

    Returns:
        List[List[str]]: Data from CSV file
    """
    try:
        df = pd.read_csv(file_path)
        # Convert to list of lists including headers
        data = [df.columns.tolist()] + df.astype(str).values.tolist()
        logger.info(f"Successfully read CSV file: {file_path}")
        return data
    except Exception as e:
        logger.error(f"Error reading CSV file {file_path}: {e}")
        return []


def auto_detect_and_read_file(file_path: str) -> List[List[str]]:
    """
    Auto-detect file type and read accordingly.

    Args:
        file_path (str): Path to file

    Returns:
        List[List[str]]: Data from file
    """
    file_path = Path(file_path)

    if not file_path.exists():
        logger.error(f"File not found: {file_path}")
        return []

    if file_path.suffix.lower() in ['.xlsx', '.xls']:
        return read_excel_file(str(file_path))
    elif file_path.suffix.lower() == '.csv':
        return read_csv_file(str(file_path))
    else:
        logger.error(f"Unsupported file format: {file_path.suffix}")
        return []

"""# DATA PROCESSING FUNCTIONS"""

def process_exam_data(data: List[List[str]]) -> List[Dict[str, str]]:
    """
    Convert exam data to structured format.

    Args:
        data: Raw data from file

    Returns:
        List of dictionaries containing processed exam data
    """
    if not data or len(data) < 2:
        logger.error("Data tidak mencukupi atau kosong")
        return []

    headers = data[0]
    logger.info(f"📋 Headers found: {headers}")

    # Expected column indices (adjust based on your file structure)
    col_mapping = {
        'bahagian': 0,           # Bahagian
        'no_soalan': 1,          # No. Soalan
        'text_soalan': 2,        # Text Soalan
        'markah_penuh': 3,       # Markah Penuh
        'jawapan_pelajar': 4,    # Jawapan Pelajar
        'markah_pelajar': 5,     # Markah Pelajar
        'kod_skema': 6,          # Kod skema
        'skema_jawapan': 7       # Skema Jawapan
    }

    processed_data = []
    current_question = None

    for i, row in enumerate(data[1:], 1):  # Skip header
        # Skip if row doesn't have enough columns
        if len(row) < 8:
            continue

        # Handle NaN values and convert to string
        row = [str(cell).strip() if str(cell) != 'nan' else '' for cell in row]

        bahagian = row[col_mapping['bahagian']]
        no_soalan = row[col_mapping['no_soalan']]
        text_soalan = row[col_mapping['text_soalan']]
        markah_penuh = row[col_mapping['markah_penuh']]
        jawapan_pelajar = row[col_mapping['jawapan_pelajar']]
        markah_pelajar = row[col_mapping['markah_pelajar']]
        kod_skema = row[col_mapping['kod_skema']]
        skema_jawapan = row[col_mapping['skema_jawapan']]

        # Skip empty rows or rows without student answers
        if not jawapan_pelajar:
            continue

        # Determine if this is a new question or continuation
        if no_soalan:  # New question
            current_question = {
                'row_number': i,
                'bahagian': bahagian,
                'no_soalan': no_soalan,
                'text_soalan': text_soalan,
                'markah_penuh': markah_penuh,
                'jawapan_pelajar': jawapan_pelajar,
                'markah_pelajar': markah_pelajar,
                'kod_skema': kod_skema,
                'skema_jawapan': skema_jawapan
            }
            processed_data.append(current_question)
        else:  # Continuation of previous question
            if current_question:
                additional_answer = {
                    'row_number': i,
                    'bahagian': current_question['bahagian'],
                    'no_soalan': current_question['no_soalan'],
                    'text_soalan': current_question['text_soalan'],
                    'markah_penuh': current_question['markah_penuh'],
                    'jawapan_pelajar': jawapan_pelajar,
                    'markah_pelajar': markah_pelajar,
                    'kod_skema': current_question['kod_skema'],
                    'skema_jawapan': current_question['skema_jawapan']
                }
                processed_data.append(additional_answer)

    return processed_data


def create_user_messages(processed_data: List[Dict[str, str]]) -> List[Dict]:
    """
    Create user_message strings from processed exam data.

    Args:
        processed_data: List of processed exam records

    Returns:
        List of formatted user_message dictionaries
    """
    user_messages = []

    for i, record in enumerate(processed_data, 1):
        # Clean and validate data
        student_answer = record['jawapan_pelajar'].strip()
        max_score = record['markah_penuh'].strip()
        exam_schema = record['skema_jawapan'].strip()
        question_text = record['text_soalan'].strip()

        # Use kod_skema if exam_schema is empty
        if not exam_schema and record['kod_skema']:
            exam_schema = record['kod_skema'].strip()

        if not exam_schema:
            logger.warning(f"⚠️  Warning: No exam schema found for question {record['no_soalan']}")
            continue

        # Format user_message
        keywords = ['senaraikan', 'namakan', 'nyatakan']
        lower_question = question_text.lower()

        matched_keyword = next((kw for kw in keywords if kw in lower_question), None)

        if matched_keyword:
            user_message = f"student_answer: {student_answer} exam_schema: {exam_schema} max_score: {max_score} keyword: senaraikan"
        else:
            user_message = f"student_answer: {student_answer} exam_schema: {exam_schema} max_score: {max_score}"

        user_messages.append({
            'question_info': f"{record['bahagian']}{record['no_soalan']} (Row {record['row_number']})",
            'user_message': user_message,
            'original_marks': record['markah_pelajar'],
            'original_matched_scheme': record['kod_skema'],
            'question_text': record['text_soalan']
        })

    return user_messages

"""# API FUNCTIONS"""

def call_marking_api(payload: Dict) -> Dict:
    """
    Call the marking scheme API endpoint.

    Args:
        payload (Dict): API payload

    Returns:
        Dict: API response
    """
    try:
        response = requests.post(API_ENDPOINT, json=payload, timeout=30)
        response.raise_for_status()
        result = response.json()
        logger.info(f"API call successful")
        return result
    except requests.exceptions.RequestException as e:
        logger.error(f"API error: {e}")
        return {"error": str(e)}
    except json.JSONDecodeError as e:
        logger.error(f"JSON decode error: {e}")
        return {"error": f"Invalid JSON response: {e}"}


def post_process_marks(response, weight, max_score):
    """
    Post-process marks based on response and weight.

    Args:
        response: The response object from API
        weight (float): The weight factor for scoring
        max_score (int): Maximum possible score

    Returns:
        dict: Processed result with total marks
    """
    if "error" in response:
        return response

    try:
        # Handle different response structures
        if isinstance(response, dict) and 'data' in response:
            result_data = response['data']
        else:
            result_data = response

        answers = result_data.get('answers', [])

        unique_criteria = set()
        for ans in answers:
            if ans.get('correct') is True:
                matched = ans.get('matched_criteria', [])
                unique_criteria.update(matched)

        logger.info(f"Unique criteria matched: {unique_criteria}")
        total_marks = round(len(unique_criteria) * weight, 2)

        if total_marks > max_score:
            total_marks = max_score

        result_data["total_answer_marks"] = total_marks
        return result_data

    except Exception as e:
        logger.error(f"Error in post_process_marks: {e}")
        return {"error": str(e)}

"""# PROCESSING FUNCTIONS"""

def process_test_data(test_data: List[Dict]) -> List[Dict]:
    """
    Process test data and call API for each item.

    Args:
        test_data: List of test data items

    Returns:
        List of results
    """
    results = []

    for i, data in enumerate(test_data, 1):
        logger.info(f"Processing item {i}/{len(test_data)}: {data['question_info']}")

        question_info = data['question_info']
        user_message_data = data['user_message']
        original_marks = data['original_marks']
        original_matched_scheme = data['original_matched_scheme']
        question_text = data['question_text']

        try:
            # Preprocess the user input
            user_msg, weight, max_score = preprocess_user_input(user_message_data)
            logger.info(f'After preprocessing: {user_msg}')

            # Create API payload
            payload = {
                "student_answer": user_msg.split("student_answer_input: ")[1].split(";")[0].strip(),
                "exam_schema": user_msg.split("exam_scheme_input: ")[1].split(";")[0].strip(),
                "max_score": max_score,
            }

            # Add keyword if present
            if "keyword" in user_message_data:
                payload["keyword"] = "senaraikan"

            # Call API
            response = call_marking_api(payload)

            if "error" not in response:
                # Post-process the response
                processed_response = post_process_marks(response, weight, max_score)

                logger.info("============================================")
                logger.info(f'Input: {user_message_data}')
                logger.info("============================================")
                logger.info(f'Output: {json.dumps(processed_response, indent=2, ensure_ascii=False)}')
                logger.info("\n")

                answers = processed_response.get("answers", [])
                student_answers = [a.get("value", "") for a in answers]
                matched_criteria_list = ["\n".join(a.get("matched_criteria", [])) for a in answers]

                result_row = {
                    "No Soalan": question_info,
                    "Question Text": question_text,
                    "Student Answer": "\n".join(student_answers),
                    "Weight": weight,
                    "Markah AI": processed_response.get("total_answer_marks", 0),
                    "Original Markah Pelajar": original_marks,
                    "Markah Penuh": max_score,
                    "Kod Skema AI": "\n".join(matched_criteria_list),
                    "Original Kod Skema": original_matched_scheme,
                    "Exam Scheme": "\n".join([x.get("text", "") for x in processed_response.get("exam_scheme", [])])
                }

                results.append(result_row)
            else:
                logger.error(f"API error for {question_info}: {response['error']}")
                # Add error result
                result_row = {
                    "No Soalan": question_info,
                    "Question Text": question_text,
                    "Student Answer": "ERROR",
                    "Weight": weight,
                    "Markah AI": 0,
                    "Original Markah Pelajar": original_marks,
                    "Markah Penuh": max_score,
                    "Kod Skema AI": f"ERROR: {response['error']}",
                    "Original Kod Skema": original_matched_scheme,
                    "Exam Scheme": "ERROR"
                }
                results.append(result_row)

        except Exception as e:
            logger.error(f"Error processing data {question_info}: {e}")
            # Add error result
            result_row = {
                "No Soalan": question_info,
                "Question Text": question_text,
                "Student Answer": "ERROR",
                "Weight": 0,
                "Markah AI": 0,
                "Original Markah Pelajar": original_marks,
                "Markah Penuh": max_score,
                "Kod Skema AI": f"ERROR: {str(e)}",
                "Original Kod Skema": original_matched_scheme,
                "Exam Scheme": "ERROR"
            }
            results.append(result_row)
            continue

        # Small delay to avoid overwhelming the API
        time.sleep(0.5)

    return results

"""# EXPORT FUNCTIONS"""

def export_results(results: List[Dict], filename: str):
    """
    Export results to Excel file with proper formatting.

    Args:
        results: List of result dictionaries
        filename: Output filename
    """
    try:
        df = pd.DataFrame(results)

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Hasil")

            # Get worksheet and set wrap text for all cells
            workbook = writer.book
            worksheet = writer.sheets["Hasil"]

            for column_cells in worksheet.columns:
                for cell in column_cells:
                    cell.alignment = cell.alignment.copy(wrap_text=True)

        logger.info(f"Results exported to: {filename}")

    except Exception as e:
        logger.error(f"Error exporting results: {e}")

"""# MAIN FUNCTIONS"""

def main(input_file_path: str, output_filename: str = None):
    """
    Main function to run the automation testing.

    Args:
        input_file_path (str): Path to input file (Excel or CSV)
        output_filename (str): Optional output filename
    """
    logger.info("Starting local automation testing")

    # Generate timestamp for output filename
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if not output_filename:
        output_filename = f"Testing_Results_{timestamp}.xlsx"

    # Read input file
    logger.info(f"Reading input file: {input_file_path}")
    data = auto_detect_and_read_file(input_file_path)

    if not data:
        logger.error("Failed to read input file")
        return

    # Process exam data
    logger.info("Processing exam data")
    processed_data = process_exam_data(data)

    if not processed_data:
        logger.error("No valid exam data found")
        return

    # Create user messages
    logger.info("Creating user messages")
    user_messages = create_user_messages(processed_data)

    if not user_messages:
        logger.error("No valid user messages created")
        return

    logger.info(f"Found {len(user_messages)} items to process")

    # Process test data
    logger.info("Processing test data with API calls")
    results = process_test_data(user_messages)

    # Export results
    logger.info("Exporting results")
    export_results(results, output_filename)

    logger.info(f"Automation testing completed. Results saved to: {output_filename}")


if __name__ == "__main__":
    # Testing
    input_file = r"C:\Users\khali\Downloads\automation_testing\B3 2025 OKR - SJ Jawapan Pelajar (Radin).xlsx"

    # Run the automation testing
    main(input_file)