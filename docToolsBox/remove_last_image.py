import os
import sys
import logging
import time
from concurrent.futures import ThreadPoolExecutor
from docx import Document
from docx.oxml.ns import qn

# Configuration
SOURCE_EXT = ".docx"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ERROR_LOG_FILE = os.path.join(SCRIPT_DIR, "image_removal_errors.log")
PROCESSED_LOG_FILE = os.path.join(SCRIPT_DIR, "processed_image_removal.txt")
NUM_THREADS = 4

# Setup Logging
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)

file_handler = logging.FileHandler(ERROR_LOG_FILE, encoding='utf-8')
file_handler.setLevel(logging.ERROR)
file_formatter = logging.Formatter('%(asctime)s - %(message)s\n%(exc_info)s\n')
file_handler.setFormatter(file_formatter)

logger = logging.getLogger("ImageRemover")
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# Global State
processed_files = set()

def load_processed_files():
    if os.path.exists(PROCESSED_LOG_FILE):
        try:
            with open(PROCESSED_LOG_FILE, 'r', encoding='utf-8') as f:
                for line in f:
                    processed_files.add(line.strip())
            logger.info(f"Loaded {len(processed_files)} processed files from history.")
        except Exception as e:
            logger.error(f"Failed to load processed log: {e}")

def mark_as_processed(file_path):
    processed_files.add(file_path)
    try:
        with open(PROCESSED_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(file_path + "\n")
    except Exception as e:
        logger.error(f"Failed to write to processed log: {e}")

def get_all_paragraphs_in_order(parent):
    """
    Recursively yield all paragraphs in the document element in order.
    Includes paragraphs in tables.
    """
    for child in parent:
        if child.tag.endswith('p'):
            yield child
        elif child.tag.endswith('tbl'):
            for row in child.findall('.//w:tr', namespaces=child.nsmap):
                for cell in row.findall('.//w:tc', namespaces=row.nsmap):
                    yield from get_all_paragraphs_in_order(cell)

def process_file(file_path):
    try:
        logger.info(f"Processing: {file_path}")
        doc = Document(file_path)
        removed = False
        
        # We need to examine the document content in reverse order.
        # To do this correctly including tables, we first linearize all paragraphs.
        # Accessing the internal XML element 'body'
        
        body = doc.element.body
        all_paragraphs = list(get_all_paragraphs_in_order(body))
        
        # Iterate backwards
        for i in range(len(all_paragraphs) - 1, -1, -1):
            p_element = all_paragraphs[i]
            
            # Get all runs in this paragraph
            runs = [child for child in p_element if child.tag.endswith('r')]
            
            paragraph_has_text = False
            paragraph_has_image = False
            image_is_last_in_para = False
            run_to_remove = None
            
            # Iterate runs in reverse to check content
            for r_idx in range(len(runs) - 1, -1, -1):
                run = runs[r_idx]
                
                # Iterate children of run in reverse
                # We need to list them because we might modify or just to iterate backwards
                run_children = list(run)
                
                for child in reversed(run_children):
                    if child.tag.endswith('t'):
                        if child.text and child.text.strip():
                            paragraph_has_text = True
                            break
                    elif child.tag.endswith('drawing') or child.tag.endswith('pict'):
                        paragraph_has_image = True
                        image_is_last_in_para = True
                        # We found the image. 
                        # Since we are iterating in reverse and haven't hit text yet, this is the last image.
                        # And since it's the first meaningful thing we hit, it's at the end.
                        
                        # Remove this specific child from the run
                        run.remove(child)
                        removed = True
                        break
                
                if paragraph_has_text:
                    # Found text after potential image or just text at end of doc
                    break
                
                if removed:
                    logger.info(f"Removed last image from: {file_path}")
                    break
            
            if paragraph_has_text:
                logger.info(f"Document ends with text, skipping: {file_path}")
                break
            
            if removed:
                break
                
            # If paragraph is empty (no text, no image), continue to previous paragraph
        
        if removed:
            doc.save(file_path)
        else:
            logger.info(f"No image removed from: {file_path}")
            
        mark_as_processed(file_path)

    except Exception as e:
        logger.error(f"Failed to process {file_path}", exc_info=True)

def main():
    load_processed_files()
    target_dir = os.getcwd()
    
    files_to_process = []
    logger.info(f"Scanning directory: {target_dir} ...")
    for root, dirs, files in os.walk(target_dir):
        for name in files:
            if name.lower().endswith(SOURCE_EXT) and not name.startswith("~$"):
                full_path = os.path.abspath(os.path.join(root, name))
                if full_path not in processed_files:
                    files_to_process.append(full_path)
    
    total_files = len(files_to_process)
    if total_files == 0:
        logger.info("No new .docx files found to process.")
        return

    logger.info(f"Found {total_files} files to process.")
    logger.info(f"Starting processing with {NUM_THREADS} threads using python-docx...")

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        executor.map(process_file, files_to_process)

if __name__ == "__main__":
    main()
