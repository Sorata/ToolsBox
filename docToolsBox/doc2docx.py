import os
import sys
import time
import threading
import queue
import logging
import traceback
import win32com.client
import pythoncom

# Configuration
# ----------------
# Supported extensions to convert
SOURCE_EXT = ".doc"
TARGET_EXT = ".docx"
# Word File Format Constant (wdFormatXMLDocument = 12)
WD_FORMAT_XML_DOCUMENT = 12

# Log files
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ERROR_LOG_FILE = os.path.join(SCRIPT_DIR, "conversion_errors.log")
PROCESSED_LOG_FILE = os.path.join(SCRIPT_DIR, "processed_files.txt")

# Number of concurrent threads (adjust based on CPU/Memory)
# Warning: Too many threads might overwhelm the system or Word instances.
# 4 is a reasonable default for desktop automation.
NUM_THREADS = 4

# Setup Logging
# ----------------
# Console Logger
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)

# File Logger for Errors
file_handler = logging.FileHandler(ERROR_LOG_FILE, encoding='utf-8')
file_handler.setLevel(logging.ERROR)
file_formatter = logging.Formatter('%(asctime)s - %(message)s\n%(exc_info)s\n')
file_handler.setFormatter(file_formatter)

logger = logging.getLogger("DocConverter")
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# Global State
# ----------------
processed_files = set()
processed_lock = threading.Lock()

def load_processed_files():
    """Load the list of already processed files to support resuming."""
    if os.path.exists(PROCESSED_LOG_FILE):
        try:
            with open(PROCESSED_LOG_FILE, 'r', encoding='utf-8') as f:
                for line in f:
                    processed_files.add(line.strip())
            logger.info(f"Loaded {len(processed_files)} processed files from history.")
        except Exception as e:
            logger.error(f"Failed to load processed log: {e}")

def mark_as_processed(file_path):
    """Mark a file as processed and append to the log file immediately."""
    with processed_lock:
        processed_files.add(file_path)
        try:
            with open(PROCESSED_LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(file_path + "\n")
        except Exception as e:
            logger.error(f"Failed to write to processed log: {e}")

def worker(q):
    """Thread worker function that maintains a persistent Word instance."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    word_app = None
    
    try:
        # Create Word instance once per thread
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
        except Exception as e:
            logger.error("Failed to initialize Word Application in thread.")
            return

        while True:
            try:
                file_path = q.get(timeout=3) # Wait for a file, or exit if empty for a while (handled by sentinel usually, but timeout helps)
            except queue.Empty:
                # If queue is empty, we check if we are done? 
                # Better: use a sentinel value (None) to stop.
                continue
            
            if file_path is None:
                q.task_done()
                break

            try:
                process_file(word_app, file_path)
            except Exception as e:
                logger.error(f"Error processing file {file_path}")
            finally:
                q.task_done()

    except Exception as e:
        logger.error(f"Thread crashed: {e}")
    finally:
        # Cleanup Word instance
        if word_app:
            try:
                word_app.Quit()
            except:
                pass
        pythoncom.CoUninitialize()

def process_file(word_app, doc_path):
    """Convert a single file using the provided Word instance."""
    docx_path = doc_path + "x" # .doc -> .docx
    
    # Check if target already exists (optional, but good for safety)
    # But user wants "resume", which is handled by processed_files list.
    # If processed_files missed it but file exists, we might overwrite or skip.
    # Let's just overwrite to be safe.
    
    doc = None
    try:
        logger.info(f"Processing: {doc_path}")
        doc = word_app.Documents.Open(doc_path)
        doc.SaveAs(docx_path, FileFormat=WD_FORMAT_XML_DOCUMENT)
        doc.Close()
        doc = None
        
        # Delete original file
        os.remove(doc_path)
        
        mark_as_processed(doc_path)
        logger.info(f"Converted and removed: {doc_path}")
        
    except Exception as e:
        # Log the full error
        logger.error(f"Failed to convert {doc_path}", exc_info=True)
        # Attempt to close doc if open
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        raise e

def main():
    load_processed_files()
    
    # Find all .doc files
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
        logger.info("No new .doc files found to process.")
        return

    logger.info(f"Found {total_files} files to process.")

    # Setup Queue and Threads
    task_queue = queue.Queue()
    for f in files_to_process:
        task_queue.put(f)
    
    # Add sentinels to stop threads
    for _ in range(NUM_THREADS):
        task_queue.put(None)

    logger.info(f"Starting {NUM_THREADS} worker threads...")
    threads = []
    for _ in range(NUM_THREADS):
        t = threading.Thread(target=worker, args=(task_queue,))
        t.start()
        threads.append(t)

    # Wait for completion
    task_queue.join()
    
    # Wait for threads to finish cleanup
    for t in threads:
        t.join()

    logger.info("Batch conversion completed.")

if __name__ == "__main__":
    main()
