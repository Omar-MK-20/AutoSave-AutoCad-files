import os
import datetime
import shutil
import win32com.client
import logging
import sys
import traceback
from pathlib import Path

# Configure logging to both file and console
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('autocad_autosave.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Configuration
BACKUP_DIR = os.path.join(os.path.expanduser("~"), "AutoCAD_Backups")
DELETE_AFTER_DAYS = 1

def create_backup_directory():
    """Create backup directory if it doesn't exist"""
    try:
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
            logging.info(f"Created backup directory: {BACKUP_DIR}")
        else:
            logging.info(f"Backup directory already exists: {BACKUP_DIR}")
    except Exception as e:
        logging.error(f"Failed to create backup directory: {str(e)}")
        logging.error(traceback.format_exc())
        raise

def get_open_autocad_files():
    """Get list of open AutoCAD files"""
    try:
        logging.info("Attempting to connect to AutoCAD...")
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        logging.info("Successfully connected to AutoCAD")
        
        documents = acad.Documents
        open_files = []
        
        logging.info(f"Number of open documents: {documents.Count}")
        
        for i in range(documents.Count):
            try:
                doc = documents.Item(i)
                logging.debug(f"Processing document {i+1}")
                
                # Try to get the file path using different methods
                file_path = None
                
                # Method 1: Try to get the file path directly from the document
                try:
                    # Get the document's database
                    db = doc.Database
                    # Get the file path from the database
                    file_path = db.FileName
                    logging.debug(f"Method 1 - Database FileName: {file_path}")
                except:
                    logging.debug("Method 1 - Database FileName not available")
                
                # Method 2: Try to get the file path from the document's properties
                if not file_path:
                    try:
                        # Get the document's properties
                        props = doc.GetInterfaceObject("AcDbDocument")
                        # Get the file path from the properties
                        file_path = props.GetVariable("DWGPREFIX") + props.GetVariable("DWGNAME")
                        logging.debug(f"Method 2 - Document properties: {file_path}")
                    except:
                        logging.debug("Method 2 - Document properties not available")
                
                # Method 3: Try to get the file path from the document's name and path
                if not file_path:
                    try:
                        name = doc.Name
                        path = doc.Path
                        if name and path:
                            file_path = os.path.normpath(os.path.join(path, name))
                            logging.debug(f"Method 3 - Name and Path: {file_path}")
                    except:
                        logging.debug("Method 3 - Name and Path not available")
                
                # Method 4: Try to get the file path from the document's full name
                if not file_path:
                    try:
                        file_path = doc.FullName
                        logging.debug(f"Method 4 - FullName: {file_path}")
                    except:
                        logging.debug("Method 4 - FullName not available")
                
                # Verify the file exists
                if file_path and os.path.exists(file_path):
                    open_files.append(file_path)
                    logging.info(f"Found open file: {file_path}")
                else:
                    logging.warning(f"Could not determine valid file path for document {i+1}")
                    logging.warning(f"Attempted path: {file_path}")
                    
                    # Try to find the file in common locations
                    if doc.Name:
                        common_paths = [
                            os.path.expanduser("~/Desktop"),
                            os.path.expanduser("~/Documents"),
                            os.path.expanduser("~")
                        ]
                        for common_path in common_paths:
                            try:
                                potential_path = os.path.normpath(os.path.join(common_path, doc.Name))
                                if os.path.exists(potential_path):
                                    open_files.append(potential_path)
                                    logging.info(f"Found file in common location: {potential_path}")
                                    break
                            except Exception as e:
                                logging.debug(f"Error checking common path {common_path}: {str(e)}")
            except Exception as e:
                logging.error(f"Error processing document {i+1}: {str(e)}")
                logging.error(traceback.format_exc())
        
        if not open_files:
            logging.warning("No open AutoCAD files found")
        
        return open_files
    except Exception as e:
        logging.error(f"Error getting open AutoCAD files: {str(e)}")
        logging.error(traceback.format_exc())
        return []

def backup_file(file_path):
    """Create a backup of the file"""
    try:
        if not os.path.exists(file_path):
            logging.warning(f"File does not exist: {file_path}")
            return
        
        file_name = os.path.basename(file_path)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{os.path.splitext(file_name)[0]}_{timestamp}.bak"
        backup_path = os.path.join(BACKUP_DIR, backup_name)
        
        logging.info(f"Attempting to backup: {file_path}")
        logging.info(f"Backup destination: {backup_path}")
        
        shutil.copy2(file_path, backup_path)
        logging.info(f"Successfully backed up: {file_path} to {backup_path}")
    except Exception as e:
        logging.error(f"Error backing up {file_path}: {str(e)}")
        logging.error(traceback.format_exc())

def cleanup_old_backups():
    """Delete backups older than DELETE_AFTER_DAYS"""
    try:
        current_time = datetime.datetime.now()
        backup_files = os.listdir(BACKUP_DIR)
        logging.info(f"Found {len(backup_files)} backup files to check")
        
        for file_name in backup_files:
            try:
                file_path = os.path.join(BACKUP_DIR, file_name)
                file_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))
                age_days = (current_time - file_time).days
                
                if age_days >= DELETE_AFTER_DAYS:
                    logging.info(f"Deleting old backup: {file_path} (age: {age_days} days)")
                    os.remove(file_path)
                    logging.info(f"Successfully deleted: {file_path}")
            except Exception as e:
                logging.error(f"Error processing backup file {file_name}: {str(e)}")
                logging.error(traceback.format_exc())
    except Exception as e:
        logging.error(f"Error cleaning up old backups: {str(e)}")
        logging.error(traceback.format_exc())

def main():
    logging.info("Starting AutoCAD autosave script")
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Current working directory: {os.getcwd()}")
    
    try:
        create_backup_directory()
        
        open_files = get_open_autocad_files()
        if not open_files:
            logging.warning("No files to backup - is AutoCAD running?")
            return
        
        for file_path in open_files:
            backup_file(file_path)
        
        cleanup_old_backups()
        logging.info("Script completed successfully")
    except Exception as e:
        logging.error(f"Main execution error: {str(e)}")
        logging.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main() 