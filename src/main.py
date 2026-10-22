# VTU Result Automation - Main Script

import time
import os
from config import WAIT_AFTER_STARTUP, WAIT_BEFORE_CAPTCHA
from gui import AutomationGUI
from scraper import ResultScraper
from file_io import load_usns_from_file, export_results
from excel_io import (
    load_input_workbook,
    read_usn,
    create_output_workbook,
    write_headers,
    write_student_info,
    write_subject_data,
    save_workbook
)
from logger import setup_logger
import logging

logger = setup_logger(__name__)


def process_results(gui, inputs):
    """
    Main automation workflow - processes student results.
    Supports input files in Excel (.xlsx, .xls), CSV, Word (.docx), PDF, or TXT.
    Supports output files in Excel (.xlsx, .xls), CSV, Word (.docx), or PDF.
    
    Args:
        gui: AutomationGUI instance for user interaction
        inputs: Dictionary containing user inputs from GUI
    """
    driver_path = inputs["driver_path"]
    usn_file = inputs["usn_file"]
    website = inputs["website"]
    save_path = inputs["save_path"]
    auto_solve_captcha = inputs.get("auto_solve_captcha", True)
    
    scraper = None

    if not usn_file or not os.path.exists(usn_file):
        gui.show_error("Please select a valid input USN file.")
        return

    if not save_path:
        gui.show_error("Please select an output save path.")
        return

    # Load USNs from file (supports Excel, CSV, Word DOCX, PDF, TXT)
    usns = load_usns_from_file(usn_file)
    in_sheet = None
    if not usns:
        # Fallback to legacy workbook loader
        in_book, in_sheet = load_input_workbook(usn_file)
        if not in_book:
            gui.show_error(f"Failed to find or load USNs from {usn_file}. Check logs for details.")
            return

    # Parse and validate row range
    start_str = str(inputs.get("start_row", "")).strip()
    end_str = str(inputs.get("end_row", "")).strip()

    total_usns = len(usns) if usns else 100
    try:
        start_row = int(start_str) if start_str else 1
    except ValueError:
        gui.show_error("Start row must be a valid number.")
        return

    try:
        end_row = int(end_str) if end_str else total_usns
    except ValueError:
        gui.show_error("End row must be a valid number.")
        return

    if start_row < 1:
        start_row = 1
    if usns and end_row > len(usns):
        end_row = len(usns)
    if start_row > end_row:
        gui.show_error(f"Start row ({start_row}) cannot be greater than End row ({end_row}).")
        return

    logger.info(f"Processing USNs from index {start_row} to {end_row} (Auto-captcha: {auto_solve_captcha})")
    
    # Create legacy output workbook for backward compatibility
    out_book, out_sheet, orange_style = create_output_workbook()
    write_headers(out_sheet)
    
    time.sleep(WAIT_AFTER_STARTUP)
    
    # Initialize web scraper
    scraper = ResultScraper(driver_path, website)
    try:
        scraper.setup_driver()
    except Exception as e:
        logger.critical(f"Failed to setup driver: {e}")
        gui.show_error(f"Failed to setup driver: {e}")
        return
    
    # Locate page elements
    if not scraper.locate_page_elements():
        logger.error("Failed to locate page elements.")
        gui.show_error("Page elements not found. Website layout may have changed.")
        scraper.cleanup()
        return
    
    manual_captcha = None
    if not auto_solve_captcha:
        time.sleep(WAIT_BEFORE_CAPTCHA)
        manual_captcha = gui.get_captcha_input()
        if not manual_captcha:
            logger.warning("Captcha input cancelled by user.")
            gui.show_warning("Captcha input cancelled. Stopping.")
            scraper.cleanup()
            return
    
    # Process each USN
    success_count = 0
    error_count = 0
    results_data = []
    
    for row_idx in range(start_row, end_row + 1):
        if usns:
            usn = usns[row_idx - 1]
        elif in_sheet:
            usn = read_usn(in_sheet, row_idx)
        else:
            usn = None
        
        if not usn:
            continue
        
        if auto_solve_captcha:
            # Automatically solve captcha and handle retries
            opened = scraper.submit_with_auto_captcha(usn)
        else:
            # Manual entry path
            scraper.enter_usn_and_captcha(usn, manual_captcha)
            opened = scraper.submit_and_switch_to_result()
        
        if not opened:
            logger.warning(f"Could not open result window for USN: {usn}")
            error_count += 1
            continue
        
        try:
            # Extract student information
            page_usn, name = scraper.scrape_student_info()
            if not page_usn or not name:
                error_count += 1
                scraper.close_result_and_return_to_main()
                continue
            
            semester = scraper.scrape_semester() if hasattr(scraper, "scrape_semester") else ""
            
            # Extract all subject details
            subjects = scraper.scrape_subjects()
            
            # Save complete full-card screenshot of the result (never cut off)
            try:
                scraper.capture_full_result_screenshot(f"result_{page_usn}.png")
            except Exception:
                pass

            # Record structured student result
            results_data.append({
                "usn": page_usn,
                "name": name,
                "semester": semester,
                "subjects": subjects
            })
            
            # Write to legacy excel sheet
            write_student_info(out_sheet, row_idx, page_usn, name)
            write_subject_data(out_sheet, row_idx, subjects, orange_style)
            
            success_count += 1
            
        except Exception as e:
            logger.error(f"Error processing {usn}: {e}")
            error_count += 1
        
        # Close result window and return to main page
        scraper.close_result_and_return_to_main()
    
    # Save results in the user's requested format (.xlsx, .csv, .docx, .pdf, or .xls)
    try:
        ext = os.path.splitext(save_path)[1].lower()
        if ext == ".xls":
            save_workbook(out_book, save_path)
        else:
            export_results(results_data, save_path)
            
        gui.show_info(f"Processing complete.\nProcessed: {success_count}\nErrors: {error_count}\nSaved to: {save_path}")
    except Exception as e:
        logger.error(f"Failed to save results: {e}")
        gui.show_error(f"Failed to save results: {e}")
    finally:
        if scraper:
            scraper.cleanup()


def main():
    """Application entry point - creates GUI and starts the workflow"""
    gui = None
    
    def on_submit():
        if gui:
            process_results(gui, gui.get_inputs())

    gui = AutomationGUI(on_submit_callback=on_submit)
    
    try:
        logger.info("Starting application GUI.")
        gui.run()
    except Exception as e:
        logger.critical(f"Unhandled exception: {e}", exc_info=True)
        try:
            gui.show_error(f"Fatal Error: {e}")
        except:
            pass
        raise


if __name__ == "__main__":
    main()
