import argparse
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from constants import (DEFAULT_CHECK_COMMAND, TESSERACT_DATA_PATH_VAR,
                       VALID_IMAGE_EXTENSIONS, WINDOWS_CHECK_COMMAND)

# Import new modules for LLM and Excel export (optional imports with error handling)
try:
    from llm_converter import text_to_json_with_gemini
    from exporter import export_menu_to_excel
    LLM_AVAILABLE = True
except ImportError as e:
    logging.warning(f"LLM/Export modules not available: {e}")
    LLM_AVAILABLE = False


def create_directory(path):
    """
    Create directory at given path if directory does not exist
    :param path:
    :return:
    """
    if not os.path.exists(path):
        os.makedirs(path)


def check_path(path):
    """
    Check if file path exists or not
    :param path:
    :return: boolean
    """
    return bool(os.path.exists(path))


def get_command():
    """
    Check OS and return command to identify if tesseract is installed or not
    :return:
    """
    if sys.platform.startswith("win"):
        return WINDOWS_CHECK_COMMAND
    return DEFAULT_CHECK_COMMAND


def get_valid_image_files(input_path):
    """
    Efficiently get all valid image files from directory using pathlib
    :param input_path: Directory path to scan
    :return: List of valid image file paths
    """
    valid_files = []
    other_files = 0

    path = Path(input_path)

    # Use pathlib for more efficient file iteration
    for file_path in path.iterdir():
        if file_path.is_file():
            if file_path.suffix.lower() in VALID_IMAGE_EXTENSIONS:
                valid_files.append(file_path)
            else:
                other_files += 1

    return valid_files, other_files


def run_tesseract_optimized(image_path, output_path=None):
    """
    Optimized tesseract runner with better error handling and performance
    :param image_path: Path to image file
    :param output_path: Optional output directory
    :return: Tuple of (success, text_content, filename)
    """
    try:
        filename = image_path.name
        filename_without_extension = image_path.stem

        # If no output path is provided, return text directly
        if not output_path:
            # Use a single temp directory for the entire batch to reduce overhead
            temp_dir = tempfile.mkdtemp()
            temp_file = os.path.join(temp_dir, filename_without_extension)

            result = subprocess.run(
                ["tesseract", str(image_path), temp_file],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=30  # Add timeout to prevent hanging
            )

            if result.returncode == 0:
                try:
                    with open(f"{temp_file}.txt", "r", encoding="utf8") as f:
                        text = f.read()
                    shutil.rmtree(temp_dir)
                    return True, text, filename
                except Exception as e:
                    logging.warning(f"Failed to read output for {filename}: {e}")
                    shutil.rmtree(temp_dir)
                    return False, None, filename
            else:
                logging.warning(
                    f"Tesseract failed for {filename}: {result.stderr.decode()}"
                )
                shutil.rmtree(temp_dir)
                return False, None, filename
        else:
            # Write directly to output directory
            text_file_path = os.path.join(output_path, filename_without_extension)
            result = subprocess.run(
                ["tesseract", str(image_path), text_file_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=30
            )

            if result.returncode == 0:
                return True, None, filename
            else:
                logging.warning(
                    f"Tesseract failed for {filename}: {result.stderr.decode()}"
                )
                return False, None, filename

    except subprocess.TimeoutExpired:
        logging.error(f"Tesseract timeout for {filename}")
        return False, None, filename
    except Exception as e:
        logging.error(f"Unexpected error processing {filename}: {e}")
        return False, None, filename


def run_tesseract(filename, output_path, image_file_name):
    """
    Legacy function for backward compatibility
    """
    image_path = Path(image_file_name)
    success, text, _ = run_tesseract_optimized(image_path, output_path)
    return text if success else ""


def check_pre_requisites_tesseract():
    """
    Check if the pre-requisites required for running the tesseract application are satisfied or not
    :param : NA
    :return: boolean
    """
    check_command = get_command()
    logging.debug("Running `{}` to check if tesseract is installed or not.".format(check_command))

    result = subprocess.run([check_command, "tesseract"], stdout=subprocess.PIPE)
    if not result.stdout:
        logging.error("tesseract-ocr missing, install `tesseract` to resolve. Refer to README for more instructions.")
        return False
    logging.debug("Tesseract correctly installed!\n")

    if sys.platform.startswith("win"):
        environment_variables = os.environ
        logging.debug("Checking if the Tesseract Data path is set correctly or not.\n")
        if TESSERACT_DATA_PATH_VAR in environment_variables:
            if environment_variables[TESSERACT_DATA_PATH_VAR]:
                path = environment_variables[TESSERACT_DATA_PATH_VAR]
                logging.debug(
                    "Checking if the path configured for Tesseract Data Environment variable `{}` \
                as `{}` is valid or not.".format(
                        TESSERACT_DATA_PATH_VAR, path
                    )
                )
                if os.path.isdir(path) and os.access(path, os.R_OK):
                    logging.debug("All set to go!")
                    return True
                else:
                    logging.error("Configured path for Tesseract data is not accessible!")
                    return False
            else:
                logging.error(
                    "Tesseract Data path Environment variable '{}' configured to an empty string!\
                ".format(
                        TESSERACT_DATA_PATH_VAR
                    )
                )
                return False
        else:
            logging.error(
                "Tesseract Data path Environment variable '{}' needs to be configured to point to\
            the tessdata!".format(
                    TESSERACT_DATA_PATH_VAR
                )
            )
            return False
    else:
        return True


def process_images_parallel(image_files, output_path, max_workers=None):
    """
    Process images in parallel using ThreadPoolExecutor
    :param image_files: List of image file paths
    :param output_path: Output directory path
    :param max_workers: Maximum number of worker threads
    :return: Tuple of (successful_files, failed_files, results)
    """
    if max_workers is None:
        # Use number of CPU cores, but cap at 8 to avoid overwhelming tesseract
        max_workers = min(8, os.cpu_count() or 4)

    successful_files = 0
    failed_files = 0
    results = []

    logging.info(f"Processing {len(image_files)} images using {max_workers} parallel workers...")

    start_time = time.time()

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all tasks
        future_to_image = {
            executor.submit(run_tesseract_optimized, image_path, output_path): image_path
            for image_path in image_files
        }

        # Process completed tasks
        for i, future in enumerate(as_completed(future_to_image), 1):
            image_path = future_to_image[future]

            try:
                success, text, filename = future.result()
                if success:
                    successful_files += 1
                    if text:  # Only store text if we're not writing to files
                        results.append((filename, text))
                else:
                    failed_files += 1

            except Exception as e:
                logging.error(f"Error processing {image_path.name}: {e}")
                failed_files += 1

            # Progress indicator
            if i % 10 == 0 or i == len(image_files):
                elapsed = time.time() - start_time
                rate = i / elapsed if elapsed > 0 else 0
                logging.info(
                    f"Progress: {i}/{len(image_files)} "
                    f"({i/len(image_files)*100:.1f}%) - {rate:.1f} files/sec"
                )

    total_time = time.time() - start_time
    logging.info(f"Parallel processing completed in {total_time:.2f} seconds")

    return successful_files, failed_files, results


def validate_and_setup(input_path, output_path):
    """
    Validate prerequisites and setup output directory
    :param input_path: Input path to validate
    :param output_path: Output path to create if needed
    :return: True if validation passes, False otherwise
    """
    # Check if tesseract is installed or not
    if not check_pre_requisites_tesseract():
        return False

    # Check if a valid input directory is given or not
    if not check_path(input_path):
        logging.error("Nothing found at `{}`".format(input_path))
        return False

    # Create output directory
    if output_path:
        create_directory(output_path)
        logging.debug("Creating Output Path {}".format(output_path))

    return True


def process_directory(input_path, output_path, max_workers):
    """
    Process all images in a directory
    :param input_path: Directory containing images
    :param output_path: Output directory for text files
    :param max_workers: Number of parallel workers
    """
    logging.debug("The Input Path is a directory.")

    # Get valid image files efficiently
    image_files, other_files = get_valid_image_files(input_path)

    if len(image_files) == 0:
        logging.error("No valid image files found at your input location")
        logging.error(
            "Supported formats: [{}]".format(", ".join(VALID_IMAGE_EXTENSIONS))
        )
        return

    total_file_count = len(image_files) + other_files
    logging.info(
        "Found total {} file(s) ({} valid images, {} other files)\n".format(
            total_file_count, len(image_files), other_files
        )
    )

    # Process images in parallel
    successful_files, failed_files, results = process_images_parallel(image_files, output_path, max_workers)

    # Print results if not writing to files
    if not output_path:
        for filename, text in results:
            print(f"\n=== {filename} ===")
            print(text)

    # Log final results
    log_processing_results(successful_files, failed_files, other_files)


def process_single_file(input_path, output_path):
    """
    Process a single image file
    :param input_path: Path to the image file
    :param output_path: Output directory for text file
    """
    filename = os.path.basename(input_path)
    logging.debug("The Input Path is a file {}".format(filename))
    image_path = Path(input_path)
    success, text, _ = run_tesseract_optimized(image_path, output_path)
    if success and text:
        print(text)


def log_processing_results(successful_files, failed_files, other_files):
    """
    Log the results of image processing
    :param successful_files: Number of successfully processed files
    :param failed_files: Number of failed files
    :param other_files: Number of non-image files
    """
    logging.info("Parsing Completed!\n")
    logging.info("Successfully parsed images: {}".format(successful_files))
    if failed_files > 0:
        logging.warning("Failed to parse images: {}".format(failed_files))
    if other_files > 0:
        logging.info("Files with unsupported file extensions: {}".format(other_files))


def convert_menu_to_structured_data(
    input_path, 
    output_path, 
    max_workers=None, 
    gemini_model="gemini-2.0-flash-exp",
    export_json=True,
    export_excel=True,
    single_sheet=True
):
    """
    Convert menu images to structured JSON and Excel using OCR + Gemini LLM
    
    :param input_path: Path to input file or directory
    :param output_path: Path to output directory
    :param max_workers: Number of parallel workers for OCR
    :param gemini_model: Gemini model to use
    :param export_json: Whether to export JSON files
    :param export_excel: Whether to export Excel files
    """
    if not LLM_AVAILABLE:
        logging.error("LLM conversion features not available. Please install required dependencies:")
        logging.error("pip install google-generativeai pandas openpyxl")
        return False
    
    # Validate prerequisites and setup
    if not validate_and_setup(input_path, output_path):
        return False
    
    logging.info("🚀 Starting menu conversion with OCR + Gemini LLM...")
    
    # Process images and get OCR results
    if os.path.isdir(input_path):
        ocr_results = process_directory_for_conversion(input_path, max_workers)
    else:
        ocr_results = process_single_file_for_conversion(input_path)
    
    if not ocr_results:
        logging.error("No OCR results to process")
        return False
    
    # Process each OCR result with Gemini
    total_files = len(ocr_results)
    successful_conversions = 0
    failed_conversions = 0
    
    for i, (filename, ocr_text, image_path) in enumerate(ocr_results, 1):
        logging.info(f"Processing {i}/{total_files}: {filename}")
        
        try:
            # Convert OCR text to structured JSON using Gemini
            success, json_data, error = text_to_json_with_gemini(
                ocr_text, 
                str(image_path),
                gemini_model
            )
            
            if success and json_data:
                # Generate output filename base
                base_filename = Path(filename).stem
                
                # Export files
                excel_path = None
                json_path = None
                
                if export_excel or export_json:
                    try:
                        excel_path, json_path = export_menu_to_excel(
                            json_data,
                            output_path,
                            f"menu_{base_filename}",
                            include_json=export_json,
                            include_metadata=True,
                            single_sheet=single_sheet
                        )
                        
                        if export_excel and excel_path:
                            logging.info(f"✅ Excel exported: {excel_path}")
                        if export_json and json_path:
                            logging.info(f"✅ JSON exported: {json_path}")
                        
                        successful_conversions += 1
                        
                    except Exception as export_error:
                        logging.error(f"❌ Export failed for {filename}: {export_error}")
                        failed_conversions += 1
                else:
                    # Just log the JSON structure
                    logging.info(f"✅ Successfully converted {filename} to structured data")
                    successful_conversions += 1
            else:
                logging.error(f"❌ Gemini conversion failed for {filename}: {error}")
                failed_conversions += 1
                
        except Exception as e:
            logging.error(f"❌ Error processing {filename}: {e}")
            failed_conversions += 1
    
    # Log final results
    logging.info("\n📊 Conversion Summary:")
    logging.info(f"Total files processed: {total_files}")
    logging.info(f"Successful conversions: {successful_conversions}")
    if failed_conversions > 0:
        logging.warning(f"Failed conversions: {failed_conversions}")
    
    return successful_conversions > 0


def process_directory_for_conversion(input_path, max_workers):
    """
    Process all images in a directory and return OCR results for conversion
    :param input_path: Directory containing images
    :param max_workers: Number of parallel workers
    :return: List of (filename, ocr_text, image_path) tuples
    """
    # Get valid image files efficiently
    image_files, other_files = get_valid_image_files(input_path)
    
    if len(image_files) == 0:
        logging.error("No valid image files found at your input location")
        return []
    
    logging.info(f"Found {len(image_files)} valid image files")
    
    # Process images in parallel (no output_path = return text directly)
    successful_files, failed_files, results = process_images_parallel(image_files, None, max_workers)
    
    if successful_files == 0:
        logging.error("No images were successfully processed by OCR")
        return []
    
    # Convert results to include full image paths
    ocr_results = []
    for filename, text in results:
        # Find the corresponding image path
        image_path = None
        for img_path in image_files:
            if img_path.name == filename:
                image_path = img_path
                break
        
        if image_path and text.strip():
            ocr_results.append((filename, text, image_path))
    
    return ocr_results


def process_single_file_for_conversion(input_path):
    """
    Process a single image file and return OCR result for conversion
    :param input_path: Path to the image file
    :return: List with single (filename, ocr_text, image_path) tuple
    """
    image_path = Path(input_path)
    success, text, filename = run_tesseract_optimized(image_path, None)
    
    if success and text and text.strip():
        return [(filename, text, image_path)]
    else:
        logging.error(f"Failed to extract text from {filename}")
        return []


def main(input_path, output_path, max_workers=None):
    """
    Main function to process images and extract text using OCR
    :param input_path: Path to input file or directory
    :param output_path: Path to output directory
    :param max_workers: Number of parallel workers
    """
    # Validate prerequisites and setup
    if not validate_and_setup(input_path, output_path):
        return

    # Process based on input type
    if os.path.isdir(input_path):
        process_directory(input_path, output_path, max_workers)
    else:
        process_single_file(input_path, output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Image to Text converter with OCR and AI-powered menu structuring"
    )
    
    # Create subcommands
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # OCR command (original functionality)
    ocr_parser = subparsers.add_parser('ocr', help='Extract text from images using OCR only')
    ocr_parser.add_argument(
        "-i", "--input",
        help="Single image file path or images directory path",
        required=True
    )
    ocr_parser.add_argument("-o", "--output", help="(Optional) Output directory for converted text")
    ocr_parser.add_argument("-d", "--debug", action="store_true", help="Enable verbose DEBUG logging")
    ocr_parser.add_argument(
        "-w", "--workers",
        type=int,
        help="Number of parallel workers (default: auto-detect)",
        default=None
    )
    
    # Convert command (new AI-powered functionality)
    convert_parser = subparsers.add_parser(
        'convert', 
        help='Convert menu images to structured JSON/Excel using OCR + Gemini LLM'
    )
    convert_parser.add_argument(
        "-i", "--input",
        help="Single menu image file path or directory path",
        required=True
    )
    convert_parser.add_argument(
        "-o", "--output", 
        help="Output directory for structured data files (JSON/Excel)",
        required=True
    )
    convert_parser.add_argument("-d", "--debug", action="store_true", help="Enable verbose DEBUG logging")
    convert_parser.add_argument(
        "-w", "--workers",
        type=int,
        help="Number of parallel workers for OCR (default: auto-detect)",
        default=None
    )
    convert_parser.add_argument(
        "--model",
        help="Gemini model to use (default: gemini-2.0-flash-exp)",
        default="gemini-2.0-flash-exp",
        choices=["gemini-2.0-flash-exp", "gemini-2.5-pro", "gemini-1.5-pro"]
    )
    convert_parser.add_argument(
        "--no-json", 
        action="store_true",
        help="Skip JSON file export (Excel only)"
    )
    convert_parser.add_argument(
        "--no-excel", 
        action="store_true",
        help="Skip Excel file export (JSON only)"
    )
    convert_parser.add_argument(
        "--multi-sheet", 
        action="store_true",
        help="Create multi-sheet Excel format (default: single sheet)"
    )

    args = parser.parse_args()

    # Set up logging
    if hasattr(args, 'debug') and args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        logging.getLogger().setLevel(logging.INFO)

    # Check Python version
    if sys.version_info[0] < 3:
        logging.error(
            "You are using Python {0}.{1}. Please use Python>=3".format(
                sys.version_info[0], sys.version_info[1]
            )
        )
        exit()

    # Handle no command (backward compatibility)
    if not args.command:
        # If no subcommand is provided, try to determine if old-style arguments are used
        if len(sys.argv) > 1 and not sys.argv[1] in ['ocr', 'convert']:
            print("⚠️  Warning: Using legacy command format. Consider using 'ocr' command:")
            print("   python main.py ocr -i <input> -o <output>")
            print("   For AI-powered menu conversion, use:")
            print("   python main.py convert -i <input> -o <output>")
            print("")
            
            # Parse as legacy format (for backward compatibility)
            legacy_parser = argparse.ArgumentParser()
            legacy_parser.add_argument("-i", "--input", required=True)
            legacy_parser.add_argument("-o", "--output")
            legacy_parser.add_argument("-d", "--debug", action="store_true")
            legacy_parser.add_argument("-w", "--workers", type=int, default=None)
            
            legacy_args = legacy_parser.parse_args()
            
            input_path = os.path.abspath(legacy_args.input)
            output_path = os.path.abspath(legacy_args.output) if legacy_args.output else None
            
            if legacy_args.debug:
                logging.getLogger().setLevel(logging.DEBUG)
            
            logging.debug("Input Path is {}".format(input_path))
            main(input_path, output_path, legacy_args.workers)
        else:
            parser.print_help()
            exit(0)

    # Process input path
    input_path = os.path.abspath(args.input)
    logging.debug("Input Path is {}".format(input_path))

    # Handle commands
    if args.command == 'ocr':
        # Original OCR functionality
        output_path = os.path.abspath(args.output) if args.output else None
        main(input_path, output_path, args.workers)
        
    elif args.command == 'convert':
        # New AI-powered conversion
        if not LLM_AVAILABLE:
            logging.error("❌ Convert command requires additional dependencies.")
            logging.error("Install them with: pip install -r requirements.txt")
            exit(1)
        
        output_path = os.path.abspath(args.output)
        
        # Check for API key
        api_key = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
        if not api_key:
            logging.error("❌ Gemini API key required for convert command.")
            logging.error("Set environment variable: GOOGLE_API_KEY or GEMINI_API_KEY")
            logging.error("Get your API key from: https://makersuite.google.com/app/apikey")
            exit(1)
        
        success = convert_menu_to_structured_data(
            input_path=input_path,
            output_path=output_path,
            max_workers=args.workers,
            gemini_model=args.model,
            export_json=not args.no_json,
            export_excel=not args.no_excel,
            single_sheet=not args.multi_sheet
        )
        
        if not success:
            exit(1)
