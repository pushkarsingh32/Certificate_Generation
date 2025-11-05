"""
Certificate Generation System

This module generates internship completion certificates by overlaying participant
information (name and dates) onto a certificate template. It supports batch processing
from Excel files and outputs both PNG and PDF formats.

Author: Certificate Generation System
Date: 2025
"""

import logging
import os
import sys
from pathlib import Path
from typing import Tuple, Optional

import cv2
import img2pdf
import numpy as np
import pandas as pd

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('certificate_generation.log')
    ]
)
logger = logging.getLogger(__name__)

# ============================================================================
# CONFIGURATION CONSTANTS
# ============================================================================

# File paths
PARTICIPANTS_FILE = "certificate_prospects/health_interns.xlsx"
TEMPLATE_FILE = "templates/template1.png"
OUTPUT_DIRECTORY = "generated_certificates/"

# Font configuration (BGR format for OpenCV)
FONT_COLOR = (68, 102, 253)  # Professional blue color
FONT_FACE = cv2.FONT_HERSHEY_SIMPLEX
FONT_THICKNESS = 10

# Text positioning adjustments (in pixels)
class TextPosition:
    """Text positioning constants for certificate generation."""
    INTERN_NAME_X = 100
    INTERN_NAME_Y = 78
    INTERN_NAME_FONT_SIZE = 3.8

    DATE_FROM_X = 130
    DATE_Y = -80
    DATE_TO_X = 535
    DATE_FONT_SIZE = 1.5

# Certificate naming
CERTIFICATE_PREFIX = "Internship_Completion_Certificate"


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def validate_file_exists(file_path: str) -> bool:
    """
    Validate that a required file exists.

    Args:
        file_path: Path to the file to validate

    Returns:
        bool: True if file exists, False otherwise

    Raises:
        FileNotFoundError: If the file does not exist
    """
    if not os.path.exists(file_path):
        logger.error(f"Required file not found: {file_path}")
        raise FileNotFoundError(f"Required file not found: {file_path}")
    return True


def ensure_output_directory(directory: str) -> None:
    """
    Ensure output directory exists, create if it doesn't.

    Args:
        directory: Path to the output directory
    """
    Path(directory).mkdir(parents=True, exist_ok=True)
    logger.info(f"Output directory ready: {directory}")


def draw_text_on_image(
    img: np.ndarray,
    text: str,
    font: int,
    font_size: float,
    x_offset: int,
    y_offset: int,
    color: Tuple[int, int, int] = FONT_COLOR,
    thickness: int = FONT_THICKNESS
) -> None:
    """
    Draw text on an image at a calculated centered position with offsets.

    This function centers the text on the image and then applies x and y offsets
    to position it precisely on the certificate template.

    Args:
        img: The image array to draw on (OpenCV format)
        text: The text string to draw
        font: OpenCV font constant (e.g., cv2.FONT_HERSHEY_SIMPLEX)
        font_size: Font scale factor
        x_offset: Horizontal offset from center (positive = right)
        y_offset: Vertical offset from center (negative = up)
        color: Text color in BGR format
        thickness: Text thickness in pixels
    """
    # Calculate text dimensions
    text_size, _ = cv2.getTextSize(text, font, font_size, thickness)
    text_width, text_height = text_size

    # Calculate centered position with offsets
    img_height, img_width = img.shape[:2]
    text_x = int((img_width - text_width) / 2 + x_offset)
    text_y = int((img_height + text_height) / 2 - y_offset)

    # Draw text on image
    cv2.putText(
        img, text, (text_x, text_y),
        font, font_size, color, thickness,
        lineType=cv2.LINE_AA  # Anti-aliased for better quality
    )
    logger.debug(f"Drew text '{text}' at position ({text_x}, {text_y})")


def format_date_string(day: Optional[int], month: Optional[str], year: Optional[int]) -> str:
    """
    Format date components into a standardized string.

    Args:
        day: Day of the month
        month: Month abbreviation (e.g., 'Mar', 'Apr')
        year: Year (can be 2 or 4 digits)

    Returns:
        str: Formatted date string (e.g., '25 Mar 22')
    """
    if all(v is not None for v in [day, month, year]):
        return f"{day} {month} {year}"
    return "Date Not Available"


def generate_certificate(
    row: pd.Series,
    template_path: str,
    output_dir: str
) -> Tuple[bool, Optional[str]]:
    """
    Generate a single certificate (PNG and PDF) for a participant.

    Args:
        row: Pandas Series containing participant data
        template_path: Path to the certificate template image
        output_dir: Directory to save generated certificates

    Returns:
        Tuple[bool, Optional[str]]: Success status and generated file path or error message
    """
    try:
        # Load template image
        img = cv2.imread(template_path)
        if img is None:
            raise ValueError(f"Failed to load template: {template_path}")

        # Extract and format participant name
        participant_name = str(row["Name"]).title()
        logger.info(f"Generating certificate for: {participant_name}")

        # Draw participant name
        draw_text_on_image(
            img, participant_name,
            FONT_FACE, TextPosition.INTERN_NAME_FONT_SIZE,
            TextPosition.INTERN_NAME_X, TextPosition.INTERN_NAME_Y
        )

        # Format and draw start date
        # TODO: Uncomment and use actual data when date columns are available
        # from_date = format_date_string(row["from"], row["s_month"], row["s_year"])
        from_date = '25 Mar 22'  # Placeholder
        draw_text_on_image(
            img, from_date,
            FONT_FACE, TextPosition.DATE_FONT_SIZE,
            TextPosition.DATE_FROM_X, TextPosition.DATE_Y
        )
        logger.debug(f"Start date: {from_date}")

        # Format and draw end date
        # TODO: Uncomment and use actual data when date columns are available
        # to_date = format_date_string(row["to"], row["e_month"], row["e_year"])
        to_date = '14 Apr 22'  # Placeholder
        draw_text_on_image(
            img, to_date,
            FONT_FACE, TextPosition.DATE_FONT_SIZE,
            TextPosition.DATE_TO_X, TextPosition.DATE_Y
        )
        logger.debug(f"End date: {to_date}")

        # Construct output file paths
        certificate_filename = f"{CERTIFICATE_PREFIX}_{participant_name}"
        png_path = os.path.join(output_dir, f"{certificate_filename}.png")
        pdf_path = os.path.join(output_dir, f"{certificate_filename}.pdf")

        # Save PNG certificate
        success = cv2.imwrite(png_path, img)
        if not success:
            raise IOError(f"Failed to save PNG: {png_path}")
        logger.info(f"Saved PNG: {png_path}")

        # Convert to PDF
        with open(pdf_path, "wb") as pdf_file:
            pdf_file.write(img2pdf.convert(png_path))
        logger.info(f"Saved PDF: {pdf_path}")

        return True, certificate_filename

    except KeyError as e:
        error_msg = f"Missing required column in data: {e}"
        logger.error(error_msg)
        return False, error_msg
    except Exception as e:
        error_msg = f"Error generating certificate: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return False, error_msg


def generate_certificates_batch(
    participants_file: str = PARTICIPANTS_FILE,
    template_file: str = TEMPLATE_FILE,
    output_dir: str = OUTPUT_DIRECTORY
) -> dict:
    """
    Generate certificates for all participants in batch mode.

    Args:
        participants_file: Path to Excel file containing participant data
        template_file: Path to certificate template image
        output_dir: Directory to save generated certificates

    Returns:
        dict: Statistics about the generation process (success, failed, total)
    """
    logger.info("=" * 70)
    logger.info("Certificate Generation System Started")
    logger.info("=" * 70)

    stats = {"total": 0, "success": 0, "failed": 0, "errors": []}

    try:
        # Validate input files
        validate_file_exists(participants_file)
        validate_file_exists(template_file)

        # Ensure output directory exists
        ensure_output_directory(output_dir)

        # Load participant data
        logger.info(f"Loading participants from: {participants_file}")
        participants_df = pd.read_excel(participants_file)
        stats["total"] = len(participants_df)
        logger.info(f"Found {stats['total']} participants to process")

        # Process each participant
        for index, row in participants_df.iterrows():
            success, result = generate_certificate(row, template_file, output_dir)

            if success:
                stats["success"] += 1
            else:
                stats["failed"] += 1
                stats["errors"].append(f"Row {index + 1}: {result}")

        # Log summary
        logger.info("=" * 70)
        logger.info("Certificate Generation Complete")
        logger.info(f"Total Processed: {stats['total']}")
        logger.info(f"Successful: {stats['success']}")
        logger.info(f"Failed: {stats['failed']}")
        logger.info("=" * 70)

        if stats["errors"]:
            logger.warning("Errors encountered:")
            for error in stats["errors"]:
                logger.warning(f"  - {error}")

    except FileNotFoundError as e:
        logger.error(f"File not found: {e}")
        sys.exit(1)
    except pd.errors.EmptyDataError:
        logger.error("Excel file is empty or invalid")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        sys.exit(1)
    finally:
        # Clean up OpenCV windows
        cv2.destroyAllWindows()

    return stats


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main() -> None:
    """Main entry point for the certificate generation system."""
    generate_certificates_batch()


if __name__ == '__main__':
    main()


