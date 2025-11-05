"""
Simple Certificate Generator

This module generates basic certificates by overlaying participant names onto
a certificate template. It reads participant data from an Excel file and outputs
PNG certificates for each participant.

Author: Certificate Generation System
Date: 2025
"""

import logging
import os
import sys
from pathlib import Path
from typing import Tuple, Optional

import cv2
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
PARTICIPANTS_FILE = "certificate_prospects/Interns Datails copy.xlsx"
TEMPLATE_FILE = "templates/template1-1.png"
OUTPUT_DIRECTORY = "generated_certificates/"

# Font configuration (BGR format for OpenCV, not RGB)
FONT_COLOR = (68, 102, 253)  # Professional blue color in BGR
FONT_FACE = cv2.FONT_HERSHEY_SIMPLEX
FONT_SIZE = 3.8
FONT_THICKNESS = 10

# Text positioning adjustments (in pixels relative to center)
# X adjustment: pixels to the right of vertical center (can be negative for left)
COORDINATE_X_ADJUSTMENT = 7
# Y adjustment: pixels above horizontal center (can be negative for below)
COORDINATE_Y_ADJUSTMENT = 78


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def validate_file_exists(file_path: str) -> bool:
    """
    Validate that a required file exists.

    Args:
        file_path: Path to the file to validate

    Returns:
        bool: True if file exists

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


def calculate_centered_position(
    img: np.ndarray,
    text: str,
    font: int,
    font_size: float,
    thickness: int,
    x_adjustment: int = 0,
    y_adjustment: int = 0
) -> Tuple[int, int]:
    """
    Calculate the position to center text on an image with optional adjustments.

    Args:
        img: The image array (OpenCV format)
        text: The text string to position
        font: OpenCV font constant
        font_size: Font scale factor
        thickness: Text thickness in pixels
        x_adjustment: Horizontal offset from center (positive = right, negative = left)
        y_adjustment: Vertical offset from center (positive = down, negative = up)

    Returns:
        Tuple[int, int]: (x, y) coordinates for text placement
    """
    # Get text dimensions
    text_size, _ = cv2.getTextSize(text, font, font_size, thickness)
    text_width, text_height = text_size

    # Get image dimensions
    img_height, img_width = img.shape[:2]

    # Calculate centered position with adjustments
    text_x = int((img_width - text_width) / 2 + x_adjustment)
    text_y = int((img_height + text_height) / 2 - y_adjustment)

    return text_x, text_y


def generate_simple_certificate(
    participant_name: str,
    template_path: str,
    output_path: str
) -> bool:
    """
    Generate a simple certificate with just the participant's name.

    Args:
        participant_name: Name to be displayed on the certificate
        template_path: Path to the certificate template image
        output_path: Full path where the certificate should be saved

    Returns:
        bool: True if certificate was generated successfully, False otherwise
    """
    try:
        # Load template image
        img = cv2.imread(template_path)
        if img is None:
            raise ValueError(f"Failed to load template: {template_path}")

        # Format participant name
        formatted_name = participant_name.title()

        # Calculate text position
        text_x, text_y = calculate_centered_position(
            img, formatted_name,
            FONT_FACE, FONT_SIZE, FONT_THICKNESS,
            COORDINATE_X_ADJUSTMENT, COORDINATE_Y_ADJUSTMENT
        )

        # Draw text on certificate
        cv2.putText(
            img, formatted_name,
            (text_x, text_y),
            FONT_FACE, FONT_SIZE, FONT_COLOR, FONT_THICKNESS,
            lineType=cv2.LINE_AA  # Anti-aliased for better quality
        )

        # Save certificate
        success = cv2.imwrite(output_path, img)
        if not success:
            raise IOError(f"Failed to save certificate: {output_path}")

        logger.info(f"Generated certificate: {output_path}")
        return True

    except Exception as e:
        logger.error(f"Error generating certificate for '{participant_name}': {e}")
        return False


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
        dict: Statistics about the generation process (total, success, failed)
    """
    logger.info("=" * 70)
    logger.info("Simple Certificate Generation System Started")
    logger.info("=" * 70)

    stats = {"total": 0, "success": 0, "failed": 0}

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

        # Validate required columns
        if "Name" not in participants_df.columns:
            raise ValueError("Excel file must contain a 'Name' column")

        # Process each participant
        for index, row in participants_df.iterrows():
            try:
                participant_name = str(row["Name"])
                certificate_filename = f"{participant_name}.png"
                output_path = os.path.join(output_dir, certificate_filename)

                # Generate certificate
                success = generate_simple_certificate(
                    participant_name,
                    template_file,
                    output_path
                )

                if success:
                    stats["success"] += 1
                else:
                    stats["failed"] += 1

            except KeyError as e:
                logger.error(f"Row {index + 1}: Missing required column {e}")
                stats["failed"] += 1
            except Exception as e:
                logger.error(f"Row {index + 1}: Unexpected error - {e}")
                stats["failed"] += 1

        # Log summary
        logger.info("=" * 70)
        logger.info("Certificate Generation Complete")
        logger.info(f"Total Processed: {stats['total']}")
        logger.info(f"Successful: {stats['success']}")
        logger.info(f"Failed: {stats['failed']}")
        logger.info("=" * 70)

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
    """Main entry point for the simple certificate generation system."""
    generate_certificates_batch()


if __name__ == '__main__':
    main()