"""
Configuration Module for Certificate Generation System

This module contains all configuration settings for the certificate generation
system. Modify these values to customize the certificate generation behavior.

Author: Certificate Generation System
Date: 2025
"""

import cv2
from dataclasses import dataclass
from typing import Tuple


@dataclass
class FileConfig:
    """File path configuration."""

    # Input files
    participants_file: str = "certificate_prospects/health_interns.xlsx"
    participants_file_simple: str = "certificate_prospects/Interns Datails copy.xlsx"

    # Template files
    template_full: str = "templates/template1.png"
    template_simple: str = "templates/template1-1.png"

    # Output directories
    output_directory: str = "generated_certificates/"

    # Logging
    log_file: str = "certificate_generation.log"


@dataclass
class FontConfig:
    """Font and text styling configuration."""

    # Font properties
    font_face: int = cv2.FONT_HERSHEY_SIMPLEX
    font_color_bgr: Tuple[int, int, int] = (68, 102, 253)  # Professional blue
    font_thickness: int = 10

    # Font sizes
    name_font_size: float = 3.8
    date_font_size: float = 1.5


@dataclass
class PositionConfig:
    """Text positioning configuration for certificate templates."""

    # Name positioning (simple certificate)
    name_x_offset_simple: int = 7
    name_y_offset_simple: int = 78

    # Name positioning (full certificate)
    name_x_offset_full: int = 100
    name_y_offset_full: int = 78

    # Date positioning
    date_from_x_offset: int = 130
    date_to_x_offset: int = 535
    date_y_offset: int = -80


@dataclass
class CertificateConfig:
    """Certificate generation configuration."""

    # Certificate naming
    certificate_prefix: str = "Internship_Completion_Certificate"

    # Output formats
    generate_png: bool = True
    generate_pdf: bool = True

    # Image quality settings
    use_antialiasing: bool = True  # Use LINE_AA for better text quality

    # Placeholder dates (use when actual dates not available in data)
    default_start_date: str = "25 Mar 22"
    default_end_date: str = "14 Apr 22"


@dataclass
class LogConfig:
    """Logging configuration."""

    # Log levels: DEBUG, INFO, WARNING, ERROR, CRITICAL
    log_level: str = "INFO"

    # Log format
    log_format: str = "%(asctime)s - %(levelname)s - %(message)s"

    # Enable file logging
    enable_file_logging: bool = True

    # Enable console logging
    enable_console_logging: bool = True


# ============================================================================
# CONFIGURATION INSTANCES
# ============================================================================

# Create default configuration instances
file_config = FileConfig()
font_config = FontConfig()
position_config = PositionConfig()
certificate_config = CertificateConfig()
log_config = LogConfig()


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_config_summary() -> str:
    """
    Generate a summary of current configuration settings.

    Returns:
        str: Formatted configuration summary
    """
    summary = []
    summary.append("=" * 70)
    summary.append("Certificate Generation System - Configuration Summary")
    summary.append("=" * 70)
    summary.append("")
    summary.append("File Configuration:")
    summary.append(f"  Participants File: {file_config.participants_file}")
    summary.append(f"  Template File: {file_config.template_full}")
    summary.append(f"  Output Directory: {file_config.output_directory}")
    summary.append("")
    summary.append("Font Configuration:")
    summary.append(f"  Font Face: HERSHEY_SIMPLEX")
    summary.append(f"  Font Color (BGR): {font_config.font_color_bgr}")
    summary.append(f"  Name Font Size: {font_config.name_font_size}")
    summary.append(f"  Date Font Size: {font_config.date_font_size}")
    summary.append("")
    summary.append("Certificate Configuration:")
    summary.append(f"  Generate PNG: {certificate_config.generate_png}")
    summary.append(f"  Generate PDF: {certificate_config.generate_pdf}")
    summary.append(f"  Use Antialiasing: {certificate_config.use_antialiasing}")
    summary.append("")
    summary.append("=" * 70)
    return "\n".join(summary)


def validate_configuration() -> bool:
    """
    Validate configuration settings.

    Returns:
        bool: True if configuration is valid

    Raises:
        ValueError: If configuration contains invalid values
    """
    # Validate font size
    if font_config.name_font_size <= 0 or font_config.date_font_size <= 0:
        raise ValueError("Font sizes must be positive numbers")

    # Validate colors (BGR values 0-255)
    for component in font_config.font_color_bgr:
        if not 0 <= component <= 255:
            raise ValueError("Color values must be between 0 and 255")

    # Validate thickness
    if font_config.font_thickness <= 0:
        raise ValueError("Font thickness must be a positive number")

    return True


if __name__ == "__main__":
    # Print configuration summary when run directly
    print(get_config_summary())

    # Validate configuration
    try:
        validate_configuration()
        print("\nConfiguration validation: PASSED")
    except ValueError as e:
        print(f"\nConfiguration validation: FAILED - {e}")
