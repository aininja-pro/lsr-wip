"""
PDF Conversion Module

Converts .docx files to PDF using LibreOffice in headless mode.
Falls back gracefully if LibreOffice is not available.
"""

import io
import subprocess
import tempfile
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_libreoffice_available = None


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is installed and accessible."""
    global _libreoffice_available
    if _libreoffice_available is not None:
        return _libreoffice_available

    for cmd in ['libreoffice', 'soffice']:
        try:
            result = subprocess.run(
                [cmd, '--version'],
                capture_output=True, timeout=10
            )
            if result.returncode == 0:
                version = result.stdout.decode().strip()
                logger.info(f"LibreOffice found: {version}")
                _libreoffice_available = True
                return True
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue

    logger.warning("LibreOffice not found — PDF conversion will be unavailable")
    _libreoffice_available = False
    return False


def convert_docx_to_pdf(docx_bytes: io.BytesIO, output_filename: str) -> io.BytesIO | None:
    """
    Convert a .docx BytesIO to PDF using LibreOffice.

    Args:
        docx_bytes: BytesIO containing the .docx file
        output_filename: desired PDF filename (used for logging only)

    Returns:
        BytesIO containing the PDF, or None if conversion fails
    """
    if not is_libreoffice_available():
        logger.warning(f"Skipping PDF conversion for {output_filename}: LibreOffice not available")
        return None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            input_path = tmpdir_path / 'input.docx'
            output_dir = tmpdir_path / 'output'
            output_dir.mkdir()

            # Write docx to temp file
            input_path.write_bytes(docx_bytes.getvalue())

            # Find the LibreOffice command
            cmd = 'libreoffice'
            try:
                subprocess.run(['libreoffice', '--version'], capture_output=True, timeout=5)
            except FileNotFoundError:
                cmd = 'soffice'

            # Convert to PDF
            result = subprocess.run(
                [cmd, '--headless', '--convert-to', 'pdf',
                 '--outdir', str(output_dir), str(input_path)],
                capture_output=True, timeout=60
            )

            if result.returncode != 0:
                stderr = result.stderr.decode(errors='replace')
                logger.error(f"LibreOffice conversion failed for {output_filename}: {stderr}")
                return None

            # Find the output PDF
            pdf_files = list(output_dir.glob('*.pdf'))
            if not pdf_files:
                logger.error(f"No PDF output found for {output_filename}")
                return None

            pdf_bytes = io.BytesIO(pdf_files[0].read_bytes())
            pdf_bytes.seek(0)
            logger.info(f"Converted {output_filename}: {len(pdf_bytes.getvalue())} bytes")
            return pdf_bytes

    except subprocess.TimeoutExpired:
        logger.error(f"LibreOffice timed out converting {output_filename}")
        return None
    except Exception as e:
        logger.error(f"PDF conversion error for {output_filename}: {e}")
        return None
