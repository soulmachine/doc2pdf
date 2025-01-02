#!/usr/bin/env python3

import argparse
import email
from pathlib import Path
import base64
from email import policy

from bs4 import BeautifulSoup
from weasyprint import HTML
from xhtml2pdf import pisa


def _extract_html(input_path: Path) -> str:
    """Extract HTML content from the MHTML file."""
    with open(input_path, 'r', encoding='utf-8') as f:
        # Use policy.default to handle modern email formats
        message = email.message_from_file(f, policy=policy.default)
        html_content = None
        images = {}

        for part in message.walk():
            content_type = part.get_content_type()
            content_location = part.get('Content-Location', '')
            content_transfer_encoding = part.get('Content-Transfer-Encoding', '')

            if content_type == 'text/html':
                html_content = part.get_payload(decode=True).decode('utf-8')
                html_content = _clean_html(html_content) # Clean the HTML
            elif (content_type.startswith('image/') or
                  content_type == 'application/octet-stream') and \
                 content_transfer_encoding.lower() == 'base64':
                if content_location:
                    # Extract filename from Content-Location
                    filename = content_location.split('/')[-1]
                    # Decode base64 content
                    images[filename] = f"data:{content_type};base64," + \
                        base64.b64encode(part.get_payload(decode=True)).decode('ascii')

        # Replace image references with data URIs
        if html_content and images:
            for filename, data_uri in images.items():
                # Replace both src and data-image-src attributes
                html_content = html_content.replace(f'src="{filename}"', f'src="{data_uri}"')
                html_content = html_content.replace(f'data-image-src="{filename}"', f'data-image-src="{data_uri}"')

        return html_content

def _clean_html(html_content: str) -> str:
    """Clean the HTML content by fixing encodings and removing unnecessary elements."""
    if not html_content:
        raise ValueError("No HTML content found in the file")


    html_content = html_content.replace('=3D', '=') # Decode QuotedPrintable content (=3D becomes =)
    html_content = html_content.replace('=2D', ' ') # Replace `=20` with a space
    html_content = html_content.replace('=\n', '')  # Remove soft line breaks

    # Parse HTML with BeautifulSoup for better cleaning
    soup = BeautifulSoup(html_content, 'html.parser')

    # Remove Microsoft Office specific elements
    for element in soup.find_all(['xml', 'o:p']):
        element.decompose()

    # Remove empty divs
    for div in soup.find_all('div', recursive=True):
        if div.string is None and len(div.contents) == 0:
            div.decompose()

    return str(soup)

def mhtml_to_html(mhtml_file: Path) -> str:
    """Convert MHTML to clean HTML."""
    # Extract HTML content
    return _extract_html(mhtml_file)

def html_to_pdf(html_content: str, pdf_file: Path) -> bool:
    """Convert HTML to PDF using WeasyPrint.

    Note: While WeasyPrint supports embedded images, some layout issues may occur
    with images being partially cut off. For complex layouts, consider manual
    adjustments to the HTML/CSS.
    """
    try:
        HTML(string=html_content).write_pdf(pdf_file)
        return True
    except Exception as e:
        return False

def html_to_pdf_xhtml2pdf(html_content: str, pdf_file: Path) -> bool:
    """Convert clean HTML to PDF using xhtml2pdf.

    Note: This implementation has limited support for embedded images as xhtml2pdf
    doesn't fully support data URIs for images. Consider using html_to_pdf() for
    better image handling.
    """
    try:
        with open(pdf_file, 'w+b') as output_file:
            pisa.CreatePDF(html_content, dest=output_file)
            return True
    except Exception as e:
        return False

def convert_mhtml_to_pdf(input_path: Path, output_path: Path) -> bool:
    """Convert the MHTML file to PDF."""
    try:
        # Clean the HTML
        cleaned_html = _extract_html(input_path)

        # Write html to file in same location as PDF
        html_output_path = output_path.with_suffix('.html')
        with open(html_output_path, 'w', encoding='utf-8') as f:
            f.write(cleaned_html)

        if html_to_pdf(cleaned_html, output_path):
            print(f"Successfully converted {input_path} to {output_path}")
            return True
        else:
            print(f"Error converting {input_path} to {output_path}")
            return False
    except Exception as e:
        print(f"Error converting file: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert MHTML/DOC files to PDF')
    parser.add_argument('input', help='Input .doc or .mhtml file or directory')
    parser.add_argument('output', help='Output PDF file or directory')
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if input_path.is_dir():
        # Process directory
        if not output_path.exists():
            output_path.mkdir(parents=True)
        # Support multiple extensions
        for ext in ['*.mhtml', '*.mht', '*.doc']:
            for input_file in input_path.rglob(ext):
                output_file = output_path / input_file.with_suffix('.pdf').name
                convert_mhtml_to_pdf(input_file, output_file)
    else:
        # Process single file
        if input_path.suffix.lower() not in ['.mhtml', '.mht', '.doc']:
            raise ValueError("Input file must be .mhtml, .mht, or .doc")
        convert_mhtml_to_pdf(input_path, output_path)

if __name__ == '__main__':
    main()