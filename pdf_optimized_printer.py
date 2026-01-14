#!/usr/bin/env python3
"""
PDF Slide Optimizer - Converts PowerPoint-exported PDFs to print-optimized format
Fits multiple slides per page when possible on standard 8.5x11" paper
"""

import fitz  # PyMuPDF
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import io
import os
import sys

def optimize_pdf_for_printing(input_pdf, output_pdf, slides_per_page='auto', dpi=200):
    """
    Optimize PDF by fitting multiple slides per page
    
    Args:
        input_pdf: Path to input PDF file
        output_pdf: Path to output PDF file
        slides_per_page: 'auto', 2, 4, 6, or 9
        dpi: Resolution for converting PDF to images (higher = better quality but slower)
    """
    
    print("Converting PDF pages to images...")
    
    # Open the PDF
    pdf_document = fitz.open(input_pdf)
    total_slides = len(pdf_document)
    
    # Convert PDF pages to PIL images
    images = []
    zoom = dpi / 72  # Convert DPI to zoom factor (72 is default DPI)
    mat = fitz.Matrix(zoom, zoom)
    
    for page_num in range(total_slides):
        page = pdf_document[page_num]
        pix = page.get_pixmap(matrix=mat)
        
        # Convert pixmap to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
        
        if (page_num + 1) % 10 == 0:
            print(f"  Converted {page_num + 1}/{total_slides} slides...")
    
    pdf_document.close()
    
    print(f"Converted {total_slides} slides")
    
    # Get first image dimensions
    first_img = images[0]
    img_width, img_height = first_img.size
    img_ratio = img_width / img_height
    
    # Standard letter size in points (1 inch = 72 points)
    page_width, page_height = letter  # 612 x 792 points (8.5" x 11")
    
    # Determine optimal slides per page if auto
    if slides_per_page == 'auto':
        if img_ratio > 1.2:  # Landscape
            slides_per_page = 2  # 2 slides vertically
        else:  # Portrait or square
            if img_width < img_height:
                slides_per_page = 2  # 2 portrait slides
            else:
                slides_per_page = 4  # 2x2 grid
    
    print(f"Image size: {img_width} x {img_height} pixels")
    print(f"Fitting {slides_per_page} slides per page")
    
    # Calculate layout
    layouts = {
        1: (1, 1),
        2: (1, 2),  # 1 column, 2 rows
        4: (2, 2),
        6: (2, 3),
        9: (3, 3)
    }
    
    cols, rows = layouts[slides_per_page]
    
    # Calculate dimensions for each slide on the page
    margin = 36  # 0.5 inch margin
    gap = 10     # Small gap between slides
    
    available_width = page_width - (2 * margin) - (cols - 1) * gap
    available_height = page_height - (2 * margin) - (rows - 1) * gap
    
    cell_width = available_width / cols
    cell_height = available_height / rows
    
    # Calculate scaled dimensions maintaining aspect ratio
    scale_w = cell_width / img_width
    scale_h = cell_height / img_height
    scale = min(scale_w, scale_h)
    
    scaled_width = img_width * scale
    scaled_height = img_height * scale
    
    print(f"Scaled slide size: {scaled_width:.1f} x {scaled_height:.1f} points")
    
    # Create PDF with images
    c = canvas.Canvas(output_pdf, pagesize=letter)
    
    slide_idx = 0
    page_count = 0
    
    while slide_idx < total_slides:
        page_count += 1
        print(f"Creating page {page_count}...")
        
        # Place slides on this page
        for j in range(slides_per_page):
            if slide_idx >= total_slides:
                break
            
            # Calculate position
            row = j // cols
            col = j % cols
            
            # Calculate x, y position (bottom-left corner)
            x = margin + col * (cell_width + gap) + (cell_width - scaled_width) / 2
            # In reportlab, y=0 is at bottom, so we calculate from bottom up
            y = page_height - margin - (row + 1) * cell_height - row * gap + (cell_height - scaled_height) / 2
            
            # Draw the image
            img = images[slide_idx]
            c.drawImage(ImageReader(img), x, y, width=scaled_width, height=scaled_height)
            
            # Optional: Draw a light border around each slide
            c.setStrokeColorRGB(0.8, 0.8, 0.8)
            c.setLineWidth(0.5)
            c.rect(x, y, scaled_width, scaled_height)
            
            slide_idx += 1
        
        # Finish this page
        c.showPage()
    
    # Save the PDF
    c.save()
    
    print(f"\nCreated {page_count} print pages from {total_slides} slides")
    print(f"Saved to: {output_pdf}")
    
if __name__ == "__main__":
    # Hardcoded file path - change this to your file
    input_pdf = r"C:\Users\benja\Downloads\Note-2.pdf"
    output_pdf = r"C:\Users\benja\Downloads\Multivariate_Notes_1.pdf"
    slides_per_page = 'auto'  # Change to 2, 4, 6, or 9 if you want
    
    if not os.path.exists(input_pdf):
        print(f"Error: File '{input_pdf}' not found")
        sys.exit(1)
    
    optimize_pdf_for_printing(input_pdf, output_pdf, slides_per_page)
    print(f"\nDone! You can now print {output_pdf} on standard 8.5x11 paper.")