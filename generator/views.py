from django.shortcuts import render
from django.http import HttpResponse
from django.contrib import messages
from .forms import UploadFileForm
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image
import io
import os
from django.conf import settings

def upload_view(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            files = request.FILES.getlist('files')
            all_data = []
            errors = []
            
            for f in files:
                try:
                    # Read the second sheet (index 1), assume no header
                    df = pd.read_excel(f, sheet_name=1, header=None)
                    all_data.append(df)
                except Exception as e:
                    error_msg = f"Error reading {f.name}: {str(e)}"
                    print(error_msg)
                    errors.append(error_msg)
            
            if all_data:
                final_df = pd.concat(all_data, ignore_index=True)
                
                # Get voucher image path
                voucher_img_path = os.path.join(settings.BASE_DIR, 'generator', 'static', 'generator', 'img', 'voucher.png')
                font_path = os.path.join(settings.BASE_DIR, 'generator', 'static', 'generator', 'fonts', 'Instruction.otf')
                
                # Register the custom font
                pdfmetrics.registerFont(TTFont('Instruction', font_path))
                
                # Get image dimensions
                img = Image.open(voucher_img_path)
                img_width, img_height = img.size
                
                # Create PDF with custom page size based on image
                buffer = io.BytesIO()
                
                # Convert pixels to points (72 points = 1 inch, assuming 96 DPI)
                pdf_width = img_width * 72 / 96
                pdf_height = img_height * 72 / 96
                
                # Calculate total height for all vouchers
                total_height = pdf_height * len(final_df)
                
                # Create canvas with single continuous page
                c = canvas.Canvas(buffer, pagesize=(pdf_width, total_height))
                
                # Set white background
                c.setFillColorRGB(1, 1, 1)  # White
                c.rect(0, 0, pdf_width, total_height, fill=1, stroke=0)
                
                # Draw each voucher
                for idx, row in final_df.iterrows():
                    # Calculate Y position for this voucher (from bottom)
                    y_offset = total_height - (idx + 1) * pdf_height
                    
                    # Draw the voucher image
                    c.drawImage(voucher_img_path, 0, y_offset, width=pdf_width, height=pdf_height)
                    
                    # Extract data from row
                    code = str(row[0]) if pd.notna(row[0]) else ""
                    price_raw = str(row[2]) if pd.notna(row[2]) else ""
                    partner_name = str(row[4]) if pd.notna(row[4]) else ""
                    
                    # Format price as ₱10 (remove decimals)
                    try:
                        price_num = float(price_raw.replace('₱', '').replace(',', ''))
                        price = f"₱{int(price_num)}"
                    except:
                        price = price_raw
                    
                    # Partner Name - top banner (centered) - BLACK TEXT
                    c.setFont("Instruction", 24)
                    c.setFillColorRGB(1, 1, 1)  # White
                    partner_x = pdf_width / 2
                    partner_y = y_offset + pdf_height - 65
                    c.drawCentredString(partner_x, partner_y, partner_name)
                    
                    # Code - after "CODE:" label (left aligned) - BLACK TEXT
                    c.setFont("Instruction", 56)
                    c.setFillColorRGB(0, 0, 0)  # Black
                    code_x = 130
                    code_y = y_offset + pdf_height - 115
                    c.drawString(code_x, code_y, code)
                    
                    # Price - bottom right box (centered in the box) - WHITE TEXT
                    c.setFont("Instruction", 90)
                    c.setFillColorRGB(1, 1, 1)  # White
                    price_x = pdf_width - 55
                    price_y = y_offset + pdf_height - 210
                    c.drawCentredString(price_x, price_y, price)

                
                c.save()
                
                # Generate filename from first row data
                from datetime import datetime
                first_row = final_df.iloc[0]
                price_raw = str(first_row[2]) if pd.notna(first_row[2]) else "Unknown"
                partner_name = str(first_row[4]) if pd.notna(first_row[4]) else "Unknown"
                timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
                
                # Format price as ₱10
                try:
                    price_num = float(price_raw.replace('₱', '').replace(',', ''))
                    price_formatted = f"₱{int(price_num)}"
                except:
                    price_formatted = price_raw
                
                # Clean filename (remove invalid characters)
                partner_clean = partner_name.replace('/', '-').replace('\\', '-')
                filename = f"{price_formatted} {partner_clean} {timestamp}.pdf"
                
                buffer.seek(0)
                response = HttpResponse(buffer, content_type='application/pdf')
                # Use UTF-8 encoding for filename with special characters
                from urllib.parse import quote
                encoded_filename = quote(filename)
                response['Content-Disposition'] = f'attachment; filename="{filename}"; filename*=UTF-8\'\'{encoded_filename}'
                return response
            else:
                # No valid data found
                error_context = {
                    'form': form,
                    'error': 'No valid data could be extracted from the uploaded files. Please ensure the files contain data in the second sheet.'
                }
                return render(request, 'generator/upload.html', error_context)
        else:
            # Form is not valid
            print("Form errors:", form.errors)
    else:
        form = UploadFileForm()
    return render(request, 'generator/upload.html', {'form': form})
