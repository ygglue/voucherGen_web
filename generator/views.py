from django.shortcuts import render
from django.http import HttpResponse
from django.contrib import messages
from .forms import UploadFileForm
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import io

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
                
                # Generate PDF
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter)
                elements = []
                
                # Add a custom header
                headers = ['Code', 'Duration', 'Price', 'Speed', 'Partner Name']
                data = [headers] + final_df.values.tolist()
                
                # Create the table
                t = Table(data)
                
                # Add style
                style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#ecf0f1')),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                ])
                t.setStyle(style)
                elements.append(t)
                
                doc.build(elements)
                
                buffer.seek(0)
                response = HttpResponse(buffer, content_type='application/pdf')
                response['Content-Disposition'] = 'attachment; filename="vouchers.pdf"'
                return response
            else:
                # No valid data found
                error_context = {
                    'form': form,
                    'error': 'No valid data could be extracted from the uploaded files. Please ensure the files contain a "WiFi5-List" sheet.'
                }
                return render(request, 'generator/upload.html', error_context)
        else:
            # Form is not valid
            print("Form errors:", form.errors)
    else:
        form = UploadFileForm()
    return render(request, 'generator/upload.html', {'form': form})
