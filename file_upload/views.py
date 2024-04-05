from django.shortcuts import render
from .salary import open_excel
from django.http.response import HttpResponse

def upload_file(request):
    if request.method == 'POST' and request.FILES.get('file'):
        uploaded_file = request.FILES['file']
        excel_data, salary, basic_details = open_excel(uploaded_file)
        cols = ['Date', 'Day', 'Time in', 'Time out', 'Working hour', 'Overtime', 'Per hour']
        return render(request, 'file_upload/success.html', {'filename': uploaded_file.name, 'excel_data': excel_data, 'cols': cols, 'salary': salary, 'basic_details': basic_details})
    return render(request, 'file_upload/upload.html')
