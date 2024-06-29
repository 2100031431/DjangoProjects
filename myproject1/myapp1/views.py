# myapp1/views.py
from django.shortcuts import render, redirect
from django.http import HttpResponse
import pandas as pd
from io import BytesIO
from .forms import MyForm

def index(request):
    if 'submissions' not in request.session:
        request.session['submissions'] = []

    if request.method == 'POST':
        form = MyForm(request.POST)
        if form.is_valid():
            submission = {
                'Name': form.cleaned_data['name'],
                'Email': form.cleaned_data['email'],
                'Message': form.cleaned_data['message']
            }
            submissions = request.session['submissions']
            submissions.append(submission)
            request.session['submissions'] = submissions
            return redirect('index')  # Redirect to the same page to clear the form
    else:
        form = MyForm()

    return render(request, 'index.html', {'form': form})

def download_excel(request):
    submissions = request.session.get('submissions', [])

    if not submissions:
        return HttpResponse('No data to export', content_type='text/plain')

    df = pd.DataFrame(submissions)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    response = HttpResponse(output, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=form_data.xlsx'
    return response
