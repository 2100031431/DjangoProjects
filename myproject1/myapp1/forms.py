# myapp1/forms.py
from django import forms
from .models import FormData

class MyForm(forms.ModelForm):
    class Meta:
        model = FormData
        fields = ['name', 'email', 'message']
