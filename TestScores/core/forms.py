from django import forms

class UploadFileForm(forms.Form):
    overwrite_file = forms.FileField(label='Please choose the file you would like to overwrite:')
    extract_file = forms.FileField(label='Please choose the file you would like to extract scores from:')
    column_to_overwrite = forms.CharField(
        max_length=2,
        label='Please enter the specific Column that you would like to overwrite scores in:',
        help_text="<br>*NOTE: Please enter the column letter that needs to be overwritten in CAPITAL LETTER (e.g., D, F, G, etc.)"
    )
