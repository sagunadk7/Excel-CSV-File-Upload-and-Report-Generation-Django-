from django.shortcuts import render
from django.core.mail import EmailMessage
from django.conf import settings
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
from io import BytesIO
import tempfile
import os

def upload_file(request):
    if request.method == 'POST' and 'file' in request.FILES:
        file = request.FILES['file']

        try:
            if file.name.endswith('.xls') or file.name.endswith('.xlsx'):
                df = pd.read_excel(file)
            elif file.name.endswith('.csv'):
                df = pd.read_csv(file)
            else:
                return render(request, 'index.html', {'error': 'Invalid file format. Please upload an Excel or CSV file.'})
        except Exception as e:
            return render(request, 'index.html', {'error': f'Error processing file: {e}'})

        try:
            numeric_df = df.select_dtypes(include=['number'])
            if not numeric_df.empty:
                summary_html = numeric_df.describe().to_html()
            else:
                summary_html = "<p>No numeric data to summarize.</p>"
        except Exception as e:
            return render(request, 'index.html', {'error': f'Error generating summary: {e}'})

        try:
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--headless")
            driver = webdriver.Chrome(service=service, options=options)

            with tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.html') as temp_html_file:
                temp_html_file.write(summary_html)
                temp_html_path = temp_html_file.name

            driver.get("file://" + os.path.abspath(temp_html_path))

            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_screenshot_file:
                screenshot_path = temp_screenshot_file.name
            driver.save_screenshot(screenshot_path)
            driver.quit()

            with Image.open(screenshot_path) as img:
                jpeg_image = BytesIO()
                img.convert("RGB").save(jpeg_image, format="JPEG")
                jpeg_image.seek(0)

            email = EmailMessage(
                subject=f'Python Assignment - {request.user.username}',
                body='Here is your summary report in JPEG format.',
                from_email=settings.DEFAULT_FROM_EMAIL,
                to=['your-mail@mail.com']
            )
            email.attach('summary_report.jpeg', jpeg_image.read(), 'image/jpeg')
            email.send()

            os.remove(temp_html_path)
            os.remove(screenshot_path)

        except Exception as e:
            return render(request, 'index.html', {'error': f'Error converting HTML to JPEG or sending email: {e}'})

        return render(request, 'index.html', {'success': 'Summary report has been sent successfully.'})

    return render(request, 'index.html')
