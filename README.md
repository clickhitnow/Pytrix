Pytrix
Pytrix is a Flask-based reporting tool that analyzes Excel data to automatically categorize support tickets, track department performance, and generate insightful monthly reports.

📌 Features
Upload and analyze Excel data seamlessly.
Automatically categorize requests as within or outside working hours.
Generate monthly summaries and reports.
Identify top customers and frequent ticket sources.
User-friendly web interface built with Flask.

🚀 Installation
Step 1: Clone the Repository
git clone <repository_url>
cd Pytrix


Step 2: Set Up Environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

Step 3: Install Dependencies

⚙️ Configuration
Adjust the upload and report folder paths in app.py:
UPLOAD_FOLDER = 'path/to/upload'
REPORT_FOLDER = 'path/to/reports'

Ensure these directories exist or will be created automatically by the application.

▶️ Running the Application
python app.py
Navigate to http://localhost:5000 in your web browser.

📁 Usage
Upload Excel file: Use the provided interface to upload your Excel file.
Preview Data: View a summary before downloading the full report.
Download Report: Generate and download a comprehensive Excel report.

🛠 Built With
Flask
Pandas
XlsxWriter

📃 License
This project is open-source under the MIT License.

🤝 Contributing
Contributions are welcome. Please open issues or submit pull requests!




