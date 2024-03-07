from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import re
from flask import Flask, flash, request, url_for, redirect, session, render_template
import datetime
import smtplib
import requests
import re
from bson import ObjectId
import gspread

from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from authlib.integrations.flask_client import OAuth
from pymongo import MongoClient
from googleapiclient.discovery import build

app = Flask(__name__)
app.secret_key = 'justtry something'

# MongoDB connection
client = MongoClient("mongodb+srv://tajulislam06:trynothing@cluster0.x3lpi33.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0")
db = client['LLM_Marking']
teachers_collection = db['teacher']  # Collection for storing teachers
courses_collection = db['courses']  # Collection for storing courses
students_collection = db['students']
questions_collection = db['questions']
form_links_collection = db['form_links']
rubrics_collection = db['rubrics']

# oauth config
oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id='948138056692-u688i32prmmuhk4o9i3i00r4m3r4tipl.apps.googleusercontent.com',
    client_secret='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',
    access_token_url='https://accounts.google.com/o/oauth2/token',
    access_token_params=None,
    authorize_url='https://accounts.google.com/o/oauth2/auth',
    authorize_params=None,
    api_base_url='https://www.googleapis.com/oauth2/v1/',
    userinfo_endpoint='https://openidconnect.googleapis.com/v1/userinfo',
    client_kwargs={'scope': 'email profile https://www.googleapis.com/auth/forms'}
)


def extract_google_sheets_key(google_sheets_link):
    pattern = re.compile(r'^https?://docs.google.com/spreadsheets/d/([a-zA-Z0-9-_]+)')
    match = pattern.match(google_sheets_link)
    if match:
        return match.group(1)
    else:
        return None


@app.route('/')
def home():
    email = dict(session).get('email', None)
    return render_template('login.html', email=email)


@app.route('/login')
def login():
    redirect_uri = url_for('authorize', _external=True)
    return google.authorize_redirect(redirect_uri)


@app.route('/authorize')
def authorize():
    google = oauth.create_client('google')
    try:
        token = google.authorize_access_token()
        resp = google.get('userinfo')
        resp.raise_for_status()
        user_info = resp.json()
        session['email'] = user_info['email']
        # Check if teacher exists in MongoDB
        teacher = teachers_collection.find_one({'email': user_info['email']})
        if teacher is None:
            # Add teacher to MongoDB
            teachers_collection.insert_one({'name': user_info['name'], 'email': user_info['email']})

    except Exception as e:
        print(f"Error in authorize route: {e}")
        return "An error occurred during authorization", 500  # Return a 500 Internal Server Error status code
    return redirect('/dashboard')


@app.route('/dashboard')
def dashboard():
    email = dict(session).get('email', None)
    if email:
        courses = courses_collection.find({'teacher_email': email})
        return render_template('dashboard.html', email=email, courses=courses)
    else:
        return redirect('/')


@app.route('/add_course', methods=['GET', 'POST'])
def add_course():
    if request.method == 'POST':
        # Extract course details from the form data
        course_name = request.form.get('course_name')
        course_section = request.form.get('course_section')
        description = request.form.get('description')
        teacher_email = dict(session).get('email')
        start_date = request.form.get('start_date')
        end_date = request.form.get('end_date')

        # Store course details in the database along with teacher email
        courses_collection.insert_one({
            'name': course_name,
            'section': course_section,
            'description': description,
            'teacher_email': teacher_email,
            'start_date': start_date,
            'end_date': end_date

            # Additional fields as needed
        })

        flash('Course added successfully')
        return redirect('/dashboard')

    return render_template('add_course.html')  # Render the form template


@app.route('/delete_course/<course_id>', methods=['POST'])
def delete_course(course_id):
    # Retrieve the course from the database
    course = courses_collection.find_one({'_id': ObjectId(course_id)})
    if course:
        # Delete the course
        courses_collection.delete_one({'_id': ObjectId(course_id)})
        flash('Course deleted successfully')
    else:
        flash('Course not found')
    return redirect('/dashboard')


@app.route('/edit_course', methods=['POST'])
def edit_course():
    course_id = request.form.get('course_id')
    name = request.form.get('name')
    description = request.form.get('description')
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')

    # Update the course in the database
    courses_collection.update_one(
        {'_id': ObjectId(course_id)},
        {'$set': {'name': name, 'description': description, 'start_date': start_date, 'end_date': end_date}}
    )

    flash('Course updated successfully')
    return redirect('/dashboard')


@app.route('/course/<course_id>', methods=['GET'])
def course_dashboard(course_id):
    email = session.get('email')
    if not email:
        return redirect('/dashboard')

    course = courses_collection.find_one({'_id': ObjectId(course_id)})
    if not course:
        return redirect('/dashboard')

    if request.method == 'GET':
        # Fetch assessments and their rubrics
        assessments = questions_collection.find({'course_id': ObjectId(course_id)})
        rubrics = rubrics_collection.find({'course_id': ObjectId(course_id)})
        return render_template('course_dashboard.html', email=email, course=course, assessments=assessments, rubrics=rubrics)


def extract_and_add_students(course_id, google_sheets_link):
    # Define the scope for Google Sheets API
    scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # Load credentials from the JSON key file (make sure you have this file)
    credentials = ServiceAccountCredentials.from_json_keyfile_name('creds1.json', scope)

    # Authenticate using the credentials
    client = gspread.authorize(credentials)

    # Extract the Google Sheets key from the provided link
    sheets_key = extract_google_sheets_key(google_sheets_link)

    if sheets_key:
        try:
            # Open the Google Sheets document by key
            doc = client.open_by_key(sheets_key)

            # Get the first sheet of the document
            worksheet = doc.sheet1

            # Extract all values from the worksheet
            data = worksheet.get_all_values()

            students_data = []
            for row in data:
                name, student_id, student_email = row
                students_data.append({
                    'name': name,
                    'student_id': student_id,
                    'email': student_email,
                    'course_id': ObjectId(course_id)  # Convert course_id to ObjectId
                })

            # Add students to the database
            students_collection.insert_many(students_data)

            return True  # Return True if students were successfully added
        except Exception as e:
            print(f"Error extracting or adding students: {e}")
            return False  # Return False if an error occurred
    else:
        print("Invalid Google Sheets link")
        return False  # Return False if the Google Sheets link is invalid
@app.route('/add_student/<course_id>', methods=['GET', 'POST'])
def add_student(course_id):
    email = session.get('email')
    if not email:
        return redirect('/dashboard')

    course = courses_collection.find_one({'_id': ObjectId(course_id)})
    if not course:
        return redirect('/dashboard')

    if request.method == 'POST':
        google_sheets_link = request.form.get('google_sheets_link')

        # Extract and add students from the Google Sheets link
        if extract_and_add_students(course_id, google_sheets_link):
            flash('Students added successfully')  # Flash a success message
        else:
            flash('Failed to add students. Please check the Google Sheets link.')

    flash('Use the Google Sheets link to add students.')
    return redirect(url_for('course_dashboard', course_id=course_id))
 
     

@app.route('/students/<course_id>')
def get_students(course_id):
    email = dict(session).get('email', None)
    if email:
        course = courses_collection.find_one({'_id': ObjectId(course_id)})
        if course:
            students = students_collection.find({'course_id': ObjectId(course_id)})
            return render_template('students_list.html', email=email, course=course, students=students)
    return redirect('/dashboard')


@app.route('/create_assessment/<course_id>', methods=['GET', 'POST'])
def create_assessment(course_id):
    if request.method == 'POST':
        google_sheets_link = request.form.get('google_sheets_link')
        deadline = request.form.get('deadline')
        assessment_name = request.form.get('assessment_name')

        # Extract questions from Google Sheets
        questions = extract_questions_from_sheets(google_sheets_link)

        # Generate a unique identifier for the assessment
        assessment_id = ObjectId()

        # Save assessment data to the database
        assessment_data = {
            '_id': assessment_id,
            'course_id': ObjectId(course_id),
            'assessment_name': assessment_name,
            'deadline': deadline,
            'questions': questions
        }
        questions_collection.insert_one(assessment_data)

        flash('Assessment created successfully')
        return redirect(url_for('course_dashboard', course_id=course_id))

    return render_template('create_assessment.html', course_id=course_id)


def extract_questions_from_sheets(google_sheets_link):
    # Define the scope for Google Sheets API
    scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # Load credentials from the JSON key file (make sure you have this file)
    credentials = ServiceAccountCredentials.from_json_keyfile_name('creds1.json', scope)

    # Authenticate using the credentials
    client = gspread.authorize(credentials)

    # Extract the Google Sheets key from the provided link
    sheets_key = extract_google_sheets_key(google_sheets_link)

    # Open the Google Sheets document by key
    doc = client.open_by_key(sheets_key)

    # Get the first sheet of the document
    worksheet = doc.get_worksheet(0)  # Assuming questions are on the first sheet

    # Extract all values from the worksheet
    values = worksheet.get_all_values()

    # Assuming the second row contains the questions
    questions = values[1]

    return questions


def extract_google_sheets_key(google_sheets_link):
    pattern = re.compile(r'/spreadsheets/d/([a-zA-Z0-9-_]+)')
    match = pattern.search(google_sheets_link)
    if match:
        return match.group(1)
    else:
        return None


# Test the function with the provided link
google_sheets_link = "https://docs.google.com/spreadsheets/d/12n149OGaAm5lkb8Xv5b89Xkkl4U3ghEhCexQOkvpbTo/edit?usp=sharing"
spreadsheet_id = extract_google_sheets_key(google_sheets_link)
print("Spreadsheet ID:", spreadsheet_id)


def extract_rubric_data_from_sheets(google_sheets_link):
    scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('creds1.json', scope)
    client = gspread.authorize(credentials)

    sheets_key = extract_google_sheets_key(google_sheets_link)
    if sheets_key:
        try:
            doc = client.open_by_key(sheets_key)
            worksheet = doc.sheet1  # Assuming the rubric data is in the first sheet
            rubric_data = worksheet.get_all_values()
            return rubric_data
        except Exception as e:
            print(f"Error extracting rubric data from Google Sheets: {e}")
            return None
    else:
        return None


@app.route('/assessment_list/<course_id>', methods=['GET'])
def assessment_list(course_id):
    email = session.get('email')
    if not email:
        return redirect('/')

    course = courses_collection.find_one({'_id': ObjectId(course_id)})
    if not course:
        return redirect('/dashboard')

    assessments = questions_collection.find({'course_id': ObjectId(course_id)})
    return render_template('assessment_list.html', email=email, course=course, assessments=assessments)


@app.route('/upload_rubric/<assessment_id>', methods=['POST'])
def upload_rubric(assessment_id):
    if request.method == 'POST':
        rubric_link = request.form.get('rubric_link')

        # Extract rubric data from the Google Sheets link (you can implement this function)
        rubric_data = extract_rubric_data_from_sheets(rubric_link)

        if rubric_data:
            # Save rubric data to the database
            rubrics_collection.insert_one({
                'assessment_id': ObjectId(assessment_id),
                'rubric_data': rubric_data
            })

            flash('Rubric uploaded successfully')
        else:
            flash('Failed to extract rubric data from the provided Google Sheets link. Please check the link.')

        return redirect(url_for('course_dashboard', course_id=assessment_id))
    return redirect('/dashboard')  # Redirect to dashboard if not a POST request





def generate_google_form(assessment_id):
    # Initialize credentials
    credentials = service_account.Credentials.from_service_account_file(
        'creds1.json',  # Replace with your service account credentials file path
        scopes=['https://www.googleapis.com/auth/forms']
    )

    # Build the Google Forms API service
    service = build('forms', 'v1', credentials=credentials)

    # Fetch questions from the database
    assessment = questions_collection.find_one({'_id': ObjectId(assessment_id)})
    if not assessment:
        return None  # Assessment not found

    # Extract questions from the assessment document
    questions = assessment.get('questions', [])
    if not questions:
        return None  # No questions found in the assessment

    # Construct form payload based on fetched questions
    form_payload = {
        "title": f"Assessment Form - {assessment_id}",
        "description": "Please complete this assessment form.",
        "items": []
    }

    # Add each question to the form payload
    for idx, question in enumerate(questions, start=1):
        form_payload['items'].append({
            "type": "TEXT",
            "title": f"Question {idx}: {question}",
            "isRequired": True
        })

    # Call the Google Forms API to create the form
    response = service.forms().create(body=form_payload).execute()

    # Extract the form URL from the response
    form_url = response.get('link', None)

    # Store the form URL in the assessment document
    if form_url:
        assessment['form_link'] = form_url
        questions_collection.update_one({'_id': ObjectId(assessment_id)}, {'$set': assessment})

    return form_url


def send_email(receiver_email, subject, body):
    # Your Gmail credentials
    sender_email = 'llm.markin299@gmail.com'
    password = 'cse299project'

    # Create a multipart message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Add body to email
    message.attach(MIMEText(body, 'plain'))

    try:
        # Create SMTP session for sending the email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            # Start TLS for security
            server.starttls()

            # Login to Gmail
            server.login(sender_email, password)

            # Convert the message to a string and send it
            server.sendmail(sender_email, receiver_email, message.as_string())

            print("Email sent successfully")
    except Exception as e:
        print(f"An error occurred while sending the email: {e}")


def send_form_link_to_students(assessment_id, form_url):
    # Fetch students enrolled in the course associated with the assessment
    students = students_collection.find({'course_id': ObjectId(assessment_id)})
    if students:
        for student in students:
            # Send the form link to each student via email or any other communication method
            student_email = student.get('email')
            if student_email:
                send_email(student_email, "Assessment Form", f"Please complete the assessment using the following link: {form_url}")


# Publish assessment route
@app.route('/publish_assessment/<assessment_id>', methods=['POST'])
def publish_assessment(assessment_id):
    # Generate Google Form based on assessment questions
    form_url = generate_google_form(assessment_id)

    if form_url:
        # Send form link to students
        send_form_link_to_students(assessment_id, form_url)

        # Update assessment status or any other necessary actions
        questions_collection.update_one({'_id': ObjectId(assessment_id)}, {'$set': {'published': True}})

        flash('Assessment published successfully')
    else:
        flash('Failed to publish assessment')

    return redirect(url_for('course_dashboard', course_id=assessment_id))


@app.route('/logout')
def logout():
    for key in list(session.keys()):
        session.pop(key)
    return redirect('/')


if __name__ == "__main__":
    app.run(debug=True)
