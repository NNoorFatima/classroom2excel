from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import json
import re

SCOPES = [
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.student-submissions.students.readonly',
    'https://www.googleapis.com/auth/classroom.student-submissions.me.readonly',
    'https://www.googleapis.com/auth/classroom.rosters.readonly'
]

def authenticate():
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return build('classroom', 'v1', credentials=creds)

def select_from_list(items, key_name):
    for i, item in enumerate(items):
        print(f"{i+1}. {item[key_name]}")
    choice = int(input("Enter number: ")) - 1
    return items[choice]

def extract_roll_number(name, attachments):
    patterns = [
        r"[A-Z]\d{2}-\d{4}",#I22-1234
        r"\d{2}[A-Z]-\d{4}",#22I-1234
        r"\d{2}[A-Z]\d{4}",#22I1234
        r"[A-Z]\d{6}", #I221234
        r"[A-Z]\d{2}_\d{4}", #I22_1234
        r"\d{2}[A-Z]_\d{4}", #22I_1234
    ]
    clean_name = name.replace(" ", "").upper()
    for pattern in patterns:
        match = re.search(pattern, clean_name)
        if match:
            return match.group()[-4:]
    for filename in attachments:
        filename = filename.upper()
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                return match.group()[-4:]
    return None

def main():
    service = authenticate()

    courses = service.courses().list().execute().get('courses', [])
    if not courses:
        print("No courses found.")
        return

    print("\nAvailable Courses:")
    selected_course = select_from_list(courses, 'name')
    course_id = selected_course['id']

    coursework = service.courses().courseWork().list(courseId=course_id).execute().get('courseWork', [])
    if not coursework:
        print("No assignments found.")
        return

    print("\nAvailable Assignments:")
    selected_work = select_from_list(coursework, 'title')
    coursework_id = selected_work['id']

    submissions = service.courses().courseWork().studentSubmissions().list(
        courseId=course_id, courseWorkId=coursework_id).execute().get('studentSubmissions', [])

    gcr_data = []

    for sub in submissions:
        grade = sub.get('assignedGrade') or sub.get('draftGrade')
        if grade is not None:
            user_id = sub['userId']
            profile = service.userProfiles().get(userId=user_id).execute()
            name = profile.get('name', {}).get('fullName')
            attachments = []

            if 'assignmentSubmission' in sub:
                assignment = sub['assignmentSubmission']
                if 'attachments' in assignment:
                    for att in assignment['attachments']:
                        if 'driveFile' in att and 'title' in att['driveFile']:
                            attachments.append(att['driveFile']['title'])

            roll = extract_roll_number(name, attachments)
            gcr_data.append({
                "name": name,
                "roll_number_last4": roll,
                "marks": f"{grade}/100"
            })

    with open("gcr_marks.txt", "w", encoding='utf-8') as f:
        json.dump(gcr_data, f, indent=4)

    print("\n✅ Draft marks saved to gcr_marks.txt")

if __name__ == "__main__":
    main()
