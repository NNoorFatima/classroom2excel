from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import json

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

def main():
    service = authenticate()

    # Get list of courses
    courses = service.courses().list().execute().get('courses', [])
    if not courses:
        print("No courses found.")
        return

    print("\nAvailable Courses:")
    selected_course = select_from_list(courses, 'name')
    course_id = selected_course['id']

    # Get coursework (assignments)
    coursework = service.courses().courseWork().list(courseId=course_id).execute().get('courseWork', [])
    if not coursework:
        print("No assignments found.")
        return

    print("\nAvailable Assignments:")
    selected_work = select_from_list(coursework, 'title')
    coursework_id = selected_work['id']

    # Get student submissions
    submissions = service.courses().courseWork().studentSubmissions().list(
        courseId=course_id, courseWorkId=coursework_id).execute().get('studentSubmissions', [])

    gcr_marks = {}
    for sub in submissions:
        # grade = sub.get('assignedGrade')
        grade = sub.get('assignedGrade') or sub.get('draftGrade')

        if grade is not None:
            user_id = sub['userId']
            profile = service.userProfiles().get(userId=user_id).execute()
            name = profile.get('name', {}).get('fullName')
            gcr_marks[name] = f"{grade}/100"

    # Save to file
    with open("gcr_marks.txt", "w", encoding='utf-8') as f:
        f.write("gcr_marks = {\n")
        for name, mark in gcr_marks.items():
            f.write(f'    "{name}": "{mark}",\n')
        f.write("}\n")

    print("\n✅ Draft marks saved to gcr_marks.txt")

if __name__ == "__main__":
    main()
