#student.py
class Student:
    def __init__(self, student_id, first_name, last_name):
        self.student_id = student_id
        self.first_name = first_name
        self.last_name = last_name

class StudentManager:
    def __init__(self):
        self.students = []

    def add_student(self, student):
        self.students.append(student)

    def find_student_by_id(self, student_id):
        for student in self.students:
            if student.student_id == student_id:
                return student
        return None

    def get_all_students(self):
        return [(student.student_id, student.first_name, student.last_name) for student in self.students]

def save_to_excel(Students):
    # Assuming this function remains unchanged
        pass