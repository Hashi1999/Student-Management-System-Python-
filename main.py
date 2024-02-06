# main.py

import tkinter as tk
from tkinter import ttk
import pandas as pd
from student import Student, StudentManager, save_to_excel

class StudentGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Student Management System")

        # Load previous data from the Excel sheet
        self.student_manager = StudentManager()
        self.load_previous_data()

        self.create_widgets()

    def load_previous_data(self):
        try:
            # Attempt to read the previous Excel sheet
            df = pd.read_excel('students.xlsx')
            print("Loaded data from Excel:")
            print(df)

            if not df.empty:
                for index, row in df.iterrows():
                    student = Student(row['ID'], row['First Name'], row['Last Name'])
                    self.student_manager.add_student(student)
            else:
                print("No data found in Excel.")

        except FileNotFoundError:
            # If the file is not found, it's the first run, and there's no need to load previous data
            print("File not found.")

    def save_data_to_excel(self):
        try:
            with pd.ExcelWriter('students.xlsx', mode='w', engine='openpyxl') as writer:
                students_data = self.student_manager.get_all_students()
                df = pd.DataFrame(students_data, columns=['ID', 'First Name', 'Last Name'])
                df.to_excel(writer, index=False)
            print("Data saved to Excel.")
        except Exception as e:
            print(f"Error saving data to Excel: {e}")

    def create_widgets(self):
        # Styling
        style = ttk.Style()
        style.configure('TButton', padding=5, width=20, font=('Helvetica', 12))
        style.configure('TLabel', font=('Helvetica', 12))

        # Buttons
        self.add_button = ttk.Button(self.master, text="Add Student", command=self.show_add_student_dialog)
        self.add_button.grid(row=0, column=0, pady=10)

        self.find_button = ttk.Button(self.master, text="Find Student by ID", command=self.show_find_student_dialog)
        self.find_button.grid(row=1, column=0, pady=10)

        self.display_button = ttk.Button(self.master, text="Display All Students", command=self.display_all_students)
        self.display_button.grid(row=2, column=0, pady=10)

        self.exit_button = ttk.Button(self.master, text="Exit", command=self.exit_application)
        self.exit_button.grid(row=3, column=0, pady=10)

        # Text widget for displaying information
        self.text_display = tk.Text(self.master, height=10, width=50)
        self.text_display.grid(row=4, column=0, pady=10)

    def show_add_student_dialog(self):
        add_dialog = tk.Toplevel(self.master)
        add_dialog.title("Add Student")

        tk.Label(add_dialog, text="Student ID:", font=('Helvetica', 12)).grid(row=0, column=0, padx=10, pady=5)
        student_id_entry = tk.Entry(add_dialog)
        student_id_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(add_dialog, text="First Name:", font=('Helvetica', 12)).grid(row=1, column=0, padx=10, pady=5)
        first_name_entry = tk.Entry(add_dialog)
        first_name_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(add_dialog, text="Last Name:", font=('Helvetica', 12)).grid(row=2, column=0, padx=10, pady=5)
        last_name_entry = tk.Entry(add_dialog)
        last_name_entry.grid(row=2, column=1, padx=10, pady=5)

        add_button = ttk.Button(add_dialog, text="Add", command=lambda: self.add_student(
            student_id_entry.get(), first_name_entry.get(), last_name_entry.get(), add_dialog))
        add_button.grid(row=3, column=0, columnspan=2, pady=10)

    def add_student(self, student_id, first_name, last_name, dialog):
        # Check if the student with the given ID already exists
        existing_student = self.student_manager.find_student_by_id(student_id)
        if existing_student:
            error_message = f"Error: Student with ID {student_id} already exists!"
            self.show_error_dialog(error_message)
        else:
            # If the student does not exist, add them to the manager
            student = Student(student_id, first_name, last_name)
            self.student_manager.add_student(student)
            print("Student added successfully!")
            dialog.destroy()

    def show_error_dialog(self, message):
        error_dialog = tk.Toplevel(self.master)
        error_dialog.title("Error")

        tk.Label(error_dialog, text=message, fg="red", font=('Helvetica', 12)).pack(padx=10, pady=10)

        ok_button = ttk.Button(error_dialog, text="OK", command=error_dialog.destroy)
        ok_button.pack(pady=10)

    def show_find_student_dialog(self):
        find_dialog = tk.Toplevel(self.master)
        find_dialog.title("Find Student by ID")

        tk.Label(find_dialog, text="Student ID:", font=('Helvetica', 12)).grid(row=0, column=0, padx=10, pady=5)
        student_id_entry = tk.Entry(find_dialog)
        student_id_entry.grid(row=0, column=1, padx=10, pady=5)

        find_button = ttk.Button(find_dialog, text="Find", command=lambda: self.find_student(
            student_id_entry.get(), find_dialog))
        find_button.grid(row=1, column=0, columnspan=2, pady=10)

    def find_student(self, student_id, dialog):
        student = self.student_manager.find_student_by_id(student_id)
        if student:
            info = f"Student found: {student.first_name} {student.last_name}"
            self.text_display.delete(1.0, tk.END)  # Clear previous content
            self.text_display.insert(tk.END, info)
        else:
            info = "Student not found."
            self.text_display.delete(1.0, tk.END)  # Clear previous content
            self.text_display.insert(tk.END, info)

        dialog.destroy()

    def display_all_students(self):
        students = self.student_manager.get_all_students()
        if students:
            info = ""
            for student in students:
                info += f"ID: {student[0]}, Name: {student[1]} {student[2]}\n"
            self.text_display.delete(1.0, tk.END)  # Clear previous content
            self.text_display.insert(tk.END, info)
        else:
            info = "No students to display."
            self.text_display.delete(1.0, tk.END)  # Clear previous content
            self.text_display.insert(tk.END, info)

    def exit_application(self):
        # Save the data to the Excel sheet before exiting
        self.save_data_to_excel()
        self.master.destroy()

def main():
    root = tk.Tk()
    app = StudentGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()