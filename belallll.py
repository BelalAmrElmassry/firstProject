import pandas as pd

class Student:
    # Constructor method to initialize student name and marks
    def __init__(self, name, marks):
        self.name = name
        self.marks = marks

    # Calculate total marks
    def calculate_total(self):
        return sum(self.marks)

    # Calculate average marks
    def calculate_average(self):
        return self.calculate_total() / len(self.marks)

    # Determine grade based on average marks
    def determine_grade(self):
        average = self.calculate_average()
        if average >= 60:
            return 'A'
        elif average >= 50:
            return 'B'
        elif average >= 40:
            return 'C'
        elif average >= 30:
            return 'D'
        else:
            return 'F'

    # Check if the student passed (based on both average and individual subject marks)
    def check_pass(self):
        # Pass if all subjects have marks >= 40 and average is >= 40
        return self.calculate_average() >= 40 and all(mark >= 40 for mark in self.marks)

    # Return a dictionary of the student's details
    def get_summary(self):
        return {
            'Student Name': self.name,
            'Total Marks': self.calculate_total(),
            'Average Marks': self.calculate_average(),
            'Grade': self.determine_grade(),
            'Passed': 'Yes' if self.check_pass() else 'No'
        }

class Classroom:
    def __init__(self, students):
        self.students = students

    # Calculate class average marks
    def class_average(self):
        return sum(student.calculate_average() for student in self.students) / len(self.students)

    # Find the top performer in the class
    def top_performer(self):
        return max(self.students, key=lambda student: student.calculate_total())

    # Find the lowest performer in the class
    def lowest_performer(self):
        return min(self.students, key=lambda student: student.calculate_total())

    # Count how many students passed
    def count_passed_students(self):
        return sum(1 for student in self.students if student.check_pass())

    # Generate a summary of all students
    def generate_student_summaries(self):
        return [student.get_summary() for student in self.students]

# Function to load student data from Excel and create Student objects
def load_student_data_from_excel(file_path):
    students = []
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Iterate through each row in the DataFrame
        for index, row in df.iterrows():
            name = row['Student Name']

            # Convert marks to a list, handling non-numeric and missing values
            marks = []
            for mark in row[3:]:
                try:
                    marks.append(float(mark))  # Ensure marks are numeric
                except ValueError:
                    marks.append(0)  # Default to 0 if data is invalid

            # Add the student to the list
            students.append(Student(name, marks))

    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

    return students

# Function to export the student summaries to a new Excel file
def export_student_summaries_to_excel(students, output_file):
    # Convert list of dictionaries to DataFrame
    summaries = pd.DataFrame(students)
    # Write the DataFrame to an Excel file
    summaries.to_excel(output_file, index=False)
    print(f"Student summaries have been successfully exported to {output_file}")

# Example usage
if __name__== "__main__":
    # Path to the Excel file
    file_path = 'Finalresult.xlsx'

    # Load student data from Excel
    students = load_student_data_from_excel(file_path)

    # Check if any students were loaded
    if students:
        # Create a Classroom object with the list of students
        classroom = Classroom(students)

        # Display details for each student
        for student in students:
            print(f"Student Name: {student.name}")
            print(f"Total Marks: {student.calculate_total()}")
            print(f"Average Marks: {student.calculate_average():.2f}")
            print(f"Grade: {student.determine_grade()}")
            print(f"Passed: {'Yes' if student.check_pass() else 'No'}")
            print("-" * 40)

        # Classroom-level statistics
        print("Classroom Statistics:")
        print(f"Class Average Marks: {classroom.class_average():.2f}")
        print(f"Top Performer: {classroom.top_performer().name}")
        print(f"Lowest Performer: {classroom.lowest_performer().name}")
        print(f"Number of Passed Students: {classroom.count_passed_students()}")
        print("-" * 40)

        # Generate and export student summaries
        student_summaries = classroom.generate_student_summaries()
        export_student_summaries_to_excel(student_summaries, 'ProcessedResults.xlsx')

           
