import openpyxl
import os

FILE_NAME = "Students.xlsx"


# Create Excel File
def create_excel():

    if not os.path.exists(FILE_NAME):

        wb = openpyxl.Workbook()
        ws = wb.active

        ws.title = "Student Records"

        ws.append([
            "ID",
            "Name",
            "Contact",
            "Class",
            "Address"
        ])

        wb.save(FILE_NAME)


# Add New Student
def add_student():

    create_excel()

    print("\n===== ADD NEW STUDENT =====")

    name = input("Enter Student Name: ")
    contact = input("Enter Contact Number: ")
    student_class = input("Enter Class: ")
    address = input("Enter Address: ")

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    # Check duplicate contact number
    for row in ws.iter_rows(min_row=2, values_only=True):

        if str(row[2]) == contact:
            print("\nThis Contact Number Already Exists!")
            wb.close()
            return

    student_id = ws.max_row

    ws.append([
        student_id,
        name,
        contact,
        student_class,
        address
    ])

    wb.save(FILE_NAME)
    wb.close()

    print("\nStudent Record Saved Successfully!")


# View All Students
def view_students():

    create_excel()

    print("\n===== ALL STUDENT RECORDS =====")

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    if ws.max_row == 1:

        print("No Student Records Found!")

    else:

        for row in ws.iter_rows(min_row=2, values_only=True):

            print("--------------------------------")
            print("ID      :", row[0])
            print("Name    :", row[1])
            print("Contact :", row[2])
            print("Class   :", row[3])
            print("Address :", row[4])

    wb.close()


# Delete Student
def delete_student():

    create_excel()

    print("\n===== DELETE STUDENT =====")

    contact = input("Enter Contact Number to Delete: ")

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    found = False

    for row in range(2, ws.max_row + 1):

        if str(ws.cell(row, 3).value) == contact:

            ws.delete_rows(row, 1)
            found = True
            break

    if found:

        wb.save(FILE_NAME)
        print("\nStudent Record Deleted Successfully!")

    else:

        print("\nStudent Record Not Found!")

    wb.close()


# Main Menu
def main():

    create_excel()

    while True:

        print("\n================================")
        print(" STUDENT RECORD MANAGEMENT SYSTEM")
        print("================================")

        print("1. Add New Student")
        print("2. View Students")
        print("3. Delete Student")
        print("4. Exit")

        choice = input("\nEnter Your Choice: ")

        if choice == "1":

            add_student()

        elif choice == "2":

            view_students()

        elif choice == "3":

            delete_student()

        elif choice == "4":

            print("\nThank You! Program Closed.")
            break

        else:

            print("\nInvalid Choice! Try Again.")


main()