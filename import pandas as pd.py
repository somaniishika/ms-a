import pandas as pd
import cv2
from pyzbar.pyzbar import decode
import openpyxl
from datetime import datetime

input_file = r"D:\Mini Project\Final Saturday\output\student_enrollment.xlsx"
output_file = r"D:\Mini Project\Final Saturday\output\attendance_record.xlsx"
user_accounts_file = r"D:\Mini Project\Final Saturday\output\user_accounts.xlsx"

# Step 1: Load admin username and password
admin_username = "admin"
admin_password = "admin"

# Step 2: Define function to load user accounts
def load_user_accounts():
    try:
        return pd.read_excel(user_accounts_file)
    except FileNotFoundError:
        # Create an empty DataFrame if the file doesn't exist
        return pd.DataFrame(columns=["Username", "Password"])

# Step 3: Define function to save user accounts
def save_user_accounts(accounts_df):
    accounts_df.to_excel(user_accounts_file, index=False)

# Step 4: Define function for user login
def user_login(username, password):
    if username == admin_username and password == admin_password:
        return True, "admin"
    else:
        accounts_df = load_user_accounts()
        if not accounts_df.empty:
            user_data = accounts_df.loc[accounts_df['Username'] == username]
            if not user_data.empty and user_data.iloc[0]['Password'] == password:
                return True, username
    return False, None

# Step 5: Define function for user signup
def user_signup(username, password):
    accounts_df = load_user_accounts()
    if username not in accounts_df['Username'].values:
        accounts_df = accounts_df._append({'Username': username, 'Password': password}, ignore_index=True)
        save_user_accounts(accounts_df)
        return True
    return False

# Step 6: Define function to mark attendance
def mark_attendance(enrollment_number):
    wb = openpyxl.load_workbook(output_file)
    sheet = wb.active
    marked = False
    for row in sheet.iter_rows(min_row=2):
        if str(row[0].value).strip().upper() == enrollment_number.strip().upper():
            row[-1].value = "Present"
            marked = True
            print(f"{enrollment_number} is marked present")
            break
    if not marked:
        print(f"Enrollment number {enrollment_number} not found")
    wb.save(output_file)

# Step 7: Prompt for login or signup
while True:
    choice = input("Are you an existing user? (yes/no): ").lower()
    if choice == 'yes':
        username = input("Enter your username: ")
        password = input("Enter your password: ")
        logged_in, user_type = user_login(username, password)
        if logged_in:
            print(f"Welcome, {username}!")
            if user_type == "admin":
                print("You are logged in as admin.")
            else:
                print("You are logged in as a teacher.")
            break
        else:
            print("Invalid username or password. Please try again.")
    elif choice == 'no':
        admin_access = input("Do you have admin access? (yes/no): ").lower()
        if admin_access == 'yes':
            admin_username_input = input("Enter admin username: ")
            admin_password_input = input("Enter admin password: ")
            logged_in, user_type = user_login(admin_username_input, admin_password_input)
            if logged_in and user_type == "admin":
                new_username = input("Enter new username: ")
                new_password = input("Enter new password: ")
                if user_signup(new_username, new_password):
                    print("Account created successfully!")
                    print("You can now log in.")
                else:
                    print("Username already exists. Please choose a different username.")
                break
            else:
                print("Invalid admin credentials. Access denied.")
        else:
            print("You need admin access to create a new account.")
    else:
        print("Invalid choice. Please enter 'yes' or 'no'.")

# Step 8: Continue with attendance marking if logged in
if logged_in:
    # Load the input Excel sheet and set up the output Excel sheet
    students_df = pd.read_excel(input_file)
    today = datetime.today().strftime('%Y-%m-%d')

    if today not in students_df.columns:
        students_df[today] = "Absent"

    students_df.to_excel(output_file, index=False)
    
    # Select webcam index
    webcam_index = int(input("Enter the index of the webcam you want to use: "))
    
    # Capture and decode QR codes
    cap = cv2.VideoCapture(webcam_index)

    while True:
        ret, frame = cap.read()
        if not ret:
            print("Failed to grab frame")
            break

        for barcode in decode(frame):
            enrollment_number = barcode.data.decode('utf-8')
            print(f"Detected QR code with enrollment number: {enrollment_number}")
            mark_attendance(enrollment_number)

            pts = barcode.polygon
            if len(pts) == 4:
                pts = [pt for pt in pts]
                pts = pts[:4]
            else:
                hull = cv2.convexHull(pts)
                pts = [tuple(pt) for pt in hull]

            n = len(pts)
            for j in range(0, n):
                cv2.line(frame, pts[j], pts[(j + 1) % n], (0, 255, 0), 3)

        cv2.imshow('QR Code Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()


