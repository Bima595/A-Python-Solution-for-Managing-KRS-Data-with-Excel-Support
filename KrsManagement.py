import tkinter as tk
from openpyxl import Workbook, load_workbook
import time
# Data login (contoh)
user_data = {
    "user123": {"password": "pass123", "ipk": 3.8, "selected_courses": []},
    "22": {"password": "22", "ipk": 3.2, "selected_courses": []},
    # Add more users with their respective passwords and IPK
}

# Data mata kuliah
matakuliah = {
    "Matkul 1": 4,
    "Matkul 2": 3,
    "Matkul 3": 2,
    "Matkul 4": 3,
    "Matkul 5": 4,
    "Matkul 6": 2,
    "Matkul 7": 3,
    "Matkul 8": 2,
    "Matkul 9": 3,
    "Matkul 10": 4,
}

# Data batasan SKS
batasan_sks_ipk_35 = {
    "min_sks": 20,
    "max_sks_high_ipk": 24
}


def load_data_from_excel():
    global user_data, selected_courses, checkbox_vars
    try:
        wb = load_workbook("KrsDatabase.xlsx")
        ws = wb.active

        for row in ws.iter_rows(min_row=2, max_col=13, max_row=ws.max_row):
            username = row[0].value
            password = row[1].value
            stored_selected_courses = [cell.value for cell in row[2:12]]
            total_sks = row[12].value

            # Check if user_data already contains the user
            if username in user_data:
                # Update the selected courses for existing users
                user_data[username]["password"] = password
                user_data[username]["selected_courses"] = stored_selected_courses
            else:
                # Add new users to user_data
                user_data[username] = {"password": password, "ipk": 0, "selected_courses": stored_selected_courses}
                
                # Initialize IntVar objects for new courses
                for course in stored_selected_courses:
                    checkbox_vars[course] = tk.IntVar()

    except FileNotFoundError:
        # Handle the case where the Excel file doesn't exist yet
        pass

def update_status_label():
    global label_status, current_user, user_data, selected_courses
    total_sks = calculate_total_sks()
    status_text = f"Total SKS: {total_sks}"
    label_status.config(text=status_text)


# Initialize root window
root = tk.Tk()

# Global variable for checkbox_vars
checkbox_vars = {}

# Global counter for each user's row in the Excel sheet
user_row_counter = {}

# Global variable for selected_courses
selected_courses = []

# Global variable to track submit status
submit_clicked = False




def login():
    global current_user, entry_username, entry_password, label_status, submit_clicked

    current_user = None
    selected_courses = []  # Reset selected_courses on login
    submit_clicked = False  # Reset submit_clicked on login
    last_submission_time = 0  # Reset last_submission_time on login


    if submit_clicked:
        return  # Prevent login after submit without logout

    username = entry_username.get()
    password = entry_password.get()

    if username in user_data and password == user_data[username]["password"]:
        current_user = username
        show_krs_page(user_data[username]["ipk"])
    else:
        label_status.config(text="Login gagal. Coba lagi.")
        # Clear the entry fields for a new login attempt
        entry_username.delete(0, tk.END)
        entry_password.delete(0, tk.END)

def login_page():
    global entry_username, entry_password, label_status

    # Atur tata letak GUI untuk login
    label_username = tk.Label(root, text="NIM:")
    label_password = tk.Label(root, text="PIC:")
    entry_username = tk.Entry(root)
    entry_password = tk.Entry(root, show="*")  # Show '*' for password
    button_login = tk.Button(root, text="Login", command=login)
    label_status = tk.Label(root, text="")

    label_username.grid(row=0, column=0)
    label_password.grid(row=1, column=0)
    entry_username.grid(row=0, column=1)
    entry_password.grid(row=1, column=1)
    button_login.grid(row=2, column=0, columnspan=2)
    label_status.grid(row=3, column=0, columnspan=2)


row_counter = 3
calculate_total_sks = lambda: sum([matakuliah[course] for course in selected_courses])

def show_krs_page(ipk):
    global label_status, root, checkbox_vars, current_user, user_row_counter, selected_courses, submit_clicked

    # Destroy widgets using grid_forget to preserve grid configuration
    for widget in root.grid_slaves():
        widget.grid_forget()

    label_status = tk.Label(root, text="Total SKS: 0")  # Initial total SKS label
    label_status.grid(row=0, column=0, columnspan=2)

    label_username = tk.Label(root, text=f"NIM: {current_user}")
    label_password = tk.Label(root, text=f"PIC: {user_data[current_user]['password']}")
    label_username.grid(row=1, column=0, columnspan=2)
    label_password.grid(row=2, column=0, columnspan=2)

    selected_courses = user_data[current_user]["selected_courses"]

    update_status_label()

    def get_max_sks_limit():
        global current_user, batasan_sks_ipk_35
        return batasan_sks_ipk_35["max_sks_high_ipk"] if user_data[current_user]["ipk"] > 3.5 else batasan_sks_ipk_35["min_sks"]

    def update_status():
        global selected_courses
        total_sks = calculate_total_sks()
        update_status_label()  # Call the function to update the label

    def calculate_total_sks():
        return sum(matakuliah.get(course, 0) for course in selected_courses)

    def update_selected_courses(course):
        global selected_courses, checkbox_vars

        if checkbox_vars[course].get() == 1:
            if course not in selected_courses and calculate_total_sks() + matakuliah[course] <= get_max_sks_limit():
                selected_courses.append(course)
            else:
                checkbox_vars[course].set(0)  # Uncheck the checkbox if SKS limit is exceeded
        else:
             if course in selected_courses:
                selected_courses.remove(course)
        
        selected_courses = [course for course in selected_courses if course is not None]

        update_status()

    

        


    row_counter = 3  # Start with the next row for checkbuttons

    for mata_kuliah, sks in matakuliah.items():
        var = tk.IntVar()
        # Set the initial state of the checkbox based on user's selection
        initial_state = 1 if mata_kuliah in selected_courses else 0
        checkbutton = tk.Checkbutton(root, text=f"{mata_kuliah} - SKS: {sks}", variable=var,
                                     command=lambda m=mata_kuliah: update_selected_courses(m), onvalue=1, offvalue=0)
        checkbutton.select() if initial_state else checkbutton.deselect()
        checkbutton.grid(row=row_counter, column=0, columnspan=2, sticky='w')
        checkbox_vars[mata_kuliah] = var
        row_counter += 1

    # Submit button with functionality
    button_submit = tk.Button(root, text="Submit", command=submit)
    button_submit.grid(row=row_counter, column=0, columnspan=2)

    # Logout button with functionality
    button_logout = tk.Button(root, text="Logout", command=logout)
    button_logout.grid(row=row_counter, column=2)

last_submission_time = 0

# Time delay between submissions (in seconds)
submission_cooldown = 60  # Adjust this value as needed

def submit():
    global user_row_counter, selected_courses, submit_clicked, last_submission_time

    current_time = time.time()

    # Check if the submission cooldown has passed
    if current_time - last_submission_time < submission_cooldown:
        label_status.config(text=f"Harap tunggu {int(submission_cooldown - (current_time - last_submission_time))} detik sebelum submit lagi.")
        return

    last_submission_time = current_time
    if submit_clicked:
        return  # Prevent multiple submissions without logout

    submit_clicked = True  # Set submit_clicked to True

    # Store selected courses in user_data
    user_data[current_user]["selected_courses"] = selected_courses

    # Load existing workbook or create a new one if it doesn't exist
    try:
        wb = load_workbook("KrsDatabase.xlsx")
    except FileNotFoundError:
        wb = Workbook()

    # Select the sheet for all users or create a new one
    ws = wb.active

    # Find the row with the same username if it exists
    existing_row = None
    for row in ws.iter_rows(min_row=2, max_col=1, max_row=ws.max_row):
        if row[0].value == current_user:
            existing_row = row[0].row
            break

    if existing_row is not None:
        # Update the existing row with the new data
        for i in range(10):
            if i < len(selected_courses):
                ws.cell(row=existing_row, column=i + 3, value=selected_courses[i])
            else:
                ws.cell(row=existing_row, column=i + 3, value="")
        ws.cell(row=existing_row, column=13, value=calculate_total_sks())  # Total SKS
        label_status.config(text="Anda Telah Berhasil Melakukan Krs Database")
    else:
        # If the user's row counter is not set, find the next available row
        if current_user not in user_row_counter:
            user_row_counter[current_user] = ws.max_row + 1

        # Write data to Excel file in the next available row
        data = [current_user, user_data[current_user]["password"]]
        for i in range(10):
            if i < len(selected_courses):
                data.append(selected_courses[i])
            else:
                data.append("")
        data.append(calculate_total_sks())  # Total SKS
        ws.append(data)
        label_status.config(text="Anda Telah Berhasil Melakukan Krs Database")
        user_row_counter[current_user] += 1

    # Save workbook to file
    wb.save("KrsDatabase.xlsx")

def logout():
    global current_user, entry_username, entry_password, label_status, submit_clicked, selected_courses
    current_user = None
    selected_courses = []  # Reset selected_courses on login
    submit_clicked = False  # Reset submit_clicked on login
    last_submission_time = 0  # Reset last_submission_time on login
    for widget in root.grid_slaves():
        widget.grid_forget()
    login_page()

calculate_total_sks = lambda: sum(matakuliah.get(course, 0) for course in selected_courses)

# Run GUI
load_data_from_excel()
login_page()
root.mainloop()