import tkinter as tk
import pandas as pd
from tkinter import simpledialog, messagebox

def calculate_bmi(weight, height):
    return weight / (height ** 2)

def collect_data():
    global user_name, user_age, user_dob, user_number, user_quali, user_blood
    global user_height, user_weight, user_nationality, user_sex, user_password

    # Collect user details
    user_number = simpledialog.askstring("Input", "Please enter your phone number:")
    user_name = simpledialog.askstring("Input", "Please enter your name:")
    user_dob = simpledialog.askstring("Input", "What is your date of birth:")
    user_age = simpledialog.askinteger("Input", "What is your age:")
    user_quali = simpledialog.askstring("Input", "What is your qualification:")
    user_blood = simpledialog.askstring("Input", "What is your blood type:")
    user_height = simpledialog.askfloat("Input", "What is your height (in meters):")
    user_weight = simpledialog.askinteger("Input", "What is your weight (in kilograms):")
    user_nationality = simpledialog.askstring("Input", "What is your nationality:")
    user_sex = simpledialog.askstring("Input", "What is your gender:")
    user_password = simpledialog.askstring("Input", "Please create a password:", show='*')

    # Calculate BMI
    BMI = calculate_bmi(user_weight, user_height)
    if BMI < 18.5:
        bmi_status = "Underweight"
    elif 18.5 <= BMI < 25:
        bmi_status = "Normal"
    elif 25 <= BMI < 30:
        bmi_status = "Overweight"
    else:
        bmi_status = "Obese"

    # Display BMI status
    messagebox.showinfo("BMI Status", f"Your BMI is {BMI:.2f} ({bmi_status})")

    # Save data to DataFrame
    data = {
        'Name': [user_name],
        'Age': [user_age],
        'Date of Birth': [user_dob],
        'Phone Number': [user_number],
        'Qualification': [user_quali],
        'Blood Type': [user_blood],
        'Height': [user_height],
        'Weight': [user_weight],
        'BMI': [BMI],
        'Nationality': [user_nationality],
        'Gender': [user_sex]
    }

    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    excel_file = 'bio_data.xlsx'
    df.to_excel(excel_file, index=False)
    print(f"Data has been saved to {excel_file}")

# Create the main window
root = tk.Tk()
root.title("Bio Data Collector")
root.configure(bg="cyan")

# Welcome message
welcome_label = tk.Label(root, text="Hi, Mister! This is the bio data collector.", fg="blue", bg="white", font=("Helvetica", 16))
welcome_label.pack(pady=10)

# Start message
start_label = tk.Label(root, text="I will collect a few of your details. Let's start!", fg="green", bg="white", font=("Helvetica", 14))
start_label.pack(pady=10)

# Start button to collect data
start_button = tk.Button(root, text="Start", command=collect_data, bg="white", font=("Helvetica", 12))
start_button.pack(pady=20)

# Run the application
root.mainloop()
