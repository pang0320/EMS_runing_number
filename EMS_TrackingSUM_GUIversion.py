import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd

def get_check_digit(num: int) -> int:
    """Calculate the S10 check digit."""
    weights = [8, 6, 4, 2, 3, 5, 9, 7]  # S10 weight
    total_sum = 0
    
    # Loop through the digits and calculate the weighted sum
    for i, digit in enumerate(f"{num:08}"):  # Convert to string to ensure 8 digits
        total_sum += weights[i] * int(digit)
    
    # Calculate the check digit
    check_digit = 11 - (total_sum % 11)
    
    # Handle special cases
    if check_digit == 10:
        check_digit = 0
    elif check_digit == 11:
        check_digit = 5
    
    return check_digit

def calculate_ems_range(prefix: str, start: int, end: int, suffix: str):
    """Generate EMS numbers with check digits for a range of numbers."""
    results = []
    for ems_num in range(start, end + 1):
        check_digit = get_check_digit(ems_num)
        ems_full = f"{prefix.upper()}{ems_num:08}{check_digit}{suffix.upper()}"
        results.append(ems_full)
    return results

def on_submit():
    # Get input from Entry fields
    prefix = entry_prefix.get().upper()
    start_ems = entry_start.get()
    end_ems = entry_end.get()
    suffix = entry_suffix.get().upper()

    # Validate the input
    if (prefix.isalpha() and len(prefix) == 2 and
        suffix.isalpha() and len(suffix) == 2 and
        start_ems.isdigit() and end_ems.isdigit() and
        len(start_ems) == 8 and len(end_ems) == 8):
        
        start_ems_num = int(start_ems)
        end_ems_num = int(end_ems)

        # Generate EMS numbers with check digits
        global results
        results = calculate_ems_range(prefix, start_ems_num, end_ems_num, suffix)
        result_text = "\n".join(results)

        # Display the results in the text box
        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)
    else:
        messagebox.showerror("Error", "Please enter valid input (2 letters and 8-digit numbers).")

def save_to_excel():
    if results:
        # Open a file dialog for saving the file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            # Create a DataFrame from the results
            df = pd.DataFrame(results, columns=["EMS Code"])
            
            # Save to Excel
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "File saved successfully!")
    else:
        messagebox.showwarning("Error", "No results to save.")

def save_to_txt():
    if results:
        # Open a file dialog for saving the file
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            # Write the results to the .txt file
            with open(file_path, 'w') as file:
                file.write("\n".join(results))
            messagebox.showinfo("Success", "TXT file saved successfully!")
    else:
        messagebox.showwarning("Error", "No results to save.")

# Create the main window
root = tk.Tk()
root.title("EMS Number Generation Program")

# Set background color for the window
root.configure(bg="#F0F8FF")

# Create the main frame for layout
main_frame = tk.Frame(root, bg="#F0F8FF")
main_frame.pack(pady=20, padx=20)

# Function to create labels and entry fields
def create_label_entry(frame, label_text):
    label = tk.Label(frame, text=label_text, bg="#F0F8FF", font=("Arial", 12), anchor="w")
    label.grid(sticky="w", padx=5, pady=5)
    entry = tk.Entry(frame, font=("Arial", 12))
    entry.grid(sticky="ew", padx=5, pady=5)
    return entry

# Create Label and Entry for the 2-letter prefix
entry_prefix = create_label_entry(main_frame, "Enter 2-letter prefix (e.g., RK):")

# Create Label and Entry for the starting EMS number
entry_start = create_label_entry(main_frame, "Enter starting EMS number (8 digits):")

# Create Label and Entry for the ending EMS number
entry_end = create_label_entry(main_frame, "Enter ending EMS number (8 digits):")

# Create Label and Entry for the 2-letter suffix
entry_suffix = create_label_entry(main_frame, "Enter 2-letter suffix (e.g., TH):")

# Create buttons for submit and save
button_frame = tk.Frame(main_frame, bg="#F0F8FF")
button_frame.grid(sticky="ew", padx=5, pady=10)

submit_button = tk.Button(button_frame, text="Generate", command=on_submit, bg="#4CAF50", fg="white", font=("Arial", 12), width=15)
submit_button.pack(side=tk.LEFT, padx=10)

save_excel_button = tk.Button(button_frame, text="Save as Excel", command=save_to_excel, bg="#2196F3", fg="white", font=("Arial", 12), width=15)
save_excel_button.pack(side=tk.LEFT, padx=10)

save_txt_button = tk.Button(button_frame, text="Save as TXT", command=save_to_txt, bg="#FF5722", fg="white", font=("Arial", 12), width=15)
save_txt_button.pack(side=tk.LEFT, padx=10)

# Create a text box to display the results
text_result = tk.Text(main_frame, height=10, width=50, font=("Arial", 12))
text_result.grid(sticky="ew", padx=5, pady=10)

# Add copyright notice at the bottom
copyright_label = tk.Label(root, text="© Copyright Boonraksa Wanichyanon. All Rights Reserved.", bg="#F0F8FF", font=("Arial", 10))
copyright_label.pack(pady=10)

# Start the program
root.mainloop()
