import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd

def get_check_digit(num: int) -> int:
    """คำนวณ S10 check digit."""
    weights = [8, 6, 4, 2, 3, 5, 9, 7]  # น้ำหนักตามหลัก S10
    total_sum = 0
    
    # วนลูปเพื่อคำนวณตามหลักน้ำหนักที่กำหนด
    for i, digit in enumerate(f"{num:08}"):  # แปลงเป็น string เพื่อให้เป็น 8 หลัก
        total_sum += weights[i] * int(digit)
    
    # คำนวณตัวเลขสุดท้าย
    check_digit = 11 - (total_sum % 11)
    
    # ตรวจสอบเงื่อนไขพิเศษ
    if check_digit == 10:
        check_digit = 0
    elif check_digit == 11:
        check_digit = 5
    
    return check_digit

def calculate_ems_range(prefix: str, start: int, end: int, suffix: str):
    """คำนวณตัวเลข EMS พร้อมเลขตรวจสอบสำหรับช่วงหมายเลข"""
    results = []
    for ems_num in range(start, end + 1):
        check_digit = get_check_digit(ems_num)
        ems_full = f"{prefix.upper()}{ems_num:08}{check_digit}{suffix.upper()}"
        results.append(ems_full)
    return results

def on_submit():
    # รับค่าจาก Entry
    prefix = entry_prefix.get().upper()
    start_ems = entry_start.get()
    end_ems = entry_end.get()
    suffix = entry_suffix.get().upper()

    # ตรวจสอบว่า input ถูกต้องหรือไม่
    if (prefix.isalpha() and len(prefix) == 2 and
        suffix.isalpha() and len(suffix) == 2 and
        start_ems.isdigit() and end_ems.isdigit() and
        len(start_ems) == 8 and len(end_ems) == 8):
        
        start_ems_num = int(start_ems)
        end_ems_num = int(end_ems)

        # คำนวณหมายเลข EMS พร้อมกับเลขตรวจสอบ
        global results
        results = calculate_ems_range(prefix, start_ems_num, end_ems_num, suffix)
        result_text = "\n".join(results)

        # แสดงผลลัพธ์ในกล่องข้อความ
        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)
    else:
        messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกข้อมูลให้ถูกต้อง (ตัวอักษร 2 ตัวและเลข 8 หลัก).")

def save_to_excel():
    if results:
        # เปิดกล่อง dialog ให้ผู้ใช้เลือกที่บันทึกไฟล์
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            # สร้าง DataFrame จากผลลัพธ์
            df = pd.DataFrame(results, columns=["EMS Code"])
            
            # บันทึกลงไฟล์ Excel
            df.to_excel(file_path, index=False)
            messagebox.showinfo("สำเร็จ", "บันทึกไฟล์สำเร็จแล้ว!")
    else:
        messagebox.showwarning("ข้อผิดพลาด", "ไม่มีผลลัพธ์ให้บันทึก")

def save_to_txt():
    if results:
        # เปิดกล่อง dialog ให้ผู้ใช้เลือกที่บันทึกไฟล์
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            # เขียนผลลัพธ์ลงในไฟล์ .txt
            with open(file_path, 'w') as file:
                file.write("\n".join(results))
            messagebox.showinfo("สำเร็จ", "บันทึกไฟล์ TXT สำเร็จแล้ว!")
    else:
        messagebox.showwarning("ข้อผิดพลาด", "ไม่มีผลลัพธ์ให้บันทึก")

# สร้างหน้าต่างหลัก
root = tk.Tk()
root.title("โปรแกรมคำนวณเลข EMS พร้อมเลขตรวจสอบ")

# เพิ่มสีพื้นหลังให้กับหน้าต่าง
root.configure(bg="#F0F8FF")

# สร้างกรอบหลักสำหรับจัดวางองค์ประกอบ
main_frame = tk.Frame(root, bg="#F0F8FF")
main_frame.pack(pady=20, padx=20)

# ฟังก์ชันสำหรับจัดวาง Label และ Entry ในลักษณะเดียวกัน
def create_label_entry(frame, label_text):
    label = tk.Label(frame, text=label_text, bg="#F0F8FF", font=("Arial", 12), anchor="w")
    label.grid(sticky="w", padx=5, pady=5)
    entry = tk.Entry(frame, font=("Arial", 12))
    entry.grid(sticky="ew", padx=5, pady=5)
    return entry

# สร้าง Label และ Entry สำหรับรับตัวอักษร 2 ตัวแรก
entry_prefix = create_label_entry(main_frame, "กรอกตัวอักษร 2 ตัวแรก (เช่น RK):")

# สร้าง Label และ Entry สำหรับรับหมายเลข EMS เริ่มต้น
entry_start = create_label_entry(main_frame, "กรอกหมายเลข EMS เริ่มต้น (8 หลัก):")

# สร้าง Label และ Entry สำหรับรับหมายเลข EMS สิ้นสุด
entry_end = create_label_entry(main_frame, "กรอกหมายเลข EMS สิ้นสุด (8 หลัก):")

# สร้าง Label และ Entry สำหรับรับตัวอักษร 2 ตัวท้าย
entry_suffix = create_label_entry(main_frame, "กรอกตัวอักษร 2 ตัวท้าย (เช่น TH):")

# สร้างปุ่ม Submit และปุ่มบันทึก
button_frame = tk.Frame(main_frame, bg="#F0F8FF")
button_frame.grid(sticky="ew", padx=5, pady=10)

submit_button = tk.Button(button_frame, text="คำนวณ", command=on_submit, bg="#4CAF50", fg="white", font=("Arial", 12), width=15)
submit_button.pack(side=tk.LEFT, padx=10)

save_excel_button = tk.Button(button_frame, text="บันทึกเป็น Excel", command=save_to_excel, bg="#2196F3", fg="white", font=("Arial", 12), width=15)
save_excel_button.pack(side=tk.LEFT, padx=10)

save_txt_button = tk.Button(button_frame, text="บันทึกเป็น TXT", command=save_to_txt, bg="#FF5722", fg="white", font=("Arial", 12), width=15)
save_txt_button.pack(side=tk.LEFT, padx=10)

# สร้างกล่องข้อความสำหรับแสดงผลลัพธ์
text_result = tk.Text(main_frame, height=10, width=50, font=("Arial", 12))
text_result.grid(sticky="ew", padx=5, pady=10)

# เริ่มต้นโปรแกรม
root.mainloop()
