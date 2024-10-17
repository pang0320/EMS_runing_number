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

def calculate_ems_range(start: int, end: int):
    """คำนวณตัวเลข EMS พร้อมเลขตรวจสอบสำหรับช่วงหมายเลข"""
    for ems_num in range(start, end + 1):
        check_digit = get_check_digit(ems_num)
        ems_full = f"RK{ems_num:08}{check_digit}TH"  # เพิ่มวงเล็บและอัญประกาศ
        print(ems_full)

# ตัวอย่างการใช้งาน
start_ems = input("กรอกหมายเลข EMS เริ่มต้น (8 หลัก): ")
end_ems = input("กรอกหมายเลข EMS สิ้นสุด (8 หลัก): ")

# ตรวจสอบว่าหมายเลข EMS ที่กรอกถูกต้องหรือไม่
if start_ems.isdigit() and end_ems.isdigit() and len(start_ems) == 8 and len(end_ems) == 8:
    calculate_ems_range(int(start_ems), int(end_ems))
else:
    print("กรุณากรอกหมายเลข EMS ที่ถูกต้อง (8 หลัก).")
