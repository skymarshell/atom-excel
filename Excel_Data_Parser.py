import pandas as pd
import re
import datetime
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import subprocess
import sys




# Function to ensure dependencies are installed
def install_dependencies():
    required_libraries = ["pandas", "openpyxl"]
    for lib in required_libraries:
        try:
            __import__(lib)
        except ImportError:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Error", f"ไม่สามารถติดตั้ง {lib}: {str(e)}")
                sys.exit(1)

# Ensure required libraries are installed
install_dependencies()

# ฟังก์ชันสำหรับประมวลผลข้อมูลและบันทึกเป็น Excel
def process_and_save():
    raw_text = text_area.get("1.0", tk.END).strip()
    filename = entry_name.get().strip()
    
    if not raw_text:
        messagebox.showerror("Error", "กรุณากรอกข้อมูล Raw Data")
        return
    
    if not filename:
        messagebox.showerror("Error", "กรุณากรอกชื่อไฟล์")
        return
    
    if re.search(r'[\\/:*?"<>|!@#$%^&*()_+]', filename):
        messagebox.showerror("Error", "ชื่อไฟล์มีอักขระต้องห้าม")
        return
    
    pattern = r"(\d{2}/\d{2}/\d{2}) (\d{2}:\d{2}) (X\d) (\w+) ([\d,]+\.\d{2}) ([\d,]+\.\d{2}) (.+)"
    matches = re.findall(pattern, raw_text)
    
    if not matches:
        messagebox.showerror("Error", "ข้อมูลไม่อยู่ในรูปแบบที่ถูกต้อง")
        return
    
    data = []
    for m in matches:
        try:
            date, time, _, _, amount, _, description = m
            money = float(amount.replace(",", ""))
            recive = money if "transfer" in description.lower() else 0
            expense_loq_1000 = money if money <= 1000 else 0
            expense_g_1000 = money if money > 1000 else 0
            data.append([date, time, description, recive, expense_loq_1000, expense_g_1000])
        except Exception as e:
            print(f"Error processing line: {m}, Error: {e}")
    
    df = pd.DataFrame(data, columns=["Date", "Time", "Description", "เงินเข้า", "เงินออก <= 1000", "เงินออก > 1000"])
    df['Time'] = pd.to_datetime(df['Time'], format='%H:%M').dt.time
    timestamp = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    output_filename = f"{filename}-{timestamp}.xlsx"
    
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=output_filename)
    if file_path:
        df.to_excel(file_path, index=False, engine="openpyxl")
        messagebox.showinfo("Success", f"บันทึกไฟล์สำเร็จ: {file_path}")

# สร้าง GUI
top = tk.Tk()
top.title("Excel Data Parser")
top.geometry("1200x800")


tk.Label(top, text="Raw Data:").pack(anchor='w')
text_area = scrolledtext.ScrolledText(top, width=70, height=15) 
text_area.pack(side='top',fill='both',expand=True)

frame = tk.Frame(top)
frame.pack(pady=5,padx=10) 

tk.Label(frame, text="ชื่อไฟล์:").pack()
entry_name = tk.Entry(frame, width=50)
entry_name.pack()
tk.Button(top, text="บันทึกเป็น Excel", command=process_and_save).pack(pady=10)



top.mainloop()
