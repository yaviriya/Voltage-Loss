import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, time, date, timedelta
import openpyxl
from openpyxl.styles import PatternFill
from collections import defaultdict

# ใช้ตำแหน่งของสคริปต์เป็น base path เพื่อให้แอปทำงานได้ทุกเครื่อง
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


class VoltageAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Voltage Loss Calculation")
        self.root.geometry("650x650")

        # ตัวแปรเก็บพาธของไฟล์ (Voltage, Current, Power Factor, Normal Voltage)
        self.input_paths = [None, None, None, None]
        self.multiply_factor_var = tk.StringVar(value="30")  # CT Ratio เริ่มต้น 30

        # ตัวแปรสำหรับเก็บผลรวมของ P Loss แยกตามสี (Peak / Off-Peak / Holiday)
        self.red_total_p_loss = 0
        self.green_total_p_loss = 0
        self.blue_total_p_loss = 0

        # เตรียมสีที่จะใช้
        self.red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')   # สีแดง
        self.green_fill = PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid')  # สีเขียว
        self.blue_fill = PatternFill(start_color='FF00B0F0', end_color='FF00B0F0', fill_type='solid')   # สีฟ้า

        # วันหยุดราชการ
        self.holidays = {
            date(2025, 1, 1),
            date(2025, 2, 12),
            date(2025, 4, 7),
            date(2025, 4, 14),
            date(2025, 4, 15),
            date(2025, 4, 16),
            date(2025, 5, 1),
            date(2025, 5, 5),
            date(2025, 5, 9),
            date(2025, 5, 12),
            date(2025, 6, 2),
            date(2025, 6, 3),
            date(2025, 7, 10),
            date(2025, 7, 11),
            date(2025, 7, 28),
            date(2025, 8, 11),
            date(2025, 8, 12),
            date(2025, 10, 13),
            date(2025, 10, 23),
            date(2025, 12, 5),
            date(2025, 12, 10),
            date(2025, 12, 31),
            # ใส่วันอื่น ๆ ตามต้องการ
        }

        self.create_widgets()

    def create_widgets(self):
        """สร้าง UI elements"""
        # สร้างเฟรมสำหรับเลือกไฟล์
        file_frame = ttk.LabelFrame(self.root, text="Upload Excel Files")
        file_frame.pack(fill="x", padx=10, pady=10)

        # สร้างปุ่มและแสดงพาธของไฟล์ที่เลือก (4 ไฟล์)
        file_labels = ["Voltage", "Current", "Power Factor", "Normal Voltage"]
        self.file_path_vars = [tk.StringVar() for _ in range(4)]

        for i, label in enumerate(file_labels):
            ttk.Label(file_frame, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=5)
            ttk.Entry(file_frame, textvariable=self.file_path_vars[i], width=50).grid(row=i, column=1, padx=5, pady=5)
            ttk.Button(file_frame, text="Browse...", command=lambda idx=i: self.browse_file(idx)).grid(row=i, column=2, padx=5, pady=5)

        # สร้างเฟรมสำหรับตั้งค่าพารามิเตอร์
        param_frame = ttk.LabelFrame(self.root, text="CT")
        param_frame.pack(fill="x", padx=10, pady=10)

        # เพิ่มช่องสำหรับกรอกค่า CT Ratio (multiply_factor)
        ttk.Label(param_frame, text="Ratio:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(param_frame, textvariable=self.multiply_factor_var, width=10).grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # สร้างปุ่มเริ่มการคำนวณ
        ttk.Button(self.root, text="Execute", command=self.process_files).pack(pady=20)

        # สร้างเฟรมสำหรับแสดงผลลัพธ์
        result_frame = ttk.LabelFrame(self.root, text="Results")
        result_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # สร้างตัวแปรสำหรับแสดงผลลัพธ์
        self.output_path_var = tk.StringVar()
        self.red_p_loss_var = tk.StringVar()
        self.green_p_loss_var = tk.StringVar()
        self.blue_p_loss_var = tk.StringVar()
        self.total_p_loss_var = tk.StringVar()

        # แสดงพาธของไฟล์ผลลัพธ์
        ttk.Label(result_frame, text="Output File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(result_frame, textvariable=self.output_path_var, width=70, state="readonly").grid(row=0, column=1, padx=5, pady=5)

        # แสดงผลรวมของ P Loss แยกตามสี
        ttk.Label(result_frame, text="P Loss Peak:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(result_frame, textvariable=self.red_p_loss_var, width=15, state="readonly").grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(result_frame, text="kWh").grid(row=1, column=1, sticky="w", padx=(120, 5), pady=5)

        ttk.Label(result_frame, text="P Loss Off-Peak:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(result_frame, textvariable=self.green_p_loss_var, width=15, state="readonly").grid(row=2, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(result_frame, text="kWh").grid(row=2, column=1, sticky="w", padx=(120, 5), pady=5)

        ttk.Label(result_frame, text="P Loss Holiday:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(result_frame, textvariable=self.blue_p_loss_var, width=15, state="readonly").grid(row=3, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(result_frame, text="kWh").grid(row=3, column=1, sticky="w", padx=(120, 5), pady=5)

        ttk.Label(result_frame, text="Total P Loss:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(result_frame, textvariable=self.total_p_loss_var, width=15, state="readonly").grid(row=4, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(result_frame, text="kWh").grid(row=4, column=1, sticky="w", padx=(120, 5), pady=5)

        # สร้างปุ่มออกจากโปรแกรม
        ttk.Button(self.root, text="Exit", command=self.root.destroy).pack(pady=10)

    def browse_file(self, index):
        """เปิดหน้าต่างให้เลือกไฟล์ Excel"""
        file_path = filedialog.askopenfilename(
            title=f"Select Excel File {index+1}",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_paths[index] = file_path
            self.file_path_vars[index].set(file_path)

    def is_weekend_or_holiday(self, dt):
        """ตรวจสอบว่าเป็นวันเสาร์-อาทิตย์ หรือวันหยุดราชการ"""
        if dt.date() in self.holidays:
            return True
        if dt.weekday() >= 5:  # 5=Saturday, 6=Sunday
            return True
        return False

    def find_header_row(self, worksheet):
        """หาแถวที่เป็นหัวตาราง โดยหาแถวแรกที่คอลัมน์ A มีรูปแบบวันที่"""
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=1), start=1):
            cell_a_value = str(row[0].value).strip() if row[0].value else ""
            if self.is_valid_datetime(cell_a_value):
                if row_idx > 1:
                    return row_idx - 1
                else:
                    return row_idx
        return None

    def is_valid_datetime(self, text):
        """ตรวจสอบว่าข้อความเป็นวันที่หรือไม่"""
        if not text:
            return False
        return any(c.isdigit() for c in text) and ('/' in text or '-' in text)

    def read_excel_data(self, file_path):
        """อ่านข้อมูลจากไฟล์ Excel"""
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        header_row = self.find_header_row(ws)

        if header_row is None:
            messagebox.showwarning("Warning", f"ไม่พบแถวหัวตารางในไฟล์ {os.path.basename(file_path)}")
            return None

        data = {}
        is_thai_date = False

        # ตรวจสอบแถวแรกที่มีข้อมูลวันที่เพื่อดูว่าเป็นวันที่ไทยหรือไม่
        for row in ws.iter_rows(min_row=header_row + 1):
            raw_value = str(row[0].value) if row[0].value is not None else ""
            if self.is_valid_datetime(raw_value):
                clean_value = raw_value.replace('\xa0', '').strip()
                if ' ' in clean_value:
                    parts = clean_value.split()
                    date_str = parts[0]
                else:
                    date_str = clean_value

                if '/' in date_str:
                    day, month, year = date_str.split('/')
                elif '-' in date_str:
                    day, month, year = date_str.split('-')
                else:
                    continue

                try:
                    year_int = int(year)
                    if year_int - 543 > 2000:
                        is_thai_date = True
                except ValueError:
                    pass
                break

        # อ่านข้อมูลทั้งหมด
        for row in ws.iter_rows(min_row=header_row + 1):
            try:
                raw_value = str(row[0].value) if row[0].value is not None else ""
                if not self.is_valid_datetime(raw_value):
                    continue

                clean_value = raw_value.replace('\xa0', '').strip()

                if ' ' in clean_value:
                    parts = clean_value.split()
                    date_str = parts[0]
                    time_str = parts[1]
                else:
                    if '/' in clean_value:
                        date_parts = clean_value.split('/')
                        if len(date_parts) >= 3:
                            date_str = '/'.join(date_parts[:3])
                            time_str = "00:00"
                        else:
                            continue
                    elif '-' in clean_value:
                        date_parts = clean_value.split('-')
                        if len(date_parts) >= 3:
                            date_str = '-'.join(date_parts[:3])
                            time_str = "00:00"
                        else:
                            continue
                    else:
                        continue

                if '/' in date_str:
                    day, month, year = date_str.split('/')
                elif '-' in date_str:
                    day, month, year = date_str.split('-')
                else:
                    continue

                if is_thai_date:
                    year = str(int(year) - 543)

                std_date_str = f"{day}/{month}/{year}"

                if ':' in time_str:
                    time_parts = time_str.split(':')
                    hour = time_parts[0]
                    minute = time_parts[1]
                    if len(minute) > 2:
                        minute = minute[:2]
                    std_time_str = f"{hour}.{minute}"
                else:
                    std_time_str = time_str

                if std_time_str == "24.00" or std_time_str == "24:00" or (hour == "24" if 'hour' in locals() else False):
                    dt_date = datetime.strptime(std_date_str, "%d/%m/%Y")
                    dt = dt_date + timedelta(days=1)
                else:
                    std_date_time_str = f"{std_date_str} {std_time_str}"
                    try:
                        dt = datetime.strptime(std_date_time_str, "%d/%m/%Y %H.%M")
                    except ValueError:
                        if '.' in std_time_str:
                            h, m = std_time_str.split('.')
                            std_time_str = f"{h}:{m}"
                        elif ':' in std_time_str:
                            h, m = std_time_str.split(':')
                            std_time_str = f"{h}.{m}"
                        std_date_time_str = f"{std_date_str} {std_time_str}"
                        try:
                            dt = datetime.strptime(std_date_time_str, "%d/%m/%Y %H.%M")
                        except ValueError:
                            dt = datetime.strptime(std_date_str, "%d/%m/%Y")

                data[dt] = [cell.value for cell in row]
            except Exception:
                continue

        return data, header_row, ws

    def compute_v_average(self, normal_voltage_data):
        """คำนวณค่าเฉลี่ย V Phase A, B, C จากไฟล์ Normal Voltage

        สมมติว่าคอลัมน์ลำดับ 1, 2, 3 ของแต่ละแถวคือ V Phase A, B, C ตามลำดับ
        (ลำดับเดียวกับไฟล์ Voltage) — datetime ของไฟล์นี้ไม่จำเป็นต้องตรงกับ
        ไฟล์อื่น เพราะใช้แค่หาค่าเฉลี่ยของแต่ละคอลัมน์
        """
        sums = [0.0, 0.0, 0.0]
        counts = [0, 0, 0]
        for row in normal_voltage_data.values():
            for i in range(3):
                idx = i + 1
                if idx >= len(row):
                    continue
                val = row[idx]
                if val is None:
                    continue
                try:
                    sums[i] += float(val)
                    counts[i] += 1
                except (ValueError, TypeError):
                    continue
        return [sums[i] / counts[i] if counts[i] > 0 else 0.0 for i in range(3)]

    def process_files(self):
        """ประมวลผลไฟล์ Excel ทั้งหมด"""
        # ตรวจสอบว่าได้เลือกไฟล์ครบทั้ง 4 ไฟล์หรือไม่
        if None in self.input_paths:
            messagebox.showerror("Error", "โปรดเลือกให้ครบทั้ง 4 ไฟล์เพื่อทำการคำนวณ")
            return

        try:
            self.multiply_factor = float(self.multiply_factor_var.get())
        except ValueError:
            messagebox.showerror("Error", "กรุณาป้อนตัวคูณเป็นตัวเลข")
            return

        # รีเซ็ตตัวแปรผลรวม P Loss
        self.red_total_p_loss = 0
        self.green_total_p_loss = 0
        self.blue_total_p_loss = 0

        # อ่านข้อมูลจาก 3 ไฟล์หลัก (Voltage, Current, Power Factor)
        all_data = []
        for path in self.input_paths[:3]:
            result = self.read_excel_data(path)
            if result is None:
                messagebox.showerror("Error", "บางไฟล์ไม่สามารถอ่านได้ กรุณาตรวจสอบไฟล์ของคุณ")
                return
            data, _, _ = result
            all_data.append(data)

        # อ่านไฟล์ Normal Voltage แยก เพื่อคำนวณค่าเฉลี่ย
        normal_result = self.read_excel_data(self.input_paths[3])
        if normal_result is None:
            messagebox.showerror("Error", "ไม่สามารถอ่านไฟล์ Normal Voltage ได้")
            return
        normal_data, _, _ = normal_result
        v_avg_a, v_avg_b, v_avg_c = self.compute_v_average(normal_data)

        if v_avg_a == 0 and v_avg_b == 0 and v_avg_c == 0:
            messagebox.showerror("Error", "ไฟล์ Normal Voltage ไม่มีข้อมูลสำหรับคำนวณค่าเฉลี่ย")
            return

        # threshold = V_avg * 0.975
        threshold_a = v_avg_a * 0.975
        threshold_b = v_avg_b * 0.975
        threshold_c = v_avg_c * 0.975

        # รวมข้อมูลจากทั้ง 3 ไฟล์หลัก
        merged_data = defaultdict(list)
        for dt in sorted(set().union(*[d.keys() for d in all_data])):
            row_data = []
            if dt.hour == 0 and dt.minute == 0:
                prev_day = dt - timedelta(days=1)
                datetime_str = prev_day.strftime("%d/%m/%Y 24.00")
            else:
                datetime_str = dt.strftime("%d/%m/%Y %H.%M")
            row_data.append(datetime_str)

            for data in all_data:
                if dt in data:
                    row_data.extend([val for val in data[dt][1:] if val is not None])

            merged_data[dt] = row_data

        # สร้างไฟล์ Excel ใหม่
        output_wb = openpyxl.Workbook()
        output_ws = output_wb.active

        headers = [
            "DateTime",
            "V Phase A", "V Phase B", "V Phase C",
            "I Phase A", "I Phase B", "I Phase C",
            "Power Factor",
            "V Avg A", "V Avg B", "V Avg C",
            "V Loss A", "V Loss B", "V Loss C",
            "P Loss A", "P Loss B", "P Loss C",
            "P Loss Total",
        ]
        output_ws.append(headers)

        # ตำแหน่งคอลัมน์ที่ต้องใช้ (1-based)
        v_a_col = headers.index("V Phase A") + 1
        v_b_col = headers.index("V Phase B") + 1
        v_c_col = headers.index("V Phase C") + 1
        i_a_col = headers.index("I Phase A") + 1
        i_b_col = headers.index("I Phase B") + 1
        i_c_col = headers.index("I Phase C") + 1
        pf_col = headers.index("Power Factor") + 1
        p_loss_total_col = headers.index("P Loss Total") + 1

        max_input_index = max(v_a_col, v_b_col, v_c_col, i_a_col, i_b_col, i_c_col, pf_col) - 1

        for dt, row_data in sorted(merged_data.items()):
            # แปลงค่า Power Factor ให้เป็นค่าบวก
            if len(row_data) > pf_col - 1:
                pf_value = row_data[pf_col - 1]
                if pf_value is not None:
                    try:
                        pf_float = float(pf_value)
                        if pf_float < 0:
                            row_data[pf_col - 1] = abs(pf_float)
                    except (ValueError, TypeError):
                        pass

            # ดึงค่าจากแถว (None ถ้าไม่ครบ)
            def get_float(idx):
                if len(row_data) > idx and row_data[idx] is not None:
                    try:
                        return float(row_data[idx])
                    except (ValueError, TypeError):
                        return None
                return None

            v_a = get_float(v_a_col - 1)
            v_b = get_float(v_b_col - 1)
            v_c = get_float(v_c_col - 1)
            i_a = get_float(i_a_col - 1)
            i_b = get_float(i_b_col - 1)
            i_c = get_float(i_c_col - 1)
            pf = get_float(pf_col - 1)

            # V Loss ต่อเฟส: ถ้า V < V_avg*0.975 → V_avg*0.975 - V, ไม่งั้น 0
            def calc_v_loss(v, threshold):
                if v is None:
                    return 0.0
                return threshold - v if v < threshold else 0.0

            v_loss_a = calc_v_loss(v_a, threshold_a)
            v_loss_b = calc_v_loss(v_b, threshold_b)
            v_loss_c = calc_v_loss(v_c, threshold_c)

            # P Loss ต่อเฟส: V_loss * I * PF * CT_Ratio / 4000
            def calc_p_loss(v_loss, i_val):
                if i_val is None or pf is None:
                    return 0.0
                return (v_loss * i_val * pf * self.multiply_factor) / 4000

            p_loss_a = calc_p_loss(v_loss_a, i_a)
            p_loss_b = calc_p_loss(v_loss_b, i_b)
            p_loss_c = calc_p_loss(v_loss_c, i_c)
            p_loss_total = p_loss_a + p_loss_b + p_loss_c

            # ขยาย row_data ให้ครบ schema (ถ้าข้อมูลขาด ให้เติม None)
            while len(row_data) <= max_input_index:
                row_data.append(None)

            # ต่อท้าย: V Avg A/B/C, V Loss A/B/C, P Loss A/B/C, P Loss Total
            row_data.extend([
                v_avg_a, v_avg_b, v_avg_c,
                v_loss_a, v_loss_b, v_loss_c,
                p_loss_a, p_loss_b, p_loss_c,
                p_loss_total,
            ])

            output_ws.append(row_data)

            # ใส่สีในคอลัมน์ P Loss Total ตามช่วงเวลา (Peak / Off-Peak / Holiday)
            cell = output_ws.cell(row=output_ws.max_row, column=p_loss_total_col)
            is_24_00 = dt.hour == 0 and dt.minute == 0

            if is_24_00:
                prev_day = dt - timedelta(days=1)
                if self.is_weekend_or_holiday(prev_day):
                    cell.fill = self.blue_fill
                    self.blue_total_p_loss += p_loss_total
                else:
                    cell.fill = self.green_fill
                    self.green_total_p_loss += p_loss_total
            else:
                if self.is_weekend_or_holiday(dt):
                    cell.fill = self.blue_fill
                    self.blue_total_p_loss += p_loss_total
                else:
                    t = dt.time()
                    if time(9, 15) <= t <= time(22, 0):
                        cell.fill = self.red_fill
                        self.red_total_p_loss += p_loss_total
                    else:
                        cell.fill = self.green_fill
                        self.green_total_p_loss += p_loss_total

        # บันทึกไฟล์ — ใช้ path สัมพัทธ์กับสคริปต์ เพื่อให้รันได้ทุกเครื่อง
        x = datetime.now()
        day = x.strftime('%d')
        month = x.strftime('%m')
        year = int(x.strftime('%y')) + 43
        hours = x.strftime('%H')
        minutes = x.strftime('%M')
        date_now_be = f"{day}_{month}_{year}_{hours}{minutes}"

        output_dir = os.path.join(SCRIPT_DIR, 'files', 'output')
        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(output_dir, f'volt_output_data_{date_now_be}.xlsx')
        try:
            output_wb.save(output_path)
        except PermissionError:
            messagebox.showerror("Error", f"ไม่สามารถบันทึกไฟล์ได้ — โปรดปิดไฟล์ที่เปิดอยู่แล้วลองอีกครั้ง:\n{output_path}")
            return

        # แสดงผลลัพธ์ใน UI
        self.output_path_var.set(output_path)
        self.red_p_loss_var.set(f"{self.red_total_p_loss:.4f}")
        self.green_p_loss_var.set(f"{self.green_total_p_loss:.4f}")
        self.blue_p_loss_var.set(f"{self.blue_total_p_loss:.4f}")
        self.total_p_loss_var.set(f"{self.red_total_p_loss + self.green_total_p_loss + self.blue_total_p_loss:.4f}")

        messagebox.showinfo("Success", f"ดำเนินการเสร็จสิ้น\nบันทึกไฟล์ไว้ที่:\n{output_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = VoltageAnalyzerApp(root)
    root.mainloop()
