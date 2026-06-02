# Voltage Loss App

แอป Desktop สำหรับคำนวณ P Loss ของมิเตอร์ AMR/AMI ที่มีแรงดัน (V) ผิดปกติ

## เป้าหมายของโปรเจค

สร้าง Desktop App ที่รับ input 5 ไฟล์ — **Voltage, Current, Power Factor, Normal Voltage, Normal Current** — แล้วทำงานตามขั้นตอนดังนี้:

### 1. Fit โมเดล Linear Regression 3 ตัว (จากข้อมูลช่วง Normal)

ฝึกโมเดลด้วยไฟล์ Normal Voltage / Normal Current โดยแต่ละเฟสใช้:

| โมเดล | ตัวแปรตาม (Y) | ตัวแปรต้น (X) |
|-------|---------------|----------------|
| **Phase A** | Voltage Phase A | Voltage B, Voltage C, Current A, Current B, Current C |
| **Phase B** | Voltage Phase B | Voltage A, Voltage C, Current A, Current B, Current C |
| **Phase C** | Voltage Phase C | Voltage A, Voltage B, Current A, Current B, Current C |

> ทุก range ใช้ค่าช่วงปกติ (Normal) ตั้งแต่บรรทัดแรกจนบรรทัดสุดท้าย
> ผลลัพธ์: ได้ Model 3 ตัว (Voltage Phase A, B, C) พร้อมแสดงค่า **R-Square ทั้ง 3 ค่า** (ยิ่งใกล้ 1 ยิ่งดี)

### 2. Predict V_regression ในช่วง Calculation

นำโมเดลทั้ง 3 มาทำนายค่า V_regression Phase A, B, C ในช่วงการคำนวณ โดยป้อน input จากไฟล์ **Voltage, Current** (ช่วง Calculation)

### 3. คำนวณ V_loss และ P_loss

- **V_loss** (ทำเหมือนกันทั้ง 3 เฟส):
  ถ้า `Voltage < V_regression(Calculation) × 0.975` → ใช้ `(V_regression(Calculation) × 0.975) − Voltage`
  ถ้าไม่ → เป็น `0`
- **P_loss** Phase A, B, C = `(V_loss × Current × Power Factor × CT_Ratio) / 4000`
- **CT_Ratio** เลือกจาก Radio Button: `100/5 = 20`, `150/5 = 30`, `250/5 = 50`, `400/5 = 80`
- **P_loss Total** = P_loss Phase A + P_loss Phase B + P_loss Phase C

## โครงสร้างโปรเจค

```
voltage_loss/
├── voltage_loss_app.py     # แอปหลัก (ใช้งานจริง) — GUI + Linear Regression 3 เฟส
├── voltage_loss_gui.py     # เวอร์ชันเก่า (V Diff + เลือกทีละเฟส) — legacy ไม่ใช้แล้ว
├── voltage_loss.py         # เวอร์ชันเก่าแบบ CLI — legacy ไม่ใช้แล้ว
├── logo/wave.ico           # ไอคอนของแอป (ใช้ทั้ง .exe และหน้าต่าง)
├── captureScreen/          # เก็บภาพ error ไว้ให้ผู้ช่วยดู
├── files/output/           # ไฟล์ผลลัพธ์ (ถูก gitignore) — สร้างข้างๆ ตัวแอป
├── dist/VoltageLoss.exe    # ไฟล์ .exe ที่ build แล้ว (gitignore)
└── CLAUDE.md
```

## Tech Stack

| เทคโนโลยี | รายละเอียด |
|-----------|-----------|
| Python 3.12 | ภาษาหลัก |
| tkinter / ttk | สร้างหน้าจอ GUI |
| openpyxl | อ่าน/เขียนไฟล์ Excel (.xlsx) พร้อมใส่สีเซลล์ |
| xlrd | อ่านไฟล์ Excel รูปแบบเก่า (.xls) — เวอร์ชัน 2.x รองรับเฉพาะ .xls |
| numpy | จัดการ array สำหรับ regression |
| scikit-learn | `LinearRegression` + ค่า R-Square |
| PyInstaller | build เป็นไฟล์ `.exe` แจกจ่าย |

## สิ่งที่ทำเสร็จแล้ว

- GUI รับ 5 ไฟล์: แยกกลุ่ม **Normal Files** (Normal Voltage, Normal Current) กับ **Calculation Files** (Voltage, Current, Power Factor)
- รองรับไฟล์ทั้ง `.xlsx` (openpyxl) และ `.xls` (xlrd) ผ่านตัวโหลดกลาง `load_rows()`
- CT Ratio เป็น Radio Button แนวนอน: 100/5→20, 150/5→30, 250/5→50, 400/5→80
- คำนวณ V_regression 3 เฟสด้วย Linear Regression + แสดง R-Square ของแต่ละเฟส
- คำนวณ V_loss, P_loss แยก 3 เฟส และ P_loss Total
- ใส่สีผลลัพธ์ตามช่วงเวลา Peak / Off-Peak / Holiday และสรุปยอดแยกสี
- build เป็น `.exe` ไฟล์เดียว (ฝังไอคอน, แก้ path ให้ output ออกข้างๆ `.exe`)

## สิ่งที่ยังต้องทำ

- (ยังไม่มีรายการค้าง)

## การรันโปรเจค

**รันแบบ dev (มี Python):**
```powershell
python voltage_loss_app.py
```

**build เป็น .exe แจกจ่าย:**
```powershell
pyinstaller --onefile --windowed --name VoltageLoss --icon "logo/wave.ico" --add-data "logo/wave.ico;." `
  --exclude-module torch --exclude-module torchvision --exclude-module torchaudio `
  --exclude-module tensorflow --exclude-module matplotlib --exclude-module pandas `
  --exclude-module cv2 --exclude-module PIL --exclude-module IPython `
  --noconfirm voltage_loss_app.py
```
> ต้อง `--exclude-module` พวก torch/cuda/pandas/matplotlib ที่ติดตั้งในเครื่อง ไม่งั้น PyInstaller จะดูดมาด้วยจนไฟล์บวมถึง ~2.4 GB (ตัดออกแล้วเหลือ ~75 MB)
> ได้ไฟล์ที่ `dist\VoltageLoss.exe` — ส่งไฟล์เดียวไปเปิดเครื่องอื่นได้เลย ไม่ต้องติดตั้ง Python

## Environment ที่ติดตั้งไว้

- Python 3.12 (`C:\Users\amari\AppData\Local\Programs\Python\Python312`)
- numpy 1.26.4, scikit-learn 1.5.2, scipy 1.14.1, openpyxl 3.1.5, xlrd 2.0.2
- PyInstaller 6.12.0
- หมายเหตุ: เครื่องนี้มี torch (cu121), pandas, matplotlib, opencv ติดตั้งอยู่ด้วย แต่แอปนี้ไม่ได้ใช้ (ต้อง exclude ตอน build)

## Core Principles

1. **Never Guess** - อ่านโค้ดก่อนตอบ อย่าเดา
2. **Find Root Cause** - หาสาเหตุที่แท้จริง ไม่ใช่แค่แก้อาการ
3. **Minimize Changes** - ทำเฉพาะที่ขอ ไม่ over-engineer
4. การแก้ไข System.Environment ทุกครั้งให้ทำการแก้ผ่านหน้าจอ GUI (System Properties) เสมอ เพราะการแก้ผ่าน PowerShell ด้วย SetEnvironmentVariable จะเขียนทับ Path เดิมทั้งหมด ทำให้ node, git, ngrok และอื่นๆ หายไปจาก Path
5. ถ้าบอกให้ดู error จากรูปภาพ ให้เข้าไปดูในโฟลเดอร์ `D:\Coding\voltage_loss\captureScreen`

## บุคลิกของผู้ช่วย

ผู้ช่วย AI ในโปรเจคนี้มีนิสัยร่าเริง เป็นกันเอง สนุกสนาน และสุภาพ พูดคุยด้วยความเป็นมิตร ใช้ภาษาที่เข้าใจง่าย และพร้อมช่วยเหลือเสมอด้วยความยินดี

## ผู้พัฒนา

ยะ & Claude
