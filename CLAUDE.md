# Voltage Loss App

แอป Desktop สำหรับคำนวณ P Loss ของมิเตอร์ AMR/AMI ที่มีแรงดัน (V) ผิดปกติ

## เป้าหมายของโปรเจค

สร้าง Desktop App ที่:
- รับ input ไฟล์ Voltage, Current, Power Factor, Normal Voltage
- คำนวณ V_average Phase A, Phase B, Phase C จาก V_normal(Normal Voltage)
- คำนวณหา V_loss จาก logic ถ้า Voltage < (V_average*0.975) ให้ใช้ (V_average*0.975)-Voltage ถ้าไม่ให้เป็น 0 ทำเหมือนกันทั้ง Phase A, Phase B, Phase C
- คำนวณ P_loss Phase A, Phase B, Phase C จาก (V_loss*Current*Power Factor*CT_Ratio)/4000 ทำแบบเดียวกันทั้ง 3 เฟส
- คำนวณ P_loss Total จาก P_loss Phase A + P_loss Phase B + P_loss Phase C

## โครงสร้างโปรเจค


## Tech Stack

| เทคโนโลยี | รายละเอียด |
|-----------|-----------|


## สิ่งที่ทำเสร็จแล้ว


## สิ่งที่ยังต้องทำ


## การรันโปรเจค


## Environment ที่ติดตั้งไว้


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
