# Ceo-PA-GAS

ระบบประเมิน PA แบบ Google Apps Script Web App ที่แยกออกจาก pm-coming โดยไม่ใช้ระบบ Login

## จุดประสงค์
- แจกจ่ายเป็น Web App ได้ง่าย
- เมนูซ้ายเลือกวิทยฐานะ
- กรรมการ 3 คนให้คะแนน PA 2/ส
- ส่วนที่ 1 = 60 คะแนน
- ส่วนที่ 2 = 40 คะแนน
- รวมเป็น PA 3/ส อัตโนมัติ
- เก็บข้อมูลลง Google Sheet
- รองรับพิมพ์ / Save PDF
- Responsive สำหรับมือถือ

## Resource
- Sheet ID: 1q-y4CmaJW5lRzldfpLIm3yGT2HACYJSSu8dim59RLDs
- Drive Folder ID: 1_8W_-1Tl4JGdKnaK_gDB814Ux0bt_Gia
- Apps Script ID: 1oL8ZpFZJoMp0Eel45xBJt48Jj_Ka8TNXlIMi2QviUQWcq8MmL6eRIO9b

## ชีตที่ระบบสร้างอัตโนมัติ
- PA_RECORDS
- PA_SCORES
- PA3
- SETTINGS

## Deployment
```bash
npx @google/clasp push
npx @google/clasp deploy --description "Ceo PA GAS"
```

ตั้ง Web App เป็น Execute as: Me และ Who has access: Anyone

> หมายเหตุ: ระบบไม่มี Login ตามโจทย์ ดังนั้นผู้ที่รู้ URL Web App จะสามารถเข้าหน้าระบบได้ ควรแจก URL เฉพาะกลุ่มเป้าหมาย
