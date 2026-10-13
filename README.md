# ECS Pro

โปรแกรมแปลง Export Summary Report (.xlsx) ของ Tong Heer Fasteners (Thailand) เป็นไฟล์ `INVB_*.xlsx` สำหรับนำเข้า ECS
(ชื่อเดิม: Tong Heer INVB)

## ติดตั้ง

คัดลอกโฟลเดอร์ `ECS_Pro_Setup` (หรือแตกไฟล์ `ECS_Pro.zip`) ไปที่เครื่องที่ต้องการ แล้ว

- **Windows:** ดับเบิลคลิก `Install.bat` จะได้ไอคอน "ECS Pro" บน Desktop และ Start Menu (ไอคอน "Tong Heer INVB" เดิมจะถูกลบ ค่าที่โปรแกรมจำไว้ยังอยู่) ถอนการติดตั้งด้วย `Uninstall.bat`
- **Mac:** แตกไฟล์ `ECS_Pro.zip` แล้วดับเบิลคลิก `Install_Mac.command` ครั้งเดียว สคริปต์จะเตรียม Python และไลบรารีเอง (ถ้าเครื่องยังไม่มี Python จะดาวน์โหลดมาไว้ในโฟลเดอร์ของโปรแกรม ไม่ต้องใช้รหัสผ่านเครื่อง) สร้างแอป "ECS Pro" ใน `~/Applications` พร้อมไอคอนบน Desktop แล้วเปิดโปรแกรมให้ทันที
  - ถ้า macOS ขึ้นว่าเปิดไม่ได้เพราะไม่รู้จักผู้พัฒนา (เกิดเฉพาะครั้งแรกกับไฟล์ที่ดาวน์โหลดมา): คลิกขวาที่ไฟล์แล้วเลือก Open หรือไปที่ System Settings > Privacy & Security แล้วกด Open Anyway
  - หรือเปิด Terminal พิมพ์ `bash ` (มีเว้นวรรค) แล้วลากไฟล์ `Install_Mac.command` ไปวาง แล้วกด Enter

เปิดจากโฟลเดอร์นี้โดยไม่ติดตั้งได้ด้วย `ECS_Pro.bat` (Windows)

## Permit Issue Date และ PRICE

- ช่อง **Permit Issue Date** (มีปุ่มปฏิทิน 📅) กำหนดวันที่ออกในส่วนที่ 3 ของคอลัมน์ PERMIT1 ถ้าเว้นว่างจะใช้วันที่ปัจจุบัน
- คอลัมน์ **PRICE** (ราคาต่อหน่วย) ใส่ให้เป็นค่าเริ่มต้น เอาติ๊กออกได้ที่มุมล่างซ้าย (แบบคำสั่งใช้ `--no-price`)
- ถ้าวันที่อินวอยในรายงานสลับวัน/เดือน และสลับแล้วสอดคล้องกับ ETD โปรแกรมจะแก้ให้เองและแจ้งในจุดที่ต้องตรวจ

## กฎ MARK1 ตามลูกค้า

ถ้าช่อง Customer ในรายงานตรงกับรายการด้านล่าง โปรแกรมจะใส่ MARK1 ตามนี้แทน Marking ท้ายรายงาน (แก้/เพิ่มได้ที่ `customer_marks` ใน `tongheer_map.json`)

| Customer | MARK1 |
| --- | --- |
| LAWSON PRODUCTS,INC | LAWSON |
| HARDWARE SPECIALTY CO.,INC. | HARDWARE |
