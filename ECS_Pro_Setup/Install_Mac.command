#!/bin/bash
# ติดตั้ง ECS Pro บน Mac: ดับเบิลคลิกไฟล์นี้ (ถ้าเปิดไม่ได้ ให้เปิด Terminal แล้วพิมพ์  bash Install_Mac.command)
cd "$(dirname "$0")" || exit 1
SRC="$(pwd)/app"
APPDIR="$HOME/Library/Application Support/ECSPro"
LAUNCHER="$HOME/Desktop/ECS Pro.command"

fail() {
  echo
  echo "[ผิดพลาด] $1"
  echo "ติดตั้งไม่สำเร็จ"
  read -r -p "กด Enter เพื่อปิด"
  exit 1
}

echo "============================================"
echo "  ติดตั้งโปรแกรม ECS Pro"
echo "============================================"
echo
[ -f "$SRC/tongheer_app.pyw" ] || fail "ไม่พบโฟลเดอร์ app ข้างไฟล์ติดตั้ง - ต้องคัดลอกมาทั้งโฟลเดอร์ ECS_Pro_Setup"

echo "[1/4] ตรวจ Python ..."
command -v python3 >/dev/null 2>&1 || fail "เครื่องนี้ยังไม่มี Python 3 - ติดตั้งจาก https://www.python.org/downloads/ ก่อน แล้วรันไฟล์นี้อีกครั้ง"
python3 -c "import tkinter" 2>/dev/null || fail "Python ของเครื่องนี้ไม่มี tkinter - ติดตั้ง Python จาก https://www.python.org/downloads/ แล้วรันไฟล์นี้อีกครั้ง"

echo "[2/4] ติดตั้งไลบรารี (ต้องต่ออินเทอร์เน็ต) ..."
mkdir -p "$APPDIR" || fail "สร้างโฟลเดอร์ $APPDIR ไม่ได้"
python3 -m venv "$APPDIR/venv" || fail "สร้าง venv ไม่สำเร็จ"
"$APPDIR/venv/bin/python" -m pip install --disable-pip-version-check -q openpyxl customtkinter tkinterdnd2 \
  || fail "ติดตั้งไลบรารีไม่สำเร็จ - ตรวจอินเทอร์เน็ตแล้วลองใหม่"

echo "[3/4] คัดลอกโปรแกรมไปที่ $APPDIR ..."
cp -f "$SRC/tongheer_app.pyw" "$SRC/tongheer_invb.py" "$SRC/app.ico" "$APPDIR/" || fail "คัดลอกโปรแกรมไม่สำเร็จ"
[ -f "$SRC/DATA_TONGHEER.xlsx" ] && cp -f "$SRC/DATA_TONGHEER.xlsx" "$APPDIR/"
# ไม่เขียนทับค่าที่โปรแกรมจำไว้ ถ้าเคยติดตั้งแล้ว
[ -f "$APPDIR/tongheer_map.json" ] || { [ -f "$SRC/tongheer_map.json" ] && cp "$SRC/tongheer_map.json" "$APPDIR/"; }

echo "[4/4] สร้างไอคอน \"ECS Pro\" บน Desktop (ลบไอคอน \"Tong Heer INVB\" เดิม) ..."
rm -f "$HOME/Desktop/Tong Heer INVB.command"
cat > "$LAUNCHER" <<EOF
#!/bin/bash
exec "$APPDIR/venv/bin/python" "$APPDIR/tongheer_app.pyw" "\$@"
EOF
chmod +x "$LAUNCHER" || fail "สร้างไอคอนไม่สำเร็จ"

echo
echo "ติดตั้งเสร็จแล้ว - เปิดโปรแกรมจากไอคอน \"ECS Pro\" บน Desktop"
read -r -p "กด Enter เพื่อปิด"
