#!/bin/bash
# ติดตั้ง ECS Pro บน Mac: ดับเบิลคลิกไฟล์นี้ครั้งเดียว จะได้แอป "ECS Pro" บน Desktop พร้อมใช้งาน
# (ถ้า macOS ไม่ยอมเปิด: คลิกขวาที่ไฟล์ > Open หรือเปิด Terminal แล้วพิมพ์  bash  ตามด้วยลากไฟล์นี้ไปวาง แล้วกด Enter)
set -o pipefail
cd "$(dirname "$0")" || exit 1
SETUP="$(pwd)"
SRC="$SETUP/app"
APPDIR="$HOME/Library/Application Support/ECSPro"
VENV="$APPDIR/venv"
APP="$HOME/Applications/ECS Pro.app"
PKGS="openpyxl customtkinter tkinterdnd2"

fail() {
  echo
  echo "[ผิดพลาด] $1"
  echo "ติดตั้งไม่สำเร็จ"
  read -r -p "กด Enter เพื่อปิด"
  exit 1
}

# Python ที่ใช้ได้ต้องมี tkinter (Tk 8.6 ขึ้นไป) - ไม่ใช้ /usr/bin/python3 ของ macOS เพราะ Tk เก่าและเด้งให้ลง Xcode tools
find_python() {
  local p
  for p in /Library/Frameworks/Python.framework/Versions/3.*/bin/python3 /opt/homebrew/bin/python3 /usr/local/bin/python3; do
    [ -x "$p" ] || continue
    if "$p" -c 'import sys, tkinter; sys.exit(0 if sys.version_info >= (3, 9) and tkinter.TkVersion >= 8.6 else 1)' 2>/dev/null; then
      echo "$p"
      return 0
    fi
  done
  return 1
}

echo "============================================"
echo "  ติดตั้งโปรแกรม ECS Pro"
echo "============================================"
echo
[ -f "$SRC/tongheer_app.pyw" ] || fail "ไม่พบโฟลเดอร์ app ข้างไฟล์ติดตั้ง - ต้องแตกไฟล์ zip ออกมาทั้งโฟลเดอร์ ECS_Pro_Setup ก่อน"
# ไฟล์ที่ดาวน์โหลดมาจะถูก macOS กักไว้ (quarantine) - ปลดให้ทั้งโฟลเดอร์ติดตั้ง
xattr -dr com.apple.quarantine "$SETUP" 2>/dev/null
mkdir -p "$APPDIR" || fail "สร้างโฟลเดอร์ $APPDIR ไม่ได้"

echo "[1/4] เตรียม Python ..."
rm -rf "$VENV"
PY="$(find_python)"
if [ -n "$PY" ]; then
  echo "      ใช้ Python ที่มีในเครื่อง: $PY"
  "$PY" -m venv "$VENV" || fail "สร้าง venv ไม่สำเร็จ"
  echo "[2/4] ติดตั้งไลบรารี (ต้องต่ออินเทอร์เน็ต) ..."
  "$VENV/bin/python" -m pip install --disable-pip-version-check -q $PKGS \
    || fail "ติดตั้งไลบรารีไม่สำเร็จ - ตรวจอินเทอร์เน็ตแล้วลองใหม่"
else
  # ไม่มี Python ที่ใช้ได้: ดาวน์โหลด Python 3.12 มาไว้ในโฟลเดอร์ของโปรแกรมเอง (ไม่ต้องใช้รหัสผ่านเครื่อง)
  echo "      ยังไม่มี Python กำลังดาวน์โหลด (ประมาณ 40 MB ต้องต่ออินเทอร์เน็ต) ..."
  export UV_PYTHON_INSTALL_DIR="$APPDIR/python"
  curl -LsSf https://astral.sh/uv/install.sh | env UV_UNMANAGED_INSTALL="$APPDIR/uv" sh >/dev/null 2>&1 \
    || fail "ดาวน์โหลดตัวติดตั้ง Python ไม่สำเร็จ - ตรวจอินเทอร์เน็ตแล้วลองใหม่"
  UV="$APPDIR/uv/uv"
  "$UV" venv -q --python 3.12 "$VENV" || fail "ติดตั้ง Python ไม่สำเร็จ - ตรวจอินเทอร์เน็ตแล้วลองใหม่"
  echo "[2/4] ติดตั้งไลบรารี ..."
  "$UV" pip install -q --python "$VENV/bin/python" $PKGS || fail "ติดตั้งไลบรารีไม่สำเร็จ - ตรวจอินเทอร์เน็ตแล้วลองใหม่"
fi
"$VENV/bin/python" -c "import tkinter, customtkinter, openpyxl" 2>/dev/null \
  || fail "Python ที่ติดตั้งเปิดหน้าจอโปรแกรมไม่ได้ (tkinter/customtkinter/openpyxl) - ติดตั้ง Python จาก https://www.python.org/downloads/ แล้วรันไฟล์นี้อีกครั้ง"

echo "[3/4] คัดลอกโปรแกรม ..."
cp -f "$SRC/tongheer_app.pyw" "$SRC/tongheer_invb.py" "$SRC/app.ico" "$APPDIR/" || fail "คัดลอกโปรแกรมไม่สำเร็จ"
[ -f "$SRC/DATA_TONGHEER.xlsx" ] && cp -f "$SRC/DATA_TONGHEER.xlsx" "$APPDIR/"
# ไม่เขียนทับค่าที่โปรแกรมจำไว้ ถ้าเคยติดตั้งแล้ว
[ -f "$APPDIR/tongheer_map.json" ] || { [ -f "$SRC/tongheer_map.json" ] && cp "$SRC/tongheer_map.json" "$APPDIR/"; }

echo "[4/4] สร้างแอป \"ECS Pro\" ..."
rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources" || fail "สร้างแอปที่ $APP ไม่ได้"
cat > "$APP/Contents/MacOS/ECSPro" <<EOF
#!/bin/bash
cd "$APPDIR" || exit 1
exec "$VENV/bin/python" "$APPDIR/tongheer_app.pyw" "\$@" >"$APPDIR/last_run.log" 2>&1
EOF
chmod +x "$APP/Contents/MacOS/ECSPro" || fail "ตั้งสิทธิ์แอปไม่สำเร็จ"
cat > "$APP/Contents/Info.plist" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>ECS Pro</string>
  <key>CFBundleDisplayName</key><string>ECS Pro</string>
  <key>CFBundleIdentifier</key><string>com.tongheer.ecspro</string>
  <key>CFBundleExecutable</key><string>ECSPro</string>
  <key>CFBundleIconFile</key><string>icon</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>1.0</string>
  <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
EOF
# ไอคอน: แปลง app.ico เป็น icns ถ้าทำได้ (ไม่ได้ก็ใช้ไอคอนมาตรฐาน)
TMP_PNG="$APPDIR/icon_tmp.png"
if sips -s format png "$APPDIR/app.ico" --out "$TMP_PNG" >/dev/null 2>&1 \
   && sips -z 512 512 "$TMP_PNG" >/dev/null 2>&1; then
  sips -s format icns "$TMP_PNG" --out "$APP/Contents/Resources/icon.icns" >/dev/null 2>&1
fi
rm -f "$TMP_PNG"
xattr -dr com.apple.quarantine "$APP" "$APPDIR" 2>/dev/null
# ไอคอนบน Desktop (ลบไอคอนของเวอร์ชันเก่า)
rm -f "$HOME/Desktop/Tong Heer INVB.command" "$HOME/Desktop/ECS Pro.command"
rm -rf "$HOME/Desktop/ECS Pro.app"
ln -s "$APP" "$HOME/Desktop/ECS Pro.app" 2>/dev/null

echo
echo "ติดตั้งเสร็จแล้ว - กำลังเปิด ECS Pro ..."
echo "ครั้งต่อไปเปิดจากไอคอน \"ECS Pro\" บน Desktop หรือใน Launchpad"
open "$APP"
