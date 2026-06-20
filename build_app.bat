@echo off
setlocal
echo ==============================================
echo BẮT ĐẦU ĐÓNG GÓI ỨNG DỤNG QUẢN LÝ HỒ SƠ CSKV v2.0
echo ==============================================

echo 1. Kiểm tra và Cài đặt thư viện mới nhất...
pip install -r requirements.txt

echo 2. Xóa các bản build cũ (Dọn dẹp môi trường)...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo 3. Chạy PyInstaller đóng gói nâng cao...
:: --noconfirm: Ghi đè mà không hỏi
:: --windowed: Ẩn cửa sổ dòng lệnh đen khi chạy ứng dụng
:: --add-data: Đóng gói templates và static vào file chạy
:: --collect-all: Đảm bảo các thư viện như qrcode, openpyxl, unidecode được đóng gói đầy đủ
python -m PyInstaller --noconfirm --log-level=INFO --windowed ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --collect-all qrcode ^
    --collect-all openpyxl ^
    --collect-all unidecode ^
    app.py

echo.
echo 4. Tự động gom dữ liệu vào bộ cài (Tính năng Drop-and-Run)...
if exist data_hoso.db (
    xcopy /y data_hoso.db dist\app\
    echo - Da copy database vao bo cai.
)
if exist uploads (
    xcopy /e /i /y uploads dist\app\uploads
    echo - Da copy thu muc tai lieu uploads vao bo cai.
) else (
    mkdir dist\app\uploads
)

echo.
echo ==============================================
echo ĐÓNG GÓI HOÀN TẤT!
echo BẠN CHỈ CẦN COPY THƯ MỤC: dist\app VÀO USB LÀ XONG.
echo Mở máy cơ quan, vào thư mục 'app' và chạy 'app.exe' là dùng được ngay.
echo ==============================================
pause
