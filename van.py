"""
Module xuất ảnh thẻ từ file Excel

Mục đích:
    - Đọc file Excel chứa thông tin nhân viên và ảnh thẻ
    - Xác định vị trí ảnh trong các ô tương ứng
    - Tạm thời phóng to ảnh để lấy chất lượng gốc
    - Xuất ảnh ra thư mục với tên file theo mã nhân viên
    - Giữ nguyên định dạng và bố cục file Excel gốc

Các chức năng chính:
    1. Tạo thư mục lưu ảnh đầu ra
    2. Xác định vị trí ảnh trong các ô Excel
    3. Tạm thời phóng to ảnh để lấy chất lượng cao
    4. Xuất ảnh dưới dạng PNG chất lượng cao
    5. Khôi phục trạng thái ban đầu của Excel
"""

import os
import re
import sys
import time
import win32com.client as win32
import pythoncom
from PIL import ImageGrab
import warnings
import traceback

# Tắt cảnh báo không cần thiết
warnings.filterwarnings("ignore")

def clean_filename(name):
    """
    Làm sạch chuỗi để tạo tên file an toàn
    
    Args:
        name (str/int): Giá trị đầu vào có thể là chuỗi hoặc số
        
    Returns:
        str: Tên file đã được làm sạch
    """
    # Loại bỏ các ký tự đặc biệt không hợp lệ trong tên file
    cleaned = re.sub(r'[\\/*?:"<>|]', '', str(name)).strip()
    return cleaned if cleaned else "Unknown"

def export_images(excel_file_path, output_folder, scale_factor, wait_time, log_callback=None):
    """
    Hàm chính thực hiện xuất ảnh từ Excel
    
    Args:
        excel_file_path (str): Đường dẫn đến file Excel
        output_folder (str): Thư mục lưu ảnh đầu ra
        scale_factor (float): Hệ số phóng to ảnh để lấy chất lượng gốc
        wait_time (float): Thời gian chờ giữa các thao tác (giây)
        log_callback (function): Hàm callback để ghi log ra giao diện
        
    Returns:
        bool: True nếu thành công, False nếu có lỗi
    """
    def log(message):
        """Ghi log ra console hoặc giao diện"""
        print(message)
        if log_callback:
            log_callback(message)
    
    try:
        # Tạo thư mục lưu ảnh nếu chưa tồn tại
        os.makedirs(output_folder, exist_ok=True)
        log(f"📁 Đã tạo thư mục lưu ảnh: {os.path.abspath(output_folder)}")
        log(f"📊 Đang mở file Excel: {os.path.basename(excel_file_path)}")
        
        # Khởi tạo môi trường COM
        pythoncom.CoInitialize()
        
        # Khởi động Excel ở chế độ ẩn
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        log("🟢 Đã khởi động Excel ở chế độ ẩn")
        
        # Mở file Excel
        wb = excel.Workbooks.Open(os.path.abspath(excel_file_path))
        sheet = wb.ActiveSheet
        log("🔓 Đã mở file Excel thành công")
        
        # Xác định hàng cuối cùng có dữ liệu
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(win32.constants.xlUp).Row
        log(f"🔢 Tổng số hàng dữ liệu: {last_row - 1} (từ hàng 2 đến {last_row})")
        
        # Lấy tất cả hình ảnh trong sheet
        all_shapes = sheet.Shapes
        log(f"🖼️ Tìm thấy {all_shapes.Count} hình ảnh trong sheet")
        
        # Kiểm tra nếu không có ảnh nào
        if all_shapes.Count == 0:
            log("⚠️ Cảnh báo: Không tìm thấy hình ảnh nào trong sheet!")
            wb.Close(False)
            excel.Quit()
            return False
        
        # Lưu trữ thông tin hình ảnh
        shapes_info = []
        for i in range(1, all_shapes.Count + 1):
            shape = all_shapes.Item(i)
            try:
                # Lưu trạng thái hiện tại của ảnh
                shapes_info.append({
                    'shape': shape,
                    'original_top': shape.Top,
                    'original_left': shape.Left,
                    'original_width': shape.Width,
                    'original_height': shape.Height
                })
            except Exception as e:
                log(f"  ⚠️ Lỗi khi lấy thông tin hình ảnh: {str(e)}")
        
        log(f"ℹ️ Đã thu thập thông tin cho {len(shapes_info)} hình ảnh")
        
        # Lấy vị trí các ô trong cột C (cột 3)
        log("🔍 Đang thu thập thông tin vị trí các ô...")
        cell_positions = {}
        for row in range(2, last_row + 1):
            try:
                cell = sheet.Cells(row, 3)
                cell_positions[row] = {
                    'top': cell.Top,
                    'left': cell.Left,
                    'height': cell.Height,
                    'width': cell.Width
                }
            except Exception as e:
                log(f"⚠️ Lỗi tại hàng {row}: {str(e)}")
        
        # Ánh xạ hình ảnh vào các ô tương ứng
        image_mapping = {}
        unmatched_shapes = []
        
        # Bước 1: Ánh xạ dựa trên trung tâm
        log("\n🔍 Bắt đầu ánh xạ ảnh vào các ô (Pass 1: Dựa trên trung tâm)...")
        for shape_info in shapes_info:
            shape = shape_info['shape']
            closest_row = None
            min_distance = float('inf')
            
            # Tính tọa độ trung tâm hình ảnh
            center_x = shape.Left + shape.Width / 2
            center_y = shape.Top + shape.Height / 2
            
            # Tìm ô gần nhất với ảnh
            for row, cell_info in cell_positions.items():
                cell_center_x = cell_info['left'] + cell_info['width'] / 2
                cell_center_y = cell_info['top'] + cell_info['height'] / 2
                
                # Tính khoảng cách Euclid
                distance = ((center_x - cell_center_x) ** 2 + 
                            (center_y - cell_center_y) ** 2) ** 0.5
                
                if distance < min_distance:
                    min_distance = distance
                    closest_row = row
            
            if closest_row:
                cell_info = cell_positions[closest_row]
                cell_center_x = cell_info['left'] + cell_info['width'] / 2
                cell_center_y = cell_info['top'] + cell_info['height'] / 2
                
                # Kiểm tra các điều kiện ánh xạ
                center_in_cell = (
                    cell_info['left'] <= center_x <= cell_info['left'] + cell_info['width'] and
                    cell_info['top'] <= center_y <= cell_info['top'] + cell_info['height']
                )
                
                tolerance = 20
                near_center = (
                    abs(center_x - cell_center_x) < tolerance and
                    abs(center_y - cell_center_y) < tolerance
                )
                
                within_boundary = (
                    shape.Left >= cell_info['left'] - tolerance and
                    shape.Top >= cell_info['top'] - tolerance and
                    shape.Left + shape.Width <= cell_info['left'] + cell_info['width'] + tolerance and
                    shape.Top + shape.Height <= cell_info['top'] + cell_info['height'] + tolerance
                )
                
                # Nếu thỏa mãn điều kiện, thêm vào ánh xạ
                if center_in_cell or near_center or within_boundary:
                    image_mapping[closest_row] = shape_info
                    condition = "trong ô" if center_in_cell else "gần trung tâm ô" if near_center else "trong ranh giới ô"
                    log(f"  ✅ Ánh xạ ảnh vào hàng {closest_row} (khoảng cách: {min_distance:.2f}, điều kiện: {condition})")
                else:
                    unmatched_shapes.append((shape_info, closest_row, min_distance))
                    log(f"  ⚠️ Ảnh gần hàng {closest_row} nhưng không đủ điều kiện (khoảng cách: {min_distance:.2f})")
        
        # Bước 2: Ánh xạ cho các ảnh chưa được xử lý (dùng dung sai lớn hơn)
        log("\n🔍 Bắt đầu ánh xạ bổ sung (Pass 2: Dùng dung sai rộng hơn)...")
        for shape_info, closest_row, min_distance in unmatched_shapes:
            shape = shape_info['shape']
            cell_info = cell_positions[closest_row]
            tolerance = 30
            
            within_boundary = (
                shape.Left >= cell_info['left'] - tolerance and
                shape.Top >= cell_info['top'] - tolerance and
                shape.Left + shape.Width <= cell_info['left'] + cell_info['width'] + tolerance and
                shape.Top + shape.Height <= cell_info['top'] + cell_info['height'] + tolerance
            )
            
            if within_boundary:
                if closest_row not in image_mapping:
                    image_mapping[closest_row] = shape_info
                    log(f"  ✅ [Pass 2] Ánh xạ ảnh vào hàng {closest_row} (điều kiện: trong ranh giới ô mở rộng)")
                else:
                    log(f"  ⚠️ [Pass 2] Hàng {closest_row} đã có ảnh, bỏ qua ảnh thứ hai")
            else:
                log(f"  ❌ [Pass 2] Không ánh xạ được ảnh cho hàng {closest_row}")
        
        log(f"📊 Tổng số ảnh đã ánh xạ: {len(image_mapping)}/{len(shapes_info)}")
        
        # Xuất ảnh
        processed_count = 0
        missing_images = 0
        log("\n🚀 Bắt đầu xuất ảnh chất lượng cao...")
        
        for row in range(2, last_row + 1):
            try:
                # Đọc thông tin nhân viên
                ma_nv = sheet.Cells(row, 1).Value
                ho_ten = sheet.Cells(row, 2).Value
                
                # Bỏ qua nếu thiếu thông tin
                if not ma_nv or not ho_ten:
                    log(f"  ⏩ Hàng {row}: Bỏ qua vì thiếu mã NV hoặc họ tên")
                    continue
                
                # Chuẩn hóa mã nhân viên
                if isinstance(ma_nv, float) and ma_nv.is_integer():
                    ma_nv = int(ma_nv)
                
                # Tạo tên file
                filename = f"{clean_filename(ma_nv)}_.png"
                filepath = os.path.join(output_folder, filename)
                
                # Xử lý nếu có ảnh ánh xạ
                if row in image_mapping:
                    shape_info = image_mapping[row]
                    shape = shape_info['shape']
                    
                    try:
                        # LƯU TRẠNG THÁI HIỆN TẠI
                        current_top = shape.Top
                        current_left = shape.Left
                        current_width = shape.Width
                        current_height = shape.Height
                        
                        # TẠM THỜI PHÓNG TO ẢNH ĐỂ LẤY CHẤT LƯỢNG GỐC
                        shape.Width = current_width * scale_factor
                        shape.Height = current_height * scale_factor
                        
                        # DI CHUYỂN ẢNH RA KHỎI VÙNG HIỂN THỊ
                        shape.Top = -1000
                        shape.Left = -1000
                        
                        # SAO CHÉP ẢNH Ở CHẤT LƯỢNG CAO
                        shape.Copy()
                        time.sleep(wait_time)
                        
                        # LẤY ẢNH TỪ CLIPBOARD VÀ LƯU
                        image = ImageGrab.grabclipboard()
                        
                        if image:
                            image.save(filepath, format='PNG')
                            processed_count += 1
                            log(f"  ✅ Đã lưu ảnh chất lượng cao: {filename} ({image.width}x{image.height} px)")
                        else:
                            log(f"  ❌ Không có ảnh trong clipboard tại hàng {row}")
                        
                        # KHÔI PHỤC TRẠNG THÁI BAN ĐẦU
                        shape.Top = current_top
                        shape.Left = current_left
                        shape.Width = current_width
                        shape.Height = current_height
                        
                    except Exception as e:
                        log(f"  ❌ Lỗi khi xử lý ảnh hàng {row}: {str(e)}")
                        # Cố gắng khôi phục trạng thái nếu có lỗi
                        try:
                            shape.Top = current_top
                            shape.Left = current_left
                            shape.Width = current_width
                            shape.Height = current_height
                        except:
                            pass
                else:
                    missing_images += 1
                    log(f"  ❌ Hàng {row}: Không có ảnh được ánh xạ")
                    
            except Exception as e:
                log(f"  ⚠️ Lỗi tại hàng {row}: {str(e)}")
        
        # Báo cáo kết quả
        log("\n📊 BÁO CÁO HOÀN THÀNH:")
        log(f"- Tổng số hàng đã xử lý: {last_row - 1}")
        log(f"- Số ảnh đã lưu thành công: {processed_count}")
        log(f"- Số hàng không có ảnh: {missing_images}")
        
        if processed_count == 0:
            log("\n⚠️ CẢNH BÁO: Không có ảnh nào được lưu! Nguyên nhân có thể:")
            log("   1. Không thể xác định vị trí ảnh")
            log("   2. Lỗi trong quá trình sao chép ảnh")
            log("   3. Định dạng ảnh không hỗ trợ")
            log("   4. Cấu trúc file Excel không như mong đợi")
        
        return True
    
    except Exception as e:
        log(f"❌ LỖI TỔNG THỂ: {str(e)}")
        log(traceback.format_exc())
        return False
    
    finally:
        # Đảm bảo giải phóng tài nguyên
        try:
            if 'wb' in locals():
                wb.Close(False)
            if 'excel' in locals():
                excel.Quit()
            pythoncom.CoUninitialize()
            log("✅ Đã đóng ứng dụng Excel và giải phóng tài nguyên")
        except:
            pass

if __name__ == "__main__":
    # Hàm log mặc định cho chế độ dòng lệnh
    def log(message):
        print(message)
    
    # Thông số mặc định khi chạy trực tiếp
    print("=" * 50)
    print("BẮT ĐẦU QUÁ TRÌNH XUẤT ẢNH TỪ EXCEL")
    print("=" * 50)
    
    # Chạy chương trình chính
    start_time = time.time()
    export_images(
        excel_file_path="B23N OKE.xlsx",
        output_folder="ANHTHE",
        scale_factor=3.0,
        wait_time=0.5,
        log_callback=log
    )
    
    # Tính thời gian thực thi
    elapsed_time = time.time() - start_time
    print("\n" + "=" * 50)
    print(f"HOÀN TẤT SAU {elapsed_time:.2f} GIÂY")
    print("=" * 50)