import os
import re
import xlwings as xw
from PIL import ImageGrab
import time
import sys

# Tạo thư mục lưu ảnh
output_folder = "ANHTHE"
os.makedirs(output_folder, exist_ok=True)
print(f"📁 Đã tạo thư mục lưu ảnh: {os.path.abspath(output_folder)}")

# Đường dẫn đến file Excel
excel_file = "du_lieu.xlsx"  # Thay bằng tên file của bạn
print(f"📊 Đang mở file Excel: {excel_file}")

# Hàm làm sạch tên file
def clean_filename(name):
    cleaned = re.sub(r'[\\/*?:"<>|]', '', str(name)).strip()
    return cleaned if cleaned else "Unknown"

try:
    # Khởi động Excel ở chế độ ẩn
    app = xw.App(visible=False)
    wb = app.books.open(excel_file)
    sheet = wb.sheets.active
    
    # Tìm hàng cuối cùng có dữ liệu trong cột A
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    print(f"🔢 Tổng số hàng dữ liệu: {last_row - 1} (từ hàng 2 đến {last_row})")
    
    # Lấy tất cả hình ảnh trong sheet
    all_shapes = sheet.shapes
    print(f"🖼️ Tìm thấy {len(all_shapes)} hình ảnh trong sheet")
    
    if not all_shapes:
        print("⚠️ Cảnh báo: Không tìm thấy hình ảnh nào trong sheet!")
        wb.close()
        app.quit()
        sys.exit(1)
    
    # Tạo danh sách các hình ảnh và vị trí của chúng
    shapes_info = []
    for shape in all_shapes:
        try:
            # Lấy tọa độ và kích thước hình ảnh
            top = shape.top
            left = shape.left
            height = shape.height
            width = shape.width
            
            # Tính toán vị trí trung tâm
            center_x = left + width / 2
            center_y = top + height / 2
            
            shapes_info.append({
                'shape': shape,
                'top': top,
                'left': left,
                'center_x': center_x,
                'center_y': center_y
            })
        except Exception as e:
            print(f"  ⚠️ Lỗi khi lấy thông tin hình ảnh: {str(e)}")
    
    print(f"ℹ️ Đã thu thập thông tin cho {len(shapes_info)} hình ảnh")
    
    # Lấy thông tin vị trí các ô trong cột 3
    print("🔍 Đang thu thập thông tin vị trí các ô...")
    cell_positions = {}
    for row in range(2, last_row + 1):
        try:
            cell = sheet.range((row, 3))  # Cột C
            cell_positions[row] = {
                'top': cell.top,
                'left': cell.left,
                'height': cell.height,
                'width': cell.width
            }
        except Exception as e:
            print(f"⚠️ Lỗi tại hàng {row}: {str(e)}")
    
    # Ánh xạ hình ảnh vào các ô
    image_mapping = {}
    unmatched_shapes = []
    
    # Pass 1: Center-based mapping
    for shape_info in shapes_info:
        closest_row = None
        min_distance = float('inf')
        
        # Tìm ô gần nhất với trung tâm hình ảnh
        for row, cell_info in cell_positions.items():
            cell_center_x = cell_info['left'] + cell_info['width'] / 2
            cell_center_y = cell_info['top'] + cell_info['height'] / 2
            
            # Tính khoảng cách Euclid
            distance = ((shape_info['center_x'] - cell_center_x) ** 2 + 
                        (shape_info['center_y'] - cell_center_y) ** 2) ** 0.5
            
            if distance < min_distance:
                min_distance = distance
                closest_row = row
        
        if closest_row:
            cell_info = cell_positions[closest_row]
            
            # Tính toán tọa độ trung tâm ô
            cell_center_x = cell_info['left'] + cell_info['width'] / 2
            cell_center_y = cell_info['top'] + cell_info['height'] / 2
            
            # Kiểm tra xem trung tâm ảnh có nằm trong ô không
            center_in_cell = (
                cell_info['left'] <= shape_info['center_x'] <= cell_info['left'] + cell_info['width'] and
                cell_info['top'] <= shape_info['center_y'] <= cell_info['top'] + cell_info['height']
            )
            
            # Kiểm tra xem trung tâm ảnh có gần trung tâm ô không (dung sai 20 pixel)
            tolerance = 20
            near_center = (
                abs(shape_info['center_x'] - cell_center_x) < tolerance and
                abs(shape_info['center_y'] - cell_center_y) < tolerance
            )
            
            # Kiểm tra xem ảnh có nằm trong ranh giới ô không
            within_boundary = (
                shape_info['left'] >= cell_info['left'] - tolerance and
                shape_info['top'] >= cell_info['top'] - tolerance and
                shape_info['left'] + shape_info['shape'].width <= cell_info['left'] + cell_info['width'] + tolerance and
                shape_info['top'] + shape_info['shape'].height <= cell_info['top'] + cell_info['height'] + tolerance
            )
            
            if center_in_cell or near_center or within_boundary:
                image_mapping[closest_row] = shape_info['shape']
                condition = "trong ô" if center_in_cell else "gần trung tâm ô" if near_center else "trong ranh giới ô"
                print(f"  ✅ Ánh xạ ảnh vào hàng {closest_row} (khoảng cách: {min_distance:.2f}, điều kiện: {condition})")
            else:
                unmatched_shapes.append((shape_info, closest_row, min_distance))
                # Ghi log chi tiết để gỡ lỗi
                print(f"  ⚠️ Ảnh gần hàng {closest_row} nhưng không đủ điều kiện (khoảng cách: {min_distance:.2f})")
                print(f"      Tọa độ ảnh: ({shape_info['center_x']}, {shape_info['center_y']})")
                print(f"      Tọa độ ô: ({cell_center_x}, {cell_center_y})")
    
    # Pass 2: Boundary-based mapping for unmatched shapes
    print("\n🔍 Bắt đầu pass 2: Ánh xạ theo ranh giới cho ảnh chưa được ánh xạ")
    for shape_info, closest_row, min_distance in unmatched_shapes:
        cell_info = cell_positions[closest_row]
        tolerance = 30  # Dung sai lớn hơn cho pass 2
        
        # Kiểm tra xem ảnh có nằm trong ranh giới ô không với dung sai lớn
        within_boundary = (
            shape_info['left'] >= cell_info['left'] - tolerance and
            shape_info['top'] >= cell_info['top'] - tolerance and
            shape_info['left'] + shape_info['shape'].width <= cell_info['left'] + cell_info['width'] + tolerance and
            shape_info['top'] + shape_info['shape'].height <= cell_info['top'] + cell_info['height'] + tolerance
        )
        
        if within_boundary:
            # Kiểm tra xem hàng đã có ảnh chưa
            if closest_row not in image_mapping:
                image_mapping[closest_row] = shape_info['shape']
                print(f"  ✅ [Pass 2] Ánh xạ ảnh vào hàng {closest_row} (điều kiện: trong ranh giới ô mở rộng)")
            else:
                print(f"  ⚠️ [Pass 2] Hàng {closest_row} đã có ảnh, bỏ qua ảnh thứ hai")
        else:
            print(f"  ❌ [Pass 2] Không ánh xạ được ảnh cho hàng {closest_row} ngay cả với dung sai mở rộng")
    
    print(f"📊 Tổng số ảnh đã ánh xạ: {len(image_mapping)}/{len(shapes_info)}")
    
    print(f"📊 Đã ánh xạ {len(image_mapping)} ảnh vào các ô")
    
    processed_count = 0
    missing_images = 0
    
    # Duyệt qua từng hàng dữ liệu để lưu ảnh
    print("\n🚀 Bắt đầu xuất ảnh...")
    for row in range(2, last_row + 1):
        try:
            # Đọc dữ liệu nhân viên
            ma_nv = sheet.range(f'A{row}').value
            ho_ten = sheet.range(f'B{row}').value
            
            if not ma_nv or not ho_ten:
                print(f"  ⏩ Hàng {row}: Bỏ qua vì thiếu mã NV hoặc họ tên")
                continue
            
            # Xử lý mã nhân viên: chuyển số thực thành số nguyên nếu có phần thập phân .0
            if isinstance(ma_nv, float) and ma_nv.is_integer():
                ma_nv = int(ma_nv)
            
            # Tạo tên file (chỉ sử dụng mã nhân viên)
            filename = f"{clean_filename(ma_nv)}_.png"
            filepath = os.path.join(output_folder, filename)
            
            # Kiểm tra và xử lý ảnh
            if row in image_mapping:
                shape = image_mapping[row]
                
                try:
                    max_attempts = 3
                    attempt = 0
                    image = None
                    
                    while attempt < max_attempts and image is None:
                        # Copy hình ảnh vào clipboard
                        shape.api.Copy()
                        time.sleep(0.5)  # Tăng thời gian chờ clipboard
                        image = ImageGrab.grabclipboard()
                        
                        if image:
                            # Lưu ảnh
                            image.save(filepath)
                            processed_count += 1
                            print(f"  ✅ Đã lưu ảnh: {filename}")
                        else:
                            attempt += 1
                            print(f"  ⚠️ Lần thử {attempt}: Không có ảnh trong clipboard tại hàng {row}")
                    
                    if not image:
                        print(f"  ❌ Không có ảnh sau {max_attempts} lần thử tại hàng {row}")
                except Exception as e:
                    print(f"  ❌ Lỗi khi lưu ảnh: {str(e)}")
            else:
                missing_images += 1
                print(f"  ❌ Hàng {row}: Không có ảnh được ánh xạ")
                
        except Exception as e:
            print(f"  ⚠️ Lỗi tại hàng {row}: {str(e)}")
    
    print("\n📊 KẾT QUẢ:")
    print(f"- Tổng số hàng đã xử lý: {last_row - 1}")
    print(f"- Số ảnh đã lưu thành công: {processed_count}")
    print(f"- Số hàng không có ảnh: {missing_images}")
    
    if processed_count == 0:
        print("\n⚠️ Cảnh báo: Không có ảnh nào được lưu! Nguyên nhân có thể:")
        print("   1. Không thể xác định vị trí ảnh chính xác")
        print("   2. Lỗi trong quá trình copy ảnh vào clipboard")
        print("   3. Ảnh được chèn dưới dạng không hỗ trợ")
        print("   4. Cấu trúc file khác với mong đợi")
    
except Exception as e:
    print(f"❌ LỖI TỔNG THỂ: {str(e)}")
    import traceback
    traceback.print_exc()
finally:
    # Đảm bảo đóng ứng dụng Excel
    try:
        if 'wb' in locals():
            wb.close()
        if 'app' in locals():
            app.quit()
        print("✅ Đã đóng ứng dụng Excel")
    except:
        pass
