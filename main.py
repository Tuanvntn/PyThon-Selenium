import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os

def bat_dau_chay():
    # Thầy dùng threading để khi robot chạy, cái cửa sổ giao diện không bị "đơ"
    luong_chay = threading.Thread(target=logic_cao_du_lieu)
    luong_chay.start()


def logic_cao_du_lieu():
    btn_run.config(state=tk.DISABLED)  # Bấm xong thì làm mờ nút để tránh bấm nhiều lần
    log_area.insert(tk.END, "🚀 Đang khởi động robot...\n")

    try:
        # 1. Cấu hình URL và Điểm trần
        DATA_CONFIG = {
            "Tổng hợp": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-tonghop.html", "max": "100"},
            "Công khai minh bạch": {
                "url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-congkhaiminhbach.html",
                "max": "18"},
            "Tiến độ giải quyết": {
                "url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-tiendogiaiquyet.html",
                "max": "20"},
            "DVC trực tuyến": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-dvctructuyen.html",
                               "max": "10"},
            "Mức độ hài lòng": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-mucdohailong.html",
                                "max": "18"},
            "Số hóa hồ sơ": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-mucdosohoa.html",
                             "max": "22"}
        }
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")  # Chạy ẩn danh, không hiện cửa sổ
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.maximize_window()
        wait = WebDriverWait(driver, 25)
        final_data_list = []

        try:
            for name, config in DATA_CONFIG.items():
                print(f"--- Đang trích xuất: {name} ---")
                driver.get(config["url"])

                try:
                    # 2. Chọn Tỉnh Tây Ninh
                    dropdown = wait.until(EC.element_to_be_clickable((By.ID, "select2-tinhtp-container")))
                    dropdown.click()
                    search_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
                    search_field.send_keys("Tây Ninh")
                    time.sleep(1.5)
                    search_field.send_keys("\n")

                    # 3. Chờ bảng cập nhật dữ liệu
                    time.sleep(6)

                    # 4. Tìm dòng Cầu Khởi
                    rows = driver.find_elements(By.TAG_NAME, "tr")
                    found_in_page = False

                    for row in rows:
                        if "Cầu Khởi" in row.text:
                            cells = row.find_elements(By.TAG_NAME, "td")

                            # Cột 0 hoặc 1 thường là STT (Hạng)
                            hang = cells[0].text.strip()

                            # Ô cuối cùng thường là điểm
                            raw_value = cells[-1].text.strip().split('\n')[0]

                            if "/" in raw_value:
                                final_score = raw_value
                            else:
                                final_score = f"{raw_value}/{config['max']}"

                            # Lưu cả hạng và điểm vào summary


                            final_data_list.append({
                                "Nhóm Chỉ Tiêu": name,
                                "Hạng": hang,
                                "Điểm Đạt / Tối Đa": final_score
                            })
                            found_in_page = True
                            break

                    if not found_in_page:
                        final_data_list.append({
                            "Nhóm Chỉ Tiêu": name, "hang": "N/A", "diem": f"0/{config['max']}"
                        })

                except Exception as e:
                    print(f"⚠️ Lỗi tại nhóm {name}: {e}")



            # 5. XUẤT FILE DÙNG PANDAS
            df = pd.DataFrame(final_data_list)
            home_directory = os.path.expanduser("~")
            file_path = os.path.join(home_directory, "Desktop", "KetQua_CauKhoi.xlsx")
            df.to_excel(file_path, index=False)
            driver.quit()
            log_area.insert(tk.END, "✅ Đã trích xuất xong Cầu Khởi!\n")
            log_area.insert(tk.END, f"📁 File lưu tại: Desktop\n")
            log_area.see(tk.END)
            messagebox.showinfo("Thành công", f"Đã xuất file thành công ra Desktop!")

        finally:
            time.sleep(5)
            driver.quit()
        # --- CHÈN TOÀN BỘ CODE SELENIUM CỦA EM VÀO ĐÂY ---
        # Thay vì dùng print(), em dùng: log_area.insert(tk.END, "Thông báo...\n")
        log_area.insert(tk.END, "✅ Đã trích xuất xong Cầu Khởi!\n")
        messagebox.showinfo("Thành công", "Đã xuất file Excel thành công!")
    except Exception as e:
        log_area.insert(tk.END, f"❌ Lỗi: {str(e)}\n")
    finally:
        btn_run.config(state=tk.NORMAL)


# --- THIẾT KẾ GIAO DIỆN ---
root = tk.Tk()
root.title("Công cụ trích xuất dữ liệu - Cầu Khởi - Tây Ninh")
root.geometry("500x400")

label_title = tk.Label(root, text="TỰ ĐỘNG TRÍCH XUẤT DỮ LIỆU", font=("Arial", 16, "bold"))
label_title.pack(pady=10)

btn_run = tk.Button(root, text="BẮT ĐẦU CHẠY", command=bat_dau_chay, bg="green", fg="white", font=("Arial", 12, "bold"),
                    padx=20, pady=10)
btn_run.pack(pady=10)

log_area = scrolledtext.ScrolledText(root, width=60, height=15)
log_area.pack(pady=10, padx=10)

root.mainloop()





