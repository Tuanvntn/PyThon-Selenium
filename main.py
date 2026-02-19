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
    luong_chay = threading.Thread(target=logic_cao_du_lieu)
    luong_chay.daemon = True  # Đảm bảo tắt app là tắt luôn luồng chạy
    luong_chay.start()


def logic_cao_du_lieu():
    btn_run.config(state=tk.DISABLED)
    log_area.insert(tk.END, "🚀 Đang khởi động robot...\n")
    log_area.see(tk.END)

    driver = None  # Khởi tạo driver bên ngoài để finally có thể gọi được
    try:
        DATA_CONFIG = {
            "Tổng hợp": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-tonghop.html", "max": "100"},
            "Công khai minh bạch": {
                "url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-congkhaiminhbach.html", "max": "18"},
            "Tiến độ giải quyết": {
                "url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-tiendogiaiquyet.html", "max": "20"},
            "DVC trực tuyến": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-dvctructuyen.html",
                               "max": "10"},
            "Mức độ hài lòng": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-mucdohailong.html",
                                "max": "18"},
            "Số hóa hồ sơ": {"url": "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-mucdosohoa.html",
                             "max": "22"}
        }

        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        # Thêm các options giúp chạy ổn định trên máy ảo GitHub/Linux/Mac
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 25)
        final_data_list = []

        for name, config in DATA_CONFIG.items():
            log_area.insert(tk.END, f"🔍 Đang quét: {name}...\n")
            log_area.see(tk.END)
            driver.get(config["url"])

            try:
                # Chọn Tỉnh Tây Ninh
                dropdown = wait.until(EC.element_to_be_clickable((By.ID, "select2-tinhtp-container")))
                dropdown.click()
                search_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
                search_field.send_keys("Tây Ninh")
                time.sleep(2)
                search_field.send_keys("\n")

                time.sleep(6)  # Đợi bảng load

                rows = driver.find_elements(By.TAG_NAME, "tr")
                found = False
                for row in rows:
                    if "Cầu Khởi" in row.text:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        hang = cells[0].text.strip()
                        raw_value = cells[-1].text.strip().split('\n')[0]

                        score = raw_value if "/" in raw_value else f"{raw_value}/{config['max']}"
                        final_data_list.append({"Nhóm Chỉ Tiêu": name, "Hạng": hang, "Điểm": score})
                        found = True
                        break

                if not found:
                    final_data_list.append({"Nhóm Chỉ Tiêu": name, "Hạng": "N/A", "Điểm": f"0/{config['max']}"})

            except Exception as e:
                log_area.insert(tk.END, f"⚠️ Lỗi nhóm {name}: {str(e)[:50]}...\n")

        # --- XUẤT FILE ---
        if final_data_list:
            df = pd.DataFrame(final_data_list)
            # Dùng đường dẫn linh hoạt
            file_name = "KetQua_CauKhoi.xlsx"
            home = os.path.expanduser("~")
            file_path = os.path.join(home, "Desktop", file_name)

            df.to_excel(file_path, index=False)

            log_area.insert(tk.END, "✅ Đã trích xuất xong Cầu Khởi!\n")
            log_area.insert(tk.END, f"📁 File: {file_path}\n")
            messagebox.showinfo("Thành công", f"File đã lưu tại Desktop!")
        else:
            log_area.insert(tk.END, "❌ Không lấy được dữ liệu nào!\n")

    except Exception as e:
        error_msg = str(e).strip()
        if error_msg:
            log_area.insert(tk.END, f"❌ Lỗi hệ thống: {str(e)}\n")
    finally:
        try:
            if driver:
                driver.quit()
        except:
            pass
        btn_run.config(state=tk.NORMAL)
        log_area.see(tk.END)

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





