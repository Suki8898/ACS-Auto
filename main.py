import tkinter as tk
from tkinter import ttk, filedialog
import threading
import os
import subprocess
import webbrowser
import logging
import configparser
import pyautogui
import time
import pandas as pd
from PIL import Image, ImageTk 
import keyboard

APP_NAME = "ACS Auto"
VERSION = "1.1.1"

# --- 1. Logger Setup ---
def setup_logging():
    log_file_path = os.path.join(os.path.dirname(__file__), 'log.txt')

    if not os.path.exists(log_file_path):
        open(log_file_path, 'a', encoding='utf-8').close()

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file_path, mode='a', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
    logger.info("------------- ACS Auto -------------")
    return logger

logger = setup_logging()


# --- 2. Config Manager ---
class ConfigManager:
    def __init__(self, config_file='config.ini'):
        self.config_file = os.path.join(os.path.dirname(__file__), config_file)
        self.config = configparser.ConfigParser()
        self._load_config()

    def _load_config(self):
        if not os.path.exists(self.config_file):
            logger.warning(f"Không tìm thấy file config: {self.config_file}. Đang tạo với các giá trị mặc định.")
            self._create_default_config()
        try:
            self.config.read(self.config_file, encoding='utf-8')
        except configparser.Error as e:
            logger.error(f"Lỗi đọc file config {self.config_file}: {e}")
            self._create_default_config()

    def _create_default_config(self):
        self.config['GENERAL'] = {
            'icon_folder': 'icons',
            'image_folder': 'images',
            'screenshot_delay_sec': '0.5',
            'action_delay_sec': '0.2',
            'find_image_confidence': '0.9'
        }
        self.config['IMAGE_PATHS'] = {
            'connected': 'connected.png',
            'on_off_btn': 'on_off_btn.png',
            'discover_btn': 'discover_blue.png,discover_white.png',
            'device_discovery_title': 'device_discovery_title.png',
            'scan_btn': 'scan_btn.png',
            'discovery_completed_text': 'discovery_completed_text.png',
            'tricolor_led_item': 'tricolor_led_item.png',
            'afvarionaut_pump_item': 'afvarionaut_pump_item.png',
            'dmx_slave_address_field': 'dmx_slave_address_field.png',
            'set_dmx_slave_address_btn': 'set_dmx_slave_address_btn.png',
            'run_dmx_test_btn': 'run_dmx_test_btn.png',
            'stop_dmx_test_btn': 'stop_dmx_test_btn.png',
            'dmx_slider_0': 'dmx_slider_0.png',
            'dmx_slider_0_2': 'dmx_slider_0_2.png',
            'dmx_slider_1': 'dmx_slider_1.png',
            'dmx_slider_2': 'dmx_slider_2.png',
            'dmx_slider_3': 'dmx_slider_3.png',
            'selec_all_1': 'selec_all_1.png',
            'selec_all_2': 'selec_all_2.png',
            'acs_device_manager_title': 'acs_device_manager_title.png',
            'acs_device_configuration_title': 'acs_device_configuration_title.png',
            'acs_device_configuration': 'acs_device_configuration.png',
            'acs_device_manager_1': 'acs_device_manager_1.png',
            'acs_device_manager_2': 'acs_device_manager_2.png',
            'close_window_btn': 'close_window_btn.png',
            'load_btn': 'Load.png',
            'add_btn_1': 'Add_1.png',
            'add_btn_2': 'Add_2.png',
            'adl_file': 'Adl.png',
            'open_adl_btn': 'Open.png',
            'generate_btn': 'Generate.png',
            'device_type_field': 'Device type.png',
            'submersible_pump_type_btn': 'Submersible Pump.png',
            'tricolor_led_type_btn': 'Tricolor Led.png',
            'singlecolor_led_type_btn': 'SingleColor Led.png',
            'dmx2vfd_converter_type_btn': 'Dmx2Vfd Converter.png',
            'device_power_field': 'Device power.png',
            '12w_power_btn': '12W.png',
            '36w_power_btn': '36W.png',
            '100w_power_btn': '100W.png',
            '140w_power_btn': '140W.png',
            '160w_power_btn': '160W.png',
            '150w_power_btn': '150W.png',
            '200w_power_btn': '200W.png',
            'write_btn': 'Write.png',
            'save_btn': 'Save.png',
        }
        self.save_config()

    def get(self, section, option, default=None):
        try:
            return self.config.get(section, option)
        except (configparser.NoSectionError, configparser.NoOptionError):
            logger.warning(f"Tùy chọn config '{option}' không tìm thấy trong phần '{section}'. Sử dụng mặc định: {default}")
            return default

    def set(self, section, option, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, option, str(value))
        self.save_config()

    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            logger.info(f"config đã lưu vào {self.config_file}")
        except IOError as e:
            logger.error(f"Lỗi lưu file config {self.config_file}: {e}")

config_manager = ConfigManager()

# --- 3. Automation Core ---
pyautogui.FAILSAFE = True

class AutoACSAutomation:
    def __init__(self):
        self.icon_folder = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'))
        self.image_folder = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'image_folder'))
        self.screenshot_delay = float(config_manager.get('GENERAL', 'screenshot_delay_sec'))
        self.action_delay = float(config_manager.get('GENERAL', 'action_delay_sec'))
        self.confidence = float(config_manager.get('GENERAL', 'find_image_confidence'))

        self.excel_data = None
        self.current_excel_row_index = 0

        self.stop_requested = False

    def _get_image_paths_list(self, image_name_key):
        filenames_str = config_manager.get('IMAGE_PATHS', image_name_key)
        if not filenames_str:
            logger.error(f"Đường dẫn hình ảnh '{image_name_key}' không tìm thấy trong config.ini")
            return []
        
        filenames = [f.strip() for f in filenames_str.split(',') if f.strip()]
        
        full_paths = []
        for filename in filenames:
            full_path = os.path.join(self.image_folder, filename)
            if not os.path.exists(full_path):
                logger.warning(f"Không tìm thấy file hình ảnh: {full_path} cho '{image_name_key}'. Bỏ qua đường dẫn này.")
            else:
                full_paths.append(full_path)
        
        if not full_paths:
            logger.error(f"Không tìm thấy file hình ảnh cho '{image_name_key}'.")
        return full_paths

    def find_and_click(self, image_name_key, timeout=30, button='left', double_click=False, confidence_override=None):
        image_paths = self._get_image_paths_list(image_name_key)
        if not image_paths:
            return False

        current_confidence = confidence_override if confidence_override is not None else self.confidence
        logger.info(f"Đang tìm kiếm bất kỳ hình ảnh nào trong {image_name_key} (thời gian chờ={timeout}s, độ tin cậy={current_confidence})")
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.stop_requested:
                return "Đã dừng."
            found_any_image_in_this_attempt = False
            for image_path in image_paths:
                try:
                    time.sleep(self.screenshot_delay)
                    location = pyautogui.locateOnScreen(image_path, confidence=current_confidence)
                    if location:
                        center = pyautogui.center(location)
                        logger.info(f"Hình ảnh '{os.path.basename(image_path)}' tìm thấy tại {center}. Đang nhấp{' hai lần' if double_click else ''}...")
                        if double_click:
                            pyautogui.doubleClick(center.x, center.y, interval=0.1)
                        else:
                            pyautogui.click(center.x, center.y, button=button)
                        time.sleep(self.action_delay)
                        return True
                except pyautogui.ImageNotFoundException:
                    logger.debug(f"Không tìm thấy hình ảnh '{os.path.basename(image_path)}', đang thử tiếp nếu có sẵn.")
                    pass 
                except Exception as e:
                    logger.error(f"Lỗi bất ngờ với hình ảnh '{os.path.basename(image_path)}': {e}", exc_info=True)
                    pass
            time.sleep(0.5) 
        logger.warning(f"Không tìm thấy bất kỳ hình ảnh nào cho '{image_name_key}' sau {timeout} giây.")
        return False

    def find_and_right_click(self, image_name_key, timeout=30, button='right', double_click=False, confidence_override=None):
        image_paths = self._get_image_paths_list(image_name_key)
        if not image_paths:
            return False

        current_confidence = confidence_override if confidence_override is not None else self.confidence
        logger.info(f"Đang tìm kiếm bất kỳ hình ảnh nào trong {image_name_key} (thời gian chờ={timeout}s, độ tin cậy={current_confidence})")
        start_time = time.time()
        while time.time() - start_time < timeout:
            found_any_image_in_this_attempt = False
            for image_path in image_paths:
                try:
                    time.sleep(self.screenshot_delay)
                    location = pyautogui.locateOnScreen(image_path, confidence=current_confidence)
                    if location:
                        center = pyautogui.center(location)
                        logger.info(f"Hình ảnh '{os.path.basename(image_path)}' tìm thấy tại {center}. Đang nhấp{' hai lần' if double_click else ''}...")
                        if double_click:
                            pyautogui.doubleClick(center.x, center.y, interval=0.1)
                        else:
                            pyautogui.click(center.x, center.y, button=button)
                        time.sleep(self.action_delay)
                        return True
                except pyautogui.ImageNotFoundException:
                    logger.debug(f"Không tìm thấy hình ảnh '{os.path.basename(image_path)}', đang thử tiếp nếu có sẵn.")
                    pass 
                except Exception as e:
                    logger.error(f"Lỗi bất ngờ với hình ảnh '{os.path.basename(image_path)}': {e}", exc_info=True)
                    pass
            time.sleep(0.5) 
        logger.warning(f"Không tìm thấy bất kỳ hình ảnh nào cho '{image_name_key}' sau {timeout} giây.")
        return False

    def type_text(self, text, image_name_key=None, timeout=10, select_all_first=False):
        if image_name_key:
            if not self.find_and_click(image_name_key, timeout=timeout):
                logger.error(f"Không thể tìm thấy trường nhập '{image_name_key}' để gõ văn bản.")
                return False
        
        logger.info(f"Đang gõ văn bản: '{text}'")
        try:
            if select_all_first:
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
            pyautogui.typewrite(str(text), interval=0.05)
            time.sleep(self.action_delay)
            return True
        except Exception as e:
            logger.error(f"Lỗi khi gõ văn bản '{text}': {e}")
            return False

    def press_key(self, key):
        logger.info(f"Đang nhấn phím: '{key}'")
        try:
            pyautogui.press(key)
            time.sleep(self.action_delay)
            return True
        except Exception as e:
            logger.error(f"Lỗi khi nhấn phím '{key}': {e}")
            return False

    def wait_for_image(self, image_name_key, timeout=30, confidence_override=None):
        image_paths = self._get_image_paths_list(image_name_key)
        if not image_paths:
            return False

        current_confidence = confidence_override if confidence_override is not None else self.confidence
        logger.info(f"Đang chờ bất kỳ hình ảnh nào trong {image_name_key} (thời gian chờ={timeout}s, độ tin cậy={current_confidence})")
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.stop_requested:
                return "Đã dừng."

            for image_path in image_paths:
                try:
                    time.sleep(self.screenshot_delay)
                    location = pyautogui.locateOnScreen(image_path, confidence=current_confidence)
                    if location:
                        logger.info(f"Hình ảnh '{os.path.basename(image_path)}' đã tìm thấy.")
                        return True
                except pyautogui.ImageNotFoundException:
                    pass
                except Exception as e:
                    logger.error(f"Lỗi khi chờ hình ảnh {os.path.basename(image_path)}: {e}")
                    return False
            time.sleep(0.5)
        logger.warning(f"Không tìm thấy bất kỳ hình ảnh nào cho '{image_name_key}' sau {timeout} giây.")
        return False
    
    # --- Đọc Excel ---
    def import_excel_data(self, file_path):
        try:
            df = pd.read_excel(file_path, header=0)
            df = df.iloc[:, [0, 1, 2]]
            df.columns = ['No.', 'Pump', 'Led']
            self.excel_data = df.to_dict('records')
            self.current_excel_row_index = 0
            logger.info(f"Đã nhập {len(self.excel_data)} hàng từ file Excel: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi nhập file Excel: {e}", exc_info=True)
            return False

    def get_current_excel_row(self):
        if self.excel_data and self.current_excel_row_index < len(self.excel_data):
            return self.excel_data[self.current_excel_row_index]
        return None

    def increment_excel_row_index(self):
        if self.excel_data and self.current_excel_row_index < len(self.excel_data):
            self.current_excel_row_index += 1
            logger.info(f"Excel đã tăng 1 hàng {self.current_excel_row_index}")
            return True
        return False

    def reset_excel_row_index(self):
        self.current_excel_row_index = 0
        logger.info("Hàng Excel đã được đặt lại.")

    # --- Các hàm logic ---

    def xac_dinh_vi_tri_thiet_bi(self, timeout=10):
        try:
            discovery_window_location = pyautogui.locateOnScreen(
                self._get_image_paths_list('device_discovery_title')[0],
                confidence=0.8
            )

            if not discovery_window_location:
                logger.warning("Không thể xác định vị trí tiêu đề cửa sổ Device Discovery. Đang tìm kiếm trên toàn màn hình.")
                search_region = None
            else:
                search_region = (
                    discovery_window_location.left,
                    discovery_window_location.top,
                    discovery_window_location.width,
                    discovery_window_location.height
                )
                logger.info(f"Đang tìm kiếm thiết bị chỉ trong cửa sổ Device Discovery tại {search_region}")
        except Exception as e:
            logger.error(f"Lỗi khi xác định vị trí cửa sổ Device Discovery, đang tìm kiếm trên toàn màn hình: {e}", exc_info=True)
            search_region = None

        led_locations = []
        pump_locations = []

        led_image_paths = self._get_image_paths_list('tricolor_led_item')
        pump_image_paths = self._get_image_paths_list('afvarionaut_pump_item')

        locate_confidence = min(0.75, self.confidence * 0.9)
        if locate_confidence < 0.6:
            locate_confidence = 0.6

        logger.info(f"Vị trí thiết bị bắt đầu trong cửa sổ Discover (thời gian chờ = {timeout} s, độ tin cậy = {locate_confidence})")
            
        start_time = time.time()
        while time.time() - start_time < timeout:
        
            current_iteration_led_locs = []
            for led_path in led_image_paths:
                try:
                    if search_region:
                        found_locs = list(pyautogui.locateAllOnScreen(led_path, confidence=locate_confidence, region=search_region))
                    else:
                        found_locs = list(pyautogui.locateAllOnScreen(led_path, confidence=locate_confidence))

                    if found_locs:
                        current_iteration_led_locs.extend(found_locs)

                except pyautogui.ImageNotFoundException:
                    logger.debug(f"Hình ảnh LED '{os.path.basename(led_path)}' không tìm thấy, đang thử tiếp nếu có sẵn.")
                    pass
                except Exception as e:
                    logger.error(f"Lỗi khi xác định vị trí hình ảnh LED '{os.path.basename(led_path)}': {e}", exc_info=True)
                    pass
            
            current_iteration_pump_locs = []
            for pump_path in pump_image_paths:
                try:
                    if search_region:
                        found_locs = list(pyautogui.locateAllOnScreen(pump_path, confidence=locate_confidence, region=search_region))
                    else:
                        found_locs = list(pyautogui.locateAllOnScreen(pump_path, confidence=locate_confidence))
                    if found_locs:
                        current_iteration_pump_locs.extend(found_locs)
                except pyautogui.ImageNotFoundException:
                    logger.debug(f"Hình ảnh PUMP '{os.path.basename(pump_path)}' không tìm thấy, đang thử tiếp nếu có sẵn.")
                    pass
                except Exception as e:
                    logger.error(f"Lỗi khi xác định vị trí hình ảnh PUMP '{os.path.basename(pump_path)}': {e}", exc_info=True)
                    pass
            
            if current_iteration_led_locs or current_iteration_pump_locs:
                logger.info(f"Tìm thấy {len(current_iteration_led_locs)} LED(s) và {len(current_iteration_pump_locs)} PUMP(s) trong Device Discovery.")
                return current_iteration_led_locs, current_iteration_pump_locs
            
            time.sleep(0.5)
                
        logger.warning("Không tìm thấy thiết bị LED hoặc PUMP trong Device Discovery trong thời gian chờ.")
        return [], []

    def chon_thiet_bi(self, device_type_name, device_location):
        logger.info(f"Đang xử lý {device_type_name} tại {pyautogui.center(device_location)}")

        center = pyautogui.center(device_location)
        logger.info(f"Nhấp đúp {device_type_name} tại {center}...")
        pyautogui.doubleClick(center.x, center.y, interval=0.1)
        time.sleep(self.action_delay)
        

    def chon_thiet_bi_va_ghi(self, device_type_name, device_location, address):
        logger.info(f"Đang xử lý {device_type_name} tại {pyautogui.center(device_location)} với địa chỉ {address}")

        center = pyautogui.center(device_location)
        logger.info(f"Nhấp đúp {device_type_name} tại {center}...")
        pyautogui.doubleClick(center.x, center.y, interval=0.1)
        time.sleep(self.action_delay)
        
        time.sleep(1) 

        if not self.type_text(address, image_name_key='dmx_slave_address_field', select_all_first=True, timeout=10):
            return f"Thất bại: Không thể gõ Địa chỉ DMX Slave cho {device_type_name}."
        
        if not self.find_and_click('set_dmx_slave_address_btn', timeout=10):
            return f"Thất bại: Không thể tìm thấy nút 'Ghi địa chỉ DMX slave' cho {device_type_name}."

        logger.info(f"Ghi địa chỉ {device_type_name} thành công.")
        return f"Ghi địa chỉ {device_type_name} thành công."

    def find(self, image_name_key, timeout=30, confidence_override=None):
        image_paths = self._get_image_paths_list(image_name_key)
        if not image_paths:
            return False

        current_confidence = confidence_override if confidence_override is not None else self.confidence
        logger.info(f"Đang tìm kiếm bất kỳ hình ảnh nào trong {image_name_key} (thời gian chờ={timeout}s, độ tin cậy={current_confidence})")
        start_time = time.time()
        while time.time() - start_time < timeout:
            for image_path in image_paths:
                try:
                    time.sleep(self.screenshot_delay)
                    location = pyautogui.locateOnScreen(image_path, confidence=current_confidence)
                    if location:
                        logger.info(f"Hình ảnh '{os.path.basename(image_path)}' đã tìm thấy.")
                        return True
                except pyautogui.ImageNotFoundException:
                    logger.debug(f"Không tìm thấy hình ảnh '{os.path.basename(image_path)}', đang thử tiếp nếu có sẵn.")
                    pass
                except Exception as e:
                    logger.error(f"Lỗi bất ngờ với hình ảnh '{os.path.basename(image_path)}': {e}", exc_info=True)
                    pass
            time.sleep(0.5)
        logger.warning(f"Không tìm thấy bất kỳ hình ảnh nào cho '{image_name_key}' sau {timeout} giây.")
        return False

    def ghi_dia_chi(self):
        logger.info("Đang bắt đầu quy trình Ghi địa chỉ...")

        auto_acs.stop_requested = False

        if not self.excel_data:
            return "Thất bại: Vui lòng nhập file Excel trước."
        
        row_data = self.get_current_excel_row()
        if not row_data:
            return "Thất bại: Không có dữ liệu trong file Excel hoặc không thể lấy hàng hiện tại."

        pump_address = row_data.get('Pump')
        led_address = row_data.get('Led')
        
        logger.info(f"Xử lý hàng {self.current_excel_row_index + 1}/{len(self.excel_data)}: Pump={pump_address}, Led={led_address}")

        results = []

        if self.find('device_discovery_title', timeout=0.2):
            pass
        else:
            if not self.find('connected', timeout=0.2):
                if not self.find_and_click('on_off_btn', timeout=0.2):
                    return "Thất bại: Không thể tìm thấy nút 'On/Off'."
            else:
                pass

            if not self.find_and_click('discover_btn', timeout=0.2):
                return "Thất bại: Không thể tìm thấy nút 'Discover'."
        
        if not self.find_and_click('scan_btn', timeout=0.2):
            return "Thất bại: Không tìm thấy nút 'Scan' trong Device Discovery."
        
        logger.info("Đang chờ 'Discovery completed'...")
        if not self.wait_for_image('discovery_completed_text', timeout=20): 
            return "Thất bại: 'Discovery completed' không xuất hiện trong thời gian chờ."
        
        led_locs, pump_locs = self.xac_dinh_vi_tri_thiet_bi(timeout=5)
     
        if led_locs:
            led_result = self.chon_thiet_bi_va_ghi("Tricolor Led", led_locs[0], led_address)
            results.append(f"LED ({led_address}): {led_result}")
            if "Failed" in led_result:
                self.increment_excel_row_index()
                return f"Thất bại khi 'Ghi địa chỉ' (LED): {led_result}"
            
            time.sleep(1)
            
            if pump_locs:
                logger.info("Cả LED và PUMP đều được phát hiện. Đang discover lại cho PUMP sau LED.")
                if not self.find_and_click('discover_btn', timeout=15):
                    return "Thất bại: Không thể tìm thấy nút 'Discover'."
                
                _, pump_locs_after_rediscover = self.xac_dinh_vi_tri_thiet_bi(timeout=5)
                
                if pump_locs_after_rediscover:
                    pump_result = self.chon_thiet_bi_va_ghi("AFVarionaut Pump", pump_locs_after_rediscover[0], pump_address)
                    results.append(f"PUMP ({pump_address}): {pump_result}")
                    if "Failed" in pump_result:
                        self.increment_excel_row_index()
                        return f"Thất bại khi 'Ghi địa chỉ' (PUMP): {pump_result}"
                else:
                    self.increment_excel_row_index()
                    return f"Thất bại: Không tìm thấy PUMP sau khi discover lại cho địa chỉ {pump_address}."
        
        elif pump_locs:
            pump_result = self.chon_thiet_bi_va_ghi("AFVarionaut Pump", pump_locs[0], pump_address)
            results.append(f"PUMP ({pump_address}): {pump_result}")
            if "Failed" in pump_result:
                self.increment_excel_row_index()
                return f"Thất bại khi 'Ghi địa chỉ' (PUMP): {pump_result}"
        
        else:
            self.increment_excel_row_index()
            return "Thất bại: Không tìm thấy thiết bị (LED/PUMP) trong cửa sổ Device Discovery sau khi quét."

        if not self.find_and_click('on_off_btn', timeout=0.2):
            return "Thất bại: Không thể tìm thấy nút 'On/Off'."
        
        self.increment_excel_row_index() 
        return "\n ".join(results) + f" (Đã hoàn thành hàng {self.current_excel_row_index}/{len(self.excel_data)})"

    def test(self):
        logger.info("Đang bắt đầu quy trình Test...")

        auto_acs.stop_requested = False

        if self.find('device_discovery_title', timeout=0.2):
            pass
        else:
            if not self.find('connected', timeout=0.2):
                if not self.find_and_click('on_off_btn', timeout=0.2):
                    return "Thất bại: Không thể tìm thấy nút 'On/Off'."
            else:
                pass

            if not self.find_and_click('discover_btn', timeout=0.2):
                return "Thất bại: Không thể tìm thấy nút 'Discover'."
        
        if not self.find_and_click('scan_btn', timeout=2):
            return "Thất bại: Không tìm thấy nút 'Scan' trong Device Discovery."
        
        logger.info("Đang chờ 'Discovery completed'...")
        if not self.wait_for_image('discovery_completed_text', timeout=20): 
            return "Thất bại: Văn bản 'Discovery completed' không xuất hiện sau khi quét trong thời gian chờ. Discover thiết bị có thể đã thất bại."
        
        led_locs, pump_locs = self.xac_dinh_vi_tri_thiet_bi(timeout=5)

        if led_locs:
            led_result = self.chon_thiet_bi("Tricolor Led", led_locs[0])

            if not self.find_and_click('run_dmx_test_btn', timeout=10):
                return "Thất bại: Không thể tìm thấy nút 'Run DMX Test'."

            time.sleep(0.2) 

            if not self.find_and_click('selec_all_2', timeout=0.2):
                logger.info("Không tìm thấy 'select all' để tắt, bỏ qua.")

            if not self.is_slider_already_moved('dmx_slider_0_2', timeout=0.2):
                if not self.drag_slider('dmx_slider_0', 0, -100, duration=0.6):
                    return "Thất bại: Không thể kéo thanh trượt 0."
            else:
                logger.info("Thanh trượt 0 đã được kéo.")

            if not self.drag_slider('dmx_slider_1', 0, -85, duration=0.6): # Kéo lên 85 pixel, 1 giây
                return "Thất bại: Không thể kéo thanh trượt 1 lên."
            time.sleep(0.2)
            if not self.drag_slider('dmx_slider_1', 0, 85, duration=0.6):  # Kéo xuống 85 pixel, 1 giây
                return "Thất bại: Không thể kéo thanh trượt 1 xuống."

            if not self.drag_slider('dmx_slider_2', 0, -85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 2 lên."
            time.sleep(0.2)
            if not self.drag_slider('dmx_slider_2', 0, 85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 2 xuống."

            if not self.drag_slider('dmx_slider_3', 0, -85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 3 lên."
            time.sleep(0.2)
            if not self.drag_slider('dmx_slider_3', 0, 85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 3 xuống."
            
            if not self.find_and_click('selec_all_1', timeout=1):
                return "Thất bại: Không thể bật select all."
            
            if not self.drag_slider('dmx_slider_1', 0, -85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 1 lên."
            time.sleep(2)
            if not self.drag_slider('dmx_slider_1', 0, 85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 1 xuống." 
                       
            if not self.find_and_click('selec_all_2', timeout=1):
                return "Thất bại: Không thể tắt select all."
            
            if not self.find_and_click('stop_dmx_test_btn', timeout=5):
                return "Thất bại: Không thể tìm thấy nút 'Stop DMX Test'."

            if pump_locs:
                logger.info("Cả LED và PUMP đều được phát hiện. Đang discover lại cho PUMP sau LED.")
                if not self.find_and_click('discover_btn', timeout=15):
                    return "Thất bại: Không thể tìm thấy nút 'Discover'."
                
                _, pump_locs_after_rediscover = self.xac_dinh_vi_tri_thiet_bi(timeout=5)
                
                if pump_locs_after_rediscover:
                    pump_result = self.chon_thiet_bi("AFVarionaut Pump", pump_locs_after_rediscover[0])

                if not self.find_and_click('run_dmx_test_btn', timeout=10):
                    return "Thất bại: Không thể tìm thấy nút 'Run DMX Test'."

                time.sleep(0.2)

                if not self.find_and_click('selec_all_2', timeout=0.2):
                    logger.info("Không tìm thấy 'select all' để tắt, bỏ qua.")

                # Kiểm tra trạng thái của thanh trượt 0 trước khi kéo
                if not self.is_slider_already_moved('dmx_slider_0_2', timeout=0.2):
                    # Kéo thanh trượt 0 lên cao nếu chưa được kéo
                    if not self.drag_slider('dmx_slider_0', 0, -100, duration=0.5):
                        return "Thất bại: Không thể kéo thanh trượt 0."
                else:
                    logger.info("Thanh trượt 0 đã được kéo.")

                if not self.drag_slider('dmx_slider_1', 0, -85, duration=1):
                    return "Thất bại: Không thể kéo thanh trượt 1 lên."
                time.sleep(1)
                if not self.drag_slider('dmx_slider_1', 0, 85, duration=1):
                    return "Thất bại: Không thể kéo thanh trượt 1 xuống."

                if not self.find_and_click('stop_dmx_test_btn', timeout=10):
                    return "Thất bại: Không thể tìm thấy nút 'Stop DMX Test'."

        elif pump_locs:
            pump_result = self.chon_thiet_bi("AFVarionaut Pump", pump_locs[0])

            if not self.find_and_click('run_dmx_test_btn', timeout=10):
                return "Thất bại: Không thể tìm thấy nút 'Run DMX Test'."

            time.sleep(0.2)

            if not self.is_slider_already_moved('dmx_slider_0_2', timeout=0.2):
                if not self.drag_slider('dmx_slider_0', 0, -100, duration=0.5):
                    return "Thất bại: Không thể kéo thanh trượt 0."
            else:
                logger.info("Thanh trượt 0 đã được kéo.")

            if not self.drag_slider('dmx_slider_1', 0, -85, duration=1):
                return "Thất bại: Không thể kéo thanh trượt 1 lên."
            time.sleep(1) 
            if not self.drag_slider('dmx_slider_1', 0, 85, duration=1):
                return "Thất bại: Không thể kéo thanh trượt 1 xuống."

            if not self.find_and_click('stop_dmx_test_btn', timeout=10):
                return "Thất bại: Không thể tìm thấy nút 'Stop DMX Test'."
        
        else:
            self.increment_excel_row_index()
            return "Thất bại: Không thấy thiết bị (LED/PUMP) trong cửa sổ Device Discovery sau khi quét."
        
        if not self.find_and_click('on_off_btn', timeout=0.2):
            return "Thất bại: Không thể tìm thấy nút 'On/Off'."
        
        return f"Đã test xong."

    def ghi_dia_chi_va_test(self):
        logger.info("Đang bắt đầu quy trình Ghi địa chỉ & Test...")

        auto_acs.stop_requested = False

        if not self.excel_data:
            return "Thất bại: Vui lòng nhập file Excel trước."
        
        row_data = self.get_current_excel_row()
        if not row_data:
            return "Thất bại: Không có dữ liệu trong file Excel hoặc không thể lấy hàng hiện tại."

        pump_address = row_data.get('Pump')
        led_address = row_data.get('Led')
        
        logger.info(f"Xử lý hàng {self.current_excel_row_index + 1}/{len(self.excel_data)}: Pump={pump_address}, Led={led_address}")

        results = []

        if self.find('device_discovery_title', timeout=0.2):
            pass
        else:
            if not self.find('connected', timeout=0.2):
                if not self.find_and_click('on_off_btn', timeout=0.2):
                    return "Thất bại: Không thể tìm thấy nút 'On/Off'."
            else:
                pass

            if not self.find_and_click('discover_btn', timeout=0.2):
                return "Thất bại: Không thể tìm thấy nút 'Discover'."
        
        if not self.find_and_click('scan_btn', timeout=10):
            return "Thất bại: Không tìm thấy nút 'Scan' trong Device Discovery."
        
        logger.info("Đang chờ 'Discovery completed'...")
        if not self.wait_for_image('discovery_completed_text', timeout=20): 
            return "Thất bại: 'Discovery completed' không xuất hiện trong thời gian chờ."
        
        led_locs, pump_locs = self.xac_dinh_vi_tri_thiet_bi(timeout=5)

        if led_locs:
            led_result = self.chon_thiet_bi_va_ghi("Tricolor Led", led_locs[0], led_address)

            if not self.find_and_click('run_dmx_test_btn', timeout=10):
                return "Thất bại: Could not find 'Run DMX Test' button."

            time.sleep(0.5)

            if not self.find_and_click('selec_all_2', timeout=0.2):
                logger.info("Không tìm thấy 'select all' để tắt, bỏ qua.")

            if not self.is_slider_already_moved('dmx_slider_0_2', timeout=0.2):
                if not self.drag_slider('dmx_slider_0', 0, -100, duration=0.6):
                    return "Thất bại: Không thể kéo thanh trượt 0."
            else:
                logger.info("Slider 0 is already moved, skipping dragging it.")

            if not self.drag_slider('dmx_slider_1', 0, -85, duration=0.6): # Kéo lên 85 pixel, 1 giây
                return "Thất bại: Không thể kéo thanh trượt 1 lên."
            time.sleep(0.2)
            if not self.drag_slider('dmx_slider_1', 0, 85, duration=0.6):  # Kéo xuống 85 pixel, 1 giây
                return "Thất bại: Không thể kéo thanh trượt 1 xuống."

            if not self.drag_slider('dmx_slider_2', 0, -85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 2 lên."
            time.sleep(0.2)
            if not self.drag_slider('dmx_slider_2', 0, 85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 2 xuống."

            if not self.drag_slider('dmx_slider_3', 0, -85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 3 lên."
            time.sleep(0.2)
            if not self.drag_slider('dmx_slider_3', 0, 85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 3 xuống."
            
            if not self.find_and_click('selec_all_1', timeout=1):
                return "Thất bại: Không thể bật select all."
            
            if not self.drag_slider('dmx_slider_1', 0, -85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 1 lên."
            time.sleep(2)
            if not self.drag_slider('dmx_slider_1', 0, 85, duration=0.6):
                return "Thất bại: Không thể kéo thanh trượt 1 xuống." 
                       
            if not self.find_and_click('selec_all_2', timeout=1):
                return "Thất bại: Không thể tắt select all."
            
            if not self.find_and_click('stop_dmx_test_btn', timeout=10):
                return "Thất bại: Không thể tìm thấy nút 'Stop DMX Test'."

            if pump_locs:
                logger.info("Cả LED và PUMP đều được phát hiện. Đang discover lại cho PUMP sau LED.")
                if not self.find_and_click('discover_btn', timeout=15):
                    return "Thất bại: Không tìm thấy nút 'Discover'."
                
                _, pump_locs_after_rediscover = self.xac_dinh_vi_tri_thiet_bi(timeout=5)
                
                if pump_locs_after_rediscover:
                    pump_result = self.chon_thiet_bi_va_ghi("AFVarionaut Pump", pump_locs_after_rediscover[0], pump_address)

                if not self.find_and_click('run_dmx_test_btn', timeout=10):
                    return "Thất bại: Không tìm thấy nút 'Run DMX Test'."

                time.sleep(0.2)

                if not self.find_and_click('selec_all_2', timeout=0.2):
                    logger.info("Không tìm thấy 'select all' để tắt, bỏ qua.")

                if not self.is_slider_already_moved('dmx_slider_0_2', timeout=0.2):
                    if not self.drag_slider('dmx_slider_0', 0, -100, duration=0.5):
                        return "Thất bại: Không thể kéo thanh trượt 0."
                else:
                    logger.info("Slider 0 is already moved, skipping dragging it.")

                if not self.drag_slider('dmx_slider_1', 0, -85, duration=1):
                    return "Thất bại: Không thể kéo thanh trượt 1 lên."
                time.sleep(1)
                if not self.drag_slider('dmx_slider_1', 0, 85, duration=1):
                    return "Thất bại: Không thể kéo thanh trượt 1 xuống."

                if not self.find_and_click('stop_dmx_test_btn', timeout=10):
                    return "Thất bại: Không tìm thấy nút 'Stop DMX Test'."

        elif pump_locs:
            pump_result = self.chon_thiet_bi_va_ghi("AFVarionaut Pump", pump_locs[0], pump_address)

            if not self.find_and_click('run_dmx_test_btn', timeout=10):
                return "Thất bại: Không tìm thấy nút 'Run DMX Test'."

            time.sleep(0.2)

            if not self.is_slider_already_moved('dmx_slider_0_2', timeout=0.2):
                if not self.drag_slider('dmx_slider_0', 0, -100, duration=0.5):
                    return "Thất bại: Không thể kéo thanh trượt 0."
            else:
                logger.info("Thanh trượt 0 đã được kéo.")

            if not self.drag_slider('dmx_slider_1', 0, -85, duration=1):
                return "Thất bại: Không thể kéo thanh trượt 1 lên."
            time.sleep(1)
            if not self.drag_slider('dmx_slider_1', 0, 85, duration=1):
                return "Thất bại: Không thể kéo thanh trượt 1 xuống."

            if not self.find_and_click('stop_dmx_test_btn', timeout=10):
                return "Thất bại: Không tìm thấy nút 'Stop DMX Test'."
        
        else:
            self.increment_excel_row_index()
            return "Thất bại: Không thấy thiết bị (LED/PUMP) trong cửa sổ Device Discovery sau khi quét."
        
        if not self.find_and_click('on_off_btn', timeout=0.2):
            return "Thất bại: Không thể tìm thấy nút 'On/Off'."

        return f"Thành công: 'Ghi địa chỉ & Test' đã hoàn thành."

    def drag_slider(self, slider_image_key, offset_x, offset_y, duration=1):
        image_paths = self._get_image_paths_list(slider_image_key)
        if not image_paths:
            return False

        logger.info(f"Đang kéo thanh trượt '{slider_image_key}'...")
        try:
            location = self.locate_image(image_paths, confidence_override=0.8, timeout=10)
            if location:
                center = pyautogui.center(location)
                pyautogui.moveTo(center.x, center.y, duration=0.2)
                pyautogui.dragRel(offset_x, offset_y, duration=duration, button='left')
                time.sleep(self.action_delay)
                return True
            else:
                logger.error(f"Không tìm thấy thanh trượt '{slider_image_key}'.")
                return False
        except Exception as e:
            logger.error(f"Lỗi khi kéo thanh trượt '{slider_image_key}': {e}", exc_info=True)
            return False

    def is_slider_already_moved(self, slider_moved_image_key, timeout=5, confidence_override=None):
        image_paths = self._get_image_paths_list(slider_moved_image_key)
        if not image_paths:
            return False

        current_confidence = confidence_override if confidence_override is not None else self.confidence
        logger.info(f"Kiểm tra xem thanh trượt đã được di chuyển chưa (timeout={timeout}s, confidence={current_confidence})...")
        start_time = time.time()
        while time.time() - start_time < timeout:
            for image_path in image_paths:
                try:
                    time.sleep(self.screenshot_delay)
                    location = pyautogui.locateOnScreen(image_path, confidence=current_confidence)
                    if location:
                        logger.info(f"Thanh trượt đã được di chuyển (tìm thấy hình ảnh '{os.path.basename(image_path)}').")
                        return True
                except pyautogui.ImageNotFoundException:
                    pass
                except Exception as e:
                    logger.error(f"Lỗi khi kiểm tra thanh trượt đã được di chuyển chưa (hình ảnh '{os.path.basename(image_path)}'): {e}")
                    return False
            time.sleep(0.5)
        logger.info("Thanh trượt chưa được di chuyển.")
        return False

    def locate_image(self, image_paths, confidence_override=None, timeout=30):
        current_confidence = confidence_override if confidence_override is not None else self.confidence
        start_time = time.time()
        while time.time() - start_time < timeout:
            for image_path in image_paths:
                try:
                    time.sleep(self.screenshot_delay)
                    location = pyautogui.locateOnScreen(image_path, confidence=current_confidence)
                    if location:
                        return location
                except pyautogui.ImageNotFoundException:
                    pass
                except Exception as e:
                    logger.error(f"Lỗi khi tìm hình ảnh {os.path.basename(image_path)}: {e}")
                    return None
            time.sleep(0.5)
        logger.warning(f"Không tìm thấy hình ảnh sau {timeout} giây.")
        return None

    def _select_device_type(self, device_type):
        logger.info(f"Selecting device type: {device_type}")
        if not self.find_and_click('device_type_field', timeout=10):
            return f"Thất bại: Could not find 'Device type' field."

        device_type_map = {
            "AFVarionaut Pump": "afvarionaut_pump_type_btn",
            "Submersible Pump": "submersible_pump_type_btn",
            "Tricolor Led": "tricolor_led_type_btn",
            "SingleColor Led": "singlecolor_led_type_btn",
            "Dmx2Vfd Converter": "dmx2vfd_converter_type_btn",
        }

        if device_type in device_type_map:
            image_key = device_type_map[device_type]
            if not self.find_and_click(image_key, timeout=10):
                return f"Thất bại: Could not find '{device_type}' button."
        else:
            return f"Thất bại: Invalid device type: {device_type}"

        return f"Device type selected: {device_type}"
    
    def _select_device_power(self, device_power):
        logger.info(f"Selecting device power: {device_power}")
        if not self.find_and_click('device_power_field', timeout=10):
            return f"Thất bại: Could not find 'Device power' field."

        device_power_map = {
            "60": "60w_power_btn",
            "100": "100w_power_btn",
            "140": "140w_power_btn",
            "160": "160w_power_btn",
            "120": "120w_power_btn",
            "150": "150w_power_btn",
            "200": "200w_power_btn",
            "18": "18w_power_btn",
            "36": "36w_power_btn",
            "6": "6w_power_btn",
            "12": "12w_power_btn",
            "Unspecified": "unspecified_power_btn",
        }

        if device_power in device_power_map:
            image_key = device_power_map[device_power]
            if not self.find_and_click(image_key, timeout=10):
                return f"Thất bại: Could not find '{device_power}' button."
        else:
            return f"Thất bại: Invalid device power: {device_power}"

        return f"Device power selected: {device_power}"

auto_acs = AutoACSAutomation()

# --- 4. Main GUI ---
class AutoACSTool:
    def __init__(self, master):
        self.master = master
        master.title(APP_NAME)
        master.geometry("500x600")
        master.resizable(False, False)
        keyboard.add_hotkey('esc', self.stop_all_automation)

        self.dark_mode_colors = {
            'bg': '#1e1e1e',          # Nền tổng thể cửa sổ
            'fg': '#e0e0e0',          # Màu chữ chung (trắng nhạt)
            'button_bg': '#333333',   # Nền nút và tab không chọn
            'button_fg': '#ffffff',   # Màu chữ nút
            'entry_bg': '#2a2a2a',    # Nền ô nhập liệu và dropdown
            'entry_fg': '#ffffff',    # Màu chữ ô nhập liệu
            'frame_bg': '#282828',    # Nền khung, tab được chọn, nền status bar
            'select_bg': '#5c5c5c',   # Nền khi hover/chọn mục
            'select_fg': '#ffffff',   # Màu chữ khi hover/chọn mục
            'border': '#505050'       # Màu viền
        }

        self.master.overrideredirect(True) 
        
        self._create_custom_title_bar()

        taskbar_icon_path = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'), 'app.ico')
        if os.path.exists(taskbar_icon_path):
            try:
                self.master.iconbitmap(default=taskbar_icon_path)
                logger.info(f"Đã đặt icon cho Taskbar: {taskbar_icon_path}")
            except tk.TclError as e:
                logger.warning(f"Lỗi khi đặt icon Taskbar '{taskbar_icon_path}'. Đảm bảo là file .ico hợp lệ. Lỗi: {e}")
        else:
            logger.warning(f"Không tìm thấy file icon Taskbar tại: {taskbar_icon_path}.")

        master.configure(bg=self.dark_mode_colors['bg'])

        self.style = ttk.Style()
        self.style.theme_use('clam') 

        self.style.configure('.', font=('ZFVCutiegirl', 10), background=self.dark_mode_colors['bg'], foreground=self.dark_mode_colors['fg'])

        self.style.configure('TLabel', 
            background=self.dark_mode_colors['frame_bg'],
            foreground=self.dark_mode_colors['fg'],
            font=('ZFVCutiegirl', 10))

        self.style.configure('TButton', 
            background=self.dark_mode_colors['button_bg'],
            foreground=self.dark_mode_colors['button_fg'],
            font=('ZFVCutiegirl', 11),
            borderwidth=1,
            relief="flat")
        self.style.map('TButton',
            background=[('active', self.dark_mode_colors['select_bg'])],
            foreground=[('active', self.dark_mode_colors['fg'])])

        self.style.configure('TFrame', background=self.dark_mode_colors['frame_bg'])
        
        self.style.configure('TLabelframe', 
            background=self.dark_mode_colors['frame_bg'],
            foreground=self.dark_mode_colors['fg'],
            relief='solid',
            borderwidth=1,
            bordercolor=self.dark_mode_colors['border'])
        self.style.configure('TLabelframe.Label', 
            background=self.dark_mode_colors['frame_bg'],
            foreground=self.dark_mode_colors['fg'],
            font=('ZFVCutiegirl', 10, 'bold'))

        self.style.configure('TNotebook', 
            background=self.dark_mode_colors['bg'],
            borderwidth=0) # Loại bỏ viền Notebook
        self.style.configure('TNotebook.Tab', 
            background=self.dark_mode_colors['button_bg'], # Nền tab chưa chọn
            foreground=self.dark_mode_colors['fg'],      # Chữ tab chưa chọn
            font=('ZFVCutiegirl', 10),
            padding=[10, 5]) # Tăng padding cho tab
        self.style.map('TNotebook.Tab',
            background=[('selected', self.dark_mode_colors['frame_bg']), # Nền tab được chọn
                ('active', self.dark_mode_colors['select_bg'])], # Nền tab khi hover
            foreground=[('selected', self.dark_mode_colors['fg']),
                ('active', self.dark_mode_colors['fg'])]) # Chữ tab khi hover

        self.style.configure('TEntry', 
            fieldbackground=self.dark_mode_colors['entry_bg'], # Nền trường nhập
            foreground=self.dark_mode_colors['entry_fg'],     # Chữ trường nhập
            insertcolor=self.dark_mode_colors['fg'],          # Màu con trỏ
            borderwidth=1,
            relief="solid", # Viền rắn cho Entry
            bordercolor=self.dark_mode_colors['border'])

        self.style.configure('TCombobox', 
            fieldbackground=self.dark_mode_colors['entry_bg'], # Nền trường hiển thị
            foreground=self.dark_mode_colors['entry_fg'],     # Chữ trường hiển thị
            selectbackground=self.dark_mode_colors['select_bg'], # Nền mục được chọn trong list
            selectforeground=self.dark_mode_colors['select_fg'], # Chữ mục được chọn trong list
            background=self.dark_mode_colors['button_bg'],    # Nền của nút mũi tên
            arrowcolor=self.dark_mode_colors['fg'],           # Màu mũi tên
            borderwidth=1,
            relief="solid", # Viền rắn cho Combobox
            bordercolor=self.dark_mode_colors['border'])
            
        self.style.map('TCombobox',
            fieldbackground=[('readonly', self.dark_mode_colors['entry_bg'])],
            foreground=[('readonly', self.dark_mode_colors['entry_fg'])],
            background=[('active', self.dark_mode_colors['select_bg'])], # Màu nút mũi tên khi hover
            arrowcolor=[('active', self.dark_mode_colors['fg'])])

        self.style.configure('TCombobox.Listbox', 
            background=self.dark_mode_colors['entry_bg'], # Nền của danh sách thả xuống
            foreground=self.dark_mode_colors['entry_fg'], # Màu chữ của các mục trong danh sách
            selectbackground=self.dark_mode_colors['select_bg'], # Nền mục được chọn khi hover
            selectforeground=self.dark_mode_colors['select_fg'], # Chữ mục được chọn khi hover
            highlightbackground=self.dark_mode_colors['select_bg'], # Viền khi hover
            highlightcolor=self.dark_mode_colors['select_bg'],      # Màu highlight khi tập trung vào mục
            borderwidth=0,
            relief="flat")

        self.style.configure('Vertical.TScrollbar', # Dành cho thanh cuộn dọc
            background=self.dark_mode_colors['button_bg'], # Nền con trượt
            troughcolor=self.dark_mode_colors['frame_bg'], # Nền rãnh cuộn
            foreground=self.dark_mode_colors['fg'],       # Màu của con trượt (khi nó không ở trạng thái active)
            arrowcolor=self.dark_mode_colors['fg'],       # Màu mũi tên
            borderwidth=0,
            relief="flat")
        self.style.map('Vertical.TScrollbar',
            background=[('active', self.dark_mode_colors['select_bg'])], # Nền con trượt khi hover
            arrowcolor=[('active', self.dark_mode_colors['fg'])])

        self.style.configure('TSeparator', background=self.dark_mode_colors['border']) # Màu thanh phân cách

        self.style.configure('Large.TLabel', 
                             font=('ZFVCutiegirl', 24, 'bold'),
                             background=self.dark_mode_colors['frame_bg'],
                             foreground=self.dark_mode_colors['fg'])

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.create_acs_device_manager_tab()
        self.create_acs_device_configuration_tab()
        self.create_settings_tab()
        self.create_suki_tab()

        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self.master, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W, padding=(10, 5))
        self.status_label.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'], foreground=self.dark_mode_colors['fg'])
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.update_status("Sẵn sàng hoạt động.")
        self.update_excel_status()

    def create_acs_device_manager_tab(self):
        tab = ttk.Frame(self.notebook, padding="20")
        tab.configure(style='TFrame')
        self.notebook.add(tab, text="ACS Device Manager")

        btn_ghi_uid = ttk.Button(tab, text="Ghi UID", command=self.ghi_uid)
        btn_ghi_uid.configure(style='TButton')
        btn_ghi_uid.pack(pady=10, fill=tk.X, padx=5)

        device_type_frame = ttk.LabelFrame(tab, text="Device Type", padding="10")
        device_type_frame.configure(style='TLabelframe')
        device_type_frame.pack(fill=tk.X, padx=5, pady=5)

        self.device_type_var = tk.StringVar(value="AFVarionaut Pump")
        device_type_options = ["AFVarionaut Pump", "Submersible Pump", "Tricolor Led", "SingleColor Led", "Dmx2Vfd Converter"]
        self.device_type_dropdown = ttk.Combobox(device_type_frame, textvariable=self.device_type_var, values=device_type_options, state="readonly")
        self.device_type_dropdown.configure(style='TCombobox')
        self.device_type_dropdown.pack(pady=5, fill=tk.X, padx=5)
        self.device_type_dropdown.bind("<<ComboboxSelected>>", self.update_device_power_options)

        device_power_frame = ttk.LabelFrame(tab, text="Device Power (W)", padding="10")
        device_power_frame.configure(style='TLabelframe')
        device_power_frame.pack(fill=tk.X, padx=5, pady=5)

        self.device_power_var = tk.StringVar(value="60")
        self.device_power_options = ["60", "100", "140", "160"]
        self.device_power_dropdown = ttk.Combobox(device_power_frame, textvariable=self.device_power_var, values=self.device_power_options, state="readonly")
        self.device_power_dropdown.configure(style='TCombobox')
        self.device_power_dropdown.pack(pady=5, fill=tk.X, padx=5)

    def create_acs_device_configuration_tab(self):
        tab = ttk.Frame(self.notebook, padding="20")
        tab.configure(style='TFrame')
        self.notebook.add(tab, text="ACS Device Configuration")

        self.btn_ghi_dia_chi = ttk.Button(tab, text="Ghi địa chỉ", command=lambda: self.run_automation(auto_acs.ghi_dia_chi))
        self.btn_ghi_dia_chi.configure(style='TButton')
        self.btn_ghi_dia_chi.pack(pady=5, fill=tk.X, padx=5)

        self.btn_test = ttk.Button(tab, text="Test", command=lambda: self.run_automation(auto_acs.test))
        self.btn_test.configure(style='TButton')
        self.btn_test.pack(pady=5, fill=tk.X, padx=5)

        self.btn_ghi_dia_chi_test = ttk.Button(tab, text="Ghi địa chỉ & Test", command=lambda: self.run_automation(auto_acs.ghi_dia_chi_va_test))
        self.btn_ghi_dia_chi_test.configure(style='TButton')
        self.btn_ghi_dia_chi_test.pack(pady=5, fill=tk.X, padx=5)

        ttk.Separator(tab, orient='horizontal').pack(fill=tk.X, pady=10, padx=5)

        excel_frame = ttk.LabelFrame(tab, text="Danh sách địa chỉ", padding="10")
        excel_frame.configure(style='TLabelframe')
        excel_frame.pack(fill=tk.X, padx=5, pady=5)

        self.btn_import_excel = ttk.Button(excel_frame, text="Nhập File Excel (.xlsx)", command=self.import_excel_gui)
        self.btn_import_excel.configure(style='TButton')
        self.btn_import_excel.pack(pady=5, fill=tk.X, padx=5)

        self.excel_status_var = tk.StringVar()
        self.excel_status_label = ttk.Label(excel_frame, textvariable=self.excel_status_var, wraplength=300)
        self.excel_status_label.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
        self.excel_status_label.pack(pady=5, fill=tk.X, padx=5)


        manual_input_frame = ttk.Frame(excel_frame, padding=5)
        manual_input_frame.configure(style='TFrame')
        manual_input_frame.pack(fill=tk.X)

        label_no = ttk.Label(manual_input_frame, text="No.")
        label_no.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
        label_no.pack(side=tk.LEFT, padx=2)
        self.no_entry = ttk.Entry(manual_input_frame, width=5)
        self.no_entry.configure(style='TEntry')
        self.no_entry.pack(side=tk.LEFT, padx=2)
        self.no_entry.bind("<KeyRelease>", lambda event: self.schedule_update(delay=500, trigger="no"))

        label_pump = ttk.Label(manual_input_frame, text="Pump")
        label_pump.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
        label_pump.pack(side=tk.LEFT, padx=2)
        self.pump_entry = ttk.Entry(manual_input_frame, width=5)
        self.pump_entry.configure(style='TEntry')
        self.pump_entry.pack(side=tk.LEFT, padx=2)
        self.pump_entry.bind("<KeyRelease>", lambda event: self.schedule_update(delay=500, trigger="pump"))

        label_led = ttk.Label(manual_input_frame, text="Led")
        label_led.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
        label_led.pack(side=tk.LEFT, padx=2)
        self.led_entry = ttk.Entry(manual_input_frame, width=5)
        self.led_entry.configure(style='TEntry')
        self.led_entry.pack(side=tk.LEFT, padx=2)
        self.led_entry.bind("<KeyRelease>", lambda event: self.schedule_update(delay=500, trigger="led"))

    def create_settings_tab(self):
        tab = ttk.Frame(self.notebook, padding="20")
        tab.configure(style='TFrame')
        self.notebook.add(tab, text="Cài đặt")

        canvas = tk.Canvas(tab, highlightthickness=0, bg=self.dark_mode_colors['frame_bg'])
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview, style='Vertical.TScrollbar')
        scrollable_frame = ttk.Frame(canvas, borderwidth=0)
        scrollable_frame.configure(style='TFrame')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Dòng bắt đầu
        current_row = 0

        def add_label(text, bold=False, colspan=3, pady=(10, 5)):
            nonlocal current_row
            font = ("ZFVCutiegirl", 12, "bold") if bold else ("ZFVCutiegirl", 10)
            label = ttk.Label(scrollable_frame, text=text, font=font)
            label.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
            label.grid(row=current_row, column=0, columnspan=colspan, sticky=tk.W, pady=pady, padx=5)
            current_row += 1
            return label

        def add_entry(label_text, attr_name):
            nonlocal current_row
            label = ttk.Label(scrollable_frame, text=label_text)
            label.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
            label.grid(row=current_row, column=0, sticky=tk.W, pady=2, padx=5)
            entry = ttk.Entry(scrollable_frame, width=10)
            entry.configure(style='TEntry')
            entry.grid(row=current_row, column=1, sticky=tk.W, pady=2, padx=5)
            setattr(self, attr_name, entry)
            current_row += 1
            return entry

        def add_image_path(label_text, config_key):
            nonlocal current_row
            label = ttk.Label(scrollable_frame, text=label_text)
            label.configure(style='TLabel', background=self.dark_mode_colors['frame_bg'])
            label.grid(row=current_row, column=0, sticky=tk.W, pady=2, padx=5)
            entry = ttk.Entry(scrollable_frame, width=20)
            entry.configure(style='TEntry')
            entry.grid(row=current_row, column=1, sticky=tk.W, pady=2, padx=5)
            choose_btn = ttk.Button(scrollable_frame, text="Chọn",
                command=lambda: self.browse_image_path(entry, config_key), width=5)
            choose_btn.configure(style='TButton')
            choose_btn.grid(row=current_row, column=2, padx=5)
            setattr(self, f"{config_key}_path", entry)
            current_row += 1
            return entry

        # ---- Cài đặt chung ----
        add_label("Cài đặt chung:", bold=True)
        add_entry("Độ trễ chụp màn hình (giây):", "screenshot_delay_entry")
        add_entry("Độ trễ hành động (giây):", "action_delay_entry")
        add_entry("Độ tin cậy tìm ảnh (0.0-1.0):", "confidence_entry")

        # ---- Image Paths ----
        add_label("Đường dẫn ảnh (Image Paths):", bold=True)

        for key, label_text in [
            ("connected", "Connected!"),
            ("on_off_btn", "On/Off"),
            ("discover_btn", "Discover"),
            ("device_discovery_title", "Device Discovery title"),
            ("scan_btn", "Scan"),
            ("discovery_completed_text", "Discovery completed"),
            ("tricolor_led_item", "Tricolor Led"),
            ("afvarionaut_pump_item", "AFVarionaut Pump"),
            ("dmx_slave_address_field", "DMX Slave Address"),
            ("set_dmx_slave_address_btn", "Set DMX Slave Address"),
            ("run_dmx_test_btn", "Run DMX Test"),
            ("stop_dmx_test_btn", "Stop DMX Test"),
            ("dmx_slider_0", "DMX 0"),
            ("dmx_slider_0_2", "DMX 0_2"),
            ("dmx_slider_1", "DMX 1"),
            ("dmx_slider_2", "DMX 2"),
            ("dmx_slider_3", "DMX 3"),
            ("selec_all_1", "Uncheck select all"),
            ("selec_all_2", "Check select all"),
            ("acs_device_manager_title", "Acs device Manager title"),
            ("acs_device_configuration_title", "ACS Device Configuration title"),
            ("acs_device_configuration", "Acs device Configuration"),
            ("acs_device_manager_1", "Acs device Manager 1"),
            ("acs_device_manager_2", "Acs device Manager 2"),
            ("close_window_btn", "Close window"),
            ("load_btn", "Load"),
            ("adl_file", "Adl file"),
            ("open_adl_btn", "Open"),
            ("add_btn_1", "Add 1"),
            ("add_btn_2", "Add 2"),
            ("generate_btn", "Generate"),
            ("device_type_field", "Device type"),
            ("submersible_pump_type_btn", "Submersible Pump"),
            ("tricolor_led_type_btn", "Tricolor Led"),
            ("singlecolor_led_type_btn", "SingleColor Led"),
            ("dmx2vfd_converter_type_btn", "Dmx2Vfd Converter"),
            ("device_power_field", "Device power (W)"),
            ("_12w_power_btn", "12W"),
            ("_36w_power_btn", "36W"),
            ("_100w_power_btn", "100W"),
            ("_140w_power_btn", "140W"),
            ("_160w_power_btn", "160W"),
            ("_150w_power_btn", "150W"),
            ("_200w_power_btn", "200W"),
            ("write_btn", "Write"),
            ("save_btn", "Save"),
        ]:
            add_image_path(label_text, key)

        # ---- Nút lưu ----
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.configure(style='TFrame')
        button_frame.grid(row=current_row, column=0, columnspan=3, pady=10, padx=5)
        save_btn = ttk.Button(button_frame, text="Lưu Cài đặt", command=self.save_settings_from_gui)
        save_btn.configure(style='TButton')
        save_btn.pack(side=tk.LEFT, padx=5)

        self.load_settings_to_gui()

    def create_suki_tab(self):
        tab = ttk.Frame(self.notebook, padding="20")
        tab.configure(style='TFrame')
        self.notebook.add(tab, text="Suki UwU")

        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=0)
        tab.grid_rowconfigure(2, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=0)
        tab.grid_columnconfigure(2, weight=1)

        content_frame = ttk.Frame(tab, style='TFrame')
        content_frame.grid(row=1, column=1, sticky="nsew") 

        ttk.Label(content_frame, text=APP_NAME, style='Large.TLabel').pack(pady=(20, 10))

        ttk.Label(content_frame, text="Tự động Ghi UID, địa chỉ và test", style='TLabel').pack(pady=5)

        ttk.Label(content_frame, text=f"Version: {VERSION}", style='TLabel').pack(pady=5)

        author_frame = ttk.Frame(content_frame, style='TFrame')
        author_frame.pack(pady=5)
        ttk.Label(author_frame, text="Author: ", style='TLabel').pack(side="left")
        author_link = ttk.Label(author_frame, text="Suki", foreground="#db9aaa", cursor="hand2", style='TLabel')
        author_link.pack(side="left")
        author_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/Suki8898"))

        license_frame = ttk.Frame(content_frame, style='TFrame')
        license_frame.pack(pady=5)
        ttk.Label(license_frame, text="License: ", style='TLabel').pack(side="left")
        license_link = ttk.Label(license_frame, text="MIT", foreground="#db9aaa", cursor="hand2", style='TLabel')
        license_link.pack(side="left")
        license_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/Suki8898/ACS-Auto/blob/main/LICENSE"))

        try:
            image_path = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'image_folder'), 'Suki.png')
            
            pil_image = Image.open(image_path)
            
            max_width = 250
            max_height = 250
            
            original_width, original_height = pil_image.size
            
            if original_width > max_width or original_height > max_height:
                ratio = min(max_width / original_width, max_height / original_height)
                new_width = int(original_width * ratio)
                new_height = int(original_height * ratio)
                
                resized_pil_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            else:
                resized_pil_image = pil_image
            
            self.suki_image_ref = ImageTk.PhotoImage(resized_pil_image) 
            
            ttk.Label(content_frame, image=self.suki_image_ref, style='TLabel').pack(pady=(20, 10))
        except FileNotFoundError:
            logger.warning(f"Không tìm thấy hình ảnh Suki tại {image_path}. Bỏ qua hiển thị ảnh.")
            ttk.Label(content_frame, text="[Không tìm thấy hình ảnh Suki]", style='TLabel').pack(pady=(20, 10))
        except Exception as e:
            logger.error(f"Lỗi khi tải hoặc xử lý hình ảnh Suki: {e}", exc_info=True)
            ttk.Label(content_frame, text="[Lỗi tải hình ảnh Suki]", style='TLabel').pack(pady=(20, 10))


    def browse_image_path(self, entry_widget, config_key):
        image_folder_name = config_manager.get('GENERAL', 'image_folder')
        initial_dir = os.path.join(os.path.dirname(__file__), image_folder_name)
        if not os.path.exists(initial_dir):
            os.makedirs(initial_dir)

        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            title=f"Chọn ảnh cho {config_key} (Chỉ chọn 1 ảnh. Để thêm nhiều ảnh, gõ tên file vào ô, cách nhau bằng dấu phẩy)",
            filetypes=(("PNG files", "*.png"), ("All files", "*.*"))
        )
        if file_path:
            filename = os.path.basename(file_path)
            current_text = entry_widget.get()
            if current_text:
                if filename not in [f.strip() for f in current_text.split(',')]:
                    entry_widget.insert(tk.END, f", {filename}")
            else:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, filename)

    def load_settings_to_gui(self):
        self.screenshot_delay_entry.delete(0, tk.END)
        self.screenshot_delay_entry.insert(0, config_manager.get('GENERAL', 'screenshot_delay_sec'))
        self.action_delay_entry.delete(0, tk.END)
        self.action_delay_entry.insert(0, config_manager.get('GENERAL', 'action_delay_sec'))
        self.confidence_entry.delete(0, tk.END)
        self.confidence_entry.insert(0, config_manager.get('GENERAL', 'find_image_confidence'))

        def _set_image_entry(entry_widget, config_key):
            filenames_str = config_manager.get('IMAGE_PATHS', config_key)
            entry_widget.delete(0, tk.END)
            if filenames_str:
                entry_widget.insert(0, filenames_str)

        _set_image_entry(self.connected_path, 'connected')
        _set_image_entry(self.on_off_btn_path, 'on_off_btn')
        _set_image_entry(self.discover_btn_path, 'discover_btn')
        _set_image_entry(self.device_discovery_title_path,  'device_discovery_title')
        _set_image_entry(self.scan_btn_path, 'scan_btn')
        _set_image_entry(self.discovery_completed_text_path, 'discovery_completed_text')
        _set_image_entry(self.tricolor_led_item_path, 'tricolor_led_item')
        _set_image_entry(self.afvarionaut_pump_item_path, 'afvarionaut_pump_item')
        _set_image_entry(self.dmx_slave_address_field_path, 'dmx_slave_address_field')
        _set_image_entry(self.set_dmx_slave_address_btn_path, 'set_dmx_slave_address_btn')
        _set_image_entry(self.run_dmx_test_btn_path, 'run_dmx_test_btn')
        _set_image_entry(self.stop_dmx_test_btn_path, 'stop_dmx_test_btn')
        _set_image_entry(self.dmx_slider_0_path, 'dmx_slider_0')
        _set_image_entry(self.dmx_slider_0_2_path, 'dmx_slider_0_2')
        _set_image_entry(self.dmx_slider_1_path, 'dmx_slider_1')
        _set_image_entry(self.dmx_slider_2_path, 'dmx_slider_2')
        _set_image_entry(self.dmx_slider_3_path, 'dmx_slider_3')
        _set_image_entry(self.selec_all_1_path, 'selec_all_1')
        _set_image_entry(self.selec_all_2_path, 'selec_all_2')
        _set_image_entry(self.acs_device_manager_title_path, 'acs_device_manager_title')
        _set_image_entry(self.acs_device_configuration_title_path, 'acs_device_configuration_title')
        _set_image_entry(self.acs_device_configuration_path, 'acs_device_configuration')
        _set_image_entry(self.acs_device_manager_1_path, 'acs_device_manager_1')
        _set_image_entry(self.acs_device_manager_2_path, 'acs_device_manager_2')
        _set_image_entry(self.close_window_btn_path, 'close_window_btn')
        _set_image_entry(self.load_btn_path, 'load_btn')
        _set_image_entry(self.adl_file_path, 'adl_file')
        _set_image_entry(self.open_adl_btn_path, 'open_adl_btn')
        _set_image_entry(self.add_btn_1_path, 'add_btn_1')
        _set_image_entry(self.add_btn_2_path, 'add_btn_2')
        _set_image_entry(self.generate_btn_path, 'generate_btn')
        _set_image_entry(self.device_type_field_path, 'device_type_field')
        _set_image_entry(self.submersible_pump_type_btn_path, 'submersible_pump_type_btn')
        _set_image_entry(self.tricolor_led_type_btn_path, 'tricolor_led_type_btn')
        _set_image_entry(self.singlecolor_led_type_btn_path, 'singlecolor_led_type_btn')
        _set_image_entry(self.dmx2vfd_converter_type_btn_path, 'dmx2vfd_converter_type_btn')
        _set_image_entry(self.device_power_field_path, 'device_power_field')
        _set_image_entry(self._12w_power_btn_path, '_12w_power_btn')
        _set_image_entry(self._36w_power_btn_path, '_36w_power_btn')
        _set_image_entry(self._100w_power_btn_path, '_100w_power_btn')
        _set_image_entry(self._140w_power_btn_path, '_140w_power_btn')
        _set_image_entry(self._160w_power_btn_path, '_160w_power_btn')
        _set_image_entry(self._150w_power_btn_path, '_150w_power_btn')
        _set_image_entry(self._200w_power_btn_path, '_200w_power_btn')
        _set_image_entry(self.write_btn_path, 'write_btn')
        _set_image_entry(self.save_btn_path, 'save_btn')

    def save_settings_from_gui(self):
        try:
            config_manager.set('GENERAL', 'screenshot_delay_sec', self.screenshot_delay_entry.get())
            config_manager.set('GENERAL', 'action_delay_sec', self.action_delay_entry.get())
            config_manager.set('GENERAL', 'find_image_confidence', self.confidence_entry.get())
            
            config_manager.set('IMAGE_PATHS', 'connected', self.connected_path.get())
            config_manager.set('IMAGE_PATHS', 'on_off_btn', self.on_off_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'discover_btn', self.discover_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'device_discovery_title', self.device_discovery_title_path.get())
            config_manager.set('IMAGE_PATHS', 'scan_btn', self.scan_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'discovery_completed_text', self.discovery_completed_text_path.get())
            config_manager.set('IMAGE_PATHS', 'tricolor_led_item', self.tricolor_led_item_path.get())
            config_manager.set('IMAGE_PATHS', 'afvarionaut_pump_item', self.afvarionaut_pump_item_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx_slave_address_field', self.dmx_slave_address_field_path.get())
            config_manager.set('IMAGE_PATHS', 'set_dmx_slave_address_btn', self.set_dmx_slave_address_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'run_dmx_test_btn', self.run_dmx_test_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'stop_dmx_test_btn', self.stop_dmx_test_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx_slider_0', self.dmx_slider_0_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx_slider_0_2', self.dmx_slider_0_2_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx_slider_1', self.dmx_slider_1_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx_slider_2', self.dmx_slider_2_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx_slider_3', self.dmx_slider_3_path.get())
            config_manager.set('IMAGE_PATHS', 'selec_all_1', self.selec_all_1_path.get())
            config_manager.set('IMAGE_PATHS', 'selec_all_2', self.selec_all_2_path.get())  
            config_manager.set('IMAGE_PATHS', 'acs_device_manager_title', self.acs_device_manager_title_path.get())
            config_manager.set('IMAGE_PATHS', 'acs_device_configuration_title', self.acs_device_configuration_title_path.get())
            config_manager.set('IMAGE_PATHS', 'acs_device_configuration', self.acs_device_configuration_path.get())
            config_manager.set('IMAGE_PATHS', 'acs_device_manager_1', self.acs_device_manager_1_path.get())
            config_manager.set('IMAGE_PATHS', 'acs_device_manager_2', self.acs_device_manager_2_path.get())
            config_manager.set('IMAGE_PATHS', 'close_window_btn', self.close_window_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'load_btn', self.load_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'adl_file', self.adl_file_path.get())
            config_manager.set('IMAGE_PATHS', 'open_adl_btn', self.open_adl_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'add_btn_1', self.add_btn_1_path.get())
            config_manager.set('IMAGE_PATHS', 'add_btn_2', self.add_btn_2_path.get())
            config_manager.set('IMAGE_PATHS', 'generate_btn', self.generate_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'device_type_field', self.device_type_field_path.get())
            config_manager.set('IMAGE_PATHS', 'submersible_pump_type_btn', self.submersible_pump_type_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'tricolor_led_type_btn', self.tricolor_led_type_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'singlecolor_led_type_btn', self.singlecolor_led_type_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'dmx2vfd_converter_type_btn', self.dmx2vfd_converter_type_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'device_power_field', self.device_power_field_path.get())
            config_manager.set('IMAGE_PATHS', '12w_power_btn', self._12w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', '36w_power_btn', self._36w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', '100w_power_btn', self._100w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', '140w_power_btn', self._140w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', '160w_power_btn', self._160w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', '150w_power_btn', self._150w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', '200w_power_btn', self._200w_power_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'write_btn', self.write_btn_path.get())
            config_manager.set('IMAGE_PATHS', 'save_btn', self.save_btn_path.get())

            config_manager.save_config()
            self.update_status("Thông báo: Cài đặt đã được lưu!")


        except Exception as e:
            logger.error(f"Error saving settings: {e}", exc_info=True)
            self.update_status(f"Lỗi: Không thể lưu cài đặt: {e}")


    def browse_image_path(self, entry_widget, config_key):
        image_folder_name = config_manager.get('GENERAL', 'image_folder')
        initial_dir = os.path.join(os.path.dirname(__file__), image_folder_name)
        if not os.path.exists(initial_dir):
            os.makedirs(initial_dir)

        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            title=f"Chọn ảnh cho {config_key} (Chỉ chọn 1 ảnh. Để thêm nhiều ảnh, gõ tên file vào ô, cách nhau bằng dấu phẩy)",
            filetypes=(("PNG files", "*.png"), ("All files", "*.*"))
        )
        if file_path:
            filename = os.path.basename(file_path)
            current_text = entry_widget.get()
            if current_text:
                if filename not in [f.strip() for f in current_text.split(',')]:
                    entry_widget.insert(tk.END, f", {filename}")
            else:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, filename)

    def stop_all_automation(self):
        try:

            auto_acs.stop_requested = True
            self.update_status("🛑 Đã dừng.")
            logger.warning("Dừng toàn bộ quá trình!")

        except Exception as e:
            logger.error(f"Lỗi khi dừng bằng ESC: {e}")


    def ghi_uid(self):
        logger.info("Bắt đầu quy trình 'Ghi UID'...")

        auto_acs.stop_requested = False

        selected_device_type = self.device_type_var.get()
        selected_device_power = self.device_power_var.get()

        results = []

        if auto_acs.find('add_btn_1', timeout=0.2):
            if not auto_acs.find_and_click('load_btn', timeout=0.2):
                results.append("Thất bại: Could not find 'load_btn' button.")

            if not auto_acs.find_and_click('adl_file', timeout=1):
                results.append("Thất bại: Could not find '.adl file' button.")

            if not auto_acs.find_and_click('open_adl_btn', timeout=1):
                results.append("Thất bại: Could not find 'Open' button.")
        else:
            pass


        if not auto_acs.find_and_click('add_btn_2', timeout=0.2):
            results.append("Thất bại: Could not find 'Add' button.")
                

        if not auto_acs.find_and_click('generate_btn', timeout=10):
            results.append("Thất bại: Could not find 'Generate' button.")

        # Bỏ qua bước chọn Device Type nếu là "Afvarionaut Pump"
        if selected_device_type != "AFVarionaut Pump":
            device_type_result = auto_acs._select_device_type(selected_device_type)
            results.append(device_type_result)
        else:
            logger.info("Device type is Afvarionaut Pump, skipping Device Type selection.")
            results.append("Device type is Afvarionaut Pump, skipping Device Type selection.")

        # Bỏ qua bước chọn Device Power nếu là "6", "18", "60", "120", hoặc "Unspecified"
        skip_power_selection = selected_device_power in ("6", "18", "60", "120", "Unspecified")
        if not skip_power_selection:
            device_power_result = auto_acs._select_device_power(selected_device_power)
            results.append(device_power_result)
        else:
            logger.info(f"Device power is {selected_device_power}, skipping Device Power selection.")
            results.append(f"Device power is {selected_device_power}, skipping Device Power selection.")

        if not auto_acs.find_and_click('write_btn', timeout=10):
            results.append("Thất bại: Could not find 'Write' button.")

        if not auto_acs.find_and_click('save_btn', timeout=10):
            results.append("Thất bại: Could not find 'Save' button.")

        if "Failed" in "".join(results):
            self.update_status("Lỗi: Ghi UID thất bại. Vui lòng kiểm tra log.")
            return

        logger.info(f"Ghi UID thành công. Kết quả: {'; '.join(results)}")
        self.update_status("Thông báo: Ghi UID thành công!")

    def update_device_power_options(self, event=None):
        selected_device = self.device_type_var.get()

        if selected_device == "AFVarionaut Pump":
            self.device_power_options = ["60", "100", "140", "160"]
        elif selected_device == "Submersible Pump":
            self.device_power_options = ["120", "150", "200"]
        elif selected_device == "Tricolor Led":
            self.device_power_options = ["18", "36"]
        elif selected_device == "SingleColor Led":
            self.device_power_options = ["6", "12"]
        elif selected_device == "Dmx2Vfd Converter":
            self.device_power_options = ["Unspecified"]

        self.device_power_dropdown['values'] = self.device_power_options

        if self.device_power_var.get() not in self.device_power_options:
            self.device_power_var.set(self.device_power_options[0])

    def schedule_update(self, delay, trigger):
        if hasattr(self, 'after_id') and self.after_id:
            self.master.after_cancel(self.after_id)
        self.after_id = self.master.after(delay, lambda: self.delayed_update(trigger))

    def delayed_update(self, trigger):
        try:
            if trigger == "no":
                row_number = int(self.no_entry.get()) - 1 if self.no_entry.get() else None
            elif trigger == "pump":
                pump_value = int(self.pump_entry.get()) if self.pump_entry.get() else None
                row_number = next((i for i, row in enumerate(auto_acs.excel_data) if row['Pump'] == pump_value), None) if pump_value else None
            elif trigger == "led":
                led_value = int(self.led_entry.get()) if self.led_entry.get() else None
                row_number = next((i for i, row in enumerate(auto_acs.excel_data) if row['Led'] == led_value), None) if led_value else None
            else:
                return
            
            if row_number is not None and auto_acs.excel_data and 0 <= row_number < len(auto_acs.excel_data):
                auto_acs.current_excel_row_index = row_number
                self.update_excel_status()
                self.update_entry_fields(row_number)
            elif row_number is not None and not auto_acs.excel_data:
                self.update_status("Lỗi: Không có dữ liệu Excel để tìm kiếm.")
            elif row_number is not None:
                self.update_status("Lỗi: Số hàng không hợp lệ hoặc không tìm thấy.")
        except ValueError:
            pass
        except Exception as e:
            self.update_status(f"Lỗi: Có lỗi xảy ra: {e}")
            logger.error(f"Error in delayed_update: {e}", exc_info=True)
    
    def update_entry_fields(self, row_number):
        self.no_entry.delete(0, tk.END)
        self.pump_entry.delete(0, tk.END)
        self.led_entry.delete(0, tk.END)
        
        try:
            if auto_acs.excel_data and 0 <= row_number < len(auto_acs.excel_data):
                row = auto_acs.excel_data[row_number]
                self.no_entry.insert(0, str(row_number + 1))
                self.pump_entry.insert(0, str(row['Pump']))
                self.led_entry.insert(0, str(row['Led']))
        except Exception as e:
            logger.error(f"Không thể lấy dữ liệu hàng {row_number}: {e}")

    def update_status(self, message):
        self.status_var.set(message)
        logger.info(f"STATUS: {message}")

    def update_excel_status(self):
        if auto_acs.excel_data:
            total_rows = len(auto_acs.excel_data)
            current_row = auto_acs.current_excel_row_index
            self.excel_status_var.set(f"Đang ghi hàng: {current_row + 1}/{total_rows}")
        else:
            self.excel_status_var.set("Chưa có file Excel được nhập.")

    def import_excel_gui(self):
        file_path = filedialog.askopenfilename(
            title="Chọn File Excel (.xlsx)",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            if auto_acs.import_excel_data(file_path):
                self.update_status("Thông báo: Nhập Excel thành công!")
                self.update_excel_status()
                self.update_entry_fields(0)
            else:
                self.update_status("Lỗi: Không thể nhập file Excel. Vui lòng kiểm tra log.")
                self.update_excel_status()

    def reset_excel_index_gui(self):
        auto_acs.reset_excel_row_index()
        self.update_excel_status()
        self.update_status("Thông báo: Hàng Excel đã được reset về 0.")
        self.update_entry_fields(0)


    def run_automation(self, func):
        if hasattr(self, 'automation_thread') and self.automation_thread.is_alive():
            self.update_status("Cảnh báo: Một tác vụ tự động hóa đang chạy. Vui lòng đợi hoặc khởi động lại tool.")
            return

        self.update_status(f"Đang chạy '{func.__name__.replace('_', ' ').title()}'...")
        self.set_buttons_state(tk.DISABLED)

        self.automation_thread = threading.Thread(target=self._automation_task, args=(func,))
        self.automation_thread.start()

    def _automation_task(self, func):
        try:
            if func in [auto_acs.ghi_dia_chi, auto_acs.ghi_dia_chi_va_test] and (auto_acs.excel_data is None or not auto_acs.excel_data):
                self.update_status("Lỗi: Vui lòng nhập file Excel trước khi chạy chức năng này.")
                self.set_buttons_state(tk.NORMAL)
                return

            result = func()
            self.update_status(f"Hoàn thành: {result}")
        except Exception as e:
            logger.error(f"Lỗi trong quá trình tự động hóa: {e}", exc_info=True)
            self.update_status(f"Lỗi: Có lỗi xảy ra: {e}. Vui lòng kiểm tra log.txt.")
        finally:
            self.set_buttons_state(tk.NORMAL)
            self.update_excel_status()
            self.update_entry_fields(auto_acs.current_excel_row_index)

    def set_buttons_state(self, state):
        self.btn_ghi_dia_chi.config(state=state)
        self.btn_test.config(state=state)
        self.btn_ghi_dia_chi_test.config(state=state)
        self.btn_import_excel.config(state=state)

    def _create_custom_title_bar(self):
        self.title_bar = tk.Frame(self.master, 
                                  bg=self.dark_mode_colors['frame_bg'], 
                                  relief="raised", 
                                  bd=0, 
                                  height=30)
        self.title_bar.pack(side="top", fill="x")

        icon_path = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'), 'app.ico')
        if os.path.exists(icon_path):
            try:
                pil_icon = Image.open(icon_path)
                pil_icon = pil_icon.resize((24, 24), Image.Resampling.LANCZOS)
                self.title_icon_img = ImageTk.PhotoImage(pil_icon)
                icon_label = tk.Label(self.title_bar, 
                    image=self.title_icon_img, 
                    bg=self.dark_mode_colors['frame_bg'])
                icon_label.pack(side="left", padx=5, pady=2)
            except Exception as e:
                logger.warning(f"Lỗi khi tải icon cho custom title bar '{icon_path}': {e}")
        
        title_label = tk.Label(self.title_bar, 
            text=APP_NAME, 
            bg=self.dark_mode_colors['frame_bg'], 
            fg=self.dark_mode_colors['fg'], 
            font=('ZFVCutiegirl', 10, 'bold'))
        title_label.pack(side="left", padx=5, pady=2, expand=True, fill="x")

        self.close_button = tk.Button(self.title_bar, 
            text="❌", 
            command=self.master.destroy,
            bg=self.dark_mode_colors['frame_bg'],
            fg=self.dark_mode_colors['fg'],
            bd=0, 
            activebackground=self.dark_mode_colors['select_bg'],
            activeforeground=self.dark_mode_colors['fg'],
            highlightthickness=0,
            width=4, height=1,
            font=('Arial', 11, 'bold'))
        self.close_button.pack(side="right")
        
        self.close_button.bind("<Enter>", self._on_close_button_enter)
        self.close_button.bind("<Leave>", self._on_close_button_leave)

        self.title_bar.bind("<ButtonPress-1>", self._start_move_window)
        self.title_bar.bind("<B1-Motion>", self._move_window)
        title_label.bind("<ButtonPress-1>", self._start_move_window)
        title_label.bind("<B1-Motion>", self._move_window)
        if 'icon_label' in locals():
            icon_label.bind("<ButtonPress-1>", self._start_move_window)
            icon_label.bind("<B1-Motion>", self._move_window)

    def _start_move_window(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def _move_window(self, event):
        delta_x = event.x - self.start_x
        delta_y = event.y - self.start_y
        new_x = self.master.winfo_x() + delta_x
        new_y = self.master.winfo_y() + delta_y
        self.master.geometry(f"+{new_x}+{new_y}")

    def _on_close_button_enter(self, event):

        self.close_button.config(bg="#db9aaa", fg="#ffffff") 

    def _on_close_button_leave(self, event):
        self.close_button.config(bg=self.dark_mode_colors['frame_bg'], fg=self.dark_mode_colors['fg'])

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoACSTool(root)
    root.mainloop()