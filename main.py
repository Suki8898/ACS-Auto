import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import webbrowser
import logging
import configparser
import pyautogui
import time
import pandas as pd
from PIL import Image, ImageTk, ImageDraw
import keyboard
import colorsys
import json
import copy
import cv2
import random

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

APP_NAME = "ACS Auto"
VERSION = "2.3.0"
ACCENT_COLOR = "#c48b9a"
HOVER_COLOR = "#4a4a4a"
HOVER_2_COLOR = "#db9aaa"
DARK_BG = "#1e1e1e"
FRAME_BG = "#2b2b2b"
BUTTON_BG = "#3a3a3a"
MAIN_FONT = "ZFVCutiegirl"

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
            'device_discovery_title_1': 'device_discovery_title_1.png',
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
            'list': 'List.png',
            'add_btn': 'Add.png',
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
            'successfully_text': 'Successfully text.png',
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


pyautogui.FAILSAFE = True

class IconManager:
    def __init__(self):
        self.image_folder = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'))

        self.add = self.load_icon("add.png", size=(20, 20))
        self.edit = self.load_icon("edit.png", size=(20, 20))
        self.delete = self.load_icon("delete.png", size=(20, 20))
        self.save = self.load_icon("save.png", size=(20, 20))
        self.photo = self.load_icon("photo.png", size=(20, 20))
        self.up = self.load_icon("up.png", size=(20, 20))
        self.down = self.load_icon("down.png", size=(20, 20))
        self.file = self.load_icon("file.png", size=(20, 20))
        self.excel = self.load_icon("excel.png", size=(20, 20))
        self.code = self.load_icon("code.png", size=(20, 20))
        self.check = self.load_icon("check.png", size=(20, 20))
        self.x = self.load_icon("x.png", size=(20, 20))

    def load_icon(self, filename, size=(20, 20)):
        try:
            path = os.path.join(self.image_folder, filename)
            if os.path.exists(path):
                pil_image = Image.open(path)
                return ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=size)
            else:
                return None
        except Exception as e:
            logger.error(f"Lỗi tải icon {filename}: {e}")
            return None

icons = IconManager()

class ACSAutomation:
    def __init__(self):
        self.icon_folder = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'))
        self.image_folder = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'image_folder'))
        self.screenshot_delay = float(config_manager.get('GENERAL', 'screenshot_delay_sec'))
        self.action_delay = float(config_manager.get('GENERAL', 'action_delay_sec'))
        self.confidence = float(config_manager.get('GENERAL', 'find_image_confidence'))

        self.excel_data = None
        self.current_excel_row_index = 0

        self.stop_requested = False
        self.enable_auto_increment = True

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
            logger.warning(f"Không tìm thấy file hình ảnh cho '{image_name_key}'.")
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
    
    
    def import_excel_data(self, file_path):
        try:
            df = pd.read_excel(file_path, header=0)
            df = df.iloc[:, [0, 1, 2, 3]]
            df.columns = ['No.', 'Pump', 'Led', 'Dmx2Vfd']
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
        if not self.enable_auto_increment:
            return False
        
        if self.excel_data and self.current_excel_row_index < len(self.excel_data):
            self.current_excel_row_index += 1
            logger.info(f"Excel đã tăng 1 hàng {self.current_excel_row_index}")
            return True
        return False

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
        dmx2vfd_locations = []

        led_image_paths = self._get_image_paths_list('tricolor_led_item')
        pump_image_paths = self._get_image_paths_list('afvarionaut_pump_item')
        dmx2vfd_image_paths = self._get_image_paths_list('dmx2vfd_item')

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
            
            current_iteration_dmx2vfd_locs = []
            for dmx2vfd_path in dmx2vfd_image_paths:
                try:
                    if search_region:
                        found_locs = list(pyautogui.locateAllOnScreen(dmx2vfd_path, confidence=locate_confidence, region=search_region))
                    else:
                        found_locs = list(pyautogui.locateAllOnScreen(dmx2vfd_path, confidence=locate_confidence))
                    if found_locs:
                        current_iteration_dmx2vfd_locs.extend(found_locs)
                except pyautogui.ImageNotFoundException:
                    logger.debug(f"Hình ảnh DMX2VFD '{os.path.basename(dmx2vfd_path)}' không tìm thấy, đang thử tiếp nếu có sẵn.")
                    pass
                except Exception as e:
                    logger.error(f"Lỗi khi xác định vị trí hình ảnh DMX2VFD '{os.path.basename(dmx2vfd_path)}': {e}", exc_info=True)
                    pass

            if current_iteration_led_locs or current_iteration_pump_locs or current_iteration_dmx2vfd_locs:
                logger.info(f"Tìm thấy {len(current_iteration_led_locs)} LED(s) , {len(current_iteration_pump_locs)} PUMP(s) và {len(current_iteration_dmx2vfd_locs)} DMX2VFD(s) trong Device Discovery.")
                return current_iteration_led_locs, current_iteration_pump_locs, current_iteration_dmx2vfd_locs

            time.sleep(0.5)
                
        logger.warning("Không tìm thấy thiết bị LED, PUMP và DMX2VFD trong Device Discovery trong thời gian chờ.")
        return [], [], []

    def chon_thiet_bi(self, device_type_name, device_location):
        logger.info(f"Đang xử lý {device_type_name} tại {pyautogui.center(device_location)}")
        center = pyautogui.center(device_location)
        pyautogui.doubleClick(center.x, center.y, interval=0.1)
        time.sleep(self.action_delay)
        
    def chon_thiet_bi_va_ghi(self, device_type_name, device_location, address):
        logger.info(f"Đang xử lý {device_type_name} tại {pyautogui.center(device_location)} với địa chỉ {address}")
        center = pyautogui.center(device_location)
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
        return self.wait_for_image(image_name_key, timeout, confidence_override)

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
        return self.wait_for_image(slider_moved_image_key, timeout, confidence_override)

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
        logger.info(f"Đang chọn thiết bị: {device_type}")
        if not self.find_and_click('device_type_field', timeout=10):
            return f"Thất bại: không thể tìm thấy 'Device type'."
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
                return f"Thất bại: không thể tìm thấy nút '{device_type}'."
        else:
            return f"Thất bại: Invalid device type: {device_type}"
        return f"Device type selected: {device_type}"
    
    def _select_device_power(self, device_power):
        logger.info(f"Đang chọn công suất: {device_power}")
        if not self.find_and_click('device_power_field', timeout=10):
            return f"Thất bại: không thể tìm thấy 'Device power'."
        device_power_map = {
            "60": "60w_power_btn", "100": "100w_power_btn", "140": "140w_power_btn",
            "160": "160w_power_btn", "120": "120w_power_btn", "150": "150w_power_btn",
            "200": "200w_power_btn", "18": "18w_power_btn", "36": "36w_power_btn",
            "6": "6w_power_btn", "12": "12w_power_btn", "Unspecified": "unspecified_power_btn",
        }
        if device_power in device_power_map:
            image_key = device_power_map[device_power]
            if not self.find_and_click(image_key, timeout=10):
                return f"Thất bại: không thể tìm thấy '{device_power}'."
        else:
            return f"Thất bại: Invalid device power: {device_power}"
        return f"Device power selected: {device_power}"

    def click_acs_device_configuration(self, x, y):
        target_title = "ACS Device Configuration - Version 1.5.0"
        try:
            windows = pyautogui.getWindowsWithTitle(target_title)
            if not windows:
                logger.warning(f"Không tìm thấy cửa sổ: '{target_title}'")
                return
            window = windows[0]
            if window.isMinimized:
                window.restore()
            abs_x = window.left + x
            abs_y = window.top + y
            try:
                window.activate()
            except Exception:
                pass
            logger.info(f"Đang click tại ({x}, {y}) trên cửa sổ ACS...")
            pyautogui.click(abs_x, abs_y)
        except Exception as e:
            logger.error(f"Lỗi khi click tọa độ: {e}")
    

acs_auto = ACSAutomation()

class AutoACSTool(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("500x650")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        
        self.bind("<Map>", lambda e: self.after(50, lambda: (self.attributes("-topmost", True), self.lift())))

        keyboard.add_hotkey('esc', self.stop_all_automation)
        keyboard.add_hotkey('f1', lambda: self.execute_category_script("uid_col1", self.get_uid_col1_context))
        keyboard.add_hotkey('f2', lambda: self.execute_category_script("uid_col2", self.get_uid_col2_context))
        keyboard.add_hotkey('f3', lambda: self.execute_category_script("address", self.get_excel_context))
        keyboard.add_hotkey('f4', lambda: self.execute_category_script("test", self.get_excel_context))
        keyboard.add_hotkey('f5', lambda: self.execute_category_script("address_test", self.get_excel_context))
        keyboard.add_hotkey('f7', lambda: self.click_acs_device_configuration(520, 220))
        keyboard.add_hotkey('f8', lambda: self.click_acs_device_configuration(560, 220))
        keyboard.add_hotkey('f9', lambda: self.click_acs_device_configuration(610, 220))
        keyboard.add_hotkey('f10', lambda: self.click_acs_device_configuration(650, 220))

        self.overrideredirect(True) 
        self.create_custom_title_bar()
        self.configure(fg_color=DARK_BG)

        self.tabview = ctk.CTkTabview(self, width=480, height=580, fg_color=FRAME_BG, segmented_button_selected_color=ACCENT_COLOR, segmented_button_selected_hover_color=HOVER_2_COLOR) 
        self.tabview.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tabview._segmented_button.configure(font=(MAIN_FONT, 12, "bold"))
        
        self.tabview.add("ACS Device Manager")
        self.tabview.add("ACS Device Configuration")
        self.tabview.add("Cài đặt")
        self.tabview.add("Suki UwU")

        self.create_acs_device_manager_tab()
        self.create_acs_device_configuration_tab()
        self.create_settings_tab()
        self.create_suki_tab()

        logger.info("Sẵn sàng hoạt động.")
        self.update_excel_status()

    def click_acs_device_configuration(self, x, y):
        target_title = "ACS Device Configuration - Version 1.5.0"
        try:
            windows = pyautogui.getWindowsWithTitle(target_title)
            if not windows:
                logger.warning(f"Không tìm thấy cửa sổ: '{target_title}'")
                return
            window = windows[0]
            if window.isMinimized:
                window.restore()
            abs_x = window.left + x
            abs_y = window.top + y
            try:
                window.activate()
            except Exception:
                pass
            logger.info(f"Đang click tại ({x}, {y}) trên cửa sổ ACS...")
            pyautogui.click(abs_x, abs_y)
        except Exception as e:
            logger.error(f"Lỗi khi click tọa độ F7-F10: {e}")
    
    def get_excel_context(self):
        if not acs_auto.excel_data:
            logger.warning("Vui lòng nhập file Excel trước.")
            return None
        row_data = acs_auto.get_current_excel_row()
        if not row_data:
            logger.warning("Không lấy được dữ liệu hàng hiện tại.")
            return None
        return {
            'pump_address': row_data.get('Pump'),
            'led_address': row_data.get('Led'),
            'dmx2vfd_address': row_data.get('Dmx2Vfd'),
            'current_row_index': acs_auto.current_excel_row_index,
            'total_rows': len(acs_auto.excel_data)
        }

    def get_uid_col1_context(self):
        return {
            'selected_device_type': self.device_type_var_col1.get(),
            'selected_device_power': self.device_power_var_col1.get()
        }

    def get_uid_col2_context(self):
        return {
            'selected_device_type': self.device_type_var_col2.get(),
            'selected_device_power': self.device_power_var_col2.get()
        }

    def execute_category_script(self, category, context_func=None):
        if hasattr(self, 'automation_thread') and self.automation_thread.is_alive():
            logger.warning("Một tác vụ đang chạy. Vui lòng đợi.")
            return

        active_script = script_manager.get_active_script(category)
        if not active_script:
            messagebox.showinfo("Thông báo", f"Chưa có kịch bản nào được chọn (Active) cho '{category}'. Hãy bấm nút 📄 để chọn.")
            return

        steps = active_script['steps']
        script_name = active_script['name']

        extra_context = {}
        if context_func:
            extra_context = context_func()
            if extra_context is None:
                return

        logger.info(f"🚀 Đang chạy kịch bản: {script_name}")
        self.set_buttons_state("disabled")

        suki_path = os.path.join(acs_auto.image_folder, "Working.mp4")
        self.player_manager.start_active(suki_path) 

        self.automation_thread = threading.Thread(
            target=self._run_dynamic_script, 
            args=(steps, extra_context)
        )
        self.automation_thread.start()

    def _run_dynamic_script(self, steps, extra_context):
        results = []
        try:
            context = {
                'acs_auto': acs_auto,
                'logger': logger,
                'time': time,
                'pyautogui': pyautogui,
                'results': results,
                'script_stop': False, 
                **extra_context
            }

            for i, step in enumerate(steps):
                if acs_auto.stop_requested:
                    logger.warning("🛑 Đã dừng kịch bản.")
                    results.append("Đã dừng bởi người dùng.")
                    break
                if context.get('script_stop', False):
                    logger.info("⏹ Kịch bản dừng lại do lệnh 'script_stop = True'.")
                    break

                step_name = step.get('name', f'Step {i+1}')
                code_str = step.get('code', 'pass')

                logger.info(f"▶ Step: {step_name}")
                try:
                    exec(code_str, globals(), context)
                    if context.get('script_stop', False):
                        continue 
                except Exception as e:
                    err_msg = f"❌ Lỗi '{step_name}': {e}"
                    logger.error(err_msg)
                    results.append(err_msg)

            final_msg = " | ".join([str(r) for r in results])
            if final_msg:
                logger.info(f"Kết quả: {final_msg}")

        except Exception as e:
            logger.error(f"Lỗi hệ thống script: {e}", exc_info=True)
        finally:
            self.after(0, self.player_manager.start_idle)
            self.after(0, self.player_config.start_idle)
            self.after(0, lambda: self.set_buttons_state("normal"))
            self.after(0, self.update_excel_status)
            self.after(0, lambda: self.update_entry_fields(acs_auto.current_excel_row_index))

    def create_acs_device_manager_tab(self):
        tab = self.tabview.tab("ACS Device Manager")

        cols_frame = ctk.CTkFrame(tab, fg_color="transparent")
        cols_frame.pack(fill="both", padx=5, pady=5, expand=True)

        left_col = ctk.CTkFrame(cols_frame, fg_color="transparent")
        left_col.pack(side="left", fill="both", expand=True, padx=(0,5))

        btn_frame_1 = ctk.CTkFrame(left_col, fg_color="transparent")
        btn_frame_1.pack(fill="x", padx=5)
        
        self.btn_uid_col1 = ctk.CTkButton(btn_frame_1, text="Ghi UID (F1)", fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14, "bold"),
            command=lambda: self.execute_category_script("uid_col1", self.get_uid_col1_context))
        self.btn_uid_col1.pack(side="left", fill="x", expand=True)
        
        btn_script_1 = ctk.CTkButton(btn_frame_1, text="", image=icons.file, compound="left", width=40, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14),
            command=lambda: ScriptSelector(self, "uid_col1", lambda: None))
        btn_script_1.pack(side="left", padx=(5,0))

        device_type_frame_1 = ctk.CTkFrame(left_col)
        device_type_frame_1.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(device_type_frame_1, text="Device Type", font=(MAIN_FONT, 12, "bold")).pack(anchor="w", padx=10, pady=(5,0))

        self.device_type_var_col1 = ctk.StringVar(value="AFVarionaut Pump")
        device_type_options = ["AFVarionaut Pump", "Submersible Pump", "Tricolor Led", "SingleColor Led", "Dmx2Vfd Converter"]
        self.device_type_dropdown_col1 = ctk.CTkComboBox(device_type_frame_1, variable=self.device_type_var_col1, values=device_type_options, command=lambda e: self.update_device_power_options_col(1), font=(MAIN_FONT, 12), dropdown_font=(MAIN_FONT, 12))
        self.device_type_dropdown_col1.pack(pady=5, fill="x", padx=10)

        device_power_frame_1 = ctk.CTkFrame(left_col)
        device_power_frame_1.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(device_power_frame_1, text="Device Power (W)", font=(MAIN_FONT, 12, "bold")).pack(anchor="w", padx=10, pady=(5,0))

        self.device_power_var_col1 = ctk.StringVar(value="60")
        self.device_power_options_col1 = ["60", "100", "140", "160"]
        self.device_power_dropdown_col1 = ctk.CTkComboBox(device_power_frame_1, variable=self.device_power_var_col1, values=self.device_power_options_col1, font=(MAIN_FONT, 12), dropdown_font=(MAIN_FONT, 12))
        self.device_power_dropdown_col1.pack(pady=5, fill="x", padx=10)

        right_col = ctk.CTkFrame(cols_frame, fg_color="transparent")
        right_col.pack(side="left", fill="both", expand=True, padx=(5,0))

        btn_frame_2 = ctk.CTkFrame(right_col, fg_color="transparent")
        btn_frame_2.pack(fill="x", padx=5)

        self.btn_uid_col2 = ctk.CTkButton(btn_frame_2, text="Ghi UID (F2)", fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14, "bold"),
            command=lambda: self.execute_category_script("uid_col2", self.get_uid_col2_context))
        self.btn_uid_col2.pack(side="left", fill="x", expand=True)

        btn_script_2 = ctk.CTkButton(btn_frame_2, text="", image=icons.file, compound="left", width=40, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14),
            command=lambda: ScriptSelector(self, "uid_col2", lambda: None))
        btn_script_2.pack(side="left", padx=(5,0))

        device_type_frame_2 = ctk.CTkFrame(right_col)
        device_type_frame_2.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(device_type_frame_2, text="Device Type", font=(MAIN_FONT, 12, "bold")).pack(anchor="w", padx=10, pady=(5,0))

        self.device_type_var_col2 = ctk.StringVar(value="AFVarionaut Pump")
        self.device_type_dropdown_col2 = ctk.CTkComboBox(device_type_frame_2, variable=self.device_type_var_col2, values=device_type_options, command=lambda e: self.update_device_power_options_col(2), font=(MAIN_FONT, 12), dropdown_font=(MAIN_FONT, 12))
        self.device_type_dropdown_col2.pack(pady=5, fill="x", padx=10)

        device_power_frame_2 = ctk.CTkFrame(right_col)
        device_power_frame_2.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(device_power_frame_2, text="Device Power (W)", font=(MAIN_FONT, 12, "bold")).pack(anchor="w", padx=10, pady=(5,0))

        self.device_power_var_col2 = ctk.StringVar(value="60")
        self.device_power_options_col2 = ["60", "100", "140", "160"]
        self.device_power_dropdown_col2 = ctk.CTkComboBox(device_power_frame_2, variable=self.device_power_var_col2, values=self.device_power_options_col2, font=(MAIN_FONT, 12), dropdown_font=(MAIN_FONT, 12))
        self.device_power_dropdown_col2.pack(pady=5, fill="x", padx=10)

        self.update_device_power_options_col(1)
        self.update_device_power_options_col(2)
        
        video_path = os.path.join(acs_auto.image_folder, "Working.mp4")
        self.player_manager = VideoPlayer(tab, acs_auto.image_folder)
        self.player_manager.start_idle()

    def update_device_power_options_col(self, col):
        if col == 1:
            selected = self.device_type_var_col1.get()
            if selected == "AFVarionaut Pump": options = ["60", "100", "140", "160"]
            elif selected == "Submersible Pump": options = ["120", "150", "200"]
            elif selected == "Tricolor Led": options = ["18", "36"]
            elif selected == "SingleColor Led": options = ["6", "12"]
            elif selected == "Dmx2Vfd Converter": options = ["Unspecified"]
            else: options = []
            self.device_power_options_col1 = options
            self.device_power_dropdown_col1.configure(values=options)
            if self.device_power_var_col1.get() not in options and options:
                self.device_power_var_col1.set(options[0])
        else:
            selected = self.device_type_var_col2.get()
            if selected == "AFVarionaut Pump": options = ["60", "100", "140", "160"]
            elif selected == "Submersible Pump": options = ["120", "150", "200"]
            elif selected == "Tricolor Led": options = ["18", "36"]
            elif selected == "SingleColor Led": options = ["6", "12"]
            elif selected == "Dmx2Vfd Converter": options = ["Unspecified"]
            else: options = []
            self.device_power_options_col2 = options
            self.device_power_dropdown_col2.configure(values=options)
            if self.device_power_var_col2.get() not in options and options:
                self.device_power_var_col2.set(options[0])

    def create_acs_device_configuration_tab(self):
        tab = self.tabview.tab("ACS Device Configuration")

        def create_config_btn_row(parent, btn_text, script_cat, context_func):
            row_frame = ctk.CTkFrame(parent, fg_color="transparent")
            row_frame.pack(pady=5, fill="x", padx=10)
            
            btn_run = ctk.CTkButton(row_frame, text=btn_text, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14, "bold"),
                command=lambda: self.execute_category_script(script_cat, context_func))
            btn_run.pack(side="left", fill="x", expand=True)

            btn_script = ctk.CTkButton(row_frame, text="", image=icons.file, compound="left", width=40, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14),
                command=lambda: ScriptSelector(self, script_cat, lambda: None))
            btn_script.pack(side="left", padx=(5, 0))
            return btn_run

        self.btn_ghi_dia_chi = create_config_btn_row(tab, "Ghi địa chỉ (F3)", "address", self.get_excel_context)
        self.btn_test = create_config_btn_row(tab, "Test (F4)", "test", self.get_excel_context)
        self.btn_ghi_dia_chi_test = create_config_btn_row(tab, "Ghi địa chỉ & Test (F5)", "address_test", self.get_excel_context)

        ctk.CTkProgressBar(tab, height=2, progress_color=ACCENT_COLOR).pack(fill="x", pady=10, padx=20)

        excel_frame = ctk.CTkFrame(tab)
        excel_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(excel_frame, text="Danh sách địa chỉ", font=(MAIN_FONT, 12, "bold")).pack(anchor="w", padx=10, pady=(5,0))

        self.btn_import_excel = ctk.CTkButton(excel_frame, text="Nhập File Excel", image=icons.excel, compound="left", command=self.import_excel_gui, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14, "bold"))
        self.btn_import_excel.pack(pady=5, fill="x", padx=10)

        self.excel_status_var = ctk.StringVar()
        self.excel_status_label = ctk.CTkLabel(excel_frame, textvariable=self.excel_status_var, wraplength=300, font=(MAIN_FONT, 12))
        self.excel_status_label.pack(pady=5, fill="x", padx=10)

        manual_input_frame = ctk.CTkFrame(excel_frame, fg_color="transparent")
        manual_input_frame.pack(fill="x", padx=5)

        def create_entry(parent, label_text, trigger):
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(side="left", padx=2)
            ctk.CTkLabel(f, text=label_text, width=40, font=(MAIN_FONT, 12)).pack(side="left")
            entry = ctk.CTkEntry(f, width=50, font=(MAIN_FONT, 12))
            entry.pack(side="left")
            entry.bind("<KeyRelease>", lambda event: self.schedule_update(delay=500, trigger=trigger))
            return entry

        self.no_entry = create_entry(manual_input_frame, "No.", "no")
        self.pump_entry = create_entry(manual_input_frame, "Pump", "pump")
        self.led_entry = create_entry(manual_input_frame, "Led", "led")
        self.dmx2vfd_entry = create_entry(manual_input_frame, "Dmx2Vfd", "dmx2vfd")

        self.auto_inc_var = ctk.BooleanVar(value=True)
        
        def on_auto_inc_toggle():
            acs_auto.enable_auto_increment = self.auto_inc_var.get()

        chk_auto_inc = ctk.CTkCheckBox(excel_frame, text="Tự động xuống hàng", variable=self.auto_inc_var, command=on_auto_inc_toggle, fg_color=ACCENT_COLOR, hover_color=HOVER_COLOR, font=(MAIN_FONT, 12, "bold"))
        chk_auto_inc.pack(pady=(10, 10), anchor="w", padx=20)
        
        acs_auto.enable_auto_increment = True
        video_path = os.path.join(acs_auto.image_folder, "Working.mp4")
        self.player_config = VideoPlayer(tab, acs_auto.image_folder)
        self.player_config.start_idle()

    def create_settings_tab(self):
        tab = self.tabview.tab("Cài đặt")

        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=2)
        tab.rowconfigure(0, weight=1)

        left_frame = ctk.CTkFrame(tab)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        ctk.CTkLabel(left_frame, text="Danh sách Key", font=(MAIN_FONT, 14, "bold")).pack(anchor="w", pady=5, padx=5)

        listbox_container = ctk.CTkFrame(left_frame, fg_color="transparent")
        listbox_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.keys_listbox = tk.Listbox(listbox_container, bg=FRAME_BG, fg="white", font=(MAIN_FONT, 10), selectbackground=ACCENT_COLOR, bd=0, highlightthickness=0, exportselection=False)
        self.keys_listbox.pack(side="left", fill="both", expand=True)
        
        scrollbar = ctk.CTkScrollbar(listbox_container, command=self.keys_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.keys_listbox.config(yscrollcommand=scrollbar.set)
        self.keys_listbox.bind('<<ListboxSelect>>', self.on_key_selected)

        btn_key_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        btn_key_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkButton(btn_key_frame, text="", image=icons.add, compound="left", width=50, command=self.add_new_key, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 12)).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(btn_key_frame, text="", image=icons.edit, compound="left", width=50, command=self.rename_current_key, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 12)).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(btn_key_frame, text="", image=icons.delete, compound="left", width=50, command=self.delete_current_key, fg_color="#cc3333", hover_color="#aa2222", font=(MAIN_FONT, 12)).pack(side="left", padx=5, pady=5)

        right_frame = ctk.CTkFrame(tab)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        ctk.CTkLabel(right_frame, text="Chi tiết cấu hình", font=(MAIN_FONT, 14, "bold")).pack(anchor="w", pady=5, padx=5)

        gen_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        gen_frame.pack(fill="x", pady=(0, 10), padx=5)
        
        def add_gen_entry(parent, label, attr_name, default_val):
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(fill="x", pady=2)
            ctk.CTkLabel(f, text=label, width=150, anchor="w", font=(MAIN_FONT, 12)).pack(side="left")
            entry = ctk.CTkEntry(f, font=(MAIN_FONT, 12))
            entry.pack(side="left", fill="x", expand=True)
            entry.insert(0, default_val)
            setattr(self, attr_name, entry)

        add_gen_entry(gen_frame, "Delay chụp màn hình", "screenshot_delay_entry", config_manager.get('GENERAL', 'screenshot_delay_sec'))
        add_gen_entry(gen_frame, "Delay thao tác", "action_delay_entry", config_manager.get('GENERAL', 'action_delay_sec'))
        add_gen_entry(gen_frame, "Độ tin cậy (0.0 - 1.0)", "confidence_entry", config_manager.get('GENERAL', 'find_image_confidence'))

        ctk.CTkProgressBar(right_frame, height=2, progress_color=ACCENT_COLOR).pack(fill="x", pady=10, padx=10)

        self.lbl_current_key = ctk.CTkLabel(right_frame, text="Đang chọn: (Chưa chọn)", font=(MAIN_FONT, 12, "bold"))
        self.lbl_current_key.pack(anchor="w", pady=(0, 5), padx=5)

        img_listbox_container = ctk.CTkFrame(right_frame, fg_color="transparent")
        img_listbox_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.images_listbox = tk.Listbox(img_listbox_container, bg=FRAME_BG, fg="white", font=(MAIN_FONT, 10),
                                         selectbackground=ACCENT_COLOR, bd=0, highlightthickness=0, exportselection=False)
        self.images_listbox.pack(side="left", fill="both", expand=True)
        img_scroll = ctk.CTkScrollbar(img_listbox_container, command=self.images_listbox.yview)
        img_scroll.pack(side="right", fill="y")
        self.images_listbox.config(yscrollcommand=img_scroll.set)

        btn_img_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        btn_img_frame.pack(fill="x", pady=5, padx=5)

        ctk.CTkButton(btn_img_frame, text="Thêm ảnh", image=icons.photo, compound="left", command=self.add_image_to_key, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 14, "bold")).pack(fill="x", expand=True, side="left", padx=5)
        ctk.CTkButton(btn_img_frame, text="Xóa ảnh", image=icons.delete, compound="left", command=self.remove_image_from_key, fg_color="#cc3333", hover_color="#aa2222", font=(MAIN_FONT, 14, "bold")).pack(fill="x", expand=True, side="left", padx=5)

        save_btn = ctk.CTkButton(right_frame, text="LƯU TOÀN BỘ CÀI ĐẶT", image=icons.save, compound="left", command=self.save_settings_dynamic, fg_color="#6a9955", hover_color="#88c46d", font=(MAIN_FONT, 14, "bold"))
        save_btn.pack(fill="x", pady=10, padx=10)

        self.refresh_keys_list()

    def refresh_keys_list(self):
        self.keys_listbox.delete(0, tk.END)
        if config_manager.config.has_section('IMAGE_PATHS'):
            keys = config_manager.config.options('IMAGE_PATHS')
            for key in keys:
                self.keys_listbox.insert(tk.END, key)

    def on_key_selected(self, event):
        selection = self.keys_listbox.curselection()
        if not selection:
            return
        key = self.keys_listbox.get(selection[0])
        self.lbl_current_key.configure(text=f"Đang chọn: {key}")
        
        self.images_listbox.delete(0, tk.END)
        raw_val = config_manager.get('IMAGE_PATHS', key, "")
        if raw_val:
            paths = [p.strip() for p in raw_val.split(',') if p.strip()]
        for p in paths:
            self.images_listbox.insert(tk.END, p)

    def _spawn_inline_entry(self, listbox, index, initial_text, on_commit):
        bbox = listbox.bbox(index)
        if not bbox: return 
        entry = tk.Entry(listbox, borderwidth=0, highlightthickness=1, 
                        bg=FRAME_BG, fg="white", font=(MAIN_FONT, 10))
        x, y, w, h = bbox
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, initial_text)
        entry.select_range(0, tk.END)
        entry.focus_set()

        def commit(event=None):
            new_val = entry.get().strip()
            entry.destroy()
            on_commit(new_val)

        def cancel(event=None):
            entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", lambda e: commit()) 
        entry.bind("<Escape>", cancel)

    def add_new_key(self):  
        temp_name = "new_key"
        self.keys_listbox.insert(tk.END, temp_name)
        idx = self.keys_listbox.size() - 1
        self.keys_listbox.see(idx)
        self.keys_listbox.selection_clear(0, tk.END)
        self.keys_listbox.selection_set(idx)

        def on_add_commit(new_key_name):
            self.keys_listbox.delete(idx)
            final_name = new_key_name.strip().replace(" ", "_").lower()
            if not final_name: return

            if config_manager.config.has_option('IMAGE_PATHS', final_name):
                messagebox.showwarning("Lỗi", "Key này đã tồn tại!")
                self.refresh_keys_list()
                return
            config_manager.set('IMAGE_PATHS', final_name, "")
            self.refresh_keys_list()
            try:
                new_idx = self.keys_listbox.get(0, tk.END).index(final_name)
                self.keys_listbox.selection_set(new_idx)
                self.keys_listbox.event_generate("<<ListboxSelect>>")
            except: pass

        self._spawn_inline_entry(self.keys_listbox, idx, "", on_add_commit)

    def rename_current_key(self):
        selection = self.keys_listbox.curselection()
        if not selection: return
        idx = selection[0]
        old_key = self.keys_listbox.get(idx)

        def on_rename_commit(new_key_name):
            final_name = new_key_name.strip().replace(" ", "_").lower()
            if not final_name or final_name == old_key: return

            if config_manager.config.has_option('IMAGE_PATHS', final_name):
                messagebox.showwarning("Lỗi", "Key này đã tồn tại!")
                return

            val = config_manager.get('IMAGE_PATHS', old_key)
            config_manager.config.remove_option('IMAGE_PATHS', old_key)
            config_manager.set('IMAGE_PATHS', final_name, val)
            
            self.refresh_keys_list()
            try:
                new_idx = self.keys_listbox.get(0, tk.END).index(final_name)
                self.keys_listbox.selection_set(new_idx)
                self.keys_listbox.event_generate("<<ListboxSelect>>")
            except: pass

        self._spawn_inline_entry(self.keys_listbox, idx, old_key, on_rename_commit)

    def delete_current_key(self):
        selection = self.keys_listbox.curselection()
        if not selection: return
        key = self.keys_listbox.get(selection[0])
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa Key '{key}' không?"):
            config_manager.config.remove_option('IMAGE_PATHS', key)
            config_manager.save_config() 
            self.refresh_keys_list()
            self.images_listbox.delete(0, tk.END)
            self.lbl_current_key.configure(text="Đang chọn: (Chưa chọn)")

    def add_image_to_key(self):
        selection = self.keys_listbox.curselection()
        if not selection:
            messagebox.showwarning("Chú ý", "Vui lòng chọn một Key bên trái trước.")
            return
        
        key = self.keys_listbox.get(selection[0])
        image_folder_name = config_manager.get('GENERAL', 'image_folder')
        initial_dir = os.path.join(os.path.dirname(__file__), image_folder_name)
        if not os.path.exists(initial_dir): os.makedirs(initial_dir)

        file_paths = filedialog.askopenfilenames(
            initialdir=initial_dir,
            title=f"Chọn ảnh cho {key}",
            filetypes=(("PNG files", "*.png"), ("All files", "*.*"))
        )
        
        if file_paths:
            current_images = self.images_listbox.get(0, tk.END)
            for file_path in file_paths:
                filename = os.path.basename(file_path)
                if filename not in current_images:
                    self.images_listbox.insert(tk.END, filename)
            self._update_config_images_from_listbox(key)

    def remove_image_from_key(self):
        selection = self.keys_listbox.curselection()
        if not selection: 
            messagebox.showwarning("Chú ý", "Chưa chọn Key nào.")
            return
        key = self.keys_listbox.get(selection[0])
        img_sel = self.images_listbox.curselection()
        if not img_sel: 
            messagebox.showwarning("Chú ý", "Chưa chọn ảnh để xóa.")
            return
        for index in reversed(img_sel):
            self.images_listbox.delete(index)
        self._update_config_images_from_listbox(key)

    def _update_config_images_from_listbox(self, key):
        images = self.images_listbox.get(0, tk.END)
        val_str = ",".join(images)
        config_manager.set('IMAGE_PATHS', key, val_str)

    def save_settings_dynamic(self):
        try:
            config_manager.set('GENERAL', 'screenshot_delay_sec', self.screenshot_delay_entry.get())
            config_manager.set('GENERAL', 'action_delay_sec', self.action_delay_entry.get())
            config_manager.set('GENERAL', 'find_image_confidence', self.confidence_entry.get())
            config_manager.save_config()
            logger.info("Người dùng đã lưu cài đặt mới.")
        except Exception as e:
            logger.error(f"Lỗi khi lưu cài đặt: {e}", exc_info=True)

    def create_suki_tab(self):
        tab = self.tabview.tab("Suki UwU")
        
        content_frame = ctk.CTkFrame(tab, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        ctk.CTkLabel(content_frame, text=APP_NAME, font=(MAIN_FONT, 24, "bold")).pack(pady=(40, 10))
        ctk.CTkLabel(content_frame, text="Tự động Ghi UID, địa chỉ và test", font=(MAIN_FONT, 12, "bold")).pack(pady=5)
        ctk.CTkLabel(content_frame, text=f"Version: {VERSION}", font=(MAIN_FONT, 12, "bold")).pack(pady=5)
        
        author_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        author_frame.pack(pady=5)
        ctk.CTkLabel(author_frame, text="Author: ", font=(MAIN_FONT, 12, "bold")).pack(side="left")
        author_link = ctk.CTkLabel(author_frame, text="Suki", text_color=ACCENT_COLOR, cursor="hand2", font=(MAIN_FONT, 12, "bold"))
        author_link.pack(side="left")
        author_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/Suki8898"))
        
        try:
            image_path = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'image_folder'), 'Suki.png')
            pil_image = Image.open(image_path)
            
            max_width, max_height = 250, 250
            original_width, original_height = pil_image.size
            new_width, new_height = original_width, original_height
            
            if original_width > max_width or original_height > max_height:
                ratio = min(max_width / original_width, max_height / original_height)
                new_width, new_height = int(original_width * ratio), int(original_height * ratio)
                pil_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            radius = 30
            sampling_factor = 4
            mask_w = new_width * sampling_factor
            mask_h = new_height * sampling_factor
            mask = Image.new("L", (mask_w, mask_h), 0)
            draw = ImageDraw.Draw(mask)   
            draw.rounded_rectangle((0, 0, mask_w, mask_h), radius=radius * sampling_factor, fill=255)
            mask = mask.resize((new_width, new_height), Image.Resampling.LANCZOS)
            pil_image = pil_image.convert("RGBA")
            pil_image.putalpha(mask)

            self.suki_image_ref = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(new_width, new_height))
            ctk.CTkLabel(content_frame, image=self.suki_image_ref, text="").pack(pady=(20, 10))
        except Exception as e:
            logger.error(f"Lỗi xử lý ảnh Suki: {e}")
            pass

    def stop_all_automation(self):
        if acs_auto.stop_requested == False:
            acs_auto.stop_requested = True
            logger.info("🛑 Đã dừng.")
        else:
            acs_auto.stop_requested = False
            logger.info("▶ Hủy dừng.")

    def schedule_update(self, delay, trigger):
        if hasattr(self, 'after_id') and self.after_id:
            self.after_cancel(self.after_id)
        self.after_id = self.after(delay, lambda: self.delayed_update(trigger))

    def delayed_update(self, trigger):
        try:
            if trigger == "no":
                row_number = int(self.no_entry.get()) - 1 if self.no_entry.get() else None
            elif trigger == "pump":
                pump_value = int(self.pump_entry.get()) if self.pump_entry.get() else None
                row_number = next((i for i, row in enumerate(acs_auto.excel_data) if row['Pump'] == pump_value), None) if pump_value else None
            elif trigger == "led":
                led_value = int(self.led_entry.get()) if self.led_entry.get() else None
                row_number = next((i for i, row in enumerate(acs_auto.excel_data) if row['Led'] == led_value), None) if led_value else None
            elif trigger == "dmx2vfd":
                dmx2vfd_value = int(self.dmx2vfd_entry.get()) if self.dmx2vfd_entry.get() else None
                row_number = next((i for i, row in enumerate(acs_auto.excel_data) if row['Dmx2Vfd'] == dmx2vfd_value), None) if dmx2vfd_value else None
            else:
                return
            
            if row_number is not None and acs_auto.excel_data and 0 <= row_number < len(acs_auto.excel_data):
                acs_auto.current_excel_row_index = row_number
                self.update_excel_status()
                self.update_entry_fields(row_number)
            elif row_number is not None and not acs_auto.excel_data:
                logger.warning("Không có dữ liệu Excel để tìm kiếm.")
        except ValueError:
            pass
        except Exception as e:
            logger.error(f"Lỗi delayed_update: {e}", exc_info=True)
    
    def update_entry_fields(self, row_number):
        self.no_entry.delete(0, tk.END)
        self.pump_entry.delete(0, tk.END)
        self.led_entry.delete(0, tk.END)
        self.dmx2vfd_entry.delete(0, tk.END)  
        try:
            if acs_auto.excel_data and 0 <= row_number < len(acs_auto.excel_data):
                row = acs_auto.excel_data[row_number]
                self.no_entry.insert(0, str(row_number + 1))
                self.pump_entry.insert(0, str(row['Pump']))
                self.led_entry.insert(0, str(row['Led']))
                self.dmx2vfd_entry.insert(0, str(row['Dmx2Vfd'])) 
        except Exception as e:
            logger.warning(f"Không thể lấy dữ liệu hàng {row_number}: {e}")

    def update_excel_status(self):
        if acs_auto.excel_data:
            total_rows = len(acs_auto.excel_data)
            current_row = acs_auto.current_excel_row_index
            self.excel_status_var.set(f"Đang ghi hàng: {current_row}/{total_rows}")
        else:
            self.excel_status_var.set("Chưa có file Excel được nhập.")

    def import_excel_gui(self):
        file_path = filedialog.askopenfilename(
            title="Chọn File Excel (.xlsx)",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            if acs_auto.import_excel_data(file_path):
                logger.info("Nhập Excel thành công!")
                self.update_excel_status()
                self.update_entry_fields(0)
            else:
                logger.error("Không thể nhập file Excel. Vui lòng kiểm tra log.")
                self.update_excel_status()

    def set_buttons_state(self, state):
        self.btn_uid_col1.configure(state=state)
        self.btn_uid_col2.configure(state=state)
        self.btn_ghi_dia_chi.configure(state=state)
        self.btn_test.configure(state=state)
        self.btn_ghi_dia_chi_test.configure(state=state)
        self.btn_import_excel.configure(state=state)

    def create_custom_title_bar(self):
        self.title_bar = ctk.CTkFrame(self, fg_color=FRAME_BG, corner_radius=0, height=30)
        self.title_bar.pack(side="top", fill="x")

        icon_path = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'), 'app.ico')
        if os.path.exists(icon_path):
            try:
                pil_icon = Image.open(icon_path)
                self.title_icon_img = ctk.CTkImage(light_image=pil_icon, dark_image=pil_icon, size=(24, 24))
                ctk.CTkLabel(self.title_bar, image=self.title_icon_img, text="").pack(side="left", padx=5, pady=2)
            except: pass
        
        def animate_title_rainbow(parent, text="A C S  A u t o", font=(MAIN_FONT, 11, "bold"), bg="#2b2b2b", speed=50):
            canvas = tk.Canvas(parent, bg=bg, highlightthickness=0, height=28)
            canvas.pack(side="left", fill="x", expand=True)
            hue = 0.0
            def draw_text():
                nonlocal hue
                canvas.delete("all")
                for i, ch in enumerate(reversed(text)):
                    offset = (hue + i * 0.01) % 1.0
                    r, g, b = [int(255 * v) for v in colorsys.hsv_to_rgb(offset, 1, 1)]
                    canvas.create_text(175 + (len(text) - i - 1) * 5, 15, text=ch, fill=f'#{r:02x}{g:02x}{b:02x}', font=font, anchor="w")
                hue = (hue + 0.02) % 1.0
                canvas.after(speed, draw_text)
            draw_text()
            return canvas

        canvas = animate_title_rainbow(self.title_bar, bg=FRAME_BG)

        self.close_button = ctk.CTkButton(self.title_bar, text="❌", command=self.destroy,
            fg_color="transparent", hover_color="#cc3333", width=30, height=30, font=(MAIN_FONT, 12))
        self.close_button.pack(side="right")
        self.min_button = ctk.CTkButton(self.title_bar, text="—", command=self.minimize_window,
            fg_color="transparent", hover_color="#444", width=30, height=30, font=(MAIN_FONT, 12))
        self.min_button.pack(side="right")
        
        self.title_bar.bind("<ButtonPress-1>", self._start_move_window)
        self.title_bar.bind("<B1-Motion>", self._move_window)
        canvas.bind("<ButtonPress-1>", self._start_move_window)
        canvas.bind("<B1-Motion>", self._move_window)

    def minimize_window(self):
        self.withdraw()
        self.show_mini_bar()

    def show_mini_bar(self):
        if hasattr(self, "mini_bar") and self.mini_bar.winfo_exists():
            self.mini_bar.lift()
            return

        self.mini_bar = ctk.CTkToplevel(self)
        self.mini_bar.overrideredirect(True)
        self.mini_bar.attributes("-topmost", True)
        self.mini_bar.configure(fg_color="#2b2b2b")

        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty()
        width, height = 140, 40

        self.mini_bar.geometry(f"{width}x{height}+{x}+{y}")
        self.mini_bar.attributes("-alpha", 0.92)

        def create_rainbow_canvas(parent, text="A C S  A u t o  :3", font=(MAIN_FONT, 11, "bold"), bg="#2b2b2b", speed=50):
            canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
            canvas.pack(expand=True, fill="both")
            hue = 0.0
            def draw_text():
                nonlocal hue
                canvas.delete("all")
                for i, ch in enumerate(reversed(text)):
                    offset = (hue + i * 0.01) % 1.0
                    r, g, b = [int(255 * v) for v in colorsys.hsv_to_rgb(offset, 1, 1)]
                    color = f'#{r:02x}{g:02x}{b:02x}'
                    canvas.create_text(25 + (len(text) - i - 1) * 5, 20, text=ch, fill=color, font=font, anchor="w")
                hue = (hue + 0.02) % 1.0
                canvas.after(speed, draw_text)
            draw_text()
            return canvas

        canvas = create_rainbow_canvas(self.mini_bar, text="A C S  A u t o  :3")
        canvas.bind("<Button-1>", self.start_move)
        canvas.bind("<B1-Motion>", self.do_move)
        canvas.bind("<Double-Button-1>", lambda e: self.restore_main_window())
        canvas.bind("<Enter>", lambda e: canvas.configure(bg="#3c3c3c"))
        canvas.bind("<Leave>", lambda e: canvas.configure(bg="#2b2b2b"))

    def start_move(self, event):
        self._x, self._y = event.x, event.y

    def do_move(self, event):
        x, y = event.x_root - self._x, event.y_root - self._y
        self.mini_bar.geometry(f"+{x}+{y}")

    def restore_main_window(self):
        if hasattr(self, "mini_bar") and self.mini_bar.winfo_exists():
            x, y = self.mini_bar.winfo_x(), self.mini_bar.winfo_y()
            self.mini_bar.destroy()
            self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)

    def _start_move_window(self, event):
        self.start_x, self.start_y = event.x, event.y

    def _move_window(self, event):
        new_x = self.winfo_x() + (event.x - self.start_x)
        new_y = self.winfo_y() + (event.y - self.start_y)
        self.geometry(f"+{new_x}+{new_y}")

class ScriptManager:
    def __init__(self, filepath='scripts.json'):
        self.filepath = os.path.join(os.path.dirname(__file__), filepath)
        self.scripts = self.load_scripts()

    def load_scripts(self):
        if not os.path.exists(self.filepath):
            return self.create_default_scripts()
        try:
            with open(self.filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                for cat in data:
                    for i, script in enumerate(data[cat]):
                        if 'active' not in script:
                            script['active'] = (i == 0)
                return data
        except Exception as e:
            logger.error(f"Lỗi đọc script json: {e}")
            return self.create_default_scripts()

    def save_scripts(self):
        try:
            with open(self.filepath, 'w', encoding='utf-8') as f:
                json.dump(self.scripts, f, indent=4, ensure_ascii=False)
            logger.info("Đã lưu kịch bản.")
        except Exception as e:
            logger.error(f"Lỗi lưu script json: {e}")

    def get_scripts_by_category(self, category):
        return self.scripts.get(category, [])

    def get_active_script(self, category):
        scripts = self.get_scripts_by_category(category)
        for script in scripts:
            if script.get('active', False):
                return script
        if scripts:
            return scripts[0]
        return None

    def create_default_scripts(self):
        default_steps = [{"name": "Ví dụ", "code": "pass"}]
        data = {
            "uid_col1": [{"name": "Mặc định Col 1", "active": True, "steps": default_steps}],
            "uid_col2": [{"name": "Mặc định Col 2", "active": True, "steps": default_steps}],
            "address": [{"name": "Ghi địa chỉ chuẩn", "active": True, "steps": default_steps}],
            "test": [{"name": "Test chuẩn", "active": True, "steps": default_steps}],
            "address_test": [{"name": "Ghi & Test", "active": True, "steps": default_steps}]
        }
        return data

script_manager = ScriptManager()

def setup_custom_window(window, title_text, is_resizable=False, width=None, height=None):
    window.configure(fg_color=DARK_BG)
    window.overrideredirect(True)  
    window.attributes("-topmost", True) 
   
    if width and height:
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = int((screen_width - width) / 2)
        y = int((screen_height - height) / 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    main_border = ctk.CTkFrame(window, fg_color="#505050")
    main_border.pack(fill="both", expand=True, padx=1, pady=1)
    
    inner_content = ctk.CTkFrame(main_border, fg_color=DARK_BG)
    inner_content.pack(fill="both", expand=True)

    title_bar = ctk.CTkFrame(inner_content, fg_color=FRAME_BG, height=30, corner_radius=0)
    title_bar.pack(side="top", fill="x")
    
    try:
        icon_path = os.path.join(os.path.dirname(__file__), config_manager.get('GENERAL', 'icon_folder'), 'app.ico')
        if os.path.exists(icon_path):
            pil_icon = Image.open(icon_path).resize((18, 18), Image.Resampling.LANCZOS)
            window.icon_img_ref = ctk.CTkImage(light_image=pil_icon, dark_image=pil_icon, size=(18, 18))
            ctk.CTkLabel(title_bar, image=window.icon_img_ref, text="").pack(side="left", padx=(5, 5))
    except: pass

    title_lbl = ctk.CTkLabel(title_bar, text=title_text, text_color="white", font=(MAIN_FONT, 10, "bold"))
    title_lbl.pack(side="left", padx=5)

    close_btn = ctk.CTkButton(title_bar, text="✕", command=window.destroy,
                          fg_color="transparent", hover_color="#cc3333", width=30, height=25, font=(MAIN_FONT, 10))
    close_btn.pack(side="right", padx=2)

    def start_move(event):
        window.x = event.x
        window.y = event.y

    def do_move(event):
        deltax = event.x - window.x
        deltay = event.y - window.y
        x = window.winfo_x() + deltax
        y = window.winfo_y() + deltay
        window.geometry(f"+{x}+{y}")

    title_bar.bind("<ButtonPress-1>", start_move)
    title_bar.bind("<B1-Motion>", do_move)
    title_lbl.bind("<ButtonPress-1>", start_move)
    title_lbl.bind("<B1-Motion>", do_move)

    if is_resizable:
        grip_size = 15
        grip_frame = ctk.CTkFrame(inner_content, fg_color=DARK_BG, height=grip_size, cursor="sizing")
        grip_frame.pack(side="bottom", fill="x") 
        grip = ctk.CTkLabel(grip_frame, text="◢", text_color="#505050", cursor="bottom_right_corner", font=(MAIN_FONT, 10))
        grip.pack(side="right", anchor="se")

        def start_resize(event):
            window.start_x = event.x_root
            window.start_y = event.y_root
            window.start_width = window.winfo_width()
            window.start_height = window.winfo_height()

        def do_resize(event):
            dx = event.x_root - window.start_x
            dy = event.y_root - window.start_y
            new_width = max(window.start_width + dx, 400) 
            new_height = max(window.start_height + dy, 300) 
            window.geometry(f"{new_width}x{new_height}")

        grip.bind("<ButtonPress-1>", start_resize)
        grip.bind("<B1-Motion>", do_resize)
        
        return inner_content, grip_frame 
    
    return inner_content, None

class ScriptSelector(ctk.CTkToplevel):
    def __init__(self, master, category, on_script_selected):
        super().__init__(master)
        
        self.content_frame, _ = setup_custom_window(self, f"Quản lý Kịch bản: {category}", is_resizable=False, width=550, height=450)
        
        self.category = category
        self.on_script_selected = on_script_selected
        self.editing_index = -1 
        
        canvas_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(canvas_frame, bg=FRAME_BG, highlightthickness=0)
        self.scrollbar = ctk.CTkScrollbar(canvas_frame, orientation="vertical", command=self.canvas.yview)
        
        self.list_frame = ctk.CTkFrame(self.canvas, fg_color=FRAME_BG)
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.list_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))

        btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10, padx=10)
        
        btn_new = ctk.CTkButton(btn_frame, text="Tạo kịch bản mới", image=icons.add, compound="left", command=self.create_new, fg_color=BUTTON_BG, hover_color=HOVER_COLOR, font=(MAIN_FONT, 12, "bold"))
        btn_new.pack(side="left")

        btn_close = ctk.CTkButton(btn_frame, text="Đóng", command=self.destroy, fg_color="#cc3333", hover_color="#aa2222", font=(MAIN_FONT, 12, "bold"))
        btn_close.pack(side="right")

        self.refresh_list()

    def refresh_list(self):
        for w in self.list_frame.winfo_children(): w.destroy()
        
        self.scripts_list = script_manager.get_scripts_by_category(self.category)
        
        header = ctk.CTkFrame(self.list_frame, fg_color="#333333")
        header.pack(fill="x", pady=0)
        ctk.CTkLabel(header, text="Active", width=50, font=(MAIN_FONT, 12, "bold")).pack(side="left")
        ctk.CTkLabel(header, text="Tên Kịch Bản", font=(MAIN_FONT, 12, "bold")).pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkLabel(header, text="Thao tác", width=100, font=(MAIN_FONT, 12, "bold")).pack(side="right", padx=5)

        if not self.scripts_list:
            ctk.CTkLabel(self.list_frame, text="(Chưa có kịch bản nào)", text_color="#888", font=(MAIN_FONT, 12)).pack(pady=20)

        for i, script in enumerate(self.scripts_list):
            bg_color = FRAME_BG
            row_frame = ctk.CTkFrame(self.list_frame, fg_color=bg_color)
            row_frame.pack(fill="x", pady=1)
            
            is_active = script.get('active', False)
            var = ctk.BooleanVar(value=is_active)
            chk = ctk.CTkCheckBox(row_frame, text="", variable=var, command=lambda idx=i: self.set_active_script(idx), width=20, fg_color=ACCENT_COLOR, font=(MAIN_FONT, 12))
            chk.pack(side="left", padx=15)

            name_container = ctk.CTkFrame(row_frame, fg_color="transparent")
            name_container.pack(side="left", padx=5, fill="x", expand=True)

            if i == self.editing_index:
                entry_name = ctk.CTkEntry(name_container, height=28, font=(MAIN_FONT, 12))
                entry_name.pack(fill="x")
                entry_name.insert(0, script['name'])
                entry_name.focus_set() 
                entry_name.select_range(0, tk.END) 
                
                entry_name.bind('<Return>', lambda e, idx=i, ent=entry_name: self.save_name(idx, ent.get()))
                entry_name.bind('<Escape>', lambda e: self.cancel_rename())
                
            else:
                fg_color = ACCENT_COLOR if is_active else "white"
                name_lbl = ctk.CTkLabel(name_container, text=script['name'], text_color=fg_color, 
                                    font=(MAIN_FONT, 12, 'bold' if is_active else 'normal'), anchor="w")
                name_lbl.pack(fill="x", expand=True)
                name_lbl.bind("<Double-Button-1>", lambda e, idx=i: self.start_rename(idx))

            btn_box = ctk.CTkFrame(row_frame, fg_color="transparent")
            btn_box.pack(side="right", padx=5)

            if i == self.editing_index:
                ctk.CTkButton(btn_box, text="", image=icons.check, compound="left", fg_color="#6a9955", hover_color="#88c46d", width=30, height=25,
                          command=lambda idx=i, ent=entry_name: self.save_name(idx, ent.get())).pack(side="left", padx=1)
                ctk.CTkButton(btn_box, text="", image=icons.x, compound="left", fg_color="#cc3333", hover_color="#e66666", width=30, height=25,
                          command=self.cancel_rename).pack(side="left", padx=1)
            else:
                ctk.CTkButton(btn_box, text="", image=icons.edit, compound="left", fg_color=BUTTON_BG, hover_color=HOVER_COLOR, width=30, height=25,
                          command=lambda idx=i: self.start_rename(idx)).pack(side="left", padx=1)

                ctk.CTkButton(btn_box, text="", image=icons.code, compound="left", fg_color=BUTTON_BG, hover_color=HOVER_COLOR, width=30, height=25,
                          command=lambda idx=i: self.edit_content(idx)).pack(side="left", padx=1)
                
                if len(self.scripts_list) > 1:
                    ctk.CTkButton(btn_box, text="", image=icons.delete, compound="left", fg_color="#cc3333", hover_color="#e66666", width=30, height=25,
                              command=lambda idx=i: self.delete_script(idx)).pack(side="left", padx=1)

    def create_new(self):
        base_name = "Kịch bản mới"
        count = 1
        existing_names = [s['name'] for s in self.scripts_list]
        new_name = f"{base_name} {count}"
        while new_name in existing_names:
            count += 1
            new_name = f"{base_name} {count}"

        new_script = {"name": new_name, "active": False, "steps": [{"name": "New Block", "code": "pass"}]}
        
        if self.category not in script_manager.scripts: 
            script_manager.scripts[self.category] = []
        
        if not script_manager.scripts[self.category]: 
            new_script['active'] = True
            
        script_manager.scripts[self.category].append(new_script)
        script_manager.save_scripts()
        
        self.editing_index = len(script_manager.scripts[self.category]) - 1
        self.refresh_list()

    def start_rename(self, index):
        self.editing_index = index
        self.refresh_list()

    def cancel_rename(self):
        self.editing_index = -1
        self.refresh_list()

    def save_name(self, index, new_name):
        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("Lỗi", "Tên không được để trống!", parent=self)
            return

        self.scripts_list[index]['name'] = new_name
        script_manager.save_scripts()
        self.editing_index = -1
        self.refresh_list()

    def set_active_script(self, index):
        for s in self.scripts_list: s['active'] = False
        self.scripts_list[index]['active'] = True
        script_manager.save_scripts()
        self.refresh_list()

    def edit_content(self, index):
        ScriptEditor(self.master, self.category, index, self.refresh_list)

    def delete_script(self, index):
        if messagebox.askyesno("Xác nhận", "Xóa kịch bản này?", parent=self):
            del script_manager.scripts[self.category][index]
            if not any(s.get('active') for s in script_manager.scripts[self.category]) and script_manager.scripts[self.category]:
                script_manager.scripts[self.category][0]['active'] = True
            script_manager.save_scripts()
            if self.editing_index == index:
                self.editing_index = -1
            elif self.editing_index > index:
                self.editing_index -= 1
            self.refresh_list()

class FindReplaceDialog(ctk.CTkToplevel):
    def __init__(self, master, text_widget):
        super().__init__(master)
        
        self.content_frame, _ = setup_custom_window(
            self, 
            title_text="Tìm kiếm & Thay thế", 
            is_resizable=False, 
            width=400, 
            height=220
        )
        
        self.text_widget = text_widget
        self.create_interface()
        self.bind('<Return>', lambda e: self.find_next())
        self.bind('<Escape>', lambda e: self.destroy())
        self.find_entry.bind("<KeyRelease>", self.live_highlight) 
        self.find_entry.focus_set()

    def live_highlight(self, event=None):
        query = self.find_entry.get()
        self.text_widget.tag_remove('match', '1.0', tk.END)
        self.text_widget.tag_remove('current_match', '1.0', tk.END)
        if not query: return
        start_pos = '1.0'
        while True:
            pos = self.text_widget.search(query, start_pos, stopindex=tk.END, nocase=True)
            if not pos:
                break
            end_pos = f"{pos}+{len(query)}c"
            self.text_widget.tag_add('match', pos, end_pos)
            start_pos = end_pos

    def create_interface(self):
        main_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(main_frame, text="Từ khóa:", font=(MAIN_FONT, 12)).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.find_entry = ctk.CTkEntry(main_frame, font=(MAIN_FONT, 12))
        self.find_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=(0, 5))

        ctk.CTkLabel(main_frame, text="Thay bằng:", font=(MAIN_FONT, 12)).grid(row=1, column=0, sticky="w", pady=5)
        self.replace_entry = ctk.CTkEntry(main_frame, font=(MAIN_FONT, 12))
        self.replace_entry.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)

        main_frame.grid_columnconfigure(1, weight=1)

        btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        btn_frame.pack(fill="x", side="bottom", padx=10, pady=10)

        def create_btn(parent, text, cmd):
            btn = ctk.CTkButton(parent, text=text, command=cmd, height=30, fg_color="#333", hover_color="#444", font=(MAIN_FONT, 12))
            btn.pack(side="right", padx=5)
            return btn

        create_btn(btn_frame, "Thay tất cả", self.replace_all)
        create_btn(btn_frame, "Thay thế", self.replace_one)
        create_btn(btn_frame, "Tìm tiếp", self.find_next)

    def find_next(self):
        query = self.find_entry.get()
        if not query: return
        try:
            start_idx = self.text_widget.index("sel.last")
        except tk.TclError:
            start_idx = self.text_widget.index(tk.INSERT)

        pos = self.text_widget.search(query, start_idx, stopindex=tk.END, nocase=True)
        if not pos:
            pos = self.text_widget.search(query, '1.0', stopindex=tk.END, nocase=True)
            
        if pos:
            end_pos = f"{pos}+{len(query)}c"
            self.text_widget.see(pos)
            self.text_widget.tag_remove("current_match", "1.0", tk.END)
            self.text_widget.tag_add("current_match", pos, end_pos)
            self.text_widget.tag_remove(tk.SEL, "1.0", tk.END)
            self.text_widget.tag_add(tk.SEL, pos, end_pos)
            self.text_widget.mark_set(tk.INSERT, end_pos)
            self.find_entry.focus_set()
            return pos, end_pos
        else:
            messagebox.showinfo("Thông báo", f"Không tìm thấy '{query}'", parent=self)
            self.find_entry.focus_set()
            return None, None

    def replace_one(self):
        query = self.find_entry.get()
        replacement = self.replace_entry.get()
        if not query: return
        try:
            sel_first = self.text_widget.index("sel.first")
            sel_last = self.text_widget.index("sel.last")
            current_selection = self.text_widget.get(sel_first, sel_last)
            if current_selection.lower() == query.lower():
                self.text_widget.delete(sel_first, sel_last)
                self.text_widget.insert(sel_first, replacement)
                self.live_highlight()
                self.find_next()
            return
        except tk.TclError:
            pass 
        self.find_next()

    def replace_all(self):
        query = self.find_entry.get()
        replacement = self.replace_entry.get()
        if not query: return
        
        count = 0
        start_idx = '1.0'
        self.text_widget.config(state=tk.NORMAL)
        while True:
            pos = self.text_widget.search(query, start_idx, stopindex=tk.END, nocase=True)
            if not pos: break
            end_pos = f"{pos}+{len(query)}c"
            self.text_widget.delete(pos, end_pos)
            self.text_widget.insert(pos, replacement)
            start_idx = f"{pos}+{len(replacement)}c"
            count += 1
        self.live_highlight()
        messagebox.showinfo("Hoàn tất", f"Đã thay thế {count} mục.", parent=self)
        self.find_entry.focus_set()          

class ScriptEditor(ctk.CTkToplevel):
    def __init__(self, master, category, script_index, on_save_callback):
        super().__init__(master)
        
        self.content_frame, self.status_bar = setup_custom_window(self, f"Chỉnh sửa Kịch bản: {category}", is_resizable=True, width=1000, height=700)
        
        self.category = category
        self.script_index = script_index
        self.on_save_callback = on_save_callback
        
        self.script_data = copy.deepcopy(script_manager.scripts[category][script_index])
        self.steps = self.script_data['steps']
        self.current_step_index = -1
        self.editing_block_index = -1 
        
        self.setup_ui(self.content_frame)
        self.refresh_blocks()
        
        self.bind('<Control-s>', lambda e: self.save_all())
        self.bind('<Control-f>', lambda e: self.show_find_replace())

    def setup_ui(self, parent):
        self.paned = tk.PanedWindow(parent, orient=tk.HORIZONTAL, bg=DARK_BG, sashwidth=4, sashrelief="flat")
        self.paned.pack(fill="both", expand=True, padx=5, pady=5)

        self.left_frame = ctk.CTkFrame(self.paned, width=250)
        self.paned.add(self.left_frame)

        self.toolbar = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.toolbar.pack(fill="x", pady=2)
        
        self.create_tool_btn(self.toolbar, "", command=self.add_block, image=icons.add)
        self.create_tool_btn(self.toolbar, "", command=self.rename_current_block, image=icons.edit)
        self.create_tool_btn(self.toolbar, "", command=self.delete_block, image=icons.delete)
        self.create_tool_btn(self.toolbar, "", command=lambda: self.move_block(-1), image=icons.up)
        self.create_tool_btn(self.toolbar, "", command=lambda: self.move_block(1), image=icons.down)

        self.canvas = tk.Canvas(self.left_frame, bg="#282828", highlightthickness=0)
        self.scrollbar = ctk.CTkScrollbar(self.left_frame, orientation="vertical", command=self.canvas.yview)
        self.block_container = ctk.CTkFrame(self.canvas)
        
        self.block_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.block_container, anchor="nw", width=230)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))

        self.right_frame = ctk.CTkFrame(self.paned, fg_color="transparent")
        self.paned.add(self.right_frame)

        header_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 5))
        self.lbl_step_name = ctk.CTkLabel(header_frame, text="Chọn khối để sửa", font=(MAIN_FONT, 12, "bold"))
        self.lbl_step_name.pack(side="left")
        ctk.CTkLabel(header_frame, text="Ctrl+S: Lưu | Ctrl+F: Tìm", text_color="#888", font=(MAIN_FONT, 12)).pack(side="right")

        self.code_frame = tk.Frame(self.right_frame, bg="#1e1e1e")
        self.code_frame.pack(fill="both", expand=True)

        self.code_scroll = ctk.CTkScrollbar(self.code_frame, orientation="vertical")
        self.code_scroll.pack(side="right", fill="y")

        self.h_code_scroll = ctk.CTkScrollbar(self.code_frame, orientation="horizontal")
        self.h_code_scroll.pack(side="bottom", fill="x")

        self.line_numbers = tk.Text(self.code_frame, width=4, padx=5, pady=5, bg="#252526", fg="#858585", 
                                    bd=0, font=(MAIN_FONT, 11), state="disabled", highlightthickness=0)
        self.line_numbers.pack(side="left", fill="y")

        self.code_text = tk.Text(self.code_frame, bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
                                 font=(MAIN_FONT, 11), undo=True, autoseparators=True, maxundo=-1,
                                 yscrollcommand=self.sync_scroll_code, 
                                 xscrollcommand=self.h_code_scroll.set, 
                                 wrap="none", 
                                 borderwidth=0)
        self.code_text.pack(side="left", fill="both", expand=True)

        self.code_scroll.configure(command=self.sync_scroll_bar)
        self.h_code_scroll.configure(command=self.code_text.xview) 

        self.code_text.bind("<KeyRelease>", self.on_code_change)
        self.code_text.bind("<Tab>", self.insert_spaces)
        self.code_text.bind("<Control-y>", self.redo)
        self.code_text.bind("<Control-A>", self.select_all)
        self.code_text.bind("<Control-a>", self.select_all)
        self.code_text.bind("<MouseWheel>", self.on_scroll)
        self.code_text.bind("<Shift-Tab>", self.unindent)
        self.code_text.bind("<<Selection>>", self.auto_highlight_selection)
        
        self.setup_tags()

        btn_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=5)
        
        ctk.CTkButton(btn_frame, text="Lưu", image=icons.save, compound="left", command=self.save_all, fg_color="#6a9955", hover_color="#88c46d", font=(MAIN_FONT, 12, "bold")).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="Hủy", command=self.destroy, fg_color="#333", hover_color="#444", font=(MAIN_FONT, 12, "bold")).pack(side="right", padx=5)

    def create_tool_btn(self, parent, text, command, image=None):
        btn = ctk.CTkButton(parent, text=text, command=command, image=image, fg_color="#3c3c3c", hover_color="#555",  width=30, height=30, font=(MAIN_FONT, 10))
        btn.pack(side="left", padx=1)

    def refresh_blocks(self):
        for w in self.block_container.winfo_children(): w.destroy()
        for i, step in enumerate(self.steps):
            base_bg = "#333333"
            if i == self.current_step_index:
                base_bg = ACCENT_COLOR 
            
            frame_bg = "#2b2b2b" if i == self.editing_block_index else base_bg
            f = ctk.CTkFrame(self.block_container, fg_color=frame_bg)
            f.pack(fill="x", pady=1, padx=2)

            if i == self.editing_block_index:
                entry = ctk.CTkEntry(f, height=28, font=(MAIN_FONT, 12))
                entry.pack(side="left", fill="x", expand=True, padx=2)
                entry.insert(0, step['name'])
                entry.focus_set()
                
                ctk.CTkButton(f, text="", image=icons.check, compound="left", command=lambda idx=i, ent=entry: self.save_block_name(idx, ent.get()),
                          fg_color="#6a9955", width=30, height=25, font=(MAIN_FONT, 10)).pack(side="right", padx=1)
                
                ctk.CTkButton(f, text="", image=icons.x, compound="left", command=self.cancel_block_rename,
                          fg_color="#cc3333", width=30, height=25, font=(MAIN_FONT, 10)).pack(side="right", padx=1)

                entry.bind("<Return>", lambda e, idx=i, ent=entry: self.save_block_name(idx, ent.get()))
                entry.bind("<Escape>", lambda e: self.cancel_block_rename())
            else:
                fg_color = "black" if i == self.current_step_index else "white"
                font_style = (MAIN_FONT, 10, "bold") if i == self.current_step_index else (MAIN_FONT, 10, "bold")
                
                lbl = ctk.CTkLabel(f, text=f"{i+1}. {step['name']}", text_color=fg_color, font=font_style, anchor="w")
                lbl.pack(fill="both", expand=True, padx=10, pady=5)
                
                lbl.bind("<Button-1>", lambda e, idx=i: self.select_block(idx))
                lbl.bind("<Double-Button-1>", lambda e, idx=i: self.start_block_rename(idx))

    def add_block(self):
        base_name = "Bước mới"
        count = 1
        existing_names = [s['name'] for s in self.steps]
        new_name = f"{base_name} {count}"
        while new_name in existing_names:
            count += 1
            new_name = f"{base_name} {count}"

        self.steps.append({"name": new_name, "code": "# Code python here\npass"})     
        new_index = len(self.steps) - 1
        self.select_block(new_index)
        self.start_block_rename(new_index)

    def rename_current_block(self):
        if self.current_step_index >= 0:
            self.start_block_rename(self.current_step_index)

    def start_block_rename(self, index):
        self.editing_block_index = index
        self.refresh_blocks()

    def save_block_name(self, index, new_name):
        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("Lỗi", "Tên khối không được để trống!", parent=self)
            return
        
        self.steps[index]['name'] = new_name
        self.editing_block_index = -1
        
        if index == self.current_step_index:
            self.lbl_step_name.configure(text=f"Editing: {new_name}")
            
        self.refresh_blocks()

    def cancel_block_rename(self):
        self.editing_block_index = -1
        self.refresh_blocks()

    def delete_block(self):
        if self.current_step_index >= 0:
            if messagebox.askyesno("Xác nhận", "Xóa khối lệnh này?", parent=self):
                del self.steps[self.current_step_index]
                self.current_step_index = -1
                self.editing_block_index = -1
                self.lbl_step_name.configure(text="Chọn khối để sửa")
                self.code_text.delete(1.0, tk.END)
                self.refresh_blocks()

    def select_block(self, index):
        if self.editing_block_index != -1 and self.editing_block_index != index:
            self.cancel_block_rename()

        self.current_step_index = index
        self.lbl_step_name.configure(text=f"Editing: {self.steps[index]['name']}")
        
        self.code_text.delete(1.0, tk.END)
        self.code_text.insert(tk.END, self.steps[index]['code'])
        self.code_text.edit_reset()
        
        self.highlight_syntax()
        self.update_line_numbers()
        self.refresh_blocks()

    def move_block(self, direction):
        idx = self.current_step_index
        if idx < 0: return
        new_idx = idx + direction
        if 0 <= new_idx < len(self.steps):
            self.steps[idx], self.steps[new_idx] = self.steps[new_idx], self.steps[idx]
            if self.editing_block_index == idx:
                self.editing_block_index = new_idx
            self.select_block(new_idx)

    def setup_tags(self):
        self.code_text.tag_configure("keyword", foreground="#569cd6")
        self.code_text.tag_configure("string", foreground="#ce9178")
        self.code_text.tag_configure("comment", foreground="#6a9955")
        self.code_text.tag_configure("number", foreground="#b5cea8")
        self.code_text.tag_configure("function", foreground="#dcdcaa")
        self.code_text.tag_configure("match", background="#3e3e42") 
        self.code_text.tag_configure("current_match", background="#d7ba7d", foreground="black")
        self.code_text.tag_raise("current_match")

    def unindent(self, event):
        try:
            try:
                index_start = self.code_text.index("sel.first")
                index_end = self.code_text.index("sel.last")
            except tk.TclError:
                index_start = self.code_text.index("insert")
                index_end = index_start

            start_line = int(index_start.split('.')[0])
            end_line = int(index_end.split('.')[0])
            
            if index_end.split('.')[1] == '0' and start_line != end_line:
                end_line -= 1

            for line in range(start_line, end_line + 1):
                start_pos = f"{line}.0"
                end_pos = f"{line}.4"
                text_start = self.code_text.get(start_pos, end_pos)
                if not text_start.strip(): 
                    self.code_text.delete(start_pos, f"{line}.{len(text_start)}")

            return "break"          
        except Exception as e:
            return "break"

    def auto_highlight_selection(self, event=None):
        self.code_text.tag_remove("match", "1.0", tk.END)
        try:
            if not self.code_text.tag_ranges("sel"):
                return
            selected_text = self.code_text.get("sel.first", "sel.last").strip()
            if len(selected_text) < 2 or len(selected_text) > 50:
                return
            start_pos = "1.0"
            while True:
                pos = self.code_text.search(selected_text, start_pos, stopindex=tk.END, nocase=False)
                if not pos:
                    break
                end_pos = f"{pos}+{len(selected_text)}c"
                self.code_text.tag_add("match", pos, end_pos)
                start_pos = end_pos
        except tk.TclError:
            pass

    def highlight_syntax(self, event=None):
        for tag in ["keyword", "string", "comment", "number", "function"]:
            self.code_text.tag_remove(tag, "1.0", tk.END)
        text = self.code_text.get("1.0", tk.END)
        import re
        keywords = r"\b(import|from|def|class|if|elif|else|while|for|return|pass|break|continue|try|except|finally|True|False|None|and|or|not|in|is|lambda|global|with|as)\b"
        for match in re.finditer(keywords, text): self.code_text.tag_add("keyword", f"1.0+{match.start()}c", f"1.0+{match.end()}c")
        for match in re.finditer(r"\b([a-zA-Z_][a-zA-Z0-9_]*)\s*\(", text): self.code_text.tag_add("function", f"1.0+{match.start(1)}c", f"1.0+{match.end(1)}c")
        for match in re.finditer(r"\b\d+\b", text): self.code_text.tag_add("number", f"1.0+{match.start()}c", f"1.0+{match.end()}c")
        for match in re.finditer(r"(\".*?\"|\'.*?\')", text): self.code_text.tag_add("string", f"1.0+{match.start()}c", f"1.0+{match.end()}c")
        for match in re.finditer(r"#.*", text): self.code_text.tag_add("comment", f"1.0+{match.start()}c", f"1.0+{match.end()}c")

    def update_line_numbers(self):
        self.line_numbers.config(state="normal")
        self.line_numbers.delete("1.0", tk.END)
        line_count = int(self.code_text.index(tk.END).split('.')[0])
        self.line_numbers.insert("1.0", "\n".join(str(i) for i in range(1, line_count)))
        self.line_numbers.config(state="disabled")

    def sync_scroll_code(self, *args):
        self.line_numbers.yview_moveto(args[0])
        self.code_scroll.set(*args)

    def sync_scroll_bar(self, *args):
        self.code_text.yview(*args)
        self.line_numbers.yview(*args)

    def on_scroll(self, event):
        self.code_text.yview_scroll(int(-1*(event.delta/120)), "units")
        self.line_numbers.yview_scroll(int(-1*(event.delta/120)), "units")
        return "break"

    def on_code_change(self, event=None):
        if self.current_step_index >= 0:
            self.steps[self.current_step_index]['code'] = self.code_text.get("1.0", "end-1c")
        self.update_line_numbers()
        if hasattr(self, "_highlight_job"): self.after_cancel(self._highlight_job)
        self._highlight_job = self.after(200, self.highlight_syntax)

    def insert_spaces(self, event):
        try:
            sel_start = self.code_text.index("sel.first")
            sel_end = self.code_text.index("sel.last")
        except tk.TclError:
            self.code_text.insert(tk.INSERT, "    ")
            return 'break'
        start_line = int(sel_start.split('.')[0])
        end_line = int(sel_end.split('.')[0])
        if sel_end.split('.')[1] == '0' and start_line != end_line:
            end_line -= 1
        for line in range(start_line, end_line + 1):
            self.code_text.insert(f"{line}.0", "    ")
        return 'break'

    def redo(self, event=None):
        try: self.code_text.edit_redo()
        except: pass
        return "break"
    
    def select_all(self, event=None):
        self.code_text.tag_add(tk.SEL, "1.0", tk.END)
        self.code_text.mark_set(tk.INSERT, "1.0")
        return 'break'

    def show_find_replace(self):
        if self.current_step_index < 0: return
        FindReplaceDialog(self, self.code_text)

    def save_all(self):
        if self.current_step_index >= 0:
            self.steps[self.current_step_index]['code'] = self.code_text.get("1.0", "end-1c")
        script_manager.scripts[self.category][self.script_index] = self.script_data
        script_manager.save_scripts()
        if self.on_save_callback: self.on_save_callback()

class VideoPlayer:
    def __init__(self, parent, image_folder, width=199, height=150):
        self.parent = parent
        self.image_folder = image_folder
        self.width = width
        self.height = height
        self.cap = None
        self.is_playing = False
        
        self.idle_files = ["Idle_1.mp4", "Idle_2.mp4", "Idle_3.mp4", "Idle_4.mp4"]
        self.idle_weights = [0.55, 0.15, 0.15, 0.15] 
        
        self.current_mode = "idle" 
        
        self.label = ctk.CTkLabel(parent, text="")
        self.label.pack_forget()

    def start_active(self, video_path):
        self.current_mode = "active"
        self._load_video(video_path)

    def start_idle(self):
        self.current_mode = "idle"
        self._play_random_idle()

    def _play_random_idle(self):
        try:
            chosen_file = random.choices(self.idle_files, weights=self.idle_weights, k=1)[0]
            video_path = os.path.join(self.image_folder, chosen_file)
            
            if not os.path.exists(video_path):
                for f in self.idle_files:
                    p = os.path.join(self.image_folder, f)
                    if os.path.exists(p):
                        video_path = p
                        break
            
            self._load_video(video_path)
        except Exception as e:
            logger.error(f"Lỗi chọn video idle: {e}")

    def _load_video(self, path):
        if self.cap:
            self.cap.release()
        
        if not os.path.exists(path):
            return

        self.cap = cv2.VideoCapture(path)
        self.is_playing = True
        self.label.pack(pady=10, side="bottom")
        
        if not hasattr(self, 'update_job') or self.update_job is None:
            self.update_frame()

    def stop(self):
        self.is_playing = False
        if self.cap:
            self.cap.release()
            self.cap = None
        self.label.pack_forget()
        self.label.configure(image=None)
        self.update_job = None

    def update_frame(self):
        if not self.is_playing or not self.cap:
            self.update_job = None
            return

        if not self.label.winfo_viewable():
            self.update_job = self.label.after(500, self.update_frame)
            return

        ret, frame = self.cap.read()
        
        if not ret:
            if self.current_mode == "active":
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                ret, frame = self.cap.read()
            else:
                self._play_random_idle()
                if self.cap:
                    ret, frame = self.cap.read()

        if ret:
            try:
                frame = cv2.resize(frame, (self.width, self.height), interpolation=cv2.INTER_LINEAR)
                
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame)
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(self.width, self.height))
                
                self.label.configure(image=ctk_img)
                self.label.image = ctk_img
            except Exception:
                pass

            self.update_job = self.label.after(25, self.update_frame)
        else:
            self.update_job = self.label.after(100, self.update_frame)

if __name__ == "__main__":
    app = AutoACSTool()
    app.mainloop()