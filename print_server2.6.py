#!/usr/bin/env python
# -*- coding: future_fstrings -*-
#作者：忆痕
#仓库地址：https://github.com/a937750307/lan-printing
 
import os
from flask import Flask, request, render_template_string, send_from_directory, redirect, url_for, flash, jsonify
# 打印相关
import win32print
import win32api
import win32con
import subprocess
import time
import requests
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from pdf2image import convert_from_path
# 扫描相关
import comtypes.client as cc
import win32com.client as win32com_client
# 托盘相关
import threading
import sys
import pystray
from PIL import Image, ImageDraw
import socket
import winreg
import json
import math
import io
#编码相关
import ctypes
import locale
import parse
import chardet
#系统相关
import platform
import wmi
import werkzeug

# Windows DeviceCapabilities 常量
DC_DUPLEX = 7
DC_COLORDEVICE = 32
DC_PAPERS = 2
DC_PAPERNAMES = 16
DC_ENUMRESOLUTIONS = 13
DC_ORIENTATION = 17
DC_COPIES = 18
DC_TRUETYPE = 28
DC_DRIVER = 11

# 控制台窗口相关全局变量
CONSOLE_WINDOW = None
CONSOLE_VISIBLE = True

# IP配置状态跟踪
IP_CONFIG_STATE = {'is_static': False, 'last_set_ip': None}

# 打印和扫描互斥状态管理
DEVICE_STATUS = {
    'is_printing': False,
    'is_scanning': False,
    'print_start_time': None,
    'scan_start_time': None,
    'print_client': '',
    'scan_client': ''
}

# 调试模式：如果True，则显示虚拟打印机（用于无物理打印机的测试环境）
# 调试模式：如果False，则不显示虚拟打印机
# 设置环境变量 PRINTING_DEBUG=1 或在此修改来启用
DEBUG_MODE = False

# Windows纸张大小常量
DMPAPER_LETTER = 1
DMPAPER_A4 = 9
DMPAPER_A3 = 8
DMPAPER_A5 = 11
DMPAPER_B4 = 12
DMPAPER_B5 = 13
DMPAPER_LEGAL = 5
DMPAPER_EXECUTIVE = 7
DMPAPER_TABLOID = 3

# 纸张名称映射
PAPER_NAMES = {
    1: "Letter (8.5 x 11 in)",
    3: "Tabloid (11 x 17 in)",
    5: "Legal (8.5 x 14 in)",
    7: "Executive (7.25 x 10.5 in)",
    8: "A3 (297 x 420 mm)",
    9: "A4 (210 x 297 mm)",
    11: "A5 (148 x 210 mm)",
    12: "B4 (250 x 354 mm)",
    13: "B5 (182 x 257 mm)",
}


def ensure_printer_connection(pr_name):
    """确保对 UNC 网络共享打印机建立临时连接。
    对以 \\ 开头的 printer 名称尝试 AddPrinterConnection，若失败则使用 printui 作为备用。
    返回 True 表示已尝试连接（不保证打印机会可用），False 表示发生异常。
    """
    """确保对 UNC 网络共享打印机建立临时连接"""
    if not pr_name:
        return False
    pn = pr_name.strip()
    if pn.startswith('\\'):
        try:
            import win32print
            print(f"尝试连接网络共享打印机: {pn}")
            try:
                win32print.AddPrinterConnection(pn)
                print(f"已添加打印机连接: {pn}")
            except Exception as e:
                print(f"AddPrinterConnection 失败: {e}")
                try:
                    import subprocess
                    cmd = ['rundll32.exe', 'printui.dll,PrintUIEntry', '/in', '/n', pn]
                    subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
                    print(f"已尝试通过 printui 添加打印机: {pn}")
                except Exception as e2:
                    print(f"通过 printui 添加打印机失败: {e2}")
            if 'printer_cache' in globals():
                try:
                    printer_cache.refresh_cache()
                    print("打印机缓存已刷新")
                except Exception:
                    pass
        except Exception as e:
            print(f"ensure_printer_connection 内部错误: {e}")
            return False
    return True


class PathManager:
    """统一路径管理器"""
    
    def __init__(self):
        self._is_packaged = hasattr(sys, '_MEIPASS')
        if self._is_packaged:
            self._resource_dir = sys._MEIPASS
            self._app_dir = os.path.dirname(sys.executable)
            self._data_dir = self._app_dir
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self._resource_dir = script_dir
            self._app_dir = script_dir
            self._data_dir = script_dir
    
    @property
    def is_packaged(self):
        return self._is_packaged
    @property
    def app_dir(self):
        return self._app_dir
    
    def get_resource_path(self, relative_path):
        return os.path.join(self._resource_dir, relative_path)
    def get_data_path(self, relative_path):
        return os.path.join(self._data_dir, relative_path)
    def get_config_path(self):
        return self.get_data_path('config.json')
    def get_log_path(self):
        return self.get_data_path('print_log.txt')
    def get_upload_dir(self):
        return self.get_data_path('uploads')
    def get_scan_dir(self):
        return self.get_data_path('scanned_files')
    def get_executable_name(self):
        return os.path.basename(sys.executable) if self._is_packaged else os.path.basename(sys.argv[0])
    def ensure_data_dirs(self):
        try:
            os.makedirs(self.get_upload_dir(), exist_ok=True)
            os.makedirs(self.get_scan_dir(), exist_ok=True)
            return True
        except Exception as e:
            print(f"创建数据目录失败: {e}")
            return False

# 创建全局路径管理器实例
path_manager = PathManager()


def get_poppler_path():
    """确定用于 pdf2image 的 poppler 可执行文件路径。
    优先级：环境变量 `POPPLER_PATH` -> path_manager 中配置（若有）-> 打包内置目录 -> 项目相对 `third_party/poppler`。
    返回 None 表示使用系统 PATH 查找。
    """
    try:
        # 1. 环境变量覆盖
        env_path = os.environ.get('POPPLER_PATH')
        if env_path and os.path.isdir(env_path):
            return env_path

        # 2. path_manager 支持（如果实现了此方法）
        try:
            pm_path = getattr(path_manager, 'get_poppler_path', None)
            if pm_path:
                p = pm_path()
                if p and os.path.isdir(p):
                    return p
        except Exception:
            pass

        # 3. PyInstaller 打包时资源会解压到 sys._MEIPASS
        base = None
        if getattr(sys, '_MEIPASS', None):
            base = sys._MEIPASS
        else:
            base = os.path.abspath(os.path.dirname(__file__))

        # 不要写死路径；检查若干常见位置，包括仓库根目录下的 `poppler` 文件夹
        candidates = [
            os.path.join(base, 'poppler_bin'),
            os.path.join(base, 'third_party', 'poppler', 'Library', 'bin'),
            os.path.join(base, 'third_party', 'poppler', 'bin'),
            os.path.join(base, 'poppler', 'Library', 'bin'),
            os.path.join(base, 'poppler', 'bin'),
            os.path.join(base, 'poppler'),
        ]

        for c in candidates:
            if os.path.isdir(c):
                return c

    except Exception:
        pass
    return None


# 在程序启动时打印 poppler 路径检测结果，便于调试和确认
try:
    _detected_poppler = get_poppler_path()
    if _detected_poppler:
        print(f"Poppler path detected: {_detected_poppler}")
    else:
        print("Poppler path not detected: will use system PATH or pdf2image default")
except Exception as _e:
    print(f"检测 Poppler 路径时出错: {_e}")

# 其他启动时诊断信息，便于快速确认运行环境
try:
    print(f"Python executable: {sys.executable}")
    print(f"Working directory: {os.getcwd()}")
    # pdf2image 可用性
    try:
        import pdf2image
        print("pdf2image available")
    except Exception as _e:
        print(f"pdf2image not available: {_e}")

    # 检查默认打印机（如果 win32print 可用）
    try:
        import win32print
        try:
            default_pr = win32print.GetDefaultPrinter()
            print(f"Default printer: {default_pr}")
        except Exception:
            print("Default printer: 未能检测到")
    except Exception:
        print("win32print not available, 跳过默认打印机检测")
except Exception as _e:
    print(f"启动时环境诊断出错: {_e}")

# 全局服务状态管理
class ServiceManager:
    """服务管理器，用于管理Flask服务和程序重启 - 优化稳定性版本"""
    def __init__(self):
        self.flask_thread = None
        self.cleaner_thread = None
        self.monitor_thread = None
        self.should_restart = False
        self.restart_port = None
        self.service_running = False
        self.last_health_check = time.time()
        self.health_check_interval = 600  # 10分钟检查一次
        self.health_fail_count = 0  # 健康检查失败计数
        self.start_time = None  # 服务启动时间
        
        # 根据Windows版本优化参数
        self._optimize_for_windows_version()
        
        self.is_shutting_down = False  # 是否正在关闭
    
    def _optimize_for_windows_version(self):
        """统一服务管理参数，不再区分Windows版本"""
        # 统一参数设置
        self.health_check_interval = 600  # 10分钟检查一次
        self.max_restart_attempts = 5     # 最多重启5次
        self.restart_cooldown = 300       # 5分钟冷却
        self.restart_count = 0
        self.last_restart_time = 0
        print("服务管理参数已统一设置：检查间隔600秒，最多重启5次，冷却300秒")
    
    def set_restart(self, port):
        """设置重启标志和新端口"""
        self.should_restart = True
        self.restart_port = port
    
    def is_restart_requested(self):
        """检查是否需要重启"""
        return self.should_restart
    
    def get_restart_port(self):
        """获取重启端口"""
        return self.restart_port
    
    def clear_restart(self):
        """清除重启标志"""
        self.should_restart = False
        self.restart_port = None
    
    def mark_service_running(self):
        """标记服务运行状态"""
        self.service_running = True
        self.last_health_check = time.time()
        if self.start_time is None:  # 只在第一次启动时设置
            self.start_time = time.time()
        self.health_fail_count = 0  # 重置失败计数
    
    def mark_service_stopped(self):
        """标记服务停止状态"""
        self.service_running = False
    
    def is_service_healthy(self):
        """检查服务健康状态"""
        if not self.service_running:
            return False
        # 检查Flask线程是否仍然活跃
        if self.flask_thread and not self.flask_thread.is_alive():
            return False
        return True
    
    def update_health_check(self):
        """更新健康检查时间"""
        self.last_health_check = time.time()
    
    def restart_flask_service(self):
        """重启Flask服务 - 增强稳定性版本"""
        current_time = time.time()
        
        # 检查是否正在关闭
        if self.is_shutting_down:
            print("程序正在关闭，跳过服务重启")
            return False
        
        # 检查重启冷却时间
        if current_time - self.last_restart_time < self.restart_cooldown:
            remaining = self.restart_cooldown - (current_time - self.last_restart_time)
            print(f"重启冷却中，还需等待 {remaining:.1f} 秒")
            return False
        
        # 检查重启次数限制
        if self.restart_count >= self.max_restart_attempts:
            print(f"已达到最大重启次数 ({self.max_restart_attempts})，停止重启")
            return False
        
        try:
            self.restart_count += 1
            self.last_restart_time = current_time
            
            print(f"检测到服务异常，正在重启Flask服务... (第{self.restart_count}次)")
            
            # 标记停止状态
            self.mark_service_stopped()
            
            # 等待旧线程结束
            if self.flask_thread and self.flask_thread.is_alive():
                print("等待旧服务线程结束...")
                self.flask_thread.join(timeout=10)  # 最多等10秒
                if self.flask_thread.is_alive():
                    print("️旧服务线程未能及时结束")
            
            # 短暂延迟，让端口完全释放
            time.sleep(2)
            
            # 启动新服务
            port = getattr(app, 'current_port', 5000)
            app.current_port = port
            
            if os.environ.get('USE_WSGI', '').lower() == 'true':
                self.flask_thread = threading.Thread(target=run_wsgi, daemon=True, name="FlaskWSGI")
            else:
                self.flask_thread = threading.Thread(target=run_flask, daemon=True, name="FlaskDev")
            
            self.flask_thread.start()
            
            # 等待服务启动
            time.sleep(3)
            
            if self.flask_thread.is_alive():
                self.mark_service_running()
                print(f" Flask服务重启成功 (第{self.restart_count}次)")
                
                # 重启成功后重置部分计数
                if self.restart_count >= 3:
                    print(" 连续重启成功，重置重启计数")
                    self.restart_count = max(0, self.restart_count - 2)
                
                return True
            else:
                print(f" Flask服务重启失败，线程未能启动")
                return False
            
        except Exception as e:
            print(f"Flask服务重启异常: {e}")
            return False

service_manager = ServiceManager()

def clean_old_files(folder=None, expire_seconds=3600):
    """定期清理指定目录下超过expire_seconds的文件，并启动日志清理 - 优化I/O版本"""
    if folder is None:
        folder = path_manager.get_upload_dir()
    
    # 启动日志清理线程（只启动一次）
    if not hasattr(clean_old_files, 'log_cleanup_started'):
        import threading
        log_cleanup_thread = threading.Thread(target=periodic_log_cleanup, daemon=True)
        log_cleanup_thread.start()
        clean_old_files.log_cleanup_started = True
        print("日志自动清理功能已启动")
    
    # 启动扫描文件清理线程（只启动一次）
    if not hasattr(clean_old_files, 'scan_cleanup_started'):
        import threading
        scan_cleanup_thread = threading.Thread(target=periodic_scan_cleanup, daemon=True)
        scan_cleanup_thread.start()
        clean_old_files.scan_cleanup_started = True
        print("扫描文件自动清理功能已启动（30分钟过期）")
    
    last_cleanup_time = 0
    cleanup_interval = 300  # 5分钟检查一次，降低频率
    files_to_check = []
    
    while True:
        current_time = time.time()
        
        # 降低检查频率
        if current_time - last_cleanup_time < cleanup_interval:
            time.sleep(30)
            continue
            
        try:
            # 批量获取文件列表，减少系统调用
            if os.path.exists(folder):
                files_to_check = []
                try:
                    for fname in os.listdir(folder):
                        fpath = os.path.join(folder, fname)
                        if os.path.isfile(fpath):
                            files_to_check.append((fpath, fname))
                except (OSError, PermissionError) as e:
                    print(f"️ 扫描上传目录失败: {e}")
                    time.sleep(60)
                    continue
                
                # 批量检查和删除过期文件
                deleted_count = 0
                for fpath, fname in files_to_check:
                    try:
                        file_age = current_time - os.path.getmtime(fpath)
                        if file_age > 600:  # 10分钟
                            os.remove(fpath)
                            deleted_count += 1
                    except (OSError, FileNotFoundError):
                        # 文件可能已被其他进程删除，忽略错误
                        continue
                    except Exception as e:
                        # 记录其他异常但继续执行
                        print(f"️ 删除文件 {fname} 失败: {e}")
                
                if deleted_count > 0:
                    print(f"文件清理: 删除了 {deleted_count} 个过期文件")
                
                last_cleanup_time = current_time
            else:
                # 上传目录不存在，尝试创建
                try:
                    os.makedirs(folder, exist_ok=True)
                except Exception as e:
                    print(f"️ 创建上传目录失败: {e}")
                    
        except Exception as e:
            print(f"文件清理异常: {e}")
            
        # 动态调整清理间隔，长时间稳定后降低清理频率
        uptime = current_time - (service_manager.start_time or current_time)
        if uptime > 7200:  # 2小时后
            cleanup_interval = 600  # 10分钟检查一次
        elif uptime > 3600:  # 1小时后
            cleanup_interval = 450  # 7.5分钟检查一次
        
        time.sleep(cleanup_interval)

def periodic_scan_cleanup():
    """定期清理扫描文件 - 30分钟过期策略"""
    scan_folder = path_manager.get_scan_dir()
    expire_seconds = 1800  # 30分钟
    
    last_cleanup_time = 0
    cleanup_interval = 300  # 5分钟检查一次
    
    while True:
        current_time = time.time()
        
        # 降低检查频率
        if current_time - last_cleanup_time < cleanup_interval:
            time.sleep(30)
            continue
        
        try:
            if os.path.exists(scan_folder):
                files_to_check = []
                try:
                    for fname in os.listdir(scan_folder):
                        fpath = os.path.join(scan_folder, fname)
                        if os.path.isfile(fpath):
                            files_to_check.append((fpath, fname))
                except (OSError, PermissionError) as e:
                    print(f"⚠️ 扫描文件夹失败: {e}")
                    time.sleep(60)
                    continue
                
                # 批量检查和删除过期文件（30分钟）
                deleted_count = 0
                for fpath, fname in files_to_check:
                    try:
                        file_age = current_time - os.path.getmtime(fpath)
                        if file_age > expire_seconds:  # 30分钟
                            os.remove(fpath)
                            deleted_count += 1
                    except (OSError, FileNotFoundError):
                        continue
                    except Exception as e:
                        print(f"⚠️ 删除扫描文件 {fname} 失败: {e}")
                
                if deleted_count > 0:
                    print(f"扫描文件清理: 删除了 {deleted_count} 个过期文件（30分钟以上）")
                
                last_cleanup_time = current_time
        except Exception as e:
            print(f"扫描文件清理异常: {e}")
        
        time.sleep(cleanup_interval)

def monitor_service_health():
    """监控服务健康状态，发现异常时自动重启 - 优化性能版本"""
    
    startup_message_shown = False
    last_check_time = 0
    while True:
        try:
            current_time = time.time()
            # 动态调整检查间隔
            if hasattr(service_manager, 'start_time') and service_manager.start_time:
                uptime = current_time - service_manager.start_time
                if uptime > 3600:
                    check_interval = 1800
                elif uptime > 1800:
                    check_interval = 900
                else:
                    check_interval = service_manager.health_check_interval
            else:
                check_interval = service_manager.health_check_interval
            if current_time - last_check_time < check_interval:
                time.sleep(30)
                continue
            last_check_time = current_time
            if not service_manager.is_service_healthy():
                print("️ 检测到Flask服务异常")
                service_manager.restart_flask_service()
                startup_message_shown = False
                continue
            if service_manager.start_time and (current_time - service_manager.start_time) < 10:
                if not startup_message_shown:
                    print(" 服务启动中，健康检查暂停10秒...")
                    startup_message_shown = True
                continue
            # 轻量级socket检查
            try:
                import socket
                port = getattr(app, 'current_port', 5000)
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                result = sock.connect_ex(('127.0.0.1', port))
                sock.close()
                if result == 0:
                    service_manager.update_health_check()
                    if hasattr(service_manager, 'health_fail_count'):
                        service_manager.health_fail_count = 0
                else:
                    if not hasattr(service_manager, 'health_fail_count'):
                        service_manager.health_fail_count = 0
                    service_manager.health_fail_count += 1
                    if service_manager.health_fail_count >= 2:
                        print(f"️ Socket健康检查失败: {service_manager.health_fail_count}/2")
                    if service_manager.health_fail_count >= 2:
                        print(" 连续Socket检查失败，重启服务")
                        service_manager.restart_flask_service()
                        service_manager.health_fail_count = 0
            except Exception as e:
                print(f"️ Socket健康检查异常: {e}")
                service_manager.health_fail_count = getattr(service_manager, 'health_fail_count', 0) + 1
                if service_manager.health_fail_count >= 2:
                    service_manager.restart_flask_service()
                    service_manager.health_fail_count = 0
        except Exception as e:
            print(f"服务监控异常: {e}")
            time.sleep(30)

# 使用路径管理器获取配置文件路径
CONFIG_FILE = path_manager.get_config_path()

def load_config():
    """加载配置文件"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                print(f" 配置文件加载成功: {CONFIG_FILE}")
                return config
        print("配置文件不存在，使用默认配置")
        return {}
    except Exception as e:
        print(f"配置文件加载失败: {e}，使用默认配置")
        return {}

def save_config(config):
    """保存配置文件"""
    try:
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"配置已保存到: {CONFIG_FILE}")
        return True
    except Exception as e:
        print(f"配置保存失败: {e}")
        return False

def get_config_port():
    """从配置文件获取端口号"""
    config = load_config()
    return config.get('port', 5000)  # 默认端口5000

def save_port_config(port):
    """保存端口配置"""
    config = load_config()
    config['port'] = port
    return save_config(config)

# ================== 扫描功能相关代码 ==================

def get_available_scanners():
    """获取系统中可用的扫描仪列表 - 简化版本，不依赖PowerShell"""
    scanners = []
    try:
        # 简单缓存，避免频繁枚举设备
        if not hasattr(get_available_scanners, '_cache'):
            get_available_scanners._cache = {'time': 0, 'scanners': []}
            get_available_scanners._cache_timeout = 60  # 60秒缓存
        
        import time
        if time.time() - get_available_scanners._cache['time'] < get_available_scanners._cache_timeout:
            return list(get_available_scanners._cache['scanners'])

        # 方法1: 使用 WIA COM 直接枚举（最通用，所有Windows都有）
        try:
            import threading
            wia_scanners = []
            
            def try_wia():
                try:
                    import win32com.client
                    dm = win32com.client.Dispatch('WIA.DeviceManager')
                    for i in range(1, min(dm.DeviceInfos.Count + 1, 10)):  # 最多检查10个设备
                        try:
                            dev = dm.DeviceInfos.Item(i)
                            name = dev.Properties('Name').Value
                            dev_id = dev.Properties('DeviceID').Value
                            wia_scanners.append({
                                'name': name,
                                'id': dev_id,
                                'type': 'WIA',
                                'available': True
                            })
                            print(f"检测到扫描仪: {name}")
                        except:
                            pass
                except Exception as e:
                    print(f"WIA枚举失败: {e}")
            
            # 用线程执行WIA，设置超时防止卡死
            thread = threading.Thread(target=try_wia, daemon=True)
            thread.start()
            thread.join(timeout=5)
            
            if wia_scanners:
                scanners.extend(wia_scanners)
        except Exception as e:
            print(f"WIA 检测异常: {e}")

        # 方法2: 如果WIA未找到，尝试 WMIC 查询（Windows Management Instrumentation Command）
        if not scanners:
            try:
                import subprocess
                # WMIC 是所有Windows都有的，不需要PowerShell
                cmd = 'wmic logicaldisk get name'
                result = subprocess.run([cmd], shell=True, capture_output=True, text=True, 
                                      timeout=5, creationflags=subprocess.CREATE_NO_WINDOW)
                
                # 简单的WMIC查询不一定能找到扫描仪，但至少验证cmd可用
                # 如果需要详细信息，可以尝试其他WMI查询
                if result.returncode == 0:
                    print("✓ WMIC 可用，继续使用WIA枚举")
            except Exception as e:
                print(f"WMIC 检测失败: {e}")

        # 方法3: 基于打印机推断扫描设备（许多多功能设备同时支持打印和扫描）
        try:
            printers = globals().get('PRINTERS') or []
            scan_keywords = ['scan', '扫描', 'mfp', 'multi', 'all-in-one', 'all in one', '多功能']
            
            for p in list(printers)[:30]:
                pname = str(p)
                lname = pname.lower()
                
                # 如果打印机名称包含扫描相关关键词，作为候选扫描仪
                if any(k in lname for k in scan_keywords):
                    if not any(s['name'] == pname for s in scanners):
                        scanners.append({
                            'name': pname,
                            'id': f'PRINTER_{pname}',
                            'type': 'Multifunction',
                            'available': True
                        })
                        print(f"检测到多功能设备: {pname}")
        except Exception as e:
            print(f"多功能设备推断失败: {e}")

        # 如果仍未找到扫描仪，添加默认项
        if not scanners:
            scanners.append({
                'name': '通用扫描（系统窗口）',
                'id': 'default',
                'type': 'Generic',
                'available': True
            })
            print("未检测到具体扫描仪，已添加通用选项")

        # 更新缓存
        try:
            get_available_scanners._cache['time'] = time.time()
            get_available_scanners._cache['scanners'] = list(scanners)
        except:
            pass
        
        return scanners
    
    except Exception as e:
        print(f"扫描仪检测出错: {e}")
        return [{
            'name': '通用扫描（系统窗口）',
            'id': 'default',
            'type': 'Generic',
            'available': True
        }]

def cleanup_port_and_restart_wia(port=5000):
    """强制清理端口占用标记并重启WIA服务"""
    try:
        print(f"[CLEANUP] 开始清理端口占用和重启WIA服务...")
        
        # 步骤1: 强制关闭占用指定端口的进程
        try:
            print(f"[PORT] 尝试清理端口 {port}...")
            result = subprocess.run(
                ['netstat', '-ano'],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if f':{port}' in line and 'ESTABLISHED' in line:
                        parts = line.split()
                        if len(parts) > 0:
                            try:
                                pid = parts[-1]
                                subprocess.run(
                                    ['taskkill', '/F', '/PID', pid],
                                    capture_output=True,
                                    timeout=5,
                                    creationflags=subprocess.CREATE_NO_WINDOW
                                )
                                print(f"[PORT] 已清理占用端口 {port} 的进程 (PID: {pid})")
                            except:
                                pass
        except Exception as e:
            print(f"[WARN] 端口清理异常: {e}")
        
        # 步骤2: 强制清理socket状态
        try:
            print("[CLEANUP] 清理socket挂起状态...")
            import gc
            gc.collect()
            print("[CLEANUP] Socket状态已清空")
        except:
            pass
        
        # 步骤3: 停止WIA服务
        try:
            print("[WIA] 停止WIA服务...")
            result = subprocess.run(
                ['sc', 'stop', 'wiaservc'],
                capture_output=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode == 0:
                print("[WIA] WIA服务已停止")
                import time
                time.sleep(1)
        except Exception as e:
            print(f"[WARN] 停止WIA服务异常: {e}")
        
        # 步骤4: 启动WIA服务
        try:
            print("[WIA] 启动WIA服务...")
            result = subprocess.run(
                ['sc', 'start', 'wiaservc'],
                capture_output=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode == 0:
                print("[WIA] WIA服务已启动")
                return True
        except Exception as e:
            print(f"[WARN] 启动WIA服务异常: {e}")
        
        print("[SUCCESS] 端口和WIA服务清理完成")
        return True
        
    except Exception as e:
        print(f"[ERROR] 清理端口和重启WIA失败: {e}")
        return False

def force_release_wia_device():
    """强制释放被锁定的WIA设备"""
    import subprocess
    try:
        print("尝试强制释放WIA设备...")
        
        # 方法1: 使用 sc 命令停止 WIA 服务（比 net 更可靠）
        try:
            result = subprocess.run(['sc', 'stop', 'wiaservc'],
                                  capture_output=True,
                                  creationflags=subprocess.CREATE_NO_WINDOW,
                                  timeout=10)
            stderr_msg = result.stderr.decode('utf-8', errors='ignore') if result.stderr else ""
            stdout_msg = result.stdout.decode('utf-8', errors='ignore') if result.stdout else ""
            
            print(f"WIA服务停止: 返回码={result.returncode} {stdout_msg.strip()}")
            
            import time
            time.sleep(2)  # 等待服务完全停止
            
            # 启动服务
            result = subprocess.run(['sc', 'start', 'wiaservc'],
                                  capture_output=True,
                                  creationflags=subprocess.CREATE_NO_WINDOW,
                                  timeout=10)
            stdout_msg = result.stdout.decode('utf-8', errors='ignore') if result.stdout else ""
            print(f"WIA服务启动: 返回码={result.returncode} {stdout_msg.strip()}")
            
            time.sleep(1)
            return True
        except Exception as e:
            print(f"WIA服务重启失败: {e}")
        
        # 方法2: 杀死占用扫描仪的进程
        process_names = [
            'scanwiz.exe',      # Windows 图片采集向导
            'wiafbdrv.exe',     # WIA 驱动框架
            'svchost.exe',      # 可能托管WIA服务
            'wiaservc.exe',     # WIA 服务可执行文件
            'mspaint.exe',      # 画图
            'explorer.exe'      # 文件管理器
        ]
        
        killed_processes = []
        for process_name in process_names:
            try:
                result = subprocess.run(['taskkill', '/F', '/IM', process_name],
                                     capture_output=True,
                                     creationflags=subprocess.CREATE_NO_WINDOW,
                                     timeout=5)
                if result.returncode == 0:
                    killed_processes.append(process_name)
                    print(f"已清理进程: {process_name}")
            except:
                pass
        
        if killed_processes:
            print(f"成功清理进程: {', '.join(killed_processes)}")
            return True
        
        # 方法3: 尝试使用 Python 进程管理清理
        try:
            import psutil
            wia_related = ['scanwiz.exe', 'wiafbdrv.exe', 'wiaservc.exe']
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'].lower() in wia_related:
                        p = psutil.Process(proc.info['pid'])
                        p.terminate()
                        print(f"已终止进程: {proc.info['name']}")
                except:
                    pass
            return True
        except ImportError:
            print("psutil 未安装，跳过进程管理器方法")
        except Exception as e:
            print(f"进程管理器方法失败: {e}")
        
        # 方法4: 清理 COM 对象缓存（内存层面）
        try:
            import gc
            import ctypes
            gc.collect()
            print("已清理COM对象缓存和内存")
        except:
            pass
        
        print("WIA设备已释放")
        import sys
        sys.stdout.flush()
        return True
    except Exception as e:
        print(f"强制释放WIA设备异常: {e}")
        print("WIA设备释放失败")
        return False

def start_scan_silent(scanner_id, scanner_name, scan_format='PNG', scan_path=None):
    """扫描功能 - 全程静默扫描，不打开任何窗口或文件夹"""
    import subprocess
    import datetime
    import time
    import threading
    import winreg
    
    try:
        print(f"启动扫描: {scanner_name} ({scanner_id})")
        
        if scan_path is None:
            # 创建扫描文件专用目录 - 统一使用PathManager管理
            scan_path = path_manager.get_scan_dir()
            if not os.path.exists(scan_path):
                os.makedirs(scan_path)
        
        # 获取开始时扫描目录中的文件
        try:
            initial_files = set(os.listdir(scan_path)) if os.path.exists(scan_path) else set()
        except:
            initial_files = set()
        
        # 保存和禁用自动打开文件夹的功能
        original_auto_open = None
        
        def disable_auto_open_folder():
            """禁用Windows自动打开扫描完成后的文件夹"""
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers",
                                    0, winreg.KEY_READ)
                winreg.CloseKey(key)
                
                # 禁用USB设备自动运行
                try:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                        r"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer",
                                        0, winreg.KEY_WRITE)
                    original_auto_open = winreg.QueryValueEx(key, 'NoDriveTypeAutorun')[0]
                    winreg.SetValueEx(key, 'NoDriveTypeAutorun', 0, winreg.REG_DWORD, 255)
                    winreg.CloseKey(key)
                except:
                    pass
            except Exception as e:
                print(f"禁用自动打开失败: {e}")
        
        def restore_auto_open_folder():
            """恢复自动打开文件夹功能"""
            try:
                if original_auto_open is not None:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                        r"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer",
                                        0, winreg.KEY_WRITE)
                    winreg.SetValueEx(key, 'NoDriveTypeAutorun', 0, winreg.REG_DWORD, original_auto_open)
                    winreg.CloseKey(key)
            except:
                pass
        
        disable_auto_open_folder()
        print("正在进行静默扫描，请在扫描仪上操作...")
        
        try:
            # 方案1: 使用WIA COM对象进行直接的静默扫描（无窗口）
            def silent_wia_scan():
                device_manager = None
                device = None
                item = None
                image = None
                try:
                    import win32com.client
                    import ctypes
                    import gc
                    
                    # 初始化COM环境
                    try:
                        ctypes.windll.ole32.CoInitialize(None)
                    except:
                        pass
                    
                    try:
                        # 初始化WIA设备管理器
                        device_manager = win32com.client.Dispatch("WIA.DeviceManager")
                        devices = device_manager.DeviceInfos
                        
                        if len(devices) > 0:
                            # 使用第一个扫描仪设备
                            device_info = devices(1)
                            device = device_info.Connect()
                            
                            # 设置扫描参数
                            item = device.Items(1)
                            scanner_item = item
                            
                            # 配置扫描设置（颜色、分辨率等）
                            for prop in scanner_item.Properties:
                                if prop.Name == 'Horizontal Resolution':
                                    prop.Value = 200  # 200dpi
                                elif prop.Name == 'Vertical Resolution':
                                    prop.Value = 200
                                elif prop.Name == 'Color Mode':
                                    prop.Value = 1  # 彩色
                            
                            # 执行扫描
                            image = scanner_item.Transfer()
                            
                            # 保存扫描结果
                            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"scan_{timestamp}.{scan_format.lower() if scan_format else 'bmp'}"
                            filepath = os.path.join(scan_path, filename)
                            
                            # 保存图像
                            image.SaveFile(filepath)
                            
                            print(f"WIA静默扫描完成: {filename}")
                            result_path = filepath
                            
                            # 立即释放COM对象，不与其他用户冲突
                            try:
                                if image is not None:
                                    image = None
                                if item is not None:
                                    item = None
                                if scanner_item is not None:
                                    scanner_item = None
                                if device_info is not None:
                                    device_info = None
                                if device is not None:
                                    device = None
                                if device_manager is not None:
                                    device_manager = None
                                gc.collect()
                            except:
                                pass
                            
                            # 释放COM环境
                            try:
                                ctypes.windll.ole32.CoUninitialize()
                            except:
                                pass
                            
                            return True, result_path
                    finally:
                        # 最终安全检查：确保所有对象都被释放
                        try:
                            if image is not None:
                                image = None
                            if item is not None:
                                item = None
                            if device is not None:
                                device = None
                            if device_manager is not None:
                                device_manager = None
                        except:
                            pass
                        
                        # 释放COM环境
                        try:
                            ctypes.windll.ole32.CoUninitialize()
                        except:
                            pass
                        
                except Exception as e:
                    print(f"WIA扫描失败: {e}")
                    return False, None
            
            # 尝试WIA静默扫描
            success, filepath = silent_wia_scan()
            if success and filepath:
                file_size = os.path.getsize(filepath)
                filename = os.path.basename(filepath)
                print(f"扫描完成: {filename} ({file_size} 字节)")
                print("设备已释放")
                import sys
                sys.stdout.flush()
                restore_auto_open_folder()
                return True, f"扫描成功！文件已保存到: {filename}"
            
            # 如果WIA扫描失败，检查是否是设备被锁定导致的
            print("WIA直接扫描未成功，尝试强制释放设备...")
            force_release_wia_device()
            import time
            time.sleep(2)  # 等待设备释放
            
            # 再次尝试WIA静默扫描
            print("重新尝试WIA扫描...")
            success, filepath = silent_wia_scan()
            if success and filepath:
                file_size = os.path.getsize(filepath)
                filename = os.path.basename(filepath)
                print(f"扫描完成: {filename} ({file_size} 字节)")
                print("释放后扫描成功")
                import sys
                sys.stdout.flush()
                restore_auto_open_folder()
                return True, f"扫描成功！文件已保存到: {filename}"
            
            # 方案3: 使用scanimage命令行工具进行静默扫描（如果已安装）或使用更简单的WIA方法
            print("使用静默扫描命令...")
            vbs_success = False
            
            try:
                # 生成扫描文件名
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                scan_filename = f"scan_{timestamp}.jpg"
                scan_filepath = os.path.join(scan_path, scan_filename)
                
                # 尝试方案A: 使用scanimage（SANE - 如果已安装）
                try:
                    result = subprocess.run(['scanimage', '--scan-mode', 'Color', '--resolution', '200', '-o', scan_filepath],
                                          capture_output=True,
                                          creationflags=subprocess.CREATE_NO_WINDOW,
                                          timeout=60)
                    if result.returncode == 0 and os.path.exists(scan_filepath):
                        print(f"scanimage扫描成功: {scan_filename}")
                        restore_auto_open_folder()
                        return True, f"扫描成功！文件已保存到: {scan_filename}"
                except:
                    pass
                
                # 方案B: 使用Windows WIA COM自动化 - 改进版本（移除类型检查）
                vbs_script = f"""
Dim objScanner, objDevice, objItem, objImage, objDeviceInfo, fso
Dim deviceCount, i
On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("{scan_filepath}") Then
    fso.DeleteFile("{scan_filepath}")
End If

Set objScanner = CreateObject("WIA.DeviceManager")
If Err.Number <> 0 Then
    WScript.Echo "Error 1: Cannot create WIA.DeviceManager - " & Err.Description
    WScript.Quit 1
End If

' 检查设备数量
deviceCount = objScanner.DeviceInfos.Count
If deviceCount = 0 Then
    WScript.Echo "Error 2: No device found"
    WScript.Quit 2
End If

WScript.Echo "Found " & deviceCount & " device(s)"

' 尝试找到可用设备（不检查类型）
Set objDevice = Nothing
For i = 1 To deviceCount
    Err.Clear
    Set objDeviceInfo = objScanner.DeviceInfos(i)
    WScript.Echo "Device " & i & ": " & objDeviceInfo.Name
    
    Set objDevice = objDeviceInfo.Connect()
    If Err.Number = 0 Then
        WScript.Echo "Successfully connected to device: " & objDeviceInfo.Name
        Exit For
    Else
        WScript.Echo "Failed to connect to device " & i & ": " & Err.Description
    End If
    Set objDevice = Nothing
Next

If objDevice Is Nothing Then
    WScript.Echo "Error 3: Cannot connect to any device"
    WScript.Quit 3
End If

' 获取扫描项
Err.Clear
Set objItem = objDevice.Items(1)
If Err.Number <> 0 Then
    WScript.Echo "Error 4: Cannot get device item: " & Err.Description
    WScript.Quit 4
End If

' 执行扫描并保存
Err.Clear
Set objImage = objItem.Transfer()
If Err.Number <> 0 Then
    WScript.Echo "Error 5: Scan transfer failed: " & Err.Description
    WScript.Quit 5
End If

Err.Clear
objImage.SaveFile "{scan_filepath}"
If Err.Number <> 0 Then
    WScript.Echo "Error 6: Cannot save file: " & Err.Description
    Set objImage = Nothing
    Set objItem = Nothing
    Set objDevice = Nothing
    Set objDeviceInfo = Nothing
    Set objScanner = Nothing
    Set fso = Nothing
    WScript.Quit 6
Else
    WScript.Echo "Scan success: {scan_filepath}"
End If

' 立即清理COM对象引用，不与其他用户冲突
Set objImage = Nothing
Set objItem = Nothing
Set objDeviceInfo = Nothing
Set objDevice = Nothing
Set objScanner = Nothing
Set fso = Nothing

' 扫描完成，返回成功状态
WScript.Quit 0
"""
                
                # 写入临时VBS文件
                vbs_path = os.path.join(path_manager.app_dir, 'temp_scan.vbs')
                # 在Windows上，cscript更喜欢由于ANSI/GBK编码的脚本
                try:
                    with open(vbs_path, 'w', encoding='gbk', errors='ignore') as f:
                        f.write(vbs_script)
                except:
                    with open(vbs_path, 'w', encoding='utf-8') as f:
                        f.write(vbs_script)
                
                print(f"执行VBS扫描脚本...")
                result = subprocess.run(['cscript.exe', '//Nologo', vbs_path],
                                      capture_output=True,
                                      creationflags=subprocess.CREATE_NO_WINDOW,
                                      timeout=60)
                
                # 获取输出信息
                def decode_output(output_bytes):
                    if not output_bytes: return ""
                    for enc in ['gbk', 'utf-8', 'utf-16']:
                        try:
                            return output_bytes.decode(enc).strip()
                        except:
                            continue
                    return output_bytes.decode('utf-8', errors='ignore').strip()

                stdout_str = decode_output(result.stdout)
                stderr_str = decode_output(result.stderr)
                
                if stdout_str:
                    print(f"VBS输出: {stdout_str}")
                if stderr_str:
                    print(f"VBS错误: {stderr_str}")
                
                print(f"VBS返回码: {result.returncode}")
                
                # 删除临时VBS文件
                try:
                    os.remove(vbs_path)
                except:
                    pass
                
                # 检查扫描文件是否生成
                time.sleep(1)
                if os.path.exists(scan_filepath) and os.path.getsize(scan_filepath) > 0:
                    file_size = os.path.getsize(scan_filepath)
                    print(f"VBS扫描成功: {scan_filename} ({file_size} 字节)")
                    restore_auto_open_folder()
                    return True, f"扫描成功！文件已保存到: {scan_filename}"
                else:
                    print("扫描未生成有效文件")
                    
            except subprocess.TimeoutExpired:
                print("扫描超时")
            except Exception as e:
                print(f"扫描执行异常: {e}")
            
            # 扫描失败
            restore_auto_open_folder()
            return False, "扫描执行失败。请确保扫描仪已连接。"
        
        except Exception as e:
            print(f"扫描异常: {e}")
            restore_auto_open_folder()
            return False, f"扫描异常: {str(e)}"
    
    except Exception as e:
        print(f"扫描功能异常: {e}")
        return False, f"扫描异常: {str(e)}"


# 网络多功能打印机发现已禁用 - 使用更稳定的扫描方案

 
# 获取本机局域网IP
def get_local_ip():
    """获取本机IP地址 - 支持内网穿透"""
    try:
        # 尝试连接外部服务器获取本机IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(2)  # 减少超时时间，快速失败
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        # 如果外网不通，尝试获取本地网络接口IP
        try:
            # 方案2：通过获取本机hostname对应的IP
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
            if ip and ip != '127.0.0.1':
                return ip
        except Exception:
            pass
        
        """获取本机局域网IP，优先返回非回环地址"""
        import socket
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(('8.8.8.8', 80))
            ip = s.getsockname()[0]
            s.close()
            if ip and not ip.startswith('127.') and not ip.startswith('169.254.'):
                return ip
        except Exception:
            pass
        return '127.0.0.1'

def get_external_ip():
    """获取公网IP地址（用于内网穿透检测）"""
    import urllib.request
    services = [
        'https://ipv4.icanhazip.com',
        'https://api.ipify.org',
        'https://checkip.amazonaws.com',
        'https://ipinfo.io/ip'
    ]
    for service in services:
        try:
            with urllib.request.urlopen(service, timeout=3) as response:
                external_ip = response.read().decode('utf-8').strip()
                if external_ip and '.' in external_ip and not external_ip.startswith(('192.168.', '10.', '172.')):
                    return external_ip
        except Exception:
            continue
    return None

def detect_network_mode():
    """检测网络模式：内网/公网/内网穿透"""
    local_ip = get_local_ip()
    external_ip = get_external_ip()
    is_private = (local_ip.startswith(('192.168.', '10.', '172.')) or local_ip == '127.0.0.1')
    if not is_private:
        return "public"
    elif local_ip == '127.0.0.1':
        return "private"
    else:
        return "private"

def get_current_ip_config():
    """获取当前IP配置状态 - 极简版本"""
    try:
        current_ip = get_local_ip()
        if current_ip and current_ip != '127.0.0.1':
            global IP_CONFIG_STATE
            dhcp_enabled = not IP_CONFIG_STATE.get('is_static', False)
            return {
                'index': '1',
                'description': '网络适配器',
                'ip': current_ip,
                'subnet': '255.255.255.0',
                'gateway': '',
                'dhcp_enabled': dhcp_enabled
            }
        else:
            return {}
    except Exception as e:
        print(f"获取IP配置失败: {e}")
        return {}

def set_static_ip(ip_address, subnet_mask='255.255.255.0', gateway=''):
    """设置静态IP地址 - 简易版本"""
    try:
        if not gateway:
            ip_parts = ip_address.split('.')
            if len(ip_parts) == 4:
                gateway = f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.1"
        print(f"设置静态IP: {ip_address}, 网关: {gateway}")
        adapter_names = ['以太网', 'Ethernet', '本地连接', 'WLAN', 'Wi-Fi']
        import subprocess
        for name in adapter_names:
            try:
                cmd = f'netsh interface ip set address name="{name}" static {ip_address} {subnet_mask} {gateway}'
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='gbk')
                if result.returncode == 0:
                    print(f"设置成功: {name}")
                    global IP_CONFIG_STATE
                    IP_CONFIG_STATE['is_static'] = True
                    IP_CONFIG_STATE['last_set_ip'] = ip_address
                    import time
                    time.sleep(2)
                    return True, f"静态IP设置成功"
            except Exception:
                continue
        return False, "未找到可用的网络适配器或设置失败"
    except Exception as e:
        return False, f"设置失败: {str(e)}"


def set_dhcp():
    """启用DHCP动态获取IP - 简易版本"""
    try:
        print("启用DHCP...")
        adapter_names = ['以太网', 'Ethernet', '本地连接', 'WLAN', 'Wi-Fi']
        import subprocess
        for name in adapter_names:
            try:
                cmd = f'netsh interface ip set address name="{name}" dhcp'
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='gbk')
                if result.returncode == 0:
                    print(f"DHCP设置成功: {name}")
                    global IP_CONFIG_STATE
                    IP_CONFIG_STATE['is_static'] = False
                    IP_CONFIG_STATE['last_set_ip'] = None
                    import time
                    time.sleep(3)
                    return True, f"已启用DHCP动态获取IP"
            except Exception:
                continue
        return False, "未找到可用的网络适配器或启用DHCP失败"
    except Exception as e:
        return False, f"启用DHCP失败: {str(e)}"

def suggest_static_ip():
    """建议一个可用的静态IP地址 - 简易版本"""
    current_ip = get_local_ip()
    if current_ip and current_ip != '127.0.0.1':
        ip_parts = current_ip.split('.')
        if len(ip_parts) == 4:
            return f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.100"
    return "192.168.1.100"
 
def detect_remote_desktop():
    """检测是否在远程桌面环境中运行"""
    try:
        # 方法1: 检查SESSIONNAME环境变量
        session_name = os.environ.get('SESSIONNAME', '')
        if session_name.startswith('RDP-Tcp'):
            return True
        
        # 方法2: 检查CLIENTNAME环境变量
        client_name = os.environ.get('CLIENTNAME', '')
        if client_name and client_name != os.environ.get('COMPUTERNAME', ''):
            return True
        
        # 方法3: 检查TS_SESSION_ID环境变量
        ts_session = os.environ.get('TS_SESSION_ID', '')
        if ts_session and ts_session != '0':
            return True
            
        # 方法4: 通过Windows API检查
        try:
            import ctypes
            from ctypes import wintypes
            
            # GetSystemMetrics(SM_REMOTESESSION)
            SM_REMOTESESSION = 0x1000
            user32 = ctypes.windll.user32
            is_remote = user32.GetSystemMetrics(SM_REMOTESESSION)
            if is_remote:
                return True
        except Exception:
            pass
            
        return False
    except Exception as e:
        print(f"检测远程桌面环境失败: {e}")
        return False

def get_print_queue_jobs(printer_name=None):
    """获取指定打印机的打印队列任务"""
    try:
        import win32print
        jobs = []
        
        if printer_name:
            # 获取指定打印机的任务
            try:
                printer_handle = win32print.OpenPrinter(printer_name)
                job_list = win32print.EnumJobs(printer_handle, 0, -1, 1)
                for job in job_list:
                    jobs.append({
                        'job_id': job['JobId'],
                        'printer': printer_name,
                        'document': job['pDocument'],
                        'user': job['pUserName'],
                        'status': job['Status'],
                        'pages': job['PagesPrinted'],
                        'size': job['Size']
                    })
                win32print.ClosePrinter(printer_handle)
            except Exception as e:
                print(f"获取打印机 {printer_name} 队列失败: {e}")
        else:
            # 获取所有打印机的任务
            for printer in PRINTERS:
                try:
                    printer_handle = win32print.OpenPrinter(printer)
                    job_list = win32print.EnumJobs(printer_handle, 0, -1, 1)
                    for job in job_list:
                        jobs.append({
                            'job_id': job['JobId'],
                            'printer': printer,
                            'document': job['pDocument'],
                            'user': job['pUserName'],
                            'status': job['Status'],
                            'pages': job['PagesPrinted'],
                            'size': job['Size']
                        })
                    win32print.ClosePrinter(printer_handle)
                except Exception as e:
                    print(f"获取打印机 {printer} 队列失败: {e}")
        
        return jobs
    except Exception as e:
        print(f"获取打印队列失败: {e}")
        return []

# Windows打印任务状态常量
JOB_STATUS_QUEUED = 0x0000
JOB_STATUS_PAUSED = 0x0001
JOB_STATUS_ERROR = 0x0002
JOB_STATUS_DELETING = 0x0004
JOB_STATUS_SPOOLING = 0x0008
JOB_STATUS_PRINTING = 0x0010
JOB_STATUS_OFFLINE = 0x0020
JOB_STATUS_PAPEROUT = 0x0040
JOB_STATUS_PRINTED = 0x0080
JOB_STATUS_DELETED = 0x0100
JOB_STATUS_BLOCKED_DEVQ = 0x0200
JOB_STATUS_USER_INTERVENTION = 0x0400
JOB_STATUS_RESTART = 0x0800
JOB_STATUS_COMPLETE = 0x1000

def get_job_status_description(status):
    """获取打印任务状态描述"""
    status_descriptions = []
    if status & JOB_STATUS_QUEUED:
        status_descriptions.append("排队中")
    if status & JOB_STATUS_PAUSED:
        status_descriptions.append("已暂停")
    if status & JOB_STATUS_ERROR:
        status_descriptions.append("错误")
    if status & JOB_STATUS_DELETING:
        status_descriptions.append("删除中")
    if status & JOB_STATUS_SPOOLING:
        status_descriptions.append("后台处理中")
    if status & JOB_STATUS_PRINTING:
        status_descriptions.append("正在打印")
    if status & JOB_STATUS_OFFLINE:
        status_descriptions.append("离线")
    if status & JOB_STATUS_PAPEROUT:
        status_descriptions.append("缺纸")
    if status & JOB_STATUS_PRINTED:
        status_descriptions.append("已打印")
    if status & JOB_STATUS_COMPLETE:
        status_descriptions.append("已完成")
    
    return ", ".join(status_descriptions) if status_descriptions else "未知状态"

def is_job_actively_printing(status):
    """检查任务是否正在打印"""
    return bool(status & (JOB_STATUS_PRINTING | JOB_STATUS_SPOOLING))

def is_job_cancellable(status):
    """检查任务是否可以取消"""
    # 不可取消的状态：已完成、已打印、正在删除
    uncancellable = (JOB_STATUS_PRINTED | JOB_STATUS_COMPLETE | JOB_STATUS_DELETED | JOB_STATUS_DELETING)
    return not bool(status & uncancellable)

def cancel_print_jobs_by_document(document_name, printer_name=None, cancel_active=False):
    """根据文档名取消打印任务
    
    Args:
        document_name: 文档名
        printer_name: 指定打印机名称，为None则搜索所有打印机
        cancel_active: 是否取消正在打印的任务（默认False）
    """
    global DEVICE_STATUS
    
    try:
        import win32print
        cancelled_jobs = []
        skipped_jobs = []
        has_cancelled_any = False
        
        # 获取打印队列任务
        jobs = get_print_queue_jobs(printer_name)
        
        for job in jobs:
            # 检查文档名是否匹配（支持部分匹配）
            if document_name.lower() in job['document'].lower() or \
               os.path.splitext(document_name)[0].lower() in job['document'].lower():
                
                job_status = job['status']
                status_desc = get_job_status_description(job_status)
                is_printing = is_job_actively_printing(job_status)
                is_cancellable = is_job_cancellable(job_status)
                
                print(f" 找到相关任务: {job['document']} (状态: {status_desc})")
                
                # 检查任务是否可取消
                if not is_cancellable:
                    print(f"[SKIP] 跳过任务 {job['document']}: 任务已完成或正在删除")
                    skipped_jobs.append({
                        'job_id': job['job_id'],
                        'printer': job['printer'],
                        'document': job['document'],
                        'reason': '任务已完成或正在删除',
                        'status': status_desc
                    })
                    continue
                
                # 检查是否正在打印
                if is_printing and not cancel_active:
                    print(f"[SKIP] 跳过正在打印的任务: {job['document']} (状态: {status_desc})")
                    print(f"    提示: 如需强制取消正在打印的任务，请使用带参数的API")
                    skipped_jobs.append({
                        'job_id': job['job_id'],
                        'printer': job['printer'],
                        'document': job['document'],
                        'reason': '正在打印，需要显式授权才能取消',
                        'status': status_desc
                    })
                    continue
                
                # 尝试取消任务
                try:
                    # 打开打印机句柄
                    printer_handle = win32print.OpenPrinter(job['printer'])
                    
                    # 取消打印任务
                    win32print.SetJob(printer_handle, job['job_id'], 0, None, win32print.JOB_CONTROL_CANCEL)
                    
                    cancelled_jobs.append({
                        'job_id': job['job_id'],
                        'printer': job['printer'],
                        'document': job['document'],
                        'status': status_desc,
                        'was_printing': is_printing
                    })
                    
                    action = "已强制取消" if is_printing else "已取消"
                    print(f" {action}打印任务: {job['document']} (任务ID: {job['job_id']}, 状态: {status_desc})")
                    
                    win32print.ClosePrinter(printer_handle)
                    has_cancelled_any = True
                    
                except Exception as e:
                    print(f" 取消打印任务失败: {job['document']} - {e}")
                    skipped_jobs.append({
                        'job_id': job['job_id'],
                        'printer': job['printer'],
                        'document': job['document'],
                        'reason': f'取消失败: {e}',
                        'status': status_desc
                    })
        
        # 如果成功取消了任何打印任务，重置全局打印状态并释放WIA设备
        if has_cancelled_any and DEVICE_STATUS['is_printing']:
            print("[RESET] 重置打印设备状态（已取消打印任务）")
            DEVICE_STATUS['is_printing'] = False
            DEVICE_STATUS['print_start_time'] = None
            DEVICE_STATUS['print_client'] = ''
            
            # 强制清理端口占用和重启WIA服务
            print("[CLEANUP] 强制清理后台占用资源...")
            try:
                port = getattr(app, 'current_port', 5000)
                cleanup_port_and_restart_wia(port)
            except Exception as e:
                print(f"[WARN] 端口清理异常（非致命）: {e}")
            
            # 自动释放WIA设备，避免与扫描冲突
            print("[INFO] 释放WIA扫描设备以避免冲突...")
            try:
                force_release_wia_device()
            except Exception as e:
                print(f"[WARN] WIA设备释放异常（非致命）: {e}")
        
        return {
            'cancelled': cancelled_jobs,
            'skipped': skipped_jobs,
            'total_found': len(cancelled_jobs) + len(skipped_jobs)
        }
        
    except Exception as e:
        print(f"取消打印任务失败: {e}")
        return {
            'cancelled': [],
            'skipped': [],
            'total_found': 0,
            'error': str(e)
        }

def clear_all_print_queues():
    """清空所有打印机的打印队列"""
    global DEVICE_STATUS
    
    try:
        import win32print
        cleared_count = 0
        
        for printer in PRINTERS:
            try:
                printer_handle = win32print.OpenPrinter(printer)
                
                # 获取所有任务
                job_list = win32print.EnumJobs(printer_handle, 0, -1, 1)
                
                # 取消所有任务
                for job in job_list:
                    try:
                        win32print.SetJob(printer_handle, job['JobId'], 0, None, win32print.JOB_CONTROL_CANCEL)
                        cleared_count += 1
                        print(f" 已取消: {printer} - {job['pDocument']} (任务ID: {job['JobId']})")
                    except Exception as e:
                        print(f" 取消任务失败: {job['pDocument']} - {e}")
                
                win32print.ClosePrinter(printer_handle)
                
            except Exception as e:
                print(f"清理打印机 {printer} 队列失败: {e}")
        
        # 如果成功清空了任何打印任务，重置全局打印状态并释放WIA设备
        if cleared_count > 0 and DEVICE_STATUS['is_printing']:
            print("[RESET] 重置打印设备状态（已清空所有打印队列）")
            DEVICE_STATUS['is_printing'] = False
            DEVICE_STATUS['print_start_time'] = None
            DEVICE_STATUS['print_client'] = ''
            
            # 强制清理端口占用和重启WIA服务
            print("[CLEANUP] 强制清理后台占用资源...")
            try:
                port = getattr(app, 'current_port', 5000)
                cleanup_port_and_restart_wia(port)
            except Exception as e:
                print(f"[WARN] 端口清理异常（非致命）: {e}")
            
            # 自动释放WIA设备，避免与扫描冲突
            print("[INFO] 释放WIA扫描设备以避免冲突...")
            try:
                force_release_wia_device()
            except Exception as e:
                print(f"[WARN] WIA设备释放异常（非致命）: {e}")
        
        return cleared_count
        
    except Exception as e:
        print(f"清空打印队列失败: {e}")
        return 0

# 开机自启注册表操作
def set_autostart(enable=True):
    exe_path = sys.executable
    key = r'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    name = 'PrintServerApp'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_ALL_ACCESS) as regkey:
        if enable:
            winreg.SetValueEx(regkey, name, 0, winreg.REG_SZ, exe_path)
        else:
            try:
                winreg.DeleteValue(regkey, name)
            except FileNotFoundError:
                pass
 
def get_autostart():
    key = r'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    name = 'PrintServerApp'
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ) as regkey:
            val, _ = winreg.QueryValueEx(regkey, name)
            return bool(val)
    except FileNotFoundError:
        return False
 
app = Flask(__name__, static_folder=None)
app.secret_key = 'print_server_secret_key'

# 配置静态文件路由 - 用于提供本地 Bootstrap 文件
@app.route('/static/<path:filename>')
def serve_static(filename):
    """提供静态文件服务 - Bootstrap CSS 和 JS"""
    import os
    from flask import send_file, abort
    
    # 安全文件名检查 - 防止路径遍历攻击
    if '..' in filename or filename.startswith('/'):
        abort(403)
    
    # 定义允许的静态文件类型
    allowed_files = {
        'bootstrap.min.css',
        'bootstrap.bundle.min.js'
    }
    
    if filename not in allowed_files:
        abort(403)
    
    # 获取文件的完整路径 - 使用路径管理器支持打包后的路径
    file_path = path_manager.get_resource_path(filename)
    
    # 验证文件存在
    if not os.path.exists(file_path):
        print(f"静态文件未找到: {file_path}")
        # 尝试备用路径 - 当前工作目录
        file_path = os.path.join(os.getcwd(), filename)
        if not os.path.exists(file_path):
            print(f"静态文件备用路径也未找到: {file_path}")
            abort(404)
    
    try:
        # 发送文件 - 带有适当的 MIME 类型
        response = None
        if filename.endswith('.css'):
            response = send_file(file_path, mimetype='text/css')
        elif filename.endswith('.js'):
            response = send_file(file_path, mimetype='application/javascript')
        else:
            response = send_file(file_path)
        
        # 添加缓存头
        if response:
            response.cache_control.max_age = 3600
        return response
    except Exception as e:
        print(f"静态文件服务错误 ({filename}): {e}")
        abort(500)

# Flask统一配置 - 所有Windows版本通用
def get_flask_config():
    """统一的Flask配置 - 所有Windows版本都使用相同配置"""
    return {
        'MAX_CONTENT_LENGTH': 100 * 1024 * 1024,      # 100MB 最大上传文件大小
        'PERMANENT_SESSION_LIFETIME': 3600,            # 1小时会话时长
        'TEMPLATES_AUTO_RELOAD': False,                # 禁用模板自动重载（生产环境）
        'JSON_AS_ASCII': False,                        # 允许 JSON 中的非 ASCII 字符
        'JSONIFY_PRETTYPRINT_REGULAR': False,          # JSON 响应不美化（减少体积）
        'SEND_FILE_MAX_AGE_DEFAULT': 300,              # 5分钟缓存静态文件
    }

app.config.update(get_flask_config())

# 兼容PyInstaller打包的路径处理
# 使用路径管理器配置文件夹和文件路径
UPLOAD_FOLDER = path_manager.get_upload_dir()
LOG_FILE = path_manager.get_log_path()
path_manager.ensure_data_dirs()

# 添加请求后钩子，进行资源清理
@app.after_request
def after_request(response):
    """请求处理后的清理工作 - 优化内存管理"""
    try:
        # 优化请求计数管理
        if not hasattr(app, 'request_count'):
            app.request_count = 0
        app.request_count += 1
        
        # 降低垃圾回收频率，减少性能影响
        if app.request_count % 500 == 0:  # 从100改为500
            import gc
            collected = gc.collect()
            if collected > 0:
                print(f"内存清理: 回收了 {collected} 个对象")
            
            # 重置计数器防止溢出
            if app.request_count > 10000:
                app.request_count = 0
        
        # 优化响应头设置
        response.headers.update({
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0',
            'X-Content-Type-Options': 'nosniff',
            'X-Frame-Options': 'DENY'
        })
        
        # 只在需要时关闭连接
        if request.method in ['POST', 'PUT', 'DELETE']:
            response.headers['Connection'] = 'close'
        
    except Exception as e:
        print(f"️ 请求后清理异常: {e}")
        
    return response

# 添加错误处理器
@app.errorhandler(500)
def internal_error(error):
    """内部服务器错误处理"""
    print(f"️ 内部服务器错误: {error}")
    return jsonify({
        'error': '服务器内部错误',
        'message': '请稍后重试，如果问题持续请重启服务'
    }), 500

@app.errorhandler(413)
def too_large(error):
    """请求实体过大错误处理"""
    return jsonify({
        'error': '文件过大',
        'message': '上传文件大小不能超过100MB'
    }), 413



# 健康检查端点，便于内网穿透监控
@app.route('/health', methods=['GET'])
def health():
    try:
        status = {
            'status': 'ok',
            'time': int(time.time()),
            'service_running': getattr(service_manager, 'service_running', False),
            'uptime': int(time.time() - service_manager.start_time) if getattr(service_manager, 'start_time', None) else 0
        }
        resp = jsonify(status)
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp
    except Exception as e:
        return jsonify({'status': 'error', 'error': str(e)}), 500



# API: 列出扫描仪（JSON），便于远程调用与内网穿透场景
@app.route('/api/scanners', methods=['GET', 'OPTIONS'])
def api_list_scanners():
    """获取可用扫描仪列表 - 支持内网穿透跨域请求"""
    # 处理 CORS 预检请求
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'GET, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Access-Control-Max-Age'] = '3600'
        return resp
    
    try:
        scanners = get_available_scanners()
        out = [{
            'name': str(s.get('name')),
            'id': str(s.get('id')),
            'type': str(s.get('type')),
            'available': bool(s.get('available', False))
        } for s in scanners]
        resp = jsonify({
            'status': 'success',
            'count': len(out),
            'scanners': out,
            'timestamp': int(time.time())
        })
        # 内网穿透支持：允许所有跨域请求
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'GET, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        return resp
    except Exception as e:
        return jsonify({
            'status': 'error',
            'error': '扫描仪列表获取失败',
            'detail': str(e)
        }), 500


# API: 强制释放被锁定的扫描仪设备
@app.route('/api/release_scanner', methods=['POST', 'OPTIONS'])
def api_release_scanner():
    """强制释放被锁定的扫描仪设备"""
    global DEVICE_STATUS
    
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp
    
    try:
        print("执行强制释放扫描仪设备...")
        
        # 强制清理端口占用和重启WIA服务
        print("[CLEANUP] 强制清理后台占用资源...")
        try:
            port = getattr(app, 'current_port', 5000)
            cleanup_port_and_restart_wia(port)
        except Exception as e:
            print(f"[WARN] 端口清理异常（非致命）: {e}")
        
        success = force_release_wia_device()
        
        # 同时重置扫描状态
        if DEVICE_STATUS['is_scanning']:
            print("重置扫描设备状态")
            DEVICE_STATUS['is_scanning'] = False
            DEVICE_STATUS['scan_start_time'] = None
            DEVICE_STATUS['scan_client'] = ''
        
        if success:
            print("[SUCCESS] 扫描仪设备已成功释放")
            return jsonify({
                'status': 'success',
                'message': '扫描仪设备已成功释放，请稍候后重试扫描',
                'timestamp': int(time.time())
            })
        else:
            return jsonify({
                'status': 'error',
                'error': '释放失败',
                'message': '无法释放扫描仪设备，请检查权限或手动重启设备'
            }), 500
    except Exception as e:
        print(f"释放扫描仪失败: {e}")
        return jsonify({
            'status': 'error',
            'error': '操作异常',
            'detail': str(e)
        }), 500


# API: 触发扫描（返回扫描结果信息或错误）- 支持内网穿透
@app.route('/api/scan', methods=['POST', 'OPTIONS'])
def api_trigger_scan():
    """触发扫描操作 - 支持内网穿透跨域请求
    
    POST JSON参数：
    {
        "scanner_id": "扫描仪ID（可选，默认为'default'）",
        "scanner_name": "扫描仪名称（用于显示）",
        "format": "扫描格式，如PNG、JPG（可选，默认为PNG）"
    }
    
    返回示例 (成功)：
    {
        "status": "success",
        "success": true,
        "message": "扫描成功！文件已保存到: scan_20260107_120000.PNG",
        "scan_time": 15,
        "client_ip": "192.168.1.100"
    }
    
    返回示例 (警告)：
    {
        "status": "warning",
        "success": true,
        "message": "扫描窗口已打开。请注意：...",
        "scan_time": 30,
        "client_ip": "192.168.1.100"
    }
    """
    # 处理 CORS 预检请求（内网穿透必需）
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Access-Control-Max-Age'] = '3600'
        return resp
    
    try:
        data = request.get_json() or {}
        scanner_id = data.get('scanner_id', 'default')
        scanner_name = data.get('scanner_name', '通用扫描')
        scan_format = data.get('format', 'PNG').upper()
        
        # 获取客户端IP（支持内网穿透的代理转发）
        client_ip = request.headers.get('X-Forwarded-For', '').split(',')[0].strip() or request.remote_addr
        
        # 验证扫描格式
        allowed_formats = ('PNG', 'JPG', 'JPEG', 'BMP', 'TIFF')
        if scan_format not in allowed_formats:
            return jsonify({
                'status': 'error',
                'error': '不支持的扫描格式',
                'detail': f'支持的格式: {", ".join(allowed_formats)}',
                'client_ip': client_ip
            }), 400
        
        # 检查扫描器是否正在忙碌
        if DEVICE_STATUS.get('is_scanning'):
            return jsonify({
                'status': 'error',
                'error': '扫描器正在忙碌中',
                'detail': '请稍后再试',
                'client_ip': client_ip
            }), 409
        
        # 标记扫描开始
        DEVICE_STATUS['is_scanning'] = True
        DEVICE_STATUS['scan_start_time'] = time.time()
        DEVICE_STATUS['scan_client'] = client_ip
        
        try:
            # 执行扫描
            ok, message = start_scan_silent(scanner_id, scanner_name, scan_format)
            
            # 计算扫描耗时
            scan_time = int(time.time() - DEVICE_STATUS['scan_start_time'])
            
            # 返回结果
            resp = jsonify({
                'status': 'success' if ok else 'warning',
                'success': ok,
                'message': message,
                'scan_time': scan_time,
                'client_ip': client_ip,
                'timestamp': int(time.time())
            })
            # 内网穿透支持：允许所有跨域请求
            resp.headers['Access-Control-Allow-Origin'] = '*'
            resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
            resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
            resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            
            # HTTP状态码：200成功 206部分成功
            return resp, (200 if ok else 206)
        
        finally:
            # 标记扫描结束
            DEVICE_STATUS['is_scanning'] = False
            DEVICE_STATUS['scan_start_time'] = None
            DEVICE_STATUS['scan_client'] = ''
    
    except Exception as e:
        error_msg = str(e)
        
        # 检测是否是WIA设备被占用的错误
        if 'WIA' in error_msg or 'busy' in error_msg.lower() or 'occupy' in error_msg.lower():
            print(f"[ERROR] 检测到WIA设备被占用错误，尝试自动释放: {error_msg}")
            try:
                force_release_wia_device()
                error_msg += "\n(已尝试自动释放设备，请重试)"
            except Exception as release_error:
                print(f"[WARN] 自动释放WIA设备失败: {release_error}")
        
        return jsonify({
            'status': 'error',
            'error': '触发扫描失败',
            'detail': error_msg,
            'client_ip': request.headers.get('X-Forwarded-For', '').split(',')[0].strip() or request.remote_addr
        }), 500

# 虚拟打印机名称列表（这些不是真正的物理打印机）
VIRTUAL_PRINTERS = {
    '导出为WPS PDF', 'WPS PDF', 'Microsoft Print to PDF', 'Microsoft XPS Document Writer',
    'Fax', '传真', 'OneNote', 'OneNote (Desktop)', 'Send To OneNote 2016',
    'Adobe PDF', 'Foxit Reader PDF Printer', 'PDF Creator', 'CutePDF Writer',
    'novaPDF', 'PDFCreator', 'Bullzip PDF Printer', 'doPDF', 'PDF24',
    'Virtual PDF Printer', '虚拟PDF打印机', 'Send to Kindle', '发送到WPS高级打印'
}

# 打印机缓存管理
class PrinterCache:
    def __init__(self):
        self.cache_time = 0
        self.all_printers = []
        self.physical_printers = []
        self.default_printer = None
        self.cache_timeout = self._detect_cache_timeout()

    def _detect_cache_timeout(self):
        """根据Windows版本设置缓存超时"""
        import platform, sys
        try:
            windows_version = platform.release()
            windows_build = getattr(getattr(sys, 'getwindowsversion', lambda: None)(), 'build', None)
            if windows_version == "7":
                print(" Win7打印机缓存：3分钟")
                return 180
            elif windows_build and windows_build >= 22000:
                print(" Win11打印机缓存：10分钟")
                return 600
            else:
                print(" Win10打印机缓存：5分钟")
                return 300
        except Exception as e:
            print(f"打印机缓存配置失败，使用默认5分钟: {e}")
            return 300

    def is_cache_valid(self):
        import time
        return (time.time() - self.cache_time) < self.cache_timeout

    def refresh_cache(self):
        import win32print, time
        try:
            self.all_printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            # 调试模式：如果启用，显示所有打印机；否则只显示物理打印机
            if DEBUG_MODE:
                self.physical_printers = self.all_printers  # 调试模式：显示所有打印机
                print(f"[调试模式] 显示所有打印机，包括虚拟打印机: {self.all_printers}")
            else:
                self.physical_printers = [p for p in self.all_printers if p not in VIRTUAL_PRINTERS]
            try:
                default = win32print.GetDefaultPrinter()
                self.default_printer = default if default in self.physical_printers else (self.physical_printers[0] if self.physical_printers else None)
            except:
                self.default_printer = self.physical_printers[0] if self.physical_printers else None
            self.cache_time = time.time()
            return True
        except Exception as e:
            print(f"刷新打印机缓存失败: {e}")
            return False

    def get_printers(self):
        if not self.is_cache_valid():
            self.refresh_cache()
        return self.physical_printers

    def get_default_printer(self):
        if not self.is_cache_valid():
            self.refresh_cache()
        return self.default_printer


# 创建全局打印机缓存并初始化
printer_cache = PrinterCache()
printer_cache.refresh_cache()
ALL_PRINTERS = printer_cache.all_printers
PRINTERS = printer_cache.physical_printers

def get_default_printer():
    """获取系统默认打印机 - 使用缓存优化"""
    return printer_cache.get_default_printer()

def refresh_printer_list():
    """刷新打印机列表 - 使用缓存优化"""
    global ALL_PRINTERS, PRINTERS
    try:
        success = printer_cache.refresh_cache()
        if success:
            ALL_PRINTERS = printer_cache.all_printers
            PRINTERS = printer_cache.physical_printers
            print(f"打印机列表已刷新，检测到 {len(PRINTERS)} 台物理打印机")
        return success
    except Exception as e:
        print(f"刷新打印机列表失败: {e}")
        return False

HTML = '''
{% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}
        <div class="container mt-3">
            {% for category, msg in messages %}
                <div class="alert alert-{{category}} alert-dismissible fade show" role="alert">
                    {{msg}}
                    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                </div>
            {% endfor %}
        </div>
    {% endif %}
{% endwith %}
<!doctype html>
<html lang="zh-cn" spellcheck="false">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>内网打印及扫描服务</title>
    <!-- 使用本地 Bootstrap CSS -->
    <link href="/static/bootstrap.min.css" rel="stylesheet">
    <style>
        /* Bootstrap核心样式备用 */
        body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif; margin: 0; padding: 0; }
        .container { max-width: 800px; margin: 0 auto; padding: 20px; }
        
        /* 禁用所有拼写检查波浪线 */
        *, *::before, *::after {
            -webkit-text-decoration-skip: none;
            text-decoration-skip: none;
        }
        select, input, textarea {
            -webkit-text-decoration-line: none !important;
            text-decoration-line: none !important;
            -webkit-text-decoration: none !important;
            text-decoration: none !important;
        }
        .btn { display: inline-block; padding: 6px 12px; margin-bottom: 0; font-size: 14px; font-weight: 400; line-height: 1.42857143; text-align: center; white-space: nowrap; vertical-align: middle; cursor: pointer; border: 1px solid transparent; border-radius: 4px; text-decoration: none; }
        .btn-primary { color: #fff; background-color: #007bff; border-color: #007bff; }
        .btn-outline-secondary { color: #6c757d; border-color: #6c757d; background-color: transparent; }
        .btn-warning { color: #212529; background-color: #ffc107; border-color: #ffc107; }
        .form-control { display: block; width: 100%; padding: 6px 12px; font-size: 14px; line-height: 1.42857143; color: #555; background-color: #fff; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; outline: none; transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out; }
        .form-control:focus { border-color: #007bff; box-shadow: 0 0 0 0.2rem rgba(0, 123, 255, 0.25); }
        .form-select { background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3e%3cpath fill='none' stroke='%23343a40' stroke-linecap='round' stroke-linejoin='round' stroke-width='2' d='m1 6 7 7 7-7'/%3e%3c/svg%3e"); }
        .alert { padding: 15px; margin-bottom: 20px; border: 1px solid transparent; border-radius: 4px; }
        .alert-success { color: #155724; background-color: #d4edda; border-color: #c3e6cb; }
        .alert-danger { color: #721c24; background-color: #f8d7da; border-color: #f5c6cb; }
        .alert-warning { color: #856404; background-color: #fff3cd; border-color: #ffeaa7; }
        .alert-info { color: #0c5460; background-color: #d1ecf1; border-color: #bee5eb; }
        .table { width: 100%; margin-bottom: 20px; border-collapse: collapse; }
        .table th, .table td { padding: 8px; text-align: left; border-top: 1px solid #ddd; }
        .table-light th { background-color: #f8f9fa; }
        .nav { display: flex; flex-wrap: wrap; padding-left: 0; margin-bottom: 0; list-style: none; }
        .nav-pills .nav-link { border-radius: 0.25rem; }
        .nav-pills .nav-link.active { color: #fff; background-color: #007bff; }
        .nav-item { margin-right: 10px; }
        .nav-link { display: block; padding: 8px 16px; text-decoration: none; color: #007bff; cursor: pointer; border: 1px solid transparent; }
        .tab-content { margin-top: 20px; }
        .tab-pane { display: none; }
        .tab-pane.show.active { display: block; }
        .row { display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px; }
        .col-md-3, .col-md-4, .col-md-6, .col-md-8, .col-12 { position: relative; width: 100%; padding-right: 15px; padding-left: 15px; }
        @media (min-width: 768px) {
            .col-md-3 { flex: 0 0 25%; max-width: 25%; }
            .col-md-4 { flex: 0 0 33.333333%; max-width: 33.333333%; }
            .col-md-6 { flex: 0 0 50%; max-width: 50%; }
            .col-md-8 { flex: 0 0 66.666667%; max-width: 66.666667%; }
        }
        .card { position: relative; display: flex; flex-direction: column; background-color: #fff; border: 1px solid rgba(0,0,0,.125); border-radius: 0.25rem; }
        .card-header { padding: 0.75rem 1.25rem; background-color: rgba(0,0,0,.03); border-bottom: 1px solid rgba(0,0,0,.125); }
        .card-body { flex: 1 1 auto; padding: 1.25rem; }
        .form-text { margin-top: 0.25rem; font-size: 0.875em; color: #6c757d; }
        .badge { display: inline-block; padding: 0.25em 0.4em; font-size: 75%; font-weight: 700; line-height: 1; text-align: center; white-space: nowrap; vertical-align: baseline; border-radius: 0.25rem; }
        .bg-success { background-color: #28a745 !important; color: #fff; }
        .bg-primary { background-color: #007bff !important; color: #fff; }
        .list-group { display: flex; flex-direction: column; padding-left: 0; margin-bottom: 0; }
        .list-group-item { position: relative; display: block; padding: 0.75rem 1.25rem; background-color: #fff; border: 1px solid rgba(0,0,0,.125); }
        .g-3 > * { margin-bottom: 1rem; }
        .mb-3, .mb-4 { margin-bottom: 1.5rem; }
        .mt-4 { margin-top: 1.5rem; }
        .text-end { text-align: right; }
        .text-center { text-align: center; }
        .px-4 { padding-left: 1.5rem; padding-right: 1.5rem; }
        .justify-content-center { justify-content: center; }
        /* 原有的自定义样式 */
        body { background: #f8f9fa; }
        .container { max-width: 800px; margin-top: 40px; background: #fff; border-radius: 12px; box-shadow: 0 2px 12px #0001; padding: 32px; }
        h1 { font-size: 2rem; margin-bottom: 1.5rem; }
        .author-info { font-size: 1.2rem; color: #6c757d; font-weight: normal; }
        .form-label { font-weight: 500; }
        .print-form-grid {
            row-gap: 1rem;
        }
        .print-form-grid .form-control,
        .print-form-grid .form-select,
        .print-form-grid .btn {
            min-width: 0;
        }
        .print-form-grid .print-field,
        .print-form-grid .print-file-area {
            min-width: 180px;
        }
        .print-form-grid .print-file-area {
            width: 100%;
        }
        .print-form-grid .print-file-area .file-drop-area {
            width: 100%;
            min-width: 100%;
        }
        .print-form-grid .print-submit-row {
            min-width: 100%;
        }
        .copies-custom-wrap {
            display: none;
            margin-top: 0.5rem;
        }
        .table { background: #fff; }
        .log-list { max-height: 200px; overflow-y: auto; font-size: 0.95em; }
        .nav-pills .nav-link.active { background-color: #0d6efd; }
        .tab-content { margin-top: 20px; }
        .ip-status { padding: 15px; background: #f8f9fa; border-radius: 8px; margin-bottom: 20px; }
        .ip-status .badge { font-size: 0.9em; }
        
        /* 拖拽文件区域样式 */
        .file-drop-area {
            border: 2px dashed #ccc;
            border-radius: 8px;
            padding: 40px;
            text-align: center;
            background-color: #f9f9f9;
            transition: all 0.3s ease;
            margin-bottom: 20px;
            cursor: pointer;
        }
        
        .file-drop-area:hover {
            border-color: #007bff;
            background-color: #e3f2fd;
        }
        
        .file-drop-area.drag-over {
            border-color: #007bff;
            background-color: #e3f2fd;
            transform: scale(1.02);
        }
        
        .file-drop-area .drop-icon {
            font-size: 48px;
            color: #6c757d;
            margin-bottom: 15px;
        }
        
        .file-drop-area.drag-over .drop-icon {
            color: #007bff;
        }
        
        .file-list {
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            padding: 10px;
            margin-top: 15px;
            background-color: #fff;
        }
        
        .file-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 12px;
            border-bottom: 1px solid #eee;
            background-color: #f8f9fa;
            border-radius: 4px;
            margin-bottom: 5px;
        }
        
        .file-item:last-child {
            margin-bottom: 0;
        }
        
        .file-item .file-name {
            font-weight: 500;
            color: #495057;
            flex-grow: 1;
            margin-right: 10px;
        }
        
        .file-item .file-size {
            color: #6c757d;
            font-size: 0.9em;
            margin-right: 10px;
        }
        
        .file-item .remove-btn {
            background: none;
            border: none;
            color: #dc3545;
            cursor: pointer;
            font-size: 16px;
            padding: 2px 6px;
            border-radius: 3px;
        }
        
        .file-item .remove-btn:hover {
            background-color: #dc3545;
            color: white;
        }
        
        /* 队列表格样式 */
        .queue-table {
            background-color: #fff;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .queue-table .file-name {
            font-weight: 500;
            color: #495057;
        }
        
        .queue-table .btn-group-sm .btn {
            padding: 0.25rem 0.5rem;
            font-size: 0.875rem;
        }
        
        /* 删除按钮样式 */
        .btn-outline-danger:hover {
            transform: scale(1.05);
            transition: transform 0.2s ease-in-out;
        }
        
        /* 空队列提示样式 */
        .empty-queue {
            padding: 2rem;
            text-align: center;
            color: #6c757d;
            background-color: #f8f9fa;
            border-radius: 8px;
            border: 2px dashed #dee2e6;
        }
        
        /* 文件类型徽章样式 */
        .file-type-badge {
            display: inline-block;
            padding: 0.25em 0.5em;
            font-size: 0.75em;
            font-weight: 500;
            line-height: 1;
            text-align: center;
            white-space: nowrap;
            vertical-align: baseline;
            border-radius: 0.25rem;
            text-transform: uppercase;
        }
        
        .file-type-pdf { background-color: #dc3545; color: white; }
        .file-type-doc, .file-type-docx { background-color: #2b579a; color: white; }
        .file-type-xls, .file-type-xlsx { background-color: #217346; color: white; }
        .file-type-ppt, .file-type-pptx { background-color: #d24726; color: white; }
        .file-type-txt { background-color: #6f42c1; color: white; }
        .file-type-jpg, .file-type-jpeg, .file-type-png { background-color: #fd7e14; color: white; }
        .file-type-unknown { background-color: #6c757d; color: white; }
        
        /* 页面切换按钮样式 */
        .tab-switch-buttons {
            margin-bottom: 2rem;
        }
        
        .tab-switch-buttons .btn {
            padding: 0.75rem 2rem;
            font-weight: 500;
            border-radius: 25px;
            margin: 0 0.5rem;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .tab-switch-buttons .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        
        .tab-switch-buttons .btn.btn-primary {
            background: linear-gradient(45deg, #007bff, #0056b3);
            border: none;
        }
        
        .tab-switch-buttons .btn.btn-outline-primary:hover {
            background: linear-gradient(45deg, #007bff, #0056b3);
            border-color: transparent;
        }

        .inline-refresh-btn {
            border: none !important;
            background: transparent !important;
            color: #6c757d;
            padding: 0.1rem 0.35rem;
            margin-left: 0.35rem;
            box-shadow: none !important;
            line-height: 1.2;
            min-width: auto;
        }

        .inline-refresh-btn:hover,
        .inline-refresh-btn:focus,
        .inline-refresh-btn:active {
            border: none !important;
            background: rgba(13, 110, 253, 0.08) !important;
            color: #0d6efd;
            box-shadow: none !important;
        }

        .inline-refresh-btn:focus {
            outline: none;
        }

        .collapsible-help {
            border-radius: 8px;
            overflow: hidden;
        }

        .collapsible-help.collapsed {
            padding-top: 0.45rem;
            padding-bottom: 0.45rem;
        }

        .collapsible-help-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            cursor: pointer;
            user-select: none;
            padding: 0.9rem 1rem;
        }

        .collapsible-help.collapsed .collapsible-help-header {
            padding: 0.35rem 0.9rem;
        }

        .collapsible-help-header h6 {
            margin: 0;
        }

        .collapsible-help.collapsed .collapsible-help-header h6 {
            font-size: 0.98rem;
        }

        .collapsible-help-toggle {
            font-size: 0.9rem;
            color: #0c5460;
            white-space: nowrap;
        }

        .collapsible-help.collapsed .collapsible-help-toggle {
            font-size: 0.86rem;
        }

        .collapsible-help-body {
            display: block;
        }

        .collapsible-help.collapsed .collapsible-help-body {
            display: none;
        }

        .collapsible-help.collapsed .collapsible-help-toggle::after {
            content: '展开';
        }

        .collapsible-help:not(.collapsed) .collapsible-help-toggle::after {
            content: '收纳';
        }
        
        /* 标签页内容淡入效果 */
        .tab-content {
            animation: fadeIn 0.3s ease-in-out;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        /* 扫描区域特殊样式 */
        #scanTab .card {
            border-left: 4px solid #17a2b8;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: transform 0.2s ease;
        }
        
        #scanTab .card:hover {
            transform: translateY(-2px);
        }
    </style>
</head>
<body>
<div class="container">
    <h1 class="mb-4 text-center">内网打印及扫描服务v2.6</br>
    <span class="author-info"><a href="https://www.937788.xyz/" target="_blank" style="color: #6c757d; text-decoration: none;">（yckj666@52PJ，作者：忆痕）</a>
    </p><a href="https://github.com/a937750307/lan-printing" target="_blank" style="color: #007bff; text-decoration: none;">【Github】</a>
    <a href="https://zanzhu.937788.xyz" target="_blank" style="color: #28a745; text-decoration: none;">【赞助】</a>
    <a href="https://print.937788.xyz" target="_blank" style="color: #dc3545; text-decoration: none;">【版本更新】</a></span></h1>
    
    <!-- 功能切换导航 -->
    <div class="text-center mb-4">
        <div class="btn-group" role="group" aria-label="功能切换">
            <button type="button" class="btn btn-primary" id="printTabBtn" onclick="switchTab('print')">
                ️ 打印
            </button>
            <button type="button" class="btn btn-outline-primary" id="scanTabBtn" onclick="switchTab('scan')">
                 扫描
            </button>
        </div>
        <p class="text-muted mt-2">可右键托盘栏图标进行网络配置和其他设置</p>
    </div>

    <!-- 打印管理内容 -->
    <div class="tab-content" id="printTab">
        <div class="main-content">
            <form method="post" enctype="multipart/form-data" class="row g-3 mb-4" id="printForm">
                <input type="hidden" name="action" value="print">
                <input type="hidden" name="device_name" id="deviceNameField" value="">
                <div class="col-md-6 print-field">
                    <label class="form-label">选择打印机 
                            <button type="button" class="btn btn-sm inline-refresh-btn" onclick="refreshPrinterList()" title="刷新打印机列表">
                             刷新
                        </button>
                    </label>
                    <select name="printer" class="form-select" id="printerSelect" spellcheck="false">
                        {% if printers %}
                            {% for p in printers %}
                                <option value="{{p}}" {% if p == default_printer %}selected{% endif %}>{{p}}{% if p == default_printer %} (默认){% endif %}</option>
                            {% endfor %}
                        {% else %}
                            <option value="">未检测到可用打印机</option>
                        {% endif %}
                    </select>
                    <div class="form-text text-muted">
                        {% if printers %}
                            <strong>️ 重要提醒:</strong> 请仔细选择打印机！程序会严格按照您的选择发送到指定打印机，不会回退到默认打印机。
                            <br>已过滤虚拟打印机，自动选择默认打印机，可手动刷新列表
                        {% else %}
                            <span class="text-warning">️ 未检测到物理打印机，请检查打印机连接后点击刷新</span>
                        {% endif %}
                    </div>
                </div>
                <div class="col-md-6 print-field">
                    <label class="form-label">打印份数</label>
                    <select name="copies" class="form-select" id="copiesSelect" spellcheck="false">
                        <option value="1" selected>1</option>
                        <option value="2">2</option>
                        <option value="3">3</option>
                        <option value="4">4</option>
                        <option value="5">5</option>
                        <option value="6">6</option>
                        <option value="7">7</option>
                        <option value="8">8</option>
                        <option value="9">9</option>
                        <option value="10">10</option>
                        <option value="custom">自定义</option>
                    </select>
                    <div class="copies-custom-wrap" id="copiesCustomWrap">
                        <input type="number" class="form-control" id="copiesCustomInput" min="1" max="9999" step="1" value="1" spellcheck="false">
                        <div class="form-text text-muted">输入自定义份数，提交时会自动使用该值</div>
                    </div>
                </div>
                <div class="col-md-4 print-field">
                    <label class="form-label">单双面
                        {% if printer_caps and printer_caps.get('duplex_support') %}
                        <span class="badge bg-success ms-1">支持</span>
                        {% endif %}
                    </label>
                    <select name="duplex" class="form-select" id="duplexSelect" spellcheck="false">
                        <option value="1"> 单面打印</option>
                        {% if printer_caps and printer_caps.get('duplex_support') %}
                        <option value="2"> 长边翻转 (书本式)</option>
                        <option value="3"> 短边翻转 (翻页式)</option>
                        {% endif %}
                    </select>
                    {% if printer_caps %}
                        {% if printer_caps.get('duplex_support') %}
                        <div class="form-text text-success">
                            <small> 支持双面打印
                            {% if printer_caps.get('duplex_modes') %}
                            - {{ printer_caps.get('duplex_modes')|join(', ') }}
                            {% endif %}
                            </small>
                        </div>
                        {% else %}
                        <div class="form-text text-warning">
                            <small>️ 打印机不支持双面打印，将使用单面模式</small>
                        </div>
                        {% endif %}
                    {% endif %}
                </div>
                <div class="col-md-4 print-field">
                    <label class="form-label">纸张大小</label>
                    <select name="papersize" class="form-select" id="paperSelect" spellcheck="false">
                        {% if printer_caps and printer_caps.get('papers') %}
                            {% for p in printer_caps.papers %}
                            <option value="{{p.id}}" {% if p.id == 9 %}selected{% endif %}>{{p.name}}</option>
                            {% endfor %}
                        {% else %}
                            <option value="9" selected>A4</option>
                        {% endif %}
                    </select>
                </div>
                <div class="col-md-4 print-field">
                    <label class="form-label">打印分辨率</label>
                    <select name="quality" class="form-select" id="qualitySelect" spellcheck="false">
                        {% if printer_caps and printer_caps.get('resolutions') %}
                            {% for r in printer_caps.resolutions %}
                            <option value="{{r}}">{{r}}</option>
                            {% endfor %}
                        {% else %}
                            <option value="600x600">600x600</option>
                        {% endif %}
                    </select>
                </div>
                <div class="col-12 print-file-area">
                    <label class="form-label">选择文件（支持PDF/JPG/PNG/DOC/DOCX/PPT/PPTX/XLS/XLSX/TXT，支持多选和拖拽）</label>
                    
                    <!-- 拖拽上传区域 -->
                    <div class="file-drop-area" id="fileDropArea">
                        <div class="drop-icon"></div>
                        <h5>拖拽文件到此处</h5>
                        <p>或者 <strong>点击选择文件</strong></p>
                        <p class="text-muted small">支持多个文件同时上传</p>
                        <input type="file" name="file" multiple class="form-control" id="fileInput" style="display: none;" spellcheck="false">
                    </div>
                    
                    <!-- 选中的文件列表 -->
                    <div class="file-list" id="fileList" style="display: none;">
                        <h6>已选择的文件：</h6>
                        <div id="selectedFiles"></div>
                    </div>
                </div>
                <div class="col-12 text-end print-submit-row">
                    {% if printers %}
                        <button type="submit" class="btn btn-primary px-4" id="printButton">上传并打印</button>
                    {% else %}
                        <button type="button" class="btn btn-secondary px-4" disabled title="无可用打印机">无法打印 - 请检查打印机</button>
                    {% endif %}
                </div>
            </form>

            <!-- 打印设置说明 -->
            <div class="alert alert-info collapsible-help" data-collapse-key="print-help">
                <div class="collapsible-help-header" role="button" tabindex="0" onclick="toggleHelpSection('print-help')" onkeydown="handleHelpToggleKey(event, 'print-help')">
                    <h6><i class="bi bi-info-circle"></i> 静默打印说明</h6>
                    <span class="collapsible-help-toggle"></span>
                </div>
                <div class="collapsible-help-body">
                <ul class="mb-0 small">
                    <li><strong>静默打印:</strong> 无需手动确认，文件会自动发送到选择的打印机</li>
                    <li><strong>PDF文件:</strong> 优先使用WPS、Office或Adobe Reader进行静默打印，自动选择最佳方案</li>
                    <li><strong>️图片文件:</strong> 支持JPG、PNG格式，使用Windows图片查看器静默打印</li>
                    <li><strong>Office文档:</strong> 支持DOC/DOCX、XLS/XLSX、PPT/PPTX，使用Office应用程序或COM对象</li>
                    <li><strong>文本文件:</strong> 支持TXT格式，直接发送到打印机</li>
                    <li><strong>️打印参数:</strong> 直接应用您设置的打印参数（双面、纸张、质量）到实际打印任务</li>
                    <li><strong>老旧打印机兼容性:</strong> 已对2000年左右的部分老旧打印机进行测试，不支持TXT格式，支持PDF、JPG、PNG、DOC、DOCX、PPT、PPTX、XLS、XLSX格式。采用转换为BMP流方案和直接发送原始字节流方案。</li>
                    <li><strong>主流打印机支持格式:</strong> PDF、JPG、PNG、DOC、DOCX、PPT、PPTX、XLS、XLSX、TXT</li>
                    <li><strong>成功标识:</strong> 看到绿色表示打印任务已成功发送</li>
                    <li><strong>备用方案:</strong> 如果主要方法失败，系统会自动尝试备用打印方案</li>
                </ul>
                </div>
            </div>
        </div>
        
        <!-- 环境状态提示（仅打印页面显示） -->
        {% if env_status %}
        <div class="alert alert-{{env_status.type}}">
            <h6><i class="bi bi-{{env_status.icon}}"></i> {{env_status.title}}</h6>
            <div class="small">{{env_status.message|safe}}</div>
        </div>
        {% endif %}
        
        <div class="alert alert-info collapsible-help" data-collapse-key="queue-help">
            <div class="collapsible-help-header" role="button" tabindex="0" onclick="toggleHelpSection('queue-help')" onkeydown="handleHelpToggleKey(event, 'queue-help')">
                <h6><i class="bi bi-lightbulb"></i> 队列管理功能</h6>
                <span class="collapsible-help-toggle"></span>
            </div>
            <div class="collapsible-help-body">
            <ul class="mb-0 small">
                <li><strong>️取消打印:</strong> 如果文件还未实际打印，点击删除按钮可取消打印</li>
                <li><strong>自动清理:</strong> 上传的文件会在10分钟后自动清理，也可以点击删除按钮手动删除</li>
                <li><strong>清空队列:</strong> 可以一键删除所有待打印文件，节省时间</li>
                <li><strong>️文件预览:</strong> 打印前可以预览文件内容，确保正确性</li>
                <li><strong>文件信息:</strong> 显示文件大小、类型和上传时间，便于管理</li>
                <li><strong>使用建议:</strong> 打印前检查队列，删除不需要的文件可以节省纸张</li>
            </ul>
            </div>
        </div>
        
        <div class="alert alert-warning collapsible-help" data-collapse-key="advanced-print-help">
            <div class="collapsible-help-header" role="button" tabindex="0" onclick="toggleHelpSection('advanced-print-help')" onkeydown="handleHelpToggleKey(event, 'advanced-print-help')">
                <h6><i class="bi bi-info-circle"></i> 高级打印功能提示</h6>
                <span class="collapsible-help-toggle"></span>
            </div>
            <div class="collapsible-help-body">
            <div class="small">
                如需使用<strong>横版打印</strong>、<strong>自定义页码范围</strong>、<strong>布局调整</strong>、<strong>黑白打印</strong>等功能，请先使用 <strong>Microsoft Office</strong> 或 <strong>WPS Office</strong> 等办公软件在本地进行编辑/设置后导出，再发送到此服务进行打印。 这样可以获得最佳的打印效果！
            </div>
            </div>
        </div>
        
        <div class="d-flex justify-content-between align-items-center mt-4">
            <h4 class="mb-0">打印队列</h4>
            {% if files %}
            <button type="button" class="btn btn-outline-danger btn-sm" onclick="deleteAllFiles()" title="清空所有待打印文件">
                ️ 清空队列
            </button>
            {% endif %}
        </div>
        <table class="table table-sm table-hover align-middle mt-2 queue-table">
            <thead class="table-light"><tr><th>文件名</th><th>大小</th><th>上传时间</th><th>操作</th></tr></thead>
            <tbody>
            {% for f in files %}
                <tr>
                    <td>
                        <div class="d-flex align-items-center">
                            <span class="file-type-badge file-type-{{f.extension}} me-2">{{f.extension}}</span>
                            <div>
                                <div class="file-name">{{f.name}}</div>
                                {% if f.extension in ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'] %}
                                    <small class="text-muted">办公文档</small>
                                {% elif f.extension in ['jpg', 'jpeg', 'png', 'gif', 'bmp'] %}
                                    <small class="text-muted">图片文件</small>
                                {% elif f.extension == 'txt' %}
                                    <small class="text-muted">文本文件</small>
                                {% else %}
                                    <small class="text-muted">其他文件</small>
                                {% endif %}
                            </div>
                        </div>
                    </td>
                    <td class="text-muted">{{f.size}}</td>
                    <td class="text-muted">{{f.time}}</td>
                    <td>
                        <div class="btn-group btn-group-sm" role="group">
                            <a href="/preview/{{f.name}}" target="_blank" class="btn btn-outline-primary btn-sm" title="预览文件">
                                ️ 预览
                            </a>
                            <button type="button" class="btn btn-outline-danger btn-sm" 
                                    onclick="deleteFile('{{f.name}}')" title="从队列中删除">
                                ️ 删除
                            </button>
                        </div>
                    </td>
                </tr>
            {% else %}
                <tr>
                    <td colspan="4">
                        <div class="empty-queue">
                            <div class="mb-3"></div>
                            <h6 class="mb-2">队列为空</h6>
                            <p class="mb-0">还没有待打印的文件，请先上传文件</p>
                        </div>
                    </td>
                </tr>
            {% endfor %}
            </tbody>
        </table>
        
        {% if files %}
        <div class="alert alert-info">
            <small>
                 当前队列中有 <strong>{{files|length}}</strong> 个文件 | 
                ️ 点击删除按钮可以从队列中移除文件 | 
                 文件会在10分钟后自动清理
            </small>
        </div>
        {% endif %}

        <h4 class="mt-4">打印日志</h4>
        <ul class="list-group log-list mb-0">
            {% for l in logs %}
                <li class="list-group-item">{{l}}</li>
            {% endfor %}
        </ul>
    </div>

    <!-- 扫描管理内容 -->
    <div class="tab-content" id="scanTab" style="display: none;">
        <div class="main-content">
            <!-- 扫描区域 -->
            <div class="card mb-4" style="border-left: 4px solid #17a2b8;">
                <div class="card-header bg-light">
                    <h5 class="mb-0"><i class="bi bi-upc-scan"></i>  扫描功能</h5>
                </div>
                <div class="card-body">
                    <form id="scanForm" class="row g-3">
                        <div class="col-md-6">
                            <label class="form-label">选择扫描仪 
                                  <button type="button" class="btn btn-sm inline-refresh-btn" onclick="refreshScannerList()" title="刷新扫描仪列表">
                                     刷新
                                </button>
                            </label>
                            <select name="scanner_id" class="form-select" id="scannerSelect" spellcheck="false">
                                <option value="">正在检测扫描仪...</option>
                            </select>
                            <div class="form-text text-muted" id="scannerHelpText">
                                <small> 支持WIA兼容的扫描仪设备，包括多功能一体机</small>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">扫描格式</label>
                            <select name="format" class="form-select" id="formatSelect" spellcheck="false">
                                <option value="PNG" selected>PNG (推荐)</option>
                                <option value="JPEG">JPEG</option>

                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">&nbsp;</label>
                            <button type="button" class="btn btn-info w-100" id="scanButton" onclick="startScan()">
                                 开始扫描
                            </button>
                        </div>
                    </form>
                    
                    <!-- 醒目的操作提示 -->
                    <div class="alert alert-warning mt-3" style="border-left: 4px solid #ffc107; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);">
                        <h6 class="mb-3" style="color: #856404; font-weight: bold;">
                            ️ 扫描前请确认
                        </h6>
                        <div class="row">
                            <div class="col-md-6">
                                <ul class="mb-0 small" style="color: #856404;">
                                    <li><strong>已放入文件：</strong>已放入要扫描的文件</li>
                                    <li><strong>位置正确：</strong>文件位置和方向摆放正确</li>
                                </ul>
                            </div>
                            <div class="col-md-6">
                                <ul class="mb-0 small" style="color: #856404;">
                                    <li><strong>盖子关好：</strong>扫描仪盖子已完全盖好</li>
                                    <li><strong>避免冲突：</strong>扫描期间请勿使用打印功能</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    
                    <!-- 设备状态提示 -->
                    <div id="deviceStatusAlert" class="alert alert-danger mt-2" style="display: none;">
                        <h6><i class="bi bi-exclamation-circle"></i> 设备忙碌中</h6>
                        <p id="deviceStatusMessage" class="mb-0"></p>
                    </div>
                    
                    <!-- 扫描状态显示区域 -->
                    <div id="scanStatusArea" class="mt-3" style="display: none;">
                        <div class="alert alert-success" id="scanSuccessAlert" style="display: none;">
                            <h6> 扫描成功</h6>
                            <p id="scanSuccessMessage"></p>
                        </div>
                        <div class="alert alert-danger" id="scanErrorAlert" style="display: none;">
                            <h6> 扫描失败</h6>
                            <p id="scanErrorMessage"></p>
                        </div>
                    </div>
                    
                    <div class="alert alert-info mt-3 collapsible-help" data-collapse-key="scan-help">
                        <div class="collapsible-help-header" role="button" tabindex="0" onclick="toggleHelpSection('scan-help')" onkeydown="handleHelpToggleKey(event, 'scan-help')">
                            <h6><i class="bi bi-info-circle"></i> 扫描说明</h6>
                            <span class="collapsible-help-toggle"></span>
                        </div>
                        <div class="collapsible-help-body">
                        <ul class="mb-0 small">
                            <li><strong>静默扫描:</strong> 完全后台执行，无弹窗干扰，不会影响目标电脑的正常使用</li>
                            <li><strong>扫描仪检测:</strong> 自动检测WIA兼容的扫描仪和多功能一体机</li>
                            <li><strong>支持格式:</strong> PNG、JPEG格式，建议使用PNG获得最佳质量</li>
                            <li><strong>自动设置:</strong> 尝试自动设置300 DPI分辨率和彩色扫描模式</li>
                            <li><strong>自动保存:</strong> 扫描完成的文件会保存在【扫描文件】区域，支持预览、下载、打印和删除</li>
                            <li><strong>自动删除:</strong> 扫描文件超过30分钟会自动删除，避免占用磁盘空间</li>
                            <li><strong>兼容性:</strong> 兼容大多数品牌的扫描仪，如Canon、HP、Epson等</li>
                            <li><strong>失败处理:</strong> 如果确实安装了扫描仪但扫描失败，请在重启扫描仪或人工操作</li>
                        </ul>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- 扫描文件管理区域 -->
            <div class="card mb-4" style="border-left: 4px solid #28a745;">
                <div class="card-header bg-light d-flex justify-content-between align-items-center">
                    <h5 class="mb-0"><i class="bi bi-folder-fill"></i>  扫描文件管理</h5>
                    <div class="btn-group btn-group-sm">
                        <button type="button" class="btn btn-outline-success" onclick="refreshScannedFiles()" title="刷新扫描文件列表">
                             刷新
                        </button>
                        <button type="button" class="btn btn-outline-danger" onclick="clearAllScannedFiles()" title="清空所有扫描文件">
                             清空队列
                        </button>
                    </div>
                </div>
                <div class="card-body">
                    <!-- 扫描文件列表 -->
                    <div id="scannedFilesList">
                        <div class="text-center text-muted py-4">
                            <i class="bi bi-folder2-open" style="font-size: 2em;"></i>
                            <p class="mt-2">正在加载扫描文件...</p>
                        </div>
                    </div>
                    
                    <!-- 文件操作提示 -->
                    <div class="alert alert-success mt-3">
                        <h6><i class="bi bi-info-circle"></i> 文件操作说明</h6>
                        <ul class="mb-0 small">
                            <li><strong>️ 预览:</strong> 点击图片文件可直接预览，PDF等文件需下载后查看</li>
                            <li><strong> 下载:</strong> 点击下载按钮将文件保存到本地</li>
                            <li><strong>️ 打印:</strong> 直接将扫描文件发送到打印机</li>
                            <li><strong>️ 删除:</strong> 不需要的文件可以删除以节省空间</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    </div>







    </div>
</div>
<!-- 使用本地 Bootstrap JS -->
<script src="/static/bootstrap.bundle.min.js"></script>
<script>
// 获取设备名并设置到请求头
function getDeviceName() {
    let deviceName = '';
    
    // 尝试多种方法获取设备名
    try {
        // 方法1: 尝试获取网络信息中的主机名 (某些浏览器支持)
        if (navigator.connection && navigator.connection.effectiveType) {
            // 现代浏览器可能提供网络信息
        }
        
        // 方法2: 从User Agent解析设备信息
        const ua = navigator.userAgent;
        
        // Android设备
        if (/Android/i.test(ua)) {
            const match = ua.match(/Android.*?;\s*([^)]+)/);
            if (match) {
                deviceName = match[1].trim();
            } else {
                deviceName = 'Android设备';
            }
        }
        // iOS设备
        else if (/iPhone/i.test(ua)) {
            deviceName = 'iPhone';
        }
        else if (/iPad/i.test(ua)) {
            deviceName = 'iPad';
        }
        // Windows设备
        else if (/Windows/i.test(ua)) {
            const winMatch = ua.match(/Windows NT ([\d.]+)/);
            if (winMatch) {
                const version = winMatch[1];
                const versionNames = {
                    '10.0': 'Win10/11电脑',
                    '6.3': 'Win8.1电脑',
                    '6.2': 'Win8电脑',
                    '6.1': 'Win7电脑'
                };
                deviceName = versionNames[version] || `Windows NT ${version}电脑`;
            } else {
                deviceName = 'Windows电脑';
            }
        }
        // Mac设备
        else if (/Mac|Macintosh/i.test(ua)) {
            const macMatch = ua.match(/Mac OS X ([\d_]+)/);
            if (macMatch) {
                const version = macMatch[1].replace(/_/g, '.');
                deviceName = `macOS ${version}`;
            } else {
                deviceName = 'Mac电脑';
            }
        }
        // Linux设备
        else if (/Linux/i.test(ua)) {
            deviceName = 'Linux电脑';
        }
        else {
            deviceName = '未知设备';
        }
    } catch (e) {
        deviceName = '设备信息获取失败';
    }
    
    return deviceName;
}

// 为AJAX请求添加设备名 (用于删除操作等)
function addDeviceNameToRequests() {
    const deviceName = encodeURIComponent(getDeviceName());
    
    // 拦截fetch请求
    const originalFetch = window.fetch;
    window.fetch = function(url, options = {}) {
        options.headers = options.headers || {};
        options.headers['X-Device-Name'] = deviceName;
        return originalFetch(url, options);
    };
}

// 简化的警告消息处理
document.addEventListener('DOMContentLoaded', function() {
    // 初始化AJAX请求拦截
    addDeviceNameToRequests();
    
    // 设置设备名到隐藏字段
    const deviceNameField = document.getElementById('deviceNameField');
    if (deviceNameField) {
        deviceNameField.value = getDeviceName();
    }
    
    // 为表单提交添加设备名
    const printForm = document.getElementById('printForm');
    if (printForm) {
        printForm.addEventListener('submit', function() {
            const deviceNameField = document.getElementById('deviceNameField');
            if (deviceNameField) {
                deviceNameField.value = getDeviceName();
            }
        });
    }
    
    // 警告消息自动关闭功能
    const alerts = document.querySelectorAll('.alert');
    alerts.forEach(alert => {
        const closeBtn = alert.querySelector('.btn-close');
        if (closeBtn) {
            closeBtn.addEventListener('click', function() {
                alert.style.display = 'none';
            });
        }
        
        // 5秒后自动关闭成功消息
        if (alert.classList.contains('alert-success')) {
            setTimeout(() => {
                alert.style.display = 'none';
            }, 5000);
        }
    });
});

// 根据所选打印机实时获取并填充分辨率与纸张列表
function refreshPrinterInfo() {
    const printerSelect = document.getElementById('printerSelect');
    const paperSelect = document.getElementById('paperSelect');
    const qualitySelect = document.getElementById('qualitySelect');
    if (!printerSelect) return;
    const selectedPrinter = printerSelect.value;
    if (!selectedPrinter) return;

    fetch('/api/printer_info?printer=' + encodeURIComponent(selectedPrinter))
        .then(r => r.json())
        .then(data => {
            if (!data.success) return;
            const caps = data.capabilities || {};
            // 填充纸张
            if (paperSelect) {
                const prev = paperSelect.value;
                paperSelect.innerHTML = '';
                if (caps.papers && caps.papers.length) {
                    let a4Index = -1;
                    caps.papers.forEach((p, idx) => {
                        const opt = document.createElement('option');
                        opt.value = p.id;
                        opt.textContent = p.name;
                        paperSelect.appendChild(opt);
                        if (p.id === 9 || (typeof p.name === 'string' && p.name.toUpperCase().includes('A4'))) {
                            a4Index = idx;
                        }
                    });
                    // 优先恢复之前选择；否则默认选A4；否则选第一项
                    if (prev && Array.from(paperSelect.options).some(o => String(o.value) === String(prev))) {
                        paperSelect.value = prev;
                    } else if (a4Index >= 0) {
                        paperSelect.selectedIndex = a4Index;
                    } else {
                        paperSelect.selectedIndex = 0;
                    }
                } else {
                    const opt = document.createElement('option');
                    opt.value = '9'; // A4 ID
                    opt.textContent = 'A4';
                    paperSelect.appendChild(opt);
                }
            }
            // 填充分辨率
            if (qualitySelect) {
                qualitySelect.innerHTML = '';
                if (caps.resolutions && caps.resolutions.length) {
                    caps.resolutions.forEach(r => {
                        const opt = document.createElement('option');
                        opt.value = r;
                        opt.textContent = r;
                        qualitySelect.appendChild(opt);
                    });
                } else {
                    const opt = document.createElement('option');
                    opt.value = '600x600';
                    opt.textContent = '600x600';
                    qualitySelect.appendChild(opt);
                }
            }
        })
        .catch(() => {});
}



// 添加表单提交验证
document.addEventListener('DOMContentLoaded', function() {
    restoreHelpSections();

    const uploadForm = document.querySelector('form[enctype="multipart/form-data"]');
    const printButton = document.getElementById('printButton');
    const copiesSelect = document.getElementById('copiesSelect');
    const copiesCustomWrap = document.getElementById('copiesCustomWrap');
    const copiesCustomInput = document.getElementById('copiesCustomInput');

    function syncCopiesField() {
        if (!copiesSelect) return;
        const isCustom = copiesSelect.value === 'custom';
        if (copiesCustomWrap) {
            copiesCustomWrap.style.display = isCustom ? 'block' : 'none';
        }
        if (isCustom && copiesCustomInput && (!copiesCustomInput.value || parseInt(copiesCustomInput.value, 10) < 1)) {
            copiesCustomInput.value = '1';
        }
    }

    if (copiesSelect) {
        copiesSelect.addEventListener('change', syncCopiesField);
        syncCopiesField();
    }
    
    if (uploadForm) {
        uploadForm.addEventListener('submit', function(e) {
            const printerSelect = document.getElementById('printerSelect');
            const selectedPrinter = printerSelect ? printerSelect.value : '';
            const copiesField = document.getElementById('copiesSelect');
            const customCopiesInput = document.getElementById('copiesCustomInput');
            const copiesValue = copiesField ? copiesField.value : '1';

            if (copiesField && copiesValue === 'custom') {
                const customCopies = customCopiesInput ? parseInt(customCopiesInput.value, 10) : NaN;
                if (!customCopies || customCopies < 1) {
                    e.preventDefault();
                    alert('请输入有效的自定义份数，至少为 1');
                    return false;
                }
                copiesField.value = String(customCopies);
            }
            
            // 检查是否选择了有效的打印机
            if (!selectedPrinter || selectedPrinter === '' || selectedPrinter === '未检测到可用打印机') {
                e.preventDefault();
                alert('请先选择一个有效的打印机！\\n\\n如果没有看到打印机，请检查：\\n1. 打印机是否正确连接\\n2. 打印机驱动是否安装\\n3. 打印机是否处于联机状态');
                return false;
            }
            
            // 检查是否选择了文件
            const fileInput = document.querySelector('input[type="file"]');
            if (fileInput && fileInput.files.length === 0) {
                e.preventDefault();
                alert('请选择要打印的文件！\\n\\n您可以：\\n1. 点击拖拽区域选择文件\\n2. 直接拖拽文件到上传区域');
                return false;
            }
            
            // 显示加载状态
            if (printButton) {
                printButton.disabled = true;
                printButton.innerHTML = ' 处理中...';
                
                // 5秒后恢复按钮状态（防止页面未刷新）
                setTimeout(() => {
                    printButton.disabled = false;
                    printButton.innerHTML = '上传并打印';
                }, 5000);
            }
            
            return true;
        });
    }
});

// 刷新打印机列表的函数
function refreshPrinterList() {
    const refreshButton = document.querySelector('button[onclick="refreshPrinterList()"]');
    const printerSelect = document.getElementById('printerSelect');
    
    if (refreshButton) {
        refreshButton.disabled = true;
        refreshButton.innerHTML = ' 刷新中...';
    }
    
    // 发送刷新请求
    fetch('/api/refresh_printers')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // 清空当前选项
                printerSelect.innerHTML = '';
                
                if (data.printers && data.printers.length > 0) {
                    // 添加新的打印机选项
                    data.printers.forEach(printer => {
                        const option = document.createElement('option');
                        option.value = printer;
                        option.textContent = printer;
                        
                        // 如果是默认打印机，添加标记并选中
                        if (printer === data.default_printer) {
                            option.textContent += ' (默认)';
                            option.selected = true;
                        }
                        
                        printerSelect.appendChild(option);
                    });
                    
                    // 显示成功消息
                    alert(data.message);

                    // 刷新当前选中打印机的能力（纸张/分辨率）
                    refreshPrinterInfo();
                } else {
                    // 没有找到打印机
                    const option = document.createElement('option');
                    option.value = '';
                    option.textContent = '未检测到可用打印机';
                    printerSelect.appendChild(option);
                    
                    alert('未检测到可用的物理打印机');
                }
            } else {
                alert('刷新失败: ' + data.error);
            }
        })
        .catch(error => {
            console.error('刷新打印机列表失败:', error);
            alert('刷新失败，请检查网络连接');
        })
        .finally(() => {
            // 恢复按钮状态
            if (refreshButton) {
                refreshButton.disabled = false;
                refreshButton.innerHTML = ' 刷新';
            }
        });
}

// ============= 页面切换功能 =============

function switchTab(tabName) {
    // 获取标签页按钮
    const printTabBtn = document.getElementById('printTabBtn');
    const scanTabBtn = document.getElementById('scanTabBtn');
    
    // 获取标签页内容
    const printTab = document.getElementById('printTab');
    const scanTab = document.getElementById('scanTab');
    
    if (tabName === 'print') {
        // 显示打印标签页
        printTab.style.display = 'block';
        scanTab.style.display = 'none';
        
        // 更新按钮样式
        printTabBtn.className = 'btn btn-primary';
        scanTabBtn.className = 'btn btn-outline-primary';
        
        // 更新页面标题
        document.title = '内网打印及扫描服务 - 打印';
        
    } else if (tabName === 'scan') {
        // 显示扫描标签页
        printTab.style.display = 'none';
        scanTab.style.display = 'block';
        
        // 更新按钮样式
        printTabBtn.className = 'btn btn-outline-primary';
        scanTabBtn.className = 'btn btn-primary';
        
        // 更新页面标题
        document.title = '内网打印及扫描服务 - 扫描';
        
        // 如果是第一次切换到扫描标签，刷新扫描仪列表和扫描文件列表
        if (typeof refreshScannerList === 'function') {
            refreshScannerList();
        }
        if (typeof refreshScannedFiles === 'function') {
            refreshScannedFiles();
        }
    }
    
    // 保存当前标签页到localStorage
    try {
        localStorage.setItem('currentTab', tabName);
    } catch (e) {
        // 忽略localStorage错误
    }
}

// 页面加载时恢复上次选择的标签页
function restoreLastTab() {
    try {
        const lastTab = localStorage.getItem('currentTab');
        if (lastTab && (lastTab === 'print' || lastTab === 'scan')) {
            switchTab(lastTab);
        } else {
            // 默认显示打印标签页
            switchTab('print');
        }
    } catch (e) {
        // 如果localStorage不可用，默认显示打印标签页
        switchTab('print');
    }
}

function setHelpSectionState(sectionKey, isExpanded) {
    const section = document.querySelector('[data-collapse-key="' + sectionKey + '"]');
    if (!section) return;

    if (isExpanded) {
        section.classList.remove('collapsed');
    } else {
        section.classList.add('collapsed');
    }

    try {
        localStorage.setItem('helpSectionState:' + sectionKey, isExpanded ? 'expanded' : 'collapsed');
    } catch (e) {
        // 忽略localStorage错误
    }
}

function toggleHelpSection(sectionKey) {
    const section = document.querySelector('[data-collapse-key="' + sectionKey + '"]');
    if (!section) return;
    setHelpSectionState(sectionKey, section.classList.contains('collapsed'));
}

function handleHelpToggleKey(event, sectionKey) {
    if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        toggleHelpSection(sectionKey);
    }
}

function restoreHelpSections() {
    const helpSections = document.querySelectorAll('[data-collapse-key]');
    helpSections.forEach(section => {
        const sectionKey = section.getAttribute('data-collapse-key');
        let isExpanded = true;
        try {
            const savedState = localStorage.getItem('helpSectionState:' + sectionKey);
            if (savedState === 'collapsed') {
                isExpanded = false;
            }
        } catch (e) {
            // 忽略localStorage错误，默认展开
        }
        setHelpSectionState(sectionKey, isExpanded);
    });
}

// ============= 扫描功能JavaScript =============

// 刷新扫描仪列表
function refreshScannerList() {
    const scannerSelect = document.getElementById('scannerSelect');
    const refreshButton = document.querySelector('button[onclick="refreshScannerList()"]');
    const helpText = document.getElementById('scannerHelpText');
    
    if (refreshButton) {
        refreshButton.disabled = true;
        refreshButton.innerHTML = ' 检测中...';
    }
    
    // 显示加载状态
    scannerSelect.innerHTML = '<option value="">正在检测扫描仪...</option>';
    
    fetch('/api/scanners')
        .then(response => response.json())
        .then(data => {
            scannerSelect.innerHTML = '';
            
            if (data.status === 'success' && data.scanners && data.scanners.length > 0) {
                data.scanners.forEach(scanner => {
                    const option = document.createElement('option');
                    option.value = scanner.id;
                    option.textContent = scanner.name;
                    option.setAttribute('data-type', scanner.type);
                    option.setAttribute('data-available', scanner.available);
                    
                    if (!scanner.available) {
                        option.textContent += ' (不可用)';
                        option.disabled = true;
                    }
                    
                    scannerSelect.appendChild(option);
                });
                
                // 自动选择第一个可用的扫描仪
                const firstAvailable = data.scanners.find(s => s.available);
                if (firstAvailable) {
                    scannerSelect.value = firstAvailable.id;
                }
                
                // 检查是否只有默认选项（即未检测到真实扫描仪）
                const realScanners = data.scanners.filter(s => s.type !== 'Default');
                if (realScanners.length > 0) {
                    helpText.innerHTML = `<small> 检测到 ${realScanners.length} 台扫描设备</small>`;
                } else {
                    helpText.innerHTML = '<small>️ 未检测到扫描仪，将尝试使用系统默认设备，可尝试扫描</small>';
                }
                
                // 启用扫描按钮
                const scanButton = document.getElementById('scanButton');
                if (scanButton) {
                    scanButton.disabled = false;
                }
            } else {
                const option = document.createElement('option');
                option.value = 'default';
                option.textContent = '未检测到扫描仪';
                scannerSelect.appendChild(option);
                
                helpText.innerHTML = '<small>️ 未检测到扫描仪，将尝试使用系统默认设备，可尝试扫描</small>';
            }
        })
        .catch(error => {
            console.error('获取扫描仪列表失败:', error);
            scannerSelect.innerHTML = '<option value="default">未检测到扫描仪</option>';
            helpText.innerHTML = '<small> 扫描仪检测失败，将使用默认设备</small>';
        })
        .finally(() => {
            if (refreshButton) {
                refreshButton.disabled = false;
                refreshButton.innerHTML = ' 刷新';
            }
        });
}

// 开始扫描
function startScan() {
    const scannerSelect = document.getElementById('scannerSelect');
    const formatSelect = document.getElementById('formatSelect');
    const scanButton = document.getElementById('scanButton');
    
    if (!scannerSelect.value) {
        alert('请先选择一个扫描仪！');
        return;
    }
    
    // 弹出确认对话框
    const confirmMessage = `️ 扫描确认\n\n请确认以下操作已完成：\n\n 扫描仪中已放入要扫描的文件\n 文件位置和方向正确\n 扫描仪盖子已盖好\n 当前没有打印任务在进行\n\n扫描仪: ${scannerSelect.options[scannerSelect.selectedIndex].text}\n格式: ${formatSelect.value}\n\n️ 扫描期间请勿使用打印功能\n\n确定开始扫描吗？`;
    
    if (confirm(confirmMessage)) {
        // 用户确认后开始扫描
        performScan();
    }
}

// 执行实际的扫描操作
function performScan() {
    const scannerSelect = document.getElementById('scannerSelect');
    const formatSelect = document.getElementById('formatSelect');
    const scanButton = document.getElementById('scanButton');
    
    // 显示扫描进度
    scanButton.disabled = true;
    scanButton.innerHTML = ' 扫描中... 请勿关闭';
    
    // 创建JSON数据
    const requestData = {
        scanner_id: scannerSelect.value,
        scanner_name: scannerSelect.options[scannerSelect.selectedIndex].text,
        format: formatSelect.value
    };
    
    // 发送扫描请求
    fetch('/api/scan', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestData)
    })
    .then(response => response.json())
    .then(data => {
        if (data.status === 'success' || data.status === 'warning') {
            showAlert(data.status === 'success' ? 'success' : 'warning', ` ${data.message}`);
            
            // 扫描成功或部分成功后刷新文件列表
            setTimeout(() => {
                if (typeof refreshScannedFiles === 'function') {
                    refreshScannedFiles();
                }
            }, 2000);
        } else {
            showAlert('danger', ` 扫描失败: ${data.error}`);
        }
    })
    .catch(error => {
        console.error('扫描请求失败:', error);
        showAlert('danger', ` 扫描请求失败: ${error.message || error}`);
    })
    .finally(() => {
        // 恢复扫描按钮
        scanButton.disabled = false;
        scanButton.innerHTML = ' 开始扫描';
    });
}

// 显示扫描进行中的遮罩层


// 页面加载完成后的初始化
document.addEventListener('DOMContentLoaded', function() {
    // 恢复上次选择的标签页
    restoreLastTab();
    
    const printerSelect = document.getElementById('printerSelect');
    if (printerSelect && printerSelect.value) {
        refreshPrinterInfo();
        printerSelect.addEventListener('change', refreshPrinterInfo);
    }
    

    
    // 初始化拖拽文件功能
    initFileDragDrop();
    
    // 注意：不在这里初始化扫描功能，而是在切换到扫描标签时才初始化
});

// 拖拽文件功能
function initFileDragDrop() {
    const dropArea = document.getElementById('fileDropArea');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('fileList');
    const selectedFiles = document.getElementById('selectedFiles');
    
    if (!dropArea || !fileInput) return;
    
    let currentFiles = [];
    
    // 支持的文件类型
    const allowedTypes = ['pdf', 'jpg', 'jpeg', 'png', 'txt', 'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx'];
    
    // 点击区域触发文件选择
    dropArea.addEventListener('click', function(e) {
        // 只阻止链接跳转，不阻止点击事件
        if (e.target.tagName === 'A') {
            e.preventDefault();
        }
        fileInput.click();
    });
    
    // 文件输入框变化
    fileInput.addEventListener('change', function(e) {
        const files = Array.from(e.target.files);
        addFiles(files);
    });
    
    // 阻止默认的拖拽行为
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    // 只在document上阻止拖拽，不影响点击
    ['dragenter', 'dragover'].forEach(eventName => {
        document.body.addEventListener(eventName, function(e) {
            if (e.target !== dropArea && !dropArea.contains(e.target)) {
                e.preventDefault();
                e.stopPropagation();
            }
        }, false);
    });
    
    // 高亮拖拽区域
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    // 处理拖拽文件
    dropArea.addEventListener('drop', handleDrop, false);
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    function highlight() {
        dropArea.classList.add('drag-over');
        dropArea.querySelector('.drop-icon').textContent = '';
    }
    
    function unhighlight() {
        dropArea.classList.remove('drag-over');
        dropArea.querySelector('.drop-icon').textContent = '';
    }
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = Array.from(dt.files);
        addFiles(files);
    }
    
    function addFiles(newFiles) {
        // 过滤允许的文件类型
        const validFiles = newFiles.filter(file => {
            const extension = file.name.split('.').pop().toLowerCase();
            return allowedTypes.includes(extension);
        });
        
        if (validFiles.length !== newFiles.length) {
            const invalidCount = newFiles.length - validFiles.length;
            alert(`有 ${invalidCount} 个文件格式不支持，已忽略。\\n支持的格式: ${allowedTypes.join(', ')}`);
        }
        
        // 添加有效文件到列表（避免重复）
        validFiles.forEach(file => {
            const exists = currentFiles.some(f => f.name === file.name && f.size === file.size);
            if (!exists) {
                currentFiles.push(file);
            }
        });
        
        updateFileList();
        updateFileInput();
    }
    
    function removeFile(index) {
        currentFiles.splice(index, 1);
        updateFileList();
        updateFileInput();
    }
    
    function updateFileList() {
        if (currentFiles.length === 0) {
            fileList.style.display = 'none';
            return;
        }
        
        fileList.style.display = 'block';
        selectedFiles.innerHTML = '';
        
        currentFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span class="file-name">${file.name}</span>
                <span class="file-size">${formatFileSize(file.size)}</span>
                <div class="file-actions" style="display:inline-block; margin-left:10px;">
                    <button type="button" class="btn btn-sm btn-outline-secondary me-1" onclick="previewLocalFile(${index})" title="预览文件">️ 预览</button>
                    <button type="button" class="remove-btn" onclick="removeFileFromList(${index})" title="移除文件">&times;</button>
                </div>
            `;
            selectedFiles.appendChild(fileItem);
        });
    }
    
    function updateFileInput() {
        // 兼容 WinXP/旧版浏览器：DataTransfer 和 input.files 赋值并非所有浏览器都支持。
        // 原生 file input 已经包含用户实际选择的文件时，不强制重建 FileList，避免上传失败。
        try {
            if (typeof DataTransfer === 'undefined') {
                return;
            }
            const dt = new DataTransfer();
            currentFiles.forEach(function(file) {
                if (dt.items && typeof dt.items.add === 'function') {
                    dt.items.add(file);
                }
            });
            try {
                fileInput.files = dt.files;
            } catch (assignError) {
                console.warn('当前浏览器不支持设置 input.files，保留原生文件选择结果:', assignError);
            }
        } catch (e) {
            console.warn('DataTransfer 不可用，使用原生文件输入:', e);
        }
    }
    
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    // 全局函数，供HTML调用
    window.removeFileFromList = function(index) {
        try {
            const file = currentFiles[index];
            if (file) {
                // 如果 file 上有临时预览 URL，则立即撤销并清理定时器
                try {
                    if (file._previewUrl) {
                        try { URL.revokeObjectURL(file._previewUrl); } catch (e) {}
                        file._previewUrl = null;
                    }
                    if (file._revokeTimer) {
                        clearTimeout(file._revokeTimer);
                        file._revokeTimer = null;
                    }
                } catch (e) {
                    // 忽略撤销错误
                }
            }

            currentFiles.splice(index, 1);
            updateFileList();
            updateFileInput();
        } catch (e) {
            console.error('removeFileFromList error:', e);
        }
    };

    // 全局函数：预览本地待上传文件（使用临时 Blob URL）
    window.previewLocalFile = function(index) {
        try {
            const file = currentFiles[index];
            if (!file) return;

            // 创建模态框容器（如果尚未创建）
            let modal = document.getElementById('localPreviewModal');
            if (!modal) {
                modal = document.createElement('div');
                modal.id = 'localPreviewModal';
                modal.className = 'modal fade';
                modal.innerHTML = `
                    <div class="modal-dialog modal-xl modal-dialog-centered">
                        <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title">文件预览</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="关闭"></button>
                            </div>
                            <div class="modal-body" style="max-height:80vh; overflow:auto;">
                                <div id="localPreviewContent"></div>
                            </div>
                            <div class="modal-footer">
                                <a id="localPreviewDownload" class="btn btn-sm btn-outline-secondary" href="#" download style="display:none;">⬇️ 下载</a>
                                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">关闭</button>
                            </div>
                        </div>
                    </div>`;
                document.body.appendChild(modal);

                // 当模态关闭时撤销任何临时 URL（如果 modal 关闭，说明用户结束预览）
                modal.addEventListener('hidden.bs.modal', function() {
                    try {
                        // 撤销所有文件上的 preview URL（保守清理）
                        currentFiles.forEach(f => {
                            try {
                                if (f && f._previewUrl) {
                                    try { URL.revokeObjectURL(f._previewUrl); } catch (e) {}
                                    f._previewUrl = null;
                                }
                                if (f && f._revokeTimer) {
                                    try { clearTimeout(f._revokeTimer); } catch (e) {}
                                    f._revokeTimer = null;
                                }
                            } catch (e) {}
                        });
                    } catch (e) {}
                });
            }

            const content = document.getElementById('localPreviewContent');
            const downloadLink = document.getElementById('localPreviewDownload');
            content.innerHTML = '';
            downloadLink.style.display = 'none';
            downloadLink.href = '#';

            const mime = file.type || '';
            const name = file.name || '文件';

            // 文本类型：使用 FileReader 直接读取并显示（避免使用 objectURL）
            if (mime.startsWith('text/') || /\.(txt|md|csv)$/i.test(name)) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    const pre = document.createElement('pre');
                    pre.style.whiteSpace = 'pre-wrap';
                    pre.style.wordBreak = 'break-word';
                    pre.textContent = e.target.result;
                    content.appendChild(pre);
                };
                reader.onerror = function() {
                    content.innerHTML = '<div class="text-danger">无法读取文本内容</div>';
                };
                reader.readAsText(file, 'utf-8');
            }
            // 图片：使用 objectURL 嵌入 <img>
            else if (mime.startsWith('image/') || /\.(jpg|jpeg|png|gif|bmp|webp)$/i.test(name)) {
                let url = file._previewUrl;
                if (!url) {
                    url = URL.createObjectURL(file);
                    try { file._previewUrl = url; } catch (e) {}
                    try {
                        file._revokeTimer = setTimeout(() => {
                            try { URL.revokeObjectURL(url); } catch (e) {}
                            try { file._previewUrl = null; } catch (e) {}
                            try { file._revokeTimer = null; } catch (e) {}
                        }, 5 * 60 * 1000);
                    } catch (e) {}
                }
                const img = document.createElement('img');
                img.src = url;
                img.className = 'img-fluid';
                img.style.maxHeight = '70vh';
                content.appendChild(img);
                downloadLink.href = url;
                downloadLink.download = name;
                downloadLink.style.display = 'inline-block';
            }
            // PDF：使用 objectURL 嵌入 <iframe>
            else if (mime === 'application/pdf' || /\.pdf$/i.test(name)) {
                let url = file._previewUrl;
                if (!url) {
                    url = URL.createObjectURL(file);
                    try { file._previewUrl = url; } catch (e) {}
                    try {
                        file._revokeTimer = setTimeout(() => {
                            try { URL.revokeObjectURL(url); } catch (e) {}
                            try { file._previewUrl = null; } catch (e) {}
                            try { file._revokeTimer = null; } catch (e) {}
                        }, 5 * 60 * 1000);
                    } catch (e) {}
                }
                const iframe = document.createElement('iframe');
                iframe.src = url;
                iframe.style.width = '100%';
                iframe.style.height = '70vh';
                iframe.frameBorder = '0';
                content.appendChild(iframe);
                downloadLink.href = url;
                downloadLink.download = name;
                downloadLink.style.display = 'inline-block';
            }
            // 其他文件类型：提示并提供下载链接（不自动下载）
            else {
                let url = file._previewUrl;
                if (!url) {
                    url = URL.createObjectURL(file);
                    try { file._previewUrl = url; } catch (e) {}
                    try {
                        file._revokeTimer = setTimeout(() => {
                            try { URL.revokeObjectURL(url); } catch (e) {}
                            try { file._previewUrl = null; } catch (e) {}
                            try { file._revokeTimer = null; } catch (e) {}
                        }, 5 * 60 * 1000);
                    } catch (e) {}
                }
                content.innerHTML = `<div>此文件类型无法在浏览器内预览。您可以点击下方“下载”按钮在本地打开。</div>`;
                downloadLink.href = url;
                downloadLink.download = name;
                downloadLink.style.display = 'inline-block';
            }

            // 显示模态框
            const bsModal = new bootstrap.Modal(modal);
            bsModal.show();

        } catch (e) {
            alert('无法预览该文件: ' + (e && e.message ? e.message : e));
        }
    };

    // 页面卸载时撤销所有尚未撤销的 Blob URL
    window.addEventListener('beforeunload', function() {
        try {
            currentFiles.forEach(file => {
                try {
                    if (file && file._previewUrl) {
                        try { URL.revokeObjectURL(file._previewUrl); } catch (e) {}
                        file._previewUrl = null;
                    }
                    if (file && file._revokeTimer) {
                        try { clearTimeout(file._revokeTimer); } catch (e) {}
                        file._revokeTimer = null;
                    }
                } catch (e) {}
            });
        } catch (e) {}
    });
}

// 删除队列中的文件
function deleteFile(filename) {
    if (confirm(`确定要从队列中删除文件 "${filename}" 吗？\\n\\n删除后无法恢复，如果需要打印需要重新上传。`)) {
        // 显示删除中状态
        const deleteButtons = document.querySelectorAll(`button[onclick="deleteFile('${filename}')"]`);
        deleteButtons.forEach(btn => {
            btn.disabled = true;
            btn.innerHTML = ' 删除中...';
        });
        
        // 发送删除请求
        fetch('/api/delete_file', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                filename: filename
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // 删除成功，刷新页面或移除表格行
                const row = document.querySelector(`button[onclick="deleteFile('${filename}')"]`).closest('tr');
                if (row) {
                    row.style.backgroundColor = '#f8f9fa';
                    row.style.opacity = '0.5';
                    setTimeout(() => {
                        location.reload(); // 刷新页面以更新队列
                    }, 500);
                }
                
                // 显示成功消息
                showAlert('success', ` 文件 "${filename}" 已从队列中删除`);
            } else {
                showAlert('danger', ` 删除失败: ${data.error}`);
                // 恢复按钮状态
                deleteButtons.forEach(btn => {
                    btn.disabled = false;
                    btn.innerHTML = '️ 删除';
                });
            }
        })
        .catch(error => {
            console.error('删除文件时发生错误:', error);
            showAlert('danger', ` 删除文件时发生网络错误: ${error.message || error}`);
            // 恢复按钮状态
            deleteButtons.forEach(btn => {
                btn.disabled = false;
                btn.innerHTML = '️ 删除';
            });
        });
    }
}

// 显示提示消息
function showAlert(type, message) {
    // 创建提示框
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
    alertDiv.style.position = 'fixed';
    alertDiv.style.top = '20px';
    alertDiv.style.right = '20px';
    alertDiv.style.zIndex = '9999';
    alertDiv.style.minWidth = '300px';
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" onclick="this.parentElement.remove()" aria-label="Close"></button>
    `;
    
    document.body.appendChild(alertDiv);
    
    // 3秒后自动消失
    setTimeout(() => {
        if (alertDiv.parentElement) {
            alertDiv.remove();
        }
    }, 3000);
}

// 批量删除功能
function deleteAllFiles() {
    const fileRows = document.querySelectorAll('table tbody tr');
    const fileCount = fileRows.length;
    
    // 排除空队列的情况
    const emptyRow = document.querySelector('table tbody tr td[colspan]');
    if (emptyRow) {
        showAlert('info', '队列为空，没有文件需要删除');
        return;
    }
    
    if (confirm(`确定要删除队列中的所有 ${fileCount} 个文件吗？\\n\\n删除后无法恢复！`)) {
        fetch('/api/delete_all_files', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert('success', ` 已删除 ${data.count} 个文件`);
                setTimeout(() => {
                    location.reload();
                }, 1000);
            } else {
                showAlert('danger', ` 批量删除失败: ${data.error}`);
            }
        })
        .catch(error => {
            console.error('批量删除时发生错误:', error);
            showAlert('danger', ` 批量删除时发生网络错误: ${error.message || error}`);
        });
    }
}

// 清空所有扫描文件
function clearAllScannedFiles() {
    const filesList = document.getElementById('scannedFilesList');
    
    // 检查是否有文件
    if (!filesList.textContent.includes('个扫描文件')) {
        showAlert('info', '暂无扫描文件需要清空');
        return;
    }
    
    // 确认删除
    if (!confirm('确定要清空所有扫描文件吗？此操作不可恢复。')) {
        return;
    }
    
    const clearButton = document.querySelector('button[onclick="clearAllScannedFiles()"]');
    if (clearButton) {
        clearButton.disabled = true;
        clearButton.innerHTML = '清空中...';
    }
    
    fetch('/api/clear_scanned_files', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.status === 'success') {
            showAlert('success', `已清空 ${data.deleted_count} 个扫描文件`);
            refreshScannedFiles();
        } else {
            showAlert('danger', `清空失败: ${data.error || data.message}`);
        }
    })
    .catch(error => {
        console.error('清空扫描文件时发生错误:', error);
        showAlert('danger', `网络错误: ${error.message}`);
    })
    .finally(() => {
        if (clearButton) {
            clearButton.disabled = false;
            clearButton.innerHTML = ' 清空队列';
        }
    });
}

// ================== 扫描文件管理功能 ==================

// 刷新扫描文件列表
function refreshScannedFiles() {
    const refreshButton = document.querySelector('button[onclick="refreshScannedFiles()"]');
    const originalText = refreshButton ? refreshButton.innerHTML : '';
    
    if (refreshButton) {
        refreshButton.disabled = true;
        refreshButton.innerHTML = '刷新中...';
    }
    
    fetch('/api/scanned_files')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                displayScannedFiles(data.files);
            } else {
                console.error('获取扫描文件列表失败:', data.error);
                document.getElementById('scannedFilesList').innerHTML = 
                    `<div class="text-center text-danger py-4">
                        <i class="bi bi-exclamation-circle" style="font-size: 2em;"></i>
                        <p class="mt-2">获取扫描文件失败: ${data.error}</p>
                    </div>`;
            }
        })
        .catch(error => {
            console.error('获取扫描文件列表时发生错误:', error);
            document.getElementById('scannedFilesList').innerHTML = 
                `<div class="text-center text-danger py-4">
                    <i class="bi bi-wifi-off" style="font-size: 2em;"></i>
                    <p class="mt-2">网络错误: ${error.message}</p>
                </div>`;
        })
        .finally(() => {
            if (refreshButton) {
                refreshButton.disabled = false;
                refreshButton.innerHTML = originalText;
            }
        });
}

// 显示扫描文件列表
function displayScannedFiles(files) {
    const filesList = document.getElementById('scannedFilesList');
    
    if (!files || files.length === 0) {
        filesList.innerHTML = 
            `<div class="text-center text-muted py-4">
                <i class="bi bi-folder2-open" style="font-size: 2em;"></i>
                <p class="mt-2">暂无扫描文件</p>
                <small>扫描完成的文件会显示在这里</small>
            </div>`;
        return;
    }
    
    let html = `
        <div class="mb-3">
            <span class="text-muted">共 ${files.length} 个扫描文件</span>
        </div>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead class="table-light">
                    <tr>
                        <th>文件名</th>
                        <th>类型</th>
                        <th>大小</th>
                        <th>创建时间</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    files.forEach(file => {
        const typeIcon = getFileTypeIcon(file.type, file.extension);
        const canPreview = file.type === 'image';
        
        html += `
            <tr>
                <td>
                    ${typeIcon} 
                    ${canPreview ? 
                        `<a href="#" onclick="previewScannedFile('${file.filename}')" class="text-decoration-none">${file.filename}</a>` :
                        file.filename
                    }
                </td>
                <td>
                    <span class="badge bg-secondary">${file.extension.toUpperCase()}</span>
                </td>
                <td class="text-muted">${file.size_str}</td>
                <td class="text-muted">${file.created}</td>
                <td>
                    <div class="btn-group btn-group-sm">
                        ${canPreview ? 
                            `<button class="btn btn-outline-info" onclick="previewScannedFile('${file.filename}')" title="预览">
                                预览
                            </button>` : ''
                        }
                        <button class="btn btn-outline-success" onclick="downloadScannedFile('${file.filename}')" title="下载">
                            下载
                        </button>
                        <button class="btn btn-outline-primary" onclick="printScannedFile('${file.filename}')" title="打印">
                            打印
                        </button>
                        <button class="btn btn-outline-danger" onclick="deleteScannedFile('${file.filename}')" title="删除">
                            删除
                        </button>
                    </div>
                </td>
            </tr>
        `;
    });
    
    html += `
                </tbody>
            </table>
        </div>
    `;
    
    filesList.innerHTML = html;
}

// 获取文件类型图标
function getFileTypeIcon(type, extension) {
    switch (type) {
        case 'image':
            return '️';
        case 'pdf':
            return '';
        default:
            return '';
    }
}

// 预览扫描文件
function previewScannedFile(filename) {
    const modal = document.createElement('div');
    modal.className = 'modal fade';
    modal.innerHTML = `
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title"> 扫描文件预览: ${filename}</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body text-center">
                    <img src="/api/scanned_files/${filename}/preview" 
                         class="img-fluid" 
                         style="max-height: 70vh;" 
                         onerror="this.src='data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAwIiBoZWlnaHQ9IjIwMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB3aWR0aD0iMTAwJSIgaGVpZ2h0PSIxMDAlIiBmaWxsPSIjZGRkIi8+PHRleHQgeD0iNTAlIiB5PSI1MCUiIGZvbnQtc2l6ZT0iMTIiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGR5PSIuM2VtIj7ml6Dms5XpooTop4g8L3RleHQ+PC9zdmc+'; this.alt='预览失败';"
                         alt="扫描文件预览">
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-success" onclick="downloadScannedFile('${filename}')">
                         下载
                    </button>
                    <button type="button" class="btn btn-primary" onclick="printScannedFile('${filename}')">
                        ️ 打印
                    </button>
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">关闭</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    const bsModal = new bootstrap.Modal(modal);
    bsModal.show();
    
    // 模态框关闭后移除DOM元素
    modal.addEventListener('hidden.bs.modal', () => {
        document.body.removeChild(modal);
    });
}

// 下载扫描文件
function downloadScannedFile(filename) {
    const link = document.createElement('a');
    link.href = `/api/scanned_files/${filename}`;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    showAlert('info', ` 正在下载: ${filename}`);
}

// 打印扫描文件
function printScannedFile(filename) {
    const printerSelect = document.getElementById('printerSelect');
    const copiesInput = document.getElementById('copiesInput');
    
    if (!printerSelect || !printerSelect.value) {
        showAlert('warning', '️ 请先选择打印机');
        return;
    }
    
    const printer = printerSelect.value;
    const copies = copiesInput ? parseInt(copiesInput.value) || 1 : 1;
    
    if (confirm(`确定要打印扫描文件 "${filename}" 吗？\\n\\n打印机: ${printer}\\n份数: ${copies}`)) {
        fetch(`/api/scanned_files/${filename}/print`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                printer: printer,
                copies: copies
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert('success', ` ${data.message}`);
            } else {
                showAlert('danger', ` 打印失败: ${data.error}`);
            }
        })
        .catch(error => {
            console.error('打印扫描文件时发生错误:', error);
            showAlert('danger', ` 打印时发生网络错误: ${error.message}`);
        });
    }
}

// 删除扫描文件
function deleteScannedFile(filename) {
    if (confirm(`确定要删除扫描文件 "${filename}" 吗？\\n\\n删除后无法恢复！`)) {
        fetch(`/api/scanned_files/${filename}/delete`, {
            method: 'DELETE'
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showAlert('success', ` ${data.message}`);
                // 刷新扫描文件列表
                refreshScannedFiles();
            } else {
                showAlert('danger', ` 删除失败: ${data.error}`);
            }
        })
        .catch(error => {
            console.error('删除扫描文件时发生错误:', error);
            showAlert('danger', ` 删除时发生网络错误: ${error.message}`);
        });
    }
}

</script>
</body>
</html>
'''
 
# 允许的文件类型
ALLOWED_EXT = {'pdf', 'jpg', 'jpeg', 'png', 'txt', 'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx'}
 
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

def is_physical_printer(printer_name):
    """检查是否为真正的物理打印机
    
    调试模式下返回True（允许使用虚拟打印机测试）
    """
    # 调试模式：允许所有打印机
    if DEBUG_MODE:
        return True
    
    if printer_name in VIRTUAL_PRINTERS:
        return False
    
    # 检查打印机名称中是否包含虚拟打印机的关键词
    virtual_keywords = ['pdf', 'fax', '传真', 'xps', 'onenote', 'virtual', '虚拟', 'send to', 'export', '导出']
    printer_lower = printer_name.lower()
    
    for keyword in virtual_keywords:
        if keyword in printer_lower:
            return False
    
    return True
 
def get_client_info():
    """获取客户端设备信息"""
    try:
        # 获取客户端IP
        client_ip = request.remote_addr or '未知IP'
        
        # 尝试通过各种方式获取计算机名/设备名
        device_name = None
        
        # 方法0: 优先检查表单中的设备名 (最准确的方法)
        try:
            form_device = request.form.get('device_name') if hasattr(request, 'form') and request.form else None
            if form_device:
                device_name = form_device.strip()
        except Exception:
            # 如果无法访问表单数据（比如JSON请求），跳过
            pass
        
        # 方法0.5: 检查是否有自定义的设备名请求头 (客户端可以主动发送)
        if not device_name:
            custom_device = request.headers.get('X-Device-Name') or request.headers.get('Device-Name')
            if custom_device:
                try:
                    # 解码URL编码的设备名
                    import urllib.parse
                    device_name = urllib.parse.unquote(custom_device.strip())
                except Exception:
                    device_name = custom_device.strip()
        
        # 方法1: 检查HTTP请求头中的计算机名信息
        user_agent = request.headers.get('User-Agent', '')
        
        # 方法2: 尝试通过IP地址反向解析获取计算机名
        try:
            import socket
            if client_ip and client_ip != '127.0.0.1' and client_ip != 'localhost':
                try:
                    hostname = socket.gethostbyaddr(client_ip)[0]
                    if hostname and hostname != client_ip:
                        # 尝试处理中文计算机名
                        try:
                            # 如果是中文，尝试正确编码
                            if any(ord(c) > 127 for c in hostname):
                                # 检测到非ASCII字符（可能是中文）
                                hostname_safe = hostname.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
                                if hostname_safe:
                                    device_name = hostname_safe
                                else:
                                    device_name = hostname
                                print(f"检测到中文计算机名: {hostname}")
                            else:
                                device_name = hostname
                        except Exception as e:
                            print(f"计算机名编码处理异常: {e}，使用原值")
                            device_name = hostname
                except socket.herror as e:
                    # DNS反向解析失败（正常情况，继续尝试其他方法）
                    print(f"无法通过DNS反向解析获取计算机名 (IP: {client_ip})")
                except UnicodeDecodeError as e:
                    # 可能是中文计算机名导致的解码错误
                    print(f"计算机名可能包含中文字符，解码失败: {e}")
                    print(f"  建议在系统设置中将计算机名改为英文")
                except Exception as e:
                    print(f"获取计算机名异常: {type(e).__name__}: {e}")
        except Exception as e:
            print(f"socket操作异常: {type(e).__name__}: {e}")
        
        # 方法3: 从User-Agent中提取更详细的设备信息
        if user_agent:
            import re
            
            # Android设备 - 提取具体型号
            if 'android' in user_agent.lower():
                # 尝试多种Android设备型号模式
                patterns = [
                    r'Android.*?;\s*([^)]+?)\s*Build/',  # 标准Android模式
                    r'Android.*?;\s*(.*?)\)',  # 备用模式
                    r'\(([^;]+);\s*wv\)',  # WebView模式
                ]
                for pattern in patterns:
                    match = re.search(pattern, user_agent)
                    if match:
                        model = match.group(1).strip()
                        # 清理一些常见的无用信息
                        model = re.sub(r'^\w+\s*', '', model)  # 移除开头的语言代码
                        if model and len(model) > 2 and model not in ['Mobile', 'Mobile Safari', 'Safari']:
                            device_name = model
                            break
            
            # iPhone设备 - 提取型号
            elif 'iphone' in user_agent.lower():
                iphone_match = re.search(r'iPhone\s*OS\s*[\d_]+.*?\)', user_agent)
                if iphone_match:
                    # 尝试提取更具体的iPhone型号信息
                    cpu_match = re.search(r'iPhone(\d+,\d+)', user_agent)
                    if cpu_match:
                        device_name = f"iPhone({cpu_match.group(1)})"
                    else:
                        device_name = "iPhone"
                else:
                    device_name = "iPhone"
            
            # iPad设备
            elif 'ipad' in user_agent.lower():
                ipad_match = re.search(r'iPad(\d+,\d+)', user_agent)
                if ipad_match:
                    device_name = f"iPad({ipad_match.group(1)})"
                else:
                    device_name = "iPad"
            
            # Windows设备 - 尝试提取Windows版本
            elif 'windows' in user_agent.lower():
                win_match = re.search(r'Windows NT ([\d.]+)', user_agent)
                if win_match:
                    version = win_match.group(1)
                    version_names = {
                        '10.0': 'Win10/11',
                        '6.3': 'Win8.1',
                        '6.2': 'Win8',
                        '6.1': 'Win7'
                    }
                    win_version = version_names.get(version, f'Windows NT {version}')
                    if not device_name:  # 如果没有通过DNS获取到计算机名
                        device_name = f"{win_version}电脑"
                else:
                    if not device_name:
                        device_name = "Windows电脑"
            
            # Mac设备
            elif 'mac' in user_agent.lower() or 'macintosh' in user_agent.lower():
                mac_match = re.search(r'Mac OS X ([\d_]+)', user_agent)
                if mac_match:
                    mac_version = mac_match.group(1).replace('_', '.')
                    if not device_name:
                        device_name = f"macOS {mac_version}"
                else:
                    if not device_name:
                        device_name = "Mac电脑"
            
            # Linux设备
            elif 'linux' in user_agent.lower():
                if not device_name:
                    device_name = "Linux电脑"
        
        # 如果所有方法都失败，使用默认值
        if not device_name:
            device_name = "未知设备"
        
        return f"{client_ip}({device_name})"
    except Exception as e:
        return f"未知客户端(获取信息失败: {e})"

def log_print(filename, printer, copies, duplex, papersize, quality, client_info=None):
    # 改进双面打印日志显示
    duplex_text = {
        1: "单面",
        2: "双面(长边翻转)",
        3: "双面(短边翻转)"
    }.get(int(duplex), f"未知({duplex})")
    
    # 如果没有提供客户端信息，尝试获取
    if client_info is None:
        try:
            client_info = get_client_info()
        except:
            client_info = "未知客户端"
    
    try:
        # 尝试确保所有字符串都能正确编码为UTF-8
        filename_safe = filename.encode('utf-8', errors='replace').decode('utf-8')
        printer_safe = printer.encode('utf-8', errors='replace').decode('utf-8')
        client_info_safe = client_info.encode('utf-8', errors='replace').decode('utf-8')
        
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now()} 客户端: {client_info_safe} 打印: {filename_safe} 打印机: {printer_safe} 份数: {copies} 模式: {duplex_text} 纸张: {papersize} 质量: {quality}\n")
    except Exception as e:
        # 如果日志写入失败，至少输出错误提示
        try:
            print(f"日志写入失败: {e}")
            print(f"  → 可能原因: 文件路径或文件名包含非法字符")
            print(f"  → 日志文件路径: {LOG_FILE}")
            print(f"  → 建议检查计算机名和路径是否包含中文")
        except:
            pass

def log_scan(scanner_name, scan_format, client_info, message):
    """记录扫描操作日志"""
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now()} 客户端: {client_info} 扫描: 扫描仪={scanner_name} 格式={scan_format} 结果={message}\n")
    except Exception as e:
        print(f"记录扫描日志失败: {e}")

def get_scanned_files():
    """获取扫描文件列表"""
    files = []
    try:
        # 确保扫描目录存在
        scan_folder = path_manager.get_scan_dir()
        if not os.path.exists(scan_folder):
            os.makedirs(scan_folder)
            return files
        
        # 获取所有扫描文件
        for filename in os.listdir(scan_folder):
            filepath = os.path.join(scan_folder, filename)
            if os.path.isfile(filepath) and not filename.startswith('.'):
                try:
                    # 获取文件信息
                    stat_result = os.stat(filepath)
                    file_size = stat_result.st_size
                    created_time = datetime.fromtimestamp(stat_result.st_ctime)
                    
                    # 获取文件扩展名
                    file_ext = os.path.splitext(filename)[1].lower()
                    
                    # 确定文件类型
                    if file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
                        file_type = 'image'
                    elif file_ext == '.pdf':
                        file_type = 'pdf'
                    else:
                        file_type = 'other'
                    
                    files.append({
                        'filename': filename,
                        'filepath': filepath,
                        'size': file_size,
                        'size_str': format_file_size(file_size),
                        'created': created_time.strftime('%Y-%m-%d %H:%M:%S'),
                        'type': file_type,
                        'extension': file_ext
                    })
                    
                except Exception as e:
                    print(f"读取扫描文件 {filename} 信息失败: {e}")
                    continue
        
        # 按创建时间倒序排列
        files.sort(key=lambda x: x['created'], reverse=True)
        
    except Exception as e:
        print(f"获取扫描文件列表失败: {e}")
    
    return files

def format_file_size(size_bytes):
    """格式化文件大小"""
    if size_bytes == 0:
        return "0 B"
    size_names = ["B", "KB", "MB", "GB"]
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_names[i]}"

def print_file_with_settings(filepath, printer_name, copies=1, duplex=1, papersize='A4', quality='normal'):
    """使用获取到的真实打印设置进行打印"""
    # 保存原始打印机双面设置（如果需要临时修改）
    saved_duplex = None
    
    try:
        # 如果是 UNC 共享打印机（形如 \\server\printer），尝试建立连接并刷新缓存
        try:
            ensure_printer_connection(printer_name)
        except Exception:
            pass

        print(f"开始打印文件: {filepath}")
        print(f"目标打印机: {printer_name}")
        print(f"打印份数: {copies}")
        print(f"双面设置: {duplex}")
        print(f"纸张大小: {papersize}")
        print(f"打印质量: {quality}")

        # 获取文件扩展名
        file_ext = os.path.splitext(filepath)[1].lower()

        # 如果需要双面打印，先临时启用打印机的双面功能
        if duplex > 1:
            print(f"检测到双面打印需求，准备临时修改打印机设置...")
            saved_duplex = apply_printer_duplex_setting(printer_name, duplex)
            if saved_duplex is not None:
                print(f"已保存打印机原始设置，将在打印完成后恢复")
            time.sleep(0.5)  # 等待设置生效

        # 优先尝试原生 API 打印（支持 DEVMODE，确保双面设置生效）
        def try_native_print():
            try:
                devmode = apply_printer_settings(printer_name, copies, duplex, papersize, quality)
                if devmode is None:
                    print("未能获取有效 DEVMODE，跳过原生打印")
                    return False
                printer_handle = win32print.OpenPrinter(printer_name)
                try:
                    # 打开打印任务
                    doc_info = {
                        'pDocName': os.path.basename(filepath),
                        'pOutputFile': None,
                        'pDatatype': None
                    }
                    job_id = win32print.StartDocPrinter(printer_handle, 1, doc_info)
                    win32print.StartPagePrinter(printer_handle)
                    # 读取文件内容并写入打印机
                    with open(filepath, 'rb') as f:
                        data = f.read()
                        win32print.WritePrinter(printer_handle, data)
                    win32print.EndPagePrinter(printer_handle)
                    win32print.EndDocPrinter(printer_handle)
                    print("原生打印成功，双面参数已生效")
                    return True
                finally:
                    win32print.ClosePrinter(printer_handle)
            except Exception as e:
                print(f"原生打印失败: {e}")
                return False

        # 仅对 PDF、TXT、图片尝试原生打印，Office 文件暂不支持
        if file_ext in ['.pdf', '.txt', '.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            if try_native_print():
                print("打印成功")
                return True
            else:
                print("原生打印未成功，回退到外部程序打印")

        # 其它类型或原生打印失败，回退原有逻辑
        if file_ext == '.pdf':
            return print_pdf_with_settings(filepath, printer_name, copies, duplex, papersize, quality)
        elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            return print_image_silent(filepath, printer_name, copies)
        elif file_ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
            return print_office_silent(filepath, printer_name, copies)
        elif file_ext == '.txt':
            return print_text_file_simple(filepath, printer_name, copies)
        else:
            print(f"未知文件类型 {file_ext}，尝试使用系统默认打印方式")
            return print_with_shell_execute(filepath, printer_name, copies)

    except Exception as e:
        print(f"打印操作失败: {e}")
        return print_file_silent_fallback(filepath, printer_name, copies)
    
    finally:
        # 打印完成后，恢复原始打印机设置
        if saved_duplex is not None:
            print("打印任务完成，准备恢复打印机设置...")
            time.sleep(1)  # 等待打印队列处理
            restore_printer_duplex_setting(printer_name, saved_duplex)
            print("打印机设置已恢复")


def validate_duplex_setting(printer_name, duplex_value):
    """验证双面打印设置是否被打印机支持"""
    try:
        # 获取打印机能力
        caps = get_printer_capabilities(printer_name)
        
        # 如果要求双面打印但打印机不支持
        if duplex_value > 1 and not caps.get('duplex_support', False):
            print(f"️ 打印机 '{printer_name}' 不支持双面打印，将改为单面打印")
            return 1  # 强制改为单面
        
        # 验证具体的双面模式
        duplex_modes = caps.get('duplex_modes', [])
        if duplex_value == 2 and 'long_edge' not in duplex_modes:
            print(f"️ 打印机不支持长边翻转，尝试使用其他双面模式")
            if 'short_edge' in duplex_modes:
                return 3  # 改为短边翻转
            else:
                return 1  # 改为单面
        
        if duplex_value == 3 and 'short_edge' not in duplex_modes:
            print(f"️ 打印机不支持短边翻转，尝试使用其他双面模式") 
            if 'long_edge' in duplex_modes:
                return 2  # 改为长边翻转
            else:
                return 1  # 改为单面
        
        return duplex_value  # 设置有效，保持原值
        
    except Exception as e:
        print(f"验证双面设置时出错: {e}，使用原设置")
        return duplex_value

def save_printer_duplex_setting(printer_name):
    """保存打印机当前的双面打印设置，用于后续恢复"""
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.GetPrinter(printer_handle, 2)['pDevMode']
            if devmode:
                current_duplex = devmode.Duplex
                print(f"保存打印机当前双面设置: {current_duplex}")
                return current_duplex
        finally:
            win32print.ClosePrinter(printer_handle)
    except Exception as e:
        print(f"保存打印机双面设置失败: {e}")
    return None


def restore_printer_duplex_setting(printer_name, original_duplex):
    """恢复打印机的原始双面打印设置"""
    if original_duplex is None:
        print("跳过恢复：未保存原始设置")
        return False
    
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.GetPrinter(printer_handle, 2)['pDevMode']
            if devmode:
                # 恢复原始双面设置
                devmode.Duplex = original_duplex
                devmode.Fields |= win32con.DM_DUPLEX
                
                # 应用到打印机默认设置
                win32print.SetPrinter(printer_name, 2, {'pDevMode': devmode}, 0)
                print(f"已恢复打印机双面设置为: {original_duplex}")
                return True
        finally:
            win32print.ClosePrinter(printer_handle)
    except Exception as e:
        print(f"恢复打印机双面设置失败: {e}")
    
    return False


def apply_printer_duplex_setting(printer_name, duplex):
    """临时应用打印机双面设置到打印机硬件配置"""
    if duplex == 1:
        print("单面打印，无需临时修改打印机设置")
        return None
    
    try:
        # 保存当前设置
        original_duplex = save_printer_duplex_setting(printer_name)
        
        # 应用新的双面设置
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.GetPrinter(printer_handle, 2)['pDevMode']
            if devmode:
                if duplex == 2:
                    devmode.Duplex = win32con.DMDUP_VERTICAL
                    print("临时启用打印机: 双面打印 - 长边翻转")
                elif duplex == 3:
                    devmode.Duplex = win32con.DMDUP_HORIZONTAL
                    print("临时启用打印机: 双面打印 - 短边翻转")
                
                devmode.Fields |= win32con.DM_DUPLEX
                
                # 应用到打印机默认设置
                win32print.SetPrinter(printer_name, 2, {'pDevMode': devmode}, 0)
                print(f"已临时修改打印机双面设置，原始设置已保存")
                return original_duplex
        finally:
            win32print.ClosePrinter(printer_handle)
    except Exception as e:
        print(f"应用打印机双面设置失败: {e}")
    
    return None


def apply_printer_settings(printer_name, copies, duplex, papersize, quality):
    """应用打印机设置，返回设备模式"""
    try:
        # 检查打印机名称
        if not printer_name or printer_name.strip() == "":
            print("错误: 打印机名称为空")
            return None
            
        if printer_name == "未检测到可用打印机":
            print("错误: 无可用打印机")
            return None
        
        # 验证并修正双面打印设置
        original_duplex = duplex
        duplex = validate_duplex_setting(printer_name, duplex)
        if duplex != original_duplex:
            print(f"双面设置已从 {original_duplex} 调整为 {duplex}")
        
        print(f"验证后的双面设置: {duplex}")
        
        # 获取打印机的默认设备模式
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            # 获取设备模式
            devmode = win32print.GetPrinter(printer_handle, 2)['pDevMode']
            if devmode is None:
                print("无法获取设备模式，使用默认设置")
                return None
            
            # 设置打印份数
            if copies > 1:
                devmode.Copies = copies
                print(f"设置打印份数: {copies}")
            
            # 设置双面打印（完整的双面打印逻辑）
            try:
                if duplex == 1:
                    # 单面打印
                    devmode.Duplex = win32con.DMDUP_SIMPLEX
                    print("设置打印模式: 单面打印")
                elif duplex == 2:
                    # 长边翻转双面打印（默认的双面模式）
                    devmode.Duplex = win32con.DMDUP_VERTICAL
                    print("设置打印模式: 双面打印 - 长边翻转")
                elif duplex == 3:
                    # 短边翻转双面打印
                    devmode.Duplex = win32con.DMDUP_HORIZONTAL
                    print("设置打印模式: 双面打印 - 短边翻转")
                else:
                    # 默认为单面打印
                    devmode.Duplex = win32con.DMDUP_SIMPLEX
                    print(f"未知双面设置值 {duplex}，默认使用单面打印")
                    
                # 确保设置生效
                devmode.Fields |= win32con.DM_DUPLEX
                
            except Exception as e:
                print(f"设置双面打印失败: {e}")
                # 失败时默认单面打印
                try:
                    devmode.Duplex = win32con.DMDUP_SIMPLEX
                    devmode.Fields |= win32con.DM_DUPLEX
                    print("双面设置失败，回退到单面打印")
                except:
                    print("双面打印设置完全失败，将使用打印机默认设置")
            
            # 设置纸张大小：支持直接传入DMPAPER数值ID
            try:
                if isinstance(papersize, int) or (isinstance(papersize, str) and papersize.isdigit()):
                    devmode.PaperSize = int(papersize)
                    print(f"设置纸张大小ID: {devmode.PaperSize}")
                else:
                    # 兼容老的名称映射（尽量少用）
                    paper_size_map = {
                        'A4': win32con.DMPAPER_A4,
                        'A3': win32con.DMPAPER_A3,
                        'Letter': win32con.DMPAPER_LETTER,
                        'Legal': win32con.DMPAPER_LEGAL
                    }
                    if papersize in paper_size_map:
                        devmode.PaperSize = paper_size_map[papersize]
                        print(f"设置纸张大小: {papersize}")
            except Exception as e:
                print(f"设置纸张大小失败: {e}")
            
            # 设置打印质量：支持 "600x600" 或 "600 x 600"
            try:
                if isinstance(quality, str) and ('x' in quality or 'X' in quality):
                    parts = quality.lower().replace(' ', '').split('x')
                    if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                        devmode.PrintQuality = int(parts[0])
                        devmode.YResolution = int(parts[1])
                        print(f"设置打印分辨率: {devmode.PrintQuality}x{devmode.YResolution}")
                elif isinstance(quality, int) or (isinstance(quality, str) and quality.isdigit()):
                    devmode.PrintQuality = int(quality)
                    print(f"设置打印质量(单值): {devmode.PrintQuality}")
                else:
                    # 兼容旧的关键字
                    if quality == 'high':
                        devmode.PrintQuality = win32con.DMRES_HIGH
                        print("设置打印质量: 高质量")
                    else:
                        devmode.PrintQuality = win32con.DMRES_MEDIUM
                        print("设置打印质量: 普通")
            except Exception as e:
                print(f"设置打印质量失败: {e}")
            
            return devmode
        finally:
            win32print.ClosePrinter(printer_handle)
    except Exception as e:
        print(f"设置打印参数失败: {e}")
        return None

def print_pdf_with_settings(filepath, printer_name, copies, duplex, papersize, quality):
    """使用设置参数打印PDF文件，优先使用WPS和Office"""
    try:
        print(f"打印PDF文件: {filepath}")
        
         # 尝试使用Adobe Reader静默打印
        
        adobe_paths = [
            r"C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe",
            r"C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe",
            r"C:\\Program Files\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe"
        ]
        
        for adobe_path in adobe_paths:
            if os.path.exists(adobe_path):
                try:
                    # 构建Adobe Reader打印命令
                    cmd = f'"{adobe_path}" /p /h "{filepath}"'
                    print(f"使用Adobe Reader打印: {cmd}")
                    
                    # 尝试应用打印设置（部分阅读器可能不生效）
                    _ = apply_printer_settings(printer_name, copies, duplex, papersize, quality)
                    
                    # 执行打印命令
                    result = os.system(cmd)
                    if result == 0:
                        print("Adobe Reader打印命令执行成功")
                        return True
                except Exception as e:
                    print(f"Adobe Reader打印失败: {e}")
                    continue
        
        # 如果Adobe Reader不可用，回退到简单打印
        return print_pdf_silent(filepath, printer_name, copies)
        
    except Exception as e:
        print(f"PDF打印失败: {e}")
        return False

def print_with_shell_execute(filepath, printer_name, copies):
    """使用ShellExecute进行应用程序调用打印"""
    try:
        print(f" 使用ShellExecute打印: {filepath} -> {printer_name}")
        success_count = 0
        for i in range(copies):
            try:
                # 使用关联的应用程序打印
                result = win32api.ShellExecute(
                    0,  # hwnd
                    'print',  # operation
                    filepath,  # file
                    None,  # parameters
                    None,  # directory
                    0  # show command (SW_HIDE)
                )
                
                if result > 32:  # ShellExecute成功
                    success_count += 1
                    time.sleep(1)  # 给应用程序时间处理
                else:
                    print(f"ShellExecute失败，错误代码: {result}")
                    
            except Exception as e:
                print(f"打印第{i+1}份时出错: {e}")
                
        if success_count > 0:
            return True, f"通过关联应用程序打印已发送 ({success_count}/{copies}份)"
        else:
            return False, "所有打印尝试都失败了"
    except Exception as e:
        return False, f"ShellExecute打印失败: {str(e)}"


def convert_file_to_bmp_bytes(filepath, max_width=2480, max_height=3508):
    """尝试将文件转换为 BMP 字节流（优先支持图片、文本和简单PDF第一页）
    如果是 PDF，则返回第一页的 BMP 字节流；多页支持请使用 `convert_pdf_to_bmp_pages`。
    """
    try:
        ext = os.path.splitext(filepath)[1].lower()
        # 图片直接打开并转换为 BMP
        if ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']:
            img = Image.open(filepath)
            img = img.convert('RGB')
        elif ext == '.txt':
            # 将文本渲染为简单白底黑字图片
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
            font = ImageFont.load_default()
            lines = text.splitlines() or [' ']
            width = min(max_width, max([font.getsize(l)[0] for l in lines]) + 20)
            height = min(max_height, (font.getsize(lines[0])[1] + 2) * len(lines) + 20)
            img = Image.new('RGB', (width, height), 'white')
            draw = ImageDraw.Draw(img)
            y = 10
            for line in lines:
                draw.text((10, y), line, fill='black', font=font)
                y += font.getsize(line)[1] + 2
        elif ext == '.pdf':
            try:
                from pdf2image import convert_from_path
                poppler = get_poppler_path()
                pages = convert_from_path(filepath, first_page=1, last_page=1, dpi=300, poppler_path=poppler)
                img = pages[0].convert('RGB')
            except Exception:
                return None
        else:
            # 尝试由 PIL 打开任意文件（失败则跳过）
            try:
                img = Image.open(filepath).convert('RGB')
            except Exception:
                return None

        # 调整尺寸以适配常见打印分辨率（A4 300dpi 大致 2480x3508）
        img.thumbnail((max_width, max_height), Image.ANTIALIAS)

        bio = io.BytesIO()
        img.save(bio, format='BMP')
        return bio.getvalue()
    except Exception:
        return None


def convert_pdf_to_bmp_pages(filepath, dpi=300, max_width=3508, max_height=4961):
    """将 PDF 转换为高质量位图列表（每页一个 BMP 字节流），返回 bytes 列表或 None"""
    try:
        from pdf2image import convert_from_path
        poppler = get_poppler_path()
        pages = convert_from_path(filepath, dpi=dpi, poppler_path=poppler)
        bmp_list = []
        for img in pages:
            img = img.convert('RGB')
            img.thumbnail((max_width, max_height), Image.ANTIALIAS)
            bio = io.BytesIO()
            img.save(bio, format='BMP')
            bmp_list.append(bio.getvalue())
        return bmp_list
    except Exception as e:
        print(f"PDF->BMP 多页转换失败: {e}")
        return None


def send_pdf_pages_to_printer(printer_name, bmp_pages, copies=1):
    """按页将 BMP 流发送至打印机，尝试在页面间做简单延迟以保证打印机接收。"""
    try:
        for c in range(copies):
            for page_bytes in bmp_pages:
                ok = send_bytes_to_printer_raw(printer_name, page_bytes)
                if not ok:
                    return False
                time.sleep(0.5)
        return True
    except Exception as e:
        print(f"发送 PDF 位图页到打印机失败: {e}")
        return False


def send_bytes_to_printer_raw(printer_name, data_bytes):
    """将原始字节流通过 WritePrinter 写入打印机（简单直接流式写入）"""
    try:
        ph = win32print.OpenPrinter(printer_name)
        try:
            docinfo = ("PythonRaw", None, "RAW")
            job_id = win32print.StartDocPrinter(ph, 1, docinfo)
            win32print.StartPagePrinter(ph)
            win32print.WritePrinter(ph, data_bytes)
            win32print.EndPagePrinter(ph)
            win32print.EndDocPrinter(ph)
            return True
        finally:
            win32print.ClosePrinter(ph)
    except Exception as e:
        print(f"直接写流到打印机失败: {e}")
        return False

def print_file_silent_fallback(filepath, printer_name, copies=1):
    """备用的静默打印方案"""
    try:
        # 方案1: 使用ShellExecute的静默打印
        for i in range(copies):
            win32api.ShellExecute(
                0, 
                'print', 
                filepath, 
                f'/d:"{printer_name}"', 
                '.', 
                win32con.SW_HIDE  # 隐藏窗口
            )
        return True, f"静默打印任务已发送到 {printer_name} ({copies}份)"
        
    except Exception as e1:
        try:
            # 方案2: 使用命令行静默打印
            import tempfile
            
            # 创建批处理文件进行静默打印
            bat_content = f'''@echo off
for /L %%i in (1,1,{copies}) do (
    start /min "" "{filepath}"
)
'''
            with tempfile.NamedTemporaryFile(mode='w', suffix='.bat', delete=False) as bat_file:
                bat_file.write(bat_content)
                bat_file_path = bat_file.name
            
            # 静默执行批处理文件
            subprocess.run([bat_file_path], 
                         creationflags=subprocess.CREATE_NO_WINDOW,
                         shell=True)
            
            # 清理临时文件
            try:
                os.unlink(bat_file_path)
            except:
                pass
                
            return True, f"静默打印任务已发送 ({copies}份) - 备用方案"
            
        except Exception as e2:
            try:
                # 方案3: 最基础的静默方式
                for i in range(copies):
                    subprocess.run(['rundll32.exe', 'mshtml.dll,PrintHTML', filepath],
                                 creationflags=subprocess.CREATE_NO_WINDOW)
                return True, f"基础静默打印已执行 ({copies}份)"
            except Exception as e3:
                # 方案4: 兼容老旧打印机的原始流方案——将文件转换成 BMP 流并直接写入打印机
                try:
                    print("尝试使用备用流式打印（BMP 流）以适配老旧打印机")
                    bmp_bytes = convert_file_to_bmp_bytes(filepath)
                    if bmp_bytes:
                        ok = True
                        for i in range(copies):
                            if not send_bytes_to_printer_raw(printer_name, bmp_bytes):
                                ok = False
                                break
                        if ok:
                            return True, f"已通过位图流发送到 {printer_name} ({copies}份)"
                        else:
                            print("位图流发送到打印机失败")
                    else:
                        print("无法将文件转换为位图流，跳过此方案")
                except Exception as e4:
                    print(f"位图流打印方案失败: {e4}")

                return False, f"所有静默打印方案都失败: {str(e3)}"

def print_text_direct_to_printer(filepath, printer_name, copies=1):
    """使用WIN32 API直接将文本文件发送到指定打印机，支持多种编码"""
    try:
        import win32print
        import chardet
        
        print(f" 使用API直接打印: {filepath}")
        
        # 智能检测文件编码
        content = read_text_with_encoding_detection(filepath)
        if not content:
            return False, "无法读取文本文件内容"
        
        print(f"文件内容长度: {len(content)} 字符")
        
        # 打开指定的打印机
        printer_handle = win32print.OpenPrinter(printer_name)
        
        try:
            success_count = 0
            for i in range(copies):
                # 开始打印作业
                job_id = win32print.StartDocPrinter(printer_handle, 1, ("Text Document", None, "RAW"))
                
                try:
                    win32print.StartPagePrinter(printer_handle)
                    
                    # 智能编码处理：尝试不同的编码方式发送到打印机
                    print_data = None
                    
                    # 尝试多种编码方式
                    encoding_attempts = ['utf-8', 'gbk', 'cp1252', 'latin1']
                    
                    for encoding in encoding_attempts:
                        try:
                            print_data = content.encode(encoding)
                            break
                        except UnicodeEncodeError:
                            continue
                    
                    # 如果所有编码都失败，使用错误替换模式
                    if print_data is None:
                        print_data = content.encode('utf-8', errors='replace')
                        print("️ 使用UTF-8错误替换模式编码")
                    
                    # 发送文本内容到打印机
                    win32print.WritePrinter(printer_handle, print_data)
                    
                    win32print.EndPagePrinter(printer_handle)
                    win32print.EndDocPrinter(printer_handle)
                    success_count += 1
                    print(f" 打印作业 {i+1} 成功")
                    
                except Exception as e:
                    print(f"打印作业 {i+1} 失败: {e}")
                    win32print.AbortPrinter(printer_handle)
                
            return True, f"直接打印到 {printer_name} 成功 ({success_count}/{copies}份)"
            
        finally:
            win32print.ClosePrinter(printer_handle)
            
    except Exception as e:
        return False, f"直接打印失败: {str(e)}"

def print_pdf_silent(filepath, printer_name, copies=1):
    """专门用于PDF文件的静默打印"""
    try:
        # 方案0: 使用Chrome/Edge浏览器的命令行打印
        browser_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        ]
        
        for browser_path in browser_paths:
            if os.path.exists(browser_path):
                try:
                    for i in range(copies):
                        # 使用浏览器的--print-to-pdf-no-header和打印机参数
                        cmd = [browser_path, '--headless', '--disable-gpu', '--print-to-printer', 
                               f'--printer-name={printer_name}', filepath]
                        result = subprocess.run(cmd, 
                                              capture_output=True, 
                                              creationflags=subprocess.CREATE_NO_WINDOW,
                                              timeout=30)
                        time.sleep(2)
                    return True, f"浏览器PDF静默打印已发送到 {printer_name} ({copies}份)"
                except subprocess.TimeoutExpired:
                    print(f"浏览器打印超时，尝试其他方案")
                except Exception as e:
                    print(f"浏览器打印失败: {e}")
        
        # 方案1: 使用Adobe Reader的命令行静默打印
        adobe_paths = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
        ]
        
        adobe_found = False
        for adobe_path in adobe_paths:
            if os.path.exists(adobe_path):
                adobe_found = True
                try:
                    for i in range(copies):
                        # 使用/t参数进行静默打印
                        cmd = [adobe_path, '/t', filepath, printer_name]
                        result = subprocess.run(cmd, 
                                              capture_output=True, 
                                              creationflags=subprocess.CREATE_NO_WINDOW,
                                              timeout=30)  # 30秒超时
                        time.sleep(2)  # 给打印机时间处理
                    return True, f"Adobe PDF静默打印已发送到 {printer_name} ({copies}份)"
                except subprocess.TimeoutExpired:
                    return False, "Adobe Reader打印超时"
                except Exception as e:
                    print(f"Adobe Reader打印失败: {e}")
                    break
        
        if not adobe_found:
            # 方案2: 使用默认PDF阅读器
            try:
                for i in range(copies):
                    # 使用 printto 动词直接打印到指定打印机，不显示对话框
                    result = win32api.ShellExecute(0, 'printto', filepath, f'"{printer_name}"', '', win32con.SW_HIDE)
                    if result <= 32:
                        raise Exception(f"ShellExecute失败，错误代码: {result}")
                    time.sleep(3)  # 给应用程序更多时间
                return True, f"默认PDF阅读器打印已发送到 {printer_name} ({copies}份)"
            except Exception as e:
                print(f"默认PDF阅读器打印失败: {e}")
        
        # 在所有基于应用程序的打印方案都失败后，尝试将 PDF 转为高质量位图并逐页发送（兼容老旧打印机）
        try:
            print("尝试将 PDF 转为高质量位图并逐页发送到打印机作为回退方案")
            bmp_pages = convert_pdf_to_bmp_pages(filepath, dpi=300)
            if bmp_pages:
                sent = send_pdf_pages_to_printer(printer_name, bmp_pages, copies)
                if sent:
                    return True, f"已通过多页位图回退方案发送到 {printer_name} ({copies}份)"
                else:
                    print("通过多页位图发送到打印机失败，继续其他回退方案")
        except Exception as e:
            print(f"多页位图回退方案出错: {e}")

        # 最终备用方案
        return print_file_silent_fallback(filepath, printer_name, copies)
        
    except Exception as e:
        print(f"PDF打印完全失败: {e}")
        return print_file_silent_fallback(filepath, printer_name, copies)

def print_text_file_simple(filepath, printer_name, copies=1):
    """改进的TXT文件打印：支持各种记事本软件创建的文件"""
    try:
        print(f" 开始打印文本文件: {filepath}")
        
        # 检测是否为远程桌面环境
        is_remote_desktop = detect_remote_desktop()
        
        # 方案1: 优先使用直接API打印 (无页码，纯文本)
        print("尝试使用直接API打印...")
        try:
            api_success = print_text_direct_to_printer(filepath, printer_name, copies)
            if api_success[0]:
                return api_success
        except Exception as e:
            print(f"直接API打印失败: {e}")
        
        # 方案2: 使用ShellExecute printto (远程桌面环境下优先使用)
        if is_remote_desktop:
            print(" 远程桌面环境，优先使用printto...")
        else:
            print("尝试使用默认程序printto...")
        sent = 0
        for i in range(copies):
            r = win32api.ShellExecute(0, 'printto', filepath, f'"{printer_name}"', None, 0)
            if r > 32:
                sent += 1
                time.sleep(1)
            else:
                # 回退到普通print
                r2 = win32api.ShellExecute(0, 'print', filepath, None, None, 0)
                if r2 > 32:
                    sent += 1
        
        # 方案3: 尝试WordPad打印 (支持更多编码，但可能有格式)
        if not is_remote_desktop:  # 远程桌面环境下跳过GUI应用
            print("尝试使用WordPad打印...")
            wordpad_success = try_wordpad_print(filepath, printer_name, copies)
            if wordpad_success[0]:
                return wordpad_success
        
        # 方案4: 最后使用记事本打印 (会产生页码)
        if not is_remote_desktop:  # 远程桌面环境下跳过GUI应用
            print(" 尝试使用Windows记事本打印(可能有页码)...")
            notepad_success = try_notepad_print(filepath, printer_name, copies)
            if notepad_success[0]:
                return notepad_success
        
        # 所有方案都失败
        return False, f"所有TXT打印方案都失败，无法发送到指定打印机 {printer_name}"
        
    except Exception as e:
        return False, f"TXT文件打印失败: {e}"

def try_notepad_print(filepath, printer_name, copies=1):
    """使用Windows自带记事本进行打印"""
    try:
        import subprocess
        notepad_path = r"C:\Windows\System32\notepad.exe"
        
        if not os.path.exists(notepad_path):
            return False, "Windows记事本未找到"
        
        success_count = 0
        for i in range(copies):
            try:
                # 使用记事本的打印功能
                # 注意：记事本没有直接的命令行打印参数，所以我们使用printto
                cmd = [notepad_path, '/p', filepath]
                result = subprocess.run(cmd, 
                                      creationflags=subprocess.CREATE_NO_WINDOW,
                                      timeout=30)
                
                # 由于notepad /p 会显示打印对话框，我们改用ShellExecute方式
                # 让系统调用notepad进行printto操作
                
                # 创建临时批处理文件来实现记事本静默打印
                temp_bat = create_notepad_print_batch(filepath, printer_name)
                if temp_bat:
                    bat_result = subprocess.run([temp_bat], 
                                              creationflags=subprocess.CREATE_NO_WINDOW,
                                              timeout=30)
                    if bat_result.returncode == 0:
                        success_count += 1
                    
                    # 清理临时文件
                    try:
                        os.remove(temp_bat)
                    except:
                        pass
                else:
                    # 如果批处理创建失败，使用备用方法
                    r = win32api.ShellExecute(0, 'open', notepad_path, f'/pt "{filepath}" "{printer_name}"', None, 0)
                    if r > 32:
                        success_count += 1
                
                time.sleep(1)
                
            except Exception as e:
                print(f"记事本打印第{i+1}份时出错: {e}")
                continue
        
        if success_count > 0:
            return True, f"Windows记事本打印成功 ({success_count}/{copies}份)"
        else:
            return False, "Windows记事本打印失败"
            
    except Exception as e:
        return False, f"记事本打印异常: {e}"

def try_wordpad_print(filepath, printer_name, copies=1):
    """使用WordPad进行打印（更好的编码支持）"""
    try:
        import subprocess
        wordpad_path = r"C:\Program Files\Windows NT\Accessories\wordpad.exe"
        
        # 64位系统的备用路径
        if not os.path.exists(wordpad_path):
            wordpad_path = r"C:\Program Files (x86)\Windows NT\Accessories\wordpad.exe"
        
        if not os.path.exists(wordpad_path):
            return False, "WordPad未找到"
        
        success_count = 0
        for i in range(copies):
            try:
                # WordPad支持 /pt 参数进行静默打印
                cmd = [wordpad_path, '/pt', filepath, printer_name]
                result = subprocess.run(cmd, 
                                      creationflags=subprocess.CREATE_NO_WINDOW,
                                      timeout=30)
                
                if result.returncode == 0:
                    success_count += 1
                    time.sleep(2)
                else:
                    print(f"WordPad打印返回代码: {result.returncode}")
                    
            except Exception as e:
                print(f"WordPad打印第{i+1}份时出错: {e}")
                continue
        
        if success_count > 0:
            return True, f"WordPad打印成功 ({success_count}/{copies}份)"
        else:
            return False, "WordPad打印失败"
            
    except Exception as e:
        return False, f"WordPad打印异常: {e}"

def read_text_with_encoding_detection(filepath):
    """智能检测文件编码并读取内容，兼容各种记事本软件"""
    try:
        # 常见的编码顺序，按优先级排序
        encodings_to_try = [
            'utf-8-sig',  # UTF-8 with BOM (Notepad3 常用)
            'utf-8',      # UTF-8 无BOM
            'gbk',        # 中文GBK
            'gb2312',     # 中文GB2312
            'cp1252',     # Windows-1252
            'latin1',     # ISO-8859-1
            'utf-16',     # UTF-16
            'utf-16le',   # UTF-16 LE
            'utf-16be'    # UTF-16 BE
        ]
        
        # 方法1: 尝试使用chardet检测编码
        try:
            with open(filepath, 'rb') as f:
                raw_data = f.read()
            
            # 使用chardet检测编码（如果可用）
            try:
                import chardet
                detected = chardet.detect(raw_data)
                if detected and detected['encoding'] and detected['confidence'] > 0.7:
                    detected_encoding = detected['encoding']
                    print(f"检测到编码: {detected_encoding} (置信度: {detected['confidence']:.2f})")
                    
                    # 将检测到的编码放在首位尝试
                    if detected_encoding not in encodings_to_try:
                        encodings_to_try.insert(0, detected_encoding)
                    else:
                        # 将检测到的编码移到首位
                        encodings_to_try.remove(detected_encoding)
                        encodings_to_try.insert(0, detected_encoding)
            except ImportError:
                print("chardet库未安装，使用默认编码顺序")
            
            # 尝试各种编码
            for encoding in encodings_to_try:
                try:
                    content = raw_data.decode(encoding)
                    print(f" 成功使用编码 {encoding} 读取文件")
                    
                    # 验证内容是否合理（不包含太多替换字符）
                    replacement_ratio = content.count('�') / len(content) if len(content) > 0 else 0
                    if replacement_ratio < 0.1:  # 替换字符少于10%
                        return content
                    else:
                        print(f"编码 {encoding} 包含过多替换字符，尝试其他编码")
                        continue
                        
                except (UnicodeDecodeError, UnicodeError) as e:
                    print(f"编码 {encoding} 失败: {e}")
                    continue
            
            # 如果所有编码都失败，使用错误处理模式
            print("️ 所有编码尝试失败，使用错误替换模式")
            return raw_data.decode('utf-8', errors='replace')
            
        except Exception as e:
            print(f"文件读取异常: {e}")
            return None
            
    except Exception as e:
        print(f"编码检测过程异常: {e}")
        return None

def create_notepad_print_batch(filepath, printer_name):
    """创建临时批处理文件实现记事本静默打印"""
    try:
        import tempfile
        
        # 创建临时批处理文件
        temp_dir = tempfile.gettempdir()
        bat_file = os.path.join(temp_dir, f"print_text_{int(time.time())}.bat")
        
        # 批处理内容：使用type命令直接发送到打印机
        bat_content = f'''@echo off
echo 正在打印文件到 {printer_name}...
type "{filepath}" > "\\\\localhost\\{printer_name}"
if errorlevel 1 (
    echo 打印失败
    exit /b 1
) else (
    echo 打印成功
    exit /b 0
)
'''
        
        with open(bat_file, 'w', encoding='gbk') as f:
            f.write(bat_content)
        
        return bat_file
        
    except Exception as e:
        print(f"创建批处理文件失败: {e}")
        return None

def print_image_silent(filepath, printer_name, copies=1):
    """专门用于图片文件的静默打印 - 使用Windows画图程序"""
    try:
        for i in range(copies):
            try:
                # 使用 Windows 自带画图程序 (mspaint.exe) 打印图片
                # /p 参数表示直接打印到默认打印机
                result = subprocess.run(
                    ['mspaint.exe', '/p', filepath],
                    capture_output=True,
                    timeout=30
                )
                if result.returncode != 0:
                    print(f"图片打印第{i+1}份失败，返回码: {result.returncode}")
                time.sleep(2)
            except subprocess.TimeoutExpired:
                print(f"图片打印第{i+1}份超时")
                return False, f"图片打印超时"
            except Exception as e:
                print(f"图片打印异常: {e}")
                return False, f"图片打印异常: {e}"
        
        return True, f"图片已发送到打印机 ({copies}份)"
    except Exception as e:
        return False, f"图片打印失败: {e}"
def print_office_silent(filepath, printer_name, copies=1):
    """简化的Office文档打印 - 使用最可靠的通用方案
    
    采用与扫描功能相同的简化模式：
    - 直接调用系统工具，不检测
    - 用线程实现超时控制
    - 用户友好的反馈
    """
    import threading
    
    try:
        file_ext = os.path.splitext(filepath)[1].lower()
        print(f"开始Office打印: {os.path.basename(filepath)}")
        
        # 方案1: 使用COM对象直接打印（最通用）
        def try_com_print():
            try:
                import win32com.client
                
                # 根据文件类型选择应用
                if file_ext in ['.doc', '.docx']:
                    app = win32com.client.Dispatch('Word.Application')
                    app.Visible = False
                    doc = app.Documents.Open(filepath, FileName=filepath)
                    
                    # 设置打印参数
                    for i in range(copies):
                        doc.PrintOut(PrintToFile=False, OutputFileName='', Printer=printer_name)
                        time.sleep(2)
                    
                    doc.Close(SaveChanges=False)
                    app.Quit()
                    return True, f"Word文档已打印 ({copies}份)"
                
                elif file_ext in ['.xls', '.xlsx']:
                    app = win32com.client.Dispatch('Excel.Application')
                    app.Visible = False
                    wb = app.Workbooks.Open(filepath)
                    
                    # 设置打印参数
                    for i in range(copies):
                        wb.PrintOut(PrintToFile=False, PrToFileName='', Printer=printer_name)
                        time.sleep(2)
                    
                    wb.Close(SaveChanges=False)
                    app.Quit()
                    return True, f"Excel表格已打印 ({copies}份)"
                
                elif file_ext in ['.ppt', '.pptx']:
                    app = win32com.client.Dispatch('PowerPoint.Application')
                    app.Visible = False
                    pres = app.Presentations.Open(filepath)
                    
                    # 设置打印参数
                    for i in range(copies):
                        pres.PrintOut(PrintToFile=False, OutputFileName='', Printer=printer_name)
                        time.sleep(3)
                    
                    pres.Close()
                    app.Quit()
                    return True, f"PowerPoint演示已打印 ({copies}份)"
                
                return False, "不支持的Office文件类型"
            
            except Exception as e:
                print(f"COM方案失败: {e}")
                return False, str(e)
        
        # 用线程执行COM打印，设置超时防止卡死
        result_holder = [False, ""]
        
        def run_print():
            result_holder[:] = try_com_print()
        
        thread = threading.Thread(target=run_print, daemon=True)
        thread.start()
        thread.join(timeout=60)  # 60秒超时
        
        if result_holder[0]:
            return result_holder
        
        print(f"✗ COM打印失败或超时，回退: {result_holder[1]}")
        
        # 方案2: 使用系统PrintTo（通用方案）
        try:
            print("正在使用系统PrintTo打印...")
            for i in range(copies):
                result = win32api.ShellExecute(
                    0, 'printto', filepath, f'"{printer_name}"', '', win32con.SW_HIDE
                )
                if result > 32:
                    print(f"第{i+1}份PrintTo成功")
                    time.sleep(3)
                else:
                    print(f"✗ 第{i+1}份PrintTo失败: {result}")
            
            return True, f"文档已通过系统PrintTo打印 ({copies}份)"
        
        except Exception as e:
            print(f"✗ PrintTo方案失败: {e}")
        
        # 方案3: 转换为PDF后打印
        try:
            print("正在转换为PDF打印...")
            pdf_path = filepath + '.print_temp.pdf'
            
            # 简单的PDF转换（使用win32print的PdfPrinter或系统转换）
            if file_ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
                # 尝试使用Microsoft Print to PDF
                result = win32api.ShellExecute(
                    0, 'print', filepath, '', '', win32con.SW_HIDE
                )
                
                if result > 32:
                    return True, f"文档已使用系统默认方式打印 ({copies}份)"
        
        except Exception as e:
            print(f"✗ 转换方案失败: {e}")
        
        # 所有方案都失败
        return False, (
            f"无法打印Office文件: {os.path.basename(filepath)}\n\n"
            "已尝试的方案：\n"
            "1. COM对象直接打印\n"
            "2. 系统PrintTo\n"
            "3. PDF转换打印\n\n"
            "建议：\n"
            "• 检查Office应用程序是否正常安装\n"
            "• 检查打印机连接\n"
            "• 尝试手动打开文件后打印\n"
            "• 将文件转换为PDF格式再打印"
        )
    
    except Exception as e:
        print(f"✗ Office打印异常: {e}")
        return False, f"Office打印异常: {str(e)}"



def print_html_silent(filepath, printer_name, copies=1):
    """专门用于HTML文件的静默打印"""
    try:
        # 方案1: 使用Internet Explorer的静默打印
        for i in range(copies):
            cmd = [
                'rundll32.exe', 
                'mshtml.dll,PrintHTML', 
                filepath
            ]
            subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
        return True, f"HTML静默打印已发送到 {printer_name} ({copies}份)"
        ps_script_direct = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 直接使用系统的PrintTo功能
    Write-Host "开始PowerPoint系统打印 {copies} 份..."
    
    for ($i = 1; $i -le {copies}; $i++) {{
        try {{
            # 方法1: 使用COM Automation但不依赖PowerPoint应用程序
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace((Get-Item "{abs_filepath.replace(chr(92), chr(92)+chr(92))}").DirectoryName)
            $item = $folder.ParseName((Get-Item "{abs_filepath.replace(chr(92), chr(92)+chr(92))}").Name)
            
            # 获取打印动词
            $verbs = $item.Verbs()
            $printVerb = $verbs | Where-Object {{ $_.Name -match "打印|Print" }}
            
            if ($printVerb) {{
                $printVerb.DoIt()
                Write-Host "PowerPoint系统打印第${{i}}份已发送"
                Start-Sleep -Seconds 3
            }} else {{
                # 备用方法：使用Start-Process PrintTo
                Start-Process -FilePath "{abs_filepath.replace(chr(92), chr(92)+chr(92))}" -Verb PrintTo -ArgumentList "{printer_name}" -WindowStyle Hidden -Wait
                Write-Host "PowerPoint PrintTo第${{i}}份完成"
                Start-Sleep -Seconds 2
            }}
        }} catch {{
            Write-Host "PowerPoint系统打印第${{i}}份失败: $_"
            
            # 最后备用：直接文件关联打印
            try {{
                cmd /c 'print /d:"{printer_name}" "{abs_filepath.replace(chr(92), chr(92)+chr(92))}"'
                Write-Host "PowerPoint命令行打印第${{i}}份完成"
            }} catch {{
                Write-Host "PowerPoint所有打印方法都失败了"
            }}
        }}
    }}
    
    Write-Output "PowerPoint直接打印完成"
}} catch {{
    Write-Host "PowerPoint直接打印失败: $_"
    exit 1
}}
'''
        
        # 尝试PowerShell直接打印
        try:
            result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_direct],
                                  capture_output=True, text=True, timeout=120,
                                  creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                print(" PowerPoint直接打印成功")
                return True, f"PowerPoint直接打印完成 ({copies}份)"
            else:
                print(f"PowerPoint直接打印stderr: {result.stderr}")
        except Exception as e:
            print(f"PowerPoint直接打印异常: {e}")
        
        # 方案2: 使用win32api直接打印
        print(" 尝试Win32API直接打印...")
        try:
            for i in range(copies):
                # 使用Windows API直接打印
                result = win32api.ShellExecute(0, 'printto', abs_filepath, f'"{printer_name}"', None, 0)
                if result > 32:  # ShellExecute成功返回值 > 32
                    print(f" PowerPoint Win32API打印第{i+1}份成功")
                    time.sleep(3)
                else:
                    print(f" PowerPoint Win32API打印第{i+1}份失败，返回码: {result}")
            
            return True, f"PowerPoint Win32API打印完成 ({copies}份)"
        except Exception as e:
            print(f"PowerPoint Win32API打印异常: {e}")
        
        # 方案3: 调用Office/WPS静默转换为PDF打印（最可靠的备用方案）
        print(" 尝试Office/WPS静默转换PDF打印...")
        try:
            temp_dir = tempfile.gettempdir()
            temp_pdf_path = os.path.join(temp_dir, f"ppt_convert_{int(time.time() * 1000)}.pdf")
            
            # 首先尝试 Microsoft PowerPoint COM 转换
            print(" 尝试Microsoft PowerPoint转换PDF...")
            ps_script_ppt_pdf = f'''
try {{
    $ErrorActionPreference = "Stop"
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $false
    $ppt.DisplayAlerts = 0
    Write-Host "PowerPoint COM创建成功"
    
    $presentation = $ppt.Presentations.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "演示文稿打开成功"
    
    # PowerPoint SaveAs PDF (格式号32)
    $presentation.SaveAs("{temp_pdf_path.replace(chr(92), chr(92)+chr(92))}", 32)
    Write-Host "PDF转换完成: {temp_pdf_path.replace(chr(92), chr(92)+chr(92))}"
    
    $presentation.Close()
    $ppt.Quit()
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    
    Write-Output "PowerPoint PDF转换成功"
}} catch {{
    Write-Host "PowerPoint PDF转换失败: $_"
    if ($ppt) {{
        try {{ 
            if ($presentation) {{ $presentation.Close() }}
            $ppt.Quit()
        }} catch {{}}
    }}
    exit 1
}}
'''
            
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_ppt_pdf],
                                      capture_output=True, text=True, timeout=60,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0 and os.path.exists(temp_pdf_path):
                    print(f" Microsoft PowerPoint PDF转换成功: {temp_pdf_path}")
                    
                    # 打印PDF
                    for i in range(copies):
                        try:
                            print_result = win32api.ShellExecute(0, 'printto', temp_pdf_path, f'"{printer_name}"', None, 0)
                            if print_result > 32:
                                print(f" PowerPoint PDF转换打印第{i+1}份成功")
                            else:
                                print(f" PowerPoint PDF转换打印第{i+1}份失败，返回码: {print_result}")
                            time.sleep(2)
                        except Exception as e:
                            print(f" 打印第{i+1}份PDF失败: {e}")
                    
                    # 清理临时PDF
                    try:
                        time.sleep(1)
                        os.remove(temp_pdf_path)
                        print(" 临时PDF已清理")
                    except:
                        pass
                    
                    return True, f"PowerPoint COM转换PDF打印完成 ({copies}份)"
            except subprocess.TimeoutExpired:
                print(" Microsoft PowerPoint PDF转换超时")
            except Exception as e:
                print(f" Microsoft PowerPoint PDF转换异常: {e}")
            
            # 尝试 WPS Presentation COM 转换
            print(" 尝试WPS Presentation转换PDF...")
            ps_script_wps_pdf = f'''
try {{
    $ErrorActionPreference = "Stop"
    $wpp = New-Object -ComObject wpp.application
    $wpp.Visible = $false
    Write-Host "WPS Presentation COM创建成功"
    
    $presentation = $wpp.Presentations.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "WPS演示文稿打开成功"
    
    # WPS 使用 ExportAsFixedFormat 方法 (格式2=PDF)
    $presentation.ExportAsFixedFormat("{temp_pdf_path.replace(chr(92), chr(92)+chr(92))}", 2)
    Write-Host "WPS PDF转换完成: {temp_pdf_path.replace(chr(92), chr(92)+chr(92))}"
    
    $presentation.Close()
    $wpp.Quit()
    
    Write-Output "WPS PDF转换成功"
}} catch {{
    Write-Host "WPS Presentation PDF转换失败: $_"
    if ($wpp) {{
        try {{ 
            if ($presentation) {{ $presentation.Close() }}
            $wpp.Quit()
        }} catch {{}}
    }}
    exit 1
}}
'''
            
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_wps_pdf],
                                      capture_output=True, text=True, timeout=60,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0 and os.path.exists(temp_pdf_path):
                    print(f" WPS Presentation PDF转换成功: {temp_pdf_path}")
                    
                    # 打印PDF
                    for i in range(copies):
                        try:
                            print_result = win32api.ShellExecute(0, 'printto', temp_pdf_path, f'"{printer_name}"', None, 0)
                            if print_result > 32:
                                print(f" WPS PDF转换打印第{i+1}份成功")
                            else:
                                print(f" WPS PDF转换打印第{i+1}份失败，返回码: {print_result}")
                            time.sleep(2)
                        except Exception as e:
                            print(f" 打印第{i+1}份PDF失败: {e}")
                    
                    # 清理临时PDF
                    try:
                        time.sleep(1)
                        os.remove(temp_pdf_path)
                        print(" 临时PDF已清理")
                    except:
                        pass
                    
                    return True, f"WPS COM转换PDF打印完成 ({copies}份)"
            except subprocess.TimeoutExpired:
                print(" WPS Presentation PDF转换超时")
            except Exception as e:
                print(f" WPS Presentation PDF转换异常: {e}")
            
            # 如果COM转换都失败，尝试 LibreOffice 作为备选
            print(" 尝试LibreOffice转换PDF...")
            libreoffice_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            
            for lo_path in libreoffice_paths:
                if os.path.exists(lo_path):
                    print(f" 找到LibreOffice: {lo_path}")
                    
                    try:
                        # LibreOffice命令行转换PDF
                        convert_cmd = [lo_path, '--headless', '--convert-to', 'pdf', '--outdir', temp_dir, abs_filepath]
                        convert_result = subprocess.run(convert_cmd, capture_output=True, 
                                                     creationflags=subprocess.CREATE_NO_WINDOW, timeout=60)
                        
                        if convert_result.returncode == 0:
                            # LibreOffice生成的PDF文件名
                            base_name = os.path.splitext(os.path.basename(abs_filepath))[0]
                            lo_generated_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
                            
                            if os.path.exists(lo_generated_pdf):
                                print(f" LibreOffice PDF转换成功: {lo_generated_pdf}")
                                
                                # 打印PDF
                                for i in range(copies):
                                    try:
                                        print_result = win32api.ShellExecute(0, 'printto', lo_generated_pdf, f'"{printer_name}"', None, 0)
                                        if print_result > 32:
                                            print(f" LibreOffice PDF打印第{i+1}份成功")
                                        else:
                                            print(f" LibreOffice PDF打印第{i+1}份失败，返回码: {print_result}")
                                        time.sleep(2)
                                    except Exception as e:
                                        print(f" 打印第{i+1}份PDF失败: {e}")
                                
                                # 清理临时PDF
                                try:
                                    time.sleep(1)
                                    os.remove(lo_generated_pdf)
                                    print(" 临时PDF已清理")
                                except:
                                    pass
                                
                                return True, f"LibreOffice转换PDF打印完成 ({copies}份)"
                    except Exception as e:
                        print(f" LibreOffice转换异常: {e}")
            
            print(" 未找到Office/WPS/LibreOffice，无法进行PDF转换")
            
        except Exception as e:
            print(f" PDF转换打印异常: {e}")
        
        print(" PowerPoint所有直接打印方法都失败")
        return False, "PowerPoint直接打印失败：所有方法都无法成功"
        
    except Exception as e:
        print(f"PowerPoint直接打印函数异常: {e}")
        return False, f"PowerPoint直接打印异常: {str(e)}"


def print_office_com(filepath, printer_name, copies, file_ext):
    """强化Office COM打印 - 支持Microsoft Office和WPS Office"""
    try:
        abs_filepath = os.path.abspath(filepath)
        print(f" 强化Office COM打印: {abs_filepath}")
        
        if file_ext in ['.doc', '.docx']:
            # Word文档 - 优先Microsoft Word，回退WPS Writer
            print(" 强化Word COM打印...")
            
            # 方案1: Microsoft Word - 智能页面边距优化
            print(" 尝试Microsoft Word...")
            ps_script_word = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 创建Microsoft Word COM对象
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = $false
    Write-Host "Microsoft Word COM创建成功"
    
    # 打开文档
    $doc = $word.Documents.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "Word文档打开成功"
    
    # 智能页面设置优化
    try {{
        $pageSetup = $doc.Range().PageSetup
        Write-Host "获取页面设置成功"
        
        # 检查是否有页眉页脚内容（保留用户设置）
        $hasHeader = $false
        $hasFooter = $false
        
        try {{
            if ($doc.Sections.Count -gt 0) {{
                $section = $doc.Sections.Item(1)
                $headerText = $section.Headers.Item(1).Range.Text
                $footerText = $section.Footers.Item(1).Range.Text
                
                # 检查是否有实际内容（不只是空白字符）
                if ($headerText -and $headerText.Trim().Length -gt 1) {{
                    $hasHeader = $true
                    Write-Host "检测到页眉内容，将保留"
                }}
                
                if ($footerText -and $footerText.Trim().Length -gt 1) {{
                    $hasFooter = $true
                    Write-Host "检测到页脚内容，将保留"
                }}
            }}
        }} catch {{
            Write-Host "页眉页脚检测失败，使用保守设置: $_"
        }}
        
        # 备份原始边距
        $originalTopMargin = $pageSetup.TopMargin
        $originalBottomMargin = $pageSetup.BottomMargin
        $originalLeftMargin = $pageSetup.LeftMargin
        $originalRightMargin = $pageSetup.RightMargin
        $originalHeaderDistance = $pageSetup.HeaderDistance
        $originalFooterDistance = $pageSetup.FooterDistance
        
        Write-Host "原始边距 - 上: $originalTopMargin, 下: $originalBottomMargin, 左: $originalLeftMargin, 右: $originalRightMargin"
        Write-Host "页眉距离: $originalHeaderDistance, 页脚距离: $originalFooterDistance"
        
        # 优化边距设置（单位：磅，1英寸=72磅）
        $optimizedTopMargin = 36      # 0.5英寸
        $optimizedBottomMargin = 36   # 0.5英寸
        $optimizedLeftMargin = 54     # 0.75英寸
        $optimizedRightMargin = 54    # 0.75英寸
        
        # 如果有页眉，保留足够空间
        if ($hasHeader) {{
            $optimizedTopMargin = [math]::Max($originalTopMargin, 54)  # 至少0.75英寸
            $pageSetup.HeaderDistance = [math]::Min($originalHeaderDistance, 18)  # 页眉距离优化
            Write-Host "保留页眉空间，顶部边距: $optimizedTopMargin"
        }} else {{
            # 没有页眉时，最小化顶部边距，最大化打印区域
            $optimizedTopMargin = 18  # 仅0.25英寸的最小边距
            $pageSetup.HeaderDistance = 0  # 无页眉时不需要页眉距离
            Write-Host "无页眉，最小化顶部边距: $optimizedTopMargin"
        }}
        
        # 如果有页脚，保留足够空间
        if ($hasFooter) {{
            $optimizedBottomMargin = [math]::Max($originalBottomMargin, 54)  # 至少0.75英寸
            $pageSetup.FooterDistance = [math]::Min($originalFooterDistance, 18)  # 页脚距离优化
            Write-Host "保留页脚空间，底部边距: $optimizedBottomMargin"
        }} else {{
            # 没有页脚时，最小化底部边距，最大化打印区域
            $optimizedBottomMargin = 18  # 仅0.25英寸的最小边距
            $pageSetup.FooterDistance = 0  # 无页脚时不需要页脚距离
            Write-Host "无页脚，最小化底部边距: $optimizedBottomMargin"
        }}
        
        # 应用优化边距
        $pageSetup.TopMargin = $optimizedTopMargin
        $pageSetup.BottomMargin = $optimizedBottomMargin
        $pageSetup.LeftMargin = $optimizedLeftMargin
        $pageSetup.RightMargin = $optimizedRightMargin
        
        Write-Host "已优化页面边距 - 上: $optimizedTopMargin, 下: $optimizedBottomMargin, 左: $optimizedLeftMargin, 右: $optimizedRightMargin"
        Write-Host "页眉状态: $hasHeader, 页脚状态: $hasFooter"
        
    }} catch {{
        Write-Host "页面设置优化失败，使用默认设置: $_"
    }}
    
    # 设置打印机
    try {{
        $word.ActivePrinter = "{printer_name}"
        Write-Host "Word打印机设置成功: {printer_name}"
    }} catch {{
        Write-Host "Word无法设置指定打印机，使用默认"
    }}
    
    # 执行打印 - 使用更可靠的参数
    Write-Host "开始Word打印 {copies} 份..."
    for ($i = 1; $i -le {copies}; $i++) {{
        try {{
            # PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)
            $doc.PrintOut([ref]$false, [ref]$false, [ref]0, [ref]"", [ref]1, [ref]($doc.Range().End), [ref]7, [ref]1)
            Write-Host "Word打印第${{i}}份完成"
            Start-Sleep -Seconds 2
        }} catch {{
            Write-Host "Word打印第${{i}}份失败: $_"
        }}
    }}
    
    $doc.Close([ref]$false)
    $word.Quit()
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Output "Microsoft Word打印成功"
}} catch {{
    Write-Host "Microsoft Word打印失败: $_"
    if ($word) {{
        try {{ $word.Quit() }} catch {{}}
    }}
    exit 1
}}
'''
            
            # 尝试Microsoft Word
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_word],
                                      capture_output=True, text=True, timeout=60,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    print(" Microsoft Word COM打印成功")
                    return True, f"Microsoft Word COM打印完成 ({copies}份)"
                else:
                    print(f" Microsoft Word COM失败: {result.stderr[:200] if result.stderr else '无错误信息'}")
            except subprocess.TimeoutExpired:
                print(" Microsoft Word COM超时")
            except Exception as e:
                print(f"Microsoft Word COM异常: {e}")
            
            # 方案2: WPS Writer - 智能页面边距优化
            print(" 尝试WPS Writer...")
            ps_script_wps_writer = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 创建WPS Writer COM对象
    $wps = New-Object -ComObject wps.application
    $wps.Visible = $false
    $wps.DisplayAlerts = $false
    Write-Host "WPS Writer COM创建成功"
    
    $doc = $null
    try {{
        # 打开文档
        $doc = $wps.Documents.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
        Write-Host "WPS Writer文档打开成功"
        
        # WPS Writer智能页面设置优化
        try {{
            $pageSetup = $doc.Range().PageSetup
            Write-Host "获取WPS页面设置成功"
            
            # 检查WPS页眉页脚内容
            $hasHeader = $false
            $hasFooter = $false
            
            try {{
                if ($doc.Sections.Count -gt 0) {{
                    $section = $doc.Sections.Item(1)
                    
                    # WPS页眉检查
                    try {{
                        $headerText = $section.Headers.Item(1).Range.Text
                        if ($headerText -and $headerText.Trim().Length -gt 1) {{
                            $hasHeader = $true
                            Write-Host "WPS检测到页眉内容"
                        }}
                    }} catch {{ }}
                    
                    # WPS页脚检查
                    try {{
                        $footerText = $section.Footers.Item(1).Range.Text
                        if ($footerText -and $footerText.Trim().Length -gt 1) {{
                            $hasFooter = $true
                            Write-Host "WPS检测到页脚内容"
                        }}
                    }} catch {{ }}
                }}
            }} catch {{
                Write-Host "WPS页眉页脚检测异常: $_"
            }}
            
            # WPS边距优化（磅为单位）
            $wpsOptTopMargin = 36      # 0.5英寸
            $wpsOptBottomMargin = 36   # 0.5英寸
            $wpsOptLeftMargin = 54     # 0.75英寸
            $wpsOptRightMargin = 54    # 0.75英寸
            
            # WPS页眉处理
            if ($hasHeader) {{
                $wpsOptTopMargin = 72      # 1英寸，为页眉保留空间
                try {{ $pageSetup.HeaderDistance = 18 }} catch {{ }}
            }} else {{
                try {{ $pageSetup.HeaderDistance = 18 }} catch {{ }}
            }}
            
            # WPS页脚处理 - 重点优化
            if ($hasFooter) {{
                $wpsOptBottomMargin = 72   # 1英寸，为页脚保留空间
                try {{ $pageSetup.FooterDistance = 18 }} catch {{ }}
                Write-Host "WPS保留页脚空间"
            }} else {{
                # 没有页脚时，大幅减少底部边距，释放预留空间
                $wpsOptBottomMargin = 30   # 约0.42英寸，最小化底部空白
                try {{ $pageSetup.FooterDistance = 12 }} catch {{ }}
                Write-Host "WPS优化底部边距，释放页脚空间"
            }}
            
            # 应用WPS优化边距
            try {{
                $pageSetup.TopMargin = $wpsOptTopMargin
                $pageSetup.BottomMargin = $wpsOptBottomMargin
                $pageSetup.LeftMargin = $wpsOptLeftMargin
                $pageSetup.RightMargin = $wpsOptRightMargin
                Write-Host "WPS页面边距已优化 - 上: $wpsOptTopMargin, 下: $wpsOptBottomMargin"
            }} catch {{
                Write-Host "WPS边距设置失败: $_"
            }}
            
        }} catch {{
            Write-Host "WPS页面优化失败，使用默认: $_"
        }}
        
        # 设置打印机
        try {{
            $wps.ActivePrinter = "{printer_name}"
            Write-Host "WPS Writer打印机设置成功"
        }} catch {{
            Write-Host "WPS Writer无法设置指定打印机，使用默认"
        }}
        
        # 执行打印
        for ($i = 1; $i -le {copies}; $i++) {{
            try {{
                $doc.PrintOut()
                Write-Host "WPS Writer打印第${{i}}份完成"
                Start-Sleep -Seconds 2
            }} catch {{
                Write-Host "WPS Writer打印第${{i}}份失败: $_"
            }}
        }}
    }} finally {{
        # 确保资源清理
        if ($doc) {{
            try {{
                $doc.Close([ref]$false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            }} catch {{
                Write-Host "清理文档对象失败: $_"
            }}
        }}
        
        if ($wps) {{
            try {{
                $wps.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wps) | Out-Null
            }} catch {{
                Write-Host "清理WPS对象失败: $_"
            }}
        }}
        
        # 强制垃圾回收
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }}
    
    Write-Output "WPS Writer打印成功"
}} catch {{
    Write-Host "WPS Writer打印失败: $_"
    exit 1
}}
'''
            
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_wps_writer],
                                      capture_output=True, text=True, timeout=45,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    print(" WPS Writer COM打印成功")
                    return True, f"WPS Writer COM打印完成 ({copies}份)"
                else:
                    print(f" WPS Writer COM失败: {result.stderr[:200] if result.stderr else '无错误信息'}")
            except subprocess.TimeoutExpired:
                print(" WPS Writer COM超时")
            except Exception as e:
                print(f"WPS Writer COM异常: {e}")
            
            print(" Word文档COM打印失败")
            return False, "Word COM对象均不可用"
        elif file_ext in ['.xls', '.xlsx']:
            # Excel文档 - 优先Microsoft Excel，回退WPS Spreadsheets
            print(" 强化Excel COM打印...")
            
            # 方案1: Microsoft Excel - 智能页面边距优化
            print(" 尝试Microsoft Excel...")
            ps_script_excel = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 创建Microsoft Excel COM对象
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Microsoft Excel COM创建成功"
    
    # 打开工作簿
    $workbook = $excel.Workbooks.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "Excel工作簿打开成功"
    
    # Excel智能页面设置优化
    try {{
        if ($workbook.Worksheets.Count -gt 0) {{
            $worksheet = $workbook.Worksheets.Item(1)
            $pageSetup = $worksheet.PageSetup
            Write-Host "获取Excel页面设置成功"
            
            # 检查Excel页眉页脚内容
            $hasHeader = $false
            $hasFooter = $false
            
            try {{
                # Excel页眉检查
                if ($pageSetup.CenterHeader -or $pageSetup.LeftHeader -or $pageSetup.RightHeader) {{
                    $headerContent = ($pageSetup.CenterHeader + $pageSetup.LeftHeader + $pageSetup.RightHeader).Trim()
                    if ($headerContent.Length -gt 0) {{
                        $hasHeader = $true
                        Write-Host "Excel检测到页眉内容"
                    }}
                }}
                
                # Excel页脚检查
                if ($pageSetup.CenterFooter -or $pageSetup.LeftFooter -or $pageSetup.RightFooter) {{
                    $footerContent = ($pageSetup.CenterFooter + $pageSetup.LeftFooter + $pageSetup.RightFooter).Trim()
                    if ($footerContent.Length -gt 0) {{
                        $hasFooter = $true
                        Write-Host "Excel检测到页脚内容"
                    }}
                }}
            }} catch {{
                Write-Host "Excel页眉页脚检测异常: $_"
            }}
            
            # Excel边距优化（英寸为单位）
            $excelOptLeftMargin = 0.75    # 0.75英寸，左右边距保持适中
            $excelOptRightMargin = 0.75   # 0.75英寸
            
            # Excel页眉处理
            if ($hasHeader) {{
                $excelOptTopMargin = 1.0      # 1英寸，为页眉保留空间
                Write-Host "Excel保留页眉空间"
            }} else {{
                # 没有页眉时，最小化顶部边距
                $excelOptTopMargin = 0.25     # 0.25英寸，最小化顶部空白
                Write-Host "Excel优化顶部边距，释放页眉空间"
            }}
            
            # Excel页脚处理 - 重点优化
            if ($hasFooter) {{
                $excelOptBottomMargin = 1.0   # 1英寸，为页脚保留空间
                Write-Host "Excel保留页脚空间"
            }} else {{
                # 没有页脚时，最小化底部边距
                $excelOptBottomMargin = 0.25  # 0.25英寸，最小化底部空白
                Write-Host "Excel优化底部边距，释放页脚空间"
            }}
            
            # 应用Excel优化边距
            try {{
                $pageSetup.TopMargin = $excel.Application.InchesToPoints($excelOptTopMargin)
                $pageSetup.BottomMargin = $excel.Application.InchesToPoints($excelOptBottomMargin)
                $pageSetup.LeftMargin = $excel.Application.InchesToPoints($excelOptLeftMargin)
                $pageSetup.RightMargin = $excel.Application.InchesToPoints($excelOptRightMargin)
                Write-Host "Excel页面边距已优化 - 上: $excelOptTopMargin, 下: $excelOptBottomMargin"
            }} catch {{
                Write-Host "Excel边距设置失败: $_"
            }}
        }}
    }} catch {{
        Write-Host "Excel页面优化失败，使用默认: $_"
    }}
    
    # 设置打印机
    try {{
        $excel.ActivePrinter = "{printer_name}"
        Write-Host "Excel打印机设置成功: {printer_name}"
    }} catch {{
        Write-Host "Excel无法设置指定打印机，使用默认"
    }}
    
    # 执行打印 - 使用更精确的参数
    Write-Host "开始Excel打印 {copies} 份..."
    for ($i = 1; $i -le {copies}; $i++) {{
        try {{
            # PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName)
            $workbook.PrintOut([Type]::Missing, [Type]::Missing, 1, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing)
            Write-Host "Excel打印第${{i}}份完成"
            Start-Sleep -Seconds 2
        }} catch {{
            Write-Host "Excel打印第${{i}}份失败: $_"
        }}
    }}
    
    $workbook.Close([ref]$false)
    $excel.Quit()
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Output "Microsoft Excel打印成功"
}} catch {{
    Write-Host "Microsoft Excel打印失败: $_"
    if ($excel) {{
        try {{ $excel.Quit() }} catch {{}}
    }}
    exit 1
}}
'''
            
            # 尝试Microsoft Excel
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_excel],
                                      capture_output=True, text=True, timeout=60,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    print(" Microsoft Excel COM打印成功")
                    return True, f"Microsoft Excel COM打印完成 ({copies}份)"
            except Exception as e:
                print(f"Microsoft Excel COM异常: {e}")
            
            # 方案2: WPS Spreadsheets - 智能页面边距优化
            print(" 尝试WPS Spreadsheets...")
            ps_script_wps_excel = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 创建WPS Spreadsheets COM对象
    $et = New-Object -ComObject et.application
    $et.Visible = $false
    Write-Host "WPS Spreadsheets COM创建成功"
    
    # 打开工作簿
    $workbook = $et.Workbooks.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "WPS Spreadsheets工作簿打开成功"
    
    # WPS Spreadsheets智能页面设置优化
    try {{
        if ($workbook.Worksheets.Count -gt 0) {{
            $worksheet = $workbook.Worksheets.Item(1)
            $pageSetup = $worksheet.PageSetup
            Write-Host "获取WPS Spreadsheets页面设置成功"
            
            # 检查WPS页眉页脚内容
            $hasHeader = $false
            $hasFooter = $false
            
            try {{
                # WPS页眉检查
                if ($pageSetup.CenterHeader -or $pageSetup.LeftHeader -or $pageSetup.RightHeader) {{
                    $headerContent = ($pageSetup.CenterHeader + $pageSetup.LeftHeader + $pageSetup.RightHeader).Trim()
                    if ($headerContent.Length -gt 0) {{
                        $hasHeader = $true
                        Write-Host "WPS Spreadsheets检测到页眉内容"
                    }}
                }}
                
                # WPS页脚检查
                if ($pageSetup.CenterFooter -or $pageSetup.LeftFooter -or $pageSetup.RightFooter) {{
                    $footerContent = ($pageSetup.CenterFooter + $pageSetup.LeftFooter + $pageSetup.RightFooter).Trim()
                    if ($footerContent.Length -gt 0) {{
                        $hasFooter = $true
                        Write-Host "WPS Spreadsheets检测到页脚内容"
                    }}
                }}
            }} catch {{
                Write-Host "WPS页眉页脚检测异常: $_"
            }}
            
            # WPS边距优化（磅为单位）
            $wpsOptTopMargin = 36      # 0.5英寸
            $wpsOptBottomMargin = 36   # 0.5英寸
            $wpsOptLeftMargin = 54     # 0.75英寸
            $wpsOptRightMargin = 54    # 0.75英寸
            
            # WPS页眉处理
            if ($hasHeader) {{
                $wpsOptTopMargin = 72      # 1英寸，为页眉保留空间
            }}
            
            # WPS页脚处理 - 重点优化
            if ($hasFooter) {{
                $wpsOptBottomMargin = 72   # 1英寸，为页脚保留空间
                Write-Host "WPS Spreadsheets保留页脚空间"
            }} else {{
                # 没有页脚时，大幅减少底部边距，释放预留空间
                $wpsOptBottomMargin = 30   # 约0.42英寸，最小化底部空白
                Write-Host "WPS Spreadsheets优化底部边距，释放页脚空间"
            }}
            
            # 应用WPS优化边距
            try {{
                $pageSetup.TopMargin = $wpsOptTopMargin
                $pageSetup.BottomMargin = $wpsOptBottomMargin
                $pageSetup.LeftMargin = $wpsOptLeftMargin
                $pageSetup.RightMargin = $wpsOptRightMargin
                Write-Host "WPS Spreadsheets页面边距已优化 - 上: $wpsOptTopMargin, 下: $wpsOptBottomMargin"
            }} catch {{
                Write-Host "WPS边距设置失败: $_"
            }}
        }}
    }} catch {{
        Write-Host "WPS Spreadsheets页面优化失败，使用默认: $_"
    }}
    
    # 设置打印机
    try {{
        $et.ActivePrinter = "{printer_name}"
        Write-Host "WPS Spreadsheets打印机设置成功"
    }} catch {{
        Write-Host "WPS Spreadsheets无法设置指定打印机，使用默认"
    }}
    
    # 执行打印
    for ($i = 1; $i -le {copies}; $i++) {{
        try {{
            $workbook.PrintOut()
            Write-Host "WPS Spreadsheets打印第${{i}}份完成"
            Start-Sleep -Seconds 2
        }} catch {{
            Write-Host "WPS Spreadsheets打印第${{i}}份失败: $_"
        }}
    }}
    
    $workbook.Close([ref]$false)
    $et.Quit()
    Write-Output "WPS Spreadsheets打印成功"
}} catch {{
    Write-Host "WPS Spreadsheets打印失败: $_"
    if ($et) {{
        try {{ $et.Quit() }} catch {{}}
    }}
    exit 1
}}
'''
            
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_wps_excel],
                                      capture_output=True, text=True, timeout=45,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    print(" WPS Spreadsheets COM打印成功")
                    return True, f"WPS Spreadsheets COM打印完成 ({copies}份)"
            except Exception as e:
                print(f"WPS Spreadsheets COM异常: {e}")
            
            print(" Excel文档COM打印失败")
            return False, "Excel COM对象均不可用"
        elif file_ext in ['.ppt', '.pptx']:
            # PowerPoint文档 - 优先Microsoft PowerPoint，回退WPS Presentation
            print("️ 强化PowerPoint COM打印...")
            
            # 方案1: Microsoft PowerPoint - 智能页面边距优化
            print(" 尝试Microsoft PowerPoint...")
            ps_script_ppt = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 创建Microsoft PowerPoint COM对象，增加权限处理
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $false
    $ppt.DisplayAlerts = 0  # 禁用所有警告
    Start-Sleep -Seconds 1
    Write-Host "Microsoft PowerPoint COM创建成功"
    
    # 打开演示文稿 - 简化参数
    $presentation = $ppt.Presentations.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "PowerPoint演示文稿打开成功"
    
    # PowerPoint智能页面设置优化
    try {{
        if ($presentation.Slides.Count -gt 0) {{
            Write-Host "开始PowerPoint页面设置优化..."
            
            # 检查PowerPoint母版中的页眉页脚
            $hasHeader = $false
            $hasFooter = $false
            
            try {{
                $slideMaster = $presentation.SlideMaster
                
                # 检查母版的页眉页脚设置
                if ($slideMaster.HeadersFooters.Header.Visible -and $slideMaster.HeadersFooters.Header.Text.Trim().Length -gt 0) {{
                    $hasHeader = $true
                    Write-Host "PowerPoint检测到页眉内容"
                }}
                
                if ($slideMaster.HeadersFooters.Footer.Visible -and $slideMaster.HeadersFooters.Footer.Text.Trim().Length -gt 0) {{
                    $hasFooter = $true
                    Write-Host "PowerPoint检测到页脚内容"
                }}
                
                # 检查页码设置
                if ($slideMaster.HeadersFooters.SlideNumber.Visible) {{
                    Write-Host "PowerPoint检测到页码显示"
                    # 页码通常在底部，按页脚处理
                    $hasFooter = $true
                }}
            }} catch {{
                Write-Host "PowerPoint母版检测异常: $_"
            }}
            
            # PowerPoint页面设置优化
            try {{
                $pageSetup = $presentation.PageSetup
                
                # PowerPoint边距优化（磅为单位）
                # PowerPoint的页面设置相对简单，主要是幻灯片尺寸
                if (-not $hasFooter) {{
                    # 如果没有页脚/页码，尝试优化底部空间
                    Write-Host "PowerPoint优化：无页脚内容，最大化内容空间"
                    
                    # 尝试调整幻灯片的内容边距（如果支持）
                    # PowerPoint的边距控制相对有限，主要通过母版调整
                    try {{
                        # 设置为更紧凑的页面设置
                        Write-Host "PowerPoint设置紧凑打印模式"
                    }} catch {{
                        Write-Host "PowerPoint边距调整有限"
                    }}
                }} else {{
                    Write-Host "PowerPoint保留页眉页脚空间"
                }}
            }} catch {{
                Write-Host "PowerPoint页面设置异常: $_"
            }}
        }}
    }} catch {{
        Write-Host "PowerPoint页面优化失败，使用默认: $_"
    }}
    
    # 设置打印机
    try {{
        $ppt.ActivePrinter = "{printer_name}"
        Write-Host "PowerPoint打印机设置成功: {printer_name}"
    }} catch {{
        Write-Host "PowerPoint无法设置指定打印机，使用默认"
    }}
    
    # 执行打印 - 先尝试PDF转换方案（更可靠）
    Write-Host "开始PowerPoint打印 {copies} 份..."
    for ($i = 1; $i -le {copies}; $i++) {{
        $success = $false
        
        # 优先使用PDF转换方案（更稳定）
        try {{
            Write-Host "尝试PowerPoint PDF转换打印第${{i}}份..."
            $tempPdf = "$env:TEMP\\ppt_print_${{i}}.pdf"
            # 导出为PDF
            $presentation.SaveAs($tempPdf, 32)  # 32 = ppSaveAsPDF
            Write-Host "PDF导出成功: $tempPdf"
            
            # 打印PDF
            # 使用正确的PrintTo动词指定打印机
            Start-Process -FilePath $tempPdf -Verb PrintTo -ArgumentList "{printer_name}" -WindowStyle Hidden -Wait
            Start-Sleep -Seconds 2
            
            # 清理临时文件
            if (Test-Path $tempPdf) {{
                Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue
            }}
            
            Write-Host "PowerPoint第${{i}}份PDF转换打印成功"
            $success = $true
        }} catch {{
            Write-Host "PowerPoint PDF转换第${{i}}份失败: $_"
        }}
        
        # 如果PDF方案失败，尝试直接打印
        if (-not $success) {{
            try {{
                Write-Host "尝试PowerPoint直接打印第${{i}}份..."
                $presentation.PrintOut()
                Write-Host "PowerPoint直接打印第${{i}}份成功"
                Start-Sleep -Seconds 3
            }} catch {{
                Write-Host "PowerPoint直接打印第${{i}}份也失败: $_"
                throw "所有PowerPoint打印方法都失败了"
            }}
        }}
    }}
    
    $presentation.Close()
    $ppt.Quit()
    
    # 清理COM对象
    try {{
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }} catch {{
        # 忽略清理错误
    }}
    
    Write-Output "Microsoft PowerPoint打印成功"
}} catch {{
    Write-Host "Microsoft PowerPoint打印失败: $_"
    if ($ppt) {{
        try {{
            if ($presentation) {{ $presentation.Close() }}
            $ppt.Quit()
        }} catch {{}}
    }}
    exit 1
}}
'''
            
            # 尝试Microsoft PowerPoint
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_ppt],
                                      capture_output=True, text=True, timeout=90,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    print(" Microsoft PowerPoint COM打印成功")
                    return True, f"Microsoft PowerPoint COM打印完成 ({copies}份)"
            except Exception as e:
                print(f"Microsoft PowerPoint COM异常: {e}")
            
            # 方案2: WPS Presentation - 智能页面边距优化
            print(" 尝试WPS Presentation...")
            ps_script_wps_ppt = f'''
try {{
    $ErrorActionPreference = "Stop"
    
    # 创建WPS Presentation COM对象
    $wpp = New-Object -ComObject wpp.application
    $wpp.Visible = $false
    Write-Host "WPS Presentation COM创建成功"
    
    # 打开演示文稿
    $presentation = $wpp.Presentations.Open("{abs_filepath.replace(chr(92), chr(92)+chr(92))}")
    Write-Host "WPS Presentation文档打开成功"
    
    # WPS Presentation智能页面设置优化
    try {{
        if ($presentation.Slides.Count -gt 0) {{
            Write-Host "开始WPS Presentation页面设置优化..."
            
            # 检查WPS Presentation母版中的页眉页脚
            $hasHeader = $false
            $hasFooter = $false
            
            try {{
                $slideMaster = $presentation.SlideMaster
                
                # 检查母版的页眉页脚设置
                if ($slideMaster.HeadersFooters.Header.Visible -and $slideMaster.HeadersFooters.Header.Text.Trim().Length -gt 0) {{
                    $hasHeader = $true
                    Write-Host "WPS Presentation检测到页眉内容"
                }}
                
                if ($slideMaster.HeadersFooters.Footer.Visible -and $slideMaster.HeadersFooters.Footer.Text.Trim().Length -gt 0) {{
                    $hasFooter = $true
                    Write-Host "WPS Presentation检测到页脚内容"
                }}
                
                # 检查页码设置
                if ($slideMaster.HeadersFooters.SlideNumber.Visible) {{
                    Write-Host "WPS Presentation检测到页码显示"
                    # 页码通常在底部，按页脚处理
                    $hasFooter = $true
                }}
            }} catch {{
                Write-Host "WPS Presentation母版检测异常: $_"
            }}
            
            # WPS Presentation页面设置优化
            try {{
                if (-not $hasFooter) {{
                    # 如果没有页脚/页码，尝试优化底部空间
                    Write-Host "WPS Presentation优化：无页脚内容，最大化内容空间"
                    
                    # WPS的演示文稿边距控制相对有限
                    try {{
                        Write-Host "WPS Presentation设置紧凑打印模式"
                        # WPS特定的打印优化设置可以在这里添加
                    }} catch {{
                        Write-Host "WPS Presentation边距调整有限"
                    }}
                }} else {{
                    Write-Host "WPS Presentation保留页眉页脚空间"
                }}
            }} catch {{
                Write-Host "WPS Presentation页面设置异常: $_"
            }}
        }}
    }} catch {{
        Write-Host "WPS Presentation页面优化失败，使用默认: $_"
    }}
    
    # 设置打印机
    try {{
        $wpp.ActivePrinter = "{printer_name}"
        Write-Host "WPS Presentation打印机设置成功"
    }} catch {{
        Write-Host "WPS Presentation无法设置指定打印机，使用默认"
    }}
    
    # 执行打印
    for ($i = 1; $i -le {copies}; $i++) {{
        try {{
            $presentation.PrintOut()
            Write-Host "WPS Presentation打印第${{i}}份完成"
            Start-Sleep -Seconds 3
        }} catch {{
            Write-Host "WPS Presentation直接打印第${{i}}份失败，尝试PDF转换: $_"
            try {{
                $tempPdf = "$env:TEMP\\wps_ppt_temp_${{i}}.pdf"
                $presentation.ExportAsFixedFormat($tempPdf, 2)  # 2=PDF格式
                # 使用正确的PrintTo动词指定打印机
                Start-Process -FilePath $tempPdf -Verb PrintTo -ArgumentList "{printer_name}" -WindowStyle Hidden -Wait
                Start-Sleep -Seconds 2
                Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue
                Write-Host "WPS Presentation第${{i}}份PDF转换打印成功"
            }} catch {{
                Write-Host "WPS Presentation第${{i}}份PDF转换失败: $_"
            }}
        }}
    }}
    
    $presentation.Close()
    $wpp.Quit()
    Write-Output "WPS Presentation打印成功"
}} catch {{
    Write-Host "WPS Presentation打印失败: $_"
    if ($wpp) {{
        try {{ $wpp.Quit() }} catch {{}}
    }}
    exit 1
}}
'''
            
            try:
                result = subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script_wps_ppt],
                                      capture_output=True, text=True, timeout=75,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    print(" WPS Presentation COM打印成功")
                    return True, f"WPS Presentation COM打印完成 ({copies}份)"
            except Exception as e:
                print(f"WPS Presentation COM异常: {e}")
            
            print(" PowerPoint文档COM打印失败")
            return False, f"""PowerPoint COM打印失败详情：

已尝试的COM方案：
1. Microsoft PowerPoint COM (包含PDF转换备用)
2. WPS Presentation COM (包含PDF转换备用)

可能原因：
- PowerPoint/WPS未正确安装或注册
- COM对象权限不足
- 文档格式损坏或不兼容
- 打印机驱动问题

建议解决方案：
1. 手动打开 {os.path.basename(filepath)} 测试是否正常
2. 尝试"打印到PDF"测试COM功能
3. 重新注册Office COM: regsvr32 /i pptcore.dll
4. 以管理员权限运行打印服务"""
        else:
            print(" 不支持的Office文档类型")
            return print_file_silent_fallback(filepath, printer_name, copies)
        
    except Exception as e:
        print(f"Office COM打印整体异常: {e}")
        return False, f"Office COM异常: {str(e)}"

def print_html_silent(filepath, printer_name, copies=1):
    """专门用于HTML文件的静默打印"""
    try:
        # 方案1: 使用Internet Explorer的静默打印
        for i in range(copies):
            cmd = [
                'rundll32.exe', 
                'mshtml.dll,PrintHTML', 
                filepath
            ]
            subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
        return True, f"HTML静默打印已发送到 {printer_name} ({copies}份)"
        
    except Exception as e1:
        try:
            # 方案2: 使用PowerShell和Internet Explorer COM对象
            ps_script = f'''
try {{
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Visible = $false
    $ie.Navigate("file:///{filepath.replace(chr(92), '/')}")
    while ($ie.Busy) {{ Start-Sleep -Milliseconds 100 }}
    for ($i = 1; $i -le {copies}; $i++) {{
        $ie.ExecWB(6, 2)  # 静默打印
    }}
    $ie.Quit()
    Write-Host "HTML打印完成"
}} catch {{
    Write-Host "HTML打印失败： $_"
}}
'''
            subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script],
                          creationflags=subprocess.CREATE_NO_WINDOW)
            return True, f"HTML PowerShell静默打印已执行 ({copies}份)"
            
        except Exception as e2:
            # 备用方案
            return print_file_silent_fallback(filepath, printer_name, copies)

def get_printer_capabilities(printer_name):
    """获取指定打印机的功能参数（原始返回）
    返回结构:
    {
        'duplex_support': bool,
        'color_support': bool,
        'papers': [{'id': int, 'name': str}],
        'resolutions': ['600x600', ...],
        'printer_status': str,
        'driver_name': str,
        'port_name': str
    }
    """
    try:
        print(f"正在获取打印机 '{printer_name}' 的实际参数...")
        
        # 检查打印机名称是否有效
        if not printer_name or printer_name.strip() == "" or printer_name == "未检测到可用打印机":
            print("打印机名称无效")
            return {
                'duplex_support': False,
                'color_support': False,
                'paper_sizes': ['A4'],
                'quality_levels': ['normal'],
                'printer_status': '离线或不可用',
                'driver_name': '未知'
            }
        
        # 尝试打开打印机并获取其属性
        printer_handle = win32print.OpenPrinter(printer_name)
        
        try:
            # 获取打印机信息
            printer_info = win32print.GetPrinter(printer_handle, 2)
            driver_name = printer_info.get('pDriverName', '未知')
            port_name = printer_info.get('pPortName', '未知')
            status = printer_info.get('Status', 0)
            
            print(f"打印机驱动: {driver_name}")
            print(f"打印机端口: {port_name}")
            print(f"打印机状态码: {status}")
            
            # 解析打印机状态
            printer_status = '在线'
            if status != 0:
                status_descriptions = {
                    0x00000001: '暂停',
                    0x00000002: '错误',
                    0x00000004: '正在删除',
                    0x00000008: '缺纸',
                    0x00000010: '缺纸',
                    0x00000020: '手动送纸',
                    0x00000040: '纸张故障',
                    0x00000080: '离线',
                    0x00000100: 'I/O 活动',
                    0x00000200: '忙',
                    0x00000400: '正在打印',
                    0x00000800: '输出槽满',
                    0x00001000: '不可用',
                    0x00002000: '等待',
                    0x00004000: '正在处理',
                    0x00008000: '正在初始化',
                    0x00010000: '正在预热',
                    0x00020000: '碳粉不足',
                    0x00040000: '没有碳粉',
                    0x00080000: '页面错误',
                    0x00100000: '用户干预',
                    0x00200000: '内存不足',
                    0x00400000: '门打开'
                }
                # 找到最符合的状态描述
                for status_bit, description in status_descriptions.items():
                    if status & status_bit:
                        printer_status = description
                        break
                else:
                    printer_status = f'未知状态 ({status})'
            
            # 获取设备功能，使用自定义的常量
            duplex_support = False
            color_support = False
            papers = []  # [{'id': int, 'name': str}]
            resolutions_list = []  # ['600x600']
            
            try:
                # 检查双面打印支持
                try:
                    duplex_caps = win32print.DeviceCapabilities(printer_name, port_name, DC_DUPLEX, None)
                    # 修复：双面打印支持应该检查 > 0，不只是 == 1
                    # duplex_caps 的含义：
                    # 0 = 不支持双面打印
                    # 1 = 支持双面打印 (仅长边翻转)
                    # 2 = 支持双面打印 (仅短边翻转) 
                    # 3 = 支持双面打印 (长边和短边都支持)
                    duplex_support = duplex_caps > 0
                    duplex_modes = []
                    if duplex_caps >= 1:
                        duplex_modes.append("long_edge")  # 长边翻转
                    if duplex_caps >= 2:
                        duplex_modes.append("short_edge")  # 短边翻转
                    
                    print(f"双面打印支持: {duplex_support} (DeviceCapabilities返回: {duplex_caps})")
                    if duplex_modes:
                        print(f"支持的双面模式: {', '.join(duplex_modes)}")
                except Exception as e:
                    print(f"检查双面打印支持失败: {e}")
                    duplex_support = False
                    duplex_modes = []
                
                # 检查颜色支持
                try:
                    color_caps = win32print.DeviceCapabilities(printer_name, port_name, DC_COLORDEVICE, None)
                    color_support = color_caps == 1
                    print(f"颜色打印支持: {color_support} (DeviceCapabilities返回: {color_caps})")
                except Exception as e:
                    print(f"检查颜色支持失败: {e}")
                    color_support = False
                
                # 获取支持的纸张（ID+名称）
                try:
                    paper_ids = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERS, None)
                    paper_names = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERNAMES, None)
                    if paper_ids and paper_names:
                        # DC_PAPERNAMES 通常返回每个名称固定长度的字节/字符串数组
                        # pywin32 会解码为字符串元组，名称末尾可能含有\x00
                        count = min(len(paper_ids), len(paper_names))
                        for i in range(count):
                            pid = paper_ids[i]
                            pname = paper_names[i]
                            if isinstance(pname, bytes):
                                try:
                                    pname = pname.decode('mbcs', errors='ignore')
                                except Exception:
                                    pname = str(pname)
                            pname = pname.replace('\x00', '').strip()
                            if pname:
                                papers.append({'id': int(pid), 'name': pname})
                        print(f"纸张(原始): {papers[:8]}{' ...' if len(papers)>8 else ''}")
                    else:
                        print("未获取到纸张列表")
                except Exception as e:
                    print(f"获取纸张列表失败: {e}")
                
                # 获取打印分辨率（原始DPI列表）
                try:
                    resolutions = win32print.DeviceCapabilities(printer_name, port_name, DC_ENUMRESOLUTIONS, None)
                    if resolutions:
                        for res in resolutions:
                            # pywin32 分辨率项通常为 dict 或 tuple，包含 xdpi/ydpi
                            if isinstance(res, dict):
                                xdpi = res.get('xdpi') or res.get('X') or 0
                                ydpi = res.get('ydpi') or res.get('Y') or 0
                            elif isinstance(res, (tuple, list)) and len(res) >= 2:
                                xdpi, ydpi = res[0], res[1]
                            else:
                                continue
                            if xdpi and ydpi:
                                resolutions_list.append(f"{xdpi}x{ydpi}")
                        print(f"分辨率(原始): {resolutions_list}")
                    else:
                        print("未获取到分辨率列表")
                except Exception as e:
                    print(f"获取分辨率失败: {e}")
                
            except Exception as e:
                print(f"获取设备功能时出错: {e}")
            
            capabilities = {
                'duplex_support': duplex_support,
                'duplex_modes': duplex_modes if 'duplex_modes' in locals() else [],
                'color_support': color_support,
                'papers': papers,
                'resolutions': resolutions_list,
                'printer_status': printer_status,
                'driver_name': driver_name,
                'port_name': port_name
            }
            
            print(f"最终获取的打印机参数: {capabilities}")
            return capabilities
            
        finally:
            win32print.ClosePrinter(printer_handle)
            
    except Exception as e:
        print(f"无法访问打印机 '{printer_name}': {e}")
        # 返回默认功能，表示打印机不可用
        return {
            'duplex_support': False,
            'duplex_modes': [],
            'color_support': False,
            'papers': [],
            'resolutions': [],
            'printer_status': '离线或不可用',
            'driver_name': '未知',
            'port_name': ''
        }
 
def get_logs():
    if not os.path.exists(LOG_FILE):
        return []
    with open(LOG_FILE, 'r', encoding='utf-8') as f:
        return f.readlines()[-50:][::-1]

def clean_old_logs():
    """清理旧的打印日志，保留最近的记录"""
    try:
        if not os.path.exists(LOG_FILE):
            return
            
        # 检查日志文件大小
        file_size = os.path.getsize(LOG_FILE)
        max_size = 5 * 1024 * 1024  # 5MB限制
        
        if file_size > max_size:
            print(f"日志文件过大({file_size/1024/1024:.1f}MB)，开始清理...")
            
            # 读取所有日志
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # 保留最近1000条记录
            keep_lines = 1000
            if len(lines) > keep_lines:
                # 保留最新的记录
                new_lines = lines[-keep_lines:]
                
                # 写回文件
                with open(LOG_FILE, 'w', encoding='utf-8') as f:
                    f.writelines(new_lines)
                
                removed_count = len(lines) - keep_lines
                print(f" 日志清理完成，删除了 {removed_count} 条旧记录，保留最新 {keep_lines} 条")
    
    except Exception as e:
        print(f"日志清理失败: {e}")

def clean_old_logs_by_date():
    """按日期清理日志，删除7天前的记录"""
    try:
        if not os.path.exists(LOG_FILE):
            return
            
        from datetime import datetime, timedelta
        
        cutoff_date = datetime.now() - timedelta(days=7)  # 7天前
        
        # 读取日志并过滤
        new_lines = []
        removed_count = 0
        
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                try:
                    # 尝试解析日志中的时间戳
                    if line.strip():
                        # 假设日志格式：2025-10-03 18:42:32 客户端: ...
                        date_str = line[:19]  # 提取前19个字符的日期时间
                        log_date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                        
                        if log_date >= cutoff_date:
                            new_lines.append(line)
                        else:
                            removed_count += 1
                except:
                    # 如果解析失败，保留这行（可能是格式异常的日志）
                    new_lines.append(line)
        
        if removed_count > 0:
            # 写回文件
            with open(LOG_FILE, 'w', encoding='utf-8') as f:
                f.writelines(new_lines)
            
            print(f"按日期清理日志完成，删除了 {removed_count} 条7天前的记录")
    
    except Exception as e:
        print(f"按日期清理日志失败: {e}")

def periodic_log_cleanup():
    """定期日志清理任务"""
    import time
    while True:
        try:
            # 每天检查一次日志大小
            time.sleep(86400)  # 24小时
            
            from datetime import datetime
            current_hour = datetime.now().hour
            
                 # 下午3点执行清理任务
            if current_hour == 15:
                print(" 执行定时日志清理...")
                # 先按大小清理
                clean_old_logs()
                # 再按日期清理
                clean_old_logs_by_date()
            else:
                # 如果不是下午3点，等到下午3点再执行
                import datetime as dt
                now = dt.datetime.now()
                target_time = now.replace(hour=15, minute=0, second=0, microsecond=0)
                if now.hour >= 15:
                    target_time += dt.timedelta(days=1)  # 明天下午3点
                
                sleep_seconds = (target_time - now).total_seconds()
                print(f" 日志清理将在 {target_time.strftime('%Y-%m-%d %H:%M')} 执行")
                time.sleep(sleep_seconds)
           
        except Exception as e:
            print(f"定期日志清理异常: {e}")
            time.sleep(3600)  # 出错后1小时再试

@app.route('/health')
def health_check():
    """健康检查端点"""
    try:
        # 更新服务管理器的健康检查状态
        service_manager.update_health_check()
        
        # 计算正确的运行时间
        uptime = 0
        if hasattr(service_manager, 'start_time') and service_manager.start_time:
            uptime = time.time() - service_manager.start_time
        
        # 返回基本的服务状态信息
        return jsonify({
            'status': 'healthy',
            'timestamp': time.time(),
            'service_running': service_manager.service_running,
            'uptime': uptime
        }), 200
        
    except Exception as e:
        return jsonify({
            'status': 'unhealthy',
            'error': str(e),
            'timestamp': time.time()
        }), 500

@app.route('/api/printer_info')
def get_printer_info_api():
    """API端点：获取指定打印机的信息"""
    try:
        printer_name = request.args.get('printer')
        if not printer_name:
            return jsonify({'success': False, 'error': '未指定打印机名称'})
        
        capabilities = get_printer_capabilities(printer_name)
        return jsonify({
            'success': True,
            'capabilities': capabilities
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@app.route('/api/delete_file', methods=['POST'])
def delete_file_api():
    """API端点：删除队列中的单个文件"""
    try:
        print(f" 收到删除文件请求")
        print(f"   Content-Type: {request.content_type}")
        print(f"   Headers: {dict(request.headers)}")
        print(f"   Method: {request.method}")
        
        # 尝试解析JSON数据
        try:
            data = request.get_json()
            print(f"   解析的JSON数据: {data}")
        except Exception as json_error:
            print(f" JSON解析错误: {json_error}")
            return jsonify({'success': False, 'error': f'JSON解析错误: {str(json_error)}'})
        
        if not data or 'filename' not in data:
            print(f" 请求数据无效: {data}")
            return jsonify({'success': False, 'error': '未提供文件名'})
        
        filename = data['filename']
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        print(f"准备删除文件: {filepath}")
        
        # 检查文件是否存在
        if not os.path.exists(filepath):
            print(f"文件不存在: {filepath}")
            return jsonify({'success': False, 'error': '文件不存在或已被删除'})
        
        # 尝试取消相关的打印任务（在删除文件之前）
        cancel_result = {'cancelled': [], 'skipped': [], 'total_found': 0}
        try:
            print(f" 检查是否有相关的打印任务需要取消...")
            
            # 默认不取消正在打印的任务，除非用户显式要求
            force_cancel = request.get_json().get('force_cancel_active', False) if request.get_json() else False
            
            cancel_result = cancel_print_jobs_by_document(filename, cancel_active=force_cancel)
            
            cancelled_count = len(cancel_result['cancelled'])
            skipped_count = len(cancel_result['skipped'])
            
            if cancelled_count > 0:
                print(f" 已取消 {cancelled_count} 个打印任务")
            if skipped_count > 0:
                print(f"[SKIP] 跳过 {skipped_count} 个任务（正在打印或已完成）")
            if cancel_result['total_found'] == 0:
                print(f" 未找到相关的打印任务")
                
        except Exception as cancel_error:
            print(f"️ 取消打印任务失败: {cancel_error}")
        
        # 删除文件
        os.remove(filepath)
        print(f"文件删除成功: {filepath}")
        
        # 记录删除日志
        try:
            client_ip = request.remote_addr or '未知IP'
            cancelled_count = len(cancel_result['cancelled'])
            cancelled_info = f", 取消了 {cancelled_count} 个打印任务" if cancelled_count > 0 else ""
            log_message = f"{datetime.now()} 客户端: {client_ip} 删除文件: {filename}{cancelled_info}"
            with open(LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(log_message + "\n")
            print(f"日志记录成功: {log_message}")
        except Exception as log_error:
            print(f"日志记录失败: {log_error}")
            # 即使日志记录失败，也继续返回成功
        
        # 返回结果，包含详细的打印任务信息
        response_message = f'文件 {filename} 已删除'
        
        cancelled_count = len(cancel_result['cancelled'])
        skipped_count = len(cancel_result['skipped'])
        
        if cancelled_count > 0:
            response_message += f'，取消了 {cancelled_count} 个打印任务'
        if skipped_count > 0:
            response_message += f'，跳过了 {skipped_count} 个任务（正在执行或已完成）'
        
        return jsonify({
            'success': True,
            'message': response_message,
            'print_queue_result': {
                'cancelled_jobs': cancelled_count,
                'skipped_jobs': skipped_count,
                'total_found': cancel_result['total_found'],
                'cancelled_details': cancel_result['cancelled'],
                'skipped_details': cancel_result['skipped']
            }
        })
        
    except Exception as e:
        print(f"删除文件API发生异常: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'服务器错误: {str(e)}'
        })

@app.route('/api/delete_all_files', methods=['POST'])
def delete_all_files_api():
    """API端点：清空队列中的所有文件"""
    try:
        files = os.listdir(UPLOAD_FOLDER)
        deleted_count = 0
        
        # 首先尝试清空所有打印机队列
        cleared_jobs = 0
        try:
            print(f"️ 正在清空所有打印机队列...")
            cleared_jobs = clear_all_print_queues()
            if cleared_jobs > 0:
                print(f" 已清空 {cleared_jobs} 个打印任务")
            else:
                print(f" 打印队列为空或无法访问")
        except Exception as clear_error:
            print(f"️ 清空打印队列失败: {clear_error}")
        
        # 然后删除所有文件
        for filename in files:
            try:
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                if os.path.isfile(filepath):
                    os.remove(filepath)
                    deleted_count += 1
            except Exception as e:
                print(f"删除文件 {filename} 时出错: {e}")
        
        # 记录删除日志
        client_ip = request.remote_addr or '未知IP'
        queue_info = f", 清空了 {cleared_jobs} 个打印任务" if cleared_jobs > 0 else ""
        log_message = f"{datetime.now()} 客户端: {client_ip} 清空队列: 删除了 {deleted_count} 个文件{queue_info}"
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_message + "\n")
        
        # 构建响应消息
        response_message = f'已删除 {deleted_count} 个文件'
        if cleared_jobs > 0:
            response_message += f'，清空了 {cleared_jobs} 个打印任务'
        
        return jsonify({
            'success': True,
            'count': deleted_count,
            'cleared_jobs': cleared_jobs,
            'message': response_message
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@app.route('/api/refresh_printers')
def refresh_printers_api():
    """API端点：刷新打印机列表"""
    try:
        # 刷新打印机列表
        success = refresh_printer_list()
        if success:
            default_printer = get_default_printer()
            return jsonify({
                'success': True,
                'printers': PRINTERS,
                'default_printer': default_printer,
                'message': f'已刷新，检测到 {len(PRINTERS)} 台物理打印机'
            })
        else:
            return jsonify({
                'success': False,
                'error': '刷新打印机列表失败'
            })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

# ================== 扫描功能路由 ==================

@app.route('/api/get_scanners')
def get_scanners():
    """获取可用扫描仪列表的API"""
    try:
        scanners = get_available_scanners()
        return jsonify({
            'success': True,
            'scanners': scanners,
            'count': len(scanners)
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'获取扫描仪列表失败: {str(e)}'
        })

@app.route('/api/device_status')
def get_device_status():
    """获取设备状态（打印和扫描）"""
    global DEVICE_STATUS
    
    current_time = time.time()
    
    # 检查超时状态并自动重置 - 打印极短超时（30秒）
    if DEVICE_STATUS['is_printing'] and DEVICE_STATUS['print_start_time']:
        if current_time - DEVICE_STATUS['print_start_time'] > 30:  # 30秒超时
            print("[TIMEOUT] 打印任务超时（30秒），重置状态")
            DEVICE_STATUS['is_printing'] = False
            DEVICE_STATUS['print_start_time'] = None
            DEVICE_STATUS['print_client'] = ''
    
    if DEVICE_STATUS['is_scanning'] and DEVICE_STATUS['scan_start_time']:
        if current_time - DEVICE_STATUS['scan_start_time'] > 60:  # 60秒超时
            print("[TIMEOUT] 扫描任务超时（60秒），重置状态")
            DEVICE_STATUS['is_scanning'] = False
            DEVICE_STATUS['scan_start_time'] = None
            DEVICE_STATUS['scan_client'] = ''
    
    # 计算运行时间
    print_duration = 0
    scan_duration = 0
    
    if DEVICE_STATUS['is_printing'] and DEVICE_STATUS['print_start_time']:
        print_duration = int(current_time - DEVICE_STATUS['print_start_time'])
    
    if DEVICE_STATUS['is_scanning'] and DEVICE_STATUS['scan_start_time']:
        scan_duration = int(current_time - DEVICE_STATUS['scan_start_time'])
    
    return jsonify({
        'success': True,
        'is_printing': DEVICE_STATUS['is_printing'],
        'is_scanning': DEVICE_STATUS['is_scanning'],
        'print_duration': print_duration,
        'scan_duration': scan_duration,
        'print_client': DEVICE_STATUS['print_client'],
        'scan_client': DEVICE_STATUS['scan_client']
    })

@app.route('/api/scan', methods=['POST'])
def scan_document():
    """执行扫描的API - 增加设备互斥检查"""
    global DEVICE_STATUS
    
    try:
        # 检查设备状态 - 不允许打印和扫描同时进行
        if DEVICE_STATUS['is_printing']:
            print_duration = int(time.time() - DEVICE_STATUS['print_start_time']) if DEVICE_STATUS['print_start_time'] else 0
            return jsonify({
                'success': False,
                'error': f'设备正在打印中，无法开始扫描\n\n正在打印: {print_duration}秒\n操作者: {DEVICE_STATUS["print_client"]}\n\n请等待打印完成后再试'
            })
        
        if DEVICE_STATUS['is_scanning']:
            scan_duration = int(time.time() - DEVICE_STATUS['scan_start_time']) if DEVICE_STATUS['scan_start_time'] else 0
            return jsonify({
                'success': False,
                'error': f'扫描仪正在使用中\n\n正在扫描: {scan_duration}秒\n操作者: {DEVICE_STATUS["scan_client"]}\n\n请稍后再试'
            })
        
        # 获取表单参数
        scanner_id = request.form.get('scanner_id', 'default')
        scanner_name = request.form.get('scanner_name', '默认扫描仪')
        scan_format = request.form.get('format', 'PNG').upper()
        
        # 获取客户端信息用于日志
        client_info = get_client_info()
        
        print(f" 收到扫描请求: 扫描仪={scanner_name}, 格式={scan_format}, 客户端={client_info}")
        
        # 设置扫描状态
        DEVICE_STATUS['is_scanning'] = True
        DEVICE_STATUS['scan_start_time'] = time.time()
        DEVICE_STATUS['scan_client'] = client_info
        
        try:
            # 执行扫描
            success, message = start_scan_silent(scanner_id, scanner_name, scan_format)
        finally:
            # 无论成功失败都要重置扫描状态
            DEVICE_STATUS['is_scanning'] = False
            DEVICE_STATUS['scan_start_time'] = None
            DEVICE_STATUS['scan_client'] = ''
        
        if success:
            print(f" 扫描成功: {message}")
            
            # 记录扫描日志
            log_scan(scanner_name, scan_format, client_info, message)
            
            return jsonify({
                'success': True,
                'message': message
            })
        else:
            print(f" 扫描失败: {message}")
            
            # 记录扫描失败日志
            log_scan(scanner_name, scan_format, client_info, f"扫描失败: {message}")
            
            return jsonify({
                'success': False,
                'error': message
            })
    
    except Exception as e:
        error_msg = f"扫描请求处理异常: {str(e)}"
        print(f" {error_msg}")
        
        return jsonify({
            'success': False,
            'error': error_msg
        })

@app.route('/api/scanned_files')
def get_scanned_files_api():
    """获取扫描文件列表的API"""
    try:
        files = get_scanned_files()
        return jsonify({
            'success': True,
            'files': files,
            'count': len(files)
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'获取扫描文件列表失败: {str(e)}'
        })

@app.route('/api/scanned_files/<filename>')
def download_scanned_file(filename):
    """下载扫描文件"""
    try:
        scan_folder = path_manager.get_scan_dir()
        
        # 安全检查：确保文件名不包含路径分隔符
        if '/' in filename or '\\' in filename or '..' in filename:
            return jsonify({'error': '非法文件名'}), 400
        
        # 检查文件是否存在
        file_path = os.path.join(scan_folder, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        file_ext = os.path.splitext(filename)[1].lower()
        if file_ext == '.pdf':
            return send_from_directory(scan_folder, filename, as_attachment=True, mimetype='application/pdf')
        else:
            return send_from_directory(scan_folder, filename, as_attachment=True)
        
    except Exception as e:
        return jsonify({'error': f'下载文件失败: {str(e)}'}), 500

@app.route('/api/scanned_files/<filename>/preview')
def preview_scanned_file(filename):
    """预览扫描文件（仅限图片）"""
    try:
        scan_folder = path_manager.get_scan_dir()
        
        # 安全检查
        if '/' in filename or '\\' in filename or '..' in filename:
            return jsonify({'error': '非法文件名'}), 400
        
        # 检查文件是否存在
        file_path = os.path.join(scan_folder, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        # 检查是否是图片文件
        file_ext = os.path.splitext(filename)[1].lower()
        if file_ext not in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
            return jsonify({'error': '该文件类型不支持预览'}), 400
        
        return send_from_directory(scan_folder, filename)
        
    except Exception as e:
        return jsonify({'error': f'预览文件失败: {str(e)}'}), 500

@app.route('/api/scanned_files/<filename>/print', methods=['POST'])
def print_scanned_file(filename):
    """打印扫描文件"""
    try:
        scan_folder = path_manager.get_scan_dir()
        
        # 安全检查
        if '/' in filename or '\\' in filename or '..' in filename:
            return jsonify({'success': False, 'error': '非法文件名'})
        
        # 检查文件是否存在
        file_path = os.path.join(scan_folder, filename)
        if not os.path.exists(file_path):
            return jsonify({'success': False, 'error': '文件不存在'})
        
        # 获取打印机参数
        printer_name = request.json.get('printer', get_default_printer())
        copies = int(request.json.get('copies', 1))
        
        # 获取客户端信息
        client_info = get_client_info()
        
        print(f"️ 收到扫描文件打印请求: {filename} -> {printer_name}, 客户端: {client_info}")
        
        # 执行打印
        success = print_file_with_settings(file_path, printer_name, copies)
        
        if success:
            message = f"扫描文件 {filename} 已发送到打印机 {printer_name}"
            print(f" {message}")
            
            # 记录打印日志
            with open(LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(f"{datetime.now()} 客户端: {client_info} 打印扫描文件: {filename} -> {printer_name} (份数: {copies})\n")
            
            return jsonify({
                'success': True,
                'message': message
            })
        else:
            error_msg = f"扫描文件 {filename} 打印失败"
            print(f" {error_msg}")
            return jsonify({
                'success': False,
                'error': error_msg
            })
    
    except Exception as e:
        error_msg = f"打印扫描文件异常: {str(e)}"
        print(f" {error_msg}")
        return jsonify({
            'success': False,
            'error': error_msg
        })

@app.route('/api/scanned_files/<filename>/delete', methods=['DELETE'])
def delete_scanned_file(filename):
    """删除扫描文件"""
    try:
        scan_folder = path_manager.get_scan_dir()
        
        # 安全检查
        if '/' in filename or '\\' in filename or '..' in filename:
            return jsonify({'success': False, 'error': '非法文件名'})
        
        # 检查文件是否存在
        file_path = os.path.join(scan_folder, filename)
        if not os.path.exists(file_path):
            return jsonify({'success': False, 'error': '文件不存在'})
        
        # 删除文件
        os.remove(file_path)
        
        # 获取客户端信息并记录日志
        client_info = get_client_info()
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now()} 客户端: {client_info} 删除扫描文件: {filename}\n")
        
        print(f"扫描文件已删除: {filename}")
        
        return jsonify({
            'success': True,
            'message': f'文件 {filename} 已删除'
        })
        
    except Exception as e:
        error_msg = f"删除扫描文件失败: {str(e)}"
        print(f"{error_msg}")
        return jsonify({
            'success': False,
            'error': error_msg
        })

@app.route('/api/clear_scanned_files', methods=['POST', 'OPTIONS'])
def clear_all_scanned_files():
    """清空所有扫描文件"""
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp
    
    try:
        scan_folder = path_manager.get_scan_dir()
        
        if not os.path.exists(scan_folder):
            return jsonify({
                'status': 'success',
                'deleted_count': 0,
                'message': '扫描文件夹不存在'
            })
        
        # 获取所有扫描文件
        deleted_count = 0
        for filename in os.listdir(scan_folder):
            file_path = os.path.join(scan_folder, filename)
            if os.path.isfile(file_path):
                try:
                    os.remove(file_path)
                    deleted_count += 1
                except Exception as e:
                    print(f"删除文件失败: {filename} - {e}")
        
        # 记录日志
        client_info = get_client_info()
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now()} 客户端: {client_info} 清空扫描队列: 删除了 {deleted_count} 个文件\n")
        
        print(f"扫描队列已清空: 删除了 {deleted_count} 个文件")
        
        return jsonify({
            'status': 'success',
            'deleted_count': deleted_count,
            'message': f'已清空 {deleted_count} 个扫描文件'
        })
    except Exception as e:
        error_msg = f"清空扫描队列失败: {str(e)}"
        print(f"{error_msg}")
        return jsonify({
            'status': 'error',
            'error': error_msg
        }), 500

@app.route('/api/print_queue', methods=['GET'])
def get_print_queue_api():
    """API端点：获取打印队列状态"""
    try:
        printer_name = request.args.get('printer')
        jobs = get_print_queue_jobs(printer_name)
        
        return jsonify({
            'success': True,
            'jobs': jobs,
            'count': len(jobs)
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@app.route('/api/clear_print_queue', methods=['POST'])
def clear_print_queue_api():
    """API端点：清空打印队列"""
    global DEVICE_STATUS
    
    try:
        data = request.get_json() or {}
        printer_name = data.get('printer')
        
        if printer_name:
            # 清空指定打印机的队列
            jobs = get_print_queue_jobs(printer_name)
            cleared_count = 0
            for job in jobs:
                try:
                    import win32print
                    printer_handle = win32print.OpenPrinter(printer_name)
                    win32print.SetJob(printer_handle, job['job_id'], 0, None, win32print.JOB_CONTROL_CANCEL)
                    win32print.ClosePrinter(printer_handle)
                    cleared_count += 1
                except Exception as e:
                    print(f"取消任务失败: {e}")
            
            message = f'已清空打印机 {printer_name} 的 {cleared_count} 个任务'
        else:
            # 清空所有打印机的队列
            cleared_count = clear_all_print_queues()
            message = f'已清空所有打印机的 {cleared_count} 个任务'
        
        # 如果成功清空了任何打印任务，重置全局打印状态并释放WIA设备
        if cleared_count > 0 and DEVICE_STATUS['is_printing']:
            print("[RESET] 重置打印设备状态（API清空打印队列）")
            DEVICE_STATUS['is_printing'] = False
            DEVICE_STATUS['print_start_time'] = None
            DEVICE_STATUS['print_client'] = ''
            
            # 强制清理端口占用和重启WIA服务
            print("[CLEANUP] 强制清理后台占用资源...")
            try:
                port = getattr(app, 'current_port', 5000)
                cleanup_port_and_restart_wia(port)
            except Exception as e:
                print(f"[WARN] 端口清理异常（非致命）: {e}")
            
            # 自动释放WIA设备，避免与扫描冲突
            print("[INFO] 释放WIA扫描设备以避免冲突...")
            try:
                force_release_wia_device()
            except Exception as e:
                print(f"[WARN] WIA设备释放异常（非致命）: {e}")
        
        # 记录日志
        client_ip = request.remote_addr or '未知IP'
        log_message = f"{datetime.now()} 客户端: {client_ip} 清空打印队列: {message}"
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_message + "\n")
        
        return jsonify({
            'success': True,
            'cleared_count': cleared_count,
            'message': message
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

# 系统修复工具API已移除，功能整合到托盘菜单中
 
def get_file_list():
    """获取上传文件夹中的文件列表（包含详细信息）"""
    file_list = []
    try:
        for filename in os.listdir(UPLOAD_FOLDER):
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(filepath):
                try:
                    # 获取文件信息
                    stat = os.stat(filepath)
                    file_size = stat.st_size
                    upload_time = datetime.fromtimestamp(stat.st_mtime)
                    
                    # 格式化文件大小
                    if file_size < 1024:
                        size_str = f"{file_size} B"
                    elif file_size < 1024 * 1024:
                        size_str = f"{file_size / 1024:.1f} KB"
                    else:
                        size_str = f"{file_size / (1024 * 1024):.1f} MB"
                    
                    # 获取文件扩展名
                    extension = os.path.splitext(filename)[1].lower().lstrip('.')
                    
                    file_info = {
                        'name': filename,
                        'size': file_size,
                        'size_str': size_str,
                        'upload_time': upload_time.strftime('%m-%d %H:%M'),
                        'extension': extension or 'unknown'
                    }
                    
                    file_list.append(file_info)
                except Exception as e:
                    print(f"获取文件 {filename} 信息时出错: {e}")
                    # 如果无法获取详细信息，至少保留文件名
                    file_list.append({
                        'name': filename,
                        'size': 0,
                        'size_str': 'Unknown',
                        'upload_time': 'Unknown',
                        'extension': 'unknown'
                    })
        
        # 按上传时间排序（最新的在前）
        file_list.sort(key=lambda x: x['upload_time'], reverse=True)
    except Exception as e:
        print(f"获取文件列表时出错: {e}")
    
    return file_list

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    files = get_file_list()  # 使用新的文件列表函数
    logs = get_logs()
    
    # 环境状态检测已取消
    env_status = None
    
    # 获取第一个打印机的功能信息（用于前端显示）
    printer_caps = {}
    if PRINTERS:
        printer_caps = get_printer_capabilities(PRINTERS[0])
    else:
        # 如果没有打印机，提供默认功能信息
        printer_caps = {
            'duplex_support': False,
            'color_support': False,
            'paper_sizes': ['A4', 'A3', 'Letter'],
            'quality_levels': ['normal'],
            'printer_status': '无可用打印机',
            'driver_name': '未知'
        }
    
    if request.method == 'POST':
        # 处理打印请求
        try:
            # 检查设备状态 - 不允许打印和扫描同时进行
            global DEVICE_STATUS
            
            if DEVICE_STATUS['is_scanning']:
                scan_duration = int(time.time() - DEVICE_STATUS['scan_start_time']) if DEVICE_STATUS['scan_start_time'] else 0
                flash(f" 设备正在扫描中，无法开始打印\n\n正在扫描: {scan_duration}秒，操作者: {DEVICE_STATUS['scan_client']}\n\n请等待扫描完成后再试", "danger")
                return redirect(url_for('upload_file'))
            
            if DEVICE_STATUS['is_printing']:
                print_duration = int(time.time() - DEVICE_STATUS['print_start_time']) if DEVICE_STATUS['print_start_time'] else 0
                flash(f" 设备正在打印中，请等待当前任务完成\n\n正在打印: {print_duration}秒，操作者: {DEVICE_STATUS['print_client']}", "warning")
                return redirect(url_for('upload_file'))
            
            # 获取客户端信息
            client_info = get_client_info()
            
            # 获取表单参数
            printer = request.form.get('printer')
            copies = int(request.form.get('copies', 1))
            duplex = int(request.form.get('duplex', 1))
            papersize = request.form.get('papersize', '9')  # 默认A4 ID
            quality = request.form.get('quality', '600x600')
            uploaded_files = request.files.getlist('file')
            
            print(f" 收到打印请求: 打印机={printer}, 份数={copies}, 文件数={len(uploaded_files)}")
            
            # 设置打印状态
            DEVICE_STATUS['is_printing'] = True
            DEVICE_STATUS['print_start_time'] = time.time()
            DEVICE_STATUS['print_client'] = client_info
            
            # 检查是否选择了文件
            if not uploaded_files or all(not f.filename for f in uploaded_files):
                flash(" 错误: 请选择要打印的文件！", "danger")
                return redirect(url_for('upload_file'))
            
            # 检查是否有可用的打印机
            if not printer or printer == "" or printer == "未检测到可用打印机":
                flash(" 错误: 未选择有效的打印机，请检查打印机连接后重试！", "danger")
                return redirect(url_for('upload_file'))
            
            # 检查是否选择了虚拟打印机
            if not is_physical_printer(printer):
                flash(f"️ 警告: '{printer}' 是虚拟打印机，不会进行实际打印，只会生成文件!", "warning")
            
            # 处理上传的文件
            success_count = 0
            total_files = 0
            
            for f in uploaded_files:
                if f and f.filename and allowed_file(f.filename):
                    total_files += 1
                    
                    # 确保文件名唯一，避免覆盖
                    # 兼容 WinXP/旧版浏览器可能提交完整本地路径的情况，只保留文件名
                    raw_filename = f.filename or ''
                    filename = raw_filename.replace('\\', '/').split('/')[-1].strip()
                    if not filename:
                        flash(" 文件名为空，请重新选择要打印的文件！", "danger")
                        continue
                    filepath = os.path.join(UPLOAD_FOLDER, filename)
                    counter = 1
                    max_attempts = 100
                    
                    # 生成唯一文件名
                    original_filename = filename
                    while os.path.exists(filepath) and counter <= max_attempts:
                        name, ext = os.path.splitext(original_filename)
                        filename = f"{name}_{counter}{ext}"
                        filepath = os.path.join(UPLOAD_FOLDER, filename)
                        counter += 1
                        
                    if counter > max_attempts:
                        flash(f" 文件 {original_filename} 名称冲突，请重命名后再上传！", "danger")
                        continue
                    
                    try:
                        # 保存文件到uploads文件夹
                        f.save(filepath)
                        
                        # 验证文件是否成功保存
                        if not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
                            flash(f" 文件 {filename} 保存失败，请重试！", "danger")
                            continue
                            
                        print(f" 文件已保存: {filepath} (大小: {os.path.getsize(filepath)} 字节)")
                        
                        # 根据文件类型选择最佳的静默打印方案
                        file_ext = os.path.splitext(filepath)[1].lower()
                        
                        print(f" 开始打印文件: {filename} -> {printer}")
                        
                        # 根据文件类型选择打印方法
                        success = False
                        message = "未知错误"
                        
                        if file_ext == '.pdf':
                            # PDF文件使用专门的静默打印方法
                            print(f" 使用PDF打印方法")
                            result = print_pdf_silent(filepath, printer, copies)
                        elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
                            # 图片文件使用专门的静默打印方法
                            print(f"️ 使用图片打印方法")
                            result = print_image_silent(filepath, printer, copies)
                        elif file_ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
                            # Office文档使用专门的静默打印方法
                            print(f" 使用Office打印方法")
                            result = print_office_silent(filepath, printer, copies)
                        elif file_ext == '.txt':
                            # 文本文件使用专门的打印方法
                            print(f" 使用文本打印方法")
                            result = print_text_file_simple(filepath, printer, copies)
                        else:
                            # 其他文件使用通用静默打印方法
                            print(f" 使用通用打印方法")
                            result = print_file_with_settings(filepath, printer, copies, duplex, papersize, quality)
                        
                        # 统一处理返回结果
                        if result and len(result) >= 2:
                            success, message = result[0], result[1]
                        elif result and isinstance(result, tuple) and len(result) == 1:
                            success, message = result[0], "打印任务已发送"
                        elif result is True:
                            success, message = True, "打印任务已发送"
                        elif result is False:
                            success, message = False, "打印失败"
                        else:
                            success, message = False, f"未知错误: {result}"
                        
                        # 记录结果
                        if success:
                            success_count += 1
                            print(f" 打印成功: {filename} -> {message}")
                            flash(f" {filename} {message}", "success")
                            log_print(filename, printer, copies, duplex, papersize, quality, client_info)
                        else:
                            print(f" 打印失败: {filename} -> {message}")
                            flash(f" {filename} 打印失败: {message}", "danger")
                            log_print(f"{filename} 失败: {message}", printer, copies, duplex, papersize, quality, client_info)
                            
                    except Exception as e:
                        print(f"️ 打印异常: {filename} -> {str(e)}")
                        error_msg = f"打印异常: {str(e)}"
                        flash(f"️ {filename} {error_msg}", "danger")
                        log_print(f"{filename} {error_msg}", printer, copies, duplex, papersize, quality, client_info)
                        import traceback
                        traceback.print_exc()
                        
                elif f and f.filename:
                    flash(f"️ 文件 {f.filename} 的格式不支持，已跳过", "warning")
            
            # 显示最终统计
            if total_files > 0:
                if success_count == total_files:
                    flash(f" 所有文件({success_count}/{total_files})都已成功发送到打印机！", "success")
                elif success_count > 0:
                    flash(f"️ 部分文件打印成功({success_count}/{total_files})，请检查失败的文件", "warning")
                else:
                    flash(f" 所有文件打印都失败，请检查打印机状态和文件格式！", "danger")
            else:
                flash(" 未找到有效的文件，请检查文件格式是否支持！", "danger")
                
        except Exception as e:
            print(f" POST请求处理异常: {str(e)}")
            flash(f" 请求处理异常: {str(e)}", "danger")
            import traceback
            traceback.print_exc()
        finally:
            # 无论成功失败都要重置打印状态
            DEVICE_STATUS['is_printing'] = False
            DEVICE_STATUS['print_start_time'] = None
            DEVICE_STATUS['print_client'] = ''
        
        return redirect(url_for('upload_file'))
    
    # 获取默认打印机
    default_printer = get_default_printer()
    
    # 获取端口配置信息
    current_port = getattr(app, 'current_port', 5000)
    config_port = get_config_port()
    port_from_config = (current_port == config_port)
    
    return render_template_string(HTML, printers=PRINTERS, files=files, logs=logs, 
                                printer_caps=printer_caps, default_printer=default_printer,
                                env_status=env_status, current_port=current_port,
                                port_from_config=port_from_config)
 
@app.route('/preview/<filename>')
def preview_file(filename):
    fpath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(fpath):
        return f'<div class="container mt-4"><div class="alert alert-danger"><h4>文件未找到</h4><p>文件 "{filename}" 不存在或已被删除！</p><p><a href="/" class="btn btn-primary">返回首页</a></p></div></div>', 404
    
    try:
        ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''
        
        if ext in {'jpg', 'jpeg', 'png'}:
            return send_from_directory(UPLOAD_FOLDER, filename, mimetype=f'image/{ext}')
        elif ext == 'pdf':
            return send_from_directory(UPLOAD_FOLDER, filename, mimetype='application/pdf')
        elif ext == 'txt':
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    content = f.read()
                return f'''
                <div class="container mt-4">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h4>文件预览: {filename}</h4>
                        <a href="/" class="btn btn-secondary">返回首页</a>
                    </div>
                    <div class="card">
                        <div class="card-body">
                            <pre style="white-space: pre-wrap; font-family: monospace;">{content}</pre>
                        </div>
                    </div>
                </div>
                '''
            except UnicodeDecodeError:
                try:
                    with open(fpath, 'r', encoding='gbk') as f:
                        content = f.read()
                    return f'''
                    <div class="container mt-4">
                        <div class="d-flex justify-content-between align-items-center mb-3">
                            <h4>文件预览: {filename}</h4>
                            <a href="/" class="btn btn-secondary">返回首页</a>
                        </div>
                        <div class="card">
                            <div class="card-body">
                                <pre style="white-space: pre-wrap; font-family: monospace;">{content}</pre>
                            </div>
                        </div>
                    </div>
                    '''
                except Exception as e:
                    return f'''
                    <div class="container mt-4">
                        <div class="alert alert-warning">
                            <h4>无法预览文件</h4>
                            <p>文件 "{filename}" 无法以文本格式预览，编码错误: {str(e)}</p>
                            <p><a href="/" class="btn btn-primary">返回首页</a></p>
                        </div>
                    </div>
                    '''
        else:
            # 对于其他文件类型，提供下载链接和基本信息
            file_size = os.path.getsize(fpath)
            if file_size < 1024:
                size_str = f"{file_size} B"
            elif file_size < 1024 * 1024:
                size_str = f"{file_size / 1024:.1f} KB"
            else:
                size_str = f"{file_size / (1024 * 1024):.1f} MB"
            
            return f'''
            <div class="container mt-4">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h4>文件信息: {filename}</h4>
                    <a href="/" class="btn btn-secondary">返回首页</a>
                </div>
                <div class="card">
                    <div class="card-body">
                        <h5 class="card-title"> {filename}</h5>
                        <p class="card-text">
                            <strong>文件类型:</strong> {ext.upper() if ext else 'Unknown'}<br>
                            <strong>文件大小:</strong> {size_str}<br>
                            <strong>说明:</strong> 此文件类型不支持在线预览
                        </p>
                        <div class="btn-group">
                            <a href="/uploads/{filename}" class="btn btn-primary" download>下载文件</a>
                            <button onclick="history.back()" class="btn btn-outline-secondary">返回</button>
                        </div>
                    </div>
                </div>
            </div>
            '''
    except Exception as e:
        return f'''
        <div class="container mt-4">
            <div class="alert alert-danger">
                <h4>预览错误</h4>
                <p>预览文件 "{filename}" 时发生错误: {str(e)}</p>
                <p><a href="/" class="btn btn-primary">返回首页</a></p>
            </div>
        </div>
        ''', 500

# 添加直接下载路由
@app.route('/uploads/<filename>')
def download_file(filename):
    """提供文件下载"""
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)
 
 
def run_flask():
    """运行Flask服务 - 优化稳定性版本"""
    port = getattr(app, 'current_port', 5000)
    max_restart_attempts = 3
    restart_count = 0
    
    while restart_count < max_restart_attempts:
        try:
            print(f" 正在启动Flask服务 (端口:{port}, 尝试:{restart_count + 1}/{max_restart_attempts})...")
            service_manager.mark_service_running()
            
            # 优化Flask配置，增强长期运行稳定性
            from werkzeug.serving import WSGIRequestHandler
            
            # 自定义请求处理器，增加错误处理
            class OptimizedRequestHandler(WSGIRequestHandler):
                def handle_one_request(self):
                    try:
                        super().handle_one_request()
                    except Exception as e:
                        # 忽略客户端断开连接等常见错误
                        if 'Connection aborted' not in str(e) and 'Broken pipe' not in str(e):
                            print(f"️ 请求处理异常: {e}")
                
                def log_error(self, format, *args):
                    # 过滤常见的无害错误日志
                    error_msg = format % args if args else format
                    if any(ignore in error_msg for ignore in [
                        'Connection aborted', 'Broken pipe', 'Connection reset',
                        'Bad file descriptor', 'Invalid HTTP method'
                    ]):
                        return
                    super().log_error(format, *args)
            
            app.run(
                host='0.0.0.0',
                port=port,
                use_reloader=False,
                threaded=True,
                debug=False,
                request_handler=OptimizedRequestHandler,
                # 优化参数
                processes=1,
                passthrough_errors=False
            )
            break  # 正常退出则跳出循环
            
        except OSError as e:
            service_manager.mark_service_stopped()
            if "Address already in use" in str(e):
                print(f" 端口 {port} 已被占用，Flask服务启动失败")
                break
            else:
                print(f" Flask服务启动失败: {e}")
                restart_count += 1
                if restart_count < max_restart_attempts:
                    print(f" {5}秒后重试...")
                    time.sleep(5)
        except Exception as e:
            service_manager.mark_service_stopped()
            print(f" Flask服务异常停止: {e}")
            restart_count += 1
            if restart_count < max_restart_attempts and service_manager.service_running:
                print(f" Flask服务将在{5 * restart_count}秒后自动重启...")
                time.sleep(5 * restart_count)  # 递增延迟
            else:
                print(" Flask服务重启次数已达上限，停止尝试")
                break

def get_waitress_config_for_windows():
    """根据Windows版本返回优化的Waitress配置"""
    import platform
    import sys
    
    try:
        windows_version = platform.release()
        windows_build = None
        
        if hasattr(sys, 'getwindowsversion'):
            win_info = sys.getwindowsversion()
            windows_build = win_info.build if hasattr(win_info, 'build') else None
        
        is_win7 = windows_version == "7"
        is_win11 = windows_build and windows_build >= 22000 if windows_build else False
        
        if is_win7:
            # Win7保守配置，降低资源消耗
            return {
                'threads': 4,                    # 较少线程
                'connection_limit': 50,          # 较少连接
                'cleanup_interval': 120,         # 更频繁清理
                'channel_timeout': 180,          # 较短超时
                'max_request_body_size': 52428800,  # 50MB
                'send_bytes': 4096,              # 较小缓冲区
                'asyncore_use_poll': False,      # Win7可能不支持poll
                'backlog': 32,
                'recv_bytes': 4096
            }
        elif is_win11:
            # Win11高性能配置
            return {
                'threads': 12,                   # 更多线程
                'connection_limit': 500,         # 更多连接
                'cleanup_interval': 30,          # 较少清理频率
                'channel_timeout': 600,          # 更长超时
                'max_request_body_size': 209715200,  # 200MB
                'send_bytes': 16384,             # 更大缓冲区
                'asyncore_use_poll': True,       # 使用高效poll
                'backlog': 128,
                'recv_bytes': 16384
            }
        else:
            # Win10标准配置
            return {
                'threads': 8,
                'connection_limit': 200,
                'cleanup_interval': 60,
                'channel_timeout': 300,
                'max_request_body_size': 104857600,  # 100MB
                'send_bytes': 8192,
                'asyncore_use_poll': True,
                'backlog': 64,
                'recv_bytes': 8192
            }
    except Exception as e:
        print(f"️ Waitress配置检测失败，使用默认配置: {e}")
        # 返回安全的默认配置
        return {
            'threads': 6,
            'connection_limit': 100,
            'cleanup_interval': 90,
            'channel_timeout': 240,
            'max_request_body_size': 104857600,
            'send_bytes': 8192,
            'asyncore_use_poll': True,
            'backlog': 64,
            'recv_bytes': 8192
        }

def run_wsgi():
    """运行WSGI服务 - 生产环境优化版本"""
    port = getattr(app, 'current_port', 5000)
    max_restart_attempts = 3
    restart_count = 0
    
    try:
        from waitress import serve
        from waitress.server import create_server
        
        while restart_count < max_restart_attempts:
            try:
                print(f" 正在启动WSGI服务 (端口:{port}, 尝试:{restart_count + 1}/{max_restart_attempts})...")
                service_manager.mark_service_running()
                
                # 根据Windows版本优化waitress配置
                config = get_waitress_config_for_windows()
                
                server = create_server(
                    app,
                    host='0.0.0.0',
                    port=port,
                    **config
                )
                
                print(f" WSGI服务器配置完成，开始监听...")
                server.run()
                break
                
            except OSError as e:
                service_manager.mark_service_stopped()
                if "Address already in use" in str(e):
                    print(f" 端口 {port} 已被占用，WSGI服务启动失败")
                    break
                else:
                    print(f" WSGI服务启动失败: {e}")
                    restart_count += 1
                    if restart_count < max_restart_attempts:
                        print(f" {5 * restart_count}秒后重试...")
                        time.sleep(5 * restart_count)
            except Exception as e:
                service_manager.mark_service_stopped()
                print(f" WSGI服务异常停止: {e}")
                restart_count += 1
                if restart_count < max_restart_attempts and service_manager.service_running:
                    print(f" WSGI服务将在{5 * restart_count}秒后自动重启...")
                    time.sleep(5 * restart_count)
                else:
                    print(" WSGI服务重启次数已达上限，停止尝试")
                    break
                    
    except ImportError:
        print("️ Waitress未安装，回退到Flask内置服务器")
        run_flask()
 
 
def on_quit(icon, item):
    print(" 正在退出程序...")
    
    try:
        # 1. 标记正在关闭，防止重启
        service_manager.is_shutting_down = True
        
        # 2. 标记服务停止状态
        service_manager.mark_service_stopped()
        
        # 3. 清除任何重启标志
        service_manager.clear_restart()
        
        # 3. 尝试关闭Flask服务器
        print(" 正在关闭Web服务...")
        try:
            if hasattr(app, 'shutdown'):
                app.shutdown()
        except Exception as e:
            print(f"关闭Flask服务时出错: {e}")
        
        # 4. 等待服务线程结束
        print("🧵 正在等待服务线程结束...")
        threads_to_wait = []
        
        if service_manager.flask_thread and service_manager.flask_thread.is_alive():
            threads_to_wait.append(("Flask服务", service_manager.flask_thread))
        
        if service_manager.cleaner_thread and service_manager.cleaner_thread.is_alive():
            threads_to_wait.append(("文件清理", service_manager.cleaner_thread))
            
        if service_manager.monitor_thread and service_manager.monitor_thread.is_alive():
            threads_to_wait.append(("服务监控", service_manager.monitor_thread))
        
        # 给重要线程更多时间优雅退出
        for thread_name, thread in threads_to_wait:
            try:
                print(f"  等待 {thread_name} 线程结束...")
                thread.join(timeout=2)  # 给每个线程2秒时间
                if thread.is_alive():
                    print(f"  ️ {thread_name} 线程未能在2秒内退出")
                else:
                    print(f"   {thread_name} 线程已结束")
            except Exception as e:
                print(f"   等待 {thread_name} 线程时出错: {e}")
        
        # 5. 检查剩余的活跃线程
        remaining_threads = [t for t in threading.enumerate() 
                           if t is not threading.current_thread() and not t.daemon]
        
        if remaining_threads:
            print(f"️ 还有 {len(remaining_threads)} 个非守护线程仍在运行")
            for t in remaining_threads:
                thread_name = getattr(t, 'name', 'Unknown')
                print(f"  - {thread_name}")
        
        # 6. 停止托盘图标
        print("️ 正在停止托盘图标...")
        icon.stop()
        
        # 7. 清理临时文件（可选）
        try:
            import tempfile
            temp_dir = tempfile.gettempdir()
            print(f"清理完成")
        except Exception as e:
            print(f"清理临时文件时出错: {e}")
            
        print(" 程序退出准备完成")
        
    except Exception as e:
        print(f" 退出过程中出现错误: {e}")
    
    finally:
        # 8. 强制退出进程 
        print(" 强制退出程序")
        try:
            # 使用os._exit确保立即退出
            import os
            os._exit(0)
        except Exception:
            # 最后的备用方案
            import sys
            sys.exit(0)

def clear_console():
    """清理控制台内容"""
    try:
        import os
        if os.name == 'nt':  # Windows
            os.system('cls')
            print(" 控制台已清理")
    except:
        pass

def auto_clear_console():
    """定期自动清理控制台的后台任务"""
    while True:
        try:
            # 每30分钟检查一次，如果控制台文本超过1000行则清理
            time.sleep(1800)  # 30分钟
            if CONSOLE_VISIBLE:
                import os
                if os.name == 'nt':
                    from ctypes import windll, create_string_buffer
                    h = windll.kernel32.GetStdHandle(-11)
                    csbi = create_string_buffer(22)
                    windll.kernel32.GetConsoleScreenBufferInfo(h, csbi)
                    lines = csbi.raw[8] * 0x10000 | csbi.raw[9]
                    if lines > 1000:
                        clear_console()
                        print(" 自动清理：控制台文本超过1000行")
        except:
            time.sleep(300)  # 出错后5分钟再试

def show_console():
    """显示控制台窗口"""
    global CONSOLE_WINDOW, CONSOLE_VISIBLE
    try:
        import ctypes
        if not CONSOLE_WINDOW:
            CONSOLE_WINDOW = ctypes.windll.kernel32.GetConsoleWindow()
        if CONSOLE_WINDOW:
            ctypes.windll.user32.ShowWindow(CONSOLE_WINDOW, 1)  # SW_SHOWNORMAL
            ctypes.windll.user32.SetForegroundWindow(CONSOLE_WINDOW)
            CONSOLE_VISIBLE = True
    except:
        pass

def hide_console():
    """隐藏控制台窗口"""
    global CONSOLE_WINDOW, CONSOLE_VISIBLE
    try:
        import ctypes
        if not CONSOLE_WINDOW:
            CONSOLE_WINDOW = ctypes.windll.kernel32.GetConsoleWindow()
        if CONSOLE_WINDOW:
            ctypes.windll.user32.ShowWindow(CONSOLE_WINDOW, 0)  # SW_HIDE
            CONSOLE_VISIBLE = False
    except:
        pass

def toggle_console_window(icon, item):
    """切换控制台窗口显示/隐藏"""
    if CONSOLE_VISIBLE:
        hide_console()
    else:
        show_console()
    # 刷新托盘菜单
    icon.menu = build_menu(icon)
 
def on_toggle_autostart(icon, item):
    current = get_autostart()
    set_autostart(not current)
    # 刷新菜单
    icon.menu = build_menu(icon)


# ===== Tkinter MessageBox 包装函数 - 修复无法关闭问题 =====
def show_message_box(msg_type, title, message):
    """
    显示消息框 - 正确处理 Tkinter 窗口生命周期
    
    Args:
        msg_type: 'info', 'error', 'warning', 'yesno', 'okcancel'
        title: 对话框标题
        message: 消息内容
    
    Returns:
        对于 yesno：True/False
        对于 okcancel：True/False
        其他类型：None
    """
    import tkinter as tk
    from tkinter import messagebox
    
    try:
        # 创建一个不可见的根窗口
        root = tk.Tk()
        root.withdraw()
        # 设置窗口在最前面
        root.attributes('-topmost', True)
        # 更新窗口以确保其在屏幕上注册
        root.update()
        
        result = None
        
        try:
            if msg_type == 'info':
                messagebox.showinfo(title, message, parent=root)
                result = None
            elif msg_type == 'error':
                messagebox.showerror(title, message, parent=root)
                result = None
            elif msg_type == 'warning':
                messagebox.showwarning(title, message, parent=root)
                result = None
            elif msg_type == 'yesno':
                result = messagebox.askyesno(title, message, parent=root)
            elif msg_type == 'okcancel':
                result = messagebox.askokcancel(title, message, parent=root)
            else:
                messagebox.showinfo(title, message, parent=root)
                result = None
        finally:
            # 确保窗口被正确销毁
            try:
                root.destroy()
            except:
                pass
        
        return result
    
    except Exception as e:
        # 如果 Tkinter 显示失败，降级为 console 输出
        print(f"[{title}] {message}")
        if msg_type == 'yesno':
            return False
        return None


def on_show_ip_config(icon, item):
    """通过浏览器打开主页，网络配置功能已集成在托盘菜单"""
    import webbrowser
    ip = get_local_ip()
    port = getattr(app, 'current_port', 5000)
    url = f"http://{ip}:{port}/"
    webbrowser.open(url)

def on_set_current_ip_static(icon, item):
    """将当前IP设置为静态IP"""
    try:
        current_ip = get_local_ip()
        
        if current_ip == '127.0.0.1':
            show_message_box("error", "无效IP", "当前IP为本地回环地址，无法设置为静态IP")
            return
        
        # 确认设置
        result = show_message_box(
            "yesno",
            "设置当前IP为静态", 
            f"确认将当前IP设置为静态IP吗？\n\n"
            f"当前IP: {current_ip}\n"
            f"子网掩码: 255.255.255.0\n"
            f"网关: {'.'.join(current_ip.split('.')[:-1])}.1\n\n"
            f"️ 这将固定当前IP地址"
        )
        
        if result:
            success, message = set_static_ip(current_ip)
            
            if success:
                # 不立即刷新托盘菜单，避免显示127.0.0.1
                show_message_box("info", "设置成功", f"已将当前IP设置为静态IP\n\n静态IP: {current_ip}\n\n注意：托盘菜单中的IP地址可能需要几秒钟才会更新")
                
                # 延迟刷新托盘菜单
                import threading
                def delayed_refresh():
                    import time
                    time.sleep(5)  # 等待5秒让IP完全生效
                    try:
                        icon.menu = build_menu(icon)
                    except:
                        pass
                
                threading.Thread(target=delayed_refresh, daemon=True).start()
            else:
                show_message_box("error", "设置失败", f"设置失败: {message}")
        
    except Exception as e:
        show_message_box("error", "错误", f"设置静态IP时发生错误: {str(e)}")





def on_enable_dhcp(icon, item):
    """启用DHCP"""
    try:
        success, message = set_dhcp()
        
        if success:
            show_message_box("info", "DHCP设置成功", f"{message}\n\n注意：新IP地址可能需要几秒钟才会在托盘菜单中显示")
            
            # 延迟刷新托盘菜单，等待DHCP获取新IP
            import threading
            def delayed_refresh():
                import time
                time.sleep(8)  # DHCP需要更多时间
                try:
                    icon.menu = build_menu(icon)
                except:
                    pass
            
            threading.Thread(target=delayed_refresh, daemon=True).start()
        else:
            show_message_box("error", "DHCP设置失败", message)
    
    except Exception as e:
        show_message_box("error", "错误", f"启用DHCP时发生错误: {str(e)}")

def on_open_github(icon, item):
    """打开GitHub仓库页面"""
    import webbrowser
    webbrowser.open("https://github.com/a937750307/lan-printing")

def on_donate(icon, item):
    """打开赞助页面"""
    try:
        import webbrowser
        url = 'https://zanzhu.937788.xyz/'
        webbrowser.open(url)
    except Exception as e:
        print(f"打开赞助页面失败: {e}")

def on_view_config(icon, item):
    """查看当前配置"""
    try:
        config = load_config()
        config_info = f"""当前配置信息：

端口设置: {config.get('port', 5000)} {'' if config.get('port') else '(默认)'}
配置文件: {CONFIG_FILE}

配置文件内容：
{json.dumps(config, ensure_ascii=False, indent=2) if config else '{}'}

说明：
• 端口设置保存后需要手动重新运行程序才能生效
• 配置文件保存在用户桌面目录  
• 可通过托盘菜单修改端口设置"""
        
        show_message_box("info", "配置信息", config_info)
    except Exception as e:
        show_message_box("error", "错误", f"查看配置时发生错误: {str(e)}")

def on_reset_config(icon, item):
    """重置配置到默认值"""
    try:
        result = show_message_box(
            "yesno",
            "重置配置确认",
            "确定要重置所有配置到默认值吗？\n\n"
            "这将：\n"
            "• 将端口重置为 5000\n"
            "• 删除当前配置文件\n"
            "• 需要重启程序生效"
        )
        
        if result:
            try:
                if os.path.exists(CONFIG_FILE):
                    os.remove(CONFIG_FILE)
                    show_message_box("info", "重置成功", "配置已重置，程序将重启以应用默认设置")
                    
                    # 重启程序
                    import subprocess
                    import sys
                    icon.stop()
                    subprocess.Popen([sys.executable] + sys.argv)
                    sys.exit(0)
                else:
                    show_message_box("info", "提示", "配置文件不存在，当前已是默认配置")
            except Exception as e:
                show_message_box("error", "错误", f"重置配置失败: {str(e)}")
    except Exception as e:
        pass

def on_change_port(icon, item):
    """更改服务端口"""
    import tkinter as tk
    from tkinter import simpledialog
    
    try:
        # 创建隐藏窗口作为对话框的父窗口
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.update()
        
        # 获取当前端口
        current_port = getattr(app, 'current_port', 5000)
        
        # 弹出输入对话框
        try:
            new_port = simpledialog.askinteger(
                "更改端口",
                f"当前端口: {current_port}\n请输入新的端口号 (1024-65535):",
                minvalue=1024,
                maxvalue=65535,
                initialvalue=current_port,
                parent=root
            )
        finally:
            try:
                root.destroy()
            except:
                pass
        
        if new_port and new_port != current_port:
            # 验证端口是否被占用
            import socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            try:
                sock.bind(('localhost', new_port))
                sock.close()
                
                # 端口可用，提示用户重启服务
                result = show_message_box(
                    "yesno",
                    "端口更改确认",
                    f"将端口从 {current_port} 更改为 {new_port}\n\n"
                    f"️ 注意：更改端口后需要手动重新运行程序才能生效\n"
                    f"程序将在保存配置后自动退出\n\n"
                    f"是否继续更改端口？"
                )
                
                if result:
                    # 保存新端口到配置文件
                    if save_port_config(new_port):
                        show_message_box(
                            "info",
                            "端口更改成功", 
                            f"端口已更改为: {new_port}\n"
                            f"新的访问地址: http://{get_local_ip()}:{new_port}\n"
                            f"配置已保存！\n\n"
                            f"️ 请手动重新运行程序以应用新端口设置"
                        )
                    else:
                        show_message_box(
                            "warning",
                            "端口更改", 
                            f"端口已更改为: {new_port}，但配置保存失败\n"
                            f"下次启动可能恢复默认端口\n\n"
                            f"️ 请手动重新运行程序以应用新端口设置"
                        )
                    
                    # 停止托盘图标，退出程序
                    icon.stop()
                    
            except socket.error:
                show_message_box("error", "端口错误", f"端口 {new_port} 已被占用，请选择其他端口")
                try:
                    sock.close()
                except:
                    pass
        
    except Exception as e:
        show_message_box("error", "错误", f"更改端口时发生错误: {str(e)}")

def on_clean_logs(icon, item):
    """手动清理日志"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        # 创建隐藏窗口
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.update()
        
        try:
            result = messagebox.askyesnocancel(
                "清理日志确认",
                "选择日志清理方式：\n\n"
                "按大小(Y) - 保留最新1000条\n"
                "按日期(N) - 删除7天前的记录\n"
                "取消 - 不清理",
                parent=root
            )
        finally:
            try:
                root.destroy()
            except:
                pass
        
        if result is True:
            # 按大小清理
            clean_old_logs()
            show_message_box("info", "清理完成", "已按大小清理日志，保留最新1000条记录")
        elif result is False:
            # 按日期清理
            clean_old_logs_by_date()
            show_message_box("info", "清理完成", "已删除7天前的日志记录")
    
    except Exception as e:
        print(f"手动清理日志失败: {e}")

def on_view_log_info(icon, item):
    """查看日志信息"""
    try:
        if os.path.exists(LOG_FILE):
            file_size = os.path.getsize(LOG_FILE)
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            size_mb = file_size / (1024 * 1024)
            
            # 获取最早和最新的日志时间
            first_date = "未知"
            last_date = "未知"
            
            try:
                if lines:
                    # 第一条记录
                    first_line = lines[0].strip()
                    if len(first_line) >= 19:
                        first_date = first_line[:19]
                    
                    # 最后一条记录
                    last_line = lines[-1].strip()
                    if len(last_line) >= 19:
                        last_date = last_line[:19]
            except:
                pass
            
            info = f""" 打印日志信息

 文件大小: {size_mb:.2f} MB
 记录总数: {len(lines)} 条
 最早记录: {first_date}
 最新记录: {last_date}
 文件路径: {LOG_FILE}

 自动清理规则:
• 文件超过5MB时保留最新1000条
• 每天下午3点清理7天前的记录
• 每天检查一次文件大小"""
            
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("日志信息", info)
            root.destroy()
        else:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("日志信息", " 日志文件不存在")
            root.destroy()
    except Exception as e:
        print(f"查看日志信息失败: {e}")

def on_clear_console(icon, item):
    """手动清理控制台"""
    if CONSOLE_VISIBLE:
        clear_console()


def on_open_qr_plugin(icon, item):
    """在默认浏览器中打开二维码生成浏览器插件的下载链接"""
    try:
        import webbrowser
        webbrowser.open('https://ext.se.360.cn/webstore/detail/afpbjjgbdimpioenaedcjgkaigggcdpp')
    except Exception as e:
        print(f"打开二维码插件链接失败: {e}")

def on_open_upgrade(icon, item):
    """打开版本升级页面"""
    try:
        import webbrowser
        webbrowser.open('https://print.937788.xyz/')
    except Exception as e:
        print(f"打开页面失败: {e}")

def build_menu(icon):
    autostart = get_autostart()
    ip = get_local_ip()
    port = getattr(app, 'current_port', 5000)  # 获取当前端口
    ip_config = get_current_ip_config()
    
    # 构建IP状态显示文本
    ip_status = f"当前IP: {ip}"
    if ip_config:
        if ip_config['dhcp_enabled']:
            ip_status += " (DHCP)"
        else:
            ip_status += " (静态)"
    
    # 检查端口是否来自配置文件
    config_port = get_config_port()
    port_status = f"当前端口: {port}"
    if port == config_port:
        port_status += " "
    else:
        port_status += " (临时)"
    
    return pystray.Menu(
        pystray.MenuItem(f'服务地址: {ip}:{port}', on_show_ip_config),
        pystray.MenuItem(ip_status, None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('网络配置（需要管理员权限）', pystray.Menu(
            pystray.MenuItem('设置当前IP为静态', on_set_current_ip_static),
            pystray.MenuItem('启用DHCP', on_enable_dhcp),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(port_status, None, enabled=False),
            pystray.MenuItem('更改端口', on_change_port),
        )),
        pystray.MenuItem('日志管理', pystray.Menu(
            pystray.MenuItem('查看日志信息', on_view_log_info),
            pystray.MenuItem('清理日志', on_clean_logs),
        )),
        pystray.MenuItem('开机自启：' + ('已开启' if autostart else '未开启'), on_toggle_autostart),
        pystray.Menu.SEPARATOR,
        # 控制台控制选项（所有模式下都可用）
        pystray.MenuItem('控制台', pystray.Menu(
            pystray.MenuItem('显示/隐藏', toggle_console_window),
            pystray.MenuItem('清理内容', on_clear_console),
            pystray.MenuItem('状态：' + ('可见' if CONSOLE_VISIBLE else '隐藏'), None, enabled=False),
        )),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('版本更新', on_open_upgrade),
        pystray.MenuItem('GitHub仓库', on_open_github),
        pystray.MenuItem('赞助作者', on_donate),
        pystray.MenuItem('二维码生成插件', on_open_qr_plugin),
        pystray.MenuItem('退出', on_quit)
    )
 
def setup_tray():
    """设置系统托盘 - 针对不同Windows版本优化"""
    import platform
    import sys
    
    # 检测Windows版本
    try:
        windows_version = platform.release()
        windows_build = None
        
        if hasattr(sys, 'getwindowsversion'):
            win_info = sys.getwindowsversion()
            windows_build = win_info.build if hasattr(win_info, 'build') else None
        
        is_win7 = windows_version == "7"
        is_win11 = windows_build and windows_build >= 22000 if windows_build else False
        
    except Exception as e:
        print(f"Windows版本检测失败: {e}")
        is_win7 = False
        is_win11 = False
    
    try:
        # 加载logo.ico文件 - 改进路径查找逻辑
        logo_path = None
        
        # 候选路径列表，按优先级排序
        candidate_paths = []
        
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller打包后的路径
            candidate_paths.extend([
                path_manager.get_resource_path('logo.ico'),  # 打包内的资源
                os.path.join(os.path.dirname(sys.executable), 'logo.ico'),  # exe同级目录
            ])
        else:
            # 源码运行时的路径
            candidate_paths.extend([
                path_manager.get_resource_path('logo.ico'),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.ico'),
                os.path.join(os.getcwd(), 'logo.ico'),
            ])
        
        # 通用备选路径
        candidate_paths.extend([
            'logo.ico',  # 当前工作目录
            path_manager.get_data_path('logo.ico'),  # 程序目录
        ])
        
        # 查找第一个存在的图标文件
        for path in candidate_paths:
            if os.path.exists(path):
                logo_path = path
                print(f"找到图标文件: {logo_path}")
                break
        
        # 加载图标文件
        if logo_path:
            try:
                image = Image.open(logo_path)
                print(f"成功加载图标文件，尺寸: {image.size}")
                    
            except Exception as e:
                print(f"加载图标失败: {e}，使用默认图标")
                logo_path = None
        
        if not logo_path:
            print("未找到logo.ico文件，创建默认图标")
            # 创建简单的默认图标
            image = Image.new('RGB', (32, 32), color='blue')
            draw = ImageDraw.Draw(image)
            draw.text((12, 12), "P", fill='white')
        
        # 创建系统托盘图标
        icon_title = '内网打印及扫描服务 - by 忆痕'
        
        icon = pystray.Icon('print_server', image, icon_title)
        icon.menu = build_menu(icon)
        
        try:
            icon.run()
        except Exception as e:
            error_msg = f"系统托盘功能启动失败，但程序核心功能正常。\n\n错误信息: {str(e)}\n\n"
            
            if is_win7:
                error_msg += "Win7系统托盘兼容性问题较常见，这是正常现象。\n"
            elif is_win11:
                error_msg += "Win11可能需要管理员权限或安全软件授权。\n"
            
            error_msg += f"您仍可以通过以下方式使用程序：\n• 直接访问: http://{get_local_ip()}:{getattr(app, 'current_port', 5000)}\n• 程序会继续在后台运行\n• 使用 Ctrl+C 可以停止程序"
            
            show_error_dialog("系统托盘启动失败", error_msg, is_critical=False)
            
            # 保持程序运行
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("程序被用户中断")
                sys.exit(0)
        
    except Exception as e:
        print(f"系统托盘初始化失败: {e}")
        
        # 显示详细错误信息
        show_error_dialog(
            "系统托盘初始化失败",
            f"无法初始化系统托盘，可能的原因：\n\n"
            f"1. 缺少图标文件 logo.ico\n"
            f"2. 系统不支持托盘功能\n"
            f"3. 相关库文件缺失\n\n"
            f"程序核心功能正常，您可以直接访问：\n"
            f"http://{get_local_ip()}:{getattr(app, 'current_port', 5000)}\n\n"
            f"错误详情: {str(e)}",
            is_critical=False
        )
        
        # 如果系统托盘失败，至少保持程序运行
        print("程序将继续运行，但没有系统托盘图标")
        print(f"您可以通过浏览器访问: http://{get_local_ip()}:{getattr(app, 'current_port', 5000)}")
        
        # 保持主线程不退出
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("程序被用户中断")
            sys.exit(0)

def show_error_dialog(title, message, is_critical=True):
    """显示友好的错误对话框"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        if is_critical:
            messagebox.showerror(title, message)
        else:
            messagebox.showwarning(title, message)
        
        root.destroy()
        return True
    except Exception:
        # 如果tkinter不可用，回退到控制台输出
        print(f"\n{'='*50}")
        print(f"错误: {title}")
        print(f"{'='*50}")
        print(message)
        print(f"{'='*50}\n")
        return False

def check_admin_privileges():
    """检查是否以管理员模式运行"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def check_windows_features():
    """检查Windows特性和服务"""
    issues = []
    suggestions = []
    
    try:
        # 检查打印服务是否运行
        result = subprocess.run(['sc', 'query', 'Spooler'], 
                              capture_output=True, text=True, timeout=5)
        if 'RUNNING' not in result.stdout:
            issues.append("Windows打印服务未运行")
            suggestions.append("启动打印服务：sc start Spooler")
    except Exception:
        pass
    
    try:
        # 检查Windows防火墙状态
        result = subprocess.run(['netsh', 'advfirewall', 'show', 'allprofiles', 'state'], 
                              capture_output=True, text=True, timeout=5)
        if 'ON' in result.stdout:
            suggestions.append("如果无法访问服务，可能需要在防火墙中允许Python或此程序")
    except Exception:
        pass
    
    return issues, suggestions

if __name__ == '__main__':
    try:
        # 统一的控制台处理逻辑（所有 Windows 版本）
        if hasattr(sys, '_MEIPASS'):
            print(" 检测到exe文件运行模式")
            print(" 内网打印及扫描服务 by 忆痕")
        
        # 无论是exe还是源码运行，都统一处理控制台
        try:
            import ctypes
            import msvcrt
            
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            
            CONSOLE_WINDOW = kernel32.GetConsoleWindow()
            
            if CONSOLE_WINDOW:
                print("\n按任意键可在5秒内保留控制台窗口，否则程序将隐藏控制台...")
                print("\n程序在托盘栏运行，右键程序图标，最上方第一个就是IP端口信息...")
                print("\n局域网下的其他设备可在浏览器地址栏里输入该IP端口，访问可视化网页，进行打印和扫描操作...")
                sys.stdout.flush()
                
                start = time.time()
                keep_console = False
                
                # 等待 5 秒或检测到按键
                while time.time() - start < 5:
                    if msvcrt.kbhit():
                        _ = msvcrt.getch()
                        keep_console = True
                        print("\n已保留控制台窗口（用户按键）")
                        CONSOLE_VISIBLE = True
                        break
                    time.sleep(0.1)
                
                if not keep_console:
                    # 5 秒后隐藏控制台
                    user32.ShowWindow(CONSOLE_WINDOW, 0)  # SW_HIDE
                    CONSOLE_VISIBLE = False
        
        except Exception as e:
            # 如果控制台处理失败，保持可见
            print(f"控制台处理异常: {e}")
            

        
        # 检查命令行参数中的端口设置，并加载配置文件
        import sys
        port = get_config_port()  # 首先从配置文件获取端口
        
        # 命令行参数可以覆盖配置文件设置（用于临时更改）
        for arg in sys.argv:
            if arg.startswith('--port='):
                try:
                    cmdline_port = int(arg.split('=')[1])
                    port = cmdline_port
                    print(f"使用命令行指定端口: {port}")
                except ValueError:
                    print(f"警告: 无效的端口参数 {arg}，使用配置文件端口 {port}")
        
        # 保存当前端口到应用对象
        app.current_port = port
        
        print("=" * 60)
        print("              内网打印及扫描服务")
        print("              作者：忆痕")
        print("    GitHub: https://github.com/a937750307/lan-printing")
        print("=" * 60)
        
        # 检测中文计算机名
        try:
            import socket
            current_hostname = socket.gethostname()
            if any(ord(c) > 127 for c in current_hostname):
                print(f"   警告: 检测到中文计算机名!")
                print(f"   当前计算机名: {current_hostname}")
                print(f"   ")
                print(f"   中文计算机名可能导致以下问题:")
                print(f"   • 网络连接时名称解析错误")
                print(f"   • 远程客户端无法正确获取本机标识")
                print(f"   • 某些网络功能异常")
                print(f"   ")
                print(f"   建议解决方案:")
                print(f"   1. 右键'此电脑' → '属性'")
                print(f"   2. 点击'重命名这台电脑'")
                print(f"   3. 改为英文名(例如: PrintServer-01)")
                print(f"   4. 重启电脑生效")
                print(f"")
        except Exception as e:
            print(f" 计算机名检测异常: {e}")
        
        # 检查管理员权限
        is_admin = check_admin_privileges()
        if is_admin:
            print(" 当前运行模式: 管理员模式 (所有功能可用)")
        else:
            print(" 当前运行模式: 非管理员模式（部分功能受限）")
            print("")
            print("  未以管理员身份运行，部分功能可能无法正常使用！")
            print("")
            print("  请关闭当前程序后，右键exe程序 → '以管理员身份运行' 重新运行以获得完整功能")
            print("")
        
        # 显示路径信息（便于调试）
        print(f"程序目录: {path_manager.app_dir}")
        print(f"上传目录: {UPLOAD_FOLDER}")
        print(f"配置文件: {CONFIG_FILE}")
        print(f"日志文件: {LOG_FILE}")
        if hasattr(sys, '_MEIPASS'):
            print(f" 运行模式: PyInstaller打包 (资源目录: {sys._MEIPASS})")
        else:
            print(f" 运行模式: 源码运行")
        
        # 显示端口信息
        config_port = get_config_port()
        if port == config_port:
            print(f" 使用配置端口: {port}")
        else:
            print(f" 使用临时端口: {port} (配置端口: {config_port})")
        
        # 初始化控制台窗口句柄（所有Windows版本通用）
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            CONSOLE_WINDOW = kernel32.GetConsoleWindow()
            print(" 控制台控制功能已就绪")
        except:
            pass

        # 检测网络状态
        local_ip = get_local_ip()
        if local_ip == '127.0.0.1':
            print(" 网络状态: 离线模式")
            print("    本地打印功能完全正常")
            print("    直接使用系统打印机")
            print("    网络恢复后自动启用完整功能")
            print(f"   本机访问: http://127.0.0.1:{port}")
            
            # 在离线模式下不显示网络异常对话框，因为这是正常情况
            # 用户可能就是想在单机环境下使用打印功能
            
        else:
            print(f" 网络状态: 在线 (IP: {local_ip})")
            print("      完整功能可用")
            print("      支持网络打印机参数获取")
        
        print(f"检测到 {len(PRINTERS)} 台物理打印机")
        if PRINTERS:
            for i, printer in enumerate(PRINTERS[:3], 1):  # 只显示前3台
                print(f"   {i}. {printer}")
            if len(PRINTERS) > 3:
                print(f"   ... 还有 {len(PRINTERS) - 3} 台打印机")
        else:
            print("       未检测到可用的物理打印机")
            print("       程序仍可运行，但打印功能可能受限")
            print("       请检查:")
            print("      - 打印机是否正确连接")
            print("      - 打印机驱动是否已安装")
            print("      - Windows打印机和扫描仪设置")
            
            # 显示打印机检测提示
            show_error_dialog(
                "打印机检测提示",
                "未检测到可用的物理打印机。\n\n"
                "请检查：\n"
                "• 打印机是否正确连接并开机\n"
                "• 打印机驱动程序是否已安装\n"
                "• Windows 设置 > 打印机和扫描仪中是否显示\n"
                "• 尝试重启程序或点击界面中的'刷新'按钮\n\n"
                "程序仍可正常运行，检测到打印机后即可使用。",
                is_critical=False
            )
        
        print("服务器将启动在: http://{}:{}".format(local_ip, port))
        print("=" * 60)
        
        # 检查Windows功能和服务
        issues, suggestions = check_windows_features()
        if issues:
            print(" 检测到以下问题：")
            for issue in issues:
                print(f"   - {issue}")
            print(" 建议解决方案：")
            for suggestion in suggestions:
                print(f"   - {suggestion}")
        
        # 检查端口是否被占用
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.bind(('localhost', port))
            sock.close()
        except socket.error:
            error_msg = f"""端口 {port} 已被占用！

可能的原因：
• 该端口被其他程序占用
• 之前的程序实例未完全关闭
• 系统服务占用了该端口

解决方案：
1. 更换端口：
   python print_server.py --port=5001
   
2. 查找占用进程：
   netstat -ano | findstr :{port}
   
3. 结束占用进程：
   taskkill /PID [进程ID] /F
   
4. 重启计算机后再试

如果问题持续，建议使用其他端口号（如5001-5010）"""
            
            show_error_dialog("端口占用错误", error_msg)
            sys.exit(1)
        
        # 启动定期清理线程
        service_manager.cleaner_thread = threading.Thread(target=clean_old_files, daemon=True)
        service_manager.cleaner_thread.start()
        
        # 启动服务监控线程
        service_manager.monitor_thread = threading.Thread(target=monitor_service_health, daemon=True)
        service_manager.monitor_thread.start()
        print("【OK】服务监控线程已启动")
        
        # 启动控制台自动清理线程
        console_cleaner_thread = threading.Thread(target=auto_clear_console, daemon=True)
        console_cleaner_thread.start()
        print("【OK】控制台自动清理线程已启动")
        
        # 记录服务启动时间
        service_manager.start_time = time.time()
        
        # 判断是否为生产环境
        if os.environ.get('USE_WSGI', '').lower() == 'true':
            service_manager.flask_thread = threading.Thread(target=run_wsgi, daemon=True)
        else:
            service_manager.flask_thread = threading.Thread(target=run_flask, daemon=True)
        service_manager.flask_thread.start()
        
        # 等待Flask服务启动
        print("正在启动Web服务...")
        time.sleep(2)
        
        # 显示启动成功提示
        print("打印服务启动完成！")
        
        if local_ip == '127.0.0.1':
            # 离线模式启动成功
            print(" 离线模式已就绪")
            if len(PRINTERS) > 0:
                print(f"    本机访问: http://127.0.0.1:{port}")
                print(f"    可用打印机: {len(PRINTERS)} 台")
                print("     本地打印功能完全正常")
            else:
                print(f"   本机访问: http://127.0.0.1:{port}")
                print("    未检测到打印机，请检查打印机连接")
                print("    网络恢复后将自动启用完整功能")
        else:
            # 在线模式启动成功
            print(" 在线模式已就绪")
            
            # 检测网络模式并显示相应提示
            network_mode = detect_network_mode()
            external_ip = get_external_ip()
            
            print(f"   访问地址: http://{local_ip}:{port}")
            
            if len(PRINTERS) > 0:
                print(f"   可用打印机: {len(PRINTERS)} 台")
            else:
                print("   未检测到打印机，请检查打印机连接")
            
            # 根据网络模式给出不同提示
            if network_mode == "internal_tunnel" and external_ip:
                print("    局域网环境")
                print(f"      • 内网IP: {local_ip}")
                if external_ip:
                    print(f"      • 路由器公网IP: {external_ip}")
            elif network_mode == "public":
                print("    公网环境 - 外部可直接访问")
            else:
                print("    局域网环境")
        
        print("右键托盘图标查看更多功能")
            
            # 启动成功提示已移除，程序静默启动
        
        setup_tray()
        
    except KeyboardInterrupt:
        print("\n程序被用户中断 (Ctrl+C)，正在优雅退出...")
        try:
            # 使用相同的优雅退出逻辑
            service_manager.mark_service_stopped()
            service_manager.clear_restart()
            
            # 等待重要线程结束
            if service_manager.flask_thread and service_manager.flask_thread.is_alive():
                print("等待Web服务结束...")
                service_manager.flask_thread.join(timeout=1)

            print("【OK】优雅退出完成")
        except Exception as e:
            print(f"优雅退出失败: {e}")
        finally:
            # 强制退出
            import os
            os._exit(0)
    except Exception as e:
        # 获取更详细的系统信息用于诊断
        try:
            import platform
            import traceback
            
            system_info = {
                'system': platform.system(),
                'release': platform.release(),
                'version': platform.version(),
                'machine': platform.machine(),
                'processor': platform.processor(),
                'python_version': platform.python_version(),
            }
            
            # Win11特有错误分析
            win11_hints = []
            error_str = str(e).lower()
            
            if 'access' in error_str or 'permission' in error_str:
                win11_hints.append("权限问题：请以管理员身份运行程序")
            
            if 'import' in error_str or 'module' in error_str:
                win11_hints.append("依赖库缺失：程序打包可能不完整")
            
            if 'socket' in error_str or 'bind' in error_str:
                win11_hints.append("网络权限：检查防火墙和Windows Defender设置")
            
            if 'file' in error_str or 'path' in error_str:
                win11_hints.append("路径问题：避免中文路径，移动到英文目录")
                
            # 生成详细错误报告
            full_traceback = traceback.format_exc()
            
        except:
            system_info = {'error': '无法获取系统信息'}
            win11_hints = []
            full_traceback = str(e)
        
        error_msg = f"""程序启动时发生严重错误：

 错误信息: {str(e)}

️ 系统信息:
• 系统: {system_info.get('system', 'Unknown')} {system_info.get('release', 'Unknown')}
• Python: {system_info.get('python_version', 'Unknown')}
• 架构: {system_info.get('machine', 'Unknown')}

 Win11专用诊断:
""" + '\n'.join(f"• {hint}" for hint in win11_hints) + f"""

 解决方案：
1. 【立即尝试】右键程序图标 → "以管理员身份运行"
2. 【Win11专用】添加Windows Defender排除项：
   • 开始菜单搜索"Windows安全中心"
   • 病毒和威胁防护 → 管理设置 → 添加或删除排除项
   • 添加文件夹：程序所在目录
3. 【路径问题】移动程序到简单英文路径：
   • 例如：C:\\Tools\\PrintService\\
4. 【网络问题】检查防火墙设置：
   • Windows设置 → 隐私和安全性 → Windows安全中心
5. 【依赖问题】重新下载完整版程序

 获取帮助：
• GitHub Issues: https://github.com/a937750307/lan-printing/issues
• 提交时请包含上述系统信息和错误详情

开发者：忆痕

--- 技术详情 (请复制给开发者) ---
{full_traceback}
系统详情: {system_info}
"""
        
        show_error_dialog("程序启动失败 - Win11兼容性", error_msg)
        print(f"\n严重错误: {e}")
        print(f"系统: {system_info}")
        if win11_hints:
            print(f" Win11提示: {', '.join(win11_hints)}")
        print("\n--- 完整错误信息 ---")
        import traceback
        traceback.print_exc()
        sys.exit(1)
