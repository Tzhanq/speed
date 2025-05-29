import pyautogui
import keyboard
import time
import threading
import os
import win32gui
import win32process
import psutil

# PotPlayer按键设置
speed_increase_key = 'c'  # 增加速度的按键（设置为3倍速）
reset_speed_key = 'z'  # 重置速度的按键（恢复1倍速）
toggle_speed_key = 'right'  # 切换速度的按键
fast_forward_key = 'right'  # 模拟快进的按键
rewind_key = 'left'  # 模拟回退的按键

# 线程控制
is_speed_key_pressed = False  # 速度切换键状态
is_ff_key_pressed = False  # 快进键状态
is_rw_key_pressed = False  # 回退键状态
key_press_start_time = 0  # 记录按键开始时间
is_speed_applied = False  # 标记是否已应用三倍速
rw_thread = None  # 回退操作线程


def is_potplayer_focused():
    """检查PotPlayer是否在运行且窗口处于焦点状态"""
    try:
        # 获取当前活动窗口句柄
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return False

        # 获取窗口标题
        window_title = win32gui.GetWindowText(hwnd)

        # 获取窗口进程ID
        _, process_id = win32process.GetWindowThreadProcessId(hwnd)

        # 检查进程是否存在
        for proc in psutil.process_iter(['name']):
            if proc.pid == process_id and 'PotPlayer' in proc.info['name']:
                return True

        # 检查窗口标题是否包含PotPlayer特征
        if "PotPlayer" in window_title or "PotPlay" in window_title:
            return True

        return False
    except:
        return False


def set_three_times_speed():
    """设置为3倍速"""
    global is_speed_applied
    if is_potplayer_focused():
        pyautogui.press(speed_increase_key)
        print("已设置为3倍速")
        is_speed_applied = True


def reset_to_normal_speed():
    """重置为正常速度"""
    global is_speed_applied
    if is_potplayer_focused():
        pyautogui.press(reset_speed_key)
        print("已恢复正常速度")
        is_speed_applied = False


def fast_forward_5_seconds():
    """模拟PotPlayer的Ctrl+方向键右键快进5秒"""
    if is_potplayer_focused():
        pyautogui.hotkey('ctrl', 'right')
        print("已快进5秒")


def rewind_5_seconds():
    """模拟PotPlayer的Ctrl+方向键左键回退5秒"""
    if is_potplayer_focused():
        pyautogui.hotkey('ctrl', 'left')
        print("已回退5秒")


def continuous_rewind():
    """持续回退操作线程函数"""
    while is_rw_key_pressed:
        if is_potplayer_focused():
            rewind_5_seconds()
        time.sleep(0.01)  # 每0.01秒回退一次


def handle_key_states():
    """处理按键状态变化"""
    global is_speed_key_pressed, is_ff_key_pressed, is_rw_key_pressed
    global key_press_start_time, is_speed_applied, rw_thread

    while True:
        # 检测速度切换键状态
        current_speed_state = keyboard.is_pressed(toggle_speed_key)

        if current_speed_state and not is_speed_key_pressed:
            # 速度切换键刚被按下，记录时间
            is_speed_key_pressed = True
            key_press_start_time = time.time()
            is_speed_applied = False
        elif not current_speed_state and is_speed_key_pressed:
            # 速度切换键刚被松开
            is_speed_key_pressed = False
            if is_speed_applied:
                reset_to_normal_speed()
            else:
                # 如果按住时间不足0.3秒，执行快进
                fast_forward_5_seconds()

        # 检测回退键状态
        current_rw_state = keyboard.is_pressed(rewind_key)

        if current_rw_state and not is_rw_key_pressed:
            # 回退键刚被按下，启动回退线程
            is_rw_key_pressed = True
            rw_thread = threading.Thread(target=continuous_rewind)
            rw_thread.daemon = True
            rw_thread.start()
        elif not current_rw_state and is_rw_key_pressed:
            # 回退键刚被松开，停止回退
            is_rw_key_pressed = False
            if rw_thread and rw_thread.is_alive():
                rw_thread.join(timeout=0.1)  # 等待线程结束

        # 检测是否需要设置三倍速
        if is_speed_key_pressed and not is_speed_applied:
            elapsed_time = time.time() - key_press_start_time
            if elapsed_time >= 0.3:
                set_three_times_speed()

        # 短暂休眠减少CPU占用
        time.sleep(0.01)


def exit_immediately():
    """立即终止程序，不执行任何清理操作"""
    print("\n强制终止程序...")
    os._exit(0)  # 立即终止程序，不执行任何清理操作


# 注册Ctrl+Alt+T组合键监听
keyboard.add_hotkey('ctrl+alt+t', exit_immediately)

# 创建并启动按键监听线程
monitor_thread = threading.Thread(target=handle_key_states)
monitor_thread.daemon = True
monitor_thread.start()

# 程序主循环
print(f"程序已启动：")
print(f"  - 按住{toggle_speed_key}键0.3秒后切换到3倍速，松开恢复1倍速")
print(f"  - 快速按下并松开{fast_forward_key}键快进5秒")
print(f"  - 按住{rewind_key}键持续回退5秒，松开关闭")
print(f"按Ctrl+Alt+T立即终止程序")

try:
    while True:
        time.sleep(1)
except Exception as e:
    print(f"发生异常: {e}")