"""
Keep-Alive System Tray Application

Prevents screen lock by sending periodic keystrokes when specific conditions are met:
- During work hours
- On AC power
- With wired ethernet connection
"""

import os
import sys
import time
import threading
from datetime import datetime
from typing import Optional, Tuple, TYPE_CHECKING

import pystray
import pyautogui
import psutil
from PIL import Image, ImageDraw
from pystray import MenuItem as item

# Configuration constants
UPDATE_PERIOD = 30  # seconds
WORK_HOURS_START = 8  # 8 AM
WORK_HOURS_END = 17   # 5 PM
ICON_SIZE = 64
APP_NAME = "Keep Alive"

# Keep-alive key options (scroll lock is least intrusive)
KEEP_ALIVE_KEY = 'scrolllock'  # Options: 'scrolllock', 'f15', 'pause', 'printscreen'


def simulate_key_press():
    """Simulate a key press to prevent screen lock."""
    pyautogui.press(KEEP_ALIVE_KEY)

# Color constants for system tray icon
COLORS = {
    'active': (34, 139, 34, 255),    # Forest green when running
    'paused': (255, 140, 0, 255),    # Orange when paused
    'stopped': (128, 128, 128, 255)  # Gray when stopped
}

# Network interface patterns to skip
SKIP_INTERFACES = ('lo', 'wifi', 'wlan', 'wireless', 'bluetooth', 'local area connection')
ETHERNET_PATTERNS = ('ethernet', 'eth', 'en')


def setup_windows_process_name() -> None:
    """Set custom process name for Windows Task Manager."""
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleTitleW(APP_NAME)
            ctypes.windll.user32.SetWindowTextW(
                ctypes.windll.kernel32.GetConsoleWindow(), 
                APP_NAME
            )
        except Exception:
            pass  # Ignore if it fails


class SystemConditions:
    """Handles checking system conditions for keep-alive activation."""
    
    @staticmethod
    def is_work_hours() -> bool:
        """Check if current time is within work hours."""
        current_hour = datetime.now().hour
        return WORK_HOURS_START <= current_hour < WORK_HOURS_END
    
    @staticmethod
    def is_on_ac_power() -> bool:
        """Check if system is plugged into AC power (not running on battery)."""
        try:
            battery = psutil.sensors_battery()
            if battery is None:
                return True  # No battery detected, assume desktop/always on AC
            return battery.power_plugged
        except Exception:
            return True  # If we can't determine, assume AC power to be safe
    
    @staticmethod
    def has_wired_ethernet_connection() -> bool:
        """Check if there's an active wired ethernet connection."""
        try:
            net_stats = psutil.net_if_stats()
            net_addrs = psutil.net_if_addrs()
            
            for interface_name in net_stats:
                if SystemConditions._should_skip_interface(interface_name):
                    continue
                
                if SystemConditions._is_ethernet_interface(interface_name):
                    if SystemConditions._interface_has_valid_connection(
                        interface_name, net_stats, net_addrs
                    ):
                        return True
            return False
        except Exception as e:
            print(f"Error checking ethernet connection: {e}")
            return False  # If we can't determine, assume no ethernet to be safe
    
    @staticmethod
    def _should_skip_interface(interface_name: str) -> bool:
        """Check if network interface should be skipped."""
        name_lower = interface_name.lower()
        return any(pattern in name_lower for pattern in SKIP_INTERFACES)
    
    @staticmethod
    def _is_ethernet_interface(interface_name: str) -> bool:
        """Check if interface appears to be ethernet."""
        name_lower = interface_name.lower()
        return any(
            name_lower.startswith(pattern) or pattern in name_lower 
            for pattern in ETHERNET_PATTERNS
        )
    
    @staticmethod
    def _interface_has_valid_connection(
        interface_name: str, 
        net_stats: dict, 
        net_addrs: dict
    ) -> bool:
        """Check if interface is up and has a valid IP address."""
        if_stats = net_stats[interface_name]
        if not if_stats.isup:
            return False
        
        if interface_name not in net_addrs:
            return False
        
        for addr in net_addrs[interface_name]:
            if (addr.family == 2 and  # AF_INET (IPv4)
                addr.address and 
                not addr.address.startswith('127.') and
                not addr.address.startswith('169.254.') and  # Exclude APIPA
                addr.address != '0.0.0.0'):
                return True
        return False


class StatusProvider:
    """Provides status information for the system tray menu."""
    
    @staticmethod
    def get_work_hours_status() -> str:
        """Get a string describing current work hours status."""
        if SystemConditions.is_work_hours():
            return f"In work hours ({WORK_HOURS_START:02d}:00 - {WORK_HOURS_END:02d}:00)"
        
        current_hour = datetime.now().hour
        if current_hour < WORK_HOURS_START:
            next_start = f"Work hours start at {WORK_HOURS_START}:00"
        else:
            next_start = f"Work hours resume tomorrow at {WORK_HOURS_START}:00"
        return f"Outside work hours. {next_start}"
    
    @staticmethod
    def get_power_status() -> str:
        """Get a string describing current power status."""
        if SystemConditions.is_on_ac_power():
            return "Connected to AC power"
        
        try:
            battery = psutil.sensors_battery()
            if battery:
                return f"On battery power ({battery.percent:.0f}%)"
            return "On battery power"
        except Exception:
            return "On battery power"
    
    @staticmethod
    def get_ethernet_status() -> str:
        """Get a string describing current ethernet connection status."""
        if SystemConditions.has_wired_ethernet_connection():
            return "Wired ethernet connected"
        return "No wired ethernet connection"


class KeepAliveApp:
    """Main application class for the keep-alive system tray app."""
    
    def __init__(self):
        self.running = False
        self.thread = None
        self.icon = None
        
    def all_conditions_met(self) -> bool:
        """Check if all conditions are met for keep-alive activation."""
        return (SystemConditions.is_work_hours() and 
                SystemConditions.is_on_ac_power() and 
                SystemConditions.has_wired_ethernet_connection())
    
    def get_failed_conditions(self) -> list[str]:
        """Get list of conditions that are not currently met."""
        reasons = []
        if not SystemConditions.is_work_hours():
            reasons.append("Outside Work Hours")
        if not SystemConditions.is_on_ac_power():
            reasons.append("On Battery")
        if not SystemConditions.has_wired_ethernet_connection():
            reasons.append("No Ethernet")
        return reasons
    
    def create_icon_image(self, status: str) -> Image.Image:
        """Create a colored circle icon based on status."""
        image = Image.new('RGBA', (ICON_SIZE, ICON_SIZE), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        color = COLORS.get(status, COLORS['stopped'])
        margin = 8
        draw.ellipse(
            [margin, margin, ICON_SIZE - margin, ICON_SIZE - margin], 
            fill=color, 
            outline=(0, 0, 0, 255)
        )
        
        return image
    
    def keep_alive_loop(self) -> None:
        """Main keep-alive functionality running in a separate thread."""
        while self.running:
            if not self.running:  # Double check for thread safety
                break
                
            if self.all_conditions_met():
                simulate_key_press()
                print(f"Keep-alive: {KEEP_ALIVE_KEY.title()} key pressed at {datetime.now().strftime('%H:%M:%S')}")
            else:
                self._log_skip_reason()
            
            self.update_icon()
            time.sleep(UPDATE_PERIOD)
    
    def _log_skip_reason(self) -> None:
        """Log why keep-alive was skipped."""
        current_time = datetime.now().strftime('%H:%M:%S')
        
        if not SystemConditions.is_work_hours():
            print(f"Keep-alive: Outside work hours ({WORK_HOURS_START}:00-{WORK_HOURS_END}:00), skipping at {current_time}")
        elif not SystemConditions.is_on_ac_power():
            print(f"Keep-alive: Running on battery power, skipping at {current_time}")
        elif not SystemConditions.has_wired_ethernet_connection():
            print(f"Keep-alive: No wired ethernet connection, skipping at {current_time}")
    
    def start_keep_alive(self, icon=None, item=None) -> None:
        """Start the keep-alive functionality."""
        if not self.running:
            self.running = True
            self.thread = threading.Thread(target=self.keep_alive_loop, daemon=True)
            self.thread.start()
            print("Keep-alive started")
            self.update_icon()
    
    def stop_keep_alive(self, icon=None, item=None) -> None:
        """Stop the keep-alive functionality."""
        if self.running:
            self.running = False
            if self.thread:
                self.thread.join(timeout=1)
            print("Keep-alive stopped")
            self.update_icon()
    
    def update_icon(self) -> None:
        """Update the system tray icon and menu based on current state."""
        if not self.icon:
            return
            
        status, title = self._get_icon_status()
        
        self.icon.icon = self.create_icon_image(status)
        self.icon.title = f"{APP_NAME} - {title}"
        self.icon.menu = self.create_menu()
    
    def _get_icon_status(self) -> Tuple[str, str]:
        """Get the current icon status and title."""
        if not self.running:
            return "stopped", "Stopped"
        elif self.all_conditions_met():
            return "active", "Running (Active)"
        else:
            failed_conditions = self.get_failed_conditions()
            return "paused", f"Paused ({', '.join(failed_conditions)})"
    
    def create_menu(self) -> pystray.Menu:
        """Create the context menu for the system tray icon."""
        return pystray.Menu(
            item(APP_NAME, None, enabled=False),
            pystray.Menu.SEPARATOR,
            item('Start', self.start_keep_alive, enabled=lambda item: not self.running),
            item('Stop', self.stop_keep_alive, enabled=lambda item: self.running),
            pystray.Menu.SEPARATOR,
            item(lambda text: StatusProvider.get_work_hours_status(), None, enabled=False),
            item(lambda text: StatusProvider.get_power_status(), None, enabled=False),
            item(lambda text: StatusProvider.get_ethernet_status(), None, enabled=False),
            pystray.Menu.SEPARATOR,
            item('Quit', self.quit_app)
        )
    
    def quit_app(self, icon=None, item=None) -> None:
        """Exit the application."""
        self.stop_keep_alive()
        if self.icon:
            self.icon.stop()
    
    def run(self) -> None:
        """Run the system tray application."""
        print(f"{APP_NAME} System Tray App started")
        print(f"Work hours: {WORK_HOURS_START}:00 - {WORK_HOURS_END}:00")
        print("Keep-alive will only be active during work hours, when on AC power, and with wired ethernet")
        print("Look for the icon in your system tray")
        print("Right-click the icon to start/stop keep-alive functionality")
        
        self.icon = pystray.Icon(
            name="keep_alive",
            icon=self.create_icon_image("stopped"),
            title=f"{APP_NAME} - Stopped",
            menu=self.create_menu()
        )
        
        self.icon.run()


def main() -> None:
    """Main entry point."""
    setup_windows_process_name()
    app = KeepAliveApp()
    app.run()


if __name__ == "__main__":
    main()
