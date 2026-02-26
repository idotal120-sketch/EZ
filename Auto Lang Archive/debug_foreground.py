"""Quick diagnostic: prints foreground window exe + title every second.
Switch to WhatsApp and change chats to see what values we get.
Press Ctrl+C to stop.
"""
import ctypes
import ctypes.wintypes as wintypes
import time

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
psapi = ctypes.windll.psapi

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = wintypes.INT
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, wintypes.INT]
user32.GetWindowTextW.restype = wintypes.INT
user32.EnumChildWindows.argtypes = [wintypes.HWND, ctypes.c_void_p, wintypes.LPARAM]
user32.EnumChildWindows.restype = wintypes.BOOL

kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL
kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL

_ENUM_CHILD_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def pid_to_exe(pid_val):
    hproc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid_val)
    if not hproc:
        return '?'
    try:
        buf_len = wintypes.DWORD(260)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if kernel32.QueryFullProcessImageNameW(hproc, 0, buf, ctypes.byref(buf_len)):
            return buf.value.rsplit('\\', 1)[-1].lower()
        return '?'
    finally:
        kernel32.CloseHandle(hproc)


def get_title(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ''
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def get_child_info(parent_hwnd):
    """List all child windows with their PID, exe, title, class."""
    parent_pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(parent_hwnd, ctypes.byref(parent_pid))
    children = []

    def cb(child_hwnd, _):
        child_pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(child_hwnd, ctypes.byref(child_pid))
        exe = pid_to_exe(child_pid.value)
        title = get_title(child_hwnd)
        cls_buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(child_hwnd, cls_buf, 256)
        children.append({
            'pid': child_pid.value,
            'same_as_parent': child_pid.value == parent_pid.value,
            'exe': exe,
            'title': title,
            'class': cls_buf.value,
        })
        return True

    callback = _ENUM_CHILD_PROC(cb)
    user32.EnumChildWindows(parent_hwnd, callback, 0)
    return parent_pid.value, children


# Add GetClassNameW
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, wintypes.INT]
user32.GetClassNameW.restype = wintypes.INT

print("=== Foreground Window Diagnostic ===")
print("Switch to WhatsApp and change chats. Press Ctrl+C to stop.\n")

last_info = ''
try:
    while True:
        hwnd = user32.GetForegroundWindow()
        if hwnd:
            pid = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            exe = pid_to_exe(pid.value)
            title = get_title(hwnd)
            
            info = f"exe={exe!r}  title={title!r}"
            
            if info != last_info:
                last_info = info
                print(f"\n{'='*60}")
                print(f"HWND: {hwnd}  PID: {pid.value}")
                print(f"EXE:   {exe}")
                print(f"TITLE: {title}")
                print(f"TITLE bytes: {[hex(ord(c)) for c in title[:50]]}")
                
                if exe == 'applicationframehost.exe':
                    print("\n  >> UWP app detected! Scanning children...")
                    parent_pid, children = get_child_info(hwnd)
                    diff_pid_children = [c for c in children if not c['same_as_parent']]
                    print(f"  >> Children with different PID: {len(diff_pid_children)}")
                    for c in diff_pid_children[:5]:
                        print(f"     exe={c['exe']!r}  class={c['class']!r}  title={c['title'][:60]!r}")
                    if not diff_pid_children:
                        print("  >> No children with different PID found!")
                        print(f"  >> Total children: {len(children)}")
                        for c in children[:5]:
                            print(f"     exe={c['exe']!r}  class={c['class']!r}  title={c['title'][:60]!r}")
        
        time.sleep(0.5)
except KeyboardInterrupt:
    print("\nDone.")
