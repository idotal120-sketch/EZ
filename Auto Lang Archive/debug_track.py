"""Quick: prints exe + title every time it changes. Switch between apps/emails/chats.
Press Ctrl+C to stop.
"""
import ctypes, ctypes.wintypes as wintypes, time

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_ENUM_CHILD_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.EnumChildWindows.argtypes = [wintypes.HWND, ctypes.c_void_p, wintypes.LPARAM]
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

def pid_to_exe(pid):
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h: return '?'
    try:
        bl = wintypes.DWORD(260); b = ctypes.create_unicode_buffer(260)
        if kernel32.QueryFullProcessImageNameW(h, 0, b, ctypes.byref(bl)):
            return b.value.rsplit('\\', 1)[-1].lower()
        return '?'
    finally:
        kernel32.CloseHandle(h)

def get_title(hwnd):
    l = user32.GetWindowTextLengthW(hwnd)
    if l <= 0: return ''
    b = ctypes.create_unicode_buffer(l + 1)
    user32.GetWindowTextW(hwnd, b, l + 1)
    return b.value

def get_real_exe(hwnd):
    pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    exe = pid_to_exe(pid.value)
    if exe == 'applicationframehost.exe':
        parent_pid = pid.value
        result = ['']
        def cb(child, _):
            cpid = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(child, ctypes.byref(cpid))
            if cpid.value and cpid.value != parent_pid:
                e = pid_to_exe(cpid.value)
                if e and e != 'applicationframehost.exe':
                    result[0] = e
                    return False
            return True
        user32.EnumChildWindows(hwnd, _ENUM_CHILD_PROC(cb), 0)
        if result[0]:
            return f"{result[0]} (via UWP)"
    return exe

print("=== Window Tracker - switch between apps/emails/chats ===")
print("Press Ctrl+C to stop.\n")

last = ''
try:
    while True:
        hwnd = user32.GetForegroundWindow()
        if hwnd:
            exe = get_real_exe(hwnd)
            title = get_title(hwnd)
            info = f"{exe}|{title}"
            if info != last:
                last = info
                print(f"[{time.strftime('%H:%M:%S')}] exe={exe!r:30s} title={title[:80]!r}")
        time.sleep(0.3)
except KeyboardInterrupt:
    print("\nDone.")
