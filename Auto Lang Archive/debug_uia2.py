"""Diagnostic: Use Windows UI Automation (raw COM) to read WhatsApp chat name.
Switch to WhatsApp and this script will try to find the active chat name.
Press Ctrl+C to stop.
"""
import time
import ctypes
import ctypes.wintypes as wintypes

# ─── Raw COM UIAutomation via ctypes ──────────────────────
# We avoid comtypes entirely and use the simpler approach:
# Use the Windows Accessibility API (MSAA / IAccessible) instead.

from ctypes import POINTER, byref, windll, HRESULT
import subprocess, sys

# Actually let's use a simpler approach - pywinauto or just plain win32 accessibility
# Simplest: use Windows' built-in oleacc.dll (IAccessible / MSAA)

oleacc = windll.oleacc
user32 = windll.user32

# AccessibleObjectFromWindow
OBJID_CLIENT = 0xFFFFFFFC  # -4

# IAccessible via oleacc
from comtypes import CoInitialize, CoUninitialize
from comtypes.automation import VARIANT
from comtypes import COMError
import comtypes

# Generate UIAutomation typelib
CoInitialize()
try:
    # Try to generate the UIAutomation type library
    from comtypes.client import GetModule, CreateObject
    # UIAutomation CLSID and TypeLib
    try:
        mod = GetModule("UIAutomationCore.dll")
    except Exception as e:
        print(f"GetModule failed: {e}")
        print("Trying alternative...")
        mod = None
    
    if mod is None:
        # Fallback: use IUIAutomation via its CLSID directly
        import comtypes.gen
        print("Trying direct COM creation...")
    
    # CUIAutomation CLSID
    CLSID_CUIAutomation = comtypes.GUID('{ff48dba4-60ef-4201-aa87-54103eef594e}')
    
    # Try creating the object
    uia = CreateObject(CLSID_CUIAutomation)
    
    print("=== WhatsApp UI Automation Diagnostic ===")
    print("Switch to WhatsApp. This checks every 2 seconds.")
    print("Press Ctrl+C to stop.\n")
    
    last_name = ''
    
    while True:
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            time.sleep(1)
            continue
        
        try:
            # Get the element from the foreground window
            element = uia.ElementFromHandle(hwnd)
            win_name = element.CurrentName
            
            if 'whatsapp' not in (win_name or '').lower():
                time.sleep(0.5)
                continue
            
            print(f"\n{'='*60}")
            print(f"WhatsApp window detected: {win_name!r}")
            
            # Get focused element - often the chat input or header
            try:
                focused = uia.GetFocusedElement()
                if focused:
                    print(f"\nFocused: name={focused.CurrentName!r}  aid={focused.CurrentAutomationId!r}  cls={focused.CurrentClassName!r}")
            except Exception as e:
                print(f"  Focused error: {e}")
            
            # Walk the tree looking for interesting elements
            walker = uia.RawViewWalker
            
            print("\nTree (3 levels deep):")
            child = walker.GetFirstChildElement(element)
            idx = 0
            while child:
                try:
                    n = child.CurrentName or ''
                    a = child.CurrentAutomationId or ''
                    c = child.CurrentClassName or ''
                    if n or a:
                        print(f"  L1[{idx}] name={n[:100]!r}  aid={a!r}  cls={c!r}")
                    
                    # Level 2
                    sub = walker.GetFirstChildElement(child)
                    sidx = 0
                    while sub:
                        try:
                            sn = sub.CurrentName or ''
                            sa = sub.CurrentAutomationId or ''
                            sc = sub.CurrentClassName or ''
                            if sn or sa:
                                print(f"    L2[{idx}.{sidx}] name={sn[:100]!r}  aid={sa!r}  cls={sc!r}")
                            
                            # Level 3
                            subsub = walker.GetFirstChildElement(sub)
                            ssidx = 0
                            while subsub and ssidx < 15:
                                try:
                                    ssn = subsub.CurrentName or ''
                                    ssa = subsub.CurrentAutomationId or ''
                                    ssc = subsub.CurrentClassName or ''
                                    if ssn or ssa:
                                        print(f"      L3[{idx}.{sidx}.{ssidx}] name={ssn[:100]!r}  aid={ssa!r}  cls={ssc!r}")
                                except:
                                    pass
                                subsub = walker.GetNextSiblingElement(subsub)
                                ssidx += 1
                        except:
                            pass
                        sub = walker.GetNextSiblingElement(sub)
                        sidx += 1
                        if sidx > 20:
                            break
                except:
                    pass
                child = walker.GetNextSiblingElement(child)
                idx += 1
                if idx > 20:
                    break
            
        except Exception as e:
            print(f"Error: {e}")
        
        time.sleep(2)

except KeyboardInterrupt:
    print("\nDone.")
except Exception as e:
    print(f"\nFatal error: {e}")
    import traceback
    traceback.print_exc()
finally:
    CoUninitialize()
