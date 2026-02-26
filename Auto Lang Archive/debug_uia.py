"""Diagnostic: Use Windows UI Automation to read WhatsApp chat name.
Switch to WhatsApp and this script will try to find the active chat name.
Press Ctrl+C to stop.
"""
import time
import ctypes
import ctypes.wintypes as wintypes
import comtypes
import comtypes.client

# UI Automation constants
UIA_NamePropertyId = 30005
UIA_ControlTypePropertyId = 30003
UIA_ClassNamePropertyId = 30012
UIA_AutomationIdPropertyId = 30011

# Control type IDs
UIA_TextControlTypeId = 50020
UIA_EditControlTypeId = 50004
UIA_DocumentControlTypeId = 50030
UIA_PaneControlTypeId = 50033
UIA_GroupControlTypeId = 50026
UIA_ListItemControlTypeId = 50007
UIA_WindowControlTypeId = 50032

# Tree scope
TreeScope_Element = 1
TreeScope_Children = 2
TreeScope_Subtree = 7
TreeScope_Descendants = 4


def get_foreground_hwnd():
    user32 = ctypes.windll.user32
    return user32.GetForegroundWindow()


def explore_whatsapp():
    """Try to find the chat name using UI Automation."""
    # Initialize COM
    comtypes.CoInitialize()
    
    # Create UIAutomation object
    uia = comtypes.client.CreateObject(
        '{ff48dba4-60ef-4201-aa87-54103eef594e}',
        interface=comtypes.gen.UIAutomationClient.IUIAutomation
    )
    
    print("=== WhatsApp UI Automation Diagnostic ===")
    print("Switch to WhatsApp. Press Ctrl+C to stop.\n")
    
    last_info = ''
    
    try:
        while True:
            hwnd = get_foreground_hwnd()
            if not hwnd:
                time.sleep(1)
                continue
            
            try:
                element = uia.ElementFromHandle(hwnd)
                name = element.CurrentName
                cls = element.CurrentClassName
                
                if 'whatsapp' not in (name or '').lower() and 'whatsapp' not in (cls or '').lower():
                    time.sleep(0.5)
                    continue
                
                # We're in WhatsApp! Let's explore the tree
                print(f"\n{'='*60}")
                print(f"Window: name={name!r} class={cls!r}")
                
                # Try to get the focused element
                try:
                    focused = uia.GetFocusedElement()
                    if focused:
                        f_name = focused.CurrentName
                        f_cls = focused.CurrentClassName
                        f_aid = focused.CurrentAutomationId
                        print(f"\nFocused element:")
                        print(f"  name={f_name!r}")
                        print(f"  class={f_cls!r}")
                        print(f"  automationId={f_aid!r}")
                except Exception as e:
                    print(f"  (focused element error: {e})")
                
                # Walk children of the main window (depth 1-2)
                print(f"\nTop-level children:")
                try:
                    walker = uia.RawViewWalker
                    child = walker.GetFirstChildElement(element)
                    depth1_count = 0
                    while child and depth1_count < 15:
                        c_name = child.CurrentName or ''
                        c_cls = child.CurrentClassName or ''
                        c_aid = child.CurrentAutomationId or ''
                        if c_name or c_aid:
                            print(f"  [{depth1_count}] name={c_name[:80]!r} class={c_cls!r} aid={c_aid!r}")
                        
                        # Go one level deeper
                        subchild = walker.GetFirstChildElement(child)
                        sub_count = 0
                        while subchild and sub_count < 10:
                            s_name = subchild.CurrentName or ''
                            s_cls = subchild.CurrentClassName or ''
                            s_aid = subchild.CurrentAutomationId or ''
                            if s_name or s_aid:
                                print(f"    [{depth1_count}.{sub_count}] name={s_name[:80]!r} class={s_cls!r} aid={s_aid!r}")
                            
                            # One more level
                            subsub = walker.GetFirstChildElement(subchild)
                            ss_count = 0
                            while subsub and ss_count < 8:
                                ss_name = subsub.CurrentName or ''
                                ss_cls = subsub.CurrentClassName or ''
                                ss_aid = subsub.CurrentAutomationId or ''
                                if ss_name or ss_aid:
                                    print(f"      [{depth1_count}.{sub_count}.{ss_count}] name={ss_name[:80]!r} class={ss_cls!r} aid={ss_aid!r}")
                                subsub = walker.GetNextSiblingElement(subsub)
                                ss_count += 1
                            
                            subchild = walker.GetNextSiblingElement(subchild)
                            sub_count += 1
                        
                        child = walker.GetNextSiblingElement(child)
                        depth1_count += 1
                except Exception as e:
                    print(f"  (tree walk error: {e})")
                
            except Exception as e:
                print(f"Error: {e}")
            
            time.sleep(2)
    
    except KeyboardInterrupt:
        print("\nDone.")
    finally:
        comtypes.CoUninitialize()


if __name__ == '__main__':
    explore_whatsapp()
