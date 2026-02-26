"""Diagnostic: Use uiautomation package to read WhatsApp chat name.
Switch to WhatsApp and this script will try to find the active chat name.
Press Ctrl+C to stop.
"""
import time
import uiautomation as auto

print("=== WhatsApp UI Automation Diagnostic ===")
print("Switch to WhatsApp and change chats.")
print("Press Ctrl+C to stop.\n")

last_focused = ''
try:
    while True:
        try:
            # Get the focused control
            focused = auto.GetFocusedControl()
            if focused:
                info = f"name={focused.Name!r} cls={focused.ClassName!r} ct={focused.ControlTypeName!r} aid={focused.AutomationId!r}"
                if info != last_focused:
                    last_focused = info
                    print(f"\n--- Focused ---")
                    print(f"  {info}")
                    
                    # Walk up to find parent info
                    parent = focused.GetParentControl()
                    depth = 0
                    while parent and depth < 5:
                        pn = parent.Name or ''
                        pcls = parent.ClassName or ''
                        paid = parent.AutomationId or ''
                        pct = parent.ControlTypeName or ''
                        if pn or paid:
                            print(f"  parent[{depth}]: name={pn[:80]!r} cls={pcls!r} ct={pct!r} aid={paid!r}")
                        if 'whatsapp' in pn.lower():
                            break
                        parent = parent.GetParentControl()
                        depth += 1
            
            # Also try to find WhatsApp window and look for header/title
            wa = auto.WindowControl(searchDepth=1, Name='WhatsApp')
            if wa.Exists(0, 0):
                # Look for specific known patterns in WhatsApp
                # The chat header usually contains the contact name
                
                # Try to find text elements at the top (header area)
                all_texts = wa.GetChildren()
                if all_texts:
                    for i, child in enumerate(all_texts[:10]):
                        cn = child.Name or ''
                        ccls = child.ClassName or ''
                        caid = child.AutomationId or ''
                        cct = child.ControlTypeName or ''
                        if cn or caid:
                            print(f"  wa_child[{i}]: name={cn[:80]!r} cls={ccls!r} ct={cct!r} aid={caid!r}")
                            
                            # Go one level deeper
                            for j, sub in enumerate(child.GetChildren()[:10]):
                                sn = sub.Name or ''
                                said = sub.AutomationId or ''
                                sct = sub.ControlTypeName or ''
                                if sn or said:
                                    print(f"    wa_sub[{i}.{j}]: name={sn[:80]!r} ct={sct!r} aid={said!r}")
        
        except Exception as e:
            err = str(e)
            if 'whatsapp' in err.lower() or len(err) < 100:
                pass  # skip noise
            else:
                print(f"Error: {err[:200]}")
        
        time.sleep(2)

except KeyboardInterrupt:
    print("\nDone.")
