#!/usr/bin/env python

# Based on pywin32 Demo RegisterDeviceNotification
import sys, time
import win32gui, win32con, win32api, win32file
import win32gui_struct, winnt
import psutil
import subprocess

GUID_DEVINTERFACE_USB_DEVICE = "{A5DCBF10-6530-11D2-901F-00C04FB951ED}"
YUBIKEY_PIDVID = 'VID_1050&PID_0116'
GPG_PROC = ('scdaemon.exe', 'gpg-agent.exe')
GPG_CMD = ('C:\Program Files (x86)\GNU\GnuPG\pub\gpg.EXE', '--card-status')

# Kill old gpg-agent
def KillGPGAgent():
    for p in psutil.process_iter():
        if p.name() in GPG_PROC:
            p.terminate()

# Run new gpg-agent
def RunGPGAgent():
    p = subprocess.Popen(GPG_CMD)
#    print GPG_CMD, p

# WM_DEVICECHANGE message handler.
def OnDeviceChange(hwnd, msg, wp, lp):
    # Unpack the 'lp' into the appropriate DEV_BROADCAST_* structure,
    # using the self-identifying data inside the DEV_BROADCAST_HDR.
    info = win32gui_struct.UnpackDEV_BROADCAST(lp)
#    print "Device change notification:", wp, str(info)
    if not info:
        return True

    # React only on Yubikey USB device
    if info.devicetype != win32con.DBT_DEVTYP_DEVICEINTERFACE or \
        YUBIKEY_PIDVID not in info.name:
        return True

    KillGPGAgent()
    if win32con.DBT_DEVICEREMOVECOMPLETE == wp:
        print 'Yubikey removed'
    else:
        print 'Yubikey connected'
        RunGPGAgent()
    return True

def YubikeyWatchdog():
    wc = win32gui.WNDCLASS()
    wc.lpszClassName = 'yubikey_notify'
    wc.style =  win32con.CS_GLOBALCLASS|win32con.CS_VREDRAW | win32con.CS_HREDRAW
    wc.hbrBackground = win32con.COLOR_WINDOW+1
    wc.lpfnWndProc={win32con.WM_DEVICECHANGE:OnDeviceChange}
    class_atom=win32gui.RegisterClass(wc)
    hwnd = win32gui.CreateWindow(wc.lpszClassName,
        'Yubikey watchdog',
        # no need for it to be visible.
        win32con.WS_CAPTION,
        100,100,900,900, 0, 0, 0, None)

    hdevs = []
    # Watch for all USB device notifications
    filter = win32gui_struct.PackDEV_BROADCAST_DEVICEINTERFACE(
                                        GUID_DEVINTERFACE_USB_DEVICE)
    hdev = win32gui.RegisterDeviceNotification(hwnd, filter,
                                               win32con.DEVICE_NOTIFY_WINDOW_HANDLE)
    hdevs.append(hdev)
    while True:
        win32gui.PumpWaitingMessages()
        time.sleep(0.01)
    # Cleanup
    win32gui.DestroyWindow(hwnd)
    win32gui.UnregisterClass(wc.lpszClassName, None)

if __name__=='__main__':
    YubikeyWatchdog()