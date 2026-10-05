"""Puts the ribbon text of make_multiuser.py into a trial file and takes pictures of a few screens.
Usage: quickshot.py <project folder> <trial .accde> <output folder> name=object ..."""
import ctypes, os, re, subprocess, sys, time
import win32com.client, win32gui, win32ui

proj, path, out = sys.argv[1:4]
os.makedirs(out, exist_ok=True)
ribbon = re.search(r"RIBBON = '''(.*?)'''", open(os.path.join(proj, 'make_multiuser.py'), encoding='utf-8').read(), re.S).group(1)
db = win32com.client.Dispatch('DAO.DBEngine.120').OpenDatabase(path)
rs = db.OpenRecordset('USysRibbons')
rs.Edit(); rs.Fields('RibbonXml').Value = ribbon; rs.Update(); rs.Close(); db.Close()

app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(path)
hwnd = app.hWndAccessApp()
win32gui.ShowWindow(hwnd, 4)
win32gui.MoveWindow(hwnd, 0, 0, 1500, 900, True)
CONVERT = "Add-Type -AssemblyName System.Drawing; $b=[System.Drawing.Image]::FromFile('%s'); $b.Save('%s',[System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()"


def shot(name):
    time.sleep(3)
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    hdc = win32gui.GetWindowDC(hwnd)
    src = win32ui.CreateDCFromHandle(hdc)
    mem = src.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(src, w, h)
    mem.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(hwnd, mem.GetSafeHdc(), 2)
    tmp = os.path.join(out, name + '.bmp')
    bmp.SaveBitmapFile(mem, tmp)
    win32gui.DeleteObject(bmp.GetHandle())
    mem.DeleteDC(); src.DeleteDC(); win32gui.ReleaseDC(hwnd, hdc)
    subprocess.run(['powershell', '-NoProfile', '-Command', CONVERT % (tmp, os.path.join(out, name + '.png'))], capture_output=True)
    os.remove(tmp)


try:
    app.Run('SetQuiet', True)
    for spec in sys.argv[4:]:
        name, obj = spec.split('=')
        if obj.startswith('rpt'):
            where = 'SaleID=1' if obj == 'rptInvoice' else ''
            app.DoCmd.OpenReport(obj, 2, '', where)
        else:
            app.DoCmd.OpenForm(obj)
        shot(name)
        app.DoCmd.Close(3 if obj.startswith('rpt') else 2, obj, 2)
        print(name)
finally:
    app.CloseCurrentDatabase()
    app.Quit()
