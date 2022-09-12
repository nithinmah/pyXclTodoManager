from baseIncludes import *

import subprocess
import win32gui
import pyautogui
import pymsgbox

def start_monitor_process_in_next_tab():
  with pyautogui.hold('ctrl'):
    sleep(0.1)
    pyautogui.press('tab')
    sleep(0.1)
  pyautogui.write("cd ~/bin")
  sleep(0.1)
  pyautogui.press('enter')
  sleep(0.1)
  pyautogui.write("./reminder_external_monitor.py")
  sleep(0.1)
  pyautogui.press('enter')
  sleep(0.1)
  with pyautogui.hold('ctrl'):
    with pyautogui.hold('shift'):
      pyautogui.press('tab')
      sleep(0.1)
  
class wpsExcelSheet:
  wps_exe_path="C:\\Users\\ethimah\\AppData\\Local\\kingsoft\\WPS Office\\ksolaunch.exe"
  autohotkey_path="C:\\Users\\ethimah\\bin\\autohotkey.exe"
  opentasksheet_ahk="C:\\Users\\ethimah\\bin\\opentasksheet.ahk"
  savesheet_ahk="C:\\Users\\ethimah\\bin\\savesheet.ahk"
  sortsheet_ahk="C:\\Users\\ethimah\\bin\\sortsheet.ahk"
  closesheet_ahk="C:\\Users\\ethimah\\bin\\closesheet.ahk"
  updateform_ahk="C:\\Users\\ethimah\\bin\\updateform.ahk"
  sheet_state = "unknown"
  #trace_obj = trace()

  def __init__(self, debug=False):
    trace_obj.prologue(currentframe().f_code.co_name)
    self.debug = debug
    trace_obj.epilogue(currentframe().f_code.co_name)

  def set_sheet_state(self, state):
    self.sheet_state = state

  def is_wps_open(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    wps_window_list = pyautogui.getWindowsWithTitle("WPS")
    if len(wps_window_list) >  0 :
      ret_val = True
    else:
      ret_val = False
    trace_obj.epilogue(currentframe().f_code.co_name)
    return ret_val  

  def focus_wps(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    pyautogui.getWindowsWithTitle("WPS")[0].activate()
    trace_obj.epilogue(currentframe().f_code.co_name)

  def minimize_wps(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    pyautogui.getWindowsWithTitle("WPS")[0].minimize()
    self.sheet_state = "unknown"
    trace_obj.epilogue(currentframe().f_code.co_name)

  def is_tasktracker_sheet_open(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    if self.is_wps_open() == False:
      trace_obj.epilogue(currentframe().f_code.co_name)
      return False
    if self.sheet_state != "unknown":
      if self.sheet_state == "open":
        trace_obj.epilogue(currentframe().f_code.co_name)
        return True
      else:
        trace_obj.epilogue(currentframe().f_code.co_name)
        return False
    wps_title_begin   = pyautogui.getWindowsWithTitle("WPS")[0].title
    #trace("open sheet: " + wps_title_begin)
    wps_title_current = wps_title_begin
    while "tasktracker" not in wps_title_current:
      pyautogui.getWindowsWithTitle("WPS")[0].activate()
      sleep_mine(0.5)
      with pyautogui.hold('ctrl'):
        pyautogui.press('tab')
        sleep_mine(0.5)
        wps_title_current = pyautogui.getWindowsWithTitle("WPS")[0].title
        #trace("open sheet now: " + wps_title_current)
        if wps_title_current == wps_title_begin: #iterated all tabs
          trace_obj.trace("couldn't find open sheet")
          trace_obj.epilogue(currentframe().f_code.co_name)
          return False
    pyautogui.getWindowsWithTitle("WPS")[0].activate() #If open bring it to focus too ?
    self.sheet_state = "open"
    trace_obj.epilogue(currentframe().f_code.co_name)
    return True
  
  def get_back_to_python(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    pyautogui.getWindowsWithTitle("MSYS")[0].activate()
    trace_obj.epilogue(currentframe().f_code.co_name)

  def open_sheet(self):
    #pyautogui.hotkey('winleft', 'd')
    trace_obj.prologue(currentframe().f_code.co_name)
    return_code = subprocess.call([self.autohotkey_path, self.opentasksheet_ahk])
    if return_code:
      pyautogui.confirm('unable to close sheet, please close manually', buttons=["OK"], timeout=60000)
    #wps_window_list = pyautogui.getWindowsWithTitle("WPS")
    #trace_obj.trace(str(len(wps_window_list)))
    #if len(wps_window_list) != 0:
    #  trace_obj.trace("WPS already open: " + wps_window_list[0].title)
    #  if "tasktracker.xlsx" in wps_window_list[0].title:
    #    trace_obj.trace("task sheet already open")
    #    self.focus_wps()
        #If it takes win32 a while to maximize wps window
        # next keystrokes could get loss. Observed one:
        # save keystroke is lost and close stops prompting
        # for save before close!
    #    sleep(0.1)
    #    trace_obj.epilogue(currentframe().f_code.co_name)
    #    return
    #subprocess.call([self.wps_exe_path, tasks_xls_path])
    #wps_window_list = pyautogui.getWindowsWithTitle("WPS")
    #while len(wps_window_list) == 0:
    #  sleep_mine(0.5)
    #  wps_window_list = pyautogui.getWindowsWithTitle("WPS")
    #while "tasktracker.xlsx" not in wps_window_list[0].title:
    #  sleep_mine(0.5)
    #  wps_window_list = pyautogui.getWindowsWithTitle("WPS")
    self.sheet_state = "open"
    #below sleep is crucial, as sometimes WPS can take more
    #time to open the sheet based on PC CPU% usage. If we dont
    #wait enough, next actions like sort, save, formula can send
    #keys before the sheet is open and they can become edit of
    #sheets rather than control key sequences. For example
    #in the alt+d+s for sort, the alt can go before the sheet is
    #open and ready to accept keys. then d+s+enter becomes edit 
    #of the cell! Patience is not optional, speed thrills but
    #kills!
    #sleep_mine(1.5)
    trace_obj.epilogue(currentframe().f_code.co_name)
  
  def save_sheet(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    #if self.sheet_state != "open":
    if self.sheet_state != "open":
      self.open_sheet()
    subprocess.call([self.autohotkey_path, self.savesheet_ahk])
    #with pyautogui.hold('ctrl'):
    #  pyautogui.press('s')
    #sleep_mine(1)
    trace_obj.epilogue(currentframe().f_code.co_name)
    
  def close_sheet(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    trace_obj.trace("sheet_state : " + self.sheet_state)
    if self.sheet_state == "closed":
      return
    self.open_sheet()
    self.save_sheet()
    return_val = subprocess.call([self.autohotkey_path, self.closesheet_ahk])
    if return_val != 0:
      pyautogui.confirm('unable to close sheet, please close manually', buttons=["OK"], timeout=60000)
    #sleep_mine(0.5)
    #with pyautogui.hold('ctrl'):
    #  pyautogui.press('w')
    #sleep_mine(1)
    #pyautogui.press('esc')
    self.sheet_state = "closed"
    #self.minimize_wps()
    trace_obj.epilogue(currentframe().f_code.co_name)
  
  def sort_sheet(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    if self.sheet_state != "open":
      self.open_sheet()
    return_val = subprocess.call([self.autohotkey_path, self.sortsheet_ahk])
    if return_val != 0:
      pyautogui.confirm('unable to sort sheet, please sort manually', buttons=["OK"], timeout=60000)
     
    #sleep(2)
    #with pyautogui.hold('alt'):
    #  pyautogui.press('d')
    #  pyautogui.press('s')
    #sleep_mine(1.5) # This is crucial, dont reduce this. Without this 
               # the below enter  goes in before the popuup shows 
               # up and we are left with popup open and sort not done
    #pyautogui.press('enter')
    #sleep_mine(0.5)
    self.save_sheet()
    trace_obj.epilogue(currentframe().f_code.co_name)

  def update_formulas_on_tracker_sheet(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    self.open_sheet()
    subprocess.call([self.autohotkey_path, self.updateform_ahk])
    #sleep_mine(0.8)
    #pyautogui.press('f9') #update
    #sleep_mine(0.5)
    self.sort_sheet()
    #self.minimize_wps()
    trace_obj.epilogue(currentframe().f_code.co_name)
  
  def clear_windows_clipboard(self):
    from ctypes import windll
    if windll.user32.OpenClipboard(None):
     windll.user32.EmptyClipboard()
     windll.user32.CloseClipboard()
  #def bring_sheet_on_focus(sheet_open_status_known=False, sheet_open=False):
  #  if sheet_open_status_known == True:
  #    if sheet_open == True:
  #      focus_wps()
  #    else:
  #      open_sheet()
  #  else:
  #    open_sheet_if_not_open()

class tida:
  firefox_exe_path = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"
  tida_url = "https://tida.syntronic.com/login"
  trace_obj = trace()

  def __init__(self, debug=False):
    trace_obj.prologue(currentframe().f_code.co_name)
    self.debug = debug
    trace_obj.epilogue(currentframe().f_code.co_name)

  def tidaUpdate(self):
    pyautogui.hotkey('winleft', 'd')
    trace_obj.prologue(currentframe().f_code.co_name)
    subprocess.call([self.firefox_exe_path, "--safe-mode"])
    sleep(1)
    pyautogui.write(self.tida_url)
    pyautogui.press('enter')
    sleep(2)
    pyautogui.press('tab')
    pyautogui.press('tab')
    confirm = pyautogui.confirm("press enter?",\
                                title="T I D A",\
                                buttons=["OK"])
    pyautogui.press('enter')
    sleep(2)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.write("https://tida.syntronic.com/manageTimeStamps")
    sleep(0.5)
    confirm = pyautogui.confirm("press enter?",\
                                title="T I D A",\
                                buttons=["OK", "Cancel"])
    if confirm == "Cancel":
      return
    pyautogui.press('enter')
    sleep(4)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    pyautogui.press('tab')
    sleep(0.5)
    confirm = pyautogui.confirm("press enter?",\
                                title="T I D A",\
                                buttons=["OK", "Cancel"])
    if confirm == "Cancel":
      return
    pyautogui.press('enter')
    sleep(1)
    pyautogui.press('tab')
    #pyautogui.press('down')
    #pyautogui.press('down')
    #pyautogui.press('down')
    #pyautogui.press('down')
    sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write("8AM")
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write("4:10PM")
    pyautogui.press('tab')
    pyautogui.press('tab')
    confirm = pyautogui.confirm("press enter?",\
                                title="T I D A",\
                                buttons=["OK", "Cancel"])
    if confirm == "Cancel":
      return
    pyautogui.press('enter')
    if datetime.now().date == 1:
      confirm = pyautogui.confirm("You need to submit salary report, check avail and submit pls",\
                                title="S A L A R Y  R E P O R T",\
                                buttons=["OK"])
    sleep(1)
    confirm = pyautogui.confirm("please remember to sign out and close tab",\
                                title="LOGOUT REMINDER",\
                                buttons=["OK"])
    trace_obj.epilogue(currentframe().f_code.co_name)
