#!/usr/bin/env python
from baseIncludes import *

import win32gui
import pyautogui
import pymsgbox
import os
import subprocess
from subprocess import PIPE

from appWindows import wpsExcelSheet
from appWindows import tida
from appWindows import start_monitor_process_in_next_tab
from task_workbook import task_workbook
from task_workbook import get_input_from_user 

#wb = openpyxl.load_workbook(tasks_xls_path, data_only=True)
#wb = openpyxl.load_workbook(tasks_xls_path)
#shortlist = wb["ShortList"]
#reminders = {} 
#tasks_rows_map = {}

class reminder:
  tida_flag_file="C:/Users/ethimah/tmp/tida_flag.txt"
  did_tida_for_today = subprocess.run("cat " + tida_flag_file,\
    shell=True, check=True, stdout=PIPE).stdout.decode("ascii").strip()
  __reminders = {} 
  __tasks_rows_map = {}
  __next_step_map = {}
  trace_obj = trace()
  did_any_reschedule = False

  def __init__(self, debug=False):
    trace_obj.prologue(currentframe().f_code.co_name)
    self.debug = debug
    self.plus_all = 0
    self.xcelSheet = wpsExcelSheet()
    #on reminder's last close we might have updated the sheet, so sort now!
    #We sort before close, so no neeed now.
   #self.xcelSheet.update_formulas_on_tracker_sheet()
    self.wb = task_workbook()
    self.tida = tida()
    #trace_obj.trace("did_tida_for_today : " + self.did_tida_for_today)
    trace_obj.epilogue(currentframe().f_code.co_name)

  def tida_check(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    #pdb.set_trace()
    if (datetime.now().hour >= 16) and \
       (datetime.now().minute > 10) and \
       (datetime.today().weekday() < 5) :
      trace_obj.trace("did_tida_for_today : " + self.did_tida_for_today)
      if self.did_tida_for_today == "notdone":
        trace_obj.trace("did_tida_for_today : " + self.did_tida_for_today)
        tida_ques="Time for Tida. Fill now?"
        confirm = pyautogui.confirm(tida_ques,\
                                    title="T I D A",\
                                    buttons=["OK", "Cancel"],\
                                    timeout=20000)
        #print("confirm is: " + confirm) 
        if confirm == "OK" or confirm == "Cancel":
          self.tida.tidaUpdate()
          self.did_tida_for_today = "done"
          subprocess.run("echo done > " + self.tida_flag_file,\
            shell=True, check=True)
    trace_obj.epilogue(currentframe().f_code.co_name)

  def setup_reminders_from_wb(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    self.__reminders.clear()
    row_id = 2
    task = self.wb.name(row_id)
    AP_date = self.wb.action_date(row_id)
    while ((AP_date.date() - datetime.now().date()).days <= 0) :
      trace_obj.trace("task: "+ task + " days left: "+ str((AP_date.date() - datetime.now().date()).days))
      remind_time = self.wb.reminder_time(row_id)
      self.__reminders[task] = [remind_time.hour, remind_time.minute]
      self.__tasks_rows_map[task] = row_id
      self.__next_step_map[task] = self.wb.next_step(row_id)
      row_id = row_id + 1;
      task = self.wb.name(row_id)
      AP_date = self.wb.action_date(row_id)
    #for key in sorted(self.reminders, key=self.reminders.get):
    #  remind_at_hour, remind_at_minute = self.reminders[key]
    #  trace("Adding reminder for " + key + " : " + str(remind_at_hour) + ":" + str(remind_at_minute))
    trace_obj.epilogue(currentframe().f_code.co_name)

  def check_sheet_requires_update_or_has_been_updated(self, last_check_time):
    trace_obj.prologue(currentframe().f_code.co_name)
    xls_last_modify_time = \
      datetime.fromtimestamp(os.path.getmtime(tasks_xls_path))
    trace_obj.trace("Checking last modify date, xls_last_modify_time: " + xls_last_modify_time.strftime("%c"))
    if (xls_last_modify_time.date() < last_check_time.date()):
      subprocess.run("echo notdone > " + self.tida_flag_file,\
        shell=True, check=True)
      self.did_tida_for_today = "notdone"
      trace_obj.trace("xcel sheet not updated today, updating")
      self.update_reminders_for_today()
    else:
      trace_obj.trace("Checking last modify time, last_check_time: " + last_check_time.strftime("%c"))
      if (xls_last_modify_time > last_check_time):
        trace_obj.trace("xcel sheet updated, need to close app and restart !!! ")
        sleep_mine(1)
        exit(0)
    trace_obj.epilogue(currentframe().f_code.co_name)

  def _spit_trace_header(self, action, datetime):
    trace_obj.header(action + " at : " + datetime.strftime("%c"))

  def remind_and_resched_task_if_wished(self, reminder_text, key, only_resched=False):
    trace_obj.prologue(currentframe().f_code.co_name)
    reschedule_map = {"5 Min": 5, "10 Min": 10, "30 Min": 30, "1 Hr": 60, \
                      "2 Hr": 120, "4 Hr": 240, "6 Hr": 360, "8 Hr": 480, "10 Hr": 600}
    reschedule_buttons = \
    ["OK", "5 Min", "10 Min", "30 Min",\
     "1 Hr", "2 Hr", "4 Hr", "6 Hr", "8 Hr", "10 Hr",\
     "Tomm", "Done", "Next Cycle", "+30 All", "+60 All"]
    if self.debug == True:
      pdb.set_trace()
    if only_resched == False:
      timeout = 60000
    else:
      timeout = 600000
    if self.plus_all == 0:
      reschedule_or_ok = pyautogui.confirm(text=reminder_text, title="PyReminder",\
        buttons=reschedule_buttons, timeout=timeout)
    else:
      if self.plus_all == 30:
        reschedule_or_ok = "30 Min"
      else:
        reschedule_or_ok = "1 Hr"
    resched_row_id = self.__tasks_rows_map[key]
    if reschedule_or_ok != "OK" and reschedule_or_ok != "Timeout":
      trace_obj.trace("rescheduling by " + reschedule_or_ok + " task: " + key)
      if self.did_any_reschedule == False:
        self.xcelSheet.close_sheet()
        self.did_any_reschedule = True
    match reschedule_or_ok:
      case "Tomm":
        self.wb.update_action_date_delta_for_task(resched_row_id, 1)
      case "5 Min" | "10 Min" | "30 Min" |\
           "1 Hr" | "2 Hr" | "4 Hr" | "6 Hr" | "8 Hr" | "10 Hr":
        self.wb.update_reminder_time_delta_for_task(resched_row_id, \
          reschedule_map[reschedule_or_ok])
        self.wb.update_action_date_delta_for_task(resched_row_id, 0)
      case "Done":
        self.wb.update_action_date_delta_for_task(resched_row_id, 120)
      case "Next Cycle":
        self.wb.update_due_month_delta_for_task(resched_row_id, 1)
      case "+30 All":
        self.plus_all = 30
        self.wb.update_reminder_time_delta_for_task(resched_row_id, \
          reschedule_map["30 Min"])
        self.wb.update_action_date_delta_for_task(resched_row_id, 0)
      case "+60 All":
        self.plus_all = 60
        self.wb.update_reminder_time_delta_for_task(resched_row_id, \
          reschedule_map["1 Hr"])
        self.wb.update_action_date_delta_for_task(resched_row_id, 0)
      case other:
        print("unexpected reschedule_or_ok: " + reschedule_or_ok)
        #should be only OK or Timeout
    new_date = self.wb.action_date(resched_row_id)
    new_remind_time = self.wb.reminder_time(resched_row_id)
    trace_obj.trace("updated reminder: " + key + " : " + datetime.strftime(new_date, "%c") +\
      " : " + str(new_remind_time)) 
    trace_obj.epilogue(currentframe().f_code.co_name)

  # Clear the reminders lest ones edited out fo the window linger.
  def update_reminders_for_today(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    #update_formulas_on_tracker_sheet()
    self.setup_reminders_from_wb()
    row_id = 2
    task = self.wb.name(row_id)
    AP_date = self.wb.action_date(row_id)
    #trace("type of date read from row: " + str(row_id) + " is: " + str(type(AP_date)))
    alerted = False
    while ((AP_date.date() - datetime.now().date()).days <= 0) :
      #trace("task: "+ task + " days left: "+ str((AP_date.date() - datetime.now().date()).days))
      if (AP_date.date() - datetime.now().date()).days < 0 :
        if alerted == False:
        #  pyautogui.alert('rows need updates')
          alerted = True
          self.xcelSheet.close_sheet()
        #self.wb.update_outdated_row(row_id)
        date_update_text = task + " needs date update \n" + self.__next_step_map[task]
        self.remind_and_resched_task_if_wished(date_update_text, task, only_resched=True)
      row_id = row_id + 1;
      task = self.wb.name(row_id)
      AP_date = self.wb.action_date(row_id) #save wb to sheet!
      #trace("type of date read from row: " + str(row_id) + " is: " + str(type(AP_date)))
    self.plus_all = 0
    if alerted == True:
      self.xcelSheet.update_formulas_on_tracker_sheet()
      exit(0) #save wb to sheet!
    trace_obj.epilogue(currentframe().f_code.co_name)

  def row_update(self, row_id):
    task = self.wb.name(row_id)
    hour_min = self.wb.reminder_time(row_id).strftime("%H:%M")
    text = str(row_id) + ". " + task + ": " + hour_min + "\n" + self.__next_step_map[task]
    self.remind_and_resched_task_if_wished(text, task)  
    

  def reminder_routine(self):
    trace_obj.prologue(currentframe().f_code.co_name)
    self.xcelSheet.set_sheet_state("unknown")
    last_check_time = datetime.now()
    self._spit_trace_header("Setting up", last_check_time)
    self.update_reminders_for_today()
    while True:
      reminder_number = 0
      did_any_reschedule = False
      self.xcelSheet.set_sheet_state("unknown")
      self._spit_trace_header("Checking", datetime.now())
      self.tida_check()
      self.check_sheet_requires_update_or_has_been_updated(last_check_time)
      last_check_time = datetime.now()
      for key in sorted(self.__reminders, key=self.__reminders.get):
        remind_at_hour, remind_at_minute = self.__reminders[key]
        hour_now = datetime.now().hour
        minute_now = datetime.now().minute
        row_id = self.__tasks_rows_map[key]
        trace_obj.trace("checking at row: " + str(row_id)+ " task: " + key +\
          " Hour: " + str(remind_at_hour) + " Min: " + str(remind_at_minute))
        if (remind_at_hour < hour_now) or (remind_at_hour == hour_now and remind_at_minute <= minute_now):
          reminder_number = reminder_number + 1
          text = str(reminder_number) + ". " + key + ": " + str(remind_at_hour) +\
            ":" + str(remind_at_minute) + "\n" + self.__next_step_map[key]
          self.remind_and_resched_task_if_wished(text, key)  
        else: #the list is sorted, so on first item failing check just break
          break
      self.plus_all = 0
      if self.did_any_reschedule == True: 
        self.xcelSheet.update_formulas_on_tracker_sheet()
        #update_formulas will rearrange/modify the sheet, exit now to re read.
        #update last_check_time so that we dont quit for our own change!
        #last_check_time = datetime.now()
        exit(0)
      trace_obj.trace("Sleeping at " + datetime.now().strftime("%c"))
      sleep(240)
     #sleep(60)

from infi.systray import SysTrayIcon
systray = SysTrayIcon("reminder.ico", "Example tray icon")

def reminder_proc_main(stdin_fn, debug=False):
  try:
    stdin = os.fdopen(stdin_fn)
    reminder_obj = reminder(debug)
    reminder_obj.reminder_routine() #
    #exit(1)
  except (KeyboardInterrupt):
    print('Interrupted')
    systray.shutdown()
    exit(0)

def monitor():
  systray.start()
  from multiprocessing import Process
  start_monitor_process_in_next_tab()
  while 1:
    stdin_fn = stdin.fileno()
    proc = Process(target=reminder_proc_main, args=(stdin_fn,))
    proc.start()
    proc.join()
    systray.shutdown()
    if proc.exitcode != 0:
      pyautogui.confirm('something went wrong... terminating', buttons=["OK"])
      exit(proc.exitcode)

if __name__ == '__main__':
    try:
        if len(argv) == 1:
          monitor()
        else:
          stdin_fn = stdin.fileno()
          reminder_proc_main(stdin_fn, debug=True)
    except (KeyError, ValueError):
        print('usage: test.py sleeptime')
        exit(1)
