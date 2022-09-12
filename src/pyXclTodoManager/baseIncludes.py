from trace import trace
import pdb
from sys import stdin 
from sys import argv 
from inspect import currentframe
from time import sleep
from datetime import datetime, timedelta

tasks_xls_path="C:\\Users\\ethimah\\Documents\\Pers\\gd\\Tasks\\tasktracker.xlsx"
trace_obj = trace()

def sleep_mine(sleep_duration):
  trace_obj.prologue(currentframe().f_code.co_name)
  trace_obj.trace("sleeping for " + str(sleep_duration))
  sleep(sleep_duration)
  trace_obj.epilogue(currentframe().f_code.co_name)

def get_input_from_user(prompt_string):
  #import msvcrt
  #while msvcrt.kbhit():
  #  msvcrt.getch()
  try:
    user_input = input(prompt_string)
    return user_input
  except EOFError:
    print("EOFError")
