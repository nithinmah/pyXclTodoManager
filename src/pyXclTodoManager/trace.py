from sys import stdout 
class trace:
  running_indent = -1
  indent = {
  0: " ",
  1: "   ",
  2: "     ",
  3: "       ",
  4: "         ",
  5: "           ",
  6: "             ",
  7: "               " ,
  8: "                 ",
  9: "                   "
  }
 
  color_codes = {\
  "RED"     : "\033[1;31m",\
  "BLUE"    : "\033[1;34m",\
  "CYAN"    : "\033[1;36m",\
  "GREEN"   : "\033[0;32m",\
  "RESET"   : "\033[0;0m" ,\
  "BOLD"    : "\033[;1m"  ,\
  "REVERSE" : "\033[;7m"   \
  }

  def print_color_and_reset(self, text, color, reset=True):
    stdout.write(self.color_codes[color])
    print(text) 
    if reset == True:
      stdout.write(self.color_codes["RESET"])

  def prologue(self, function_name):
    #pdb.set_trace()
    self.running_indent = self.running_indent + 1
    prol_text =  self.indent[self.running_indent] + "--> " + function_name
    self.print_color_and_reset(prol_text, "GREEN")
  
  def epilogue(self, function_name):
    #pdb.set_trace()
    epi_text =  self.indent[self.running_indent] + "<-- " + function_name
    self.print_color_and_reset(epi_text, "GREEN")
    self.running_indent = self.running_indent - 1
  
  def trace(self, print_val):
    #pdb.set_trace()
    trace_text = self.indent[self.running_indent] + "Trace: " + print_val
    self.print_color_and_reset(trace_text, "CYAN")

  def header(self, header_text):
    self.print_color_and_reset("", "BOLD", False)
    self.print_color_and_reset("--------------------", "BLUE", False)
    self.print_color_and_reset(header_text, "BLUE", False)
    self.print_color_and_reset("--------------------", "BLUE")
