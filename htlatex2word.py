import clipboard
import os.path
import pywinauto
import re
import sys
from time import sleep
import subprocess

if not os.path.exists("tex4ht.env"):
    print "Creating empty tex4ht.env"
    fl = open("tex4ht.env", "w")
    fl.close()

subprocess.call(["htlatex", sys.argv[1] + ".tex", "xhtml,mathml"])

all=""

for ln in open(sys.argv[1] + ".html"): all += ln.strip() + " "

if all.find("footnote-mark"):
    print "*** WARNING: your file contains footnotes. These may contain MathML code, even"
    print "*** if they do not contain math. Please try removing them in case of problems."

maths = list(re.findall('<math.*?</math>', all))

fl = open("withoutmath.html", "w")
print >>fl, re.sub('<math.*?</math>', 'ASFLKJFASASFKJ', all)
fl.close()

print "Open \"withoutmath.html\" with Word and make sure the navigation pane is closed."
raw_input("Press Enter to continue")

app = pywinauto.Application().connect(title_re="withoutmath.html - Word")
main = app.window(title_re="withoutmath.html - Word")
main.Minimize()
main.Restore()
main.SetFocus()
sleep(0.1)

first = True

for math in maths:
    clipboard.copy(math)
    
    if first:
        pywinauto.keyboard.SendKeys("^f")
        sleep(0.05)
        pywinauto.keyboard.SendKeys("ASFLKJFASASFKJ")
        sleep(0.05)
        pywinauto.keyboard.SendKeys("{ENTER}")
        sleep(0.05)
        pywinauto.mouse.click(button='left', coords=(400,250))
        sleep(0.05)
        first = False
    else:
        pywinauto.keyboard.SendKeys("+{VK_F4}")
        sleep(0.05)
        
    pywinauto.keyboard.SendKeys("^v")
    sleep(0.05)
 