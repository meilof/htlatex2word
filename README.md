# htlatex2word
Convert LaTeX to Word (with formulas) using htlatex

## Requirements

 * Pywinauto (`pip install -U pywinauto`)
 * A LaTeX installation with `htlatex` (e.g., MikTeX)

## Usage

To convert `file.tex` to Word, perform the following steps:

* Run `python htlatex2word.py file` (so without the `.tex` suffix). This will create a file called `withoutmath.html`
* Open `withoutmath.html` in Microsoft Word. Make sure the "Navigation" pane is closed and the keyboard focus is on the text
* Return to the terminal where the Python script is running and press Enter
* The script will now fill in all math formulas in `withoutmath.html`. Wait until this is finished and **do not use your computer in the mean time** since the tool works by sending key presses to the Word window
* In Microsoft Word, just save `withoutmath.html` as a Word document and you are done!

## Problems

* The tool has a fixed timeout value of `0.05` seconds between replacements
  (and `0.1` for starting the search). If your computer is not fast enough
  this may cause problems. In this case, replace all occurences of `0.05` 
  and `0.1` in `htlatex2word.py` by some larger value
