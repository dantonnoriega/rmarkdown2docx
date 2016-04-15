#!/usr/local/bin/bash
# Original Applescript by Andrew Heiss (https://gist.github.com/andrewheiss/5bb905b1cfb244ebab40)
# SCRIPT MUST BE IN SAME DIRECTORY AS HTML FILES
# use: ~$ zsh ./html2docx.sh
# OR ~$ bash ./html2docx.sh
echo
html2docx () {
echo "Converting $2.html to $2.docx..."
osascript <<EOD
set base_folder to "$1/"
set file_in to base_folder & "$2.html"
set file_out to base_folder & "$2.docx"

tell application "Microsoft Word"
  activate
  open file_in

  # tables: change font, font size; add spacing and padding; format
  set tbls to tables of active document

  repeat with tbl in tbls
    set name of font object of text object of tbl to "Consolas"
    set font size of font object of text object of tbl to 9
    set alignment of row 1 of tbl to align row center
    select tbl
    set myRange1 to text object of selection
    set myRange1 to move start of range myRange1 by a character item count - 1
    set myRange2 to collapse range myRange1 direction collapse start
    set myRange3 to collapse range myRange1 direction collapse end
    select myRange2
    type text selection text "\n"
    select myRange3
    type text selection text "\n"
    set allow page breaks of tbl to False
    select tbl
    select (row 1 of selection)
    set alignment of paragraph format of selection to align paragraph center
    set enable borders of border options of tbl to true
    set BOR1 to get border tbl which border border horizontal
    set BOR2 to get border tbl which border border vertical
    set line style of BOR1 to line style single
    set line style of BOR2 to line style single
    set line width of BOR1 to line width100 point
    set line width of BOR2 to line width450 point
    set color index of BOR1 to white
    set color index of BOR2 to white
    set outside line style of border options of tbl to line style none
    set BOR3 to get border row 1 of tbl which border border bottom
    set line style of BOR3 to line style single
    set color index of BOR3 to gray25
  end repeat

  # adjust fonts
  set name of font object of Word style style heading1 of active document to "Calibri"
  set name of font object of Word style style heading2 of active document to "Calibri"
  set name of font object of Word style style heading3 of active document to "Calibri"
  set name of font object of Word style style normal of active document to "Calibri"
  set name of font object of Word style style emphasis of active document to "Calibri"

  # scale and center images then break links
  ## first, break links
  set all_images to inline pictures of active document

  repeat with img in all_images
    try
      break link link format of img
    end try
  end repeat

  ## second, adjust size and center
  set all_inline_images to inline pictures of active document

  repeat with img in all_inline_images
    try
      set lock aspect ratio of img to true
      set height of img to inches to points inches 4
      set alignment of horizontal line format of img to horizontal line align center
    end try
  end repeat

  # remove "Previous" and "Next" if last paragraph (happens in some html conversions)
  set LastParagraph to text object of last paragraph of active document
  select LastParagraph
  set selFind to find object of selection
  set forward of selFind to true
  set wrap of selFind to find stop
  set content of selFind to "Previous"
  execute find selFind
  delete selection
  set content of selFind to "Next"
  execute find selFind
  delete selection

  set view type of view of active window to print view

  save as active document file name file_out file format format document
  close active document
end tell

EOD
echo "...done"
}

#get path
PWD=$(pwd)

# convert all .html files do docx
## works with bash and zsh
find . -maxdepth 1 -type f -name '*.html' | while read F; do FF=${F#*/}; html2docx $PWD ${FF%.html}; done;

## works with zsh but not bash
# find . -maxdepth 1 -type f -name '*.html' | while read F; do html2docx $PWD ${${F#*/}%.html}; done;

#quit word
osascript -e 'quit app "Microsoft Word"'
