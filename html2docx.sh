#!/bin/bash
echo
echo "Converting html to docx..."
osascript <<EOD
  set base_folder to "/Users/dnoriega/Github/rmarkdown2docx/"
  set file_in to base_folder & "example.html"
  set file_out to base_folder & "example.docx"

  tell application "Microsoft Word"
    activate
    open file_in

    set name of font object of Word style style heading1 of active document to "Calibri"
    set name of font object of Word style style heading2 of active document to "Calibri"
    set name of font object of Word style style heading3 of active document to "Calibri"
    set name of font object of Word style style normal of active document to "Calibri"
    set name of font object of Word style style emphasis of active document to "Calibri"

    set all_images to inline pictures of active document

    try
      repeat with img in all_images
        break link link format of img
      end repeat
    end try

    set view type of view of active window to print view

    save as active document file name file_out file format format document
    close active document
  end tell
EOD
osascript -e 'quit app "Microsoft Word"'

echo "...done"