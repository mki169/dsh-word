on run argv
  set inPath to item 1 of argv
  set outPath to item 2 of argv
  set inFile to POSIX file inPath
  set targetHFS to inFile as text
  tell application "Microsoft Word"
    set theDoc to missing value
    open inFile
    repeat 240 times
      delay 0.5
      set docCount to count of documents
      repeat with i from 1 to docCount
        try
          if full name of document i is targetHFS then
            set theDoc to document i
            exit repeat
          end if
        end try
      end repeat
      if theDoc is not missing value then exit repeat
    end repeat
    if theDoc is missing value then error "Word did not open " & inPath
    save as theDoc file name outPath file format format PDF
    close theDoc saving no
  end tell
end run
