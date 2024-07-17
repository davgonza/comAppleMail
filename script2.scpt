using terms from application "Mail"
    on perform mail action with messages theMessages for rule theRule
        tell application "Mail"
            repeat with theMessage in theMessages

                set senderEmail to (sender of theMessage)
                set attachmentCount to count of mail attachments of theMessage
                set flagColor to (flag index of theMessage)

                if senderEmail contains "info@thehernandezcontractor.com" and flagColor is 1 then -- 2 represents the orange flag
                    display dialog "Processing email with " & attachmentCount & " attachments."
                    set flagged status of theMessage to false

                    -- Calculate the closest past Sunday's date
                    set today to current date
                    set dayOfWeek to (weekday of today) as string
                    set dayNumber to 0

                    if dayOfWeek is "Sunday" then
                        set dayNumber to 0
                    else if dayOfWeek is "Monday" then
                        set dayNumber to 1
                    else if dayOfWeek is "Tuesday" then
                        set dayNumber to 2
                    else if dayOfWeek is "Wednesday" then
                        set dayNumber to 3
                    else if dayOfWeek is "Thursday" then
                        set dayNumber to 4
                    else if dayOfWeek is "Friday" then
                        set dayNumber to 5
                    else if dayOfWeek is "Saturday" then
                        set dayNumber to 6
                    end if

                    set pastSunday to today - (dayNumber * days)

                    set theYear to year of pastSunday
                    set theMonth to month of pastSunday as integer
                    set theDay to day of pastSunday

                    -- Format the date as MM.DD.YYYY
                    set formattedDate to text -2 through -1 of ("0" & theMonth) & "." & text -2 through -1 of ("0" & theDay) & "." & theYear

                    -- Create the folder path
                    set destinationFolder to "/Users/abby/Library/Mobile Documents/com~apple~CloudDocs/HC2/Invoices and Timesheets/TECHSOURCE/" & formattedDate

                    -- Convert the POSIX path to HFS path
                    set hfsDestinationFolder to POSIX file destinationFolder as text

                    -- Check if the folder exists, if not, create it
                    do shell script "mkdir -p " & quoted form of destinationFolder

                    -- Collect all attachments into a list
                    set attachmentsList to mail attachments of theMessage



                    set nameToSave to ""
                    set currentSave to ""








                    repeat with i from 1 to (count of attachmentsList)
                        set theAttachment to item i of attachmentsList
                        set attachmentName to name of theAttachment
                        set savePath to hfsDestinationFolder & ":" & attachmentName

                        set dotPos to (offset of "." in (reverse of characters of attachmentName as string))

                        set baseName to text 1 thru -(dotPos + 1) of attachmentName

                        set newName to baseName & ".txt"

                        set txtPath to hfsDestinationFolder & ":" & newName

                        save theAttachment in file savePath

                        if (attachmentName contains "WE") and (attachmentName contains "24") then
                            set appPath to "/Users/abby/pdf_thing.app"
                            set posixPath to POSIX path of (savePath as alias)
                            do shell script "open -a " & quoted form of appPath & " " & quoted form of posixPath

                            delay 1

                            set test to "/Users/abby/tmp/" & newName

                            display dialog "here 2"

                            set txtContent to do shell script "cat " & quoted form of test

                            set textLines to paragraphs of txtContent

                            set variableA to item 4 of textLines
                            set variableB to item 5 of textLines


                            display dialog "Variable A: " & variableA
                            display dialog "Variable B: " & variableB

                            if not (variableB contains "-TS ") then
                                set variableB to item 6 of textLines
                            end if


                            set companyName to word 1 of variableA
                            set lastTwoWords to (item -2 of words of variableB) & " " & (item -1 of words of variableB)

                            set AppleScript's text item delimiters to ":"
                            set pathComponents to text items of savePath

                            set lastFolder to item -2 of pathComponents

                            -- WE 06.30.2024-Hazzard-HFC-TS 034
                            set weName to "WE " & lastFolder & "-" & companyName & "-" & lastTwoWords & ".pdf"

                            set AppleScript's text item delimiters to ":"
                            set directoryPath to (text items 1 through -2 of savePath) as string
                            set AppleScript's text item delimiters to ""

                            -- Construct the new file path
                            set newFilePath to directoryPath & "/" & weName

                            -- Rename the file
                            tell application "Finder"
                                set name of file savePath to weName
                            end tell

                            -- Look ahead at the next attachment
                            if i < (count of attachmentsList) then
                                set nextAttachment to item (i + 1) of attachmentsList
                                set nextAttachmentName to name of nextAttachment
                                set currentSave to hfsDestinationFolder & ":" & nextAttachmentName
                                set newSave to hfsDestinationFolder & ":"

                            end if

                        end if
                    end repeat



                    repeat with i from 1 to (count of attachmentsList)
                        set theAttachment to item i of attachmentsList
                        set attachmentName to name of theAttachment

                        if not (attachmentName contains "WE" and attachmentName contains "24") then

                            set stName to "ST-" & nameToSave

                            set newSave to newSave & nameToSave


                            tell application "Finder"
                                set name of file currentSave to newSave
                            end tell

                        end if
                    end repeat








                end if
            end repeat
        end tell
    end perform mail action with messages
end using terms from
