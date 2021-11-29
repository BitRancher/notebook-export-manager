--
-- Notebook Export Manager: An Applescript to export data from Evernote
-- 
-- Copyright (c) 2021 William C Jacob Jr
-- Licensed under the MIT license. See accompanying documentation.
--
-- This works with the version of evernote installed from their website,
-- but not with the Apple Mac App Store version 10 which doesn't
-- support AppleScript at this time (Feb 2021).
-- Runs successfully on macOS Catalina (10.15.7) and Big Sur (11.1) with Evernote 7.14
-- 0.4.2 fix locale problem caused by date in string literal
-- 0.5.0 fix file Kind search for macOS Monterery (12.0.1): HTML files on disk now
--       show file Kind "HTML document", not "HTML text"

-- To get Finder write
use scripting additions
-- To get NSString access in urlEncode handler
-- use framework "Foundation"     removed in 0.3.2

--
-- Users may wish to change these initial values
--
global logProgress -- switch to control Log statements for debugging
set logProgress to false -- run Log commands for debugging USER SETTING
global alwaysBuildTopIndex
set alwaysBuildTopIndex to true -- force rebuild of top-level index USER SETTING
-- The following setting might be useful if you make changes to the buildNoteIndex handler
-- and you want to rebuild the index pages for each notebook without actually exporting
-- all the notes from Evernote again.
global alwaysBuildBookIndexes
set alwaysBuildBookIndexes to false -- force rebuild of notebook indexes USER SETTING
global myPListFullPath -- PList full path + file name
-- The Property List file for this application (USER SETTING):
-- Store it in [my home folder]/Library/Preferences:
set myPListFullPath to POSIX path of ((path to preferences as text) & "com.lake91.nem.plist") -- Folder for preferences
-- Or, store it in Downloads, just for testing (comment this out to use standard location)
-- set myPListFullPath to POSIX path of ((path to downloads folder as text) & "com.lake91.enexp.plist") -- Test folder

-- Global variables
--
global parentTagNames, parentTagHTML -- cache of tag names
set parentTagNames to {} -- list of tag names already looked up
set parentTagHTML to {} -- corresponding list of parent tag HTML

global offsetGMT -- offset to GMT in seconds
set offsetGMT to time to GMT -- get offset to GMT in seconds
global headerInfo -- system info to store as comment in index pages
set headerInfo to "" -- initially empty
global plExportTypeDesc -- Descriptions for options
global myCSSfile, myJSfile, myTagFile, myTopIndex -- file names for top-level folder
global nbIndexFile -- filename for notebook index pages
global FolderHTML, FolderENEX, AllNotebooks, SearchString -- Properties
global ExportAsHTML, ExportAsENEX, TimeoutMinutes, BooksSelected, ExportDate -- Properties 
global frontProcess -- name of our process
set myCSSfile to "enstyles.css" -- CSS file name for html export folder
set myJSfile to "enscripts.js" -- JS file name for html export folder
set myTagFile to "entags.html" -- master tag list file for html export folder
set myTopIndex to "index.html" -- name of the top-level index page listing all notebooks
set nbIndexFile to "index.html" -- name of the notebook index page listing all notes

-- For internal timing
global timerAll, timerSub1
set timerAll to 0
set timerSub1 to 0

if logProgress then log "Property List file:  " & myPListFullPath

set plExportTypeDesc to {"Export only in HTML format", "Export only in ENEX format", "Export in both HTML and ENEX formats"}

-- Set initial defaults for properties
set FolderHTML to path to documents folder as text -- Path for HTML exports
set FolderENEX to path to documents folder as text -- Path for ENEX exports
set AllNotebooks to true -- export all notebooks, not just some
set SearchString to "" -- assume no user-supplied search string
set ExportAsHTML to true -- Export notebooks in HTML format
set ExportAsENEX to true -- Export notebooks in ENEX format
set BooksSelected to {} -- List of notebooks to export
set ExportDate to initDate(2000, 1, 1) -- No exports have been done yet  0.4.2
set TimeoutMinutes to 30 -- Default timeout

-- Wake up Evernote and then switch back to this script
tell application "System Events" to set frontProcess to the name of every process whose frontmost is true -- get our process name
if logProgress then log "frontProcess: " & frontProcess
if application "Evernote" is not running then
	activate application "Evernote" -- start evernote if not already running
	delay 5 -- give it a few seconds to crank up
	activate frontProcess -- Switch back to our app so user sees the prompts
	delay 2
end if
--
-- Get Properties from PList file
-- Create Piist if necessary and Get current values from Property List file
--
set plistPrompt to not managePlist(1) -- Get PList values (returns false if the PList file did not exist)
activate frontProcess -- Switch back to our app so user sees the prompts
delay 1
--
-- Loop to prompt the user to start the export or to adjust the settings
--
repeat -- Until user clicks Export Now or Cancel
	if plistPrompt then -- If PList is new or it's a second pass through this loop
		-- Show Properties and allow user to make changes
		if not managePlist(2) then return -- update PList; quit if user cancels
		activate frontProcess -- Switch back to our app so user sees the prompts
		delay 1 -- give it a second
	end if
	set plistPrompt to true -- If we loop in the repeat, let user update PList
	set msg to "Will consider " & cond(AllNotebooks, "all notebooks", ("only " & (count of BooksSelected) & " selected notebooks")) & " for export to
   HTML: " & cond(ExportAsHTML, POSIX path of FolderHTML, "(not to be exported)") & "
   ENEX: " & cond(ExportAsENEX, POSIX path of FolderENEX, "(not to be exported)") & "
" & cond(length of SearchString > 0, "Only notes meeting this criteria will be exported:
" & SearchString, "All notes will be exported in each notebook") & "

Start export now or change settings."
	if ("Export Now" = (button returned of (display dialog msg buttons {"Export now", "Change settings", "Cancel"} default button 1 cancel button 3 with title "Ready to Export" with icon note))) then exit repeat
end repeat

set timerAll to getFineTime() -- get time at start ot run
--
-- Build a list of notebooks to consider, either all of them or a selected list
--
set exportBooks to BooksSelected -- assume a user-selected list
if AllNotebooks then set exportBooks to getBooks() -- get sorted list of all notebooks
--
-- Check that the volumes containing the output folders are mounted (eg, if on a USB stick)
--
if ExportAsENEX then
	if not checkRunVol(FolderENEX, "ENEX") then return -- check volume and folder; quit if not found
end if
if ExportAsHTML then
	if not checkRunVol(FolderHTML, "HTML") then return -- check volume and folder; quit if not found
end if
--
-- Set up counters for the main processing loop
--
with timeout of (TimeoutMinutes * 60) seconds -- run this code with a timeout
	set countENEXtried to 0 -- init count of attempted ENEX exports
	set countHTMLtried to 0 -- init count of attempted HTML exports
	set countENEXdone to 0 -- init count of completed ENEX exports
	set countHTMLdone to 0 -- init count of completed HTML exports
	set countBooksAttempted to length of exportBooks -- count of notebooks being attempted
	set ExportDate to current date -- Set date of main index rebuild
	set progress total steps to (length of exportBooks) -- set up progress display
	set progress completed steps to 0
	set progress description to "Exporting notebooks"
	set pcounter to 0
	--
	-- This is the main loop for this script
	-- Export each eligible notebook in sequence
	--
	repeat with aBook in exportBooks -- repeat with each notebook name
		if ExportAsENEX then -- if ENEX export was requested
			set progress additional description to "Exporting " & aBook & " as ENEX" -- set progress display detail
			set countENEXtried to countENEXtried + 1 -- count attempts
			if exportNotes(aBook, (POSIX path of FolderENEX) & my repl(aBook, "/", "_") & ".enex", "ENEX") then set countENEXdone to countENEXdone + 1 -- export as ENEX anc count successes
		end if
		if ExportAsHTML then -- if HTML export was requested
			set progress additional description to "Exporting " & aBook & " as HTML" -- set progress display detail
			set countHTMLtried to countHTMLtried + 1 -- count attempts
			set outHTML to (POSIX path of FolderHTML) & my repl(aBook, "/", "_") & "/" -- build HTML export folder path
			set expSome to exportNotes(aBook, outHTML, "HTML") -- Try to export notebook
			if expSome then set countHTMLdone to countHTMLdone + 1 -- count successful exports
			if expSome or alwaysBuildBookIndexes then buildNoteIndex(aBook, outHTML, expSome) -- build the index.html file for this notebook
		end if
		set pcounter to pcounter + 1 -- increment the progress display counter
		set progress completed steps to pcounter -- and ask Applescript to display it
	end repeat
	--
	-- Build the top-level index.html page if anything was exported
	-- or if it doesn't exist or if it is to be built every time.
	--
	tell application "Finder" to set gotIndex to (exists (FolderHTML & myTopIndex)) -- does top index exist ?
	if (countHTMLdone > 0) or (not gotIndex) or alwaysBuildTopIndex then buildTopIndex() -- rebuild top index if missing or outdated
	if (countENEXdone > 0) or (countHTMLdone > 0) then set x to managePlist(3) -- Update the Property List file with new Export Date
	set progress total steps to 0 -- shut down the progress display
	set progress completed steps to 0
	set progress description to ""
	set progress additional description to ""
	
	set timerAll to getFineTime() - timerAll -- elapsed seconds for run
	-- display alert "Run Time: " & timerAll & "   Sub Time: " & timerSub1
	--
	-- Display results and exit
	--
	display alert cond((countENEXdone + countHTMLdone) > 0, "Some notebooks exported", "No notebooks exported") message ((countBooksAttempted as text) & " Notebooks considered
" & countHTMLdone & " of " & countHTMLtried & " exported as HTML and
" & countENEXdone & " of " & countENEXtried & " exported as ENEX

(Export time (secs): " & (timerAll as integer) & ", index read: " & (timerSub1 as integer) & ")")
end timeout
return -- end execution of this script

--
-- Select notes and export them
--   bName is the notebook name (text)
--   oFIle is the full pathname of the output file (ENEX) or folder (HTML)
--   oType is either "ENEX" or "HTML"
-- Returns True if some were exported; else False
-- A lot of logic in this handler is devoted to handling notebooks whose
--   name contains a quote (").
-- 
on exportNotes(bName, oFile, oType)
	local f, fileExists, datesrch, enexp, srch, someNotes
	local selectedNotes, d, cu, b, c, i, m, w, j
	if logProgress then log "Exporting notebook: " & bName & "  to file: " & oFile & "  as " & oType
	tell application "Evernote" to set f to exists notebook bName -- does notebook exist?
	if not f then return false -- if not, stop now and return
	-- set QuoteInName to cond((offset of "\"" in bName) > 0, true, false) -- notebook name contains "
	tell application "Finder" to set fileExists to (exists oFile as POSIX file) -- check for output file existence
	set datesrch to "" -- Assume this notebook has never been exported
	set enexp to "" -- assume no export-date found in existing exported notebook
	if fileExists then -- if output file already exists
		--		if oType = "ENEX" then tell application "System Events" to set enexp to value of XML attribute "export-date" of XML element "en-export" of XML file oFile -- Get export date/time from inside ENEX file (this was very slow for large ENEX files and was replaced in ver  0.3.2)
		if oType = "ENEX" then set enexp to item 1 of readIndex(POSIX file oFile, {"<en-export export-date=\""}, {"\""}, 500) -- Get export date/time from inside ENEX file
		if oType = "HTML" then set enexp to item 1 of readIndex(POSIX file (oFile & nbIndexFile), {"<!-- bookexported "}, {" bookexported -->"}, 500) -- get export from HTML comment
		if (length of enexp) > 0 then set datesrch to " updated:" & enexp -- Set export date clause for search string
	end if
	if (offset of "\"" in bName) = 0 then -- if normal notebook name (no " in it)
		set srch to "notebook:\"" & bName & "\" " & SearchString -- Build basic search string for EXPORT
		tell application "Evernote" to set someNotes to find notes srch & datesrch -- any updated notes?
		if logProgress then log oType & " note count for: " & bName & "  is: " & (count of someNotes) & " with search: " & srch & datesrch
		if (count of someNotes) = 0 then return false -- If no new notes, don't export this notebook
		tell application "Evernote" to set selectedNotes to find notes srch -- get list of notes to export
	else -- there is a " in the notebook name - no search allowed
		--		tell application "Evernote" to set selectedNotes to every note in notebook bName -- export all notes
		if enexp ≠ "" then -- if only updated notes
			set d to current date -- init date object
			tell d to set {its year, its month, its day, its time} to {0 + (text 1 thru 4 of enexp), 0 + (text 5 thru 6 of enexp), 0 + (text 7 thru 8 of enexp), (hours * (0 + (text 10 thru 11 of enexp))) + (minutes * (0 + (text 12 thru 13 of enexp))) + (0 + (text 14 thru 15 of enexp))} -- convert date string into type date
			set d to d + offsetGMT -- adjust GMT export-date to local time
			if logProgress then log "enexp: " & enexp & "  local: " & d
			tell application "Evernote" to set cu to count of (every note in notebook bName whose modification date ≥ d) -- get count of updated notes in this notebook
			if cu = 0 then return false -- if no changed notes in notebook, don't export it
		end if
		tell application "Evernote"
			if SearchString = "" then -- did user specify a search string?
				set selectedNotes to every note in notebook bName -- if not, export every note in notebook
			else
				set b to find notes SearchString -- search all of evernote with user-specified search
				set selectedNotes to {} -- init list of notes to export
				set c to every note in notebook bName -- get all notes in notebook
				repeat with i from 1 to (length of c) -- with every note in notebook
					set m to item i of c -- get the note from the notebook
					repeat with j from 1 to length of b -- with every note in search str results
						set w to item j of b -- get note from search results
						if m is w then -- if note is in both lists
							copy m to end of selectedNotes -- copy it to the result list
							exit repeat
						end if
					end repeat
				end repeat
			end if
		end tell
		if logProgress then log oType & " note count for: " & bName & "  is: " & (count of selectedNotes) & " with quoted notebook name"
	end if
	if (count of selectedNotes) = 0 then return false -- If no notes to export, return empty list
	if oType = "ENEX" then tell application "Evernote" to export selectedNotes to oFile format ENEX with tags -- export notes as enex
	if oType = "HTML" then tell application "Evernote" to export selectedNotes to oFile format HTML -- or as HTML
	return true -- return: aome notes exported
end exportNotes

--
-- Build the index.html file for a notebook, replacing the simple
-- one produced by the evernote Export as HTML operation
-- Input: Notebook name (string), Output folder name (string),
--      Did export just occur (boolean)
-- Returns nothing
--
on buildNoteIndex(bookName, outHTML, wasExported)
	--
	-- Initialize variables before processing each note in the notebook
	--
	local bookMod, bookCreate, htmlNotes, tagRows, yy, noteBookCell, pfile, listNoteFiles
	local noteAttrs, noteHTMLfile, sNoteTitle, fileEnc, notelink, noteLinkFar
	local aNoteMod, aNoteCreate, tagsComma, allTagNames, aNoteTagList
	local aTageName, seenTag, parentCell, loopTag, spanTag, sNoteAttach, dirSize
	local aNoteAuthor, aNoteExporter, fOff, lOff, sz, expStamp, priorEnviron, priorExportDate
	local ourTimer, bookNotes, bookFiles
	-- tell application "Evernote" to set bookType to the notebook type of notebook bookName -- get type of notebook
	set bookMod to initDate(2000, 1, 1) -- default book mod date  0.4.2
	set bookCreate to initDate(2040, 1, 1) -- default book create date  0.4.2
	set htmlNotes to "" -- Building HTML for Notes list
	set tagRows to "" -- Building HTML for rows with Tag info
	set bookNotes to 0 -- init count of notes in this notebook
	set bookFiles to 0 -- init count of attachments
	set yy to my urlEncode(my repl(bookName, "/", "_")) & "/" & nbIndexFile -- index file for this notebook
	set yy to my repl(yy, ":", "%3A") -- handle filename with :
	set noteBookCell to "      <td class='linkhover nbname'><a class='linktarg' href=\"" & yy & "\">" & my textHT(bookName) & "</a></td>
"
	-- Get a list of all .html files in the exported folder for this notebook, except for index.html
	try
		-- tell application "Finder" to set listNoteFiles to every item of (POSIX file outHTML as alias) whose ((kind is "HTML text") and (name ≠ nbIndexFile)) -- comment 0.3.5
		set yy to alias (outHTML as POSIX file)
		tell application "System Events" to set listNoteFiles to ((every file of yy) whose (((kind is "HTML text") or (kind is "HTML document")) and (name ≠ nbIndexFile))) -- get every note html file 0.5.0
	on error
		return -- if no files found, return without building the notebook index
	end try
	if (count of listNoteFiles) = 0 then return -- if no files, return  0.3.5
	--
	-- If this notebook was just exported, use today's date as the export date. If not
	-- get the actual export date from the existing index.html file.
	-- 
	set expStamp to current date -- assume this notebook was actually just exported
	set priorEnviron to "" -- assume no prior "environment" comment exists
	set priorExportDate to "" -- assume no prior export date
	if not wasExported then -- if it wasn't just exported ...
		set yy to readIndex(POSIX file (outHTML & nbIndexFile), {"<!-- bookexported ", "<!-- Exported from Evernote "}, {" bookexported -->", "   -->"}, 500) -- Get export date and environment comment from old index.html
		if length of (item 1 of yy) > 0 then
			set priorExportDate to item 1 of yy -- remember export date as string
			set expStamp to getDateSrch(item 1 of yy) -- get export date from previous index file
			set priorEnviron to item 2 of yy -- remember previous environment comment
		end if
	end if
	--
	-- Read the first few hundred characters of every exported note file (html) to get metadata
	-- for that note and create appropriate table rows for the index file.
	--
	repeat with noteHTMLfile in listNoteFiles -- examine each note in the current notebook
		-- set noteAttrs to getMeta(noteHTMLfile as text) -- get Meta tags from the note html file 0.3.5
		set noteAttrs to getMeta2(noteHTMLfile) -- get Meta tags from the note html file  0.3.5
		set aNoteTitle to nname of noteAttrs -- get note title from meta data
		set aNoteAuthor to author of noteAttrs -- get author from meta data
		set fOff to offset of "<" in aNoteAuthor -- get position of email start char
		set lOff to offset of ">" in aNoteAuthor -- and end char
		if (fOff > 2) and ((offset of "@" in aNoteAuthor) > fOff) and (lOff > (offset of "@" in aNoteAuthor)) then -- if Author field contains something plus an email address
			set aNoteAuthor to "      <td title='" & textHT(text (fOff + 1) thru (lOff - 1) of aNoteAuthor) & "'>" & textHT(trim(text 1 thru (fOff - 1) of aNoteAuthor)) & " ✉️</td>
"
		else
			set aNoteAuthor to "      <td>" & cond(aNoteAuthor = "", "&nbsp;", textHT(aNoteAuthor)) & "</td>
"
		end if
		set aNoteExporter to exporter of noteAttrs -- evernote version
		set fileEnc to my urlEncode(name of noteHTMLfile) -- get current note file name (no path)
		set fileEnc to my repl(fileEnc, "/", "%3A") -- handle filename with : which evernote saves as /
		-- Build table cell containing link to Note html file (one for note index and one for top-level tag index)
		set notelink to "      <td class='linkhover slink'><a class='linktarg' href=\"" & fileEnc & "\">" & my textHT(aNoteTitle) & "</a></td>
"
		set noteLinkFar to "      <td class='linkhover nbname'><a class='linktarg' href=\"" & my repl(my urlEncode(my repl(bookName, "/", "_")), ":", "%3A") & "/" & fileEnc & "\">" & my textHT(aNoteTitle) & "</a></td>
"
		set aNoteMod to getDateMeta(updated of noteAttrs) -- get note update date (as class date)
		set aNoteCreate to getDateMeta(created of noteAttrs) -- get note create date (as class date)
		if aNoteMod > bookMod then set bookMod to aNoteMod -- track most recently changed note in notebook
		if aNoteCreate < bookCreate then set bookCreate to aNoteCreate -- and earliest created note in notebook
		set tagsComma to (tags of noteAttrs) as text -- comma separated list of tags from note html meta tags
		if tagsComma = "" then -- if no tags
			set allTagNames to "&nbsp;" -- placeholder for table cell
		else -- process the tags assigned to this note
			-- 
			-- Separate the string of tag names into a list of tags.  Then
			-- figure out if each tag has a parent tag. Keep a cache of tags
			-- we've already checked so we don't have to ask Evernote again and again.
			-- Create table cells with the tag and with the parent tags (if any)
			--
			set AppleScript's text item delimiters to ", "
			set aNoteTagList to every text item of tagsComma
			set AppleScript's text item delimiters to ""
			set allTagNames to "" -- initialize a string of tag names for the note detail line
			repeat with aTagName in aNoteTagList -- process each tag
				-- set aTagName to name of aTag -- name of current tag
				set seenTag to my item_position({aTagName}, parentTagNames) -- seen tag before?
				if seenTag = 0 then -- have not seen this tag before
					set parentCell to "" -- init hierarchy of tag names
					-- loop up through linked parent tags (if any) 
					tell application "Evernote"
						set loopTag to tag aTagName
						repeat while (parent of loopTag) is not missing value
							set parentCell to my fmtTag(name of loopTag) & parentCell -- pre-pend parent tag
							set loopTag to parent of loopTag
						end repeat
						set parentCell to my fmtTag(name of loopTag) & parentCell -- pre-pend parent tag
					end tell
					copy parentCell to end of parentTagHTML -- cache this parent tag info
					copy (aTagName) to end of parentTagNames -- and this tag
				else -- get tag parent string from previously cached list
					set parentCell to (item seenTag of parentTagHTML) -- get parent hierarchy data
				end if
				set spanTag to my fmtTag(aTagName) -- format the Tag Name
				set allTagNames to allTagNames & spanTag -- add to cell for tag table detail row
				-- Construct an HTML table row for the tag table, which will appear in
				-- the bottom half of the notebook index page and in the top-level
				-- all-tags page.
				set tagRows to tagRows & "    <tr>
      <td>" & spanTag & "</td>
" & notelink & noteLinkFar & "      <td>" & parentCell & "</td>
" & noteBookCell & "    </tr>
"
			end repeat
		end if
		--
		-- Figure out file sizes and how many attachments this note has
		--
		set sz to size of (get info for noteHTMLfile) -- size of note html file 0.3.5
		set aNoteAttach to {} -- assume no attachments for this note
		try -- handle error if no resources folder exists
			tell application "System Events" -- 0.3.5
				set resourceFolder to (POSIX path of noteHTMLfile) & ".resources"
				set sz to sz + (size of folder resourceFolder) -- add size of resources folder
				set aNoteAttach to every file of folder resourceFolder -- get list of attachments
			end tell
		end try
		set aNoteAttach to count of aNoteAttach -- count attachments
		set bookFiles to bookFiles + aNoteAttach -- count attachments in this notebook
		set bookNotes to bookNotes + 1 -- count notes in this notebook
		set htmlNotes to htmlNotes & "    <tr>
" & notelink & my dateCell(aNoteCreate) & my dateCell(aNoteMod) & fmtSize(sz) & "      <td class='clfile'>" & aNoteAttach & "</td>
" & aNoteAuthor & "      <td>" & allTagNames & "</td></tr> 
"
	end repeat -- End of processing for notes in current notebook
	--
	-- All notes in this notebook have been processed. Get ready to write the index page
	-- for the notebook.
	--
	set yy to alias (outHTML as POSIX file)
	tell application "System Events" to set dirSize to size of yy -- get size of notebook 0.3.5
	--	set dirSize to my getSize({outHTML}) -- size of this html notebook folder in text and as a number 0.3.5
	-- Build the table row for this notebook. It will be imbedded as an HTML comment in
	-- the note index file and, later, extracted to build the top-level index page
	set yy to count of listNoteFiles -- count of notes in this notebook
	set noteBookRow to "    <tr>
" & noteBookCell & "      <td class='clnote' data-notes='" & yy & "'>" & yy & "</td>
" & dateCell(bookCreate) & dateCell(bookMod) & dateCell(expStamp) & fmtSize(dirSize) & "      <td class='clfile' data-files='" & bookFiles & "'>" & bookFiles & "</td>
    </tr>
"
	set sortTags to "" -- assume no tag table to sort (fragment for the body onload attribute)
	-- if there are tags in this notebook, construct the table to show them in
	-- the notebook index page
	if (length of tagRows) > 0 then
		set tagRows to "<div class='hdr2'>Tags in this Notebook</div>
<table id='tagTable'>
  <thead>
    <tr>
      <th class='linkhover' onclick=\"sortTable('tagTable',0,0)\">Tag Name</th>
      <th class='linkhover slink' onclick=\"sortTable('tagTable',1,0)\">Note</th>
      <th class='linkhover nbname' onclick=\"sortTable('tagTable',2,0)\">Note</th>
      <th class='linkhover' onclick=\"sortTable('tagTable',3,0)\">Parents</th>
      <th class='linkhover nbname' onclick=\"sortTable('tagTable',4,0)\">Notebook</th>
    </tr>
  </thead>
  <tbody>
<!-- taglist start -->
" & tagRows & "<!-- taglist end -->
  </tbody>
</table>
"
		set sortTags to "hideNBnames('nbname'); sortTable('tagTable',0,0);" -- for onload attribute
	end if
	set {year:y1, month:m1, day:d1, hours:h1, minutes:r1, seconds:s1} to ((expStamp - offsetGMT) as date)
	set ENdate to (y1 as text) & lzer(m1 as number) & lzer(d1 as text) & "T" & lzer(h1 as text) & lzer(r1 as text) & lzer(s1 as text) & "Z" -- build export date to embed as comment in notebook index page
	if priorExportDate ≠ "" then set ENdate to priorExportDate -- use prior export date if book wasn't just exported
	--
	-- Build the notebook index page for this notebook
	--
	set htmlNotes to "<!DOCTYPE html>
<head>
<!-- bookexported " & ENdate & " bookexported -->
" & getSysInfo(priorEnviron) & "<!-- bookindex
" & noteBookRow & "bookindex -->
  <meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/>
  <title>" & bookName & "</title>
  <script src='../" & myJSfile & "'></script>
  <link rel='stylesheet' type='text/css' href='../" & myCSSfile & "'>
</head>
<body onload=\"setLocalDates();sortTable('noteTable',0,0);" & sortTags & "\">
" & pageHead("Notebook: " & my textHT(bookName), "<a href='../" & myTopIndex & "'>All Notebooks</a>&nbsp;&nbsp;<a href='../" & myTagFile & "'>All Tags</a>", "Click headings to sort. This notebook was exported from Evernote " & my fmtLongDate(expStamp) & " " & tzone() & "</br>Totals:  " & bookNotes & " notes, " & makeSize(dirSize) & " size, " & bookFiles & " files (attachments)", true) & "
<table id='noteTable'>
  <thead>
    <tr>
      <th class='linkhover notecol' onclick=\"sortTable('noteTable',0,0)\">Note Title</th>
      <th class='linkhover' onclick=\"sortTable('noteTable',1,3)\">Created</th>
      <th class='linkhover' onclick=\"sortTable('noteTable',2,3)\">Modified</th>
      <th class='linkhover' onclick=\"sortTable('noteTable',3,2)\">Size</th>
      <th class='linkhover' onclick=\"sortTable('noteTable',4,1)\">Files</th>
      <th class='linkhover' onclick=\"sortTable('noteTable',5,0)\">Author</th>
      <th class='linkhover' onclick=\"sortTable('noteTable',6,0)\">Tags</th>
	</tr>
  </thead>
  <tbody>
" & htmlNotes & "  </tbody>
</table>
" & tagRows & "
</br></br>
</body>
</html>
"
	my writeIndex(outHTML & nbIndexFile, htmlNotes) -- write the index.html for this notebook
	return
end buildNoteIndex
--
-- Build top level index.html page by reading data stored as comments in
-- each of the notebook index pages
--
on buildTopIndex()
	local bookHTML, tagHTML, buildDate, folderList, fol, pf
	local bookdata, pcounter, htmlBooks, tagOutput, exi
	local totNotes, totSize, totFiles, totBooks
	-- get names of all subfolders
	-- tell application "Finder" to set folderList to every item of alias FolderHTML whose kind is "Folder"
	tell application "System Events" to set folderList to every folder of alias FolderHTML --  0.3.5
	if logProgress then log "Folders being indexed"
	set bookHTML to "" -- init table body html for list of notebooks
	set tagHTML to "" -- init table body html for list of tags
	set buildDate to current date -- date & time these index pages are built
	set totBooks to count of folderList
	set totNotes to 0 -- init totals
	set totSize to 0
	set totFiles to 0
	set progress total steps to totBooks -- set up progress display
	set progress completed steps to 0 -- set progress counter to zero
	set progress description to "Building top-level index"
	set pcounter to 0 -- counter for progress display
	--
	--  This loop reads every notebook index page and extracts
	--  data elements that were stored as comments in the html
	--
	repeat with fol in folderList -- for every subfolder
		tell application "System Events" to set pf to name of fol -- 0.3.5
		set progress additional description to "Processing " & pf
		-- Read the index.html file from the subfolder and look for two strings:
		--    the table entry for the notebook itself in the top-level index page
		--    the list of tags for this notebook, if any
		-- we look for both in this one operation so that we don't have to
		-- read the file two times.
		-- tell application "System Events" to set pf to POSIX path of fol
		tell application "System Events" to set pf to ((path of fol) & nbIndexFile)
		set exi to true
		try
			set pf to alias pf
		on error
			set exi to false
		end try
		if exi then -- if index.html exists in this folder
			set bookdata to readIndex(pf, {"<!-- bookindex
", "<!-- taglist start -->
", "data-notes='", "data-foldersize='", "data-files='"}, {"bookindex -->", "<!-- taglist end -->", "'", "'", "'"}, 0) -- read the index.html page in the subfolder
			set bookHTML to bookHTML & "    " & item 1 of bookdata -- keep the link data
			set tagHTML to tagHTML & item 2 of bookdata -- and the tag data
			try
				set totNotes to totNotes + (0 + (item 3 of bookdata))
				set totSize to totSize + (0 + (item 4 of bookdata))
				set totFiles to totFiles + (0 + (item 5 of bookdata))
			end try
			set pcounter to pcounter + 1
			set progress completed steps to pcounter
		end if
	end repeat -- all notebook index pages have now been read
	-- Build the HTML for the top-level index.html page
	set htmlBooks to "<!DOCTYPE html>
<head>
" & getSysInfo("") & "  <meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/>
  <title>All Notebooks</title>
  <script src='" & myJSfile & "'></script>
  <link rel='stylesheet' type='text/css' href='" & myCSSfile & "'>
</head>
<body onload=\"setLocalDates();sortTable('bookTable',0,0);\">
" & pageHead("All Exported Notebooks", "<a href='" & myTagFile & "'>All Tags</a>", "Click headings to sort. This page was built " & fmtLongDate(buildDate) & " " & tzone() & ".</br>Totals:  " & totBooks & " notebooks, " & totNotes & " notes, " & makeSize(totSize) & " size, " & totFiles & " files (attachments)", true) & "
<table id=\"bookTable\">
  <thead>
    <tr>
      <th class='linkhover' onclick=\"sortTable('bookTable',0,0)\">Notebook Name</th>
      <th class='linkhover' onclick=\"sortTable('bookTable',1,1)\">Notes</th>
      <th class='linkhover' onclick=\"sortTable('bookTable',2,3)\">Created</th>
      <th class='linkhover' onclick=\"sortTable('bookTable',3,3)\">Modified</th>
      <th class='linkhover' onclick=\"sortTable('bookTable',4,3)\">Exported</th>
      <th class='linkhover' onclick=\"sortTable('bookTable',5,2)\">Size</th>
      <th class='linkhover' onclick=\"sortTable('bookTable',6,1)\">Files</th>
    </tr>
  </thead>
  <tbody>
" & bookHTML & "  </tbody>
</table>
</br></br>
</body>
</html>
"
	my writeIndex((POSIX path of FolderHTML) & myTopIndex, htmlBooks) -- write top-level index page for all notebooks
	-- build placeholder HTML for the top-level list of tags in case there are none
	set tagOutput to "<!DOCTYPE html>
<head>
  <meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/>
  <title>All Tags</title>
  <script src='" & myJSfile & "'></script>
  <link rel='stylesheet' type='text/css' href='" & myCSSfile & "'>
</head>
<body>
" & pageHead("All Tags", "<a href='" & myTopIndex & "'>All Notebooks</a>", "No tags were found in any of the exported notebooks.", false) & "
<div>&nbsp;</div>
</br></br>
</body>
</html>
"
	-- if there are tags, build HTML for the All Tags page
	if length of tagHTML > 0 then set tagOutput to "<!DOCTYPE html>
<head>
  <meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/>
  <title>All Tags</title>
  <script src='" & myJSfile & "'></script>
  <link rel='stylesheet' type='text/css' href='" & myCSSfile & "'>
</head>
<body onload=\"hideNBnames('slink');sortTable('tagTable',0,0);\">
" & pageHead("All Tags", "<a href='" & myTopIndex & "'>All Notebooks</a>", "Click headings to sort. This page was built " & fmtLongDate(buildDate) & " " & tzone() & ".", true) & "
<table id='tagTable'>
  <thead>
    <tr>
      <th class='linkhover' onclick=\"sortTable('tagTable',0,0)\">Tag Name</th>
      <th class='linkhover slink' onclick=\"sortTable('tagTable',1,0)\">Note</th>
      <th class='linkhover nbname' onclick=\"sortTable('tagTable',2,0)\">Note</th>
      <th class='linkhover' onclick=\"sortTable('tagTable',3,0)\">Parents</th>
      <th class='linkhover nbname' onclick=\"sortTable('tagTable',4,0)\">Notebook</th>
    </tr>
  </thead>
  <tbody>
" & tagHTML & "
  </tbody>
</table>
</br></br>
</body>
</html>
"
	my writeIndex((POSIX path of FolderHTML) & myTagFile, tagOutput) -- write all-tag page
	my writeJS() -- write javascript file if it's not already there		
	my writeCSS() -- write CSS file if it's not already there		
end buildTopIndex
--
-- Extract meta tag info from HTML note file
-- reads a hfs file name, expecting it to be an HTML note file
-- returns note metadata as a Record
--
on getMeta(ifile)
	local resp, XMLAppData, t, nm, cn, metas
	set ourTimer to getFineTime() -- init sub timer
	-- Read first 2000 bytes of the note html file, hoping to get the entire
	-- <head>...</head> group. Append those tags to the returned string so
	-- the system event XML processor will see them.
	set hd to "<head>" & (item 1 of readIndex(ifile, {"<head>"}, {"</head>"}, 2000)) & "</head>"
	set resp to {nname:"", author:"", created:"", tags:{""}, updated:"", exporter:""} -- init returned value
	tell application "System Events"
		--		set XMLAppData to make new XML data with properties {name:"htData", text:hd}
		--		tell XML data "htData" -- id XMLAppData
		tell (make new XML data with properties {text:hd}) -- pass the note html to the xml scanner
			tell XML element "head" -- look within the head tag
				set metas to every XML element whose name = "meta" -- get all meta tags
				repeat with t from 1 to (length of metas) -- look at each meta tag in turn
					set thismeta to item t of metas -- get this meta tag
					if exists (XML attribute "name" of thismeta) then -- if there'a a 'name' attribute
						set nm to value of XML attribute "name" of thismeta -- get the name attribute
						set cn to value of XML attribute "content" of thismeta -- and the content attribute
						if nm is "exporter-version" then set exporter of resp to cn -- extract based on name ....
						if nm is "created" then set created of resp to cn
						if nm is "updated" then set updated of resp to cn
						if nm is "author" then set author of resp to cn
						if nm is "keywords" then set tags of resp to cn
					end if
				end repeat
				set nname of resp to (value of XML element "title") as text -- get note title from title tag
			end tell
		end tell
	end tell
	set timerSub1 to timerSub1 + (getFineTime() - ourTimer) -- accumulate time for this subroutine
	return resp -- return record with meta info to caller
end getMeta
--
-- Outer function to read Note html file and then
-- call our XML processor to parse the <head> tag
-- Input:  alias of note html file
-- Returns: record of values from <head>
--
on getMeta2(infile)
	local theXMLdata, ourTimer, resp
	set theXMLdata to "<?xml version=\"1.0\" encoding=\"UTF-8\"?> <head>" & (item 1 of readIndex(infile, {"<head>"}, {"</head>"}, 2000)) & "</head>" -- get html head tags from note html file
	set ourTimer to getFineTime() -- timer for start
	set resp to parseNote(theXMLdata) -- parse the xml data
	set timerSub1 to timerSub1 + (getFineTime() - ourTimer) -- accumulate time for this subroutine
	return resp
end getMeta2
--
--  Called by getMeta2
--  This encapsulates the fromework "Foundation" call in
--  a separate script.
--  For XML processing, see https://macscripter.net/viewtopic.php?id=45479
--  and https://macscripter.net/viewtopic.php?pid=200788
--  and https://developer.apple.com/documentation/foundation/nsxmldocument
--  and https://www.w3schools.com/xml/xpath_intro.asp
--
on parseNote(xmlData)
	script xmlScript -- encase the asobjc call in a script object
		property parent : a reference to current application
		use framework "Foundation"
		on parseNote(xmlData)
			local theXMLDoc, theError, t, resp
			set resp to {nname:"", author:"", created:"", tags:"", updated:"", exporter:""} -- init returned value
			set {theXMLDoc, theError} to current application's NSXMLDocument's alloc()'s initWithXMLString:xmlData options:0 |error|:(reference) -- parse the xml data
			if theXMLDoc ≠ missing value then -- if valid xml was found
				-- if theXMLDoc = missing value then error (theError's localizedDescription() as text)
				set t to ((theXMLDoc's nodesForXPath:"/head/title" |error|:(missing value))'s valueForKey:"stringValue") as list
				if (length of t) > 0 then set nname of resp to (item 1 of t) -- copy Note Title if it's there
				set t to ((theXMLDoc's nodesForXPath:"//meta[@name='created']/@content" |error|:(missing value))'s valueForKey:"stringValue") as list
				if (length of t) > 0 then set created of resp to (item 1 of t) -- copy Creation Date if it's there
				set t to ((theXMLDoc's nodesForXPath:"//meta[@name='author']/@content" |error|:(missing value))'s valueForKey:"stringValue") as list
				if (length of t) > 0 then set author of resp to (item 1 of t) -- copy Author Name if it's there
				set t to ((theXMLDoc's nodesForXPath:"//meta[@name='updated']/@content" |error|:(missing value))'s valueForKey:"stringValue") as list
				if (length of t) > 0 then set updated of resp to (item 1 of t) -- copy Modification Date if it's there
				set t to ((theXMLDoc's nodesForXPath:"//meta[@name='keywords']/@content" |error|:(missing value))'s valueForKey:"stringValue") as list
				if (length of t) > 0 then set tags of resp to (item 1 of t) -- copy Tag Names if it's there
				set t to ((theXMLDoc's nodesForXPath:"//meta[@name='exporter-version']/@content" |error|:(missing value))'s valueForKey:"stringValue") as list
				if (length of t) > 0 then set exporter of resp to (item 1 of t) -- copy EN Version if it's there
			end if
			return resp
		end parseNote
	end script
	return xmlScript's parseNote(xmlData)
end parseNote
--
-- Make text string HTML safe (for Note titles, Notebook Names, Tag names, etc.)
--
on textHT(inString)
	local j
	set j to repl(inString, "&", "&amp;") -- must be first to avoid double-encoding of ampersands
	set j to repl(j, "\"", "&quot;")
	set j to repl(j, "'", "&#039;")
	set j to repl(j, "<", "&lt;")
	set j to repl(j, ">", "&gt;")
	return j
end textHT
--
-- Replace one string with another
--
on repl(theText, theSearchString, theReplacementString)
	local theTextItems
	set AppleScript's text item delimiters to theSearchString
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to theReplacementString
	set theText to theTextItems as string
	set AppleScript's text item delimiters to ""
	return theText
end repl
--
-- Return formatted date and time as yyyy-mm-dd hh-mm
-- input is a date object
--
on fmtLongDate(inpDate)
	set {year:y, month:m, day:d, hours:h, minutes:r} to inpDate
	return (y as text) & "-" & lzer(m as number) & "-" & lzer(d as text) & " " & lzer(h as text) & ":" & lzer(r as text)
end fmtLongDate
--
-- Return an integer as two digits, with leading zero if necessary and negative sign if < 0
--
on lzer(numb)
	if numb ≥ 0 then
		return (text -1 thru -2 of ("000" & (numb as text)))
	else
		return "-" & (text -1 thru -2 of ("000" & ((-numb) as text)))
	end if
end lzer
--
-- Return time zone offset as -hh:mm
--
on tzone()
	return lzer(offsetGMT div hours) & ":" & lzer(offsetGMT mod hours / minutes as integer)
end tzone
-- 
-- Return a date object
-- input is a date/time string from an HTML meta tag
--    as 2020-11-22 21:05:34 +0000
--
on getDateMeta(x)
	local d
	set d to initDate(2000, 1, 1) -- 0.4.2
	if x = "" then return d
	tell d to set {its year, its month, its day, its time} to {0 + (text 1 thru 4 of x), 0 + (text 6 thru 7 of x), 0 + (text 9 thru 10 of x), (hours * (0 + (text 12 thru 13 of x))) + (minutes * (0 + (text 15 thru 16 of x))) + (0 + (text 18 thru 19 of x))} -- 
	set d to d + (hours * (text 22 thru 23 of x)) + (0 + (text 24 thru 25 of x)) * (cond((text 21 thru 21 of x) = "-", -1, 1)) + (time to GMT)
	return d
end getDateMeta
-- 0.4.2 ....
-- Initialize a date object
-- input is year, month, day as integers
-- returns date object
-- Avoids locale-related issues with dates in strings
-- 
on initDate(yr, mo, dy) -- 0.4.2
	local md
	set md to current date
	tell md to set {its year, its month, its day, its time} to {yr, mo, dy, hours * 12}
	return md
end initDate
-- .... 0.4.2
-- 
-- Return a date object in local time
-- input is a date/time string as in our bookexported comment
-- and in Evernote's search syntax:  20210205T203848Z
-- With help from https://forum.latenightsw.com/t/time-to-gmt-for-any-date/2070/2
--    to calc GMT offset for arbitrary day of the year
-- and:  http://www.unicode.org/reports/tr35/tr35-31/tr35-dates.html#Date_Format_Patterns
--     for info on date pattern strings
--
on getDateSrch(x)
	script dateScript
		property parent : a reference to current application
		use framework "Foundation"
		on getDateSrch(x)
			local df, theDate, tz, theOffset, d
			if x = "" then return initDate(2000, 1, 1) -- 0.4.2
			set df to current application's NSDateFormatter's new()
			df's setDateFormat:"yyyyMMdd'T'HHmmssX" -- Pattern date
			set theDate to df's dateFromString:x
			-- get offset
			set tz to current application's NSTimeZone's localTimeZone()
			set theOffset to tz's secondsFromGMTForDate:theDate
			-- I should be able to add theDate to theOffset, but Applescript could not seem to
			-- constrain theDate from NSDate to Date.  So, I'm converting it another way....
			set d to current date -- create a date variable  0.4.2
			tell d to set {its year, its month, its day, its time} to {0 + (text 1 thru 4 of x), 0 + (text 5 thru 6 of x), 0 + (text 7 thru 8 of x), (hours * (0 + (text 10 thru 11 of x))) + (minutes * (0 + (text 12 thru 13 of x))) + (0 + (text 14 thru 15 of x))} -- convert date string to date object
			set d to d + theOffset -- adjust GMT time string to local time
			return d -- return date in our time zone
		end getDateSrch
	end script
	return dateScript's getDateSrch(x)
end getDateSrch
--
-- Write an index.html file
--
on writeIndex(thePath, theText)
	set theAlias to thePath as POSIX file as «class furl»
	try
		-- tell application "Finder"  (adding class furl below fixed this)
		-- using terms from scripting additions.  (adding class furl below fixed this)
		set theOpenedFile to open for access theAlias as «class furl» with write permission
		-- end tell
		-- end using terms from
		-- if logProgress then log ("response from open: " & theOpenedFile)
		set eof of theOpenedFile to 0
		write theText to theOpenedFile as «class utf8»
		-- if logProgress then log "Normal file close"
		close access theOpenedFile
		return true
	on error errMsg number errNum
		-- if logProgress then log "Error file close. Msg: " & errMsg & "  Num: " & errNum
		try
			close access file theAlias
		end try
		return false
	end try
	-- if logProgress then log "Leaving addtolog"
end writeIndex
--
--  Format file/folder size for display within a table cell
--  Input is a number of bytes
--  Returns a table cell (<TD>) formatted with file size
--
on fmtSize(s)
	return "      <td class='clsize' data-foldersize='" & s & "'>" & makeSize(s) & "</td>
"
end fmtSize
--
--  Format file/folder size for display
--  Input is a number of bytes
--  Returns a string with file size
--  We're using decimal calc for KB, MB and GB, in keeping
--  with how macOS Finder reports sizes
-- This newer version of the handler adds one decimal point to the result
--
on makeSize(s)
	local strSize
	if s < 1000 then -- Build a formatted text showing total size
		set strSize to "0." & ((s / 100) as integer as text) & " KB"
	else if (s / 1000) < 1000 then
		set strSize to ((s / 100) as integer as text)
		set strSize to (text 1 thru -2 of strSize) & "." & (text -1 thru -1 of strSize) & " KB"
	else if (s / (1000 ^ 2)) < 1000 then
		set strSize to ((s / (100000)) as integer as text)
		set strSize to (text 1 thru -2 of strSize) & "." & (text -1 thru -1 of strSize) & " MB"
	else
		set strSize to ((s / (1000 ^ 3 / 10)) as integer as text)
		set strSize to (text 1 thru -2 of strSize) & "." & (text -1 thru -1 of strSize) & " GB"
	end if
	return strSize
end makeSize
--
--  Format file/folder size for display
--  Input is a number of bytes
--  Returns a string with file size
--  We're using decimal calc for KB, MB and GB, in keeping
--  with how macOS Finder reports sizes
--
on makeSizeOld(s)
	local strSize
	if s < 1000 then -- Build a formatted text showing total size
		set strSize to (s as integer as text) & " B&nbsp;&nbsp;"
	else if (s / 1000) < 1000 then
		set strSize to ((s / 1000) as integer as text) & " KB"
	else if (s / (1000 ^ 2)) < 1000 then
		set strSize to ((s / (1000 ^ 2)) as integer as text) & " MB"
	else
		set strSize to ((s / (1000 ^ 3)) as integer as text) & " GB"
	end if
	return strSize
end makeSizeOld

-- Handler for Property List file.  Three functions are performed in this one routine
--    so that the relevant variables can be localized here.
-- Option 1: Create Plist if necessary then fetch current property values
--    into global variables.
--    if PList already exists, return True; if we had to create it, return false
-- Option 2: Prompt user to update properties one after the other.
--    returns FALSE if user cancels or fails to correctly set the properties
-- Option 3: Update the PList with a new ExportDate
--   returns True
on managePlist(opt)
	local plFolderHTML, plFolderENEX, plAllNotebooks, plSearchString
	local plExportAsHTML, plExportAsENEX, plBooksSelected, plPropDate
	local plExportDate, plTimeoutMinutes, foundPList, theParentDictionary
	local thePropertyListFile, myResp, exopt, buts, chc, bookNames
	local msgPart, resp
	-- Names of Property List items
	set plFolderHTML to "FolderHTML" -- Folder for HTML exports
	set plFolderENEX to "FolderENEX" -- Folder for ENEX exports
	set plAllNotebooks to "AllNotebooks" -- Export all notebooks, not just some
	set plSearchString to "SearchString" -- user's filter for notes
	set plExportAsHTML to "ExportAsHTML" -- Export as HTML?
	set plExportAsENEX to "ExportAsENEX" -- Export as ENEX?
	set plBooksSelected to "BooksSelected" -- names of notebooks to export
	set plPropDate to "PropDate" -- date this Property List last updated
	set plExportDate to "ExportDate" -- date export last completed
	set plTimeoutMinutes to "TimeoutMinutes" -- export timeout in minutes
	
	if opt = 1 then -- Option 1: create default PList if necessary and then load Plist values into Global variables
		-- set foundPList to true if PList exisits
		tell application "Finder" to set foundPList to (exists myPListFullPath as POSIX file) -- Does PList exist?
		if not foundPList then -- If it doesn't exist, create one with default values
			tell application "System Events"
				set theParentDictionary to make new property list item with properties {kind:record}
				set thePropertyListFile to make new property list file with properties {contents:theParentDictionary, name:myPListFullPath}
				tell property list items of thePropertyListFile
					make new property list item at end with properties {kind:string, name:plFolderHTML, value:FolderHTML}
					make new property list item at end with properties {kind:string, name:plFolderENEX, value:FolderENEX}
					make new property list item at end with properties {kind:boolean, name:plAllNotebooks, value:AllNotebooks}
					make new property list item at end with properties {kind:string, name:plSearchString, value:SearchString}
					make new property list item at end with properties {kind:boolean, name:plExportAsHTML, value:ExportAsHTML}
					make new property list item at end with properties {kind:boolean, name:plExportAsENEX, value:ExportAsENEX}
					make new property list item at end with properties {kind:list, name:plBooksSelected, value:BooksSelected}
					make new property list item at end with properties {kind:date, name:plPropDate, value:(current date)}
					make new property list item at end with properties {kind:date, name:plExportDate, value:ExportDate}
					make new property list item at end with properties {kind:number, name:plTimeoutMinutes, value:TimeoutMinutes}
				end tell
			end tell
		end if
		--  Fetch current PList values from the file
		tell application "System Events"
			tell property list file myPListFullPath
				set FolderHTML to value of property list item plFolderHTML
				set FolderENEX to value of property list item plFolderENEX
				set AllNotebooks to value of property list item plAllNotebooks
				set SearchString to value of property list item plSearchString
				set ExportAsHTML to value of property list item plExportAsHTML
				set ExportAsENEX to value of property list item plExportAsENEX
				set BooksSelected to value of property list item plBooksSelected as list
				set ExportDate to value of property list item plExportDate
				set TimeoutMinutes to value of property list item plTimeoutMinutes
			end tell
		end tell
		return foundPList -- Tell caller if PList alreay existed
	end if
	--
	-- This option prompts the user to review / update the properties
	-- 
	if opt = 2 then -- Option 2: Prompt for updates to Property List items
		set myResp to false -- assume update has failed
		tell application "System Events" to set frontProcess to the name of every process whose frontmost is true
		set exopt to 3 -- assume both formats
		if (not (ExportAsHTML and ExportAsENEX)) then set exopt to cond(not ExportAsHTML, 2, 1) -- if not HTML, then ENEX only
		set exopt to item_position(choose from list plExportTypeDesc with title "Choose export types" with prompt "Which format(s) should be exported?" default items {item exopt of plExportTypeDesc}, plExportTypeDesc)
		if logProgress then log "Export option: " & exopt
		if exopt = 0 then return myResp -- quit if user cancelled
		set ExportAsHTML to item exopt of {true, false, true} -- Set HTML export flag
		set ExportAsENEX to item exopt of {false, true, true} -- Set ENEX export flag
		if logProgress then log "ExportAsHTML: " & ExportAsHTML & "  ExportAsENEX: " & ExportAsENEX
		if ExportAsHTML then -- If exporting as HTML
			-- first, check that volume containing default output folder is mounted
			-- use Documents if user says to ignore prior setting
			if checkPropVol(FolderHTML, "HTML") then set FolderHTML to path to documents folder as text
			tell application "Finder" to set FolderHTML to (choose folder with prompt "Select or create folder for HTML exports" default location alias FolderHTML) as text -- prompt for folder
		end if
		if logProgress then log "FolderHTML: " & FolderHTML
		if ExportAsENEX then -- if exporting as ENEX
			-- first, check that volume containing default output folder is mounted
			-- use Documents if user says to ignore prior setting
			if checkPropVol(FolderENEX, "ENEX") then set FolderENEX to path to documents folder as text
			tell application "Finder" to set FolderENEX to (choose folder with prompt "Select or create folder for ENEX exports" default location alias FolderENEX) as text -- prompt for folder
		end if
		if logProgress then log "FolderENEX: " & FolderENEX
		activate frontProcess -- switch back from Finder to script
		set buts to {"All notebooks", "Only selected notebooks", "Cancel"}
		set chc to display dialog "Export all notebooks or only selected ones?" buttons buts default button cond(AllNotebooks, 1, 2) cancel button 3 with icon note
		set AllNotebooks to ((item 1 of buts) = (button returned of chc))
		if logProgress then log "AllNotebooks: " & AllNotebooks
		if not AllNotebooks then -- if only selected notebooks
			--			set bookNames to sortList(getBooks("")) -- sort notebook names   0.3.2
			set bookNames to getBooks() -- get sorted notebook names
			if logProgress then log "Count of Notebooks in Evernote: " & (count of bookNames)
			set BooksSelected to choose from list bookNames with title "Choose Notebooks to export" with prompt "Use command key (⌘) to select several or all" default items BooksSelected with multiple selections allowed
			if BooksSelected = false then return myResp -- if user cancelled, return to caller w 'false'
			if logProgress then log "BooksSelected: " & BooksSelected
		end if
		set resp to display dialog "Enter Evernote search terms to restrict notes that will be exported. Do not use 'notebook:', 'any:', or 'updated:' terms. Leave blank to export all notes in each notebook." default answer SearchString with title "Optional search terms" with icon note
		set SearchString to (text returned of resp) as text -- save user input
		set msgPart to "" -- clear error message fragment
		repeat -- loop to prompt for Timeout value and validate it
			set resp to display dialog ("Enter export timeout in minutes. " & msgPart) default answer (TimeoutMinutes as text) with title "Specify Timeout" with icon note
			set msgPart to "" -- clear error message fragment
			try -- see if text entered is a valid integer
				if logProgress then log "About to convert to integer."
				set TimeoutMinutes to (text returned of resp) as integer
				if logProgress then log "Convert to integer worked: " & TimeoutMinutes
			on error -- not a valid integer
				set msgPart to "
Timeout value must be a positive integer." -- set error message
				if logProgress then log "Convert to integer failed."
				set TimeoutMinutes to (text returned of resp) -- leave bad timeout value in place
			end try
			if logProgress then log "TimeoutMinutes: " & TimeoutMinutes
			if msgPart = "" then
				if TimeoutMinutes < 1 then
					set msgPart to "
Timeout value must be a positive integer." -- set error message
				else
					exit repeat -- if integer was valid, break out of the Repeat loop
				end if
			end if
		end repeat
		--
		-- With all Properties now reviewed by user, write them
		-- to the Property List file
		--
		tell application "System Events"
			tell property list file myPListFullPath
				set value of property list item plFolderHTML to FolderHTML
				set value of property list item plFolderENEX to FolderENEX
				set value of property list item plAllNotebooks to AllNotebooks
				set value of property list item plSearchString to SearchString
				set value of property list item plExportAsHTML to ExportAsHTML
				set value of property list item plExportAsENEX to ExportAsENEX
				set value of property list item plTimeoutMinutes to TimeoutMinutes
				set value of property list item plBooksSelected to BooksSelected
				set value of property list item plPropDate to current date
			end tell
		end tell
		if logProgress then log "Successfully set all properties"
		return true
	end if
	--
	-- This option saves the date of the last successful export
	-- 
	if opt = 3 then -- Option 3: Update only the ExportDate property
		tell application "System Events"
			tell property list file myPListFullPath
				set value of property list item plExportDate to ExportDate
			end tell
		end tell
		return true
	end if
end managePlist
--
-- Get list of all notebooks from Evernote
-- Return a sorted list
--
on getBooks()
	local listBooks
	set listBooks to {} -- init the list of book names	
	tell application "Evernote"
		repeat with iBook in every notebook -- for every notebook ...
			copy (name of iBook) to end of listBooks -- add notebook name to list
		end repeat
	end tell
	return sortListOfStrings(listBooks) -- sort the list and return it
end getBooks
--
-- Sort a list of strings
-- based on https://macosxautomation.com/applescript/sbrt/sbrt-00.html
-- converted to a Script object to avoid using framework "Foundation" in the main script
-- 
on sortListOfStrings(sourceList)
	script sortScript -- encase the asobjc call in a script object
		property parent : a reference to current application
		use framework "Foundation"
		on sortListOfStrings(sourceList)
			-- create a Cocoa array from the passed AppleScript list
			set the CocoaArray to current application's NSArray's arrayWithArray:sourceList
			-- sort the Cocoa array
			set the sortedItems to CocoaArray's sortedArrayUsingSelector:"localizedStandardCompare:"
			-- return the Cocoa array coerced to an AppleScript list
			return (sortedItems as list)
		end sortListOfStrings
	end script
	return sortScript's sortListOfStrings(sourceList)
end sortListOfStrings
--
-- Return time since Jan 1, 2001 in millionths of a second
-- 
on getFineTime()
	script timeScript -- encase the asobjc call in a script object
		property parent : a reference to current application
		use framework "Foundation"
		on getFineTime()
			return current application's CFAbsoluteTimeGetCurrent()
		end getFineTime
	end script
	return timeScript's getFineTime()
end getFineTime
--
-- Return the item number of an item in a list
--
on item_position(myitem, theList)
	if myitem = false then return 0
	repeat with i from 1 to the count of theList
		if (item i of theList) is (item 1 of myitem) then return i
	end repeat
	return 0
end item_position
--
-- Format a note Tag by enclosing it in html span tags
--
on fmtTag(tn)
	-- wrap tagname in a Span, replacing blanks with non-breaking space to
	-- hopefully prevent them from breaking across lines in notebook index page
	return "<span class='ENtag'>" & repl(textHT(tn), " ", "&nbsp;") & "</span> "
end fmtTag
--
-- Format a table cell containg a date
-- The local date will be visible in the cell and
-- the GMT date & time will be in a data attribute for use by
-- javascript, which will update the cell value and create a title (tool tip)
--
on dateCell(dt)
	local y, m, d, yy, mm, dd, hh, s, r
	set {year:y, month:m, day:d} to dt -- for short, local date that will visible in cell
	-- return (y as text) & "-" & lzer(m as number) & "-" & lzer(d as text)
	set {year:yy, month:mm, day:dd, hours:hh, minutes:r, seconds:s} to (dt - offsetGMT) -- for javascript to convert
	--	return (yy as text) & "-" & lzer(mm as number) & "-" & lzer(dd as text) & "T" & lzer(hh as text) & ":" & lzer(r as text) & ":" & lzer(s as text) & "Z"
	return "      <td data-gmtdate='" & ((yy as text) & "-" & lzer(mm as number) & "-" & lzer(dd as text) & "T" & lzer(hh as text) & ":" & lzer(r as text) & ":" & lzer(s as text) & "Z") & "' class='tdate'>" & ((y as text) & "-" & lzer(m as number) & "-" & lzer(d as text)) & "</td>
"
end dateCell
--
-- Format a page header with left and right content
--
on pageHead(cleft, cright, cinstruct, checkbox)
	return "  <div class='hdr'>
    <div class='lhdr'>" & cleft & "</div>
    <div class='rhdr'>" & cright & "</div>
    <div class='bhdr'>&nbsp;</div>
    <div class='instruct'>" & cinstruct & "</div>
" & cond(checkbox, "      <div>
        <form class='inpform'>
          <label for='notetarg'>Open links in new tabs</label>
          <input type='checkbox' id='notetarg' onchange=\"settarget('notetarg','linktarg');\"></form></div>", "") & "
    <div class='bhdr'>&nbsp;</div>
  </div>
"
end pageHead
--
-- Return a value conditionally
--
on cond(tst, v1, v2)
	if tst then return v1
	return v2
end cond
--
-- Get a text strings from an existing file
-- fname = full name of input file
-- frontstr = list of leading tokens to scan for
-- backstr = list of trailing tokens to scan for
-- chrs = bytes of file to read, or 0 for entire file
-- Returns list of strings found with "" for strings not found
-- 
on readIndex(fname, frontstr, backstr, chrs)
	local ex, xd, chrs, res, i, stpos, bal, enpos, infile
	try
		if logProgress then log "Reading file: " & fname & "  for " & chrs & " bytes"
	end try
	tell application "Finder" to set ex to (exists fname) -- check for file existence
	if not ex then -- if file does not exist
		set xd to "dummy string" -- set file data to a dummy string
	else -- otherwise, read the input file
		try
			set fname to fname as alias
		end try
		if chrs = 0 then
			--set xd to read (fname as «class furl») as «class utf8» -- read entire file
			set xd to read fname as «class utf8» -- read entire file
		else
			--set xd to read (fname as «class furl») for chrs as «class utf8» -- if caller specified a byte count
			repeat with ei from chrs to (chrs + 7) -- try several reads in case UTF8 char is split
				set readOK to true
				try
					set xd to read fname for ei as «class utf8» -- if caller specified a byte count
				on error
					set readOK to false
				end try
				if readOK then exit repeat
				set xd to "dummy string" -- set file data to a dummy string
			end repeat
		end if
	end if
	if logProgress then log "   Len read: " & length of xd
	set res to {} -- initialize list to be returned
	repeat with i from 1 to (length of frontstr) -- for each leading search token
		-- if logProgress then log "   frontstr: " & i & "  " & item i of frontstr
		-- if logProgress then log "   backstr : " & i & "  " & item i of backstr
		set stpos to offset of (item i of frontstr) in xd -- find front token
		if stpos = 0 then
			copy "" to end of res -- return empty string if not found
		else
			set stpos to stpos + (length of (item i of frontstr)) -- start of target text
			set bal to text stpos thru (length of xd) of xd -- remainder of file data
			set enpos to (offset of (item i of backstr) in bal) - 1
			if enpos = 0 then
				copy "" to end of res -- return empty string if not found
			else
				copy trim(text 1 thru enpos of bal) to end of res
			end if
		end if
		-- if logProgress then log "   result(" & (item i of res) & ")"
	end repeat
	return res
end readIndex
--
-- Trim leading and trailing blanks from a string
--
-- From Nigel Garvey's CSV-to-list converter
-- http://macscripter.net/viewtopic.php?pid=125444#p125444
on trim(txt)
	repeat with i from 1 to (count txt) - 1
		if (txt begins with space) then
			set txt to text 2 thru -1 of txt
		else
			exit repeat
		end if
	end repeat
	repeat with i from 1 to (count txt) - 1
		if (txt ends with space) then
			set txt to text 1 thru -2 of txt
		else
			exit repeat
		end if
	end repeat
	if (txt is space) then set txt to ""
	return txt
end trim
--
-- Check that the volume is mounted before running an export
-- Then look for folder. 
-- If both found, return false
--
on checkRunVol(fpath, omode)
	local diskNames, targetVolume, ex
	set targetVolume to text 1 thru ((offset of ":" in fpath) - 1) of fpath
	repeat
		tell application "System Events" to set diskNames to name of every disk
		if logProgress then log diskNames
		if (diskNames contains targetVolume) then exit repeat
		set ret to display alert "Volume not mounted" message "The volume '" & targetVolume & "' needed for " & omode & " export is not mounted. Please mount it and then retry. Or click cancel." buttons {"Retry", "Cancel"} default button 1 cancel button 2
	end repeat
	tell application "Finder" to set ex to (exists fpath)
	if not ex then display alert "Folder " & (POSIX path of fpath) & " not found for " & omode & " export.
Rerun this script and choose Change Settings to set " & omode & " export folder."
	return ex
end checkRunVol
--
-- Check that the volume is mounted before prompting user to select the output folder
-- Return False if found or true if not found.
--
on checkPropVol(fpath, omode)
	set targetVolume to text 1 thru ((offset of ":" in fpath) - 1) of fpath
	repeat
		tell application "System Events" to set diskNames to name of every disk
		if logProgress then log diskNames
		if (diskNames contains targetVolume) then exit repeat
		set ret to display alert "Volume not mounted" message "The volume '" & targetVolume & "' selected for " & omode & " export is not mounted. Please mount it and then retry or choose another folder or click cancel." buttons {"Retry", "Choose another", "Cancel"} default button 1 cancel button 3
		if button returned of ret = "Choose another" then return true
	end repeat
	tell application "Finder" to set ex to (exists fpath)
	activate frontProcess --	bring us to the front
	if not ex then display alert "Folder previously selected for " & omode & " export not found (" & (POSIX path of fpath) & "). 
Choose another folder."
	return not ex
end checkPropVol
--
--  Return a comment string with system info
--  to embed as a comment in the index.html pages
--  A comment string from prior index file may be supplied
--
on getSysInfo(pastEnv)
	local sys, ENacct, ENver
	if length of pastEnv > 0 then return "<!-- Exported from Evernote " & pastEnv & "   -->
" -- reuse prior comment
	if headerInfo is "" then
		set sys to system info
		tell application "Evernote"
			set ENacct to name of current account
			set ENver to version
		end tell
		set headerInfo to "<!-- Exported from Evernote ver: " & ENver & ", account: " & ENacct & "
     running on macOS: " & (system version of sys) & "  using AppleScript ver: " & (AppleScript version of sys) & "
     on computer: " & (computer name of sys) & "  by user: " & (long user name of sys) & "
     user locale: " & (user locale of sys) & "  GMT offset: " & tzone() & "  local date: " & fmtLongDate(current date) & "   -->
"
	end if
	return headerInfo
end getSysInfo
--
-- Encode a string for use as a URL in HTML (0.3.2)
-- replaces blanks and most special characters with hex codes
-- This is encased in a Script object as described here:
-- https://latenightsw.com/adding-applescriptobjc-to-existing-scripts/
-- so that we can avoid use framework "Foundation" which messes
-- up the Read and Write commands
-- 
on urlEncode(inp)
	script aScript
		property parent : a reference to current application
		use framework "Foundation"
		on urlEncode(inp)
			tell current application's NSString to set rawUrl to stringWithString_(inp)
			set theEncodedURL to rawUrl's stringByAddingPercentEscapesUsingEncoding:4 -- 4 is NSUTF8StringEncoding
			return theEncodedURL as Unicode text
		end urlEncode
	end script
	return aScript's urlEncode(inp)
end urlEncode
--
-- WriteJS handler - if the enscripts.js file doesn't exist, write it to
--   the top level HTML folder
--
on writeJS()
	set jsFile to (POSIX path of FolderHTML) & myJSfile -- name of javascript file
	tell application "Finder" to set foundJS to (exists jsFile as POSIX file) -- Does PList exist?
	if foundJS then return false
	set js to "/*
	This JS file is created initially by the applescript
	Notebook Export Manager For Evernote and placed in
	the top-level html folder of exported Notebooks.
	-- Once created, the user can modify this file, and
	   the applescript will not over-write it.
	-- This JS file contains the functions to dynamically
	   sort the main html table and perform other
	   run-time functions inside the top-level
	   Notebook list webpage and the lower-level
	   Note list index pages.
	-- If you want to modify this JS script and re-imbed
	   it into the applescript, avoid using any backslash
	   or double-quote characters, as they must be escaped
	   in applescript text constants.
*/
function setLocalDates() {
	/*  This function is to be called by the OnLoad tag
		of a BODY element. It loops through every TD
		(Table Data) element, looking for the data-gmtdate
		attribute, taking the GMT date stored there and
		using it to populate the innerHTML and the title
		of the table cell. So, dates will be displayed
		in the user's local time zone.
		*/
	var elems,c,d,i,s
	elems = document.getElementsByTagName('TD')
	for (i = 0; i < (elems.length); i++)  {
		c = elems[i].getAttribute('data-gmtdate');
		if (c) {
			d = new Date(c);
			s = d.getFullYear() + '-' + lz(d.getMonth() + 1) + '-' + lz(d.getDate()) + ' ' + lz(d.getHours()) + ':' + lz(d.getMinutes()) ;
			elems[i].innerText = s.substring(0, 10);
			elems[i].title = s;
		}
	}
}
function lz(n) {
	/* format an integer as two digits with leading zero */
 var s = '00' + n;
 return s.substring(s.length - 2);
}
function sortTable(tname,col,typ) {
	/* 	
		Sort table using javascript array sort
		based on https://stackoverflow.com/questions/14267781/sorting-html-table-with-javascript
		tname is the ID of the table
		col is the column number to sort on (0 to n)
		typ is the type of data (see Switch statement below
	*/
	var sortattr = 'data-sort', i ;  // data attribute tag name
	var tabl = document.getElementById(tname) ;  // get table object
	var tb = tabl.tBodies[0] ; // get table tbody object
	var thdr = tabl.tHead.rows[0].cells ;  // get cells in table header row
	var sd = thdr[col].getAttribute(sortattr) == 'A' ? -1 : 1 ;  // get sort direction (A/D) (A if no prior sort)
	var tr = Array.prototype.slice.call(tb.rows, 0) ; // get rows in array
	for (i = 0; i < (thdr.length); i++) {
	  thdr[i].classList.remove('trasc','trdesc');  /* clear table header classes  */
	  thdr[i].setAttribute(sortattr, 'U');  // set all cols sort direction to unsorted
  	}
  	// Sort the array of table rows based on specified cell (ie, column) and data type
	tr = tr.sort(function(a,b) {
		var cm, la, lb
		la = a.cells[col] ; // get cell being sorted
		lb = b.cells[col] ; // get cell being sorted
		switch(typ) {
			case 1:  // sort column contains integers
				cm = (parseInt(la.innerHTML) > parseInt(lb.innerHTML)) ? 1 : -1 ;
				break ;
			case 2:  // sort column contains number in an attribute
				cm = (parseFloat(la.getAttribute('data-foldersize')) > parseFloat(lb.getAttribute('data-foldersize'))) ? 1 : -1 ;
				break ;
			case 3: // sort column contains GMT date in an attribute
				cm = (la.getAttribute('data-gmtdate') > lb.getAttribute('data-gmtdate')) ? 1 : -1 ;
				break ;
			default: // sort column contains a string, possibly inside an <a> tag
				la = ((null !== la.firstElementChild) ? la.firstElementChild : la).innerHTML.toLowerCase() ;
				lb = ((null !== lb.firstElementChild) ? lb.firstElementChild : lb).innerHTML.toLowerCase() ;
				cm = la > lb ? 1 : -1 ;
		}
		return cm * sd ;  // return comparison result, adjusted for sort direction
	}) ;
	for(i = 0; i < tr.length; ++i) tb.appendChild(tr[i]); // append each row in order back into tbody
	thdr[col].classList.add(sd == -1 ? 'trdesc' : 'trasc') ;  // set color indicator in sort column
	thdr[col].setAttribute(sortattr, (sd == -1 ? 'D' : 'A'));  // remember sorted column
}
function hideNBnames(cls){
	/*  This function is used to inject CSS into all the index.html
		pages for individual notebooks, to hide table columns
		with specified class.
		*/
	var style2 = document.createElement('style');
	style2.innerHTML = '.' + cls + ' { ' +
    '   display:none; ' +
    '   } ' ;
	document.head.appendChild(style2);
}
function settarget(checkid,acls) {
	/* This function sets the target= attribute of anchor
		tags with the class 'acls' based on the checked 
		property of the element with id 'checkid'
		*/
	var atarg, elems, i
	atarg = document.getElementById(checkid).checked ? '_blank' : '_self' ;
	elems = document.getElementsByClassName(acls) ;
	for (i = 0; i < (elems.length); i++)  {
		elems[i].target = atarg ;
	}
}"
	writeIndex(jsFile, js)
	return true
end writeJS
--
-- WriteCSS handler - if the enstyles.css file doesn't exist, write it to
--   the top level HTML folder
--
on writeCSS()
	set cssFile to (POSIX path of FolderHTML) & myCSSfile -- name of css file
	tell application "Finder" to set foundCSS to (exists cssFile as POSIX file) -- Does CSS exist?
	if foundCSS then return false
	set css to "/*
This CSS stylesheet is created initially by the applescript
Notebook Export Manager to control the display
of index.html pages for exported notebooks and
for the top-level index page.
*/
body {
	font-family: Arial, Helvetica, sans-serif; 
	}
.clnote {
	text-align: center; 
	}
.clfile {
	text-align: center; 
	}
.clsize {
	text-align: right;
	}
.tdate {
	text-align: center; 
	}
table {
	border-collapse: collapse;
	border-spacing: 0;
	font-size: 14px;
	}
th {                     
	cursor: pointer ;
	}
tbody tr:nth-child(odd) { 
	background: #eee; 
	}
th, td {
	padding:4px; 
	}
td a {
	display:block; 
	width:100%; 
	}
.trasc { 
	background-color: #ccff99; 
	} 
.trdesc {
	background-color: #ff9999;
	} 
.ENtag { 
	background-color: #B7D7FF;
	}
.instruct { 
	padding-bottom: 15px; 
	font-size: 14px;
	color: gray; 
	float: left; 
	}
.inpform {
	font-size: 14px;
	background-color: lightgray ;
	float: right ;
	padding: 3px 5px 3px 5px ;
}
.lhdr {
	float: left; 
	font-size: 24px; 
	} 
.rhdr {
	float: right; 
	font-size: 18px; 
	} 
.bhdr {
	clear:both;
	bottom-margin: 15px ;
	}
.hdr2 {
	font-size: 18px;
	padding-top: 20px;
	padding-bottom: 20px;
	} 	
.notecol {
	width:40%;
	} 
.linkhover:hover { 
	 background-color: #99ccff; 
	} 
.zeroNotes { 
	background-color: #ffff99;
	}
"
	writeIndex(cssFile, css)
	return true
end writeCSS
