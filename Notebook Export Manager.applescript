--
   HTML: " & cond(ExportAsHTML, POSIX path of FolderHTML, "(not to be exported)") & "
   ENEX: " & cond(ExportAsENEX, POSIX path of FolderENEX, "(not to be exported)") & "
" & cond(length of SearchString > 0, "Only notes meeting this criteria will be exported:
" & SearchString, "All notes will be exported in each notebook") & "

Start export now or change settings."
" & countHTMLdone & " of " & countHTMLtried & " exported as HTML and
" & countENEXdone & " of " & countENEXtried & " exported as ENEX

(Export time (secs): " & (timerAll as integer) & ", index read: " & (timerSub1 as integer) & ")")
"
"
"
"
"
      <td>" & spanTag & "</td>
" & notelink & noteLinkFar & "      <td>" & parentCell & "</td>
" & noteBookCell & "    </tr>
"
" & notelink & my dateCell(aNoteCreate) & my dateCell(aNoteMod) & fmtSize(sz) & "      <td class='clfile'>" & aNoteAttach & "</td>
" & aNoteAuthor & "      <td>" & allTagNames & "</td></tr> 
"
" & noteBookCell & "      <td class='clnote' data-notes='" & yy & "'>" & yy & "</td>
" & dateCell(bookCreate) & dateCell(bookMod) & dateCell(expStamp) & fmtSize(dirSize) & "      <td class='clfile' data-files='" & bookFiles & "'>" & bookFiles & "</td>
    </tr>
"
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
", "<!-- taglist start -->
", "data-notes='", "data-foldersize='", "data-files='"}, {"bookindex -->", "<!-- taglist end -->", "'", "'", "'"}, 0) -- read the index.html page in the subfolder
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
"
Timeout value must be a positive integer." -- set error message
Timeout value must be a positive integer." -- set error message
"
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
Rerun this script and choose Change Settings to set " & omode & " export folder."
Choose another folder."
" -- reuse prior comment
     running on macOS: " & (system version of sys) & "  using AppleScript ver: " & (AppleScript version of sys) & "
     on computer: " & (computer name of sys) & "  by user: " & (long user name of sys) & "
     user locale: " & (user locale of sys) & "  GMT offset: " & tzone() & "  local date: " & fmtLongDate(current date) & "   -->
"
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