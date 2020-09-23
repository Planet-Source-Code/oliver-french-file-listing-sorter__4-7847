<div align="center">

## File Listing Sorter


</div>

### Description

This shows how to display files in a folder sorted by parameter (name,size date). You can sort in either direction (ascending or descending).
 
### More Info
 
The sorting function takes the array to sort, the direction to sort, and the index number of the first dimension to sort by.

Sort array is similar to Andy Slowey's sort array function but adapted to do multi-dimensional.

The sorting function returns a sorted array.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Oliver French](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/oliver-french.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/oliver-french-file-listing-sorter__4-7847/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = True %>
<%
Dim sSort
Dim FileArray() ' (0 is name, 1 is size, and 2 is last modified)
dim objFSO
dim objFile
dim objFolder
dim sMapPath
dim sUrlPath
Dim filename, filecollection, filesize, modified
Dim i
Dim FinalFileArray
' Make path strings (change this to the path of the folder you want to see)
sUrlPath = "iDesign/pages/Sample Files"
sMapPath = Server.MapPath(sUrlPath)
' Create FSO object and buddies
set objFSO = createobject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(sMapPath)
Set filecollection = objFolder.Files
' Resize file array to appropriate number of files to hold (with 3 attributes to hold
Redim FileArray(2,filecollection.count - 1)
Redim FinalFileArray(2,filecollection.count - 1)
' Load file data into array
i = 0
For Each objFile In objFolder.Files
	filename=right(objFile.name,len(objFile.name)-InStrRev(objFile.name, "\"))
	filesize = objFile.size
	modified = objFile.DateLastModified
	FileArray(0, i) = filename
	FileArray(1, i) = filesize
	FileArray(2, i) = modified
	i = i + 1
Next
' Sort array according to method
sSort = Request("Sort")
Select case sSort
	Case "Name"
		FinalFileArray = SortAlpha(FileArray,"DESC",0)
	Case "Size"
		FinalFileArray = SortAlpha(FileArray,"DESC",1)
	Case "Modified"
		FinalFileArray = SortAlpha(FileArray,"DESC",2)
End Select
' ***********************************************************
' ***********************************************************
' **	2-D sorting function 							 **
' ***********************************************************
' ***********************************************************
Function SortAlpha(ary, direction, indexnum)
	Dim StopWork
	Dim i
	dim i2
	Dim firstval()
	Dim secondval()
	redim firstval(ubound(ary,1))
	redim secondval(ubound(ary,1))
	StopWork=False
	Do Until StopWork=True
		StopWork=True
		For i = 0 To UBound(ary,2)
			if i=UBound(ary,2) Then Exit For
			if UCase(Direction) = "DESC" Then
				if ary(indexnum,i) < ary(indexnum,i+1) Then
					For i2 = 0 to ubound(firstval)
						firstval(i2) = ary(i2,i)
						secondval(i2) = ary(i2,i+1)
						ary(i2,i) = secondval(i2)
						ary(i2,i+1) = firstval(i2)
					Next
					StopWork=False
				End if
			Else
				if ary(indexnum,i) > ary(indexnum,i+1) Then
					For i2 = 0 to ubound(firstval)
						firstval(i2) = ary(i2,i)
						secondval(i2) = ary(i2,i+1)
						ary(i2,i) = secondval(i2)
						ary(i2, i+1) = firstval(i2)
					Next
					StopWork=False
				End if
			End if
		Next
	Loop
	SortAlpha=ary
End Function
%>
<html>
<head>
<title></title>
</head>
<body>
<table width="75%" border="2" cellspacing="0" cellpadding="2" align="center" bgcolor="#FFFFCC" bordercolor="#000000">
 <tr>
 <td> <!-- The name of this particular file was Directory.asp -->
 <div align="center"><b><a href="Directory.asp?Sort=Name">Name</a></b></div>
 </td>
 <td>
 <div align="center"><b><a href="Directory.asp?Sort=Size">Size</a></b></div>
 </td>
 <td>
 <div align="center"><b><a href="Directory.asp?Sort=Modified">Modified</a></b></div>
 </td>
 </tr>
<%
	For i = 0 to Ubound(FinalFileArray,2)
		filename=FinalFileArray(0,i)
		filesize = FinalFileArray(1,i)
		modified = FinalFileArray(2,i)
		Response.Write "<tr><td>" & FileName & "</td><td>" & filesize & "</td><td>" & modified & "</td></tr>"
	Next
%>
</table>
</body>
</html>
```

