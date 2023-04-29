
<%
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = fso.CreateTextFile("d:/tifwww/_private/storelist.csv",TRUE)

Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")


	
	
	
	
set rs = Server.CreateObject("ADODB.recordset")
rs.CursorLocation = 3 
rs.CursorType = 1

a = ""

SQL = "SELECT top 1 * FROM stores order by merchant"
'Response.Write(SQL)
rs.open SQL, conn3


rs.movefirst
			
For each item in rs.Fields

a = a & item.Name & "##"

next


rs.close

x.WriteLine(a)
a = ""
SQL = "SELECT * FROM stores order by merchant"
'Response.Write(SQL)
rs.open SQL, conn3

rs.movefirst
while not rs.EOF
			
	For each item in rs.Fields
	
	
	a = a &item.value & "##"

	next
x.WriteLine(a)	
a = ""
rs.movenext
wend

rs.close

	Set fso = Nothing			
				
	
%>
Done
<body>
</body>
</html>
