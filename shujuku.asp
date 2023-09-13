<html>
<body>

<p>
    <%set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

sql="INSERT INTO score (s_name,s_score)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("fname") & "',"
sql=sql & "'" & Request.Form("score") & "')"

on error resume next
conn.Execute sql,recaffected
if err<>0 then
  Response.Write("No update permissions!")
else 
  Response.Write("<h3>" & recaffected & " record added</h3>")
end if
conn.close
%>
</p>
<p>
<%
Dim conn, rs, sql

' 创建数据库连接对象
set conn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

' SQL查询
sql = "SELECT s_name, s_score FROM score ORDER BY s_score DESC"

' 执行查询
rs.Open sql, conn

' 显示查询结果
Response.Write "<table border='1' width='50%' cellspacing='0' cellpadding='2'>"
Response.Write "<tr><th>Name</th><th>Score</th></tr>"

Do While Not rs.EOF
    Response.Write "<tr>"
    Response.Write "<td>" & rs("s_name") & "</td>"
    Response.Write "<td>" & rs("s_score") & "</td>"
    Response.Write "</tr>"
    rs.MoveNext
Loop

Response.Write "</table>"

' 清理并关闭连接
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
%>
<%


' 创建数据库连接对象
Set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

' DELETE语句
sql = "DELETE FROM score WHERE s_score = ' '"

' 执行DELETE语句
conn.Execute sql

' 关闭连接
conn.Close
Set conn = Nothing
%>
</p>
</body>
</html>