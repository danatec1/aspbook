<%
' �Է� �����͸� �д´�.
strGCode = Request.QueryString("GCODE")
' �����ͺ��̽����� �ش� ��ǰ ������ �о�´�. 
Set objConn = Session("objConn")
sqlQuery = "SELECT * FROM SGoods WHERE GCODE='" & strGCode & "'"
set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> Goods Detail View </TITLE>
</HEAD>
<BODY>
<% 
IF objRS.BOF and objRS.EOF then %>
    <TABLE BORDER="1">
    <TR> <TD> �ش� ��ǰ�� ���� ������ �����. <TD> </TR>
    </TABLE>
<% 
Else %>
    <TABLE>
    <FORM METHOD="post" ACTION="SCInsert.asp">
<%  Do Until objRS.EOF %>
	<TR> <TD> <% = objRS("GTitle") %> </TD> </TR>
        <TR> <TD> <% = objRS("GDescription") %> </TD> </TR>
        <TR> <TD> <HR> </TD> </TR>
        <TR> <TD> 
	    ���� : <% = FormatCurrency(objRS("GPrice")) %> �� &nbsp;  
            ���� : <INPUT TYPE="text" SIZE=3 NAME="GAmount" 
		       VALUE="1"> �� &nbsp; 
	           <INPUT TYPE="hidden" NAME="GCODE" 
		       VALUE="<%=objRS("GCODE")%>">
	           <INPUT TYPE="submit" VALUE="��ٱ��Ϸ�">
	     </TD>
        </TR>
<%      objRS.MoveNext
    Loop %> 
    </FORM>
    </TABLE>
<%
End If %>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
%>
