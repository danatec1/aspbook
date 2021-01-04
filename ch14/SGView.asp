<%
' 입력 데이터를 읽는다.
strGCode = Request.QueryString("GCODE")
' 데이터베이스에서 해당 상품 정보를 읽어온다. 
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
    <TR> <TD> 해당 상품의 상세한 정보가 없어요. <TD> </TR>
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
	    가격 : <% = FormatCurrency(objRS("GPrice")) %> 원 &nbsp;  
            수량 : <INPUT TYPE="text" SIZE=3 NAME="GAmount" 
		       VALUE="1"> 개 &nbsp; 
	           <INPUT TYPE="hidden" NAME="GCODE" 
		       VALUE="<%=objRS("GCODE")%>">
	           <INPUT TYPE="submit" VALUE="장바구니로">
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
