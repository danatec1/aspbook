<%
' ���� Shopping Cart�� ã�´�. 
strSCID = Session("sesSCID")
' Shopping Cart�� ������,  
If strSCID = "" Then
%>
    <HTML>
    <HEAD> 
    <TITLE> Shopping Cart View </TITLE>
    </HEAD>
    <BODY>
    <TABLE BORDER="1">
    <TR> <TD> ��ǰ�� �������� �ʾҾ��. <TD> </TR>
    </TABLE>
    </BODY>
    </HTML>
<%
' Shopping Cart�� ������,  
' ��ٱ��Ͽ��� ������ ��ǰ���� �д´�. 
Else 
    intSCID = CInt(strSCID)
    Set objConn = Session("objConn")
    sqlQuery = "SELECT SGAmount, GCode, GTitle, GPrice FROM SCart, SGoods"
    sqlQuery = sqlQuery & " WHERE SCart.SCID=" & intSCID
    sqlQuery = sqlQuery & " and SCart.SGCODE = SGoods.GCODE"
    Set objRS = objConn.Execute(sqlQuery)
%>
    <HTML>
    <HEAD> 
    <TITLE> Shopping Cart View </TITLE>
    </HEAD>
    <BODY>
    <I>��ٱ��� ����</I>
    <% 
    IF objRS.BOF and objRS.EOF then %>
        <TABLE BORDER="1">
        <TR> <TD> ��ٱ��ϰ� ������. <TD> </TR>
        </TABLE>
    <% 
    Else %>
        <TABLE BORDER=1>
        <TR>
            <TH> ��ǰ�� </TH>
            <TH> �ܰ� </TH>
            <TH> ���� </TH>
            <TH> �ݾ� </TH>
            <TH COLSPAN=2> ����/���� </TH>
        </TR>     
        <%
	intSum = 0
        Do Until objRS.EOF 
            intPSum = objRS("SGAMOUNT") * objRS("GPrice") 
            intSum = intSum + intPSum  %>
	    <TR> 
	        <TD> <% = objRS("GTitle") %> </TD>
                <TD ALIGN="right"> 
		     <% = FormatCurrency(objRS("GPrice")) %> </TD> 
	        <FORM METHOD="post" ACTION="SCUpdate.asp"> 
                <TD> <INPUT TYPE="text" NAME="GAMOUNT" SIZE=3 
	                 VALUE="<%= objRS("SGAMOUNT") %>"> </TD> 
                <TD ALIGN="right"> 
                     <% = FormatCurrency(intPSum) %>
		     <INPUT TYPE="hidden" NAME="GCODE"  
	                    VALUE="<%= objRS("GCODE") %>"> </TD> 
                <TD> <INPUT TYPE="submit" VALUE="����"> </TD>
		</FORM> 
                <TD> <A HREF="SCDelete.asp?GCODE=<%= objRS("GCode")%>"
		         >����</A></TD>
            </TR>
        <%  objRS.MoveNext
        Loop %>
        <TR> <TD COLSPAN=4 ALIGN="right"> 
	    �� �ݾ� : <% = FormatCurrency(intSum) %> </TD>
	     <TD COLSPAN=2> &nbsp; </TD> 
	</TR>
        </TABLE> 
        <A HREF="SCEmpty.asp">����</A>
        <A HREF="SOForm.asp">�����ϱ�</A>
    <%
    End If 
    objRS.Close
    Set objRS = Nothing
    %>
    </BODY>
    </HTML>
<%
End If
%> 
