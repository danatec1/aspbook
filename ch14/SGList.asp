<%
' �����ͺ��̽����� ��ǰ ��ϵ��� �о�´�. 
Set objConn = Session("objConn")
sqlQuery = "SELECT * FROM SGoods "
Set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> Shop Goods List </TITLE>
</HEAD>
<BODY>
<I>��ǰ���</I> &nbsp; &nbsp;
<A HREF="SCView.asp" TARGET="_blank">��ٱ��Ϻ���</A>
<TABLE BORDER="1">
<TR>
    <TH> ��ǰ�� </TH>
    <TH> ���� </TH>
    <TH> ��ǰ���� </TH>
</TR>
<% 
IF objRS.BOF and objRS.EOF then %>
    <TR>
        <TD COLSPAN=3> ��ǰ�� �ϳ��� �����. </TD>
    </TR>
<% 
Else
    Do Until objRS.EOF 
%>
	<TR>
	    <TD> <% = objRS("GTitle") %> </TD>
            <TD ALIGN="right"> 
                   <% = FormatCurrency(objRS("GPrice")) %> </TD>
	    <TD> <A HREF="SGView.asp?GCODE=<% = objRS("GCODE")%>"
	             TARGET="_blank">��ǰ����</A> </TD>
        </TR>
<%      objRS.MoveNext
    Loop 
End If %>
</TABLE>
<!-------------------------------------
  �˻� �޴��� �����ִ� �κ� 
--------------------------------------->
<FORM METHOD="get" ACTION="SGSList.asp"> 
    <SELECT NAME="SField">
        <OPTION VALUE="GTitle" SELECTED> ��ǰ��
        <OPTION VALUE="GDescription"> ��ǰ����
    </SELECT> 
    <INPUT TYPE="text" NAME="SString">
    <INPUT TYPE="submit" VALUE="�˻�">
</FORM>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
%>
