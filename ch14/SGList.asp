<%
' 데이터베이스에서 상품 목록들을 읽어온다. 
Set objConn = Session("objConn")
sqlQuery = "SELECT * FROM SGoods "
Set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> Shop Goods List </TITLE>
</HEAD>
<BODY>
<I>상품목록</I> &nbsp; &nbsp;
<A HREF="SCView.asp" TARGET="_blank">장바구니보기</A>
<TABLE BORDER="1">
<TR>
    <TH> 상품명 </TH>
    <TH> 가격 </TH>
    <TH> 상품정보 </TH>
</TR>
<% 
IF objRS.BOF and objRS.EOF then %>
    <TR>
        <TD COLSPAN=3> 상품이 하나도 없어요. </TD>
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
	             TARGET="_blank">상품정보</A> </TD>
        </TR>
<%      objRS.MoveNext
    Loop 
End If %>
</TABLE>
<!-------------------------------------
  검색 메뉴를 보여주는 부분 
--------------------------------------->
<FORM METHOD="get" ACTION="SGSList.asp"> 
    <SELECT NAME="SField">
        <OPTION VALUE="GTitle" SELECTED> 상품명
        <OPTION VALUE="GDescription"> 상품정보
    </SELECT> 
    <INPUT TYPE="text" NAME="SString">
    <INPUT TYPE="submit" VALUE="검색">
</FORM>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
%>
