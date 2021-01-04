<%
' 고객의 Shopping Cart를 찾는다. 
strSCID = Session("sesSCID")
' Shopping Cart가 없으면,  
If strSCID = "" Then
%>
    <HTML>
    <HEAD> 
    <TITLE> Shopping Cart View </TITLE>
    </HEAD>
    <BODY>
    <TABLE BORDER="1">
    <TR> <TD> 상품을 구입하지 않았어요. <TD> </TR>
    </TABLE>
    </BODY>
    </HTML>
<%
' Shopping Cart가 있으면,  
' 장바구니에서 구입한 상품들을 읽는다. 
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
    <I>장바구니 보기</I>
    <% 
    IF objRS.BOF and objRS.EOF then %>
        <TABLE BORDER="1">
        <TR> <TD> 장바구니가 비었어요. <TD> </TR>
        </TABLE>
    <% 
    Else %>
        <TABLE BORDER=1>
        <TR>
            <TH> 상품명 </TH>
            <TH> 단가 </TH>
            <TH> 수량 </TH>
            <TH> 금액 </TH>
            <TH COLSPAN=2> 수정/삭제 </TH>
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
                <TD> <INPUT TYPE="submit" VALUE="수정"> </TD>
		</FORM> 
                <TD> <A HREF="SCDelete.asp?GCODE=<%= objRS("GCode")%>"
		         >삭제</A></TD>
            </TR>
        <%  objRS.MoveNext
        Loop %>
        <TR> <TD COLSPAN=4 ALIGN="right"> 
	    총 금액 : <% = FormatCurrency(intSum) %> </TD>
	     <TD COLSPAN=2> &nbsp; </TD> 
	</TR>
        </TABLE> 
        <A HREF="SCEmpty.asp">비우기</A>
        <A HREF="SOForm.asp">구매하기</A>
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
