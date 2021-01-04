<%
' 입력 데이터들을 읽는다.
strX = Request.Form.Item("varX")
strY = Request.Form("varY")
' 입력 데이터들을 정수 타입으로 변환하고 더하기 한다.
intX = CInt(strX)
intY = CInt(strY)
intSum = intX + intY
%>
<HTML>
<HEAD> 
<TITLE> Example : Form Collection(Result) </TITLE>
</HEAD>
<BODY>
<I>Form 컬렉션(데이터 처리)</I><BR>
입력 데이터 값 : 
X = <% = strX %>, Y = <% = strY %> <BR>
두 수의 합 : <% = intSum %>
</BODY>
</HTML>
