<%
strVTime  = Request.Cookies("cooVTime")
strVCount = Request.Cookies("cooVCount")
' 처음 방문하면 Visit Count를 0 으로 설정한다.
If strVCount = "" Then
    strVCount = 0
End If
intVCount = CInt(strVCount)
intVCount = intVCount + 1 
' 마지막 방문 횟수를 재 설정한다.
Response.Cookies("cooVCount") = intVCount
' 현재 날짜와 시간을 구한다.
datNow = now()
' date 형을 문자형으로 바꾼다.
strNow = CStr(datNow) 
' 마지막 방문 시간을 재 설정한다.
Response.Cookies("cooVTime") = strNow
%>
<HTML>
<HEAD> 
<TITLE> Cookie : Visited Counter and Time </TITLE>
</HEAD>
<BODY>
<I>쿠키 : 방문 횟수와 마지막 방문 시간</I> <BR> <BR>
당신은 이곳을 <% = intVCount %>번째 방문하였습니다.<BR>
<%
If strVTime = "" Then %>
첫 방문을 환영합니다.
<%
Else %>
마지막 방문 시간은 <% = strVTime %> 이었습니다.
<%
End If %>
</BODY>
</HTML>
