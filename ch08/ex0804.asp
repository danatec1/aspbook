<% 
' test를 위해 페이지를 캐쉬하지 않도록 지정
Response.Expires = 0
' Timeout 속성을 이해하기 위해서 1분으로 지정 
Session.Timeout = 1
%>
<%
' 세션 변수가 없으면 초기값을 0으로 한다.
If Session("sesVisited") = "" Then
    Session("sesVisited") = 0
End If
' 세션 변수의 값을 하나 증가시킨다.
Session("sesVisited") = Session("sesVisited") + 1
' 세션 변수의 값이 4가 되면 객체를 파괴한다.
If (Session("sesVisited") = "5") Then
    Session.Abandon()
End If
%>
<HTML>
<HEAD> 
<TITLE> SessionID, Timeout Attribute and Abandon method </TITLE>
</HEAD>
<BODY>
<I>SessionID, Timeout Attribute and Abandon method</I> <BR>
<FONT COLOR="red">1분 후에 Session이 종료된다.</FONT><BR>
SessionID = <% = Session.SessionID %> <BR>
방문수 : <% = Session("sesVisited") %> <BR>
<A HREF="ex0804.asp">다시읽기</A>
</BODY>
</HTML>
