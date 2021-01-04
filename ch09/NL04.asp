<%
strLFile = "NL01.txt"
Set objNL = Server.CreateObject("MSWC.NextLink")
intLCount = objNL.GetListCount(strLFile)
intLIndex = objNL.GetListIndex(strLFile)
%>
<HTML>
<HEAD>
<TITLE> Mac OS </TITLE>
</HEAD>
<BODY>
<I>Macintosh 운영 체제</I> <BR>
돈이 없어 제대로 이용 못해본 운영 체제 !!!<BR>
<!-- 이동 링크 만들기 시작 -->
<% 
' 이전 웹 페이지로 이동 링크
If (intLIndex > 1) Then
    strPrevURL = objNL.GetPreviousURL(strLFile) %>
    [<A HREF="<%=strPrevURL%>">이전</A>] &nbsp;
<% 
End If
' 다음 웹 페이지로 이동 링크
If (intLIndex < intLCount) Then
    strNextURL = objNL.GetNextURL(strLFile) %>
    [<A HREF="<%=strNextURL%>">다음</A>]
<% 
End If %>
<!-- 이동 링크 만들기 끝 -->
</BODY>
</HTML>
