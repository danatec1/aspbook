<HTML>
<HEAD>
<TITLE> AdRotatotor Object and Redirect File </TITLE>
</HEAD>
<BODY>
<I>AdRotator 객체와 리디렉트 파일</I> <BR>
<% 
Dim objAR, varImage
Set objAR = Server.CreateObject("MSWC.AdRotator")
varImage = objAR.GetAdvertisement("AR03.txt")
Response.Write varImage
%>
</BODY>
</HTML>
