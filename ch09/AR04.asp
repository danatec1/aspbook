<HTML>
<HEAD>
<TITLE> AdRotatotor Object and Redirect File </TITLE>
</HEAD>
<BODY>
<I>AdRotator ��ü�� ����Ʈ ����</I> <BR>
<% 
Dim objAR, varImage
Set objAR = Server.CreateObject("MSWC.AdRotator")
varImage = objAR.GetAdvertisement("AR03.txt")
Response.Write varImage
%>
</BODY>
</HTML>
