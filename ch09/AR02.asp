<HTML>
<HEAD>
<TITLE> AdRotatotor Object and Scheduler File </TITLE>
</HEAD>
<BODY>
<I>AdRotator ��ü�� ��ȹ ����</I> <BR>
<% 
Dim objAR, varImage
Set objAR = Server.CreateObject("MSWC.AdRotator")
varImage = objAR.GetAdvertisement("AR01.txt")
Response.Write varImage
%>
</BODY>
</HTML>
