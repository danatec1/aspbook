<%
Set objNL =Server.CreateObject("MSWC.NextLink")
intLCount =objNL.GetListCount("NL01.txt")
%>
<HTML>
<HEAD>
<TITLE> Operating System </TITLE>
</HEAD>
<BODY>
<I>� ü�� ����</I>
<UL>
<% For intI=1 To intLCount %>
<LI> <A HREF="<% =objNL.GetNthURL("NL01.txt", intI) %>">
      <% =objNL.GetNthDescription("NL01.txt", intI) %>
      </A>
<% Next %>
</UL>
</BODY>
</HTML>
