<HTML>
<HEAD> 
<TITLE> Example : ServerVariables Collection </TITLE>
</HEAD>
<BODY>
<I>ServerVariables Collection</I> <BR>
<%
For Each Item in Request.ServerVariables
    For intCount=1 To Request.ServerVariables(Item).Count
%>
        <% = (Item & "=" & Request.ServerVariables(Item)(intCount)) %>
        <BR>
<%  Next
Next
%>
</BODY>
</HTML>
