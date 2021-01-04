<% 
strName = Request.QueryString("name")
strAge  = Request.QueryString("age")
%>
<HTML>
<HEAD> 
<TITLE> Example : QueryString Collection(Result) </TITLE>
</HEAD>
<BODY>
<I>QueryString 컬렉션(데이터 처리)</I><BR>
당신의 이름은 <% = strName %>이고, <BR>  
나이는 <% = strAge %>살입니다.
</BODY>
</HTML>
