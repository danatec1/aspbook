<% 
strArea = Request.QueryString("area")
strSex  = Request.QueryString("sex")
%>
<HTML>
<HEAD> 
<TITLE> Example : QueryString Collection(Form Result) </TITLE>
</HEAD>
<BODY>
<I>QueryString 컬렉션(폼 데이터 처리)</I> <BR> 
당신은 <% = strArea %>에 사는 <% = strSex %>이군요.
</BODY>
</HTML>
