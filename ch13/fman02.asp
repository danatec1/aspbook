<% 
' FileUpLoad 객체를 사용하여 입력 데이터를 읽는다.
Set objFUL = Server.CreateObject("FileMan.FileUpLoad")
strName = objFUL.GetValueByName("name")
strAge  = objFUL.GetValueByName("age")
%>
<HTML>
<HEAD> 
<TITLE> FileUpload Object (Result) </TITLE>
</HEAD>
<BODY>
<I>FileUpload 객체 (결과)</I><BR>
당신의 이름은 <% = strName %>이고, <BR>  
나이는 <% = strAge %>살입니다.
</BODY>
</HTML>
