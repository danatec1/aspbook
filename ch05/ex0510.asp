<% 
strS = "I am a boy"
%>
<HTML>
<HEAD>
<TITLE> Example : String Functions </TITLE>
</HEAD>
<BODY BGCOLOR=white>
문자열 데이터 : "I am a boy" 
<HR>
문자열 길이 : (<% = Len(strS) %>) <BR>
왼쪽에서 4 자 추출 : (<% = Left(strS, 4) %>) <BR>
중간에서 2 자 추출 : (<% = Mid(strS, 3, 2) %>) <BR>
오른쪽에서 3 자 추출 : (<% = Right(strS, 3) %>) <BR>
문자열을 대문자로 : (<% = UCase(strS) %>) <BR>
문자열을 소문자로 : (<% = LCase(strS) %>) <BR>
</BODY>
</HTML>
