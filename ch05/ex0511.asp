<% 
strS = " You ARE a girl "
%>
<HTML>
<HEAD>
<TITLE> Example : String Functions </TITLE>
</HEAD>
<BODY BGCOLOR=white>
문자열 데이터 : " You ARE a girl " 
<HR>
문자열 길이 : (<% = Len(strS) %>) <BR>
왼쪽 공백 문자 제거 : (<% = LTrim(strS) %>) 
    길이 : (<% = Len(LTrim(strS)) %>) <BR>
오른쪽 공백 문자 제거 : (<% = RTrim(strS) %>)
    길이 : (<% = Len(RTrim(strS)) %>) <BR>
왼쪽 오른쪽 공백 문자 제거 : (<% = Trim(strS) %>)
    길이 : (<% = Len(Trim(strS)) %>) <BR>
문자열 바꾸기 : (<% = Replace(strS, "ARE", "are") %>)
    길이 : (<% = Len(Replace(strS, "ARE", "are")) %>)
</BODY>
</HTML>
