<% 
strS = " You ARE a girl "
%>
<HTML>
<HEAD>
<TITLE> Example : String Functions </TITLE>
</HEAD>
<BODY BGCOLOR=white>
���ڿ� ������ : " You ARE a girl " 
<HR>
���ڿ� ���� : (<% = Len(strS) %>) <BR>
���� ���� ���� ���� : (<% = LTrim(strS) %>) 
    ���� : (<% = Len(LTrim(strS)) %>) <BR>
������ ���� ���� ���� : (<% = RTrim(strS) %>)
    ���� : (<% = Len(RTrim(strS)) %>) <BR>
���� ������ ���� ���� ���� : (<% = Trim(strS) %>)
    ���� : (<% = Len(Trim(strS)) %>) <BR>
���ڿ� �ٲٱ� : (<% = Replace(strS, "ARE", "are") %>)
    ���� : (<% = Len(Replace(strS, "ARE", "are")) %>)
</BODY>
</HTML>
