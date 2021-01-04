<HTML>
<HEAD>
<TITLE> Example : For...Next Statement </TITLE>
</HEAD>
<BODY BGCOLOR=white>
1..20까지 짝수 출력 <BR>
<% For I = 2 To 20 Step 2 %> 
        &nbsp; <% = I %> 
<% Next %> <BR>
<HR>
<% For I = 1 To 7 %> 
        <FONT SIZE=<%=I%>>
        글자가 커져요 !!!
        </FONT> <BR> 
<% Next %> 
</BODY>
</HTML>
