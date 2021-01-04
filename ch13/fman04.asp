<%
' 업로드되는 파일들을 저장할 가상 디렉토리를 지정하고
'   이에 대한 물리적인 디렉토리를 구한다.
strVirUPath = "/JFiles"
strPhyUPath = Server.MapPath(strVirUPath) 
'
' FileUpLoad 객체를 사용하여 입력 데이터를 읽는다.
Set objFUL = Server.CreateObject("FileMan.FileUpLoad")
' 파일의 크기를 읽는다.
strUFSize=objFUL.TotalBytes("frmLPFile")
strUFSize=CStr(strUFSize)
' 파일 경로명을 찾는다.
strLPFile = Trim(objFUL.GetValueByName("frmLPFile"))
' 파일 이름만을 추출한다.
strUFName = Mid(strLPFile, InStrRev(strLPFile,"\") +1)
' 파일을 저장할 물리적인 경로명을 구한다. 
strUPFile = strPhyUPath & "/" & strUFName
' On Error Resume Next
' 파일을 저장한다.
strRet = objFUL.SaveFile("frmLPFile", CStr(strUPFile))
'
' 파일 UpLoad 실패
If strRet="" Then
    Set objFUL = Nothing        
%>
    <HTML>
    <HEAD> </HEAD>
    <BODY>
    파일을 저장할 때에 Error가 발생했어요.<BR>
    <%=strUFName%> 파일이 이미 있거나 <BR>
    파일 서버에 이상이 있습니다. <BR> 
    <A HREF="javascript:history.back()">되돌아가기</A>
    </BODY>
    </HTML>
<%
' 파일 UpLoad 성공
Else 
%>
    <HTML>
    <HEAD></HEAD>
    <BODY>
    파일이 성공적으로 저장되었어요. <BR>
    파일 이름은 <%=strUFName%>이고, 크기는 <%= strUFSize%> 이다.<BR>
    <A HREF="javascript:history.back()">되돌아가기</A>
    </BODY>
    </HTML>
<%
    Set objFUL = Nothing
End If 
%>
