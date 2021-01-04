<!-- #include file="SJConf.inc" -->
<%
' 사용자의 입력들을 읽는다.
Set objFUL = Server.CreateObject("FileMan.FileUpLoad")
strUName = objFUL.GetValueByName("frmUName")
strSubject = objFUL.GetValueByName("frmSubject")
strContent = objFUL.GetValueByName("frmContent")
strLPFile = objFUL.GetValueByName("frmLPFile")
' 단일(') 인용 부호 문제를 없앤다. 
strSubject = Replace(strSubject, "'", "''")
strContent = Replace(strContent, "'", "''")
'
' Upload 파일이 없으면  
If strLPFile = "" Then
'   파일 정보를 제외한 데이터들을 데이터베이스에 저장한다.
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open ConnString
    sqlQuery = "INSERT INTO JMessage VALUES (" 
    sqlQuery = sqlQuery & "'" & strSubject & "'" 
    sqlQuery = sqlQuery & ",'" & strContent & "'"
    sqlQuery = sqlQuery & ",'" & strUName & "'"
    sqlQuery = sqlQuery & ",'','')" 
    objConn.Execute(sqlQuery) 
    objConn.Close
    Set objConn = Nothing
    Set objFUL = Nothing
%>
    <HTML>
    <HEAD>
    <TITLE> New JaRyo </TITLE>
    </HEAD>
    <BODY>
    자료가 등록되었어요. <BR>
    첨부 파일은 없습니다. <BR>
    <A HREF="SJList.asp">목록보기</A>
    </BODY>
    </HTML>
<%
'
' Upload 파일이 있으면  
Else
'   파일의 크기를 읽는다.
    strUFSize=objFUL.TotalBytes("frmLPFile")
    strUFSize=CStr(strUFSize)
'   파일 이름만을 추출한다.
    strUFName = Mid(strLPFile, InStrRev(strLPFile,"\")+1)
    strUPFile = strPhyUDir & "/" & strUFName
'   On Error Resume Next
'   파일을 저장한다.
    strRet = objFUL.SaveFile("frmLPFile", CStr(strUPFile))
'   파일 UpLoad 실패
    If strRet="" Then
        Set objFUL = Nothing        
%>
        <HTML>
        <HEAD>
        <TITLE> New JaRyo </TITLE>
        </HEAD>
        <BODY>
        파일을 저장할 때에 Error가 발생했어요.<BR>
        <%=strUFName%>이 서버에 이미 있거나 <BR>
        자료실 서버에 이상이 있습니다. <BR> 
        <A HREF="SJList.asp">목록보기</A>
        </BODY>
        </HTML>
<%
'   파일 UpLoad 성공
    Else 
'       데이터베이스에 데이터들을 저장한다.
        Set objConn = Server.CreateObject("ADODB.Connection")
        objConn.Open ConnString
        sqlQuery = "INSERT INTO JMessage VALUES (" 
        sqlQuery = sqlQuery & "'" & strSubject & "'" 
        sqlQuery = sqlQuery & ",'" & strContent & "'"
        sqlQuery = sqlQuery & ",'" & strUName & "'"
        sqlQuery = sqlQuery & ",'" & strUFName & "'"
        sqlQuery = sqlQuery & ",'" & strUFSize & "')"
        objConn.Execute(sqlQuery) 
        objConn.Close
        Set objConn = Nothing
        Set objFUL = Nothing
%> 
        <HTML>
        <HEAD>
        <TITLE> New JaRyo </TITLE>
        </HEAD>
        <BODY>
        자료가 등록되었어요. <BR>
        첨부된 파일 이름은 <%=strUFName%>이고, <BR>
        파일 크기는 <%= strUFSize%> 이다. <BR>
        <A HREF="SJList.asp">목록보기</A>
        </BODY>
        </HTML>
<%
    End If
End If
%>
