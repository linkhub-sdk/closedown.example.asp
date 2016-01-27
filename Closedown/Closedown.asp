<!--#include file="Linkhub/Linkhub.asp"--> 
<%
Application("LINKHUB_TOKEN_SCOPE_CLOSEDOWN") = Array("170")
Const ServiceID= "CLOSEDOWN"
Const ServiceURL = "https://closedown.linkhub.co.kr"
Const APIVersion = "1.0"

Class Closedown
Private m_TokenDic
Private m_Linkhub
Public Sub Class_Initialize
	On Error Resume next
	If  Not(CLOSEDOWN_TOKEN_CACHE Is Nothing) Then
		Set m_TokenDic = CLOSEDOWN_TOKEN_CACHE
	Else
		Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
	End If
	On Error GoTo 0
	If isEmpty(m_TokenDic) Then
		Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
	End If
	Set m_Linkhub = New Linkhub
End Sub
Private Property Get m_scope
	m_scope = Application("LINKHUB_TOKEN_SCOPE_CLOSEDOWN")
End Property
Public Sub Initialize(linkID, SecretKey )
    m_Linkhub.LinkID = linkID
    m_Linkhub.SecretKey = SecretKey
End Sub

Public Function getSession_token()
    refresh = False
    Set m_Token = Nothing
	
	If m_TokenDic.Exists("token") Then 
		Set m_Token = m_TokenDic.Item("token")
	End If
	
    If m_Token Is Nothing Then
        refresh = True
    Else
		'CheckScope
		For Each scope In m_scope
			If InStr(m_Token.strScope,scope) = 0 Then
				refresh = True
				Exit for
			End if
		Next
		If refresh = False then
			Dim utcnow
			utcnow = CDate(Replace(left(m_linkhub.getTime,19),"T" , " " ))
			refresh = CDate(Replace(left(m_Token.expiration,19),"T" , " " )) < utcnow
		End if
    End If
    
    If refresh Then
		If m_TokenDic.Exists("token") Then m_TokenDic.remove "token"
        Set m_Token = m_Linkhub.getToken(ServiceID, null, m_scope)
		m_Token.set "strScope", Join(m_scope,"|")
		m_TokenDic.Add "token", m_Token
	End If
    
	getSession_token = m_Token.session_token
End Function

Public Function GetBalance()
    GetBalance = m_Linkhub.GetPartnerBalance(getSession_token(), ServiceID)
End Function

Public Function GetUnitCost()
	Set result = httpGET("/UnitCost", getSession_token())
	GetUnitCost = result.unitCost
End Function 

Public Function checkCorpNum(CorpNum)
	If CorpNum = "" Then
        Err.Raise -99999999, "CLOSEDOWN", "사업자번호가 입력되지 않았습니다."
	End If 
	url = "/Check?CN=" + CorpNum

	Set result = New CorpState
	Set tmp = httpGET(url, getSession_token())

	result.fromJsonInfo tmp
	
	Set checkCorpNum = result
End Function 

Public Function checkCorpNums(CorpNumList)
    If isEmpty(CorpNumList) Then
        Err.Raise -99999999, "CLOSEDOWN", "사업자번호 배열이 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("[]")
	For i=0 To UBound(CorpNumList) -1
		tmp.Set i, CorpNumList(i)
	Next

	postdata = toString(tmp)

	Set result = httpPOST("/Check", getSession_token(), postdata)
	
	Set infoObj = CreateObject("Scripting.Dictionary")
	For i=0 To result.length-1
		Set tmp = New CorpState
		tmp.fromJsonInfo result.Get(i)
		infoObj.Add i, tmp
	Next
		
	Set checkCorpNums = infoObj
End Function

'Private Functions
Public Function httpGET(url, BearerToken)
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", ServiceURL + url)
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-api-version", APIVersion)
    
    winhttp1.Send
    winhttp1.WaitForResponse
    result = winhttp1.responseText
       
    If winhttp1.Status <> 200 Then
		Set winhttp1 = Nothing
        Set parsedDic = m_Linkhub.parse(result)
        Err.Raise parsedDic.code, "CLOSEDOWN", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(result)
End Function

Public Function httpPOST(url , BearerToken , postdata)
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", ServiceURL + url)
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)    
    Call winhttp1.setRequestHeader("x-api-version", APIVersion)
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    result = winhttp1.responseText
    
    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
		Set parsedDic = m_Linkhub.parse(result)
        Err.Raise parsedDic.code, "CLOSEDOWN", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    Set httpPOST = m_Linkhub.parse(result)

End Function

public Function toString(object)
	toString = m_Linkhub.toString(object)
End Function
End Class

Class CorpState
	Public corpNum
	Public ctype
	Public state
	Public stateDate
	Public checkDate

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.corpNum) Then
				corpNum = jsonInfo.corpNum
			End If 

			If Not isEmpty(jsonInfo.type) Then
				ctype = jsonInfo.type
			End If 

			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If 

			If Not isEmpty(jsonInfo.stateDate) Then
				stateDate = jsonInfo.stateDate
			End If 

			If Not isEmpty(jsonInfo.checkDate) Then
				checkDate = jsonInfo.checkDate
			End If 
		On Error GoTo 0
	End Sub
End Class
%>