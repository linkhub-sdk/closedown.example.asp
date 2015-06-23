<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>휴폐업조회 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%

	On Error Resume Next
	remainPoint = closedownChecker.getBalance()
	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If
	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>잔여포인트 확인</legend>
				<% If code = 0 Then %>
					<ul>
						<li>잔여포인트 : <%=CStr(remainpoint)%> </li>
					</ul>
				<%	Else  %>
				<ul>
					<li>Response.code: <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>