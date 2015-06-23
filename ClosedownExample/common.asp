<!--#include virtual="/Closedown/Closedown.asp"--> 
<%
	'연동상담시 발급받은 링크아이디 
	LinkID = "TESTER"
	'연동상담시 발급받은 비밀키, 유출에 주의
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
	set  closedownChecker = new Closedown
	closedownChecker.Initialize LinkID, SecretKey
%>