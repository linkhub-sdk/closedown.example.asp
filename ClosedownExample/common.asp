<!--#include virtual="/Closedown/Closedown.asp"--> 
<%
	'�������� �߱޹��� ��ũ���̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
	set  closedownChecker = new Closedown
	closedownChecker.Initialize LinkID, SecretKey
%>