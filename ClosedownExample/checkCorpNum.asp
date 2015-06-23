<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
		<title>�������ȸ API SDK ASP Example.</title>
	</head>
	<!--#include file="common.asp"--> 
	<%
			'�������ȸ - �ܰ�

			Dim CorpNum
			
			CorpNum = request.QueryString("CorpNum")		'��ȸ�� ����ڹ�ȣ
			
			If CorpNum <> "" Then
				On Error Resume Next
				
				Set result = closedownChecker.checkCorpNum(CorpNum)

				If Err.Number <> 0 Then
					code = Err.Number
					message = Err.Description
					Err.Clears
				End If
				On Error GoTo 0
			End if

	%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�������ȸ - �ܰ�</legend>
					<div class ="fieldset4">
					<form method= "GET" id="corpnum_form" action="checkCorpNum.asp">
						<%
							If IsEmpty(result) then
						%>
								<input class= "txtCorpNum left" type="text" placeholder="����ڹ�ȣ ����" id="CorpNum" name="CorpNum"  tabindex=1/>
						<%
							Else 
						%>
								<input class= "txtCorpNum left" type="text" placeholder="����ڹ�ȣ ����" id="CorpNum" name="CorpNum"  value="<%=result.corpNum%>" tabindex=1/>
						<%
							End if	
						%>

						<p class="find_btn find_btn01 hand" onclick="search()" tabindex=2>��ȸ</p>
					</form>
					</div>
			</fieldset>
			<%
				If Not IsEmpty(result) Then  
			%>
				<fieldset class="fieldset2">
					<legend>�������ȸ - �ܰ�</legend>
					<ul>
						<li>����ڹ�ȣ(corpNum) : <%= result.corpNum%></li>		
						<li>���������(type) : <%= result.ctype%></li>	
						<li>���������(state) : <%= result.state%></li>
						<li>���������(stateDate) : <%= result.stateDate%></li>	
						<li>����û Ȯ������(checkDate) : <%= result.checkDate%></li>	
					</ul>
					<p class="info">> state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�</p>
					<p class="info">> type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������</p>
					<br/>
				</fieldset>
			<%
				End If 
				If Not IsEmpty(code) then
			%>
				<fieldset class="fieldset2">
					<legend>�������ȸ - �ܰ�</legend>
					<ul>
						<li>Response.code : <%= code %> </li>
						<li>Response.message : <%= message %></li>
					</ul>
				</fieldset>
			<%
				End If
			%>		
		 </div>

		<script type ="text/javascript">
			 window.onload=function(){
				 document.getElementById('CorpNum').focus();
			 }
			 
			 function search(){
				document.getElementById('corpnum_form').submit();
			 }		 
		 </script>
	</body>
</html>