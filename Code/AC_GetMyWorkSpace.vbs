Sub Form_OnLoad(Form)

	'Load User List
	
	Call Form.Fields.Field("UserID").RemoveAll()
	Call Form.Fields.Field("UserID").AddListItem("Please Select a User", "Select")		
	
	Dim RecordReturned
	Dim Connection, Recordset, SQL, Server, field, strAllFields
	Dim RecordLabel
	Dim RecordValue
		
	SQL = "SELECT [USERID],[USERIDEX],[FULLNAME],[USERNUM] FROM [Active].[MHGROUP].[DOCUSERS] WHERE [LOGIN] = 'Y'"
	
	Set Connection = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")

	Connection.Open "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1"
	Recordset.Open SQL,Connection
	If Recordset.EOF Then
		RecordLabel = "No records found"
		RecordValue = "No records found"
		Call Form.Fields.Field("UserID").AddListItem(RecordLabel, RecordValue)		
	Else
		Do While Not Recordset.Eof
			Dim userID : userID = Recordset("USERID")
			Dim userName : userName = Recordset("FULLNAME")
			RecordLabel = userName
			RecordValue = userID 
			Call Form.Fields.Field("UserID").AddListItem(RecordLabel, RecordValue)
			Recordset.MoveNext
		Loop
	End If
	Recordset.Close
	Set Recordset=Nothing
	Connection.Close
	Set Connection=Nothing
	
End Sub 


Function Form_OnValidate(Form)

End Function


Sub Field_OnChange(Form, FieldName, FieldValue)

	Dim RecordReturned
	Dim Connection, Recordset, SQL, Server, field, strAllFields
	Dim RecordLabel
	Dim RecordValue
	
	If FieldName = "UserID" Then
		
		'Load WorkSpace List
	
		Call Form.Fields.Field("WorkSpace").RemoveAll()
		Call Form.Fields.Field("WorkSpace").AddListItem("Please Select a WorkSpace", "Select")		
	

		
		SQL = "SELECT [PRJ_ID],[PRJ_PID],[PRJ_NAME],[PRJ_OWNER],[TYPE],[SUBTYPE] FROM [Active].[MHGROUP].[PROJECTS] WHERE [SUBTYPE] = 'work' and [PRJ_OWNER] = '" & FieldValue & "'"
	
		Set Connection = CreateObject("ADODB.Connection")
		Set Recordset = CreateObject("ADODB.Recordset")

		Connection.Open "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1"
		Recordset.Open SQL,Connection
		If Recordset.EOF Then
			RecordLabel = "No records found"
			RecordValue = "No records found"
			Call Form.Fields.Field("WorkSpace").AddListItem(RecordLabel, RecordValue)		
		Else
			Do While Not Recordset.Eof
				Dim workspaceID : workspaceID = Recordset("PRJ_ID")
				Dim workspaceName : workspaceName = Recordset("PRJ_NAME")
				RecordLabel = workspaceName
				RecordValue = workspaceID 
				Call Form.Fields.Field("WorkSpace").AddListItem(RecordLabel, RecordValue)
				Recordset.MoveNext
			Loop
		End If
		Recordset.Close
		Set Recordset=Nothing
		Connection.Close
		Set Connection=Nothing
	End If
	
		

	If FieldName = "WorkSpace" Then
		'Load WorkSpace List
	
		Call Form.Fields.Field("Folder").RemoveAll()
		Call Form.Fields.Field("Folder").AddListItem("Please Select a Folder", "Select")		
	
		SQL = "SELECT [PRJ_ID], [PRJ_NAME] FROM [Active].[MHGROUP].[PROJECTS] WHERE [PRJ_PID] ='" & FieldValue & "'"
	
		Set Connection = CreateObject("ADODB.Connection")
		Set Recordset = CreateObject("ADODB.Recordset")

		Connection.Open "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1"
		Recordset.Open SQL,Connection
		If Recordset.EOF Then
			RecordLabel = "No records found"
			RecordValue = "No records found"
			Call Form.Fields.Field("WorkSpace").AddListItem(RecordLabel, RecordValue)		
		Else
			Do While Not Recordset.Eof
				Dim folderID : folderID = Recordset("PRJ_ID")
				Dim folderName : folderName = Recordset("PRJ_NAME")
				RecordLabel = folderName
				RecordValue = folderID 
				Call Form.Fields.Field("Folder").AddListItem(RecordLabel, RecordValue)
				Recordset.MoveNext
			Loop
		End If
		Recordset.Close
		Set Recordset=Nothing
		Connection.Close
		Set Connection=Nothing	
		
	End If
End Sub


Function Field_OnValidate(FieldName, FieldValue)

End Function


Sub Button_OnClick(Form, ButtonName)

End Sub

