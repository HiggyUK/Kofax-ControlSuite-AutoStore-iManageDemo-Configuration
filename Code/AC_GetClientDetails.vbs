Sub Form_OnLoad(Form)

	'Load Client List
	
	Call Form.Fields.Field("Client").RemoveAll()
	Call Form.Fields.Field("Client").AddListItem("Please Select a Client", "Select")		
	
	Dim RecordReturned
	Dim Connection, Recordset, SQL, Server, field, strAllFields
	Dim RecordLabel
	Dim RecordValue
		
	SQL = "SELECT [CUSTOM_ALIAS],[C_DESCRIPT] FROM [Active].[MHGROUP].[CUSTOM1] WHERE [ENABLED] = 'Y'"
	
	Set Connection = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")

	Connection.Open "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1"
	Recordset.Open SQL,Connection
	If Recordset.EOF Then
		RecordLabel = "No records found"
		RecordValue = "No records found"
		Call Form.Fields.Field("Client").AddListItem(RecordLabel, RecordValue)		
	Else
		Do While Not Recordset.Eof
			Dim clientID : clientID = Recordset("CUSTOM_ALIAS")
			Dim clientName : clientName = Recordset("C_DESCRIPT")
			RecordLabel = clientName
			RecordValue = clientID 
			Call Form.Fields.Field("Client").AddListItem(RecordLabel, RecordValue)
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
	
	If FieldName = "Client" Then
		
		'Load Matter List
	
		Call Form.Fields.Field("Matter").RemoveAll()
		Call Form.Fields.Field("Matter").AddListItem("Please Select a Matter", "Select")		
	

		
		SQL = "SELECT [CUSTOM_ALIAS],[C_DESCRIPT] FROM [Active].[MHGROUP].[CUSTOM2] WHERE [ENABLED] = 'Y' AND [CPARENT_ALIAS] = '" & FieldValue & "'"
	
		Set Connection = CreateObject("ADODB.Connection")
		Set Recordset = CreateObject("ADODB.Recordset")

		Connection.Open "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1"
		Recordset.Open SQL,Connection
		If Recordset.EOF Then
			RecordLabel = "No records found"
			RecordValue = "No records found"
			Call Form.Fields.Field("Matter").AddListItem(RecordLabel, RecordValue)		
		Else
			Do While Not Recordset.Eof
				Dim matterID : matterID = Recordset("CUSTOM_ALIAS")
				Dim matterName : matterName = Recordset("C_DESCRIPT")
				RecordLabel = matterName
				RecordValue = matterID 
				Call Form.Fields.Field("Matter").AddListItem(RecordLabel, RecordValue)
				Recordset.MoveNext
			Loop
		End If
		Recordset.Close
		Set Recordset=Nothing
		Connection.Close
		Set Connection=Nothing
	End If
	
		
	If FieldName = "Matter" Then
			'Load WorkSpace List
	
		Call Form.Fields.Field("WorkSpace").RemoveAll()
		Call Form.Fields.Field("WorkSpace").AddListItem("Please Select a WorkSpace", "Select")		
	

		Dim clientID : clientID = Form.GetFieldValue("Client")
	
		SQL = "SELECT MHGROUP.PROJECTS.PRJ_ID, MHGROUP.PROJECTS.PRJ_NAME, MHGROUP.DOCMASTER.C1ALIAS, MHGROUP.DOCMASTER.C2ALIAS FROM MHGROUP.PROJECTS INNER JOIN MHGROUP.DOCMASTER ON MHGROUP.PROJECTS.DOCNUM = MHGROUP.DOCMASTER.DOCNUM AND MHGROUP.PROJECTS.VERSION = MHGROUP.DOCMASTER.VERSION WHERE  MHGROUP.DOCMASTER.C1ALIAS = '" & clientID & "' AND MHGROUP.DOCMASTER.C2ALIAS = '" & FieldValue & "'"
	
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
				Dim projectID : projectID = Recordset("PRJ_ID")
				Dim projectName : projectName = Recordset("PRJ_NAME")
				RecordLabel = projectName
				RecordValue = projectID 
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

