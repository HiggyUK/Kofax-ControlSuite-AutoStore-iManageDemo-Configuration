Sub GetClientandMatterOfWorkSpace_OnLoad
	
	SQL = "SELECT MHGROUP.PROJECTS.PRJ_ID, MHGROUP.PROJECTS.PRJ_NAME, MHGROUP.DOCMASTER.C1ALIAS, MHGROUP.DOCMASTER.C2ALIAS FROM MHGROUP.PROJECTS INNER JOIN MHGROUP.DOCMASTER ON MHGROUP.PROJECTS.DOCNUM = MHGROUP.DOCMASTER.DOCNUM AND MHGROUP.PROJECTS.VERSION = MHGROUP.DOCMASTER.VERSION WHERE  MHGROUP.PROJECTS.PRJ_ID = '" & prjID & "'"
	
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
			Dim clientID : clientID = Recordset("C1ALIAS")
			Dim matterID : matterID = Recordset("C2ALIAS")
			Recordset.MoveNext
		Loop
	End If
	Recordset.Close
	Set Recordset=Nothing
	Connection.Close
	Set Connection=Nothing	

	
	EKOManager.StatusMessage ("Client ID is " & clientID)
	EKOManager.StatusMessage ("Matter ID is " & matterID)
	
	
	Set Topic = KnowledgeContent.GetTopicInterface
	If Topic Is Nothing Then
		KnowledgeObject.Status = 2
		EKOManager.ErrorMessage "Cannot retrieve Topic Interface"
		Exit Sub
	End If
	
	Call Topic.Replace("~USR::ClientID~", clientID)
	Call Topic.Replace("~USR::MatterID~", matterID)
	

End Sub

Sub GetClientandMatterOfWorkSpace_OnUnload

End Sub
