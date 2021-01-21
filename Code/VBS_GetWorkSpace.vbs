Sub GetWorkSpace_OnLoad
	
	Dim parentID
	
	Call GetParentWorkSpace(prjID, parentID)

	Set Topic = KnowledgeContent.GetTopicInterface
	If Topic Is Nothing Then
		KnowledgeObject.Status = 2
		EKOManager.ErrorMessage "Cannot retrieve Topic Interface"
		Exit Sub
	End If
	
	Call Topic.Replace("~USR::ParentID~", CStr(prjID))

End Sub

Sub GetParentWorkSpace(prjID, parentID)
	
	
	Dim newparentID : newparentID = Null

	EKOManager.StatusMessage "Get Parent WorkSpace of " & prjID
	
	SQL = "SELECT PRJ_ID, PRJ_PID, PRJ_NAME FROM MHGROUP.PROJECTS WHERE PRJ_ID = '" & prjID & "'"
	
	Set Connection = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")

	Connection.Open "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1"
	Recordset.Open SQL,Connection
	If Not Recordset.EOF Then
		Do While Not Recordset.Eof
			newparentID = Recordset("PRJ_PID")
			Recordset.MoveNext
		Loop
	End If
	Recordset.Close
	Set Recordset=Nothing
	Connection.Close
	Set Connection=Nothing	
	
	EKOManager.StatusMessage "Prj_ID: " & prjID
	EKOManager.StatusMessage "Prj_PID: " & newparentID
	
	If newparentID <> "" Then 
		EKOManager.StatusMessage "Parent PID is not null"
		prjID = newparentID
		parentID = ""
		Call GetParentWorkSpace(prjID, parentID)
	Else
		parentID = newparentID
	End If

End Sub

Sub GetClientandMatterOfWorkSpace_OnUnload

End Sub
