Option Strict Off

Imports System
Imports System.Data.ODBC
Imports System.IO
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports NSi.AutoStore.WebCapture.Workflow

Module Script

    Sub Form_OnLoad(ByVal eventData As ClientEventData)
   
   
		' Get List of Users for Operator and Author
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim operatorList As ListField = eventData.Form.Fields.GetField("Operator")
		Dim authorList As ListField = eventData.Form.Fields.GetField("Author")
		
		operatorList.Items.Clear
		authorList.Items.Clear
		
		cmdText	= "SELECT [USERID],[USERIDEX],[FULLNAME],[USERNUM] FROM [Active].[MHGROUP].[DOCUSERS] WHERE [LOGIN] = 'Y'"
	
		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				
				Dim listItem As listItem = New ListItem(dataReader.GetString(2), dataReader.GetString(0))
				operatorList.Items.Add(listItem)
				authorList.Items.Add(listItem)
				
			End While
			
		Catch ex As Exception
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()
		
		' Load the Tree!?
		
		Dim workspaceTree As TreeField = eventData.Form.Fields.GetField("WorkSpaces")
		AddWorkSpaces(workspaceTree.RootNode,workspaceTree.RootNode.Value)
	
	End Sub
	
	Sub AddWorkSpaces (ByVal workspaceTree As TreeItem, workSpaceID As String)
		
		'Get Top Level WorkSpace
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		Dim prjID As String
		Dim prjName As String
		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		if workSpaceID = "0" then
			cmdText	= "SELECT [PRJ_ID],[PRJ_PID],[PRJ_NAME] FROM [Active].[MHGROUP].[PROJECTS] WHERE PRJ_PID IS NULL"
		else
			cmdText	= "SELECT [PRJ_ID],[PRJ_PID],[PRJ_NAME] FROM [Active].[MHGROUP].[PROJECTS] WHERE PRJ_PID = '" & workSpaceID & "'"
		end if
		
		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				prjID = left(dataReader.GetString(0), instr(dataReader.GetString(0),".")-1)
				prjName = dataReader.GetString(2)
				Dim newNode As New TreeItem
				newNode.Text = prjName
				newNode.Value = prjID 
				workspaceTree.Nodes.Add(newNode)
				AddWorkSpaces(newNode,newNode.Value)
			End While
			
		Catch ex As Exception
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()
	End Sub
	
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
		'TODO: add code here to execute when the user presses OK in the form
	End Sub

	Sub Form_OnSubmit(ByVal eventData As ClientEventData)
		'TODO: add code here to execute after the sucessfull submitting of the form
    End Sub
End Module
