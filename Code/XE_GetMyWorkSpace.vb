'LoadAssembly:System.Data.dll

Option Strict Off

Imports System
Imports System.Data.ODBC
Imports System.IO
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports XeroxEIPCapture.Workflow



Module Script
    Sub Form_OnLoad(ByVal eventData As MFPEventData)
		' Get List of Users
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim userList As ListField = eventData.Form.Fields.GetField("UserID")
		Dim operatorList As ListField = eventData.Form.Fields.GetField("Operator")
		
		userList.Items.Clear
		operatorList.Items.Clear
		
		
		cmdText	= "SELECT [USERID],[USERIDEX],[FULLNAME],[USERNUM] FROM [Active].[MHGROUP].[DOCUSERS] WHERE [LOGIN] = 'Y'"
	
		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				
				Dim listItem As listItem = New ListItem(dataReader.GetString(2), dataReader.GetString(0))
				userList.Items.Add(listItem)
				operatorList.Items.Add(listItem)
				
			End While
			
		Catch ex As Exception
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()
	End Sub
	
	Sub UserID_OnChange(ByVal eventData As MFPEventData)
		
		' Get List of WorkSpaces
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		Dim userID As String = eventData.Form.Fields.GetField("UserID").Value
		Dim prjID As String
		Dim prjName As String
		

		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim workspaceList As ListField = eventData.Form.Fields.GetField("WorkSpaces")
		
		workspaceList.Items.Clear
		
		
		cmdText	= "SELECT [PRJ_ID],[PRJ_PID],[PRJ_NAME],[PRJ_OWNER],[TYPE],[SUBTYPE] FROM [Active].[MHGROUP].[PROJECTS] WHERE [SUBTYPE] = 'work' and [PRJ_OWNER] = '" & userID & "'"


		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				prjID = left(dataReader.GetString(0), instr(dataReader.GetString(0),".")-1)
				prjName = dataReader.GetString(2)
				Dim listItem As listItem = New ListItem(prjName, prjID)
				workspaceList.Items.Add(listItem)
				
			End While
			
		Catch ex As Exception
			
			Dim listItem As listItem = New ListItem(ex.Message, ex.Message)
			workspaceList.Items.Add(listItem)
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()
		
		
	End Sub

	Sub WorkSpaces_OnChange(ByVal eventData As MFPEventData)
		
		' Get List of Folders
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		Dim prjID As String = eventData.Form.Fields.GetField("WorkSpaces").Value
		Dim folderID As String
		Dim folderName As String
		

		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim folderList As ListField = eventData.Form.Fields.GetField("Folder")
		
		folderList.Items.Clear
		
		
		cmdText	= "SELECT [PRJ_ID], [PRJ_NAME] FROM [Active].[MHGROUP].[PROJECTS] WHERE [PRJ_PID] ='" & prjID & "'"

		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				folderID = left(dataReader.GetString(0), instr(dataReader.GetString(0),".")-1)
				folderName = dataReader.GetString(1)
				
				Dim listItem As listItem = New ListItem(folderName, folderID)
				folderList.Items.Add(listItem)
				
			End While
			
		Catch ex As Exception
		
			Dim listItem As listItem = New ListItem(ex.Message, ex.Message)
			folderList.Items.Add(listItem)
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()
		
		
	End Sub
	
    Sub Form_OnSubmit(ByVal eventData As MFPEventData)
        'TODO add code here to execute when the user presses OK in the form
    End Sub
End Module
