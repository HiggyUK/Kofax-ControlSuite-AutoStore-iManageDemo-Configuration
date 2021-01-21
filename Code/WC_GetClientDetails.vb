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
		
		' Get List of Clients
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim clientList As ListField = eventData.Form.Fields.GetField("ClientID")
		
		clientList.Items.Clear
		
		
		cmdText	= "SELECT [CUSTOM_ALIAS],[C_DESCRIPT] FROM [MHGROUP].[CUSTOM1] WHERE [ENABLED] = 'Y'"
	
		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				
				Dim listItem As listItem = New ListItem(dataReader.GetString(1), dataReader.GetString(0))
				clientList.Items.Add(listItem)
				
			End While
			
		Catch ex As Exception
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()

		
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
		
	End Sub
	

	Sub ClientID_OnChange(ByVal eventData As ClientEventData)
		
		' Get List of Matters
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		Dim clientID As String = eventData.Form.Fields.GetField("ClientID").Value
		Dim matterID As String
		Dim matterName As String
		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim matterList As ListField = eventData.Form.Fields.GetField("MatterID")
		
		matterList.Items.Clear
		
		
		cmdText	= "SELECT [CUSTOM_ALIAS],[C_DESCRIPT] FROM [MHGROUP].[CUSTOM2] WHERE [ENABLED] = 'Y' AND [CPARENT_ALIAS] = '" & clientID & "'"


		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				matterID = dataReader.GetString(0)
				matterName = dataReader.GetString(1)
				Dim listItem As listItem = New ListItem(matterName, matterID)
				matterList.Items.Add(listItem)
				
			End While
			
		Catch ex As Exception
			
			Dim listItem As listItem = New ListItem(ex.Message, ex.Message)
			matterList.Items.Add(listItem)
		End Try
	
		If Not(dataReader Is Nothing) Then	
			dataReader.Close() 
		End If
		conn.Close()
		
		
	End Sub
	Sub MatterID_OnChange(ByVal eventData As ClientEventData)
		
		' Get List of WorkSpaces
		
		Dim connString As String
		Dim conn As OdbcConnection 
		Dim dataReader As OdbcDataReader 
		Dim command As OdbcCommand
		Dim cmdText As String
		Dim clientID As String = eventData.Form.Fields.GetField("ClientID").Value
		Dim matterID As String = eventData.Form.Fields.GetField("MatterID").Value
		Dim prjID As String
		Dim prjName As String
		
		connString = "DSN=IMDB;Database=Active;Uid=sa;Pwd=Password1" 
		conn = New OdbcConnection(connString)
		
		Dim workspaceList As ListField = eventData.Form.Fields.GetField("WorkSpaces")
		
		workspaceList.Items.Clear
		
		
		cmdText	= "SELECT MHGROUP.PROJECTS.PRJ_ID, MHGROUP.PROJECTS.PRJ_NAME, MHGROUP.DOCMASTER.C1ALIAS, MHGROUP.DOCMASTER.C2ALIAS FROM MHGROUP.PROJECTS INNER JOIN MHGROUP.DOCMASTER ON MHGROUP.PROJECTS.DOCNUM = MHGROUP.DOCMASTER.DOCNUM AND MHGROUP.PROJECTS.VERSION = MHGROUP.DOCMASTER.VERSION WHERE  MHGROUP.DOCMASTER.C1ALIAS = '" & clientID & "' AND MHGROUP.DOCMASTER.C2ALIAS = '" & matterID & "'"


		Try
			command = New OdbcCommand(cmdText) 
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader() 
			While dataReader.Read()
				prjID = left(dataReader.GetString(0), instr(dataReader.GetString(0),".")-1)
				prjName = dataReader.GetString(1)
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

	Sub WorkSpaces_OnChange(ByVal eventData As ClientEventData)
		
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
	Sub Form_OnValidate(ByVal eventData As ClientEventData)
      'TODO: add code here to execute when the user presses OK in the form
    End Sub

    Sub Form_OnSubmit(ByVal eventData As ClientEventData)
      'TODO: add code here to execute after the sucessfull submitting of the form
    End Sub
End Module
