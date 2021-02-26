Attribute VB_Name = "mdl_PickFileAccess"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' By..........: SILVA, ADILIO
' Contact.....: gomesadilio@outlook.com
' Date........: 1/1/2021
' Description.: Pick a folder
'---------------------------------------------------------------------------------------

'                         .
'                     /   ))     |\         )               ).
'               c--. (\  ( `.    / )  (\   ( `.     ).     ( (
'               | |   ))  ) )   ( (   `.`.  ) )    ( (      ) )
'               | |  ( ( / _..----.._  ) | ( ( _..----.._  ( (
' ,-.           | |---) V.'-------.. `-. )-/.-' ..------ `--) \._
' | /===========| |  (   |      ) ( ``-.`\/'.-''           (   ) ``-._
' | | / / / / / | |--------------------->  <-------------------------_>=-
' | \===========| |                 ..-'./\.`-..                _,,-'
' `-'           | |-------._------''_.-'----`-._``------_.-----'
'               | |         ``----''            ``----''
'               | |
'               c--`

Public Function PickFolderAccess() As String

	Dim objFd 			As FileDialog
    
    Set objFd = Application.FileDialog(msoFileDialogOpen)
    
    ChDrive "Z"
    ChDir "Z:\MyFolder\"
    
    With objFd
    
        .InitialFileName = ""
        .AllowMultiSelect = False
        .Filters.Clear        
        .Filters.Add "Txt", "*.txt", 1
        .Filters.Add "Csv", "*.csv", 2
        .FilterIndex = 1
        
        If .Show Then
            PickFolderAccess = .SelectedItems.Item(1)
        End If
        
    End With

End Sub
