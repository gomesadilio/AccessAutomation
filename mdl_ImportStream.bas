Attribute VB_Name = "Módulo1"
Option Compare Database

Const table1 As String = "MYTABLE"
Const file As String = "C:\Users\MyUser\Desktop\MyTxt.txt"

Sub ImportTxtStreamMethod()

    Dim MyArray As Variant
    Dim fso As Variant
    Dim objStream As Variant
    Dim objFile As Variant
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Long
    
    i = 0
    
    Set rs = New ADODB.Recordset
    
    rs.Open "DELETE * FROM " & table1, CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    
    sSQL = "SELECT * FROM " & table1
    
    rs.Open sSQL, CurrentProject.Connection, adOpenDynamic, adLockOptimistic
     
     
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(file) Then
        Set objStream = fso.OpenTextFile(file, 1, False, 0)
    End If
    Do While Not objStream.AtEndOfStream
        strLine = objStream.ReadLine
           ReDim MyArray(0)
        MyArray = Split(strLine, vbTab)
        
        If i > 0 Then
         rs.AddNew
         For j = 0 To rs.Fields.Count - 1
         On Error Resume Next
         rs(j) = MyArray(j)
         On Error GoTo 0
         Next
         
         rs.Update
        End If
         i = i + 1
         
    Loop
     
    MsgBox i & " new lines inserter!"

End Sub


