Attribute VB_Name = "Winman"
Global TotalNumOfWinmanFiles As Long
Global WinmanFileArray() As String

Public Function CntWinmanFiles(myDirPath As String) As Long

    Dim fso As Scripting.FileSystemObject
    Dim Folder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim File As Scripting.File
    
    Set fso = New Scripting.FileSystemObject
    'Debug.Print myDirPath
    
    If fso.FolderExists(myDirPath) Then
        Set Folder = fso.GetFolder(myDirPath)
        
        If Folder.Attributes = 22 Or Folder.Attributes = 1046 And Folder.Attributes = 16 And Folder.Attributes = 18 Then
            Exit Function
        End If
        
        For Each File In Folder.Files
            If UCase(fso.GetExtensionName(File)) = "TAX" Then
            
                'Call SelectedFile(File.Name)
                
                If UBound(WinmanFileArray) <= 0 Then
                    TotalNumOfWinmanFiles = TotalNumOfWinmanFiles + 1
                Else
                    WinmanFileArray(TotalNumOfWinmanFiles) = File.Path
                    TotalNumOfWinmanFiles = TotalNumOfWinmanFiles + 1
                End If
                
            End If
        Next File
    
        For Each SubFolder In Folder.SubFolders
            CntWinmanFiles SubFolder.Path
        Next SubFolder
        
     Else
        MsgBox "Specified folder does not exist!", vbExclamation
    End If
    
    'MsgBox WinmanFileArray(UBound(WinmanFileArray))
    Set fso = Nothing
    Set Folder = Nothing
    
End Function

Public Function SelectedFile(FileName As String)
    Dim Cnt As Integer
    
    For Cnt = 0 To myForm.FilemyFiles.ListCount - 1
        If StrComp(myForm.FilemyFiles.List(Cnt), FileName, vbTextCompare) = 0 Then
            myForm.FilemyFiles.ListIndex = Cnt
            Exit For
        End If
    Next
End Function
