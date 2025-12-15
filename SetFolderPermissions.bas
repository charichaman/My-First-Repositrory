Attribute VB_Name = "FolderPermission"
Private Declare Function SetNamedSecurityInfo Lib "advapi32.dll" Alias "SetNamedSecurityInfoA" ( _
    ByVal pObjectName As String, _
    ByVal ObjectType As Long, _
    ByVal SecurityInfo As Long, _
    ByVal pOwner As Any, _
    ByVal pGroup As Any, _
    ByVal pDacl As Any, _
    ByVal pSacl As Any) As Long

Private Const SE_FILE_OBJECT As Long = 1
Private Const DACL_SECURITY_INFORMATION As Long = &H4
Private Const ERROR_SUCCESS As Long = 0

Public Function ResetFolderPermissions(ByVal FolderPath As String) As Boolean
    Dim Result As Long

    ' Set the DACL to NULL, which removes all permissions and users
    Result = SetNamedSecurityInfo(FolderPath, SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)

    Shell "icacls """ & FolderPath & """ /grant " & "Everyone" & ":(F) /T"
    
End Function
