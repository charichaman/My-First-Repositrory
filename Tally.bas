Attribute VB_Name = "Tally"
Global Paths() As String
Global AdminInfo As String

Global TotalNumOfTallyDataFolders As Long
Global ProgressCurrentValue As Integer

Global TallyFolderPaths() As String
Global TallyFolderPathsArraySize As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub LoadFolders(myPath As String, myDir As DirListBox)
    myDir.Path = myPath
End Sub

Public Sub InitializeAll(myForm As Form)

    AdminInfo = "Nagaraj HD"
    
    Call CenterMe(myForm)
    Call CenterMe(frmPleaseWait)
    
    Call LoadFolders(Mid(myForm.DrvMyDrives, 1, 2) & "\", myForm.DirMyFolders)
    Call LoadUnLoadConfig("Read")
    
End Sub

Public Function LoadUnLoadConfig(ReadOrWrite As String)

    Dim fso As Scripting.FileSystemObject
    Dim myConfigTxtFile As Scripting.TextStream
    Dim sFolderPath As String
    Dim dFolderPath As String
    Dim TallyERP9ExePath As String
    Dim TallyPrimeExePath As String
    
    Set fso = New Scripting.FileSystemObject
    
    If Not fso.FileExists(App.Path & "\myConfigFile.txt") Then
        Set myConfigTxtFile = fso.CreateTextFile(App.Path & "\myConfigFile.txt")
            myConfigTxtFile.WriteLine "C:\"
            myConfigTxtFile.WriteLine "D:\"
            myConfigTxtFile.WriteLine "Click Me to Get Tally.ERP9 \ Tally.Exe"
            myConfigTxtFile.WriteLine "Click Me to Get TallyPrime \ Tally.Exe"
        myConfigTxtFile.Close
    Else
        If ReadOrWrite = "Read" Then
            Set myConfigTxtFile = fso.OpenTextFile(App.Path & "\myConfigFile.txt", ForReading, False, TristateMixed)
            
                sFolderPath = myConfigTxtFile.ReadLine
                    If fso.FolderExists(sFolderPath) Then
                        myForm.DrvMyDrives.Drive = Mid(sFolderPath, 1, 2)
                            If fso.FolderExists(sFolderPath) Then
                                myForm.DirMyFolders.Path = sFolderPath
                            End If
                    Else
                        myForm.DrvMyDrives.Drive = "C:\"
                        myForm.DirMyFolders.Path = "C:\"
                    End If
                        
                dFolderPath = myConfigTxtFile.ReadLine
                    If fso.FolderExists(dFolderPath) Then
                        myForm.LblClientsFolderPath = dFolderPath
                    End If
                    
                TallyERP9ExePath = myConfigTxtFile.ReadLine
                    If fso.FileExists(TallyERP9ExePath) Then
                        myForm.LblTallyERP9Path = TallyERP9ExePath
                    Else
                        myForm.LblTallyERP9Path = "Click Me to Get Tally.ERP9 \ Tally.Exe"
                    End If
                
                TallyPrimeExePath = myConfigTxtFile.ReadLine
                    If fso.FileExists(TallyPrimeExePath) Then
                        myForm.LblTallyPrimePath = TallyPrimeExePath
                    Else
                        myForm.LblTallyPrimePath = "Click Me to Get TallyPrime \ Tally.Exe"
                    End If
                
            myConfigTxtFile.Close
            
            Call tdlCompanyCreateTxtForTallyERP9
            Call tdlCompanyCreateTxtForTallyPrime
            
        ElseIf ReadOrWrite = "Write" Then
            Set myConfigTxtFile = fso.OpenTextFile(App.Path & "\myConfigFile.txt", ForWriting, False)
                myConfigTxtFile.WriteLine myForm.DirMyFolders.Path
                myConfigTxtFile.WriteLine myForm.LblClientsFolderPath.Caption
                myConfigTxtFile.WriteLine myForm.LblTallyERP9Path.Caption
                myConfigTxtFile.WriteLine myForm.LblTallyPrimePath.Caption
            myConfigTxtFile.Close
        End If
    End If
    
    Set fso = Nothing
    Set myConfigTxtFile = Nothing
    
        
    
End Function

Public Function CenterMe(ByVal myForm As Form)
    
    'MyForm.WindowState = 0
    
    myForm.Left = (Screen.Width - myForm.Width) / 2
    myForm.Top = (Screen.Height - myForm.Height) / 2

End Function

Public Function CntTallyDataFoders(myDirPath As String) As Long

    Dim fso As Scripting.FileSystemObject
    Dim folder As Scripting.folder
    Dim subFolder As Scripting.folder
    Dim myPathAttribute As Integer
    Dim TallyERP9orTallyPrime As String
    
    Set fso = New Scripting.FileSystemObject
    'Debug.Print myDirPath
    
    If fso.FolderExists(myDirPath) Then
        Set folder = fso.GetFolder(myDirPath)
        
        If folder.Attributes = 22 Or folder.Attributes = 1046 And folder.Attributes = 16 And folder.Attributes = 18 Then
            Exit Function
        End If
        
        For Each subFolder In folder.SubFolders
            If IsNumeric(fso.GetFolder(subFolder).Name) Then
            
                TallyERP9orTallyPrime = CheckForTallyFiles(subFolder)
                If TallyERP9orTallyPrime = "TallyERP9" Or TallyERP9orTallyPrime = "TallyPrime" Then
                
                        ReDim Preserve TallyFolderPaths(TotalNumOfTallyDataFolders)
                            TallyFolderPaths(TotalNumOfTallyDataFolders) = subFolder.Path
                    
                        TotalNumOfTallyDataFolders = TotalNumOfTallyDataFolders + 1
                        
                            If fso.FileExists(subFolder & "\Company.999") = True Then
'Debug.Print subFolder
                                fso.DeleteFile subFolder & "\Company.999"
                            End If
                            If fso.FileExists(subFolder & "\Company.888") = True Then
                                fso.DeleteFile subFolder & "\Company.888"
                            End If
                    End If
                End If
                CntTallyDataFoders subFolder.Path
        Next subFolder

    Else
        MsgBox "Specified folder does not exist!", vbExclamation
    End If

    ProgressCurrentValue = 0
    CntTallyDataFoders = TotalNumOfTallyDataFolders

            Set fso = Nothing
            Set folder = Nothing
            Set subFolder = Nothing

End Function

Public Function myMainPathToCreateCompany999File(DirPath As String)
    
    Dim fso As Scripting.FileSystemObject
        Dim sFolder As Scripting.folder
        Dim dFolder As Scripting.folder
        Dim FolderPath As Scripting.folder
    Dim Cnt As Long
    Dim StopExecuting As Variant
   
    Dim Company999888Created As Boolean
    Dim CompaniesSuccessFullyTransfered As Integer
    Dim CompaniesLocked As Integer
    
    CompaniesSuccessFullyTransfered = 0
    CompaniesLocked = 0
    
    Set fso = New Scripting.FileSystemObject
    
'fso.MoveFolder "G:\01. WD Cloud Data\Boss-Sys-Bkup\Old\Boss-System-Backup\D\Boss", "G:\01. WD Cloud Data\Boss-Sys-Bkup\Old\Boss-System-Backup\D\Boss-1"

    TotalNumOfTallyDataFolders = 0
    TotalNumOfTallyDataFolders = CntTallyDataFoders(DirPath)

    If TotalNumOfTallyDataFolders > 0 Then
            frmPleaseWait.Show vbModeless
            
            'frmPleaseWait.myProgressBar.
            
            frmPleaseWait.myProgressBar.Max = 1
            frmPleaseWait.myProgressBar.Max = TotalNumOfTallyDataFolders
            
        For Cnt = LBound(TallyFolderPaths) To UBound(TallyFolderPaths)

'myForm.LblCurrentFolderName = Dir(TallyFolderPaths(Cnt), vbDirectory)
myForm.LblFoldersTotal.Caption = TotalNumOfTallyDataFolders
myForm.LblFoldersRemains.Caption = TotalNumOfTallyDataFolders - (CompaniesSuccessFullyTransfered + CompaniesLocked)
    myForm.LblCurrentWorkingFolder.Caption = TallyFolderPaths(Cnt)
            
            myForm.Enabled = False
                frmPleaseWait.myProgressBar.Value = Cnt
                
                DoEvents
                        myForm.LblFolderStatus.ForeColor = &HFFFF&
                        myForm.LblFolderStatus.BackColor = &H80FF80
                        myForm.LblFolderStatus.Caption = ""

Debug.Print TallyFolderPaths(Cnt)

                If fso.FolderExists(TallyFolderPaths(Cnt)) = True Then
                
                    Set FolderPath = fso.GetFolder(TallyFolderPaths(Cnt))

                    If Company999888txtFileCreate(FolderPath) = True Then
                        CompaniesSuccessFullyTransfered = CompaniesSuccessFullyTransfered + 1
                        Set sFolder = FolderPath
                        Set dFolder = fso.GetFolder(myForm.LblClientsFolderPath.Caption)
'Debug.Print sFolder
'Debug.Print dFolder & vbCr
                        myForm.Enabled = True
                            myForm.LblFolderStatus.ForeColor = &HFFFF&
                            myForm.LblFolderStatus.BackColor = &H80FF80
                            myForm.LblFolderStatus.Caption = "Open"
                                Call FillCompanyDataToControls(sFolder)
                        myForm.Enabled = False
                        
                        Call MoveTallyDataFolderToAbhinandan(sFolder, dFolder)
                    Else
                        CompaniesLocked = CompaniesLocked + 1
                        Set sFolder = FolderPath
                        Set dFolder = fso.GetFolder(myForm.LblClientsFolderPath.Caption)
                        
                        myForm.Enabled = True
                            myForm.LblFolderStatus.ForeColor = &H80FF80
                            myForm.LblFolderStatus.BackColor = &HFF&
                            myForm.LblFolderStatus.Caption = "Locked"
                        myForm.Enabled = False
                        
                        Call MoveToProtectedFolder(sFolder, dFolder)
                    End If
                    
                End If
        Next
        
        Unload frmPleaseWait
        myForm.Enabled = True

        myForm.LblFoldersTotal.Caption = TotalNumOfTallyDataFolders
        myForm.LblFoldersRemains.Caption = TotalNumOfTallyDataFolders - CompaniesSuccessFullyTransfered

        MsgBox "Tally's  : " & CompaniesSuccessFullyTransfered & " : Data Folder(s) Transfered Successfully...!", vbOKOnly + vbInformation, AdminInfo
    Else
        MsgBox "Tally Data Folders may not be found...!", vbOKOnly + vbInformation, AdminInfo
    End If
    
    'Set fso = Nothing
    'Set sFolder = Nothing
    'Set dFolder = Nothing
    

    Set sFolder = fso.GetFolder(myForm.DirMyFolders.List(myForm.DirMyFolders.ListIndex))
    OkFolderPath = myForm.DirMyFolders.List(myForm.DirMyFolders.ListIndex)
    
    If Not Mid(OkFolderPath, Len(OkFolderPath) - 3, 4) = "(Ok)" Then
        OkFolderPath = OkFolderPath & " (Ok)"

Debug.Print sFolder
Debug.Print OkFolderPath

'MsgBox sFolder & " ---> " & OkFolderPath

            fso.MoveFolder sFolder, OkFolderPath
        myForm.DirMyFolders.Path = OkFolderPath
    End If
    
End Function
Public Function MoveToProtectedFolder(mYsFolder As Scripting.folder, mYdFolder As Scripting.folder)

    Dim fso As Scripting.FileSystemObject
    Dim sFolder As Scripting.folder
        Dim DataFolder As String
    Dim dFolder As Scripting.folder
    
    Dim retryCount As Integer
    Dim maxRetries As Integer
    
    retryCount = 0
    maxRetries = 5
    
    Set fso = New Scripting.FileSystemObject

'On Error GoTo myWait

        Set sFolder = mYsFolder
            DataFolder = Trim(fso.GetBaseName(sFolder))
        Set dFolder = mYdFolder
        
            If Not fso.FolderExists(dFolder & "\Protected") Then
                fso.CreateFolder (dFolder & "\Protected")
                Set dFolder = fso.GetFolder(dFolder & "\Protected")
            Else
                Set dFolder = fso.GetFolder(dFolder & "\Protected")
            End If
                            
CreateFolder:
    If fso.FolderExists(dFolder & "\" & DataFolder) Then
        DataFolderNumberFormat = String(Len(DataFolder), "0")
        DataFolder = Format(Val(DataFolder) + 1, DataFolderNumberFormat)
        GoTo CreateFolder:
    Else

Debug.Print
Debug.Print "Pro-S-Folder : " & sFolder
Debug.Print "Pro-D-Folder : " & dFolder & "\" & DataFolder
Debug.Print
            
    retryCount = 0
    
        Do While retryCount < maxRetries
            On Error Resume Next ' Ignore errors to retry
            fso.CopyFolder sFolder, dFolder & "\" & DataFolder
            If Err.Number = 0 Then
                ' Successfully copied the folder, now delete the original folder
                fso.DeleteFolder mYsFolder
                Exit Do
            ElseIf Err.Number = 70 Then
                ' Permission Denied error (Error 70)
                'MsgBox "Permission Denied, retrying... Attempt " & (retryCount + 1) & " of " & maxRetries
                retryCount = retryCount + 1
                Sleep 1000 ' Wait for 1 second before retrying
            Else
                ' Handle any other errors
                MsgBox "An unexpected error occurred: " & Err.Description
                Exit Do
            End If
        Loop

    End If

    myForm.DirClientsFolders.Refresh
    myForm.DirMyFolders.Refresh
        
    Call LoadToListBox(myForm)
    
End Function
Public Function Company999888txtFileCreate(myFolderPath As Scripting.folder) As Boolean

    Dim fso As Scripting.FileSystemObject
    Dim TmpTxtFile As Scripting.TextStream
    
    Dim DataFolderPath As String
    Dim DataFolderName As String
    Dim CompanyTxtFileString As String
    
        Dim RetnValue As Long
        Dim RetnMaxValue As Long
        Dim SleepValue As Integer
    
    Dim TallyERP9orTallyPrime As String
    
    Set fso = New Scripting.FileSystemObject
    
Debug.Print DataFolderPath
Debug.Print DataFolderName

        DataFolderPath = fso.GetParentFolderName(myFolderPath)
        DataFolderName = fso.GetBaseName(myFolderPath)
        
        If Not fso.FolderExists("C:\DATA") Then
            fso.CreateFolder "C:\DATA"
        End If
        
        If fso.FileExists("C:\DATA\Company.999") Then
                fso.DeleteFile "C:\DATA\Company.999"
        End If
        
        If fso.FileExists("C:\DATA\Company.888") Then
            fso.DeleteFile "C:\DATA\Company.888"
        End If
        
        'MsgBox myFolderPath
        
        TallyERP9orTallyPrime = CheckForTallyFiles(myFolderPath)
        
        If TallyERP9orTallyPrime = "TallyERP9" Then
            CompanyTxtFileString = myForm.LblTallyERP9Path.Caption & " /NOINILOAD /NOINITDL /TDL:" & App.Path & "\tdlCompanyCreateTxtForTallyERP9.txt /NOGUI /DATA:""" & DataFolderPath & """ /LOAD:" & DataFolderName & " /ACTION:CALL:CreateTxtFile"
            Shell CompanyTxtFileString, vbNormalFocus
            
                SleepValue = 6000
Wait999:
                    Sleep (SleepValue)
                        SleepValue = SleepValue - 500
                        
                        If SleepValue <= 0 Then
                            Set fso = Nothing
                            Exit Function
                        End If
                        
                        If fso.FileExists("C:\DATA\Company.999") Then
                            On Error Resume Next
                                Set TmpTxtFile = fso.OpenTextFile("C:\DATA\Company.999", ForReading, False, TristateMixed)
                                    If Err.Number <> 0 Then
                                        GoTo Wait999:
                                    End If
                                Set TmpTxtFile = Nothing
                            fso.MoveFile "C:\DATA\Company.999", myFolderPath & "\Company.999"
                                Company999888txtFileCreate = True
                            'End
                        Else
                            GoTo Wait999:
                        End If

        ElseIf TallyERP9orTallyPrime = "TallyPrime" Then
            CompanyTxtFileString = myForm.LblTallyPrimePath.Caption & " /NOINILOAD /NOINITDL /TDL:" & App.Path & "\tdlCompanyCreateTxtForTallyPrime.txt /NOGUI /DATA:""" & DataFolderPath & """ /LOAD:" & DataFolderName & " /ACTION:CALL:CreateTxtFile"

'Debug.Print CompanyTxtFileString
            Shell CompanyTxtFileString, vbNormalFocus
            
                SleepValue = 6000
                    
Wait888:
                    Sleep (SleepValue)
                        SleepValue = SleepValue - 500
                        
                        If SleepValue <= 0 Then
                            Set fso = Nothing
                            Exit Function
                        End If
                        
                        If fso.FileExists("C:\DATA\Company.888") Then
                            On Error Resume Next
                                Set TmpTxtFile = fso.OpenTextFile("C:\DATA\Company.888", ForReading, False, TristateMixed)
                                    If Err.Number <> 0 Then
                                        GoTo Wait888:
                                    End If
                                Set TmpTxtFile = Nothing
                            fso.MoveFile "C:\DATA\Company.888", myFolderPath & "\Company.888"
                                Company999888txtFileCreate = True
                            'End
                        Else
                            GoTo Wait888:
                        End If
            
        End If
            
    Set fso = Nothing
    
End Function

Public Function MoveTallyDataFolderToAbhinandan(mYsFolder As Scripting.folder, mYdFolder As Scripting.folder)
    Dim fso As Scripting.FileSystemObject
    Dim Cmp999888File As Scripting.TextStream
    
        Dim CmpName As String
        Dim CmpGST As String
        Dim CmpPAN As String
        Dim CmpFromDate As String
        Dim CmpToDate As String
        
        Dim CmpGrossProfit As String
        Dim CmpClosingStock As String
        Dim CmpSalesAccount As String
        Dim CmpDirectIncome As String
        Dim CmpNetProfit As String
        Dim CmpTransfered As String
        
        Dim myCompanyFinYear As String
        
        Dim DataFolder As String
        
        Dim CmpDetailsTxtFileName As String
        
        Dim DataFolderNumberFormat As String

        Set fso = New Scripting.FileSystemObject
        
        If fso.FileExists(mYsFolder & "\Company.999") Then
            CmpDetailsTxtFileName = mYsFolder & "\Company.999"
        ElseIf fso.FileExists(mYsFolder & "\Company.888") Then
            CmpDetailsTxtFileName = mYsFolder & "\Company.888"
        End If
        

            Set Cmp999888File = fso.OpenTextFile(CmpDetailsTxtFileName, ForReading, False, TristateMixed)
                CmpName = Cmp999888File.ReadLine
                CmpGST = Cmp999888File.ReadLine
                CmpPAN = Cmp999888File.ReadLine
                CmpFromDate = Cmp999888File.ReadLine
                CmpToDate = Cmp999888File.ReadLine
                    Cmp999888File.ReadLine
                CmpGrossProfit = Cmp999888File.ReadLine
                CmpClosingStock = Cmp999888File.ReadLine
                CmpSalesAccount = Cmp999888File.ReadLine
                CmpDirectIncome = Cmp999888File.ReadLine
                    Cmp999888File.ReadLine
                CmpNetProfit = Cmp999888File.ReadLine
                CmpTransfered = Cmp999888File.ReadLine
            Set Cmp999888File = Nothing
            
                    CmpName = Trim(Mid(CmpName, InStrRev(CmpName, "!") + 1))
                        CmpName = MakeCompanyNameProper(CmpName)
                    
                    CmpGST = Trim(Mid(CmpGST, InStrRev(CmpGST, "!") + 1))
                    CmpPAN = Trim(Mid(CmpPAN, InStrRev(CmpPAN, "!") + 1))
                    
                    CmpFromDate = Trim(Mid(CmpFromDate, InStrRev(CmpFromDate, "!") + 1))
                        CmpFromDate = Format(CmpFromDate, "dd-mmm-YYYY")
                    
                    CmpToDate = Trim(Mid(CmpToDate, InStrRev(CmpToDate, "!") + 1))
                        CmpToDate = Format(CmpToDate, "dd-mmm-YYYY")
                    
                    myCompanyFinYear = Format(CmpFromDate, "YYYY") & "-" & Val(Format(CmpFromDate, "YY")) + 1
                    
                        CmpGrossProfit = Trim(Mid(CmpGrossProfit, InStrRev(CmpGrossProfit, "!") + 1))
                        CmpClosingStock = Trim(Mid(CmpClosingStock, InStrRev(CmpClosingStock, "!") + 1))
                        CmpSalesAccount = Trim(Mid(CmpSalesAccount, InStrRev(CmpSalesAccount, "!") + 1))
                        CmpDirectIncome = Trim(Mid(CmpDirectIncome, InStrRev(CmpDirectIncome, "!") + 1))
                    
                    CmpNetProfit = Trim(Mid(CmpNetProfit, InStrRev(CmpNetProfit, "!") + 1))
                    CmpTransfered = Trim(Mid(CmpTransfered, InStrRev(CmpTransfered, "!") + 1))
 Debug.Print CmpName
 Debug.Print CmpFromDate
 
            If Len(CmpName) <> 0 Then
                If Not fso.FolderExists(mYdFolder & "\" & CmpName) Then
                    Set mYdFolder = fso.CreateFolder(mYdFolder & "\" & CmpName)
                Else
                    Set mYdFolder = fso.GetFolder(mYdFolder & "\" & CmpName)
                End If
                    If Not fso.FolderExists(mYdFolder & "\" & myCompanyFinYear) Then
                        Set mYdFolder = fso.CreateFolder(mYdFolder & "\" & myCompanyFinYear)
                    Else
                        Set mYdFolder = fso.GetFolder(mYdFolder & "\" & myCompanyFinYear)
                    End If
                    
                    DataFolder = fso.GetBaseName(mYsFolder)
                    
CreateFolder:
                        If fso.FolderExists(mYdFolder & "\" & DataFolder) Then
                            DataFolderNumberFormat = String(Len(DataFolder), "0")
                            DataFolder = Format(Val(DataFolder) + 1, DataFolderNumberFormat)
                            GoTo CreateFolder:
                        Else

Debug.Print
Debug.Print "Source Folder : " & mYsFolder
Debug.Print "Destination Folder : " & mYdFolder & "\" & DataFolder
Debug.Print

                            fso.CopyFolder mYsFolder, mYdFolder & "\" & DataFolder
                                fso.DeleteFolder mYsFolder
                                
                            myForm.DirClientsFolders.Refresh
                            myForm.DirMyFolders.Refresh
                            
                            Call LoadToListBox(myForm)
                            
                        End If
            End If
                
        
        Set fso = Nothing
        
End Function

Public Function MakeCompanyNameProper(myCmpName As String) As String

    Dim Cnt As Integer
    Dim TmpCmpName As String
    Dim myChar As String
    'TmpCmpName = myCmpName
    
    For Cnt = 1 To Len(myCmpName)
            myChar = Mid(myCmpName, Cnt, 1)
        'If myChar Like "[A-Za-z0-9. ]" Then
        If myChar Like "[A-Za-z. ]" Then
            TmpCmpName = Mid(myCmpName, 1, Cnt)
        Else
            MakeCompanyNameProper = Trim(TmpCmpName)
            Exit Function
        End If
    Next
    
    TmpCmpName = Replace(TmpCmpName, ".", Space(1))
    MakeCompanyNameProper = Trim(TmpCmpName)
    
End Function

Public Function CheckForTallyFiles(myPathLine As Scripting.folder) As String
    
    Dim fso As Scripting.FileSystemObject
    Dim textFile As Scripting.TextStream
    
    Dim CntFiles As Integer
    Dim TallyERP9Files As Variant
    Dim TallyPrimeFiles As Variant
    Dim IsTallyERP9 As Boolean
    Dim IsTallyPrime As Boolean
    
        'TallyERP9Files = Array("AddlCmp.900", "CmpSave.900", "Company.900", "LinkMgr.900", "Manager.900", "SumTran.900", "TACCESS.TSF", "TEXCL.TSF", "TranMgr.900", "TSTATE.TSF", "TUPDATE.TSF")
        'TallyERP9Files = Array("AddlCmp.900", "CmpSave.900", "Company.900", "LinkMgr.900", "Manager.900", "SumTran.900", "TACCESS.TSF", "TEXCL.TSF", "TranMgr.900", "TSTATE.TSF")
        'TallyERP9Files = Array("AddlCmp.900", "CmpSave.900", "Company.900", "LinkMgr.900", "Manager.900", "SumTran.900", "TACCESS.TSF", "TranMgr.900", "TSTATE.TSF")
        TallyERP9Files = Array("AddlCmp.900", "CmpSave.900", "Company.900", "LinkMgr.900", "Manager.900", "TACCESS.TSF", "TranMgr.900", "TSTATE.TSF")
    
        'TallyPrimeFiles = Array("AddlCmp.1800", "Aggr.1800", "CmpSave.1800", "Company.1800", "ExtMngr.1800", "LinkMgr.1800", "Manager.1800", "SecTran.1800", "StatStatus.1800", "TACCESS.TSF", "TEXCL.TSF", "TranMgr.1800", "TSTATE.TSF", "TUPDATE.TSF", "VchStatus.1800")
        TallyPrimeFiles = Array("AddlCmp.1800", "Aggr.1800", "CmpSave.1800", "Company.1800", "ExtMngr.1800", "LinkMgr.1800", "Manager.1800", "SecTran.1800", "StatStatus.1800", "TACCESS.TSF", "TEXCL.TSF", "TranMgr.1800", "TSTATE.TSF", "VchStatus.1800")
        
        
    Set fso = New Scripting.FileSystemObject
    
    
Debug.Print myPathLine
    
        If fso.GetExtensionName(Dir(myPathLine & "\TranMgr.*")) = "900" Then
            IsTallyERP9 = True
        ElseIf fso.GetExtensionName(Dir(myPathLine & "\TranMgr.*")) = "1800" Then
            IsTallyPrime = True
        End If
        
        If IsTallyERP9 = True Then
            For CntFiles = LBound(TallyERP9Files) To UBound(TallyERP9Files)
                If fso.FileExists(myPathLine & "\" & TallyERP9Files(CntFiles)) Then
                    CheckForTallyFiles = "TallyERP9"
                Else
                    CheckForTallyFiles = ""
                    Exit For
                End If
            Next
        End If
        
        If IsTallyPrime = True Then
            For CntFiles = LBound(TallyPrimeFiles) To UBound(TallyPrimeFiles)
                If fso.FileExists(myPathLine & "\" & TallyPrimeFiles(CntFiles)) Then
                    CheckForTallyFiles = "TallyPrime"
                Else
                    CheckForTallyFiles = ""
                    Exit For
                End If
            Next
        End If
    
    
    Set fso = Nothing
    
End Function

Public Function LoadToListBox(myMainForm As Form)

    Dim fso As Scripting.FileSystemObject
    Dim folder As Scripting.folder
    Dim subFolder As Scripting.folder

    Set fso = New Scripting.FileSystemObject
    
    myMainForm.LstClients.Clear
    
    If myMainForm.DirClientsFolders.ListCount > 0 Then
        Set folder = fso.GetFolder(myMainForm.DirClientsFolders.Path)
            For Each subFolder In folder.SubFolders
                myMainForm.LstClients.AddItem subFolder.Name
            Next subFolder
    End If

End Function

Public Function FillCompanyDataToControls(DataFolder As Scripting.folder)

    Dim fso As Scripting.FileSystemObject
        Dim folder As Scripting.folder
        Dim myFolderName As Scripting.folder
        Dim myTextFile As Scripting.TextStream
        
    Dim isThisIsTallyFolder As Boolean
    
    Dim myFolderNamePath As String
    
        Dim CmpName As String
        Dim CmpGST As String
        Dim CmpPAN As String
        Dim CmpFromDate As String
        Dim CmpToDate As String
        
        Dim CmpGrossProfit As String
        Dim CmpClosingStock As String
        Dim CmpSalesAccount As String
        Dim CmpDirectIncome As String
        Dim CmpNetProfit As String
        Dim CmpTransfered As String
        
        Dim CmpDetailsTxtFileName As String

    
    Set fso = New Scripting.FileSystemObject
    Set myFolderName = fso.GetFolder(DataFolder)
    'myFolderName = Dir(DirClientsFolders.List(DirClientsFolders.ListIndex), vbDirectory)
    'myFolderNamePath = DirClientsFolders.List(DirClientsFolders.ListIndex)
    
    If IsNumeric(myFolderName.ShortName) Then
        
        If fso.FileExists(myFolderName & "\Company.999") Then
            CmpDetailsTxtFileName = myFolderName & "\Company.999"
        ElseIf fso.FileExists(myFolderName & "\Company.888") Then
            CmpDetailsTxtFileName = myFolderName & "\Company.888"
        End If

    
        If fso.FileExists(CmpDetailsTxtFileName) Then
            myForm.LblCurrentFolderName = myFolderName.ShortName
            
            Set textFile = fso.OpenTextFile(CmpDetailsTxtFileName, ForReading, False, TristateMixed)
                
                CmpName = textFile.ReadLine
                CmpName = Trim(Mid(CmpName, InStrRev(CmpName, "!") + 1))
                myForm.LblCmpName.Caption = Space(1) & CmpName
                
                CmpGST = textFile.ReadLine
                CmpGST = Trim(Mid(CmpGST, InStrRev(CmpGST, "!") + 1))
                myForm.LblCmpGST = CmpGST
                
                CmpPAN = textFile.ReadLine
                CmpPAN = Trim(Mid(CmpPAN, InStrRev(CmpPAN, "!") + 1))
                myForm.LblCmpPAN = CmpPAN
                
                CmpFromDate = textFile.ReadLine
                CmpFromDate = Trim(Mid(CmpFromDate, InStrRev(CmpFromDate, "!") + 1))
                myForm.LblYearFrom = Format(CmpFromDate, "dd-MMM-yyyy")
                
                CmpToDate = textFile.ReadLine
                CmpToDate = Trim(Mid(CmpToDate, InStrRev(CmpToDate, "!") + 1))
                myForm.LblYearTo = Format(CmpToDate, "dd-MMM-yyyy")
                
                textFile.ReadLine
                
                CmpGrossProfit = textFile.ReadLine
                CmpGrossProfit = Trim(Mid(CmpGrossProfit, InStrRev(CmpGrossProfit, "!") + 1))
                myForm.LblGrossProfit = CmpGrossProfit & Space(1)
                
                CmpClosingStock = textFile.ReadLine
                CmpClosingStock = Trim(Mid(CmpClosingStock, InStrRev(CmpClosingStock, "!") + 1))
                myForm.LblClosingStock.Caption = CmpClosingStock & Space(1)
                
                CmpSalesAccount = textFile.ReadLine
                CmpSalesAccount = Trim(Mid(CmpSalesAccount, InStrRev(CmpSalesAccount, "!") + 1))
                myForm.LblSalesAccount.Caption = CmpSalesAccount & Space(1)
                
                CmpDirectIncome = textFile.ReadLine
                CmpDirectIncome = Trim(Mid(CmpDirectIncome, InStrRev(CmpDirectIncome, "!") + 1))
                myForm.LblDirectIncome.Caption = CmpDirectIncome & Space(1)
                
                textFile.ReadLine
                
                CmpNetProfit = textFile.ReadLine
                CmpNetProfit = Trim(Mid(CmpNetProfit, InStrRev(CmpNetProfit, "!") + 1))
                myForm.LblNetProfit.Caption = CmpNetProfit & Space(1)
                
                CmpTransfered = textFile.ReadLine
                CmpTransfered = Trim(Mid(CmpTransfered, InStrRev(CmpTransfered, "!") + 1))
                myForm.LblTransfered.Caption = CmpTransfered & Space(1)
                
                Set textFile = Nothing
            
        End If
    Else
        Call ClearAllControls
    End If
    
End Function

Public Function ClearAllControls()

        myForm.LblCmpName.Caption = ""
        myForm.LblCmpGST.Caption = ""
        myForm.LblCmpPAN.Caption = ""
        myForm.LblYearFrom.Caption = ""
        myForm.LblYearTo.Caption = ""
            myForm.LblGrossProfit.Caption = ""
            myForm.LblClosingStock.Caption = ""
            myForm.LblSalesAccount.Caption = ""
            myForm.LblDirectIncome.Caption = ""
        myForm.LblNetProfit.Caption = ""
        myForm.LblTransfered.Caption = ""
        
        myForm.LblParentFolder.Caption = ""
        myForm.LblCurrentWorkingFolder.Caption = ""
        
        myForm.LblCurrentFolderName.Caption = ""
        myForm.LblFoldersTotal.Caption = ""
        myForm.LblFoldersRemains.Caption = ""
        myForm.LblFolderStatus.Caption = ""
        
End Function

Public Function tdlCompanyCreateTxtForTallyERP9()

    Dim fso As Scripting.FileSystemObject
    Dim TDLtxtFile As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject

CreateTDL:
            Set TDLtxtFile = fso.CreateTextFile(App.Path & "\tdlCompanyCreateTxtForTallyERP9.txt", ForWriting, False)
            
                TDLtxtFile.WriteLine ";Use the blow said Command and Paraneters"
                TDLtxtFile.WriteLine ";Tally.exe /NOINILOAD /NOINITDL /TDL:D:\Documents\myTDL\myTDL.txt /NOGUI /DATA:" & """D:\Tally.ERP9\Data""" & "/LOAD:10000 /ACTION:CALL:CreateTxtFile"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[System : Events]"
                TDLtxtFile.WriteLine "AppStart1:System Start:True:Trigger Key:Alt+W"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[#Part:SVCompanyUser]"
                TDLtxtFile.WriteLine "    On:Focus:yes:Action:Call:CmpPassword"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[#Menu:Gateway of Tally]"
                TDLtxtFile.WriteLine "    Add:Top Button:btnClickMe"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[Button:btnClickMe]"
                TDLtxtFile.WriteLine "    Title:" & """Click Me"""
                TDLtxtFile.WriteLine "    Key:Alt+W"
                TDLtxtFile.WriteLine "    Action:Call:CreateTxtFile"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[Function:CmpPassword]"
                TDLtxtFile.WriteLine "    00:Action:Trigger Key:Esc"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[Function:CreateTxtFile]"
                TDLtxtFile.WriteLine "    10:If:$(Company,##SVCurrentCompany).IsSecurityOn=No"
                TDLtxtFile.WriteLine "    20:Open File:" & """C:\DATA\Company.999""" & ":Text:Write"
                TDLtxtFile.WriteLine "    30:Truncate File"
                TDLtxtFile.WriteLine "    40:Write File Line:" & """Company Name ! """ & "+ $$String:##SVCurrentCompany"
                TDLtxtFile.WriteLine "    50:Write File Line:" & """Company GST  ! """ & "+ $$String:@@CMPCurrGSTNumber"
                TDLtxtFile.WriteLine "    60:Write File Line:" & """PAN          ! """ & "+ $$String:@@PanNumber"
                TDLtxtFile.WriteLine "    70:Write File Line:" & """Date From    ! """ & "+ $$String:@@StartDate"
                TDLtxtFile.WriteLine "    80:Write File Line:" & """Date To      ! """ & "+ $$String:@@EndingDate"
                TDLtxtFile.WriteLine "    90:Write File Line:"""""
                TDLtxtFile.WriteLine "   100:Write File Line:" & """Gross Profit ! """ & "+ $$String:@@GrossProfitTot"
                TDLtxtFile.WriteLine "   110:Write File Line:" & """Closing Stock! """ & "+ $$String:@@BSClosingStock"
                TDLtxtFile.WriteLine "   120:Write File Line:" & """Sales Account! """ & "+ $$String:@@NettSales"
                TDLtxtFile.WriteLine "   130:Write File Line:" & """Direct Income! """ & "+ $$String:@@DirectIncome"
                TDLtxtFile.WriteLine "   140:Write File Line:"""""
                TDLtxtFile.WriteLine "   150:Write File Line:" & """Nett Profit  ! """ & "+ $$String:@@NetProfitTot"
                TDLtxtFile.WriteLine "   160:Write File Line:" & """Transfered   ! """ & "+ $$String:@@BSClosingProfit"
                TDLtxtFile.WriteLine "   170:Close Target File"
                TDLtxtFile.WriteLine "   180:End If"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[System:Formula]"
                TDLtxtFile.WriteLine "        CMPCurrGSTNumber: $GSTRegNumber:TaxUnit:@@CMPExcisePrimaryGodown"
                TDLtxtFile.WriteLine "        PanNumber       : $IncomeTAXNumber:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        StartDate       : $$Max:$$SystemPeriodFrom:$StartingFrom:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        EndingDate      : $$Min:@@FinYrEnding:$EndingAt:Company:##SVCurrentCompany"
    
                TDLtxtFile.WriteLine "        GrossProfitTot  : $$AsPositive:$$Negative:$ClosingGrossProfit:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        BSClosingStock  : $$AsPositive:$ClosingBalance:Group:$$GroupStock"
                TDLtxtFile.WriteLine "        NettSales       : $TBalClosing:Group:$$GroupSales"
                TDLtxtFile.WriteLine "        DirectIncome    : $ClosingBalance:Group:$$GroupDirectIncomes"
    
                TDLtxtFile.WriteLine "        NetProfitTot    : $$AsPositive:$$Negative:$ClosingProfit:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        BSClosingProfit : $$AsPositive:$ClosingProfit:Company:##SVCurrentCompany"
                
            TDLtxtFile.Close
            
        If Not fso.FileExists(App.Path & "\tdlCompanyCreateTxtForTallyERP9.txt") Then
            GoTo CreateTDL:
        End If
        
    Set fso = Nothing
    Set TDLtxtFile = Nothing
    
End Function

Public Function tdlCompanyCreateTxtForTallyPrime()

    Dim fso As Scripting.FileSystemObject
    Dim TDLtxtFile As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject

CreateTDL:
            Set TDLtxtFile = fso.CreateTextFile(App.Path & "\tdlCompanyCreateTxtForTallyPrime.txt", ForWriting, False)
            
                TDLtxtFile.WriteLine ";Use the blow said Command and Paraneters"
                TDLtxtFile.WriteLine ";Tally.exe /NOINILOAD /NOINITDL /TDL:D:\Documents\myTDL\myTDL.txt /NOGUI /DATA:" & """D:\Tally.ERP9\Data""" & "/LOAD:10000 /ACTION:CALL:CreateTxtFile"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[System : Events]"
                TDLtxtFile.WriteLine "AppStart1:System Start:True:Trigger Key:Alt+W"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[#Part:SVCompanyUser]"
                TDLtxtFile.WriteLine "    On:Focus:yes:Action:Call:CmpPassword"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[#Menu:Gateway of Tally]"
                TDLtxtFile.WriteLine "    Add:Top Button:btnClickMe"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[Button:btnClickMe]"
                TDLtxtFile.WriteLine "    Title:" & """Click Me"""
                TDLtxtFile.WriteLine "    Key:Alt+W"
                TDLtxtFile.WriteLine "    Action:Call:CreateTxtFile"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[Function:CmpPassword]"
                TDLtxtFile.WriteLine "    00:Action:Trigger Key:Esc"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[Function:CreateTxtFile]"
                TDLtxtFile.WriteLine "    10:If:$(Company,##SVCurrentCompany).IsSecurityOn=No"
                TDLtxtFile.WriteLine "    20:Open File:" & """C:\DATA\Company.888""" & ":Text:Write"
                TDLtxtFile.WriteLine "    30:Truncate File"
                TDLtxtFile.WriteLine "    40:Write File Line:" & """Company Name ! """ & "+ $$String:##SVCurrentCompany"
                TDLtxtFile.WriteLine "    50:Write File Line:" & """Company GST  ! """ & "+ $$String:@@CMPCurrGSTNumber"
                TDLtxtFile.WriteLine "    60:Write File Line:" & """PAN          ! """ & "+ $$String:@@PanNumber"
                TDLtxtFile.WriteLine "    70:Write File Line:" & """Date From    ! """ & "+ $$String:@@StartDate"
                TDLtxtFile.WriteLine "    80:Write File Line:" & """Date To      ! """ & "+ $$String:@@EndingDate"
                TDLtxtFile.WriteLine "    90:Write File Line:"""""
                TDLtxtFile.WriteLine "   100:Write File Line:" & """Gross Profit ! """ & "+ $$String:@@GrossProfitTot"
                TDLtxtFile.WriteLine "   110:Write File Line:" & """Closing Stock! """ & "+ $$String:@@BSClosingStock"
                TDLtxtFile.WriteLine "   120:Write File Line:" & """Sales Account! """ & "+ $$String:@@NettSales"
                TDLtxtFile.WriteLine "   130:Write File Line:" & """Direct Income! """ & "+ $$String:@@DirectIncome"
                TDLtxtFile.WriteLine "   140:Write File Line:"""""
                TDLtxtFile.WriteLine "   150:Write File Line:" & """Nett Profit  ! """ & "+ $$String:@@NetProfitTot"
                TDLtxtFile.WriteLine "   160:Write File Line:" & """Transfered   ! """ & "+ $$String:@@BSClosingProfit"
                TDLtxtFile.WriteLine "   170:Close Target File"
                TDLtxtFile.WriteLine "   180:End If"
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine
                TDLtxtFile.WriteLine "[System:Formula]"
                TDLtxtFile.WriteLine "        CMPCurrGSTNumber: $GSTRegNumber:TaxUnit:@@CMPExcisePrimaryGodown"
                TDLtxtFile.WriteLine "        PanNumber       : $IncomeTAXNumber:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        StartDate       : $$Max:$$SystemPeriodFrom:$StartingFrom:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        EndingDate      : $$Min:@@FinYrEnding:$EndingAt:Company:##SVCurrentCompany"
    
                TDLtxtFile.WriteLine "        GrossProfitTot  : $$AsPositive:$$Negative:$ClosingGrossProfit:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        BSClosingStock  : $$AsPositive:$ClosingBalance:Group:$$GroupStock"
                TDLtxtFile.WriteLine "        NettSales       : $TBalClosing:Group:$$GroupSales"
                TDLtxtFile.WriteLine "        DirectIncome    : $ClosingBalance:Group:$$GroupDirectIncomes"
    
                TDLtxtFile.WriteLine "        NetProfitTot    : $$AsPositive:$$Negative:$ClosingProfit:Company:##SVCurrentCompany"
                TDLtxtFile.WriteLine "        BSClosingProfit : $$AsPositive:$ClosingProfit:Company:##SVCurrentCompany"
                
            TDLtxtFile.Close
            
        If Not fso.FileExists(App.Path & "\tdlCompanyCreateTxtForTallyPrime.txt") Then
            GoTo CreateTDL:
        End If
        
    Set fso = Nothing
    Set TDLtxtFile = Nothing
    
End Function

