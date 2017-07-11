Attribute VB_Name = "Module2"
Option Explicit
Public i As Long, j As Long
Public mold_id As Integer
Public resin_dir As String
Public mold_dir As String
Public MyObj As Object, MySource As Object, file As Variant

'is this it???? 7/11/2017 350 pm

Public Function tom()

    j = 1
'hello
    i = 1
    Do Until Range("B" & i) = ""
        Dim StrFile As String
        StrFile = Dir("\\spfs1\stone\Engineering\Automotive Material Specifications\Resin Spec List by RES #\")
        Do While Len(StrFile) > 0
            Debug.Print j & StrFile
            StrFile = Dir
            j = j + 1
            If Right(Left(StrFile, 13), 8) = Range("A" & i) Then
                If Len(Dir("\\spfs1\stone\Mold_Books\" & Range("B" & i), vbDirectory)) = 0 Then MkDir "\\spfs1\stone\Mold_Books\" & Range("B" & i)
                FileCopy "\\spfs1\stone\Engineering\Automotive Material Specifications\Resin Spec List by RES #\" & StrFile, "\\spfs1\stone\Mold_Books\" & Range("B" & i) & "\" & StrFile
            End If
        Loop
        'checking the files
        'for next loop for resins
        '   Set MySource = MyObj.GetFolder("\\spfs1\stone\Engineering\Automotive Material Specifications\Resin Spec List by RES #")
        '   For Each file In MySource.Files
        '      If Right(Left(file.Name, 13), 8) = Range("A" & i) Then
        '      FileCopy resin_dir, mold_dir
        '      End If
        '   Next file
        i = i + 1
    Loop

End Function

Public Sub ExportSourceFiles(destPath As String)

    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
 
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
        ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
        ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
        ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
        ToFileExtension = vbNullString
    End Select
 
End Function

Public Function CommandLine(command As String, Optional ByVal keepAlive As _
                                              Boolean = True, Optional windowState As VbAppWinStyle = VbAppWinStyle.vbNormalFocus) _
                                              As Boolean

     
    '--------------------------------------------------------------------------------
    ' Procedure : CommandLine
    ' Author    : Aaron Bush (Oorang)
    ' Date      : 10/02/2007
    ' Purpose   : Provides a simple interface to execute a command lines from VBA.
    ' Input(s)  :
    '               command     : The DOS command you wish to execute.
    '               keepAlive   : Keeps the DOS window open *after* command has been
    '                             executed. Default behavior is to auto-close. (See
    '                             remarks section for additional information.)
    '               windowState : Determines the window state of the DOS prompt
    '                             *during* command execution.
    ' Output    : True if completed with no errors, False if error encountered.
    ' Remarks   : If the windowState property is set to vbHide while the keepAlive
    '             parameter is set to True, then windowState will be changed to
    '             vbNormalFocus.
    '--------------------------------------------------------------------------------
    On Error GoTo Err_Hnd
    Const lngMatch_c As Long = 0
    Const strCMD_c As String = "cmd.exe"
    Const strComSpec_c As String = "COMSPEC"
    Const strTerminate_c As String = " /c "
    Const strKeepAlive_c As String = " /k "
    Dim strCmdPath As String
    Dim strCmdSwtch As String
    If keepAlive Then
        If windowState = vbHide Then
            windowState = vbNormalFocus
        End If
        strCmdSwtch = strKeepAlive_c
    Else
        strCmdSwtch = strTerminate_c
    End If
    strCmdPath = VBA.Environ$(strComSpec_c)
    If VBA.StrComp(VBA.Right$(strCmdPath, 7), strCMD_c, vbTextCompare) <> _
       lngMatch_c Then
        strCmdSwtch = vbNullString
    End If
    VBA.Shell strCmdPath & strCmdSwtch & command, windowState
    CommandLine = True
    Exit Function
Err_Hnd:
    CommandLine = False
End Function

'CommandLine ("cd " & Application.ActiveWorkbook.Path & " & pause")
'If Len(Dir(Application.ActiveWorkbook.Path & "\.git", vbDirectory)) = 0 Then
'MkDir prefixstrFlpath & mold_id
'End If

Public Sub git_push_code()
    Dim dirloc As String
    dirloc = Application.ActiveWorkbook.Path & "\" & Application.ActiveWorkbook.Name & "code_files"
    If Dir(dirloc, vbDirectory) = "" Then
        MkDir dirloc
    End If
    If Len(Dir(Application.ActiveWorkbook.Path & "\" & Application.ActiveWorkbook.Name & "code_files\.git", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive Or vbDirectory)) = 0 Then
        'git init
        CommandLine "cd " & dirloc & " & git init " 'creating the .git
    End If
    Call CommandLine("cd " & dirloc & " & git add . & git commit -m '#" & InputBox("What did you change") & "'") 'this commits the git locally
    If True Then                                 'no remote define

        'Call CommandLine("cd " & dirloc & " & git remote add origin " & InputBox("paste web url here") & " & git push -u origin master")
    End If
    Call CommandLine("cd " & dirloc & " & git push -u origin master")
    'https://github.com/danielleevandenbosch/molds_and_resins.git
End Sub

Public Function FileExists(filename As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filename) Then FileExists = True Else FileExists = False
End Function


