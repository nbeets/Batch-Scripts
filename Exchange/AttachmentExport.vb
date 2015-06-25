Attribute VB_Name = "Module1"
Option Explicit

Public Sub ExportAttachments()
    Dim objOL As Outlook.Application
    Dim objMsg As Object
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    Dim i As Long, lngCount As Long
    Dim filesRemoved As String, fName As String, strFolder As String, saveFolder As String, savePath As String
    Dim alterEmails As Boolean, overwrite As Boolean
    Dim result
    
    ' Added variables to enable Hashing
    Dim d As Integer ' Keep track of how many emails evaluated before deleting the temporary files.
    Dim tmpPath As String, fHash As String, fNameNoExt As String, fExt As String
    Dim fArr
    
    saveFolder = BrowseForFolder("Select the folder to save attachments to.")
    If saveFolder = vbNullString Then Exit Sub
    
    result = MsgBox("Do you want to remove attachments from selected file(s)? " & vbNewLine & _
    "(Clicking no will export attachments but leave the emails alone)", vbYesNo + vbQuestion)
    alterEmails = (result = vbYes)
    
    ' Create directory used to temporarily hold attachments when calculating their hash values.
    If Dir("C:\temp\OutlookScript", vbDirectory) = "" Then
        MkDir ("C:\temp\OutlookScript")
    End If
    
    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    d = 0
    For Each objMsg In objSelection
        If objMsg.Class = olMail Then
            Set objAttachments = objMsg.Attachments
            lngCount = objAttachments.Count
            If lngCount > 0 Then
                filesRemoved = ""
                For i = lngCount To 1 Step -1
                    fName = objAttachments.Item(i).FileName
                    tmpPath = "C:\temp\OutlookScript\" & fName
                    
                    'Get File hash value
                    objAttachments.Item(i).SaveAsFile tmpPath
                    fHash = FileToSHA1Hex(tmpPath)
                    'Debug.Print objAttachments.Item(i).Type
                    
                    fArr = Split(fName, ".")
                    fExt = fArr(UBound(fArr))
                    ReDim Preserve fArr(UBound(fArr) - 1)
                    fNameNoExt = Join(fArr, ".")
                    
                    savePath = saveFolder & "\" & fNameNoExt & "_" & fHash & "." & fExt
                    overwrite = True
                    While Dir(savePath) <> vbNullString And Not overwrite
                        Dim newFName As String
                        newFName = InputBox("The file '" & fName & _
                            "' already exists. Please enter a new file name, or just hit OK overwrite.", _
                            "Confirm File Name", fName)
                        If newFName = vbNullString Then GoTo skipfile
                        If newFName = fName Then overwrite = True Else fName = newFName
                        savePath = saveFolder & "\" & fName
                    Wend
                    
                    objAttachments.Item(i).SaveAsFile savePath
                    
                    If alterEmails Then
                        filesRemoved = filesRemoved & "<br>""" & objAttachments.Item(i).FileName & """ (" & _
                                                                formatSize(objAttachments.Item(i).size) & ") " & _
                            "<a href=""" & savePath & """>[Location Saved]</a>"
                        objAttachments.Item(i).Delete
                    End If
skipfile:
                Next i
                
                If alterEmails Then
                    filesRemoved = "<b>Attachments removed</b>: " & filesRemoved & "<br><br>"
                    
                    Dim objDoc As Object
                    Dim objInsp As Outlook.Inspector
                    Set objInsp = objMsg.GetInspector
                    Set objDoc = objInsp.WordEditor

                    objMsg.HTMLBody = filesRemoved + objMsg.HTMLBody
                    objMsg.Save
                End If
            End If
        End If
        
        d = d + 1
        ' Ensure the HDD does not fill up from attachments.
        If (d Mod 50) = 0 Then
            On Error Resume Next
            Kill "C:\temp\OutlookScript\*.*" 'Delete all files.
            On Error GoTo 0
        End If
    Next
    
ExitSub:
    ' Delete all temporary files and folder.
    On Error Resume Next
    Kill "C:\temp\OutlookScript\*.*" 'Delete all files.
    RmDir "C:\temp\OutlookScript\" 'Delete Folder.
    On Error GoTo 0

    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
End Sub

Function formatSize(size As Long) As String
    Dim val As Double, newVal As Double
    Dim unit As String
    
    val = size
    unit = "bytes"
    
    newVal = Round(val / 1024, 1)
    If newVal > 0 Then
        val = newVal
        unit = "KB"
    End If
    newVal = Round(val / 1024, 1)
    If newVal > 0 Then
        val = newVal
        unit = "MB"
    End If
    newVal = Round(val / 1024, 1)
    If newVal > 0 Then
        val = newVal
        unit = "GB"
    End If
    
    formatSize = val & " " & unit
End Function

'Function purpose:  To Browser for a user selected folder.
'If the "OpenAt" path is provided, open the browser at that directory
'NOTE:  If invalid, it will open at the Desktop level
Function BrowseForFolder(Optional Prompt As String, Optional OpenAt As Variant) As String
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, Prompt, 0, OpenAt)

    On Error Resume Next
    BrowseForFolder = ShellApp.self.path
    On Error GoTo 0
    Set ShellApp = Nothing
     
    'Check for invalid or non-entries and send to the Invalid error handler if found
    'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":": If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\": If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
        Case Else: GoTo Invalid
    End Select
     
    Exit Function
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = vbNullString
End Function

Function BrowseForFile(Optional Prompt As String, Optional OpenAt As Variant) As String
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, Prompt, 16 + 16384, OpenAt)
    
    On Error Resume Next
    BrowseForFile = ShellApp.self.path
    On Error GoTo 0
    Set ShellApp = Nothing
     
    'Check for invalid or non-entries and send to the Invalid error handler if found
    'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":": If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\": If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
        Case Else: GoTo Invalid
    End Select
     
    Exit Function
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFile = vbNullString
End Function

'Source: http://stackoverflow.com/a/17858040
Private Function FileToSHA1Hex(sFileName As String) As String
    Dim enc
    Dim bytes
    Dim outstr As String
    Dim pos As Integer
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFileName)
    bytes = enc.ComputeHash_2((bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
    Next
    FileToSHA1Hex = outstr 'Returns a 40 byte/character hex string
    Set enc = Nothing
End Function

'Source: http://stackoverflow.com/a/17858040
Private Function GetFileBytes(ByVal path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    lngFileNum = FreeFile
    If LenB(Dir(path)) Then ''// Does file exist?
        Open path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function
