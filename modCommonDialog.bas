Attribute VB_Name = "modCommonDialog"
' ***********************************************
' * Common Dialog Code                          *
' * Code by: Michael Heath                      *
' * Code Date: Sometime in '99                  *
' * Email:  mheath@indy.net                     *
' * -----------------------                     *
' * Simple CommonDialog Routines for Open/Save  *
' * File operations.                            *
' * Being revised from time to time.            *
' ***********************************************

Public CurrentFileName As String     ' Holds filename
Public NoOpen As Boolean             ' Flag for file operation
Public vFileName As String           ' Name of File without the path
Public strAppName As String          ' Name of this App
Public Function OpenFile(vForm As Form, vFilter As String, SetCaption As Boolean, Optional InitialDirectory As String)
' Reset the NoOpen flag
NoOpen = False
CurrentFileName = ""
' Handle errors
    On Error GoTo OpenProblem
    
    If InitialDirectory = "" Then
        vForm.CommonDialog1.InitDir = "c:\windows"
    Else
        vForm.CommonDialog1.InitDir = InitialDirectory
    End If
    
    vForm.CommonDialog1.Filter = vFilter
    vForm.CommonDialog1.FilterIndex = 1
    ' Display an Open dialog box.
    vForm.CommonDialog1.Action = 1
    ' Set the Caption of the Parent Form if it is set to true
    If SetCaption = True Then
        vForm.Caption = strAppName & " - " & vForm.CommonDialog1.filename
        vForm.Caption = strAppName & " - " & vForm.CommonDialog1.filename
    End If
    
    CurrentFileName = vForm.CommonDialog1.filename
    vFileName = vForm.CommonDialog1.FileTitle
    If CurrentFileName = "" Then NoOpen = True
    OpenFile = CurrentFileName
    Exit Function

OpenProblem:
    ' There was a problem so let's flag NoOpen as true
    NoOpen = True
    Exit Function

End Function

Public Function SaveFile(vForm As Form, vFilter As String, SetCaption As Boolean, Optional InitialDirectory As String)
On Error GoTo SaveERR
    Dim FileNum As Integer
    ' Set Initial Directory to open and FileTypes
    If InitialDirectory = "" Then
        vForm.CommonDialog1.InitDir = App.Path & "\Save"
    Else
        vForm.CommonDialog1.InitDir = InitialDirectory
    End If
        
        vForm.CommonDialog1.Filter = vFilter
        If CurrentFileName = "" Then
            ' Code revised - Do nothing
            CurrentFileName = "NewFile"
        Else
            vForm.CommonDialog1.filename = CurrentFileName
        End If
        
        If SetCaption = True Then
            vForm.Caption = strAppName & " - " & vForm.CommonDialog1.filename
            vForm.Caption = strAppName & " - " & vForm.CommonDialog1.filename
        End If
            
            If vForm.CommonDialog1.CancelError = True Then
                SaveFile = ""
                CurrentFileName = ""
            End If
            
            vForm.CommonDialog1.ShowSave
            CurrentFileName = vForm.CommonDialog1.filename
            If FileExist = True Then
                Dim strQ As String
                strQ = MsgBox(CurrentFileName & " already exist, creating a new one will overwrite your current settings. Are you sure you want to do this?", vbYesNo + vbCritical, "Warning: File Exist")
                    Select Case strQ
                        Case vbYes
                            SaveFile = CurrentFileName
                        Case vbNo
                            MsgBox "File Creation Aborted By User!", vbOKOnly, "Aborted"
                            SaveFile = ""
                            CurrentFileName = ""
                    End Select
            Else
                SaveFile = CurrentFileName
            End If
            
Exit Function
SaveERR:
    ' No real error trap, just exit the sub and give a description of the error
    MsgBox "An error occured. " & Err.Description, vbOKOnly + vbInformation, "Save Error"
End Function

Public Function PutFileInString(sFileName As String) As String
    'sFileName must include Path and file na
    '     me
    'eg "c:\Windows\notepad.exe"
    Dim iFree As Integer, sizeOfFile As Long
    Dim sFileString As String, sTemp As String
    iFree = FreeFile
    Open sFileName For Binary Access Read As iFree
    sizeOfFile = LOF(iFree)
    sFileString = Space$(sizeOfFile)
    Get iFree, , sFileString
    Close #iFree
    PutFileInString = sFileString
    
    ' Thanks to Robert Carter for this function
End Function

Public Function FileExist() As Boolean
On Error GoTo FileErr
If FileLen(CurrentFileName) >= 0 Then
    FileExist = True
Else
    FileExist = False
End If
Exit Function
FileErr:
    FileExist = False
End Function

