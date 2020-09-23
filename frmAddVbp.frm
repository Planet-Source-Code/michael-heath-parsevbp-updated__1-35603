VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddVbp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Visual Basic Project"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Return"
      Height          =   420
      Left            =   1440
      TabIndex        =   30
      Top             =   7785
      Width           =   1230
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "&Parse"
      Height          =   420
      Left            =   225
      TabIndex        =   2
      Top             =   7785
      Width           =   1230
   End
   Begin VB.CommandButton cmdTree 
      Caption         =   "Show &Tree"
      Enabled         =   0   'False
      Height          =   420
      Left            =   1440
      TabIndex        =   31
      Top             =   7380
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame Frame5 
      Caption         =   "Title"
      Height          =   600
      Left            =   7740
      TabIndex        =   32
      Top             =   7110
      Width           =   2130
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   45
         TabIndex        =   33
         Top             =   225
         Width           =   2040
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9900
      Top             =   7245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Re&set"
      Height          =   420
      Left            =   225
      TabIndex        =   29
      Top             =   7380
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Caption         =   "Description"
      Height          =   1185
      Left            =   2925
      TabIndex        =   27
      Top             =   6030
      Width           =   2715
      Begin VB.TextBox txtDescription 
         Height          =   870
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   2490
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comments"
      Height          =   1185
      Left            =   2925
      TabIndex        =   25
      Top             =   7245
      Width           =   4830
      Begin VB.TextBox txtComment 
         Height          =   870
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   4650
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "CompanyName"
      Height          =   600
      Left            =   9855
      TabIndex        =   17
      Top             =   6615
      Width           =   2130
      Begin VB.TextBox txtComp 
         Height          =   285
         Left            =   45
         TabIndex        =   18
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "ExeName"
      Height          =   600
      Left            =   7740
      TabIndex        =   21
      Top             =   6615
      Width           =   2130
      Begin VB.TextBox txtExe 
         Height          =   285
         Left            =   45
         TabIndex        =   22
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "ProductName"
      Height          =   600
      Left            =   7740
      TabIndex        =   23
      Top             =   6030
      Width           =   4245
      Begin VB.TextBox txtProd 
         Height          =   285
         Left            =   45
         TabIndex        =   24
         Top             =   225
         Width           =   4110
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "FileDescription"
      Height          =   1185
      Left            =   5625
      TabIndex        =   19
      Top             =   6030
      Width           =   2130
      Begin VB.TextBox txtFileDes 
         Height          =   870
         Left            =   45
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   330
      Left            =   2970
      TabIndex        =   1
      Top             =   0
      Width           =   1230
   End
   Begin VB.TextBox txtProj 
      Height          =   285
      Left            =   4185
      TabIndex        =   0
      Top             =   0
      Width           =   7800
   End
   Begin MSComctlLib.TreeView treMain 
      Height          =   7260
      Left            =   45
      TabIndex        =   34
      Top             =   45
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   12806
      _Version        =   393217
      Indentation     =   176
      Style           =   7
      ImageList       =   "imgSmall"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   10395
      Top             =   7245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":118C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":2318
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":34A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":4630
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":57BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":650E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":6F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddVbp.frx":788E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame11 
      Height          =   150
      Left            =   1530
      TabIndex        =   16
      Top             =   5850
      Width           =   11220
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   2925
      TabIndex        =   3
      Top             =   360
      Width           =   9060
      Begin VB.Frame Frame10 
         Caption         =   "Dependencies"
         Height          =   2715
         Left            =   6030
         TabIndex        =   14
         Top             =   2790
         Width           =   3030
         Begin VB.ListBox lstDep 
            Height          =   2400
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   225
            Width           =   2850
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Classes"
         Height          =   2715
         Left            =   6030
         TabIndex        =   12
         Top             =   90
         Width           =   3030
         Begin VB.ListBox lstCls 
            Height          =   2400
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   225
            Width           =   2850
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Modules"
         Height          =   2715
         Left            =   3015
         TabIndex        =   10
         Top             =   90
         Width           =   3030
         Begin VB.ListBox lstMod 
            Height          =   2400
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   225
            Width           =   2850
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Forms"
         Height          =   2715
         Left            =   0
         TabIndex        =   8
         Top             =   90
         Width           =   3030
         Begin VB.ListBox lstExp 
            Height          =   2400
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   225
            Width           =   2850
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "UserControls"
         Height          =   2715
         Left            =   0
         TabIndex        =   6
         Top             =   2790
         Width           =   3030
         Begin VB.ListBox lstCon 
            Height          =   2400
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   225
            Width           =   2850
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "PropertyPages"
         Height          =   2715
         Left            =   3015
         TabIndex        =   4
         Top             =   2790
         Width           =   3030
         Begin VB.ListBox lstProp 
            Height          =   2400
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   225
            Width           =   2850
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmAddVbp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    ' Browse for a VB project file
    txtProj.Text = OpenFile(Me, "Visual Basic Projects|*.vbp", False, App.Path)
    cmdBrowse.Enabled = False
End Sub

Private Sub cmdComments_Click()

End Sub

Private Sub cmdExit_Click()
' Save the stuff and leave
' ToDo: Save the info and add to the settings file
MsgBox "ToDo:  Save the info"
Unload Me
Set frmAddVbp = Nothing

End

End Sub

Public Sub cmdParse_Click()
' Declare some variables
Dim strHold As String
Dim strTemp As String
Dim g As Integer

' Make sure there is a file in the txtproj.text box
If txtProj.Text = "" Then Exit Sub

' Disable the cmdParse button so we don't get the data reloaded, you could
' Also clear out all the objects and force the prog to restart the process
cmdParse.Enabled = False
cmdTree.Enabled = True
' Open the VB Project File and start parsing and extracting the information needed
On Error GoTo OpERR  ' If there is an error, let's handle it
    Open txtProj For Input As #1
        Do While Not EOF(1)         ' Loops through the file until we reach the end
            Input #1, strTemp       ' Raw Data read from the VB Project file
            
            ' Read to the left of the raw data - Decide what we have and
            ' Act on it according to the Select Case
            
            Select Case Left(strTemp, 5)
                Case "Form="
                    ' Extract the forms
                    strHold = Right(strTemp, Len(strTemp) - 5) ' Drop off Raw Data
                    lstExp.AddItem strHold
                Case "Class"
                    ' Extract the class modules
                    strHold = strTemp
                    For g = 1 To Len(strTemp)
                    strHold = Right(strHold, Len(strHold) - 1)
                        If InStr(strHold, ";") Then
                            ' Continue removing characters
                        Else
                            ' We have what we need, let's get out of the for/next loop
                            Exit For
                        End If
                    Next g
                    lstCls.AddItem strHold
                Case "Title"
                    strHold = Right(strTemp, Len(strTemp) - 7)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtTitle.Text = strHold
            End Select
            
            Select Case Left(strTemp, 6)
                Case "Module"
                    ' Extract the bas files
                    strHold = strTemp
                    For g = 1 To Len(strTemp)
                    strHold = Right(strHold, Len(strHold) - 1)
                        If InStr(strHold, ";") Then
                            ' Continue removing characters
                        Else
                            ' We have what we need, let's get out of the for/next loop
                            Exit For
                        End If
                    Next g
                    lstMod.AddItem strHold
                Case "Object"
                    ' Extract dependencies - ocx's
                    strHold = strTemp
                    For g = 1 To Len(strTemp)
                    strHold = Right(strHold, Len(strHold) - 1)
                        If InStr(strHold, ";") Then
                            ' Continue removing characters
                        Else
                            ' We have what we need, let's get out of the for/next loop
                            Exit For
                        End If
                    Next g
                    lstDep.AddItem strHold
            End Select
            
            Select Case Left(strTemp, 9)
                Case "ExeName32"
                    ' Extract the Executable name
                    strHold = Right(strTemp, Len(strTemp) - 11)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtExe.Text = strHold
            End Select
            
            Select Case Left(strTemp, 11)
                Case "Description"
                    ' Extract the description
                    strHold = Right(strTemp, Len(strTemp) - 13)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtDescription.Text = strHold
                Case "UserControl"
                    ' Extract the usercontrols
                    strHold = Right(strTemp, Len(strTemp) - 12)
                    lstCon.AddItem strHold
                Case "PropertyPag"
                    ' Extract the property pages
                    strHold = Right(strTemp, Len(strTemp) - 13)
                    lstProp.AddItem strHold
                Case "VersionComp"
                    ' Extract the Company name
                    strHold = Right(strTemp, Len(strTemp) - 20)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtComp.Text = strHold
                Case "VersionFile"
                    ' Extract the file description
                    strHold = Right(strTemp, Len(strTemp) - 24)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtFileDes.Text = strHold
                Case "VersionProd"
                    ' Extract the Product Name
                    strHold = Right(strTemp, Len(strTemp) - 20)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtProd.Text = strHold
            End Select
            
            Select Case Left(strTemp, 15)
                Case "VersionComments"
                    ' Extract the comments
                    strHold = Right(strTemp, Len(strTemp) - 17)
                    strHold = Left(strHold, Len(strHold) - 1)
                    txtComment.Text = strHold
            End Select
        Loop
    Close #1
        ' Populate the TreeNode
        ShowTree treMain, lstExp, lstMod, lstCls, lstCon, lstProp, lstDep, txtTitle.Text
    Exit Sub
OpERR:
' handle the error
    Dim ErrNum As Long
    ErrNum = Err.Number
    Select Case ErrNum
        Case 53
            MsgBox "File Does Not Exist or it is Corrupt.", vbOKOnly, "Error - 53"
        Case Else
            MsgBox "Error Occurred - " & Err.Number, vbOKOnly, "ERROR"
    End Select
End Sub
Private Sub Command1_Click()
End Sub

Private Sub cmdReset_Click()
    Unload Me
    Set frmAddVbp = Nothing
    Load frmAddVbp
    frmAddVbp.Show

End Sub

Private Sub Command2_Click()
End Sub

Private Sub cmdTree_Click()
    ShowTree treMain, lstExp, lstMod, lstCls, lstCon, lstProp, lstDep, txtTitle.Text
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Visual Basic Project Parser" & Chr(10) _
        & "By: Michael Heath" & Chr(10) & _
        "7 JUN 2002", vbOKOnly + vbInformation, "ABOUT PARSEVBP"
End Sub
