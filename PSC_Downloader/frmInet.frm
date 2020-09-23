VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmInet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PSC File Downloader"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Download files by author"
      Height          =   8745
      Left            =   30
      TabIndex        =   41
      Top             =   5160
      Visible         =   0   'False
      Width           =   10395
      Begin VB.CommandButton Command17 
         Caption         =   "List Files on this page"
         Height          =   285
         Left            =   8280
         TabIndex        =   46
         ToolTipText     =   "List the files on the current page allowing you to sellect the files to download."
         Top             =   210
         Width           =   1665
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Download All Results"
         Height          =   285
         Left            =   6510
         TabIndex        =   57
         Top             =   210
         Width           =   1665
      End
      Begin VB.CommandButton Command23 
         Caption         =   "PSC Advanced Search"
         Height          =   255
         Left            =   2160
         TabIndex        =   56
         Top             =   630
         Width           =   1845
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1500
         TabIndex        =   54
         ToolTipText     =   "Takes you to the psc search page"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   53
         ToolTipText     =   "stop the browser"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command21 
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   450
         TabIndex        =   52
         ToolTipText     =   "Go Forward"
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton Command21 
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   51
         ToolTipText     =   "Go Back"
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Download Files by this Author"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4110
         TabIndex        =   49
         ToolTipText     =   "When enabled allows you to download all files by the author of the current code."
         Top             =   210
         Width           =   2325
      End
      Begin VB.CommandButton Command19 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10050
         TabIndex        =   48
         ToolTipText     =   "Hide the browser"
         Top             =   0
         Width           =   345
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   7785
         Left            =   60
         TabIndex        =   45
         Top             =   900
         Width           =   10245
         ExtentX         =   18071
         ExtentY         =   13732
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Go"
         Height          =   285
         Left            =   3540
         TabIndex        =   44
         ToolTipText     =   "Search PSC for an author"
         Top             =   210
         Width           =   465
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         ToolTipText     =   "Enter the author name"
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Select a post from the list to be able to download all files by that author."
         Height          =   225
         Left            =   4140
         TabIndex        =   50
         Top             =   690
         Width           =   5595
      End
      Begin VB.Label Label4 
         Caption         =   "Search Author name"
         Height          =   225
         Left            =   60
         TabIndex        =   42
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Download Files"
      Height          =   375
      Left            =   6480
      TabIndex        =   28
      ToolTipText     =   "Download latest files uploaded to PSC."
      Top             =   4560
      Width           =   1485
   End
   Begin VB.Frame Frame4 
      Caption         =   "Latest 50 Postings"
      Height          =   4245
      Left            =   3750
      TabIndex        =   30
      Top             =   90
      Visible         =   0   'False
      Width           =   6645
      Begin VB.Frame Frame5 
         Caption         =   "Selection Options"
         Height          =   585
         Left            =   120
         TabIndex        =   35
         Top             =   3570
         Width           =   4035
         Begin VB.CommandButton Command22 
            Caption         =   "Unselect All"
            Height          =   285
            Left            =   2940
            TabIndex        =   55
            Top             =   210
            Width           =   1035
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Select All"
            Height          =   285
            Left            =   2190
            TabIndex        =   37
            Top             =   210
            Width           =   765
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Select Updated"
            Height          =   285
            Left            =   990
            TabIndex        =   38
            Top             =   210
            Width           =   1215
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Select New"
            Height          =   285
            Left            =   30
            TabIndex        =   36
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Download Selected"
         Height          =   285
         Left            =   4230
         TabIndex        =   33
         Top             =   3750
         Width           =   1635
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Hide"
         Height          =   285
         Left            =   5910
         TabIndex        =   32
         Top             =   3750
         Width           =   555
      End
      Begin VB.ListBox List1 
         Height          =   3180
         ItemData        =   "frmInet.frx":030A
         Left            =   90
         List            =   "frmInet.frx":030C
         MultiSelect     =   1  'Simple
         TabIndex        =   31
         Top             =   270
         Width           =   1035
      End
      Begin VB.ListBox List2 
         Height          =   3180
         ItemData        =   "frmInet.frx":030E
         Left            =   1110
         List            =   "frmInet.frx":0310
         MultiSelect     =   1  'Simple
         TabIndex        =   34
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resume"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8100
      TabIndex        =   27
      Top             =   4560
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pages Downloaded"
      Height          =   4245
      Index           =   1
      Left            =   7110
      TabIndex        =   21
      Top             =   90
      Width           =   3255
      Begin VB.FileListBox File2 
         Height          =   3015
         Left            =   90
         Pattern         =   "*.html;*.txt"
         TabIndex        =   24
         Top             =   210
         Width           =   3075
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Rename The Pages From PSC text"
         Height          =   285
         Left            =   90
         TabIndex        =   23
         Top             =   3450
         Width           =   3045
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move The Pages to a New Location."
         Height          =   285
         Left            =   90
         TabIndex        =   22
         Top             =   3840
         Width           =   3045
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9330
      TabIndex        =   20
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Frame Frame3 
      Caption         =   "Files Downloaded"
      Height          =   4245
      Index           =   0
      Left            =   3750
      TabIndex        =   16
      Top             =   90
      Width           =   3225
      Begin VB.CommandButton Command15 
         Caption         =   "browse"
         Height          =   225
         Left            =   2460
         TabIndex        =   40
         ToolTipText     =   "Select another path of files to rename or move"
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Move The Files to a New Location."
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   3840
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Rename The files From PSC text"
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   3450
         Width           =   3045
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   90
         Pattern         =   "*.zip"
         TabIndex        =   17
         Top             =   210
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Downloader Options"
      Height          =   4245
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   3645
      Begin VB.CommandButton Command9 
         Caption         =   "List the 50 latest postings to PSC."
         Height          =   255
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "This will create a list of the latest 50 postings including updated postings allowing you to dounload them."
         Top             =   1530
         Width           =   3375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Select to resume an eairler session."
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Resumes downloading where you left off from a previous session"
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Get The Lastest Code Posting."
         Height          =   225
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "This will retrive the number of the last file to be posted to PSC"
         Top             =   1260
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Caption         =   "File Download Progress for this session."
         Height          =   1035
         Left            =   90
         TabIndex        =   11
         Top             =   3120
         Width           =   3465
         Begin VB.TextBox txtURL 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   13
            Text            =   "0"
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   12
            Text            =   "0"
            Top             =   570
            Width           =   1455
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "# Downloaded Zip Files"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   300
            Width           =   1905
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "# Downloaded Pages"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1905
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Text            =   "0"
         Top             =   570
         Width           =   765
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Text            =   "0"
         Top             =   570
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInet.frx":0312
         Left            =   1680
         List            =   "frmInet.frx":0334
         TabIndex        =   4
         Text            =   "Visual Basic"
         Top             =   210
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2820
         TabIndex        =   3
         Text            =   "5"
         Top             =   2820
         Width           =   465
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Download Files From:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   660
         Width           =   285
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Code Type"
         Height          =   255
         Index           =   3
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInet.frx":038F
         Height          =   1245
         Left            =   180
         TabIndex        =   7
         Top             =   1830
         Width           =   3285
      End
   End
   Begin PSCDownloader.FileDownloader DF 
      Left            =   150
      Top             =   240
      _ExtentX        =   1799
      _ExtentY        =   1667
   End
   Begin VB.OptionButton optMethod 
      Caption         =   "&Post"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   510
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optMethod 
      Caption         =   "&Get"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   510
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "qqq"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   47
      ToolTipText     =   "Search PSC by Author"
      Top             =   4410
      Width           =   315
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Progress indication"
      Height          =   735
      Left            =   420
      TabIndex        =   39
      Top             =   4380
      Width           =   5985
   End
End
Attribute VB_Name = "frmInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CD As cDlg
Private Pausit As Boolean
Private StartSpot As String
Public EndString As String
Private sURL As String
Private sFileName As String
Private sText As String
Private sType As String
Private DataA As String
Private fileList() As String
Private DataB As String
Private Parts() As String
Private mFileName As String
Private a As Long
Private b As Long
Private downnew As Boolean
Private downByAuth As Boolean
Private SearchString As String
Private nNumFils As Long
Private curPageNum As Long
'
' API Constants
'
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

' CreateFile
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000

Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

Private Const CREATE_NEW As Long = 1
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const TRUNCATE_EXISTING As Long = 5

'
' API Types
'
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'
' API Functions
'
Private Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFileAPI Lib "KERNEL32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal nNoOverlapping As Long) As Long
Private Declare Function FlushFileBuffers Lib "KERNEL32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

Private Sub cmdGO_Click()
Frame4.Visible = False
Dim resp As String
Dim z As Long
Dim process As Boolean
process = False
On Error Resume Next
    Me.MousePointer = vbHourglass
    Command1.Caption = "Pause"
    Command1.Enabled = True
    Pausit = False
        If FileExist(App.Path & "\downloads.txt") = True Then
            process = True
            Open App.Path & "\downloads.txt" For Input As #1
                fileList = Split(Replace(Input(LOF(1), 1), """", ""), vbCrLf)
            Close #1
        End If
    Label3.Caption = "Please Wait while I download the files indicated" & vbCrLf & "You can pause the downloads at any time."
        For a = Val(Text1.Text) To Val(Text3.Text)
            Text1.Text = a
                If process Then
                    For z = 0 To UBound(fileList)
                        If Val(fileList(z)) = a Then
                            fileList(z) = ""
                            resp = MsgBox("The file " & a & " has been downloaded Before" & vbCrLf & "Do you want to download it again?", vbYesNo)
                                If resp = vbNo Then
                                    GoTo skipFile
                                End If
                        ElseIf Val(fileList(z)) < Val(Text1.Text) Then
                            fileList(z) = ""
                        End If
                    Next z
                End If
            Label3.Caption = "I am now downloading file #" & a & vbCrLf & "Please wait while I download the remaining #" & (Val(Text3.Text) - Val(Text1.Text)) & " files." & vbCrLf & "You can pause the downloads at any time."
            DoEvents
            sURL = "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=" & Val(Text1.Text) & EndString
            mFileName = GetLocalFileNameFromURL(sURL)
            sFileName = App.Path & "\Files\Pages\" & mFileName
            sText = ReadURL(sURL, sType)
                If InStr(1, sText, "/vb/scripts/ShowZip.asp") > 0 Then
                    Parts = Split(sText, "/vb/scripts/ShowZip.asp")
                    DataA = Left(Parts(1), InStr(1, Parts(1), "><") - 2)
                    DataB = "http://www.planetsourcecode.com/vb/scripts/ShowZip.asp" & DataA
                    DF.DownloadFile DataB, App.Path & "\Files\Zips\" & mFileName & ".zip"
                    txtURL.Text = Val(txtURL.Text) + 1
                ElseIf InStr(1, sText, "This submission no longer exists in the database") > 0 Then
                ElseIf InStr(1, sText, "This submission was disapproved by the moderator") > 0 Then
                ElseIf InStr(1, sText, "The author of this code has deleted it or it has been removed") > 0 Then
                ElseIf InStr(1, sText, "This code has not yet been processed and posted to the public") > 0 Then
                Else
                    If InStr(1, sText, "/vb/scripts/ShowCodeAsText.asp") > 0 Then
                        sURL = "http://www.planetsourcecode.com/vb/scripts/ShowCodeAsText.asp?txtCodeId=" & Val(Text1.Text) & "&amp;lngWId=1"
                        sText = ReadURL(sURL, sType)
                        sText = Replace(sText, Chr(13), "")
                        sText = Replace(sText, Chr(10), vbCrLf)
                    End If
                    Call WriteFile(sFileName, sText)
                    Text2.Text = Val(Text2.Text) + 1
                End If
            File1.Refresh
            File2.Refresh
                DoEvents
                    If Not a = Val(Text3.Text) Then
                        For b = 0 To Val(Text4.Text) * 2 'To 0 Step 1
                            Label3.Caption = "I have finished with file #" & Text1.Text & vbCrLf & "Pausing For " & (5 - (b \ 2)) & " Seconds"
                            Label3.Refresh
                            DoEvents
                                If Pausit = True Then Exit For
                            Sleep 500
                       Next b
                    End If
skipFile:
                If Pausit = True Then Exit For
          '  Text1.Text = Val(Text1.Text) + 1
       Next a
       If process Then
            DataA = ""
                For z = 0 To UBound(fileList)
                    If Len(fileList(z)) > 0 Then
                        If Len(DataA) = 0 Then
                            DataA = """" & fileList(z) & """"
                        Else
                            DataA = DataA & vbCrLf & """" & fileList(z) & """"
                        End If
                    End If
                Next z
                If Len(DataA) = 0 Then
                    Kill App.Path & "\Downloads.txt"
                Else
                    Open App.Path & "\downloads.txt" For Output As #1
                        Print #1, DataA
                    Close #1
                End If
        End If
    Call SaveSetting(App.EXEName, "Settings", StartSpot & "Start", Val(Text1.Text))
    Call SaveSetting(App.EXEName, "Settings", StartSpot & "End", Val(Text3.Text))
    Me.MousePointer = vbDefault
    Label3.Caption = "I have finished with your downloads."
End Sub


Private Sub WriteFile(ByVal sFileName As String, ByVal sContent As String)
Dim hFile As Long
Dim nBytesWritten As Long
Dim secAtt As SECURITY_ATTRIBUTES
Dim aBytes() As Byte
    hFile = CreateFile(sFileName & ".html", GENERIC_WRITE, FILE_SHARE_READ, secAtt, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0&)
        If hFile = INVALID_HANDLE_VALUE Then Exit Sub
    aBytes = StrConv(sContent, vbFromUnicode)
        If WriteFileAPI(hFile, aBytes(0), Len(sContent), nBytesWritten, 0&) <> 0 Then
            Call FlushFileBuffers(hFile)
        End If
    Call CloseHandle(hFile)
End Sub

Public Sub GetURLS()
        Select Case Combo1.ListIndex
            Case 0
                EndString = "&lngWId=1" 'VB
                StartSpot = "VB"
            Case 1
                EndString = "&lngWId=10" '.net
                StartSpot = ".Net"
            Case 2
                EndString = "&lngWId=2" 'java/java script
                StartSpot = "Java"
            Case 3
                EndString = "&lngWId=3" 'c/c++
                StartSpot = "c++"
            Case 4
                EndString = "&lngWId=4" 'asp/vb script
                StartSpot = "ASP"
            Case 5
                EndString = "&lngWId=9" 'cold fusion
                StartSpot = "ColdFus"
            Case 6
                EndString = "&lngWId=7" 'delphi
                StartSpot = "Delphi"
            Case 7
                EndString = "&lngWId=6" 'perl
                StartSpot = "Perl"
            Case 8
                EndString = "&lngWId=8" 'php
                StartSpot = "PHP"
            Case 9
                EndString = "&lngWId=5" 'sql
                StartSpot = "SQL"
            Case Else
                EndString = "&lngWId=1" 'VB
                StartSpot = "VB"
        End Select
    Text1.Text = GetSetting(App.EXEName, "Settings", StartSpot & "Start", "0")
    Text3.Text = GetSetting(App.EXEName, "Settings", StartSpot & "End", "0")
End Sub

Private Sub Combo1_Click()
    GetURLS
End Sub

Private Sub Command1_Click()
Dim Data As String
        If Command1.Caption = "Resume" Then
            Command1.Caption = "Pause"
            cmdGO_Click
        Else
            If downnew = True Then
                    For a = 0 To List1.ListCount - 1
                        If List1.Selected(a) = True Then
                            Data = Data & "::" & List1.List(a) & "*" & List2.List(a)
                        End If
                    Next a
                Call SaveSetting(App.EXEName, "Settings", StartSpot & "DownList", Data)
            ElseIf downByAuth = True Then
                Pausit = True
                'Command1.Caption = "Resume"
                'Call SaveSetting(App.EXEName, "Settings", "CurrentPage", curPageNum)
                'save the file number
                'on resume reload the current page and
                'skip the files up to the current file number
                'download the rest as normal
            Else
                Command1.Caption = "Resume"
                Call SaveSetting(App.EXEName, "Settings", StartSpot & "Start", Val(Text1.Text))
                Call SaveSetting(App.EXEName, "Settings", StartSpot & "End", Val(Text3.Text))
                MsgBox "Program will Stop after the current file is finished"
                Pausit = True
            End If
        End If
End Sub

Private Sub Command10_Click()
    For a = 0 To List1.ListCount - 1
        If Right(List1.List(a), 3) = "NEW" Then
            List1.Selected(a) = True
            List2.Selected(a) = True
        Else
            List1.Selected(a) = False
            List2.Selected(a) = False
        End If
    Next a
End Sub

Private Sub Command11_Click()
    For a = 0 To List1.ListCount - 1
        List1.Selected(a) = True
        List2.Selected(a) = True
    Next a
End Sub

Private Sub Command12_Click()
Frame4.Visible = False
cmdGo.Enabled = True
End Sub

Private Sub Command13_Click()
On Error Resume Next
Screen.MousePointer = 11
Command1.Caption = "Pause"
Command1.Enabled = True
Pausit = False
Frame4.Visible = False
cmdGo.Enabled = True
downnew = True
Label3.Caption = "Please Wait while I download the files you have selected" & vbCrLf & "You can pause the downloads at any time."
    For a = 0 To List1.ListCount - 1
        If List1.Selected(a) = True Then
                If CLng(Left(List1.List(a), 5)) > Val(Text3.Text) Then
                    Open App.Path & "\downloads.txt" For Append As #1
                        Write #1, Left(List1.List(a), 5)
                    Close #1
                End If
            Label3.Caption = "I am now downloading file #" & Left(List1.List(a), 5) & " """ & List2.List(a) & """" & vbCrLf & "Please wait while I download the remaining files." & vbCrLf & "You can pause the downloads at any time."
            DoEvents
            sURL = "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=" & Val(Left(List1.List(a), 5)) & EndString
            mFileName = GetLocalFileNameFromURL(sURL)
            sFileName = App.Path & "\Files\Pages\" & mFileName
            sText = ReadURL(sURL, sType)
                If InStr(1, sText, "/vb/scripts/ShowZip.asp") > 0 Then
                    Parts = Split(sText, "/vb/scripts/ShowZip.asp")
                    DataA = Left(Parts(1), InStr(1, Parts(1), "><") - 2)
                    DataB = "http://www.planetsourcecode.com/vb/scripts/ShowZip.asp" & DataA
                    DF.DownloadFile DataB, App.Path & "\Files\Zips\" & mFileName & ".zip"
                ElseIf InStr(1, sText, "This submission no longer exists in the database") > 0 Then
                ElseIf InStr(1, sText, "This submission was disapproved by the moderator") > 0 Then
                ElseIf InStr(1, sText, "The author of this code has deleted it or it has been removed") > 0 Then
                ElseIf InStr(1, sText, "This code has not yet been processed and posted to the public") > 0 Then
                Else
                    If InStr(1, sText, "/vb/scripts/ShowCodeAsText.asp") > 0 Then
                        sURL = "http://www.planetsourcecode.com/vb/scripts/ShowCodeAsText.asp?txtCodeId=" & Val(Left(List1.List(a), 5)) & "&amp;lngWId=1"
                        sText = ReadURL(sURL, sType)
                        sText = Replace(sText, Chr(13), "")
                        sText = Replace(sText, Chr(10), vbCrLf)
                    End If
                    Call WriteFile(sFileName, sText)
                    Text2.Text = Val(Text2.Text) + 1
                End If
            File1.Refresh
            File2.Refresh
                DoEvents
                    For b = 0 To Val(Text4.Text) * 2   'To 0 Step 1
                        Label3.Caption = "I have finished with file " & vbCrLf & List2.List(a) & vbCrLf & "Pausing For " & (5 - (b \ 2)) & " Seconds"
                        Label3.Refresh
                        DoEvents
                        If Pausit = True Then Exit For
                        Sleep 500
                    Next b
                If Pausit = True Then Exit For
            txtURL.Text = Val(txtURL.Text) + 1
            List1.Selected(a) = False
        End If
    Next a
Screen.MousePointer = 0
Label3.Caption = "I have finished with your downloads."
downnew = False
End Sub

Private Sub Command14_Click()
    For a = 0 To List1.ListCount - 1
        If Right(List1.List(a), 6) = "UPDATE" Then
            List1.Selected(a) = True
            List2.Selected(a) = True
        Else
            List1.Selected(a) = False
            List2.Selected(a) = False
        End If
    Next a
End Sub

Private Sub Command15_Click()
File1.Path = CD.BrowseForFolder("Select A Destination", "")
End Sub

Private Sub Command16_Click()
    WebBrowser1.Navigate "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=on&cmSearch=Search&optSort=Alphabetical&txtCriteria=" & Replace(Text5.Text, " ", "+") & "&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=50&chkCodeTypeZip=on&chkCodeTypeText=on&chkCodeTypeArticle=on&chkCodeDifficulty=1%2C+2%2C+3%2C+4" & EndString '&lngWId=1"
End Sub

Private Sub Command17_Click()
Dim inetPage As String
    Frame6.Visible = False
    Me.Height = 5460
    inetPage = ReadURL(WebBrowser1.LocationURL, "")
    FindFiles (inetPage)
End Sub

Private Sub Command18_Click()
        If WebBrowser1.LocationURL = "http:///" Then
            WebBrowser1.Stop  '.Navigate "http://www.planetsourcecode.com/vb/scripts/search.asp?" & Right(EndString, Len(EndString) - 1)
        End If
    Frame6.Visible = True
    Me.Height = 9165
End Sub

Private Sub Command19_Click()
    Frame6.Visible = False
    Me.Height = 5460
End Sub

Private Sub Command2_Click()
Dim inetPage As String
Dim Parts() As String
Dim Parts2() As String
Label3.Caption = "Downloading requested page please wait.": Label3.Refresh
Screen.MousePointer = 11
Dim highCount As Long
inetPage = ReadURL("http://www.planetsourcecode.com/vb/scripts/browsecategoryorsearchresults.asp?grpcategories=-1&optsort=datedescending&txtmaxnumberofentriesperpage=50&blnnewestcode=true&blnresetallvariables=true" & EndString & "&b1=go&optsort=alphabetical", "")
Parts = Split(inetPage, "/vb/scripts/ShowCode.asp?txtCodeId=")
highCount = 0
    For a = 1 To UBound(Parts) - 1
        Parts2 = Split(Parts(a), "&lngWId=")
            If Len(Parts2(0)) <= 5 Then
                If Val(Parts2(0)) > highCount Then
                highCount = Val(Parts2(0))
                End If
            End If
    Next a
Text3.Text = highCount
Screen.MousePointer = 0
Call SaveSetting(App.EXEName, "Settings", StartSpot & "End", Val(Text3.Text))
Label3.Caption = "The file number to stop downloading at has been updated to the latest code posted to PSC."
End Sub

Private Sub Command20_Click()
Dim Data() As String
Dim Data2 As String
Dim NumPages As String
Frame6.Visible = False
Me.Height = 5460
downByAuth = True
Dim inetPage As String
    Label3.Caption = "Downloading first page of postings by this author please wait.": Label3.Refresh
    inetPage = ReadURL("http://www.planetsourcecode.com" & SearchString, "")
    Data = Split(inetPage, "Page 1 of ")
    NumPages = Left(Data(1), InStr(1, Data(1), " ") - 1)
    FindFiles (inetPage)
    curPageNum = 1
    DoEvents
    Command11_Click
    Command13_Click
        For a = 2 To CLng(NumPages)
            curPageNum = a
            Label3.Caption = "Downloading page #" & a & " of postings by this author please wait.": Label3.Refresh
            inetPage = ReadURL("http://www.planetsourcecode.com" & SearchString & "&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=" & ((a - 1) * 50) + 1 & "&intLastRecordOnPage=" & a * 50 & "&intMaxNumberOfEntriesPerPage=50&intLastRecordInRecordset=" & nNumFils & "&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=&cmdGoToPage=" & a & "&lngMaxNumberOfEntriesPerPage=50", "")
            FindFiles (inetPage)
            DoEvents
            Command11_Click
            Command13_Click
        Next a
End Sub

Private Sub Command21_Click(Index As Integer)
'just the simple options to move around at psc
    Select Case Index
        Case 0
            WebBrowser1.GoBack
        Case 1
            WebBrowser1.GoForward
        Case 2
            WebBrowser1.Stop
        Case 3
            WebBrowser1.Navigate "http://www.planetsourcecode.com/vb/scripts/search.asp?" & Right(EndString, Len(EndString) - 1)
    End Select
End Sub

Private Sub Command22_Click()
    For a = 0 To List1.ListCount - 1
        List1.Selected(a) = False
        List2.Selected(a) = False
    Next a
End Sub

Private Sub Command23_Click()
    PSCAdvancedSearch.Show
End Sub

Private Sub Command24_Click()
Dim Data() As String
Dim Data2 As String
Dim NumPages As String
Frame6.Visible = False
Me.Height = 5460
Dim a As Long
Dim Parts() As String
Dim preFix As String
Dim sufFix As String
Dim Found As Boolean
Dim inetPage As String
    Label3.Caption = "Downloading first page of results please wait.": Label3.Refresh
    inetPage = ReadURL(WebBrowser1.LocationURL, "")
    Parts = Split(WebBrowser1.LocationURL, "&")
    For a = 0 To UBound(Parts)
        If InStr(1, Parts(a), "&cmdGoToPage=") > 0 Then
            Found = True
        Else
            If Found Then
                If sufFix = "" Then
                    sufFix = Parts(a)
                Else
                    sufFix = sufFix & "&" & Parts(a)
                End If
            Else
                If preFix = "" Then
                    preFix = Parts(a)
                Else
                    preFix = preFix & "&" & Parts(a)
                End If
            End If
        End If
    Next a
    Data = Split(inetPage, "Page 1 of ")
    NumPages = Left(Data(1), InStr(1, Data(1), " ") - 1)
    FindFiles (inetPage)
    curPageNum = 1
    DoEvents
    Command11_Click
    Command13_Click
        For a = 2 To CLng(NumPages)
            curPageNum = a
            Label3.Caption = "Downloading page #" & a & " of " & NumPages & " in results.": Label3.Refresh
            inetPage = ReadURL(preFix & "&cmdGoToPage=" & a & sufFix, "")
            FindFiles (inetPage)
            DoEvents
            Command11_Click
            Command13_Click
        Next a
End Sub

Private Sub Command3_Click()
    Load frmMain
    frmMain.cmdGO_Click
    frmMain.Show  'vbModal, Me
    File1.Refresh
    Unload frmMain
End Sub

Private Sub Command4_Click()
Dim desPath As String
desPath = CD.BrowseForFolder("Select A Destination", "")
    If desPath <> "" Then
        For a = 0 To File1.ListCount - 1
            FileCopy File1.Path & "\" & File1.List(a), desPath & "\" & File1.List(a)
            Kill File1.Path & "\" & File1.List(a)
        Next a
    End If
File1.Refresh
End Sub

Private Sub Command5_Click()
Set CD = Nothing
Unload Me
End Sub

Private Sub Command6_Click()
Dim desPath As String
desPath = CD.BrowseForFolder("Select A Destination", "")
    If desPath <> "" Then
        For a = 0 To File2.ListCount - 1
            FileCopy File2.Path & "\" & File2.List(a), desPath & "\" & File2.List(a)
            Kill File2.Path & "\" & File2.List(a)
        Next a
    End If
File2.Refresh
End Sub

Private Sub Command7_Click() 'Rename html Files
On Error Resume Next
Dim Data As String
Dim f As Long
'Dim g As Long
'g = FreeFile
'Open App.Path & "\TestFiles.txt" For Output As #g
'Dim Lines() As String
Dim Parts() As String
f = FreeFile
    For a = 0 To File2.ListCount - 1
        Open File2.Path & "\" & File2.List(a) For Input As #f
            Data = Input(LOF(f), f)
        Close #f
        If InStr(1, Data, "<xmp>") > 0 And InStr(1, Data, "Name") > 0 Then
            Parts = Split(Data, "Name: ")
            Parts(1) = Left(Parts(1), InStr(1, Parts(1), Chr(10)) - 1)
            Parts(1) = Replace(Parts(1), " ", "")
            Parts(1) = Replace(Parts(1), ".", "")
            Parts(1) = Replace(Parts(1), Chr(13), "")
            Parts(1) = SafeFileName(Parts(1), "txt")
'            Print #g, File2.Path & "\" & File2.List(a) & " -To- " & Parts(1)   ' & ".txt"
           Name File2.Path & "\" & File2.List(a) As File2.Path & "\" & Parts(1)
'            If Err Then
'                Err.Clear
'                Name File2.Path & "\" & File2.List(a) As File2.Path & "\" & Parts(1)
'            End If
        ElseIf InStr(1, Data, "<HTML>") > 0 Then
            Parts = Split(Data, "<title>")
            Parts(1) = Left(Parts(1), InStr(1, Parts(1), "</") - 1)
            Parts(1) = Replace(Parts(1), " ", "")
            Parts(1) = Replace(Parts(1), ".", "")
            Parts(1) = SafeFileName(Parts(1), "html")
'            Print #g, File2.Path & "\" & File2.List(a) & " -To- " & Parts(1)
            Name File2.Path & "\" & File2.List(a) As File2.Path & "\" & Parts(1)
        End If
    Next a
'Close #g
File2.Refresh
End Sub

Private Sub Command8_Click()
Dim Data As String
Dim Parts() As String
Dim Parts2() As String
Data = GetSetting(App.EXEName, "Settings", StartSpot & "DownList", "")
    If Data <> "" Then
        Parts = Split(Data, "::")
            For a = 1 To UBound(Parts)
                Parts2 = Split(Parts(a), "*")
                ListboxAddItem List1, Parts2(0)
                ListboxAddItem List2, Parts2(1)
            Next a
        Call SaveSetting(App.EXEName, "Settings", StartSpot & "DownList", "")
        Frame4.Visible = True
        cmdGo.Enabled = False
        Label3.Caption = "These are the files you had selected in an eairler session."
    Else
        Text1.Text = GetSetting(App.EXEName, "Settings", StartSpot & "Start", "0")
        Text3.Text = GetSetting(App.EXEName, "Settings", StartSpot & "End", "0")
        Label3.Caption = "Your start and stop points have been revised to where you paused downloading in an eairler session."
    End If
Command1.Enabled = True
End Sub

Private Sub Command9_Click() 'list the last 50 postings to psc and count todays postings
Dim inetPage As String
    Label3.Caption = "Please wait while I list the latest 50 postings to PSC.": Label3.Refresh
    inetPage = ReadURL("http://www.planetsourcecode.com/vb/scripts/browsecategoryorsearchresults.asp?grpcategories=-1&optsort=datedescending&txtmaxnumberofentriesperpage=50&blnnewestcode=true&blnresetallvariables=true" & EndString & "&b1=go&optsort=alphabetical", "")
    FindFiles (inetPage)
    Label3.Caption = "Finished with your list you can download all or any part of them. Select the files you want and click " & """" & "Download Selected" & "."
End Sub

Private Sub FindFiles(Data As String)
Dim Parts() As String
Dim Parts2() As String
Dim a As Long
Dim b As Long
Dim aTag As String
Dim codName As String
Screen.MousePointer = 11
List1.Clear
List2.Clear
'Main td for content
If InStr(1, Data, "Main td for content") > 0 Then
    Data = Right(Data, Len(Data) - InStr(1, Data, "Main td for content"))
End If
Parts = Split(Data, "/vb/scripts/ShowCode.asp?txtCodeId=")
    For a = 1 To UBound(Parts) '- 1
        aTag = "OLD"
            If InStr(1, Parts(a), "_top") > 0 Then aTag = "UPDATE"
        Parts2 = Split(Parts(a), "&lngWId=")
            If Len(Parts2(0)) <= 5 Then
                If Val(Parts2(0)) > Val(Text3.Text) Then aTag = "NEW"
                If List1.ListCount = 0 Then
                    List1.AddItem Parts2(0) & "-" & aTag
                    codName = Left(Parts(a), InStr(1, Parts(a), "</a>", vbTextCompare) - 1)
'                    codName = Right(codName, Len(codName) - InStr(1, codName, ">"))
                    codName = Right(codName, Len(codName) - (InStr(1, codName, "alt=") + 4))
                    ListboxAddItem List2, codName
                Else
                    For b = 0 To List1.ListCount - 1
                        If Parts2(0) = Left(List1.List(b), InStr(1, List1.List(b), "-") - 1) Then
                            Exit For
                        ElseIf b = List1.ListCount - 1 Then 'Parts2(0) <> List1.List(b) And
                             If Not InStr(1, Parts(a), "img align=") > 0 Then
                                    ListboxAddItem List1, Parts2(0) & "-" & aTag
                                    codName = Left(Parts(a), InStr(1, Parts(a), "</a>", vbTextCompare) - 1)
                                    codName = Right(codName, Len(codName) - InStr(1, codName, ">"))
                                    'If InStr(1, codName, "</a>") > 0 Then codName = Left(codName, InStr(1, codName, "</a>") - 1)
                                    If codName = "" Then codName = "Name not found"
                                    ListboxAddItem List2, codName
                            End If
                        End If
                    Next b
                End If
            End If
    Next a
    
Frame4.Visible = True
cmdGo.Enabled = False
'the number left is the latest code artical
Screen.MousePointer = 0
Erase Parts
Erase Parts2
End Sub

Private Sub Form_Load()
    Me.Height = 5460
    Frame6.Top = 30
    Combo1.ListIndex = 0
    If DoesDirectoryExist(App.Path & "\Files") = False Then MkDir App.Path & "\Files"
    If DoesDirectoryExist(App.Path & "\Files\Zips") = False Then MkDir App.Path & "\Files\Zips"
    If DoesDirectoryExist(App.Path & "\Files\Pages") = False Then MkDir App.Path & "\Files\Pages"
    File1.Path = App.Path & "\Files\Zips"
    File2.Path = App.Path & "\Files\Pages"
    Set CD = New cDlg
    GetURLS
    Me.Show
    DoEvents
    Text1.Text = GetSetting(App.EXEName, "Settings", StartSpot & "Start", "0")
    Text3.Text = GetSetting(App.EXEName, "Settings", StartSpot & "End", "0")
    If Val(Text1.Text) < Val(Text3.Text) Then Command1.Enabled = True
    SetFont
End Sub

Private Sub SetFont()
'On Error Resume Next
'    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull 'smallest
'    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull 'small
'    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull 'med
'    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull 'large
'    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull 'largest
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CD = Nothing
    Unload frmMain
End Sub

Private Sub List1_Click()
On Error Resume Next
    List2.Selected(List1.ListIndex) = List1.Selected(List1.ListIndex)
    List2.TopIndex = List1.TopIndex
End Sub

Private Sub List2_Click()
On Error Resume Next
    List1.Selected(List2.ListIndex) = List2.Selected(List2.ListIndex)
    List1.TopIndex = List2.TopIndex
End Sub
Private Sub List1_Scroll()
    List2.TopIndex = List1.TopIndex
End Sub

Private Sub List2_Scroll()
    List1.TopIndex = List2.TopIndex
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command16_Click
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Dim Data1() As String
Dim Data2 As String
Dim NumFiles As String
Dim ArthName As String
Dim inetPage As String
    Command20.Enabled = False
    Command17.Enabled = True
    Command24.Enabled = True
    inetPage = ReadURL(WebBrowser1.LocationURL, "")
        If InStr(1, inetPage, "submission(s) by this author") > 0 Then
            Data1 = Split(inetPage, " submission(s) by this author")
            Data2 = Right(Data1(0), Len(Data1(0)) - (InStrRev(Data1(0), "<a href=") + 8))
            SearchString = Left(Data2, (InStr(1, Data2, "txtMaxNumberOfEntriesPerPage=") - 1)) & "txtMaxNumberOfEntriesPerPage=50"
            NumFiles = Right(Data2, Len(Data2) - (InStrRev(Data2, "Other") + 5))
            nNumFils = CLng(NumFiles) + 1
            Data1 = Split(Data2, "AuthorName=")
            ArthName = Left(Data1(1), InStr(1, Data1(1), "&") - 1)
            Command20.Caption = "Download all " & CLng(NumFiles) + 1 & " files from " & ArthName & "?"
            Command20.Enabled = True
            Command17.Enabled = False
            Command24.Enabled = False
        End If
End Sub

