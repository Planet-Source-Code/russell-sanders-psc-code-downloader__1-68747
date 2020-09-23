VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renaming The Files."
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Tag             =   "!"
   Begin VB.Frame fraZips 
      Caption         =   "Files Ma&tching Criteria"
      Height          =   5475
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4035
      Begin VB.FileListBox filZips 
         Height          =   5160
         Hidden          =   -1  'True
         Left            =   120
         MultiSelect     =   2  'Extended
         Pattern         =   "*.zip"
         System          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   3825
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "DO IT&!"
      Height          =   345
      Left            =   2490
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FRAMEWORK by Crock
Option Explicit
Private ZipName      As String
Private FilePath     As String
Private ProgName     As String
Private strTitle As String
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub cmdGO_Click()
On Error Resume Next
  Dim x           As Integer
  Dim strFilePath As String
  Dim a           As Control
    For x = filZips.ListCount - 1 To 0 Step -1
    strFilePath = filZips.Path & "\" & filZips.List(x)
    ZipName = Right$(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\", , vbTextCompare))
    FilePath = Left$(strFilePath, Len(strFilePath) - Len(ZipName))
    uZip
    ZipName = strTitle
    If ZipName <> "NO README FOUND!" Then
      ZipName = SafeFileName(ZipName, "zip")
      SetAttr strFilePath, GetAttr(strFilePath) And Not vbReadOnly  '
      Name strFilePath As FilePath & ZipName
        If Err.Number = 58 Then
errfix:
            Err.Clear
            ZipName = Left(ZipName, Len(ZipName) - 4) & "2.zip"
            Name strFilePath As FilePath & ZipName
                If Err.Number = 58 Then
                    GoTo errfix
                Else
                    If Err.Number <> 0 Then MsgBox Err.Number & Err.Description
                End If
        Else
            If Err.Number <> 0 Then MsgBox Err.Number & Err.Description
        End If
    End If
  Next
Me.Caption = "  DONE!"
'Store The Files To The desktop
 ' Call MoveFiles
'  Unload Me
End Sub


Private Sub Form_Load()
filZips.Path = frmInet.File1.Path
End Sub

Private Sub uZip()
  Dim Crit     As Variant
  Dim i        As Long
On Error Resume Next
  strTitle = "NO README FOUND!"
  Screen.MousePointer = vbHourglass   ' Change mouse pointer to hourglass.
  Cls
  uZipInfo = vbNullString
  uZipNumber = 0   ' Holds The Number Of Zip Files
  uQuiet = 2             ' 2 = No Messages, 1 = Less, 0 = All
  uWriteStdOut = 1       ' 1 = Write To Stdout, Else 0
  uExtractList = 1       ' 0 = Extract, 1 = List Contents
  uDisplayComment = 0    ' 1 = Display Zip File Comment, Else 0
  uCaseSensitivity = 1   ' 1 = Case Insensitivity, 0 = Case Sensitivity
  uZipFileName = FilePath & ZipName        ' The Zip File Name
  Crit = Array("@PSC_ReadMe*.txt")
  For i = 0 To UBound(Crit)
    If LenB(uZipInfo) = 0 Then ' Try next criteria   ' Try next criteria
      uZipNames.uzFiles(0) = Crit(i)
      uNumberFiles = 1
      VBUnZip32
    End If
  Next i
  If LenB(uZipInfo) Then
    strTitle = Left$(uZipInfo, InStr(1, uZipInfo, vbNewLine, vbTextCompare) - 1)
        If InStr(1, strTitle, "Title: ") > 0 Then
            strTitle = Replace$(strTitle, "Title: ", vbNullString, 1, 1, vbTextCompare)
        Else
            strTitle = "NO README FOUND!"
        End If
  Else 'LENB(UZIPINFO) = FALSE/0
      strTitle = "NO README FOUND!"
  End If
  Screen.MousePointer = vbDefault ' Return mouse pointer to normal.
  ProgName = strTitle

End Sub

