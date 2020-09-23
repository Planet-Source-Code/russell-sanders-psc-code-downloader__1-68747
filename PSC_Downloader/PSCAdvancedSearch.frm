VERSION 5.00
Begin VB.Form PSCAdvancedSearch 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PSC Advanced  Search"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "PSC Advanced Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6555
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2685
      Begin VB.CommandButton Command2 
         Caption         =   "Hide"
         Height          =   255
         Left            =   1410
         TabIndex        =   23
         Top             =   6180
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   810
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Zip Files"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   1410
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Copy and paste"
         Height          =   225
         Index           =   1
         Left            =   1110
         TabIndex        =   14
         Top             =   1410
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Articles/Tuts"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   13
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "3rd Party Rev"
         Height          =   225
         Index           =   3
         Left            =   1290
         TabIndex        =   12
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Unranked"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2310
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Beginner"
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   2580
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Intermediate"
         Height          =   225
         Index           =   6
         Left            =   1290
         TabIndex        =   9
         Top             =   2310
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Advanced"
         Height          =   225
         Index           =   7
         Left            =   1290
         TabIndex        =   8
         Top             =   2580
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Scans actual code contents(takes longer)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   90
         TabIndex        =   7
         Top             =   3210
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "Alphabetical Order"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   4110
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "Newest submissions first"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   4350
         Width           =   2025
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "Oldest submissions first"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   4620
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "Most Popular submissions first"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   3
         Top             =   4920
         Width           =   2475
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   6180
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Text            =   "50"
         Top             =   5850
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Name Description Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   525
         Index           =   1
         Left            =   90
         TabIndex        =   22
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Code Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   1080
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Code Difficulty Level:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   1980
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Thorough Search:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   19
         Top             =   2880
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Display In:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   18
         Top             =   3750
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "Max # of entries to view (per page):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Index           =   6
         Left            =   90
         TabIndex        =   17
         Top             =   5250
         Width           =   2505
      End
   End
End
Attribute VB_Name = "PSCAdvancedSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Most popular, type no third party rev, dif all, rank best to worst, through
'http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?
'optSort=CountDescending
'&cmSearch=Search&txtCriteria=subclass
'&blnResetAllVariables=TRUE
'&txtMaxNumberOfEntriesPerPage=50
'&chkCodeTypeZip=on
'&chkCodeTypeText=on
'&chkCodeTypeArticle=on
'&chkCodeDifficulty=1%2C+2%2C+3%2C+4
'&chkThoroughSearch=on
'&lngWId=1

'oldest first
'http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?
'chkCode3rdPartyReview=on
'&optSort=DateAscending
'&cmSearch=Search
'&txtCriteria=subclass
'&blnResetAllVariables=TRUE
'&txtMaxNumberOfEntriesPerPage=50
'&chkCodeTypeZip=on
'&chkCodeTypeText=on
'&chkCodeTypeArticle=on
'&chkCodeDifficulty=2%2C+3%2C+4
'&chkThoroughSearch=on
'&lngWId=1

'alphabetical
'http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=on&optSort=Alphabetical&cmSearch=Search&txtCriteria=subclass&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=50&chkCodeTypeZip=on&chkCodeTypeText=on&chkCodeTypeArticle=on&chkCodeDifficulty=1%2C+2%2C+3%2C+4&chkThoroughSearch=on&lngWId=1

'no beginer newest first
'http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=on&optSort=DateDescending&cmSearch=Search&txtCriteria=subclass&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=50&chkCodeTypeZip=on&chkCodeTypeText=on&chkCodeTypeArticle=on
'&chkCodeDifficulty=1%2C+3%2C+4&lngWId=1

'http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=on
'&optSort=Alphabetical&cmSearch=Search&txtCriteria=Subclass&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=10&chkCodeTypeZip=on&chkCodeTypeText=on&chkCodeTypeArticle=on
'&chkCodeDifficulty=1%2C+2%2C+3%2C+4&lngWId=1


'http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?chkCode3rdPartyReview=on&optSort=CountDescending&cmSearch=Search&txtCriteria=subclass&blnResetAllVariables=TRUE&txtMaxNumberOfEntriesPerPage=10&chkCodeTypeZip=on&chkCodeTypeText=on&chkCodeTypeArticle=on&chkCodeDifficulty=1%2C+2%2C+3%2C+4&lngWId=1
Private Sub Command1_Click(Index As Integer)
Dim Baseadd As String
Dim difi As String
Baseadd = ""
difi = ""
Baseadd = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?" & "&blnResetAllVariables=TRUE"
    If Option1(0).Value = True Then
        Baseadd = Baseadd & "&optSort=Alphabetical"
    ElseIf Option1(1).Value = True Then
        Baseadd = Baseadd & "&optSort=DateDescending"
    ElseIf Option1(2).Value = True Then
        Baseadd = Baseadd & "&optSort=DateAscending"
    ElseIf Option1(3).Value = True Then
        Baseadd = Baseadd & "&optSort=CountDescending"
    End If
Baseadd = Baseadd & "&cmSearch=Search&txtCriteria=" & Text1(0).Text '& "&blnResetAllVariables=TRUE"
    If Check1(0).Value = 1 Then Baseadd = Baseadd & "&chkCodeTypeZip=on"
    If Check1(1).Value = 1 Then Baseadd = Baseadd & "&chkCodeTypeText=on"
    If Check1(2).Value = 1 Then Baseadd = Baseadd & "&chkCodeTypeArticle=on"
    If Check1(3).Value = 1 Then Baseadd = Baseadd & "&chkCode3rdPartyReview=on"
    
    If Check1(4).Value = 1 Then
        difi = "1"
    End If
    If Check1(5).Value = 1 Then
        If Len(difi) > 0 Then
            difi = difi & "%2c+2"
        Else
            difi = "2"
        End If
    End If
    If Check1(6).Value = 1 Then
        If Len(difi) > 0 Then
            difi = difi & "%2c+3"
        Else
            difi = "3"
        End If
    End If
    If Check1(7).Value = 1 Then
        If Len(difi) > 0 Then
            difi = difi & "%2c+4"
        Else
            difi = "4"
        End If
    End If
Baseadd = Baseadd & "&chkCodeDifficulty=" & difi
    If Check1(8).Value = 1 Then Baseadd = Baseadd & "&chkThoroughSearch=on"
Baseadd = Baseadd & "&txtMaxNumberOfEntriesPerPage=" & Val(Text1(1).Text) & frmInet.EndString
frmInet.WebBrowser1.Navigate Baseadd
End Sub
