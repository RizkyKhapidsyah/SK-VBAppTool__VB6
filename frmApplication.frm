VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmApplication 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   3420
   ClientTop       =   2055
   ClientWidth     =   4620
   Icon            =   "frmApplication.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   4620
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   210
      TabIndex        =   9
      Top             =   315
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   210
      TabIndex        =   4
      Top             =   630
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search String"
      Height          =   3165
      Left            =   210
      TabIndex        =   3
      Top             =   4515
      Width           =   4215
      Begin VB.CheckBox Check1 
         Caption         =   "Search Sub Folders"
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   210
         Width           =   2745
      End
      Begin VB.TextBox txtString 
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   525
         Width           =   3795
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search For String"
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   840
         Width           =   3795
      End
      Begin RichTextLib.RichTextBox rtfSearch 
         Height          =   1905
         Left            =   210
         TabIndex        =   8
         Top             =   1155
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   3360
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmApplication.frx":030A
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line Count"
      Height          =   1800
      Left            =   210
      TabIndex        =   1
      Top             =   2625
      Width           =   4215
      Begin RichTextLib.RichTextBox txtLineCnt 
         Height          =   960
         Left            =   210
         TabIndex        =   10
         Top             =   735
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   1693
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmApplication.frx":03FF
      End
      Begin VB.CommandButton cmdLineCnt 
         Caption         =   "&Line Count"
         Height          =   330
         Left            =   210
         TabIndex        =   2
         Top             =   315
         Width           =   3795
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7725
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5080
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "8:10 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Select folder with VB files"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   105
      Width           =   3585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     ' Force explicit variable declaration.
Option Compare Text ' Set the string comparison method to Text.

Private Sub cmdLineCnt_Click()
  LineCount
End Sub

Private Sub cmdSearch_Click()
  SearchCode
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
  Caption = App.Title
  mnuAbout.Caption = mnuAbout.Caption & " " & App.Title
  StatusBar.Panels(1).Text = "Ready"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmApplication = Nothing
End Sub

Private Sub mnuAbout_Click()
  frmAbout.ShowForm
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub SearchCode()
  Dim sPath As String
  Dim FSO As FileSystemObject
  Dim oFolder As Folder
  Dim sFiles As String
  Me.MousePointer = vbHourglass
  sPath = Dir1.Path
  rtfSearch.Text = ""
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = FSO.GetFolder(sPath)
  SearchFolders FSO, oFolder, sPath, sFiles
  rtfSearch.Text = sFiles
  Me.MousePointer = vbDefault
End Sub

Private Sub SearchFolders(FSO As FileSystemObject, oFolder As Folder, sPath As String, sFiles As String)
  Dim oFile As File
  Dim oFiles As Files
  Dim oSubFolder As Folder
  Dim oSubFolders As Folders
  Dim x As Integer
  
  'Search each file in this folder
  Set oFiles = oFolder.Files
  For Each oFile In oFiles
    SearchFile FSO, oFile, sPath, sFiles
  Next
  
  'if there is a sub folder, recursivly call SearchFolders and pass in each sub folder
  If Check1.Value = 1 Then
    If oFolder.SubFolders.Count > 0 Then
      Set oSubFolders = oFolder.SubFolders
      For Each oSubFolder In oSubFolders
        SearchFolders FSO, oSubFolder, oSubFolder.Path, sFiles
      Next
    End If
  End If
End Sub

Private Sub SearchFile(FSO As FileSystemObject, FileTemp As File, sPath As String, sFiles As String)
  Dim bFile As Boolean
  Dim sTemp As String
  Dim TextStream As TextStream
  Dim sString As String
  
  'check to see if file type is a vb file
  sTemp = CStr(FileTemp.Name)
  Select Case Right(sTemp, 3)
    Case "dsr"
      bFile = True
    Case "frm"
      bFile = True
    Case "bas"
      bFile = True
    Case "cls"
      bFile = True
    Case "ctl"
      bFile = True
    Case Else
      bFile = False
  End Select
  If bFile = True Then
    Set TextStream = FSO.OpenTextFile(sPath & "\" & sTemp)
    'get all the text of this file
    sString = TextStream.ReadAll
    'search text for search string
    If InStr(1, sString, txtString, vbTextCompare) > 0 Then
       sFiles = sFiles & sTemp & vbCrLf
    End If
    TextStream.Close
  End If
End Sub

Private Sub LineCount()
  Dim sPath As String
  Dim FSO As FileSystemObject
  Dim TextStream As TextStream
  Dim iLineCnt As Long
  Dim iCommentCnt As Long
  Dim iBlankCnt As Long
  Dim iTotalCnt As Long
  Dim F As Folder
  Dim FileTemp As File
  Dim colF As Files
  Dim sTemp As String
  Dim bProcess As Boolean
  Dim sLine As String
  Dim bStart As Boolean
  
  Me.MousePointer = vbHourglass
  sPath = Dir1.Path
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set F = FSO.GetFolder(sPath)
  Set colF = F.Files
  
  'check to see if file type is a vb file
  For Each FileTemp In colF
     bStart = False
     sTemp = CStr(FileTemp.Name)
     Select Case Right(sTemp, 3)
       Case "dsr"
         bProcess = True
       Case "frm"
         bProcess = True
       Case "bas"
         bProcess = True
       Case "cls"
         bProcess = True
       Case "ctl"
         bProcess = True
       Case Else
         bProcess = False
     End Select
     
     'if a vb code file type then
     If bProcess = True Then
       Set TextStream = FSO.OpenTextFile(sPath & "\" & sTemp)
       Do While TextStream.AtEndOfStream = False
          sLine = TextStream.ReadLine
          If InStr(1, sLine, "Attribute VB_Name", vbTextCompare) > 0 Then bStart = True
          If bStart = True Then
            'add one to total count
            iTotalCnt = iTotalCnt + 1
            If Len(Trim(sLine)) = 0 Then
              'if a blank line add one to blank line count
              iBlankCnt = iBlankCnt + 1
            ElseIf Left(Trim(sLine), 1) = "'" Then
              'if a commented line add one to comments count
              iCommentCnt = iCommentCnt + 1
            Else
              'otherwise add one to code line count
              iLineCnt = iLineCnt + 1
            End If
          End If
       Loop
       TextStream.Close
     End If
  Next
  txtLineCnt.Text = "Lines of Code: " & iLineCnt & vbCrLf & "Lines of Comments: " & iCommentCnt & vbCrLf & "Blank Lines: " & iBlankCnt & vbCrLf & "Total Lines in app: " & iTotalCnt
  Me.MousePointer = vbDefault
End Sub
