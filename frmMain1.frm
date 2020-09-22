VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "RenameControls"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   480
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   3600
      Width           =   8175
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   6300
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11721
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "4/9/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "1:14 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2760
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":077E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0890
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":09A2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0AB4
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0BC6
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0CD8
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      X1              =   240
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Control Name"
      Height          =   195
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Control"
      Height          =   195
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Save To:"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Form To Work With"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1395
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   240
      X2              =   4440
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   4440
      X2              =   4440
      Y1              =   480
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   3240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Show Examples ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_Load()
    
   Me.Width = Screen.Width * 0.89  ' Set width of form.
   Me.Height = Screen.Height * 0.89  ' Set height of form.
   Me.Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Me.Top = (Screen.Height - Height) / 2   ' Center form vertically.
    
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add , , "Control Type", ListView1.Width / 2 'Add a column
    ListView1.ColumnHeaders.Add , , "Control Name", ListView1.Width / 2 'Add a column
  
Label7.Caption = ""
Label8.Caption = ""
Label1.Visible = False
Label2.Visible = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub
Sub StripControl(FormData, ControlsName, Controls)
Dim a, b
FormData = Trim(FormData)
a = InStr(FormData, " ")
b = InStrRev(FormData, " ")
ControlsName = Mid(FormData, a + 1, (b - a))
Controls = Mid(FormData, b + 1, Len(FormData) - b)

End Sub

Private Sub ListView1_DblClick()
On Error Resume Next ' resume The Next line On a Error
Dim Message, Title, Default, NewValue, i, a, b, ControlName, Control, CurrentName, CurrentItem
  CurrentName = ListView1.SelectedItem.SubItems(1)
For i = 0 To List1.ListCount - 1

If InStr(Trim(List1.List(i)), "BeginProperty") = 0 Then
If InStr(Trim(List1.List(i)), "Begin") = 1 Then
If InStr(List1.List(i), ListView1.SelectedItem.SubItems(1)) Then
For b = i To List1.ListCount - 1
List1.Selected(b) = True 'Selected item to display in main listbox
List1.Selected(b) = False
If InStr(Trim(List1.List(b)), "End") = 1 Then List1.Selected(i) = True: Exit For
Next

Exit For
End If
End If
End If
Next i
Message = "Change Name Of Control on Form"   ' Set prompt.

Duplicate:
Title = "Controls"   ' Set title.
Default = ListView1.SelectedItem.SubItems(1)
' Display message, title, and default value.
NewValue = InputBox(Message, Title, Default)

If NewValue = "" Then List1.Selected(i) = False: Exit Sub
b = 0

For i = 0 To List1.ListCount - 1   'Check if name you want to use already exists
a = InStr(List1.List(i), NewValue & " ")
If a <> 0 Then Exit For
  Next i
If a <> 0 Then Message = NewValue & "****  is a duplicate Name  Enter another Name": GoTo Duplicate

For i = 0 To List1.ListCount - 1   'Change name of the control
a = Replace(List1.List(i), ListView1.SelectedItem.SubItems(1), NewValue)
 List1.List(i) = a
  Next i
    
    For i = 1 To ListView1.ListItems.Count  'Check and reset names in listview
    If ListView1.ListItems.Item(i).SubItems(1) = CurrentName Then
      ListView1.ListItems.Item(i).SubItems(1) = NewValue
      End If
      Next
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
       
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
       
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub
Private Sub mnuToolsOptions_Click()
    frmShowExamples.Show vbModal, Me
End Sub
Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub
Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub
Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub
Private Sub mnuFileSave_Click()
    Dim sFile As String
    Dim i, FileOut
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'set the flags and attributes of the common dialog control
        .Filter = "Visual Basic Form|*.frm"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
'now save
Label8.Caption = sFile
FileOut = FreeFile
Open sFile For Output As FileOut
    For i = 0 To List1.ListCount - 1
        Print #FileOut, List1.List(i)
    Next i
Close #FileOut

End Sub
Private Sub mnuFileClose_Click()
List1.Clear
ListView1.ListItems.Clear ' Clears all ListItems.
Label7.Caption = ""
Label8.Caption = ""
Label1.Visible = False
Label2.Visible = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
End Sub
Private Sub mnuFileOpen_Click()
    Dim sFile As String
    Dim StringData, ControlName, Control, FileIn
    List1.Clear
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        ' set the flags and attributes of the common dialog control
        .Filter = "Visual Basic Form|*.frm"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With

Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Label1.Visible = True
Label2.Visible = True

Label7.Caption = sFile
Label8.Caption = sFile

FileIn = FreeFile
Open sFile For Input As FileIn ' Get Form
    Do Until EOF(FileIn)
        Line Input #FileIn, StringData
        List1.AddItem StringData
    
    If InStr(Trim(StringData), "BeginProperty") = 0 Then
        
    If InStr(Trim(StringData), "Begin") = 1 Then
        
        StripControl StringData, ControlName, Control
    
       Dim lst As ListItem     'Set lst as a listitem
        Set lst = ListView1.ListItems.Add(, , ControlName) 'this allways adds to the 1st column
        lst.SubItems(1) = Control 'this allways adds to the second column
        End If
    End If
    Loop
  
  Close #FileIn

End Sub
