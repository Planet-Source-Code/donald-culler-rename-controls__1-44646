VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmShowExamples 
   Caption         =   "                                        Examples"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7646
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
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frmShowExamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
        Unload Me
End Sub


Private Sub Form_Load()

'    ListView1.ListItems.Remove ListView1.SelectedItem.Index 'Delete The Selected Item
    
    ListView1.View = lvwReport 'Set The Listview1 View So we Can See Our Columns/Headers
    
'   Dim colX As ColumnHeader ' Declare variable.
'   Dim intX As Integer ' Counter variable.
'   For intX = 1 To 3
'   Set colX = ListView1.ColumnHeaders.Add()
'   colX.Text = "Field " & intX
'   colX.Width = ListView1.Width / 3
'   Next intX

    'This Code is needed another way
    
    ListView1.ColumnHeaders.Add , , "Control Type", ListView1.Width / 3 'Add a column
    ListView1.ColumnHeaders.Add , , "Prefix", ListView1.Width / 3 'Add a column
    ListView1.ColumnHeaders.Add , , "Example", ListView1.Width / 3 'Add a column
    
    'End Of Needed Code
    
    ListView1.ListItems.Clear ' Clears all ListItems.

 'For i = 1 To ListView1.ListItems.Count
 '       Set ListView1.SelectedItem = ListView1.ListItems(i)
 '       ListView1.SelectedItem.Text = txtNewValue
 '   Next i

    Dim Col1 'Remember Col1
    Dim Col2 'Remember Col2
    Dim Col3 'remember Col3
    Dim DataInput
        'Adding To List Code
        
        
 Open "C:\1listview\project\test.txt" For Input As #1
 Do Until EOF(1)
 Line Input #1, DataInput
Strip DataInput, Col1, Col2, Col3
        Dim lst As ListItem 'Set lst as a listitem
        Set lst = ListView1.ListItems.Add(, , Col1) 'this allways adds to the 1st column , this lines adds the Col to the 1st Column
        lst.SubItems(1) = Col2 'this allways adds to the second column , this lines adds the Col2 to the 2nd Column
        lst.SubItems(2) = Col3 'this allways adds to the third column , this lines adds the Col3 to theb 3rd Column
        'End Of Adding To List Code
    Loop
    Close #1
    
End Sub





Private Sub ListView1_DblClick()
    On Error Resume Next ' resume The Next line On a Error
    MsgBox "Control Type   : " + ListView1.SelectedItem.Text + vbCrLf + "Prefix   : " + ListView1.SelectedItem.SubItems(1) + vbCrLf + "Control Example   : " + ListView1.SelectedItem.SubItems(2) 'Make The Msgbox
End Sub

Sub Strip(Main, AA, BB, CC)
Dim a, b, c


Main = Trim(Main)
If Main = "" Then Exit Sub
a = InStr(Main, ";")
AA = Trim(Left(Main, a - 1))

b = InStr(a + 1, Main, ";")
BB = Trim(Mid(Main, a + 1, b - (a + 1)))

c = Len(Main)
CC = Trim(Mid(Main, b + 1, c - (b)))



End Sub
