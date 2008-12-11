VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Pop-Client"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3000
      Width           =   4575
   End
   Begin MSComctlLib.ListView lsvTestView 
      Height          =   2535
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    With lsvTestView
        .ColumnHeaders.Add , , "From"
        .ColumnHeaders.Add , , "Subject"
        .ColumnHeaders.Add , , "Date"
        
        .View = lvwReport
        .ListItems.Add , , "zzA"
        .ListItems.Item(1).ListSubItems.Add , , "ffA"
        .ListItems.Item(1).ListSubItems.Add , , "ddA"
        
        .ListItems.Add , , "ttB"
        .ListItems.Item(2).ListSubItems.Add , , "rrB"
        .ListItems.Item(2).ListSubItems.Add , , "ssB"
    End With
End Sub
Private Sub lsvTestView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDescending(2) As Boolean

    If ColumnHeader.index = 1 Then 'there it starts with 1!
        'MsgBox "first one"
    ElseIf ColumnHeader.index = 2 Then
        'MsgBox "second one"
    End If
    
    With lsvTestView
        .SortKey = ColumnHeader.index - 1 'there it starts with 0!
        .SortOrder = Switch(blnDescending(ColumnHeader.index - 1), lvwDescending, True, lvwAscending)
        .Sorted = True
        blnDescending(ColumnHeader.index - 1) = Not blnDescending(ColumnHeader.index - 1)
    End With
    
    
End Sub
