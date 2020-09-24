VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTable 
   BorderStyle     =   0  'None
   Caption         =   "frmTable"
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9810
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView xmlListView 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3625
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776960
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub xmlListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

xmlListView.SortKey = ColumnHeader.Index - 1
If xmlListView.SortOrder = lvwAscending Then
   xmlListView.SortOrder = lvwDescending
Else
   xmlListView.SortOrder = lvwAscending
End If
xmlListView.Sorted = True

End Sub

Private Sub xmlListView_DblClick()
Unload Me
End Sub

Private Sub xmlListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Me
End Sub
