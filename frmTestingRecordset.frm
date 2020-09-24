VERSION 5.00
Begin VB.Form frmTestingRecordset 
   Caption         =   "Testing PrintRecordset"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PrintRecordset"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmTestingRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPrint_Click()

'The first thing is connect to any database
' Using ADO, RDO, DAO
' For the time being I use ADO

Dim l_cn As ADODB.Connection
Dim l_rsSuppliers As ADODB.Recordset
Dim l_rsProducts As ADODB.Recordset
Dim l_rsCustomers As ADODB.Recordset
Dim ConnectionString As String


' For the time being I am connecting with
' Activex object library 2.0

' You can connect to Oracle , MS SQL
' or to any damn database including XML

'First open Odbc connection  and name it "Test"
'and connect to the database NWIND

Set l_cn = New ADODB.Connection

Set l_rsSuppliers = New ADODB.Recordset
Set l_rsProducts = New ADODB.Recordset
Set l_rsSuppProd = New ADODB.Recordset
Set l_rsCustomers = New ADODB.Recordset

MsgBox "Requires the database to be in C:\Program Files\Microsoft Visual Studio\VB98\nwind.mdb only for this example", vbInformation, "Vijay Phani"

  ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
  "Persist Security Info=False;" & _
  "Data Source=C:\Program Files\Microsoft Visual Studio\VB98\nwind.mdb;" & _
  "JET OLEDB:SFP=True;"

l_cn.Open ConnectionString

l_rsSuppliers.Open "Select * from suppliers", l_cn, adOpenKeyset, adLockOptimistic
l_rsProducts.Open "Select * from Products", l_cn, adOpenKeyset, adLockOptimistic
l_rsCustomers.Open "Select * from Customers", l_cn, adOpenKeyset, adLockOptimistic

' Just Use this to print the contents of the recordsets
Dim prs As PrintRecordset.clsTable
Set prs = New PrintRecordset.clsTable

'Remove the comments if u want to print more tables

prs.PrintTable l_rsCustomers
prs.PrintTable l_rsSuppliers
prs.PrintTable l_rsProducts

' This way you can print the contents of the Recordset
' which can be from any servive provider
MsgBox "1. You can click on the tabs and the recordset will be in sorted order" & vbCrLf & "2. By double click the recordset control the control is unloaded" & vbCrLf & "3. The form can be dragged any where on the screen by holding the left mouse button" & vbCrLf & "4. Drag the tab to place it somewhere in the form" & vbCrLf & "5. Can be used on any database Oracle,foxpro,Sqlserver anydamn database" & vbCrLf & "6. Can be used for Ado,Rdo,Dao", vbInformation, "Vijay Phani"

End Sub
