VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub PrintTable(l_rs As Object)

On Error GoTo err_printTable

    Dim l_intCount As Integer

    Dim frm As frmTable
    Set frmTable = New frmTable
    Dim clmXml As ColumnHeader
    Dim Itmx As ListItem

    With frmTable.xmlListView
        .View = lvwReport
        For l_intCount = 0 To l_rs.Fields.Count - 1
            Set clmXml = .ColumnHeaders.Add(, , l_rs.Fields(l_intCount).Name)

        Next

        If l_rs.RecordCount > 0 Then
            l_rs.MoveFirst
            Do While Not l_rs.EOF
                Set Itmx = .ListItems.Add(, , l_rs.Fields(0).Value)
                For l_intCount = 1 To l_rs.Fields.Count - 1

                    Itmx.SubItems(l_intCount) = IIf(IsNull(l_rs.Fields(l_intCount).Value), "Null", l_rs.Fields(l_intCount).Value)
                Next
                l_rs.MoveNext
            Loop
        End If
    End With
    frmTable.Show
err_printTable:
If Err.Number <> 0 Then
    MsgBox Err.Number & ":" & Err.Description, vbCritical, "Unable to print the recordset"
End If

End Sub
