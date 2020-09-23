Attribute VB_Name = "Mdl_XPDatePicker"
Option Explicit

Public ResultDate As Long

Public Function GetDateSys() As String
    GetDateSys = DateSys(Date)
End Function

Public Function DateSys(StrDate As String) As String
 On Error GoTo errsystems
        DateSys = Format(StrDate, "yyyy") & "/" & Format(StrDate, "mm") & "/" & Format(StrDate, "dd")
errsystems:
End Function

Sub main()
    Load MyForm
End Sub
