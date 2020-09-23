Attribute VB_Name = "code"
Option Explicit



Enum IEstate
    START_PAGE = 0
    SOFTWARE_PAGE = 1
End Enum
Public IE_state As IEstate

'
'--prevent duplicates in a listbox
'
Function isDuplicate(Listbox As Listbox, strval As String) As Boolean
 
 Dim lcnt    As Long
  
 For lcnt = 0 To Listbox.ListCount - 1
    If strval = Listbox.List(lcnt) Then
       isDuplicate = True
       Exit Function
    End If
 Next lcnt
 
End Function
'
'-- check to see if item is array or array is initialized
'
Function IsArray(varArray As Variant) As Boolean
Dim upper As Integer
On Error Resume Next
 
  upper = UBound(varArray)
  
  If Err.Number Then
     If Err.Number = 9 Then
       IsArray = False
     End If
  Else
     IsArray = True
  End If

End Function


