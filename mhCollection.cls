VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mhCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' enhanced collection

Private colldata As Collection
Private collkeys As Collection

Private Sub Class_Initialize()
Set colldata = New Collection
Set collkeys = New Collection

End Sub

Private Sub Class_Terminate()
 ClearCollection colldata
 ClearCollection collkeys
 Set colldata = Nothing
 Set collkeys = Nothing
 
End Sub
Private Sub ClearCollection(ByRef rcolCollection As Collection)
    Dim lngIndex As Long
    For lngIndex = 1 To rcolCollection.Count
        rcolCollection.Remove 1
    Next lngIndex
End Sub
Public Sub Clear()
 ClearCollection colldata
 ClearCollection collkeys
End Sub

Public Function Exists(ByVal key As Variant) As Boolean

On Error Resume Next
Dim arg As Variant
arg = colldata.item(key)
If (IsEmpty(arg)) Then
 Exists = False
Else
 Exists = True
End If
End Function

Public Property Get Count() As Long
 Count = colldata.Count
 
End Property

Public Function item(ByVal idx As Variant) As Variant
Dim outval As Variant
If (Exists(idx) = False) Then
 item = CVar("")
 Exit Function
End If
outval = colldata(idx)
If (IsObject(outval)) Then
 Set item = outval
Else
    item = outval
End If
End Function

Public Function key(ByVal idx As Variant) As Variant
Dim outval As Variant
outval = collkeys(idx)
key = outval
End Function

Public Sub Remove(ByVal idx As Variant)
colldata.Remove idx
collkeys.Remove idx
End Sub

Public Sub Add(ByVal itemx As Variant, Optional ByVal key As Variant = Empty)

If (IsEmpty(key)) Then
 colldata.Add itemx
Else
 If (Exists(key)) Then
  Remove key
 End If
 colldata.Add itemx, key
 collkeys.Add key, key
End If

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
 Set NewEnum = colldata.[_NewEnum]
 
End Property
