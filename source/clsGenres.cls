VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection

Public Function Add(Description As String, Identifier As String) As clsGenre
On Error GoTo AddError
Const ITEM_EXISTS_ERR_NUMBER  As Integer = 457
Dim objNewMember              As clsGenre
Set objNewMember = New clsGenre
With objNewMember
   .Description = Description
   .Identifier = Identifier
End With
mCol.Add objNewMember, Identifier
Set Add = objNewMember
Set objNewMember = Nothing
AddExit:
Exit Function
AddError:
If Err.Number = ITEM_EXISTS_ERR_NUMBER Then
   Set Add = mCol(Identifier)
Else
   With Err
      .Raise .Number, .Source, .Description
   End With
End If
Resume AddExit
End Function

Public Function Exists(vntIndexKey As Variant) As Boolean
On Error GoTo ExistsError
Dim objTest As clsGenre
Exists = True
Set objTest = mCol(vntIndexKey)
ExistsExit:
   Exit Function
ExistsError:
   Exists = False
End Function

Public Property Get Item(vntIndexKey As Variant) As clsGenre
Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
mCol.Remove vntIndexKey
End Sub

Private Sub Class_Initialize()
Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
Set mCol = Nothing
End Sub
