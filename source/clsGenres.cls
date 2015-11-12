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

'local variable to hold collection
Private mCol As Collection
Public Function Add( _
      Description As String, _
      Identifier As String _
   ) As clsGenre

On Error GoTo AddError

'Constant error number raised when trying
'to add an item with a key that already
'exists within the collection
   Const ITEM_EXISTS_ERR_NUMBER  As Integer = 457

   Dim objNewMember              As clsGenre

   Set objNewMember = New clsGenre

   'set the properties passed into the method
   With objNewMember
      .Description = Description
      .Identifier = Identifier
   End With

   'Add the item using the Identifier as key

   mCol.Add objNewMember, Identifier

   'return the object created
   Set Add = objNewMember
   Set objNewMember = Nothing

AddExit:
   Exit Function

AddError:
   If Err.Number = ITEM_EXISTS_ERR_NUMBER Then
      'Return the existing object
      Set Add = mCol(Identifier)
   Else
      With Err
         .Raise .Number, .Source, .Description
      End With
   End If

   Resume AddExit

End Function
Public Function Exists( _
      vntIndexKey As Variant _
   ) As Boolean

On Error GoTo ExistsError

   Dim objTest As clsGenre

   Exists = True

   Set objTest = mCol(vntIndexKey)

ExistsExit:
   Exit Function

ExistsError:
   Exists = False
End Function
Public Property Get Item( _
      vntIndexKey As Variant _
   ) As clsGenre
'used when referencing an element in the collection
'vntIndexKey contains either the Index or Key to the collection,
'this is why it is declared as a Variant
'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)

   Set Item = mCol(vntIndexKey)

End Property
Public Property Get Count() As Long
'used when retrieving the number of elements in the
'collection. Syntax: Debug.Print x.Count
   Count = mCol.Count

End Property
Public Sub Remove(vntIndexKey As Variant)
'used when removing an element from the collection
'vntIndexKey contains either the Index or Key, which is why
'it is declared as a Variant
'Syntax: x.Remove(xyz)

   mCol.Remove vntIndexKey

End Sub
'Public Property Get NewEnum() As IUnknown
'this property allows you to enumerate
'this collection with the For...Each syntax
'eg For Each objMember In objCollection

   'Set NewEnum = mCol.[_NewEnum]

'End Property
Private Sub Class_Initialize()

'creates the collection when this class is created
   Set mCol = New Collection

End Sub
Private Sub Class_Terminate()
'destroys collection when this class is terminated
   Set mCol = Nothing

End Sub


