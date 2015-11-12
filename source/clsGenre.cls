VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variables to hold property values
Private m_Description   As String   'local copy
Private m_Identifier    As String   'local copy
Public Property Let Identifier( _
      ByVal lIdentifier As String _
   )
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Identifier = 5
   m_Identifier = lIdentifier

End Property
Public Property Get Identifier() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Identifier
   Identifier = m_Identifier

End Property
Public Property Let Description( _
      ByVal sDescription As String _
   )
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
   m_Description = sDescription

End Property
Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
   Description = m_Description

End Property


