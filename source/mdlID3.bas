Attribute VB_Name = "mdlID3"
'Option Explicit
Private m_objGenres              As clsGenres
Private Type ID3V1Tag
   Album       As String * 30
   Artist      As String * 30
   Comment     As String * 30
   Genre       As Byte
   Identifier  As String * 3
   Title       As String * 30
   Year        As String * 4
End Type
Public Type Id3Tag
   Album       As String * 30
   Artist      As String * 30
   Comment     As String * 30
   Genre       As String * 30
   Identifier  As String * 3
   Title       As String * 30
   Year        As String * 4
End Type

Public Function ms_ShowID3V1Tag(sFileName As String) As Id3Tag
   On Local Error GoTo ErrHandler
   Const ID3V1TagSize   As Integer = 127
   Dim result As Id3Tag
   Dim t                As ID3V1Tag
   Dim lFileHandle      As Long
   Dim lll              As Long
   Dim sGenre           As String
   lFileHandle = FreeFile()
   Open sFileName For Binary As #lFileHandle
   lll = LOF(lFileHandle) 'Get the length of mp3 file
   Get #lFileHandle, lll - ID3V1TagSize, t.Identifier
   With t
      If .Identifier = "TAG" Then
         Get #lFileHandle, , .Title   '30 chars
         Get #lFileHandle, , .Artist  '30 chars
         Get #lFileHandle, , .Album   '30 chars
         Get #lFileHandle, , .Year    '4 chars
         Get #lFileHandle, , .Comment '30 chars
         Get #lFileHandle, , .Genre   '1 byte (i think)
         result.Album = Trim(.Album)
         result.Artist = Trim(.Artist)
         result.Comment = Trim(.Comment)
         result.Identifier = Trim(.Identifier)
         result.Title = Trim(.Title)
         result.Year = Trim(.Year)
         'sGenre = CStr(.Genre)
         'If m_objGenres.Exists(CStr(sGenre)) Then
            'Dim g As clsGenre
            'g = m_objGenres.Item(CStr(sGenre))
            'result.Genre = g.Description
         'End If
         'ms_ShowID3V1Tag = result
      End If
   End With
   ms_ShowID3V1Tag = result
   Close
Exit Function
ErrHandler:
   MsgBox "Error: " & Err.Description
End Function

Public Sub ms_InitialiseGenres()
   Dim objXMLDocument   As Object 'MSXML2.DOMDocument
   Dim objNodeList      As Object 'MSXML2.IXMLDOMNodeList
   Dim objRoot          As Object 'MSXML2.IXMLDOMElement
   Dim objNode          As Object 'MSXML2.IXMLDOMNode
   Dim objChild         As Object 'MSXML2.IXMLDOMNode
   Dim sIdentifier      As String
   Dim sGenre           As String
   Dim XML_FILE       As String
   If Right(App.Path, 1) <> mc_DIR_SEPARATOR Then
      XML_FILE = App.Path & "\Genres.xml"
   Else
      XML_FILE = App.Path & "Genres.xml"
   End If
   Set objXMLDocument = CreateObject("Microsoft.XMLDOM") '= New MSXML2.DOMDocument
   With objXMLDocument
      .async = False
      If .Load(XML_FILE) Then
         Set objRoot = .documentElement()
         For Each objNode In objRoot.childNodes
            sGenre = vbNullString
            sIdentifier = vbNullString
            For Each objChild In objNode.childNodes
               If objChild.nodeName = "id" Then
                  sIdentifier = objChild.Text
               ElseIf objChild.nodeName = "Description" Then
                  sGenre = objChild.Text
               End If
            Next
            If sGenre <> vbNullString Then
               If sIdentifier <> vbNullString Then
                  m_objGenres.Add sGenre, sIdentifier
               End If
            End If
         Next
         'ms_LoadGenreComboBox
      Else
         MsgBox "Error loading xml file: " & XML_FILE & vbCrLf & _
            "Check if the path to the file is correct", _
            vbExclamation, "Cannot Find XML File"
      End If
   End With
End Sub

'Public Sub ms_LoadGenreComboBox(box As ComboBox)
   'Dim objGenre   As Genre
   'With box
      '.Clear
      '.AddItem vbNullString
   'End With
   'For Each objGenre In m_objGenres
      'With objGenre
         'box.AddItem .Description
      'End With
   'Next
'End Sub
