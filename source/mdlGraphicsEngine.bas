Attribute VB_Name = "mdlGraphicsEngine"
Option Explicit

'Option Explicit

'Collision detection

Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

'Graphics drawing functions

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Pixel manipulation functions

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Clipboard functions

Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

'Miscellaneous functions

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'LoadImage Constants

Public Const LR_LOADFROMFILE = &H10
Public Const CF_BITMAP = 2

'Pixel manipulation constants

Public Const PIXELGET = 0
Public Const PIXELSET = 1

Public Const PIXELS = 3

'graphics drawing enumeration

Public Enum dwRop

    WHITENESS = &HFF0062
    BLACKNESS = &H42
    SRCAND = &H8800C6
    SRCCOPY = &HCC0020
    SRCINVERT = &H660046
    SRCERASE = &H440328
    SRCPAINT = &HEE0086
    
End Enum

'LoadImage enumeration

Public Enum LoadImg

    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
    
End Enum

'Graphics drawing types

Public Type POINTAPI

  X As Long
  Y As Long
  
End Type


'graphics drawing variables

Public PointType As POINTAPI
Public rectangle As RECT

'Graphics drawing function

Public Function DoBitBlt(ByRef Destination As PictureBox, ByVal DestinationX As Long, ByVal DestinationY As Long, ByVal DestinationWidth As Long, ByVal DestinationHeight As Long, ByRef Sprite As PictureBox, ByVal SpriteX As Long, ByVal SpriteY As Long, ByVal SpriteWidth As Long, ByVal SpriteHeight As Long, ByRef Mask As PictureBox, ByVal MaskX As Long, ByVal MaskY As Long, ByVal MaskWidth As Long, ByVal MaskHeight As Long) As Long

If DestinationWidth = SpriteWidth And DestinationHeight = SpriteHeight Then
    
    DoBitBlt = BitBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Mask.hdc, MaskX, MaskY, dwRop.SRCAND)
    DoBitBlt = BitBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Sprite.hdc, SpriteX, SpriteY, dwRop.SRCPAINT)

ElseIf DestinationWidth <> SpriteWidth Or DestinationHeight <> SpriteHeight Then
    
    DoBitBlt = StretchBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Mask.hdc, MaskX, MaskY, MaskWidth, MaskHeight, dwRop.SRCAND)
    DoBitBlt = StretchBlt(Destination.hdc, DestinationX, DestinationY, DestinationWidth, DestinationHeight, Sprite.hdc, SpriteX, SpriteY, SpriteWidth, SpriteHeight, dwRop.SRCPAINT)

Else

    DoBitBlt = 0
    
End If

End Function

'LoadImage function

Public Function LoadImgFromFile(ByVal ImageType As LoadImg, ByVal Width As Long, ByVal Height As Long, ByVal Path As String, Optional ByVal Dest As PictureBox = 0) As Long

On Error Resume Next

Dim ImageFormat As Long

LoadImgFromFile = LoadImage(App.hInstance, Path, ImageType, Width, Height, LR_LOADFROMFILE)
    
OpenClipboard Form1.hwnd
EmptyClipboard
    
SetClipboardData CF_BITMAP, LoadImgFromFile
        
If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then
    
    Exit Function
    
End If
    
CloseClipboard
    
Dest.Picture = Clipboard.GetData(CF_BITMAP)
        
End Function

'Pixel manipultion function

Public Function GetOrSetPixel(ByVal DC As Long, ByVal GetOrSet As Long, ByRef Point As POINTAPI, ByRef PixelColor As ColorConstants)
    
Select Case GetOrSet
        
Case PIXELGET

    GetOrSetPixel = GetPixel(DC, Point.X, Point.Y)

Case PIXELSET

    GetOrSetPixel = SetPixelV(DC, Point.X, Point.Y, PixelColor)

Case Else

    GetOrSetPixel = 0

End Select

End Function

'Animation functions

Public Sub MultiPictureAnimation(ByVal Destination As PictureBox, ByRef Animations() As PictureBox, ByRef Masks() As PictureBox, ByVal NumberOfTimes As Integer, ByVal TimeBetween_Secs As Integer)

Dim Counter As Integer, Counter2 As Integer

For Counter = 1 To NumberOfTimes
    
    For Counter2 = 1 To UBound(Animations)
        
        Destination.Cls
        DoBitBlt Destination, 0, 0, Destination.ScaleWidth, Destination.ScaleHeight, Animations(Counter2), 0, 0, Animations(Counter2).ScaleWidth, Animations(Counter2).ScaleHeight, Masks(Counter2), 0, 0, Masks(Counter2).ScaleWidth, Masks(Counter2).ScaleHeight
        DoEvents
        Sleep TimeBetween_Secs * 1000
        
    Next Counter2
    
Next Counter

End Sub


Public Sub SinglePictureAnimation(ByRef Destination As PictureBox, ByVal NumberOfFrames As Long, ByRef Animation As PictureBox, ByRef Mask As PictureBox, ByVal NumberOfTimes As Integer, ByVal TimeBetween_Secs As Double)

Dim Counter As Integer, Counter2 As Integer

For Counter = 1 To NumberOfTimes
    
    For Counter2 = 1 To NumberOfFrames
        
        Destination.Cls
        DoBitBlt Destination, 0, 0, Destination.ScaleWidth, Destination.ScaleHeight, Animation, (Animation.ScaleWidth / NumberOfFrames) * (Counter2 - 1), 0, Animation.ScaleWidth / NumberOfFrames, Animation.ScaleHeight, Mask, (Mask.ScaleWidth / NumberOfFrames) * (Counter2 - 1), 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight
        DoEvents
        Sleep TimeBetween_Secs * 1000
        
    Next Counter2
    
Next Counter

End Sub

'Function for moving a sprite on a background

Public Function MoveSprite(ByRef Sprite As PictureBox, ByRef Mask As PictureBox, ByRef Background As PictureBox, ByVal Direction As String, ByVal Distance_Pixels As Long, ByVal startX As Single, startY As Single, ByVal Speed As Long, Optional ByVal NumberOfFrames As Long = 1) As String

Dim X As Single, Y As Single

Select Case Direction

Case "Up"
    
    X = startX
    
    For Y = startY To Distance_Pixels + startY
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, Y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
    
    Next Y

Case "Down"
    
    X = startX
    
    For Y = Distance_Pixels + startY To startY Step -1
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, Y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
        
    Next Y

Case "Left"
    
    Y = startY

    For X = Distance_Pixels + startX To startX Step -1
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, Y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
        
    Next X

Case "Right"

    Y = startY

    For X = startX To Distance_Pixels + startX
        
        Background.Picture = LoadPicture
        MoveSprite = DoBitBlt(Background, X, Y, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Sprite, 0, 0, Sprite.ScaleWidth / NumberOfFrames, Sprite.ScaleHeight, Mask, 0, 0, Mask.ScaleWidth / NumberOfFrames, Mask.ScaleHeight)
        Background.Refresh
        Sleep Speed * 4
        DoEvents
    
    Next X

End Select

End Function

'Main starting sub

Public Sub Main()

ActiveForm.ScaleMode = PIXELS

For Each PictureBox In ActiveForm
    PictureBox.ScaleMode = PIXELS
    PictureBox.AutoRedraw = True
Next

For Each Container In ActiveForm
    Container.ScaleMode = PIXELS
Next

End Sub

