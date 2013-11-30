Attribute VB_Name = "mdlTypes"
Option Explicit

Enum OSTypes
    Windows_95 = 1
    Windows_98 = 2
    Windows_2000 = 3
    Windows_NT = 4
    Windows_XP = 5
End Enum
Enum eSliderTypes
    Slider_Volume = 1
    Slider_Position = 2
    Slider_Balance = 3
End Enum
Enum eFiletypes
    Other_File = 0
    Mp3_File = 1
    Wav_File = 2
    Midi_File = 3
    Wma_File = 4
    Wmv_File = 5
    Mpeg_File = 6
End Enum
Enum eGFXFlash
    Audica_Logo = 1
End Enum
Private Type gFiles
    fFilename As String
    fFilepath As String
    fFiletype As eFiletypes
End Type
Private Type gPlaylist
    pFilename As String
    pFiles(200) As gFiles
    pCount As Integer
    pCurrent As Integer
    pVolumeButton As Boolean
End Type
Enum eCurrentLayout
    eSmWindow = 1
    eUtilityWindow = 2
    eAboutWindow = 3
    eNexENCODEWindow = 4
End Enum

Private Type gInterface
    iCurrentLayout As eCurrentLayout
    iCurrentFlash As eGFXFlash
    iFlashLoop As Integer
    iSliderType As eSliderTypes
    iStatusText As String
    iStatusDisplay As String
    iPlaying As Boolean
    iStoped As Boolean
    iOS As OSTypes
    iOsSelected As Integer
    iPauseLayout As Boolean
End Type

Private Type gDirectory
    dPath As String
    dFiletype As eFiletypes
End Type
Private Type gSettings
    sDirectorys(6) As gDirectory
    sOutputDevice As Integer
End Type

Public lSettings As gSettings
Public lInterface As gInterface
Public lPlaylist As gPlaylist
