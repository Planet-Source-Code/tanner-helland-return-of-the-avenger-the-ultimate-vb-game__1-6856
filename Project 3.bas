Attribute VB_Name = "ScreenRes_Module"
'Global Screen Size Variables
Public ScreenWidth As Long
Public ScreenHeight As Long

'Screen Resolution Variables, API Calls, Types, Constants, etc.
Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
    Public Const CCDEVICENAME = 32
    Public Const CCFORMNAME = 32
    Public Const DM_PELSWIDTH = &H80000
    Public Const DM_PELSHEIGHT = &H100000
    Public Const CDS_UPDATEREGISTRY = &H1
Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Sub ChangeRes()

Dim DevM As DEVMODE

erg& = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
DevM.dmPelsWidth = 640
DevM.dmPelsHeight = 480
DevM.dmBitsPerPel = 16
erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)

End Sub

Public Sub RestoreRes()

Dim DevM As DEVMODE
erg& = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
DevM.dmPelsWidth = ScreenWidth
DevM.dmPelsHeight = ScreenHeight
DevM.dmBitsPerPel = 32
erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)

End Sub

Public Sub SystemRes()

End Sub

