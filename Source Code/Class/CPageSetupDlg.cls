VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPageSetupDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' © 1999 MW Software, all rights reserved
Option Explicit

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type POINTAPI
        X As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type DEVMODE
        dmDeviceName(0 To 31) As Byte
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
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Type PageSetupDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As RECT
        rtMargin As RECT
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

Private Type DEVNAMES
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 100
End Type

Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (PPageSetupDlg As PageSetupDlg) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Enum eOptions
    PSD_MINMARGINS = &H1
    PSD_MARGINS = &H2
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEPRINTER = &H20
    PSD_NOWARNING = &H80
    PSD_DISABLEORIENTATION = &H100
    PSD_DISABLEPAPER = &H200
    PSD_DEFAULTMINMARGINS = &H0
    PSD_DISABLEPAGEPAINTING = &H80000
End Enum

Public Enum eErrors
    ERR_CANCEL = 0
End Enum

Public Enum ePaperSize
    DMPAPER_LETTER = 1
    DMPAPER_FIRST = DMPAPER_LETTER     '  Letter 8 1/2 x 11 in
    DMPAPER_LETTERSMALL = 2            '  Letter Small 8 1/2 x 11 in
    DMPAPER_TABLOID = 3                '  Tabloid 11 x 17 in
    DMPAPER_LEDGER = 4                 '  Ledger 17 x 11 in
    DMPAPER_LEGAL = 5                  '  Legal 8 1/2 x 14 in
    DMPAPER_STATEMENT = 6              '  Statement 5 1/2 x 8 1/2 in
    DMPAPER_EXECUTIVE = 7              '  Executive 7 1/4 x 10 1/2 in
    DMPAPER_A3 = 8                     '  A3 297 x 420 mm
    DMPAPER_A4 = 9                     '  A4 210 x 297 mm
    DMPAPER_A4SMALL = 10               '  A4 Small 210 x 297 mm
    DMPAPER_A5 = 11                    '  A5 148 x 210 mm
    DMPAPER_B4 = 12                    '  B4 250 x 354
    DMPAPER_B5 = 13                    '  B5 182 x 257 mm
    DMPAPER_FOLIO = 14                 '  Folio 8 1/2 x 13 in
    DMPAPER_QUARTO = 15                '  Quarto 215 x 275 mm
    DMPAPER_10X14 = 16                 '  10x14 in
    DMPAPER_11X17 = 17                 '  11x17 in
    DMPAPER_NOTE = 18                  '  Note 8 1/2 x 11 in
    DMPAPER_ENV_9 = 19                 '  Envelope #9 3 7/8 x 8 7/8
    DMPAPER_ENV_10 = 20                '  Envelope #10 4 1/8 x 9 1/2
    DMPAPER_ENV_11 = 21                '  Envelope #11 4 1/2 x 10 3/8
    DMPAPER_ENV_12 = 22                '  Envelope #12 4 \276 x 11
    DMPAPER_ENV_14 = 23                '  Envelope #14 5 x 11 1/2
    DMPAPER_CSHEET = 24                '  C size sheet
    DMPAPER_DSHEET = 25                '  D size sheet
    DMPAPER_ESHEET = 26                '  E size sheet
    DMPAPER_ENV_DL = 27                '  Envelope DL 110 x 220mm
    DMPAPER_ENV_C5 = 28                '  Envelope C5 162 x 229 mm
    DMPAPER_ENV_C3 = 29                '  Envelope C3  324 x 458 mm
    DMPAPER_ENV_C4 = 30                '  Envelope C4  229 x 324 mm
    DMPAPER_ENV_C6 = 31                '  Envelope C6  114 x 162 mm
    DMPAPER_ENV_C65 = 32               '  Envelope C65 114 x 229 mm
    DMPAPER_ENV_B4 = 33                '  Envelope B4  250 x 353 mm
    DMPAPER_ENV_B5 = 34                '  Envelope B5  176 x 250 mm
    DMPAPER_ENV_B6 = 35                '  Envelope B6  176 x 125 mm
    DMPAPER_ENV_ITALY = 36             '  Envelope 110 x 230 mm
    DMPAPER_ENV_MONARCH = 37           '  Envelope Monarch 3.875 x 7.5 in
    DMPAPER_ENV_PERSONAL = 38          '  6 3/4 Envelope 3 5/8 x 6 1/2 in
    DMPAPER_FANFOLD_US = 39            '  US Std Fanfold 14 7/8 x 11 in
    DMPAPER_FANFOLD_STD_GERMAN = 40    '  German Std Fanfold 8 1/2 x 12 in
    DMPAPER_FANFOLD_LGL_GERMAN = 41    '  German Legal Fanfold 8 1/2 x 13 in
    DMPAPER_CUSTOM = 256               '  Custom Paper Size
End Enum

Public Enum eOrientation
    '  orientation selections
    DMORIENT_PORTRAIT = 1
    DMORIENT_LANDSCAPE = 2
End Enum

Private tPageSetupDlg As PageSetupDlg
Private ChosenOptions As Long
Private ChosenhWnd As Long
Private ChosenError As Long
Private ChosenPrinter As String
Private ChosenPaperSize As Long
Private ChosenOrientation As Long
Private ChosenDriver As String
Private ChosenPort As String

Public Function ShowPageSetupDlg()
    Dim lTemp As Long, sTemp As String, NTemp As Long
    Dim tDevMode As DEVMODE
    Dim tDevNames As DEVNAMES
    With tPageSetupDlg
        .hwndOwner = ChosenhWnd
        ' Set the owner of the window
        tDevMode.dmOrientation = 2
        .hDevMode = 0&
        ' Set this to Null (0&)
        .hDevNames = 0&
        .flags = ChosenOptions
        ' Set the flags
        .hInstance = 0&
        ' Set to Null (0&)
        .lStructSize = Len(tPageSetupDlg)
        ' Set the size of the structure
        ChosenError = PageSetupDlg(tPageSetupDlg)
        ' Call the Dialog and return error information
    End With
    lTemp = GlobalLock(tPageSetupDlg.hDevMode)
    ' Lock the memory object in place and return a pointer
    ' to the object in memory
    Call CopyMemory(tDevMode, ByVal lTemp, Len(tDevMode))
    ' Copy the memory object using to the pointer to a
    ' devmode structure
    Call GlobalUnlock(tPageSetupDlg.hDevMode)
    ' Unlock the memory oibject using the pointer
    NTemp = GlobalLock(tPageSetupDlg.hDevNames)
    Call CopyMemory(tDevNames, ByVal NTemp, Len(tDevNames))
    Call GlobalUnlock(tPageSetupDlg.hDevNames)
    sTemp = StrConv(tDevMode.dmDeviceName, vbUnicode)
    ' Convert the string to unicode
    ChosenPrinter = Left$(sTemp, lstrlen(sTemp))
    ' Resize the string. NB: Use lstrlen to get the size
    ' of the ANSI version and not Len
    ChosenPaperSize = tDevMode.dmPaperSize
    ' Return the paper size
    ChosenOrientation = tDevMode.dmOrientation
    ' Return the orientation
    With tDevNames
    ChosenDriver = Mid(.extra, 1, .wDriverOffset)
    'ChosenDriver = Left$(sTemp, lstrlen(sTemp))
    sTemp = Mid(.extra, .wOutputOffset - 7, Len(.extra) - .wDeviceOffset)
    ChosenPort = Left$(sTemp, lstrlen(sTemp))
    End With
End Function

Public Property Let Options(lOptions As eOptions)
    ChosenOptions = lOptions
End Property

Public Property Let hwnd(lhWnd As Long)
    ChosenhWnd = lhWnd
End Property

Public Property Get Error() As eErrors
    Error = ChosenError
End Property

Public Property Get Printer() As String
    Printer = ChosenPrinter
End Property

Public Property Get X() As Long
    X = tPageSetupDlg.ptPaperSize.X / 100
End Property

Public Property Get y() As Long
    y = tPageSetupDlg.ptPaperSize.y / 100
End Property

Public Property Get Top() As Long
    Top = tPageSetupDlg.rtMargin.Top / 100
End Property

Public Property Get Bottom() As Long
    Bottom = tPageSetupDlg.rtMargin.Bottom / 100
End Property

Public Property Get Left() As Long
    Left = tPageSetupDlg.rtMargin.Left / 100
End Property

Public Property Get Right() As Long
    Right = tPageSetupDlg.rtMargin.Right / 100
End Property

Public Property Get PaperSize() As ePaperSize
    PaperSize = ChosenPaperSize
End Property

Public Property Get Orientation() As eOrientation
    Orientation = ChosenOrientation
End Property

Property Let Right(lOption As Long)
    tPageSetupDlg.rtMargin.Right = lOption * 100
End Property

Property Let Top(lOption As Long)
    tPageSetupDlg.rtMargin.Top = lOption * 100
End Property

Property Let Bottom(lOption As Long)
    tPageSetupDlg.rtMargin.Bottom = lOption * 100
End Property

Property Let Left(lOption As Long)
    tPageSetupDlg.rtMargin.Left = lOption * 100
End Property

Property Let MinRight(lOption As Long)
    tPageSetupDlg.rtMinMargin.Right = lOption * 100
End Property

Property Let MinTop(lOption As Long)
    tPageSetupDlg.rtMinMargin.Top = lOption * 100
End Property

Property Let MinBottom(lOption As Long)
    tPageSetupDlg.rtMinMargin.Bottom = lOption * 100
End Property

Property Let MinLeft(lOption As Long)
    tPageSetupDlg.rtMinMargin.Left = lOption * 100
End Property
Public Property Get DriverName() As String
    DriverName = ChosenDriver
End Property
Public Property Get PortName() As String
    PortName = ChosenPort
End Property

