VERSION 5.00
Begin VB.Form OSDwin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.Timer ScrollTimer 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   480
      Top             =   0
   End
   Begin VB.Label text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   150
   End
End
Attribute VB_Name = "OSDwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const HWND_TOPMOST        As Integer = -1
Private Const SWP_NOMOVE          As Long = &H2
Private Const SWP_NOSIZE          As Long = &H1
Private Const RC_PALETTE          As Long = &H100
Private Const SIZEPALETTE         As Long = 104
Private Const RASTERCAPS          As Long = 38
Private Type PALETTEENTRY
    peRed                           As Byte
    peGreen                         As Byte
    peBlue                          As Byte
    peFlags                         As Byte
End Type
Private Type LOGPALETTE
    palVersion                      As Integer
    palNumEntries                   As Integer
    palPalEntry(255)                As PALETTEENTRY    ' Enough for 256 colors
End Type
Private Type GUID
    Data1                           As Long
    Data2                           As Integer
    Data3                           As Integer
    Data4(7)                        As Byte
End Type
Private Type PicBmp
    Size                            As Long
Type                            As Long '<:-):SUGGESTION: Legal but ill-advised to use VB Reserved words as Variable names
    hBmp                            As Long
    hPal                            As Long
    Reserved                        As Long
End Type

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, _
                                                                      RefIID As GUID, _
                                                                      ByVal fPictureOwnsHandle As Long, _
                                                                      IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, _
                                                              ByVal wStartIndex As Long, _
                                                              ByVal wNumEntries As Long, _
                                                              lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hPalette As Long, _
                                                    ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Function CreateBitmapPicture(ByVal hBmp As Long, _
                                     ByVal hPal As Long) As Picture


    Dim R             As Long

    Dim Pic           As PicBmp
    Dim IPic          As IPicture
    Dim IID_IDispatch As GUID
    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With 'IID_IDISPATCH
    'Fill picture info
    With Pic
        .Size = Len(Pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With 'PIC
    'Create the picture
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    'Return the new picture
    Set CreateBitmapPicture = IPic

End Function

Private Sub Form_Click()
'Unload if user clicks on OSD
    Unload Me

End Sub

Private Sub Form_Load()
'Start, our Positioning and Other Stuff
    With Me
        'Me.Top = 1000
        .Left = 0
        .Width = Screen.Width
        'Get the Background Picture to make our form transparent
        Set .Picture = hDCToPicture(GetDC(0), 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
        'Make Our form Top-Most
        SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End With
    'Enable the timer will make the animation
    'Change Timer time to makethe animation fast or Slow
     ScrollTimer.Interval = 25
     ScrollTimer.Enabled = True

End Sub

Private Function hDCToPicture(ByVal hDCSrc As Long, _
                              ByVal LeftSrc As Long, _
                              ByVal TopSrc As Long, _
                              ByVal WidthSrc As Long, _
                              ByVal HeightSrc As Long) As Picture


    Dim hDCMemory       As Long
    Dim hBmp            As Long
    Dim hBmpPrev        As Long
    Dim R               As Long

    Dim hPal            As Long
    Dim hPalPrev        As Long
    Dim RasterCapsScrn  As Long
    Dim HasPaletteScrn  As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal          As LOGPALETTE
    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        With LogPal
            .palVersion = &H300
            'Number of palette entries
            .palNumEntries = 256
            'Retrieve the system palette entries
            R = GetSystemPaletteEntries(hDCSrc, 0, 256, .palPalEntry(0))
            'Create the palette
        End With 'LogPal
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If
    'Copy the source image to our compatible device context
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    'Delete our memory DC
    R = DeleteDC(hDCMemory)
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)

End Function

Private Sub ScrollTimer_Timer()
    If text.Left > -text.Width Then
        text.Left = text.Left - 125
    Else
        Unload Me
    End If

End Sub

Private Sub text_Click()
'Unload if user clicks on OSD
    Unload Me

End Sub



