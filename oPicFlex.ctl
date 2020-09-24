VERSION 5.00
Begin VB.UserControl oPicFlex 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   PropertyPages   =   "oPicFlex.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   2655
End
Attribute VB_Name = "oPicFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Enum enBorder
   None = 0
   Sunken = 1
End Enum
 
Enum enDF
   vbTwips
   vbPixels
End Enum

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Long
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean


Private oPic             As IPictureDisp 'currently active pic
Private oStoredPics()    As IPictureDisp 'pics in storage
Private oStoredPicCount  As Long 'number of stored pics

Private m_hDC As Long

'Default Property Values:
Const m_def_ActivePicture = 0
Const m_def_Stretch = 0

'Property Variables:
Dim m_ActivePicture As Long
Dim m_Stretch As Boolean
 
 
Event PictureDownloadComplete()
Event Error(ErrDescription$)
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  1/2/2005 1:42:20 AM
'
'     PaintBltFromThis: This is a conglomerate replacement for the
'     api calls BitBlt, StretchBlt, and TransparentBlt which
'     allows you to paint all or part of this controls current
'     picture and another hdc
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     outputDC            [Required | long]
'                         The device context to paint all or part
'                         of this controls current picture to
'     -------------------------------------------------------
'     x                   [Required | long]
'                         The left most point (in twips or pixels)
'                         to begin draw to the hdc pointed to by
'                         [outputDC]
'     -------------------------------------------------------
'     y                   [Required | long]
'                         The top most point (in twips or pixels)
'                         to begin draw to the hdc pointed to by
'                         [outputDC]
'     -------------------------------------------------------
'     outputWid           [Required | long]
'                         The width(in twips or pixels) of drawing
'                         to the hdc defined in [outputDC]
'     -------------------------------------------------------
'     outputHeight        [Required | long]
'                         The height(in twips or pixels) of
'                         drawing to the hdc defined in [outputDC]
'     -------------------------------------------------------
'     sourceX             [Optional | long]
'                         The left most point(in twips or pixels)
'                         to start drawing from this control
'                         (current picture) If parameter is not
'                         supplied then 0 is assumed
'     -------------------------------------------------------
'     sourceY             [Optional | long]
'                         The top most point(in twips or pixels)
'                         to start drawing from this control
'                         (current picture) If parameter is not
'                         supplied then 0 is assumed
'     -------------------------------------------------------
'     sourceWidth         [Optional | long]
'                         The width (in twips or pixels) of the
'                         current picture, from this control, to
'                         draw from. If parameter not supplied
'                         then the entire width of the controls
'                         picture is assumed
'     -------------------------------------------------------
'     sourceHeight        [Optional | long]
'                         The height(in twips or pixels) of the
'                         current picture, from this control, to
'                         draw from. If parameter not supplied
'                         then the entire height of the controls
'                         picture is assumed
'     -------------------------------------------------------
'     transparentColor    [Optional | long]
'                         Any part of this controls picture who's
'                         pixels match the color specified by this
'                         parameter, are not drawn to the output
'                         device [outputDC], and thus, are
'                         transparent.
'     -------------------------------------------------------
'     paramFormat         [Optional | enumeration]
'                         This is the scalemode of the parameters
'                         provided to this function (either twips
'                         or pixels)
'     -------------------------------------------------------
'
'ADDITIONAL INFO:
'     The scalemode of all the paramters must be
'     consistent..in other words, all the parameters or
'     arguments provided must either be provided in vbPixels
'     or vbTwips
'
'
Function PaintBltFromThis(outputDC As Long, x As Long, y As Long, _
             outputWid As Long, outputHeight As Long, _
             Optional sourceX As Long = 0, _
             Optional sourceY As Long = 0, _
             Optional sourceWidth As Long, _
             Optional sourceHeight As Long, _
             Optional transparentColor As Long = -1, _
             Optional paramFormat As enDF = vbTwips) As Boolean
   
   'if data provided in twips then convert to pixels
   If paramFormat = vbTwips Then
     x = (x / Screen.TwipsPerPixelX)
     y = (y / Screen.TwipsPerPixelY)
     outputWid = (outputWid / Screen.TwipsPerPixelX)
     outputHeight = (outputHeight / Screen.TwipsPerPixelY)
     sourceX = (sourceX / Screen.TwipsPerPixelX)
     sourceY = (sourceY / Screen.TwipsPerPixelY)
     sourceWidth = (sourceWidth / Screen.TwipsPerPixelX)
     sourceHeight = (sourceHeight / Screen.TwipsPerPixelY)
   End If
   
   If sourceWidth <= 0 Then
      sourceWidth = (ScaleWidth / Screen.TwipsPerPixelX)
   End If
   If sourceHeight <= 0 Then
      sourceHeight = (ScaleHeight / Screen.TwipsPerPixelY)
   End If
 
   If transparentColor >= 0 Then
       PaintBltFromThis = TransparentBlt(outputDC, x, y, _
                          outputWid, outputHeight, hdc, _
                          sourceX, sourceY, sourceWidth, _
                          sourceHeight, transparentColor)
    Else
       Dim b As Boolean
       Dim ret As Long

       ret = StretchBlt(outputDC, x, y, _
                        outputWid, outputHeight, hdc, _
                        sourceX, sourceY, sourceWidth, _
                        sourceHeight, vbSrcCopy)
       'translate the api return to boolean
       b = IIf(ret = 0, False, True)
       PaintBltFromThis = b
    End If
    
    'if the api fails user of the control will know y
    If PaintBltFromThis = False Then
       Dim cErr As New cApiErrorStringVal
       RaiseEvent Error(cErr.ApiErrorText(GetLastError))
    End If
End Function

'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  1/2/2005 1:29:19 AM
'
'     AddToStoredPics: This function redims [oStoredPic] (which is
'     an object of IPictureDisp...or more commonly known as
'     stdPicture) and adds a picture to this array
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     pathToPic           [Optional | string]
'                         This is a local filepath to where a
'                         picture resides. If this value isnt
'                         supplied then it is assumed that we add
'                         the controls current picture to the
'                         array
'
Function AddToStoredPics(Optional pathToPic As String)
   ReDim Preserve oStoredPics(oStoredPicCount)
   If Len(Trim(pathToPic)) = 0 Then
     Set oStoredPics(oStoredPicCount) = oPic
   Else
     Set oStoredPics(oStoredPicCount) = LoadPicture(pathToPic)
   End If
   oStoredPicCount = (oStoredPicCount + 1)
End Function
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  1/2/2005 1:30:30 AM
'
'     ClearStoredPics:  Releases reference to all stored pics in
'     the [oStoredPics] array and erases the array so it is
'     released from memory
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     pathToPic           [Optional | string]
'                         This is a local filepath to where a
'                         picture resides. If this value isnt
'                         supplied then it is assumed that we add
'                         the controls current picture to the
'                         array
'     -------------------------------------------------------
'
Function ClearStoredPics()
Dim lcnt As Long

   For lcnt = 0 To (oStoredPicCount - 1)
      Set oStoredPics(lcnt) = Nothing
   Next lcnt
   
   Erase oStoredPics
   oStoredPicCount = 0
End Function
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  1/2/2005 1:32:31 AM
'
'     LoadPicFromInternet: This uses the usercontrol.AsyncRead
'     method which allows you to download data off the internet
'     without the need for a third party control such as a winsock
'     control
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     sURL                [Required | string]
'                         This is the url or http address that
'                         points to the location of the picture
'     -------------------------------------------------------
'
Sub LoadPicFromInternet(sURL$)
Dim rightPart As String
On Error GoTo local_error:

   '--validate to make sure url ends with .gif/.jpg/.jpeg
   rightPart = LCase(Trim(Right(sURL, 4)))
   If rightPart = ".gif" Or rightPart = ".jpg" Or rightPart = "jpeg" Then
       AsyncRead sURL, vbAsyncTypePicture
   Else
      RaiseEvent Error("Url not valid and consistent with an internet image file .gif,.jpg,.jpeg")
   End If
   
Exit Sub
local_error:
   RaiseEvent Error(Err.Description)
   Exit Sub
End Sub
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
  On Error Resume Next
  
  Debug.Print Format((AsyncProp.BytesMax - AsyncProp.BytesRead) / 2 ^ 10, "#0.0 KB") & _
        " (" & Format(AsyncProp.BytesRead / AsyncProp.BytesMax, "#0.0%") & ")"
End Sub
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error GoTo local_error:

  Set oPic = AsyncProp.Value
  Call CreateFromPicture(oPic)
  RaiseEvent PictureDownloadComplete
  Debug.Print "download complete"

Exit Sub
local_error:
   RaiseEvent Error(Err.Description)
   Exit Sub
End Sub

'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  1/2/2005 1:51:05 AM
'
'     CreateFromPicture: This function was obtained (slightly
'     modified by me) from VBaccelorator.com.  Its purpose is to
'     take a picture object and make it compatible for the device
'     context of this control
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     sPic                [Required | long]
'                         The picture object to format for display
'                         to this control
'     -------------------------------------------------------
'
Public Sub CreateFromPicture(sPic As IPicture)
Dim tB As BITMAP
Dim lhDCC As Long, lhDC As Long
Dim lhBmpOld As Long
Dim lwid As Long, lhei As Long

On Error GoTo local_error:

   Cls 'erase any last picture
   GetObjectAPI sPic.Handle, Len(tB), tB
   'adjustment for picture size depending on property [Stretch]
   If Stretch = True Then
     lwid = (Width / Screen.TwipsPerPixelX)
     lhei = (Height / Screen.TwipsPerPixelY)
   Else
     lwid = tB.bmWidth
     lhei = tB.bmHeight
   End If
   lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDC = CreateCompatibleDC(lhDCC)
   lhBmpOld = SelectObject(lhDC, sPic.Handle)
   StretchBlt hdc, 0, 0, lwid, lhei, lhDC, 0, 0, tB.bmWidth, tB.bmHeight, vbSrcCopy
   SelectObject lhDC, lhBmpOld
   DeleteDC lhDC
   DeleteDC lhDCC
   
Exit Sub
local_error:
   RaiseEvent Error(Err.Description)
   Exit Sub
End Sub

'ACTIVEPICTURE
Public Property Get ActivePicture() As Long
    ActivePicture = m_ActivePicture
End Property
Public Property Let ActivePicture(ByVal New_ActivePicture As Long)
    If New_ActivePicture < oStoredPicCount Then
       m_ActivePicture = New_ActivePicture
       PropertyChanged "ActivePicture"
       Set oPic = oStoredPics(New_ActivePicture)
       Call CreateFromPicture(oPic)
    End If
End Property
'BORDERSTYLE
Public Property Get BorderStyle() As enBorder
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As enBorder)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'HDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property
'HWND
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
'PICTURE
Public Property Get Picture() As Picture
    Set Picture = oPic
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set oPic = New_Picture
    PropertyChanged "Picture"
    Call CreateFromPicture(oPic)
End Property
'STRETCH
Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "If True, the picture stretches to fit the size of this control, else, the control resizes to fit the picture"
    Stretch = m_Stretch
End Property
Public Property Let Stretch(ByVal New_Stretch As Boolean)
    m_Stretch = New_Stretch
    PropertyChanged "Stretch"
    If oPic Is Nothing Then Exit Property
    Call CreateFromPicture(oPic)
End Property
 
Private Sub UserControl_Resize()
   If oPic Is Nothing Then Exit Sub
   Call CreateFromPicture(oPic)
End Sub
Private Sub UserControl_Terminate()
   Set oPic = Nothing
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Stretch = m_def_Stretch
    m_ActivePicture = m_def_ActivePicture
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Stretch = PropBag.ReadProperty("Stretch", m_def_Stretch)
    Set oPic = PropBag.ReadProperty("Picture", Nothing)
    m_ActivePicture = PropBag.ReadProperty("ActivePicture", m_def_ActivePicture)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)

    If oPic Is Nothing Then Exit Sub
    Call CreateFromPicture(oPic)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Stretch", m_Stretch, m_def_Stretch)
    Call PropBag.WriteProperty("Picture", oPic, Nothing)
    Call PropBag.WriteProperty("ActivePicture", m_ActivePicture, m_def_ActivePicture)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
End Sub





