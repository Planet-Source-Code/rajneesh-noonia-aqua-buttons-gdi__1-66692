VERSION 5.00
Begin VB.UserControl AquaButton 
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "AquaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_StartColour As Colors
Private m_EndColor As Colors
Private m_LightIntentisity As Byte
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32.DLL" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "GDI32.DLL" (ByVal hObject As Long) As Long
Private m_MemoryDc As cMemDC

Public Property Get StartColour() As Colors
    StartColour = m_StartColour
End Property

Public Property Let StartColour(ByVal pStartColour As Colors)
    m_StartColour = pStartColour
    PropertyChanged "StartColour"
    UserControl.Refresh
End Property

Public Property Get EndColour() As Colors
    EndColour = m_EndColor
End Property

Public Property Let EndColour(ByVal pEndColor As Colors)
    m_EndColor = pEndColor
    PropertyChanged "EndColour"
    UserControl.Refresh
End Property

Public Property Get LightIntentisity() As Byte
    LightIntentisity = m_LightIntentisity
End Property

Public Property Let LightIntentisity(ByVal pLightIntentisity As Byte)
    m_LightIntentisity = pLightIntentisity
    PropertyChanged "LightIntentisity"
    UserControl.Refresh
End Property




Private Sub UserControl_Paint()
    Call SetRegion
    Call Gradient
    If m_MemoryDc.hdc <> 0 Then
        m_MemoryDc.Draw UserControl.hdc
    End If
End Sub

Private Sub UserControl_Initialize()
    m_StartColour = Green
    m_EndColor = White
    m_LightIntentisity = 250
    Set m_MemoryDc = New cMemDC
    Call modGDIPlus.GDIPlusCreate
End Sub



Private Sub Gradient()
    Dim pGraphics As Long
    Dim pStartC As Colors
    Dim pEndColor As Colors
    
    
    
    Dim pPt1 As POINTL
    Dim pPt2 As POINTL
    Dim pshadowOffset As Long
    
    Dim pRect As RECTF
    Dim pPath As Long
    Dim pBrush As Long
    
    Dim prect2 As RECTF
    Dim pPath2 As Long
    Dim pBr2 As Long
    
    Dim prect3 As RECTF
    Dim pPath3 As Long
    Dim pBr3 As Long
    
    If m_MemoryDc.hdc <> 0 Then
        Call GdipCreateFromHDC(m_MemoryDc.hdc, pGraphics)
        Call GdipGraphicsClear(pGraphics, Colors.White)
    Else
        Call GdipCreateFromHWND(UserControl.hWnd, pGraphics)
    End If
    Call GdipSetSmoothingMode(pGraphics, SmoothingModeAntiAlias)
    
    pshadowOffset = 4
    pStartC = m_StartColour
    pEndColor = m_EndColor
    pRect.Top = 0
    pRect.Left = 0
    pRect.Height = UserControl.ScaleHeight / Screen.TwipsPerPixelY - pshadowOffset
    pRect.Width = UserControl.ScaleWidth / Screen.TwipsPerPixelX - pshadowOffset
    
     
    
    pPath = GetPath(pRect, 20)
    pPt1.x = pRect.Left
    pPt1.y = pRect.Top
    
    pPt2.x = pRect.Left
    pPt2.y = pRect.Height + 2
    
    Call GdipCreateLineBrushI(pPt1, pPt2, pStartC, pEndColor, WrapModeTile, pBrush)
    pPath = GetPath(pRect, 20)
    
    'Create shadow
    LSet prect2 = pRect
    Call Offset(prect2, pshadowOffset, pshadowOffset)
    
    pPath2 = GetPath(prect2, 20)
    Call GdipCreatePathGradientFromPath(pPath2, pBr2)
    Call GdipSetPathGradientCenterColor(pBr2, pStartC)
    Call GdipSetPathGradientSurroundColorsWithCount(pBr2, Colors.White, 2)
    
    
    'Create top water color to give "aqua" effect
    
    LSet prect3 = pRect
    Call InflateRectF(prect3, -5, -5)
    prect3.Height = pRect.Height / 3
    pPath3 = GetPath(prect3, 10)
    Call GdipCreateLineBrushFromRect(prect3, ColorSetAlpha(Colors.White, m_LightIntentisity), ColorSetAlpha(Colors.White, 0), LinearGradientModeVertical, WrapModeTileFlipXY, pBr3)
    
    GdipFillPath pGraphics, pBr2, pPath2 'draw shadow
    GdipFillPath pGraphics, pBrush, pPath 'draw main
    GdipFillPath pGraphics, pBr3, pPath3 'draw top bubble
    
    GdipDeleteBrush pBrush
    GdipDeletePath pPath
    
    GdipDeleteBrush pBr2
    GdipDeletePath pPath2
    
    GdipDeleteBrush pBr3
    GdipDeletePath pPath3
    
    GdipDeleteGraphics pGraphics
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_EndColor = PropBag.ReadProperty("EndColor", Colors.White)
    m_StartColour = PropBag.ReadProperty("StartColour", Colors.DarkGreen)
    m_LightIntentisity = PropBag.ReadProperty("LightIntentisity", 250)
End Sub

Private Sub UserControl_Resize()
    Set m_MemoryDc = New cMemDC
    m_MemoryDc.Height = UserControl.ScaleHeight / Screen.TwipsPerPixelY
    m_MemoryDc.Width = UserControl.ScaleWidth / Screen.TwipsPerPixelX
End Sub


Private Sub SetRegion()
    Dim pRect As RECTF
    Dim pRegion As Long
    pRect.Top = 0
    pRect.Left = 0
    pRect.Height = UserControl.ScaleHeight / Screen.TwipsPerPixelY
    pRect.Width = UserControl.ScaleWidth / Screen.TwipsPerPixelX
    pRegion = CreateRoundRectRgn(pRect.Left, pRect.Top, pRect.Width, pRect.Height, 20, 20)
    Call SetWindowRgn(UserControl.hWnd, pRegion, True)
    Call DeleteObject(pRegion)
End Sub

Private Sub UserControl_Terminate()
        Set m_MemoryDc = Nothing
        Call modGDIPlus.GDIPlusDispose
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("EndColor", m_EndColor, Colors.White)
    Call PropBag.WriteProperty("StartColour", m_StartColour, Colors.DarkGreen)
    Call PropBag.WriteProperty("LightIntentisity", m_LightIntentisity, 250)
End Sub


