Attribute VB_Name = "modGDIPlus"

Option Explicit
Dim m_token As Long

Public Enum SegOrientation
    Horizontal = 0
    Verticle = 1
End Enum

Public Enum SegColour
    sGreen = 0
    sBlue = 1
    sRed = 2
    sSelected = 3
End Enum

Public Enum DirectionCode
    L2R = 1
    R2L = 2
    U2D = 3
    D2U = 4
End Enum


Public Function GDIPlusCreate() As Boolean
Dim GpInput As GdiplusStartupInput
Dim token As Long
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) = Ok Then
      m_token = token
      GDIPlusCreate = True
   End If
End Function

Public Sub GDIPlusDispose()
   If Not (m_token = 0) Then
      GdiplusShutdown m_token
      m_token = 0
   End If
End Sub

Public Function ColorSetAlpha(ByVal lColor As Long, _
   ByVal Alpha As Byte) As Long

   Dim bytestruct       As COLORBYTES
   Dim result           As COLORLONG

   result.longval = lColor
   LSet bytestruct = result

   bytestruct.AlphaByte = Alpha

   LSet result = bytestruct
   ColorSetAlpha = result.longval
End Function

Public Function Gdi2VbColor(ByVal GdiColor As Long) As Long
   Dim Red              As Long
   Dim Green            As Long
   Dim Blue             As Long

   GdiColor = GdiColor And &HFFFFFF 'Strip alpha(if any)
   'Get components. Rebuild with Red and Blue swapped.
   Red = (GdiColor \ 65536) And &HFF
   Green = (GdiColor \ 256) And &HFF
   Blue = GdiColor And &HFF

   Gdi2VbColor = (Blue * 65536) + (Green * 256) + Red
End Function

Public Sub InflateRectF(rct As RECTF, ByVal x As Double, ByVal y As Double)
   rct.Left = rct.Left - x
   rct.Top = rct.Top - y
   rct.Width = rct.Width + x + x
   rct.Height = rct.Height + y + y
End Sub

Public Sub Offset(rct As RECTF, ByVal x As Double, ByVal y As Double)
   rct.Left = rct.Left + x
   rct.Top = rct.Top + y
   rct.Width = rct.Width + x + x
   rct.Height = rct.Height + y + y
End Sub


Public Function SetGradientRectF(ByRef TR As RECTF, _
   ByVal GradientMode As LinearGradientMode, _
   ByVal sngScale As Single, _
   ByVal MaintainAspectRatio As Boolean) As RECTF

   Dim sngW             As Single
   Dim sngH             As Single

   LSet SetGradientRectF = TR
   sngW = SetGradientRectF.Width / sngScale
   sngH = SetGradientRectF.Height / sngScale

   Select Case GradientMode
      Case LinearGradientModeHorizontal
         SetGradientRectF.Width = sngW
      Case LinearGradientModeVertical
         SetGradientRectF.Height = sngH
      Case LinearGradientModeForwardDiagonal, LinearGradientModeBackwardDiagonal
         If MaintainAspectRatio Then
            SetGradientRectF.Width = sngW
            SetGradientRectF.Height = sngH
         Else
            If SetGradientRectF.Width < SetGradientRectF.Height Then
               SetGradientRectF.Width = sngW
            Else
               SetGradientRectF.Width = sngH
            End If
            SetGradientRectF.Height = SetGradientRectF.Width
         End If
   End Select

End Function

Public Sub GetGradColours(ByVal pColourType As SegColour, ByRef pStartColour As Colors, ByRef pEndColour As Colors)

    Select Case pColourType
        Case SegColour.sBlue:
            pEndColour = LightBlue
            pStartColour = DarkBlue
        Case SegColour.sGreen:
            pEndColour = LightGreen
            pStartColour = DarkGreen
        Case SegColour.sRed:
            pEndColour = LightPink
            pStartColour = DarkRed
        Case SegColour.sSelected:
            pEndColour = LightGray
            pStartColour = DarkGray
    End Select
End Sub


Public Sub GetGradMode(ByVal pOrientation As SegOrientation, ByRef pGradMode As LinearGradientMode)
    If pOrientation = Horizontal Then
        pGradMode = LinearGradientModeVertical
    Else
        pGradMode = LinearGradientModeHorizontal
    End If
End Sub

Public Function GetPath(ByRef rc As RECTF, ByVal r As Integer) As Long
        Dim x As Integer
        Dim y As Integer
        Dim w As Integer
        Dim h As Integer
        Dim pPath As Long
        
        x = rc.Left
        y = rc.Top
        w = rc.Width
        h = rc.Height
        Call GdipCreatePath(FillModeWinding, pPath)
        Call GdipAddPathArc(pPath, x, y, r, r, 180, 90) 'Upper left corner
        Call GdipAddPathArc(pPath, x + w - r, y, r, r, 270, 90)      'Upper right corner
        Call GdipAddPathArc(pPath, x + w - r, y + h - r, r, r, 0, 90) 'Lower right corner
        Call GdipAddPathArc(pPath, x, y + h - r, r, r, 90, 90)      'Lower left corner
        GetPath = pPath
End Function

