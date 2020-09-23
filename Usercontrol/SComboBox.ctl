VERSION 5.00
Begin VB.UserControl SComboBox 
   CanGetFocus     =   0   'False
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   KeyPreview      =   -1  'True
   ScaleHeight     =   103
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   ToolboxBitmap   =   "SComboBox.ctx":0000
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   840
      Top             =   405
   End
End
Attribute VB_Name = "SComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
'*                   Version 1.0.5                     *'
'*******************************************************'
'* Control:       SComboBox                            *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Collaboration: fred.cpp                             *'
'*                                                     *'
'*                So many thanks for his contribution  *'
'*                for this project, some styles and    *'
'*                Traduction to English of some        *'
'*                comments.                            *'
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
Option Explicit

Public Event MouseEnter()
Public Event MouseLeave()

'****************************'
'* English: Private Type.   *'
'* Español: Tipos Privados. *'
'****************************'

Private Type GRADIENT_RECT
    UpperLeft   As Long
    LowerRight  As Long
End Type

Private Type POINTAPI
    x           As Long
    y           As Long
End Type

Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

Private Type RGB
    Red         As Integer
    Green       As Integer
    Blue        As Integer
End Type

Private Type TRIVERTEX
    x             As Long
    y             As Long
    Red           As Integer
    Green         As Integer
    Blue          As Integer
    Alpha         As Integer
End Type

'*********************************************'
'* English: Public Enum of Control.          *'
'* Español: Enumeración Publica del control. *'
'*********************************************'

'* English: Appearance Combo.
'* Español: Apariencias del Combo.

Public Enum ComboAppearance
    Office = &H1             '* By fred.cpp & HACKPRO TM.
    Win98 = &H2              '* By fred.cpp.
    WinXp = &H3              '* By fred.cpp & HACKPRO TM.
    Mac = &H4                '* By fred.cpp & HACKPRO TM.
    JAVA = &H5               '* By fred.cpp.
    [Explorer Bar] = &H6     '* By HACKPRO TM.
    [GradientV] = &H7        '* By HACKPRO TM.
    [GradientH] = &H8        '* By HACKPRO TM.
    [Light Blue] = &H9       '* By HACKPRO TM.
    [Chocolate] = &HA        '* By HACKPRO TM.
    [Button Download] = &HB  '* By HACKPRO TM.
End Enum

'* English: Appearance standard style Office.
'* Español: Apariencias estándares del estilo Office.

Public Enum ComboOfficeAppearance
    [Office Xp] = &H0       '* By HACKPRO TM.
    [Office 2000] = &H1     '* By fred.cpp.
    [Office 2003] = &H2     '* By HACKPRO TM.
End Enum

'* English: Appearance standard style Xp.
'* Español: Apariencias estándares del estilo Xp.

Public Enum ComboXpAppearance
    [Windows Themed] = &H0  '* By fred.cpp
    Aqua = &H1              '* By HACKPRO TM.
    [Olive Green] = &H2     '* By HACKPRO TM.
    Silver = &H3            '* By HACKPRO TM.
    TasBlue = &H4           '* By HACKPRO TM.
    Gold = &H5              '* By HACKPRO TM.
    Blue = &H6              '* By HACKPRO TM.
    CustomXP = &H7          '* By HACKPRO TM.
End Enum

'********************************'
'* English: Private variables.  *'
'* Español: Variables privadas. *'
'********************************'
Private ControlEnabled          As Boolean
Private cValor                  As Long
Private g_Font                  As StdFont
Private iFor                    As Long
Private isFailedXP              As Boolean
Private IsOver                  As Boolean
Private isPicture               As Boolean
Private m_btnRect               As RECT
Private m_StateG                As Integer
Private myAppearanceCombo       As ComboAppearance
Private myArrowColor            As OLE_COLOR
Private myBackColor             As OLE_COLOR
Private myDisabledColor         As OLE_COLOR
Private myGradientColor1        As OLE_COLOR
Private myGradientColor2        As OLE_COLOR
Private myHighLightBorderColor  As OLE_COLOR
Private myHighLightColorText    As OLE_COLOR
Private myMouseIcon             As StdPicture
Private myMousePointer          As MousePointerConstants
Private myNormalBorderColor     As OLE_COLOR
Private myNormalColorText       As OLE_COLOR
Private myOfficeAppearance      As ComboOfficeAppearance
Private mySelectBorderColor     As OLE_COLOR
Private myText                  As String
Private myXpAppearance          As ComboXpAppearance
Private NoDown                  As Boolean
Private NoShow                  As Boolean
Private RGBColor                As RGB
Private tempBorderColor         As OLE_COLOR
Private tmpC1                   As Long
Private tmpC2                   As Long
Private tmpC3                   As Long
Private tmpcolor                As Long
Private UserText                As String

'***************************************'
'* English: Constant declares.         *'
'* Español: Declaración de Constantes. *'
'***************************************'
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_RECT = (&H1 Or &H2 Or &H4 Or &H8)
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
Private Const COLOR_GRAYTEXT As Long = 17
Private Const COLOR_HOTLIGHT As Long = 26
Private Const COLOR_INACTIVECAPTIONTEXT As Long = 19
Private Const COLOR_WINDOW = 5
Private Const defAppearanceCombo = 1
Private Const defArrowColor = &HC56A31
Private Const defDisabledColor = &H808080
Private Const defGradientColor1 = &HDAB278
Private Const defGradientColor2 = &HFFDD9E
Private Const defHighLightBorderColor = &HC56A31
Private Const defHighLightColorText = &HFFFFFF
Private Const defNormalBorderColor = &HDEEDEF
Private Const defNormalColorText = &HC56A31
Private Const defListColor = &HFFFFFF
Private Const defOfficeAppearance = 0
Private Const defSelectBorderColor = &HC56A31
Private Const DT_LEFT                As Long = &H0
Private Const DT_SINGLELINE          As Long = &H20
Private Const DT_VCENTER             As Long = &H4
Private Const DT_WORD_ELLIPSIS       As Long = &H40000
Private Const EDGE_RAISED = (&H1 Or &H4)
Private Const EDGE_SUNKEN = (&H2 Or &H8)
Private Const GRADIENT_FILL_RECT_H   As Long = &H0
Private Const GRADIENT_FILL_RECT_V   As Long = &H1
Private Const Version                As String = "SComboBox 1.0.5 By HACKPRO TM"

Public Event Click()
Public Event CloseList()
Public Event Change()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

'* Declares for Unicode support.
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
End Type

Private mWindowsNT   As Boolean

'**********************************'
'* English: Calls to the API's.   *'
'* Español: Llamadas a los API's. *'
'**********************************'
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function CreatePen Lib "gdi32" ( _
        ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" ( _
        lpPoint As POINTAPI, _
        ByVal nCount As Long, _
        ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" ( _
        ByVal hDC As Long, _
        qrc As RECT, _
        ByVal edge As Long, _
        ByVal grfFlags As Long) As Long
'Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal
'   hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long,
'   ByVal X
'   As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long
Private Declare Function DrawTextA Lib "user32" ( _
        ByVal hDC As Long, _
        ByVal lpStr As String, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" ( _
        ByVal hDC As Long, _
        ByVal lpStr As Long, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" ( _
        ByVal hTheme As Long, _
        ByVal lhDC As Long, _
        ByVal iPartId As Long, _
        ByVal iStateId As Long, _
        pRect As RECT, _
        pClipRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As
'   Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As
'   Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long,
'   ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" ( _
        lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GradientFillRect Lib "msimg32" _
        Alias "GradientFill" ( _
        ByVal hDC As Long, _
        pVertex As TRIVERTEX, _
        ByVal dwNumVertex As Long, _
        pMesh As GRADIENT_RECT, _
        ByVal dwNumMesh As Long, _
        ByVal dwMode As Long) As Long
Private Declare Function LineTo Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
        ByVal lOleColor As Long, _
        ByVal lHPalette As Long, _
        lColorRef As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" ( _
        ByVal hWnd As Long, _
        ByVal pszClassList As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As
'   Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" ( _
        lpRect As RECT, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long,
'   ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As
'   Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As
'   Long)
'   As Long
Private Declare Function SetWindowRgn Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hrgn As Long, _
        ByVal bRedraw As Boolean) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As
'   Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long,
'   ByVal
'   ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub APIFillRect(ByVal hDC As Long, ByRef rc As RECT, ByVal Color As Long)

  Dim NewBrush As Long

    '* English: The FillRect function fills a rectangle by using the specified brush. This function
    '   includes the left and Top borders, but excludes the right and bottom borders of the
    '   rectangle.
    '* Español: Pinta el rectángulo de un objeto.
    NewBrush& = CreateSolidBrush(Color&)
    Call FillRect(hDC&, rc, NewBrush&)
    Call DeleteObject(NewBrush&)

End Sub

Private Sub APILine(ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal x2 As Long, _
                    ByVal y2 As Long, _
                    ByVal lColor As Long)

  Dim PT As POINTAPI
  Dim hPen As Long
  Dim hPenOld As Long

    '* English: Use the API LineTo for Fast Drawing.
    '* Español: Pinta líneas de forma sencilla y rápida.
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hDC, hPen)
    Call MoveToEx(UserControl.hDC, X1, Y1, PT)
    Call LineTo(UserControl.hDC, x2, y2)
    Call SelectObject(hDC, hPenOld)
    Call DeleteObject(hPen)

End Sub

Private Function APIRectangle(ByVal hDC As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal W As Long, _
                              ByVal H As Long, _
                              Optional ByVal lColor As OLE_COLOR = -1) As Long

  Dim hPen As Long
  Dim hPenOld As Long
  Dim PT   As POINTAPI

    '* English: Paint a rectangle using API.
    '* Español: Pinta el rectángulo de un Objeto.
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hDC, hPen)
    Call MoveToEx(hDC, x, y, PT)
    Call LineTo(hDC, x + W, y)
    Call LineTo(hDC, x + W, y + H)
    Call LineTo(hDC, x, y + H)
    Call LineTo(hDC, x, y)
    Call SelectObject(hDC, hPenOld)
    Call DeleteObject(hPen)

End Function

Public Property Get AppearanceCombo() As ComboAppearance

    '* English: Sets/Gets the style of the Combo.
    '* Español: Devuelve o establece el estilo del Combo.
    AppearanceCombo = myAppearanceCombo

End Property

Public Property Let AppearanceCombo(ByVal new_Style As ComboAppearance)

    myAppearanceCombo = IIf(new_Style <= 0, 1, new_Style)
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("AppearanceCombo")
    Refresh

End Property

Public Property Let ArrowColor(ByVal New_Color As OLE_COLOR)

    myArrowColor = ConvertSystemColor(New_Color)
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("ArrowColor")
    Refresh

End Property

Public Property Get ArrowColor() As OLE_COLOR

    '* English: Sets/Gets the color of the arrow.
    '* Español: Devuelve o establece el color de la flecha.
    ArrowColor = myArrowColor

End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)

    myBackColor = ConvertSystemColor(GetLngColor(New_Color))
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("BackColor")
    Refresh

End Property

Public Property Get BackColor() As OLE_COLOR

    '* English: Sets/Gets the color of the Usercontrol.
    '* Español: Devuelve o establece el color del Usercontrol.
    BackColor = myBackColor

End Property

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long)

On Error GoTo BlendColors_Error
    BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) _
        + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ _
        &H10000) And &HFF)) / 2)
    Exit Function
BlendColors_Error:

End Function

Public Sub ClosedList()

    NoDown = False
    RaiseEvent CloseList
    
End Sub

Private Function ConvertSystemColor(ByVal theColor As Long) As Long

    '* English: Convert Long to System Color.
    '* Español: Convierte un long en un color del sistema.
    Call OleTranslateColor(theColor, 0, ConvertSystemColor)

End Function

Private Function CreateMacOSXRegion() As Long

  Dim pPoligon(8) As POINTAPI
  Dim lw As Long
  Dim lh As Long

    '* English: Create a nonrectangular region for the MAC OS X Style.
    '* Español: Crea el Estilo MAC OS X.
    lw = UserControl.ScaleWidth
    lh = UserControl.ScaleHeight
    pPoligon(0).x = 0:      pPoligon(0).y = 2
    pPoligon(1).x = 2:      pPoligon(1).y = 0
    pPoligon(2).x = lw - 2: pPoligon(2).y = 0
    pPoligon(3).x = lw:     pPoligon(3).y = 2
    pPoligon(4).x = lw:     pPoligon(4).y = lh - 5
    pPoligon(5).x = lw - 6: pPoligon(5).y = lh
    pPoligon(6).x = 3:      pPoligon(6).y = lh
    pPoligon(7).x = 0:      pPoligon(7).y = lh - 3
    CreateMacOSXRegion = CreatePolygonRgn(pPoligon(0), 8, 1)

End Function

Public Property Get DisabledColor() As OLE_COLOR

    '* English: Sets/Gets the color of the disabled text.
    '* Español: Devuelve o establece el color del texto deshabilitado.
    DisabledColor = ShiftColorOXP(myDisabledColor, 94)

End Property

Public Property Let DisabledColor(ByVal New_Color As OLE_COLOR)

    myDisabledColor = ConvertSystemColor(GetLngColor(New_Color))
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("DisabledColor")
    Refresh

End Property

Private Sub DrawAppearance(Optional ByVal Style As ComboAppearance = 1, _
                           Optional ByVal m_State As Integer = 1)

  Dim isText    As String
  Dim RText     As RECT
  Dim m_lRegion As Long
  Dim isH       As Integer

    '* English: Draw appearance of the control.
    '* Español: Dibuja la apariencia del control.
    Cls
    Call SetRect(m_btnRect, UserControl.ScaleWidth - 18, 2, UserControl.ScaleWidth - 2, _
        UserControl.ScaleHeight - 2)
    AutoRedraw = True
    FillStyle = 1
    m_StateG = m_State
    isH = 0
    If (Style <> 6) Then UserControl.BackColor = myBackColor
    On Error Resume Next

    Select Case Style
    Case 1
        Call DrawOfficeButton(myOfficeAppearance)

    Case 2
        '* English: Style Windows 98.
        '* Español: Estilo Windows 98.
        Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
            EDGE_SUNKEN)
        Call APIFillRect(UserControl.hDC, m_btnRect, GetSysColor(COLOR_BTNFACE))
        tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
        Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 3, EDGE_SUNKEN, _
            EDGE_RAISED))
        Call DrawStandardArrow(m_btnRect, ArrowColor)

    Case 3
        '* English: Style Windows Xp.
        '* Español: Estilo Windows Xp.
        If (myXpAppearance = 1) Then     '* Aqua.
            tmpcolor = &HB99D7F
        ElseIf (myXpAppearance = 2) Then '* Olive Green.
            tmpcolor = &H94CCBC
        ElseIf (myXpAppearance = 3) Then '* Silver.
            tmpcolor = &HA29594
        ElseIf (myXpAppearance = 4) Then '* TasBlue.
            tmpcolor = &HF09F5F
        ElseIf (myXpAppearance = 5) Then '* Gold.
            tmpcolor = &HBFE7F0
        ElseIf (myXpAppearance = 6) Then '* Blue.
            tmpcolor = ShiftColorOXP(&HA0672F, 123)
        ElseIf (myXpAppearance = 7) Or (myXpAppearance = 0) Then '* Custom.

            If (m_StateG = 1) Then
                tmpcolor = NormalBorderColor
            ElseIf (m_StateG = 2) Then
                tmpcolor = HighLightBorderColor
            ElseIf (m_StateG = 3) Then
                tmpcolor = SelectBorderColor
            End If

        End If
        Call DrawWinXPButton(myXpAppearance, tmpcolor)

        If (myXpAppearance <> 0) Or (isFailedXP = True) Then
            Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
            Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
            Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, _
                UserControl.BackColor)
            Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, _
                UserControl.BackColor)
        End If

    Case 4
        '* English: Style MAC.
        '* Español: Estilo MAC.
        isH = 2
        Call DrawMacOSXCombo

    Case 5
        '* English: Style JAVA.
        '* Español: Estilo JAVA.
        tmpcolor = ShiftColorOXP(NormalBorderColor, 52)
        tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
        Call DrawJavaBorder(0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, _
            GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
        Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 2, tmpcolor, IIf(m_StateG <> -1, _
            NormalBorderColor, ShiftColorOXP(NormalBorderColor, 192))))
        Call DrawJavaBorder(m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, _
            m_btnRect.Bottom - m_btnRect.Top - 1, GetSysColor(COLOR_BTNSHADOW), _
            GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
        Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(ArrowColor, 166), _
            ArrowColor))

    Case 6
        Call DrawExplorerBarButton(m_StateG)

    Case 7
        Call DrawGradientButton(1)

    Case 8
        Call DrawGradientButton(2)

    Case 9
        Call DrawLightBlueButton

    Case 10
        Call DrawChocolateButton

    Case 11
        Call DrawButtonDownload
    End Select

    If (Style = 4) Then
        If (m_lRegion <> 0) Then Call DeleteObject(m_lRegion)
        m_lRegion = CreateMacOSXRegion
        Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
    Else
        m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
        Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
    End If

    cValor = 5

    With UserControl
        isText = myText
        .CurrentX = cValor
        .CurrentY = Int(UserControl.ScaleHeight / 2) - 7
        Set .Font = g_Font
        If (Enabled = False) Then
            Call SetTextColor(.hDC, DisabledColor)
        Else
            Call SetTextColor(.hDC, NormalColorText)
        End If

        RText = m_btnRect
        RText.Left = cValor
        RText.Top = -1
        RText.Bottom = UserControl.ScaleHeight
        RText.Right = UserControl.ScaleWidth - 20

        If (mWindowsNT = True) Then
            Call DrawTextW(.hDC, StrPtr(isText), Len(isText), RText, DT_VCENTER Or DT_LEFT Or _
                DT_SINGLELINE Or DT_WORD_ELLIPSIS)
        Else
            Call DrawTextA(.hDC, isText, Len(isText), RText, DT_VCENTER Or DT_LEFT Or _
                DT_SINGLELINE Or DT_WORD_ELLIPSIS)
        End If
        
    End With

End Sub

Private Sub DrawButtonDownload()

    '* English: Draw Button Download appearance.
    '* Español: Crea la apariencia de un Botón de Descarga.
    cValor = IIf(m_StateG = -1, ShiftColorOXP(&H92603C), &H92603C)
    tempBorderColor = cValor
    tmpC3 = IIf(m_StateG = -1, ShiftColorOXP(&HE0C6AE), &HE0C6AE)

    If (m_StateG = 1) Or (m_StateG = 3) Then
        tmpC1 = &HBE8F63
        tmpC2 = &HE8DBCB
        tmpcolor = ArrowColor
    ElseIf (m_StateG = 2) Then
        tmpC1 = ShiftColorOXP(&HBE8F63, 49)
        tmpC2 = ShiftColorOXP(&HE8DBCB, 49)
        tmpcolor = ShiftColorOXP(ArrowColor, 89)
    Else
        tmpC1 = ShiftColorOXP(&HBE8F63)
        tmpC2 = ShiftColorOXP(&HE8DBCB)
        tmpcolor = ShiftColorOXP(&HC0C0C0, 85)
    End If

    Call DrawGradient(UserControl.hDC, m_btnRect.Left + 2, m_btnRect.Top, m_btnRect.Right, _
        m_btnRect.Bottom, tmpC1, tmpC2, 1)
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        cValor)
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, _
        UserControl.ScaleHeight, ShiftColorOXP(cValor, 5))
    Call DrawXpArrow(tmpcolor)
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, _
        Int(UserControl.ScaleHeight / 2) + 4, UserControl.ScaleWidth - 6, _
        Int(UserControl.ScaleHeight / 2) + 4, tmpcolor)
    Call DrawShadow(tmpC3, tmpC3, False)

End Sub

Private Sub DrawChocolateButton()

    '* English: Chocolate Style.
    '* Español: Estilo Chocolate.
    cValor = IIf(m_StateG = -1, ShiftColorOXP(&H4A464B), &H4A464B)
    tempBorderColor = cValor
    tmpC3 = &HFFFFFF

    If (m_StateG = 1) Or (m_StateG = 3) Then
        tmpC1 = &H686567
        tmpC2 = ShiftColorOXP(&H292929, 89)
        tmpcolor = &H0
    ElseIf (m_StateG = 2) Then
        tmpC1 = ShiftColorOXP(&H686567, 89)
        tmpC2 = ShiftColorOXP(&H292929, 178)
        tmpcolor = ShiftColorOXP(&H0, 89)
    Else
        tmpC1 = ShiftColorOXP(&H838181)
        tmpC2 = ShiftColorOXP(&H292929)
        tmpcolor = ShiftColorOXP(&H0)
    End If

    Call DrawGradient(UserControl.hDC, m_btnRect.Left + 2, m_btnRect.Top, m_btnRect.Right, _
        m_btnRect.Bottom, tmpC1, tmpC2, 2)
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        cValor)
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 0, 19, _
        UserControl.ScaleHeight, ShiftColorOXP(cValor, 5))
    Call DrawShadow(tmpC3, tmpcolor, False)
    m_btnRect.Bottom = m_btnRect.Bottom / 2 + 4
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, m_btnRect.Bottom + _
        2, UserControl.ScaleWidth - 5, m_btnRect.Bottom + 2, tmpcolor)
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, m_btnRect.Bottom + _
        3, UserControl.ScaleWidth - 5, m_btnRect.Bottom + 3, tmpcolor)

    For iFor = 4 To 7
        Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor - 1, _
            m_btnRect.Bottom - iFor + 5, UserControl.ScaleWidth - iFor - 1, m_btnRect.Bottom - _
            iFor + 5, tmpcolor)
        Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor - 1, _
            m_btnRect.Bottom + iFor, UserControl.ScaleWidth - (iFor + 1), m_btnRect.Bottom + iFor, _
            tmpcolor)
        Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor, _
            m_btnRect.Bottom - iFor + 5, UserControl.ScaleWidth - (iFor + 2), m_btnRect.Bottom - _
            iFor + 5, &HFFFFFF)
        Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor, _
            m_btnRect.Bottom + iFor, UserControl.ScaleWidth - (iFor + 2), m_btnRect.Bottom + iFor, _
            &HFFFFFF)
    Next

End Sub

Private Sub DrawCtlEdge(ByVal hDC As Long, _
                        ByVal x As Single, _
                        ByVal y As Single, _
                        ByVal W As Single, _
                        ByVal H As Single, _
                        Optional ByVal Style As Long = EDGE_RAISED, _
                        Optional ByVal Flags As Long = BF_RECT)

  Dim R As RECT

    '* English: The DrawEdge function draws one or more edges of rectangle. using the specified
    '   coords.
    '* Español: Dibuja uno ó más bordes del rectángulo.

    With R
        .Left = x
        .Top = y
        .Right = x + W
        .Bottom = y + H
    End With

    Call DrawEdge(hDC, R, Style, Flags)

End Sub

Private Sub DrawCtlEdgeByRect(ByVal hDC As Long, _
                              ByRef RT As RECT, _
                              Optional ByVal Style As Long = EDGE_RAISED, _
                              Optional ByVal Flags As Long = BF_RECT)

    '* English: Draws the edge in a rect.
    '* Español: Colorea uno ó más bordes del rectángulo del Control.
    Call DrawEdge(hDC, RT, Style, Flags)

End Sub

Private Sub DrawExplorerBarButton(ByVal m_StateG As Long)

  Dim isBackColor As OLE_COLOR

    '* English: Style ExplorerBar.
    '* Español: Estilo ExplorerBar.
    isBackColor = ShiftColorOXP(&HDEEAF0, 184)
    UserControl.BackColor = isBackColor
    Call DrawRectangleBorder(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 2, _
        UserControl.ScaleHeight - 2, &HEAF3F7)

    If (m_StateG = 1) Then
        cValor = ShiftColorOXP(&HB6BFC3, 91)
        iFor = &HEAF3F7
        tmpcolor = ShiftColorOXP(&HB6BFC3, 162)
    ElseIf (m_StateG = 2) Then
        cValor = ShiftColorOXP(&HB6BFC3, 31)
        iFor = &HDCEBF1
        tmpcolor = ShiftColorOXP(&HB6BFC3, 132)
    ElseIf (m_StateG = 3) Then
        cValor = ShiftColorOXP(&HB6BFC3, 21)
        iFor = &HCEE3EC
        tmpcolor = ShiftColorOXP(&HB6BFC3, 112)
        tempBorderColor = ShiftColorOXP(&HB6BFC3, 21)
    Else
        UserControl.BackColor = ShiftColorOXP(&HEAF3F7, 124)
        cValor = ShiftColorOXP(&HB6BFC3, 84)
        tmpC1 = ShiftColorOXP(&HEAF3F7, 124)
        iFor = ShiftColorOXP(&HEAF3F7, 123)
        tmpcolor = ShiftColorOXP(&HB6BFC3, 132)
    End If

    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        cValor)
    If (m_StateG = -1) Then Call DrawRectangleBorder(UserControl.hDC, 1, 1, UserControl.ScaleWidth _
        - 2, UserControl.ScaleHeight - 2, tmpC1)
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, _
        UserControl.ScaleHeight - 4, iFor, False)
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, _
        UserControl.ScaleHeight - 4, tmpcolor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, _
        UserControl.BackColor)
    Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, _
        UserControl.BackColor)
    Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 196), ArrowColor))

End Sub

Private Sub DrawGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal X1 As Long, _
                         ByVal Y1 As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal Direction As Integer)

  Dim Vert(1) As TRIVERTEX
  Dim gRect As GRADIENT_RECT

    '* English: Draw a gradient in the selected coords and hDC.
    '* Español: Dibuja el objeto en forma degradada.
    Call LongToRGB(Color1)

    With Vert(0)
        .x = x
        .y = y
        .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
        .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
        .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
        .Alpha = 1
    End With

    Call LongToRGB(Color2)

    With Vert(1)
        .x = X1
        .y = Y1
        .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
        .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
        .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
        .Alpha = 0
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    If (Direction = 1) Then
        Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)
    Else
        Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H)
    End If

End Sub

Private Sub DrawGradientButton(ByVal WhatGradient As Long)

    '* English: Draw a Vertical or Horizontal Gradient style appearance.
    '* Español: Dibuja la apariencia degradada bien sea vertical ó horizontal.

    If (m_StateG = 1) Then
        tmpcolor = ShiftColorOXP(&HC56A31, 133)
        cValor = GetLngColor(&HFFFFFF)
        iFor = GetLngColor(&HD8CEC5)
    ElseIf (m_StateG = 2) Then
        tmpcolor = ShiftColorOXP(&HC56A31, 113)
        cValor = GetLngColor(&HFFFFFF)
        iFor = GetLngColor(&HD6BEB5)
    ElseIf (m_StateG = 3) Then
        tmpcolor = ShiftColorOXP(&HC56A31, 93)
        cValor = GetLngColor(&HFFFFFF)
        iFor = GetLngColor(&HB3A29B)
        tempBorderColor = tmpcolor
    Else
        tmpcolor = CLng(ShiftColorOXP(&H0&))
        cValor = GetLngColor(&HC0C0C0)
        iFor = GetLngColor(&HFFFFFF)
    End If

    Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, _
        m_btnRect.Bottom, cValor, iFor, WhatGradient)
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        tmpcolor)
    Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 2, 16, _
        UserControl.ScaleHeight - 4, tmpcolor, True)
    Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))

End Sub

Private Sub DrawJavaBorder(ByVal x As Long, _
                           ByVal y As Long, _
                           ByVal W As Long, _
                           ByVal H As Long, _
                           ByVal lColorShadow As Long, _
                           ByVal lColorLight As Long, _
                           ByVal lColorBack As Long)

    '* English: Draw the edge with a JAVA style.
    '* Español: Dibuja el borde estilo JAVA.
    Call APIRectangle(UserControl.hDC, x, y, W - 1, H - 1, lColorShadow)
    Call APIRectangle(UserControl.hDC, x + 1, y + 1, W - 1, H - 1, lColorLight)
    Call SetPixel(UserControl.hDC, x, y + H, lColorBack)
    Call SetPixel(UserControl.hDC, x + W, y, lColorBack)
    Call SetPixel(UserControl.hDC, x + 1, y + H - 1, BlendColors(lColorLight, lColorShadow))
    Call SetPixel(UserControl.hDC, x + W - 1, y + 1, BlendColors(lColorLight, lColorShadow))

End Sub

Private Sub DrawLightBlueButton()

  Dim PT      As POINTAPI
  Dim cx As Long
  Dim cy As Long
  Dim hPenOld As Long
  Dim hPen    As Long

    '* English: Style LightBlue.
    '* Español: Estilo LightBlue.

    If (m_StateG = 1) Or (m_StateG = 3) Then
        cValor = GetLngColor(&HFFFFFF)
        iFor = GetLngColor(&HA87057)
        tmpcolor = &HA69182
        tempBorderColor = tmpcolor
    ElseIf (m_StateG = 2) Then
        cValor = GetLngColor(&HFFFFFF)
        iFor = GetLngColor(&HCFA090)
        tmpcolor = &HAF9080
    Else
        cValor = GetLngColor(&HFFFFFF)
        iFor = ShiftColorOXP(GetLngColor(&HA87057))
        tmpcolor = ShiftColorOXP(&HA69182, 146)
    End If

    Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, _
        m_btnRect.Bottom + 1, cValor, iFor, 1)
    Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        tmpcolor)

    If (m_StateG = 2) Then
        Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 17, _
            UserControl.ScaleHeight, &H53969F)
        Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 16, 1, 15, _
            UserControl.ScaleHeight - 2, &H92C4D8)
        tmpcolor = &H3EB4DE
    ElseIf (m_StateG <> -1) Then
        Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 17, 0, 19, _
            UserControl.ScaleHeight, ShiftColorOXP(tmpcolor, 5))
        tmpcolor = ArrowColor
    End If

    cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
    cy = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 + 2
    hPen = CreatePen(0, 1, IIf(m_StateG <> -1, tmpcolor, ShiftColorOXP(&HC0C0C0, 97)))
    hPenOld = SelectObject(UserControl.hDC, hPen)
    Call MoveToEx(UserControl.hDC, cx - 3, cy - 1, PT)
    Call LineTo(UserControl.hDC, cx + 1, cy - 1)
    Call LineTo(UserControl.hDC, cx, cy)
    Call LineTo(UserControl.hDC, cx - 2, cy)
    Call LineTo(UserControl.hDC, cx, cy + 2)
    Call SelectObject(hDC, hPenOld)
    Call DeleteObject(hPen)
    hPen = CreatePen(0, 1, IIf(m_StateG <> -1, tmpcolor, ShiftColorOXP(&HC0C0C0, 97)))
    hPenOld = SelectObject(UserControl.hDC, hPen)
    cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 3
    Call MoveToEx(UserControl.hDC, cx - 4, cy - 3, PT)
    Call LineTo(UserControl.hDC, cx, cy - 3)
    Call LineTo(UserControl.hDC, cx - 2, cy - 5)
    Call LineTo(UserControl.hDC, cx - 3, cy - 4)
    Call LineTo(UserControl.hDC, cx - 1, cy - 3)
    Call SelectObject(hDC, hPenOld)
    Call DeleteObject(hPen)

End Sub

Private Sub DrawMacOSXCombo()

  Dim PT      As POINTAPI
  Dim cy   As Long
  Dim cx      As Long
  Dim Color1 As Long
  Dim ColorG As Long
  Dim hPen    As Long
  Dim hPenOld As Long
  Dim Color2 As Long
  Dim Color3 As Long
  Dim ColorH As Long
  Dim Color4  As Long
  Dim Color5   As Long
  Dim Color6 As Long
  Dim Color7 As Long
  Dim ColorI As Long
  Dim Color8  As Long
  Dim Color9   As Long
  Dim ColorA As Long
  Dim ColorB As Long
  Dim ColorC  As Long
  Dim ColorD   As Long
  Dim ColorE As Long
  Dim ColorF As Long

    '* English: Draw the Mac OS X combo (this is a cool style!).
    '* Español: Dibujar el combo estilo Mac OS X (este es un estilo chevere).
    m_btnRect.Left = m_btnRect.Left - 4
    tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
    '* English: Button gradient Top.
    ColorA = &HA0A0A0
    UserControl.BackColor = myBackColor

    If (m_StateG = 1) Then
        Color1 = ShiftColorOXP(&HFDF2C3, 9)
        Color2 = ShiftColorOXP(&HDE8B45, 9)
        Color3 = ShiftColorOXP(&HDD873E, 9)
        Color4 = ShiftColorOXP(&HB33A01, 9)
        Color5 = ShiftColorOXP(&HE9BD96, 9)
        Color6 = ShiftColorOXP(&HB9B2AD, 9)
        Color7 = ShiftColorOXP(&H968A82, 9)
        Color8 = ShiftColorOXP(&HA25022, 9)
        Color9 = ShiftColorOXP(&HB8865E, 9)
        ColorB = ShiftColorOXP(&HDFBC86, 9)
        ColorC = ShiftColorOXP(&HFFBA77, 9)
        ColorD = ShiftColorOXP(&HE3D499, 9)
        ColorE = ShiftColorOXP(&HFFD996, 9)
        ColorF = ShiftColorOXP(&HE1A46D, 9)
        ColorG = ShiftColorOXP(&HCBA47B, 9)
        ColorH = ShiftColorOXP(&HDFDFDF, 9)
        ColorI = ShiftColorOXP(&HD0D0D0, 9)
    ElseIf (m_StateG = 2) Then
        Color1 = ShiftColorOXP(&HFDF2C3, 89)
        Color2 = ShiftColorOXP(&HDE8B45, 89)
        Color3 = ShiftColorOXP(&HDD873E, 89)
        Color4 = ShiftColorOXP(&HB33A01, 99)
        Color5 = ShiftColorOXP(&HE9BD96, 109)
        Color6 = ShiftColorOXP(&HB9B2AD, 109)
        Color7 = ShiftColorOXP(&H968A82, 109)
        Color8 = ShiftColorOXP(&HA25022, 109)
        Color9 = ShiftColorOXP(&HB8865E, 109)
        ColorB = ShiftColorOXP(&HDFBC86, 109)
        ColorC = ShiftColorOXP(&HFFBA77, 109)
        ColorD = ShiftColorOXP(&HE3D499, 109)
        ColorE = ShiftColorOXP(&HFFD996, 109)
        ColorF = ShiftColorOXP(&HE1A46D, 109)
        ColorG = ShiftColorOXP(&HCBA47B, 109)
        ColorH = ShiftColorOXP(&HDFDFDF, 109)
        ColorI = ShiftColorOXP(&HD0D0D0, 109)
    ElseIf (m_StateG = 3) Then
        Color1 = ShiftColorOXP(&HFDF2C3, 15)
        Color2 = ShiftColorOXP(&HDE8B45, 15)
        Color3 = ShiftColorOXP(&HDD873E, 15)
        Color4 = ShiftColorOXP(&HB33A01, 15)
        Color5 = ShiftColorOXP(&HE9BD96, 15)
        Color6 = ShiftColorOXP(&HB9B2AD, 15)
        Color7 = ShiftColorOXP(&H968A82, 15)
        Color8 = ShiftColorOXP(&HA25022, 15)
        Color9 = ShiftColorOXP(&HB8865E, 15)
        ColorB = ShiftColorOXP(&HDFBC86, 15)
        ColorC = ShiftColorOXP(&HFFBA77, 15)
        ColorD = ShiftColorOXP(&HE3D499, 15)
        ColorE = ShiftColorOXP(&HFFD996, 15)
        ColorF = ShiftColorOXP(&HE1A46D, 15)
        ColorG = ShiftColorOXP(&HCBA47B, 15)
        ColorH = ShiftColorOXP(&HDFDFDF, 15)
        ColorI = ShiftColorOXP(&HD0D0D0, 15)
    Else
        Color1 = ShiftColorOXP(&H808080, 195)
        Color2 = ShiftColorOXP(&H808080, 135)
        Color3 = ShiftColorOXP(&H808080, 135)
        Color4 = ShiftColorOXP(&H808080, 5)
        Color5 = Color1
        Color6 = GetLngColor(Parent.BackColor)
        Color7 = Color6
        Color8 = ShiftColorOXP(&H808080, 65)
        Color9 = Color6
        ColorA = Color6
        ColorB = Color4
        ColorC = Color4
        ColorD = Color4
        ColorE = Color4
        ColorF = Color4
        ColorG = Color4
        ColorH = Color6
        ColorI = Color6
    End If

    Call DrawVGradient(Color1, Color2, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, _
        UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 3)
    '* English: Button gradient bottom.
    Call DrawVGradient(Color3, Color1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, _
        UserControl.ScaleHeight / 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight * 2 / 3 - _
        4)
    '* English: Lines for the text area.
    Call APILine(2, 0, UserControl.ScaleWidth - 3, 0, &HA1A1A1)
    Call APILine(1, 0, 1, UserControl.ScaleHeight - 3, &HA1A1A1)
    '* English: Left shadow.

    If (m_StateG <> -1) Then
        Call DrawVGradient(ColorH, &HBBBBBB, 0, 0, 1, 3)
        Call DrawVGradient(&HBBBBBB, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
        Call DrawVGradient(ColorA, &HBBBBBB, 0, UserControl.ScaleHeight / 2, 1, _
            UserControl.ScaleHeight / 2 - 5)
        Call DrawVGradient(&HBBBBBB, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
    Else
        Call DrawVGradient(ColorH, ColorH, 0, 0, 1, 3)
        Call DrawVGradient(ColorA, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
        Call DrawVGradient(ColorA, ColorA, 0, UserControl.ScaleHeight / 2, 1, _
            UserControl.ScaleHeight / 2 - 5)
        Call DrawVGradient(ColorH, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
    End If

    '* English: Bottom shadows.
    Call APILine(1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, _
        UserControl.ScaleHeight - 3, &H747474)
    Call APILine(1, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, _
        UserControl.ScaleHeight - 2, &HA1A1A1)
    Call APILine(2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 4, _
        UserControl.ScaleHeight - 1, &HDDDDDD)
    '* English: Lines for the button area.
    Call DrawVGradient(ColorB, Color3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, _
        UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight / 3)
    Call DrawVGradient(Color3, ColorB, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, _
        UserControl.ScaleHeight / 3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, _
        UserControl.ScaleHeight * 2 / 3 - 4)
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 0, _
        UserControl.ScaleWidth - 3, 0, Color4)
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, 1, _
        UserControl.ScaleWidth - 4, 1, Color5)
    '* English: Right shadow.
    Call DrawVGradient(ColorH, ColorI, UserControl.ScaleWidth - 1, 2, UserControl.ScaleWidth, 3)
    Call DrawVGradient(ColorI, ColorA, UserControl.ScaleWidth - 1, 3, UserControl.ScaleWidth, _
        UserControl.ScaleHeight / 2 - 6)
    Call DrawVGradient(ColorA, ColorI, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 2 - 2, _
        UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
    Call DrawVGradient(ColorI, ColorH, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 8, _
        UserControl.ScaleWidth, 3)
    '* English: Layer1.
    Call DrawVGradient(Color4, Color3, UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 1, _
        UserControl.ScaleHeight - 7)
    '* English: Layer2.
    Call DrawVGradient(Color4, ColorC, UserControl.ScaleWidth - 3, 1, UserControl.ScaleWidth - 2, _
        UserControl.ScaleHeight - 6)
    '* English: Doted Area / 1-Bottom.
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4, ColorG)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 4, Color7)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5, ColorF)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 5, Color7)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6, Color9)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 4, Color6)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, Color6)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 2, _
        &HCACACA)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 2, _
        &HBFBFBF)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 1, _
        &HE4E4E4)
    '* English: Doted Area / 2-Botom
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 4, ColorD)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 5, ColorE)
    '* English: Doted Area / 3-Top.
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 0, IIf(m_StateG <> -1, &HA76E4A, _
        ShiftColorOXP(&H808080, 55)))
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 0, Color6)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 1, Color8)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, 1, IIf(m_StateG <> -1, &HB3A49D, _
        GetLngColor(Parent.BackColor)))
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, 1, Color9)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 1, Color8)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 2, Color9)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 3, Color8)
    '* English: Draw Twin Arrows.
    cx = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
    cy = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 - 1
    hPen = CreatePen(0, 1, IIf(m_StateG <> -1, &H0&, ShiftColorOXP(&H0&)))
    hPenOld = SelectObject(UserControl.hDC, hPen)
    '* English: Down Arrow.
    Call MoveToEx(UserControl.hDC, cx - 3, cy + 1, PT)
    Call LineTo(UserControl.hDC, cx + 1, cy + 1)
    Call LineTo(UserControl.hDC, cx, cy + 2)
    Call LineTo(UserControl.hDC, cx - 2, cy + 2)
    Call LineTo(UserControl.hDC, cx - 2, cy + 3)
    Call LineTo(UserControl.hDC, cx, cy + 3)
    Call LineTo(UserControl.hDC, cx - 1, cy + 4)
    Call LineTo(UserControl.hDC, cx - 1, cy + 6)
    '* English: Up Arrow.
    Call MoveToEx(UserControl.hDC, cx - 3, cy - 2, PT)
    Call LineTo(UserControl.hDC, cx + 1, cy - 2)
    Call LineTo(UserControl.hDC, cx, cy - 3)
    Call LineTo(UserControl.hDC, cx - 2, cy - 3)
    Call LineTo(UserControl.hDC, cx - 2, cy - 4)
    Call LineTo(UserControl.hDC, cx, cy - 4)
    Call LineTo(UserControl.hDC, cx - 1, cy - 5)
    Call LineTo(UserControl.hDC, cx - 1, cy - 7)
    '* English: Destroy PEN.
    Call SelectObject(hDC, hPenOld)
    Call DeleteObject(hPen)
    '* English: Undo the offset.
    m_btnRect.Left = m_btnRect.Left + 4

End Sub

Private Sub DrawOfficeButton(ByVal WhatOffice As ComboOfficeAppearance)

  Dim tmpRect As RECT

    '* English: Draw Office Style appearance.
    '* Español: Dibuja la apariencia de Office.
    tmpRect = m_btnRect

    Select Case WhatOffice
    Case 0
        '* English: Style Office Xp, appearance default.
        '* Español: Estilo Office Xp, apariencia por defecto.

        If (m_StateG = 1) Then
            '* English: Normal Color.
            '* Español: Color Normal.
            tmpcolor = NormalBorderColor
        ElseIf (m_StateG = 2) Then
            '* English: Highlight Color.
            '* Español: Color de Selección MouseMove.
            tmpcolor = HighLightBorderColor
            cValor = 185
        ElseIf (m_StateG = 3) Then
            '* English: Down Color.
            '* Español: Color de Selección MouseDown.
            tmpcolor = SelectBorderColor
            tempBorderColor = tmpcolor
            cValor = 125
        Else
            '* English: Disabled Color.
            '* Español: Color deshabilitado.
            tmpcolor = ConvertSystemColor(ShiftColorOXP(NormalBorderColor, 41))
        End If

        If (m_StateG > 1) Then
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), _
                tmpcolor, B
            UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 14, _
                UserControl.ScaleHeight - 2), ShiftColorOXP(tmpcolor, cValor), BF
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), _
                tmpcolor, B
        Else
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), _
                tmpcolor, B
            UserControl.Line (UserControl.ScaleWidth - 3, 2)-(UserControl.ScaleWidth - 13, _
                UserControl.ScaleHeight - 3), tmpcolor, BF
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), _
                tmpcolor, B
        End If

        Call DrawStandardArrow(m_btnRect, ArrowColor)

    Case 1
        '* English: Style Office 2000.
        '* Español: Estilo Office 2000.

        If (m_StateG = 1) Then
            '* English: Flat.
            '* Español: Normal.
            tmpcolor = NormalBorderColor
            Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 13, 1, 12, _
                UserControl.ScaleHeight - 2, ShiftColorOXP(tmpcolor, 175), False)
        ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
            '* English: Mouse Hover or Mouse Pushed.
            '* Español: Mouse presionado o MouseMove.

            If (m_StateG = 2) Then
                tmpcolor = ShiftColorOXP(HighLightBorderColor)
            Else
                tmpcolor = ShiftColorOXP(SelectBorderColor)
                tempBorderColor = tmpcolor
            End If

            Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
                BDR_SUNKENOUTER)
            tmpRect.Left = tmpRect.Left + 4
            Call APIFillRect(UserControl.hDC, tmpRect, tmpcolor)
            tmpRect.Left = tmpRect.Left - 1
            Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, _
                UserControl.ScaleHeight - 3, tmpcolor)
            Call APILine(tmpRect.Left, tmpRect.Top, tmpRect.Left, tmpRect.Bottom, tmpcolor)

            If (m_StateG = 2) Then
                Call DrawCtlEdgeByRect(UserControl.hDC, tmpRect, BDR_RAISEDINNER)
            Else
                Call DrawCtlEdgeByRect(UserControl.hDC, tmpRect, BDR_SUNKENOUTER)
            End If

            m_btnRect.Left = m_btnRect.Left - 3
        Else
            '* English: Disabled control.
            '* Español: Control deshabilitado.
            Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, _
                UserControl.ScaleHeight, ShiftColorOXP(&HC0C0C0, 36))
            Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 18, 1, 17, _
                UserControl.ScaleHeight - 2, myBackColor, False)
        End If

        tmpRect.Left = tmpRect.Left + 4
        Call DrawStandardArrow(tmpRect, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))

    Case 2
        '* English: Style Office 2003.
        '* Español: Estilo Office 2003.

        If (m_StateG <> -1) Then
            tmpC2 = GetSysColor(COLOR_WINDOW)
        Else
            tmpC2 = ShiftColorOXP(GetSysColor(COLOR_BTNFACE))
        End If

        tmpC1 = ArrowColor
        UserControl.BackColor = tmpC2
        tmpcolor = GetSysColor(COLOR_HOTLIGHT)

        If (m_StateG = 1) Then
            cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), _
                GetSysColor(29)), 109)
            iFor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_INACTIVECAPTIONTEXT), _
                GetSysColor(COLOR_GRADIENTINACTIVECAPTION)))
        ElseIf (m_StateG = 2) Then
            cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), _
                GetSysColor(29)), 170)
            iFor = cValor
        ElseIf (m_StateG = 3) Then
            cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), _
                GetSysColor(29)), 140)
            iFor = cValor
        Else
            tmpC1 = GetSysColor(COLOR_GRAYTEXT)
            Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, _
                UserControl.ScaleHeight, tmpC1)
            GoTo DrawNowArrow
        End If

        Call DrawGradient(UserControl.hDC, m_btnRect.Left + 4, tmpRect.Top - 1, tmpRect.Right + 1, _
            tmpRect.Bottom + 1, iFor, cValor, 1)

        If (m_StateG = 2) Or (m_StateG = 3) Then
            Call DrawRectangleBorder(UserControl.hDC, 0, 0, UserControl.ScaleWidth, _
                UserControl.ScaleHeight, tmpcolor)
            Call DrawRectangleBorder(UserControl.hDC, UserControl.ScaleWidth - 15, 0, 17, _
                UserControl.ScaleHeight, tmpcolor, True)
            tempBorderColor = tmpcolor
        End If

DrawNowArrow:
        Call DrawStandardArrow(tmpRect, tmpC1)
        myBackColor = tmpC2
    End Select

End Sub

Private Sub DrawRectangleBorder(ByVal hDC As Long, _
                                ByVal x As Long, _
                                ByVal y As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal Color As Long, _
                                Optional ByVal SetBorder As Boolean = True)

  Dim hBrush As Long
  Dim TempRect As RECT

    '* English: Draw a rectangle.
    '* Español: Crea el rectángulo.
    On Error Resume Next
    TempRect.Left = x
    TempRect.Top = y
    TempRect.Right = x + Width
    TempRect.Bottom = y + Height
    hBrush = CreateSolidBrush(Color)

    If (SetBorder = True) Then
        Call FrameRect(hDC, TempRect, hBrush)
    Else
        Call FillRect(hDC, TempRect, hBrush)
    End If

    Call DeleteObject(hBrush)

End Sub

Private Sub DrawShadow(ByVal iColor1 As Long, _
                       ByVal iColor2 As Long, _
                       Optional ByVal SoftColor As Boolean = True)

    '* English: Set a Shadow Border.
    '* Español: Coloca un borde con sombra.
    tmpC2 = 15

    If (SoftColor = True) Then
        tmpC3 = 178
        iFor = 10
    Else
        tmpC3 = 0
        iFor = 0
    End If

    For tmpC1 = 1 To 16
        tmpC2 = tmpC2 - 1
        '* Horizontal Top Border.
        Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, 1, _
            UserControl.ScaleWidth - tmpC1, 1, ShiftColorOXP(iColor1, tmpC3))
        '* Horizontal Bottom Border.
        Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, _
            UserControl.ScaleHeight - 2, UserControl.ScaleWidth - tmpC1, UserControl.ScaleHeight - _
            2, IIf(m_StateG = -1, ShiftColorOXP(iColor2, tmpC3), ShiftColorOXP(iColor2, iFor)))

        If (SoftColor = True) Then
            tmpC3 = tmpC3 - 5
            iFor = iFor + 5
        End If

    Next
    m_btnRect.Bottom = m_btnRect.Bottom - 11

    If (SoftColor = True) Then
        tmpC3 = 128
        iFor = 70
    End If

    For tmpC1 = 0 To 12
        '* Vertical Left Border.
        Call APILine(m_btnRect.Left + 1, m_btnRect.Top + tmpC1 - 1, m_btnRect.Left + 1, _
            m_btnRect.Bottom + tmpC1 - 1, ShiftColorOXP(iColor1, tmpC3))
        '* Vertical Right Border.
        Call APILine(UserControl.ScaleWidth - 2, m_btnRect.Top + tmpC1 - 1, UserControl.ScaleWidth _
            - 2, m_btnRect.Bottom + tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(iColor2, tmpC3), _
            ShiftColorOXP(iColor2, iFor)))

        If (SoftColor = True) Then
            tmpC3 = tmpC3 + 5
            iFor = iFor - 5
        End If

    Next

End Sub

Private Sub DrawStandardArrow(ByRef RT As RECT, ByVal lColor As Long)

  Dim PT   As POINTAPI
  Dim hPenOld As Long
  Dim cx As Long
  Dim hPen As Long
  Dim cy           As Long

    '* English: Draw the standard arrow in a Rect.
    '* Español: Dibuje la flecha normal en un Rect.

    If (AppearanceCombo = 1) And (OfficeAppearance = 1) Or (AppearanceCombo = 10) Or _
        (AppearanceCombo = 17) Then
        hPen = 1
    ElseIf ((OfficeAppearance = 2) Or (OfficeAppearance = 0)) And (AppearanceCombo = 1) Then
        hPen = 2
    End If

    cx = RT.Left + (RT.Right - RT.Left) - (7 - hPen)
    cy = RT.Top + (RT.Bottom - RT.Top) / 2
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hDC, hPen)
    Call MoveToEx(UserControl.hDC, cx - 3, cy - 1, PT)
    Call LineTo(UserControl.hDC, cx + 1, cy - 1)
    Call LineTo(UserControl.hDC, cx, cy)
    Call LineTo(UserControl.hDC, cx - 2, cy)
    Call LineTo(UserControl.hDC, cx, cy + 2)
    Call SelectObject(hDC, hPenOld)
    Call DeleteObject(hPen)

End Sub

Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           rtRect As RECT) As Boolean

  Dim hTheme  As Long '* hTheme Handle.
  Dim lResult As Long '* Temp Variable.

    '* If a error occurs then or we are not running XP or the visual style is Windows Classic.
    On Error GoTo NoXP
    '* Get out hTheme Handle.
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
    '* Did we get a theme handle?.

    If (hTheme) Then
        '* Yes! Draw the control Background.
        lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
        '* If drawing was successful, return true, or false If not.
        DrawTheme = IIf(lResult, False, True)
    Else
        '* No, we couldn't get a hTheme, drawing failed.
        DrawTheme = False
    End If

    '* Close theme.
    Call CloseThemeData(hTheme)
    '* Exit the function now.
    Exit Function
NoXP:
    '* An Error was detected, drawing Failed.
    DrawTheme = False

End Function

Private Sub DrawVGradient(ByVal lEndColor As Long, _
                          ByVal lStartcolor As Long, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal x2 As Long, _
                          ByVal y2 As Long)

  Dim dR As Single
  Dim dG As Single
  Dim dB As Single
  Dim ni As Long
  Dim sR As Single
  Dim sG As Single
  Dim sB As Single
  Dim eR As Single
  Dim eG As Single
  Dim eB As Single

    '* English: Draw a Vertical Gradient in the current hDC.
    '* Español: Dibuja un degradado en forma vertical.
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / y2
    dG = (sG - eG) / y2
    dB = (sB - eB) / y2

    For ni = 0 To y2
        Call APILine(x, y + ni, x2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
    Next

End Sub

Private Sub DrawWinXPButton(ByVal XpAppearance As ComboXpAppearance, ByVal tmpcolor As OLE_COLOR)

  Dim tmpXPAppearance   As ComboXpAppearance
  Dim isState           As Integer
  Dim bDrawThemeSuccess As Boolean
  Dim tmpRect           As RECT

    '* English: This Sub Draws the XpAppearance Button.
    '* Español: Este procedimiento dibuja el Botón estilo XP.
    isFailedXP = False

    If (XpAppearance = 0) Then
        '* Draw the XP Themed Style.
        isState = IIf(m_StateG < 0, 4, m_StateG)
        Call SetRect(tmpRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
        bDrawThemeSuccess = DrawTheme("Edit", 2, isState, tmpRect)
        Call SetRect(tmpRect, m_btnRect.Left - 1, m_btnRect.Top - 1, m_btnRect.Right + 1, _
            m_btnRect.Bottom + 1)
        bDrawThemeSuccess = DrawTheme("ComboBox", 1, isState, tmpRect)

        If (bDrawThemeSuccess = True) Then
            Exit Sub
        Else '* If themed failed, then use the Next Style.
            tmpXPAppearance = 7 '* If failed, use custom colors.
            isFailedXP = True
            GoTo noUxThemed
            'myArrowColor = vbBlack
            'Call DrawAppearance(Win98, 1)
            Exit Sub
        End If

    Else
        tmpXPAppearance = XpAppearance
    End If

noUxThemed:

    If (tmpXPAppearance = 7) And (m_StateG <> -1) Then
        UserControl.BackColor = BackColor
    ElseIf (m_StateG <> -1) Then
        UserControl.BackColor = GetSysColor(COLOR_WINDOW)
    Else
        UserControl.BackColor = &HE5ECEC
    End If

    Call APIRectangle(UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - _
        1, IIf(m_StateG <> -1, tmpcolor, &HC2C9C9))
    Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - _
        3, GetSysColor(COLOR_WINDOW))

    Select Case tmpXPAppearance
    Case 1
        '* English: Style WinXp Aqua.
        '* Español: Estilo WinXp Aqua.
        cValor = &H85614D
        tempBorderColor = &HC56A31
        tmpC2 = &HB99D7F

        If (m_StateG = 1) Then
            tmpC3 = &HF5C8B3
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 2) Then
            tmpC3 = ShiftColorOXP(&HF5C8B3, 58)
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 3) Then
            tmpC3 = &HF9A477
            tmpcolor = &HFFFFFF
        End If

    Case 2
        '* English: Style WinXp Olive Green.
        '* Español: Estilo WinXp Olive Green.
        cValor = &HFFFFFF
        tempBorderColor = &H668C7D
        tmpC2 = &H94CCBC

        If (m_StateG = 1) Then
            tmpC3 = &H8BB4A4
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 2) Then
            tmpC3 = &HA7D7CA
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 3) Then
            tmpC3 = &H80AA98
            tmpcolor = &HFFFFFF
        End If

    Case 3
        '* English: Style WinXp Silver.
        '* Español: Estilo WinXp Silver.
        tempBorderColor = &HA29594
        cValor = &H48483E
        tmpC2 = &HA29594

        If (m_StateG = 1) Then
            tmpC3 = &HDACCCB
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 2) Then
            tmpC3 = ShiftColorOXP(&HDACCCB, 58)
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 3) Then
            tmpC3 = &HE5D1CF
            tmpcolor = &HFFFFFF
        End If

    Case 4
        '* English: Style WinXp TasBlue.
        '* Español: Estilo WinXp TasBlue.
        tempBorderColor = &HF09F5F
        cValor = ShiftColorOXP(&H703F00, 58)
        tmpC2 = &HF09F5F

        If (m_StateG = 1) Then
            tmpC3 = &HF0AF70
            tmpcolor = &HFFE7CF
        ElseIf (m_StateG = 2) Then
            tmpC3 = ShiftColorOXP(&HF0BF80, 58)
            tmpcolor = &HFFEFD0
        ElseIf (m_StateG = 3) Then
            tmpC3 = &HF09F5F
            tmpcolor = &HFFEFD0
        End If

    Case 5
        '* English: Style WinXp Gold.
        '* Español: Estilo WinXp Gold.
        tempBorderColor = &HBFE7F0
        cValor = ShiftColorOXP(&H6F5820, 45)
        tmpC2 = &HBFE7F0

        If (m_StateG = 1) Then
            tmpC3 = ShiftColorOXP(&HCFFFFF, 54)
            tmpcolor = &HBFF0FF
        ElseIf (m_StateG = 2) Then
            tmpC3 = &HBFEFFF
            tmpcolor = ShiftColorOXP(&HCFFFFF, 58)
        ElseIf (m_StateG = 3) Then
            tmpC3 = &HCFFFFF
            tmpcolor = &HBFE8FF
        End If

    Case 6
        '* English: Style WinXp Blue.
        '* Español: Estilo WinXp Blue.
        tempBorderColor = ShiftColorOXP(&HA0672F, 123)
        cValor = &H6F5820
        tmpC2 = ShiftColorOXP(&HA0672F, 123)

        If (m_StateG = 1) Then
            tmpC3 = &HEFF0F0
            tmpcolor = &HF0F7F0
        ElseIf (m_StateG = 2) Then
            tmpC3 = &HF0F8FF
            tmpcolor = &HF0F7F0
        ElseIf (m_StateG = 3) Then
            tmpC3 = &HF1946E
            tmpcolor = &HEEC2B4
        End If

    Case 7
        '* English: Style WinXp Custom.
        '* Español: Estilo WinXp Custom.
        tempBorderColor = SelectBorderColor
        cValor = ArrowColor

        If (m_StateG = 1) Then
            tmpC3 = NormalBorderColor
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 2) Then
            tmpC3 = HighLightBorderColor
            tmpcolor = &HFFFFFF
        ElseIf (m_StateG = 3) Then
            tmpC3 = SelectBorderColor
            tmpcolor = &HFFFFFF
        End If

        tmpC2 = tmpC3
    End Select

    If (m_StateG = -1) Then
        tmpcolor = &HE5ECEC
        tmpC3 = m_btnRect.Bottom - m_btnRect.Top
        tmpC1 = m_btnRect.Bottom - 1

        For iFor = 3 To tmpC1
            Call APILine(m_btnRect.Left + 1, tmpC3 - iFor + 3, m_btnRect.Right - 1, tmpC3 - iFor + _
                3, tmpcolor)
        Next

        tmpC1 = ShiftColorOXP(&HC2C9C9, 19)
    Else
        tmpC1 = tmpC2
        Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, _
            m_btnRect.Bottom, tmpcolor, tmpC3, 1)
    End If

    Call APIRectangle(hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, _
        m_btnRect.Bottom - m_btnRect.Top - 1, tmpC1)
    Call DrawXpArrow(IIf(m_StateG = -1, &HC2C9C9, cValor))

End Sub

Private Sub DrawXpArrow(Optional ByVal iColor3 As OLE_COLOR = &H0)

    '* English: Draw The XP Style Arrow.
    '* Español: Dibuja la flecha estilo Xp.
    tmpC1 = m_btnRect.Right - m_btnRect.Left
    tmpC2 = m_btnRect.Bottom - m_btnRect.Top + 1
    tmpC1 = m_btnRect.Left + tmpC1 / 2 + 1
    tmpC2 = m_btnRect.Top + tmpC2 / 2
    If (iColor3 = &H0) Then iColor3 = ArrowColor
    Call APILine(tmpC1 - 5, tmpC2 - 2, tmpC1, tmpC2 + 3, iColor3)
    Call APILine(tmpC1 - 4, tmpC2 - 2, tmpC1, tmpC2 + 2, iColor3)
    Call APILine(tmpC1 - 4, tmpC2 - 3, tmpC1, tmpC2 + 1, iColor3)
    Call APILine(tmpC1 + 3, tmpC2 - 2, tmpC1 - 2, tmpC2 + 3, iColor3)
    Call APILine(tmpC1 + 2, tmpC2 - 2, tmpC1 - 2, tmpC2 + 2, iColor3)
    Call APILine(tmpC1 + 2, tmpC2 - 3, tmpC1 - 2, tmpC2 + 1, iColor3)

End Sub

Public Property Get Enabled() As Boolean

    '* English: Sets/Gets the Enabled property of the control.
    '* Español: Devuelve o establece si el Usercontrol esta habilitado ó deshabilitado.
    Enabled = ControlEnabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled = New_Enabled
    ControlEnabled = New_Enabled
    tmrFocus.Enabled = False
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("Enabled")
    Refresh

End Property

Public Property Get Font() As StdFont

    '* English: Sets/Gets the Font of the control.
    '* Español: Devuelve o establece el tipo de fuente del texto.
    Set Font = g_Font

End Property

Public Property Set Font(ByVal New_Font As StdFont)

    On Error Resume Next

    With g_Font
        .Name = New_Font.Name
        .Size = IIf(New_Font.Size > 12, 8, New_Font.Size)
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With

    Call isEnabled(ControlEnabled)
    Call PropertyChanged("Font")
    Refresh

End Property

Public Function GetControlVersion() As String

    '* English: Control Version.
    '* Español: Version del Control.
    GetControlVersion = Version & " © " & Year(Now)

End Function

Private Function GetLngColor(ByVal Color As Long) As Long

    '* English: The GetSysColor function retrieves the current color of the specified display
    '   element. Display elements are the parts of a window and the Windows display that appear on
    '   the
    '   system display screen.
    '* Español: Recupera el color actual del elemento de despliegue especificado.

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If

End Function

Public Property Let GradientColor1(ByVal New_Color As OLE_COLOR)

    myGradientColor1 = ConvertSystemColor(New_Color)
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("GradientColor1")
    Refresh

End Property

Public Property Get GradientColor1() As OLE_COLOR

    '* English: Sets/Gets the color First gradient color.
    '* Español: Devuelve o establece el color Gradient 1.
    GradientColor1 = myGradientColor1

End Property

Public Property Get GradientColor2() As OLE_COLOR

    '* English: Sets/Gets the Second gradient color.
    '* Español: Devuelve o establece el color Gradient 2.
    GradientColor2 = myGradientColor2

End Property

Public Property Let GradientColor2(ByVal New_Color As OLE_COLOR)

    myGradientColor2 = ConvertSystemColor(New_Color)
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("GradientColor2")
    Refresh

End Property

Public Property Get HighLightBorderColor() As OLE_COLOR

    '* English: Sets/Gets the color of the border of the control when the the control is
    '   highlighted.
    '* Español: Devuelve o establece el color del borde del control cuando el pasa sobre él.
    HighLightBorderColor = myHighLightBorderColor

End Property

Public Property Let HighLightBorderColor(ByVal New_Color As OLE_COLOR)

    myHighLightBorderColor = ConvertSystemColor(New_Color)
    Call PropertyChanged("HighLightBorderColor")
    Refresh

End Property

Public Property Get HighLightColorText() As OLE_COLOR

    '* English: Sets/Gets the color of the selection of the text.
    '* Español: Devuelve o establece el color de selección del texto.
    HighLightColorText = myHighLightColorText

End Property

Public Property Let HighLightColorText(ByVal New_Color As OLE_COLOR)

    myHighLightColorText = ConvertSystemColor(New_Color)
    Call PropertyChanged("HighLightColorText")
    Refresh

End Property

Public Property Get hWnd() As Long

    '* English: Returns a handle to a form or control.
    '* Español: Devuelve el controlador de un formulario o un control.
    hWnd = UserControl.hWnd

End Property

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean

  Dim mPos As POINTAPI
  Dim oRect As RECT

    '* English: Verifies if the mouse is on the object or if one makes clic outside of him.
    '* Español: Verifica si el mouse se encuentra sobre el objeto ó si se hace clic fuera de él.
    Call GetCursorPos(mPos)
    Call GetWindowRect(ObjecthWnd, oRect)
    UserControl.MousePointer = myMousePointer

    If (mPos.x >= oRect.Left) And (mPos.x <= oRect.Right) And (mPos.y >= oRect.Top) And (mPos.y <= _
        oRect.Bottom) Then
        InFocusControl = True
    End If

End Function

Private Sub isEnabled(ByVal isTrue As Boolean)

    '* English: Shows the state of Enabled or Disabled of the Control.
    '* Español: Muestra el estado de Habilitado ó Deshabilitado del Control.

    If (isTrue = True) Then
        Call DrawAppearance(myAppearanceCombo, 1)
    Else
        Call DrawAppearance(myAppearanceCombo, -1)
    End If

End Sub

Private Sub LongToRGB(ByVal lColor As Long)

    '* English: Convert a Long to RGB format.
    '* Español: Convierte un Long en formato RGB.
    RGBColor.Red = lColor And &HFF
    RGBColor.Green = (lColor \ &H100) And &HFF
    RGBColor.Blue = (lColor \ &H10000) And &HFF

End Sub

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)

    Set myMouseIcon = New_MouseIcon

End Property

Public Property Get MouseIcon() As StdPicture

    '* English: Sets a custom mouse icon.
    '* Español: Establece un icono escogido por el usuario.
    Set MouseIcon = myMouseIcon

End Property

Public Property Get MousePointer() As MousePointerConstants

    '* English: Sets/Gets the type of mouse pointer displayed when over part of an object.
    '* Español: Devuelve o establece el tipo de puntero a mostrar cuando el mouse pase sobre el
    '   objeto.
    MousePointer = myMousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

    myMousePointer = New_MousePointer

End Property

Public Property Get NormalBorderColor() As OLE_COLOR

    '* English: Sets/Gets the normal border color of the control.
    '* Español: Devuelve o establece el color normal del borde del control.
    NormalBorderColor = myNormalBorderColor

End Property

Public Property Let NormalBorderColor(ByVal New_Color As OLE_COLOR)

    myNormalBorderColor = ConvertSystemColor(New_Color)
    If (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
    Call PropertyChanged("NormalBorderColor")
    Refresh

End Property

Public Property Get NormalColorText() As OLE_COLOR

    '* English: Sets/Gets the normal text color in the control.
    '* Español: Devuelve o establece el color del texto normal.
    NormalColorText = myNormalColorText

End Property

Public Property Let NormalColorText(ByVal New_Color As OLE_COLOR)

    myNormalColorText = ConvertSystemColor(New_Color)
    If (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
    Call PropertyChanged("NormalColorText")
    Refresh

End Property

Public Property Get OfficeAppearance() As ComboOfficeAppearance

    '* English: Sets/Gets the office apperance.
    '* Español: Devuelve o establece la apariencia de Office.
    OfficeAppearance = myOfficeAppearance

End Property

Public Property Let OfficeAppearance(ByVal New_Apperance As ComboOfficeAppearance)

    myOfficeAppearance = New_Apperance
    If (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
    Call PropertyChanged("OfficeAppearance")
    Refresh

End Property

Public Sub Refresh()
    
    Call isEnabled(True)
    
End Sub

Public Property Let SelectBorderColor(ByVal New_Color As OLE_COLOR)

    mySelectBorderColor = ConvertSystemColor(New_Color)
    Call PropertyChanged("SelectBorderColor")
    Refresh

End Property

Public Property Get SelectBorderColor() As OLE_COLOR

    '* English: Sets/Gets the color of the border of the control when It has the focus.
    '* Español: Devuelve o establece el color del borde del control cuando el tenga el enfoque.
    SelectBorderColor = mySelectBorderColor

End Property

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long

  Dim cRed   As Long
  Dim cBlue   As Long
  Dim Delta  As Long
  Dim cGreen As Long

    cBlue = ((theColor \ &H10000) Mod &H100)
    cGreen = ((theColor \ &H100) Mod &H100)
    cRed = (theColor And &HFF)
    Delta = &HFF - Base
    cBlue = Base + cBlue * Delta \ &HFF
    cGreen = Base + cGreen * Delta \ &HFF
    cRed = Base + cRed * Delta \ &HFF
    If (cRed > 255) Then cRed = 255
    If (cGreen > 255) Then cGreen = 255
    If (cBlue > 255) Then cBlue = 255
    ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue

End Function

Public Property Get Text() As String

    '* English: Sets/Gets the text of the selected item.
    '* Español: Devuelve o establece el texto del item seleccionado.
    Text = myText

End Property

Public Property Let Text(ByVal NewText As String)

    myText = NewText
    UserText = myText
    If (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
    Call PropertyChanged("Text")

End Property

Private Sub tmrFocus_Timer()

    If (InFocusControl(UserControl.hWnd) = False) Then
        tmrFocus.Enabled = False
        IsOver = False
        Call DrawAppearance(myAppearanceCombo, 1)
        RaiseEvent MouseLeave
    End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    If (AppearanceCombo = 18) Then
         If (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
    End If

End Sub

Private Sub UserControl_ExitFocus()

    NoDown = False
    IsOver = False
    Call UserControl_LostFocus

End Sub

Private Sub UserControl_Initialize()

  Dim OS As OSVERSIONINFO

    '* Get the operating system version for text drawing purposes.
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_InitProperties()

    '* English: Setup properties values.
    '* Español: Establece propiedades iniciales.
    ControlEnabled = True
    isPicture = False
    myAppearanceCombo = defAppearanceCombo
    myArrowColor = defArrowColor
    myBackColor = defListColor
    myDisabledColor = defDisabledColor
    myGradientColor1 = defGradientColor1
    myGradientColor2 = defGradientColor2
    myHighLightBorderColor = defHighLightBorderColor
    myHighLightColorText = defHighLightColorText
    myNormalBorderColor = defNormalBorderColor
    myNormalColorText = defNormalColorText
    myOfficeAppearance = defOfficeAppearance
    mySelectBorderColor = defSelectBorderColor
    myText = Ambient.DisplayName
    Text = myText
    myXpAppearance = 1
    Set g_Font = Ambient.Font

End Sub

Private Sub UserControl_LostFocus()

    If (NoDown = False) Then
        If (Ambient.UserMode = True) Then Call isEnabled(ControlEnabled)
    End If
    NoShow = False

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button = vbLeftButton) And (NoDown = False) Then
        NoDown = True
        Call DrawAppearance(myAppearanceCombo, 3)
        RaiseEvent Click
    ElseIf (Button = vbLeftButton) And (NoDown = True) Then
        NoDown = False
        Call DrawAppearance(myAppearanceCombo, 3)
        Call Wait(0.2)
        Call DrawAppearance(myAppearanceCombo, 2)
        RaiseEvent CloseList
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button < 2) Then
        If (InFocusControl(UserControl.hWnd) = False) Then
            Call DrawAppearance(myAppearanceCombo, 1)
        ElseIf (Button = 0) And Not (IsOver = True) Then
            tmrFocus.Enabled = True
            IsOver = True
            Call DrawAppearance(myAppearanceCombo, 2)
            RaiseEvent MouseEnter
        End If

    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Call DrawAppearance(myAppearanceCombo, 2)
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next
    AppearanceCombo = PropBag.ReadProperty("AppearanceCombo", defAppearanceCombo)
    ArrowColor = PropBag.ReadProperty("ArrowColor", defArrowColor)
    BackColor = PropBag.ReadProperty("BackColor", defListColor)
    DisabledColor = PropBag.ReadProperty("DisabledColor", defDisabledColor)
    Enabled = PropBag.ReadProperty("Enabled", True)
    GradientColor1 = PropBag.ReadProperty("GradientColor1", defGradientColor1)
    GradientColor2 = PropBag.ReadProperty("GradientColor2", defGradientColor2)
    Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
    HighLightBorderColor = PropBag.ReadProperty("HighLightBorderColor", defHighLightBorderColor)
    HighLightColorText = PropBag.ReadProperty("HighLightColorText", defHighLightColorText)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    NormalBorderColor = PropBag.ReadProperty("NormalBorderColor", defNormalBorderColor)
    NormalColorText = PropBag.ReadProperty("NormalColorText", defNormalColorText)
    OfficeAppearance = PropBag.ReadProperty("OfficeAppearance", defOfficeAppearance)
    SelectBorderColor = PropBag.ReadProperty("SelectBorderColor", defSelectBorderColor)
    Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
    XpAppearance = PropBag.ReadProperty("XpAppearance", 1)
    On Error GoTo 0

End Sub

Private Sub UserControl_Resize()

    If (Ambient.UserMode = False) Then Call isEnabled(ControlEnabled)

End Sub

Private Sub UserControl_Show()

    tmrFocus.Enabled = False
    Call isEnabled(ControlEnabled)

End Sub

Private Sub UserControl_Terminate()

    tmrFocus.Enabled = False

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AppearanceCombo", myAppearanceCombo, defAppearanceCombo)
    Call PropBag.WriteProperty("ArrowColor", myArrowColor, defArrowColor)
    Call PropBag.WriteProperty("BackColor", myBackColor, defListColor)
    Call PropBag.WriteProperty("DisabledColor", myDisabledColor, defDisabledColor)
    Call PropBag.WriteProperty("Enabled", ControlEnabled, True)
    Call PropBag.WriteProperty("Font", g_Font, Ambient.Font)
    Call PropBag.WriteProperty("GradientColor1", myGradientColor1, defGradientColor1)
    Call PropBag.WriteProperty("GradientColor2", myGradientColor2, defGradientColor2)
    Call PropBag.WriteProperty("HighLightBorderColor", myHighLightBorderColor, _
        defHighLightBorderColor)
    Call PropBag.WriteProperty("HighLightColorText", myHighLightColorText, defHighLightColorText)
    Call PropBag.WriteProperty("MouseIcon", myMouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", myMousePointer, 0)
    Call PropBag.WriteProperty("NormalBorderColor", myNormalBorderColor, defNormalBorderColor)
    Call PropBag.WriteProperty("NormalColorText", myNormalColorText, defNormalColorText)
    Call PropBag.WriteProperty("OfficeAppearance", myOfficeAppearance, defOfficeAppearance)
    Call PropBag.WriteProperty("SelectBorderColor", mySelectBorderColor, defSelectBorderColor)
    Call PropBag.WriteProperty("Text", myText, Ambient.DisplayName)
    Call PropBag.WriteProperty("XpAppearance", myXpAppearance, 1)

End Sub

Private Sub Wait(ByVal Segundos As Single)

  Dim ComienzoSeg As Single
  Dim FinSeg As Single

    '* English: Wait a certain time.
    '* Español: Esperar un determinado tiempo.
    ComienzoSeg = Timer
    FinSeg = ComienzoSeg + Segundos

    Do While FinSeg > Timer
        DoEvents
        If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
    Loop

End Sub

Public Property Get XpAppearance() As ComboXpAppearance

    '* English: Sets the appearance in Xp Mode.
    '* Español: Establece la apariencia en modo Xp.
    XpAppearance = myXpAppearance

End Property

Public Property Let XpAppearance(ByVal new_Style As ComboXpAppearance)

    myXpAppearance = new_Style
    Call isEnabled(ControlEnabled)
    Call PropertyChanged("XpAppearance")

End Property

