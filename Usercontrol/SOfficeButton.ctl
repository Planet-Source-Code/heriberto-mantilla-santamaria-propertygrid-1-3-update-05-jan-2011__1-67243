VERSION 5.00
Begin VB.UserControl SOfficeButton 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   PropertyPages   =   "SOfficeButton.ctx":0000
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "SOfficeButton.ctx":0035
End
Attribute VB_Name = "SOfficeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************'
'*        All Rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
'*                   Version 1.0.3                     *'
'*******************************************************'
'* Control:       SOfficeButton                        *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Description:   This usercontrol simulates a Office  *'
'*                Button.                              *'
'*                                                     *'
'*                This button is based on the origi-   *'
'*                nal code of fred.cpp, please see     *'
'*                the [CodeId = 56053].                *'
'*                                                     *'
'*                Also many thanks to Paul Caton for   *'
'*                it's spectacular self-subclassing    *'
'*                usercontrol template, please see     *'
'*                the [CodeId = 54117].                *'
'*******************************************************'
'* Started on:    Sunday, 09-jan-2005.                 *'
'*******************************************************'
'* Release date:  Monday, 18-jul-2005.                 *'
'*******************************************************'
'*                                                     *'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*        Please rate my work on this control.         *'
'*    That lives the Soccer and the América of Cali    *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All Rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
Option Explicit

'* Private Types.

Private Type RECT
    xLeft    As Long
    xTop     As Long
    xRight   As Long
    xBottom  As Long
End Type

'*******************************************************'
'*                Subclasser Declarations              *'
'*                                                     *'
'* Author: Paul Caton.                                 *'
'* Mail:   Paul_Caton@hotmail.com                      *'
'* Web:    None                                        *'
'*******************************************************'

'-uSelfSub declarations---------------------------------------------------------------------------

Private Enum eMsgWhen                                                       'When to callback
    MSG_BEFORE = 1                                                            'Callback before the
    '   original WndProc
    MSG_AFTER = 2                                                             'Callback after the
    '   original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and
    '   after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table
'   entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the
'   WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc
'   index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of
'   the shutdown flag
Private Const IDX_hWnd      As Long = 2                                     'Thunk data index of
'   the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of
'   the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of
'   the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of
'   the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of
'   the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(654)           As Long                                         'Thunk machine-code
'   initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address
'   collection

Private Declare Function CallWindowProcA Lib "user32" ( _
        ByVal lpPrevWndFunc As Long, _
        ByVal hWnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
        ByVal hModule As Long, _
        ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
        ByVal hWnd As Long, _
        lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long, _
        ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
        ByVal Destination As Long, _
        ByVal Source As Long, _
        ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Event MouseEnter()
Public Event MouseLeave()

Private Const WM_MOUSEMOVE         As Long = &H200
Private Const WM_MOUSELEAVE        As Long = &H2A3
Private Const WM_THEMECHANGED      As Long = &H31A
Private Const WM_SYSCOLORCHANGE    As Long = &H15 '21

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                      As Long
    dwFlags                     As TRACKMOUSEEVENT_FLAGS
    hWndTrack                   As Long
    dwHoverTime                 As Long
End Type

Private bTrack                As Boolean
Private bTrackUser32          As Boolean
Private isInCtrl              As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" ( _
        lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" _
        Alias "_TrackMouseEvent" ( _
        lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'*******************************************************'

'*******************************************************'
'*                     Tool Tip Class                  *'
'*                                                     *'
'* Author: Mark Mokoski                                *'
'* Mail: markm@cmtelephone.com                         *'
'* Web:  www.rjillc.com                                *'
'*******************************************************'

'******************************************************
'* API Functions.                                     *
'******************************************************
Private Declare Function CreateWindowEx Lib "user32" _
        Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'******************************************************
'* Constants.                                         *
'******************************************************

'* Windows API Constants.
Private Const CW_USEDEFAULT = &H80000000
Private Const hWnd_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const WM_USER = &H400

'* Tooltip Window Constants.
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_BALLOON = &H40
Private Const TTS_NOPREFIX = &H2

'* Tool Tip Icons.
Private Const TTI_ERROR                   As Long = 3
Private Const TTI_INFO                    As Long = 1
Private Const TTI_NONE                    As Long = 0
Private Const TTI_WARNING                 As Long = 2

'* Tool Tip API Class.
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'******************************************************
'* Types.                                             *
'******************************************************

'* Tooltip Window Types.

Private Type TOOLINFO
    lSize                             As Long
    lFlags                            As Long
    lhWnd                             As Long
    lId                               As Long
    lpRect                            As RECT
    hInstance                         As Long
    lpStr                             As String
    lParam                            As Long
End Type

'******************************************************
'* Local Class variables and Data .                   *
'******************************************************

'* Local variables to hold property values.
Private ToolActive                        As Boolean
Private ToolBackColor                     As Long
Private ToolCentered                      As Boolean
Private ToolForeColor                     As Long
Private ToolIcon                          As ToolIconType
Private TOOLSTYLE                         As ToolStyleEnum
Private ToolText                          As String
Private ToolTitle                         As String

'* Private Data for Class.
Private m_ltthWnd                         As Long
Private TI                                As TOOLINFO

Public Enum ToolIconType
    TipNoIcon = TTI_NONE            '= 0
    TipIconInfo = TTI_INFO          '= 1
    TipIconWarning = TTI_WARNING    '= 2
    TipIconError = TTI_ERROR        '= 3
End Enum

Public Enum ToolStyleEnum
    StyleStandard = 0
    StyleBalloon = 1
End Enum

'*******************************************************'

'* Private Types.

Private Type POINTAPI
    x      As Long
    y      As Long
End Type

'* Private Enum's.

Public Enum OfficeAlign
    ACenter = &H0
    ALeft = &H1
    ARight = &H2
    ATop = &H3
    ABottom = &H4
End Enum

Public Enum OfficeState
    OfficeNormal = &H0
    OfficeHighLight = &H1
    OfficeHot = &H2
    OfficeDisabled = &H3
End Enum

Public Enum ShapeBorder
    Rectangle = &H0
    [Round Rectangle] = &H1
End Enum

'* Private variables.
Private g_Font           As StdFont
Private isAutoSizePic    As Boolean
Private isBackColor      As OLE_COLOR
Private isBorderColor    As OLE_COLOR
Private isButtonShape    As ShapeBorder
Private isCaption        As String
Private isDisabledColor  As OLE_COLOR
Private isEnabled        As Boolean
Private isFocus          As Boolean
Private isFontAlign      As OfficeAlign
Private isForeColor      As OLE_COLOR
Private isHeight         As Long
Private isHighLightColor As OLE_COLOR
Private isHotColor       As OLE_COLOR
Private isHotTitle       As Boolean
Private isMultiLine      As Boolean
Private isPicture        As StdPicture
Private isPictureAlign   As OfficeAlign
Private isPictureSize    As Integer
Private isSetBorder      As Boolean
Private isSetBorderH     As Boolean
Private isSetGradient    As Boolean
Private isSetHighLight   As Boolean
Private isShadowText     As Boolean
Private isShowFocus      As Boolean
Private isState          As OfficeState
Private isSystemColor    As Boolean
Private isTxtRect        As RECT
Private isWidth          As Long
Private isXPos           As Integer
Private isYPos           As Integer
Private m_bGrayIcon      As Boolean
Private RectButton       As RECT

'* Private Constants.
Private Const defBackColor      As Long = vbButtonFace
Private Const defBorderColor    As Long = vbHighlight
Private Const defDisabledColor  As Long = vbGrayText
Private Const defForeColor      As Long = vbButtonText
Private Const defHighLightColor As Long = vbHighlight
Private Const defHotColor       As Long = vbHighlight
Private Const defShape          As Integer = &H0
Private Const DSS_DISABLED      As Long = &H20
Private Const DSS_MONO          As Long = &H80
Private Const DSS_NORMAL        As Long = &H0
Private Const DST_BITMAP        As Long = &H4
Private Const DST_ICON          As Long = &H3
Private Const DT_BOTTOM         As Long = &H8
Private Const DT_CENTER         As Long = &H1
Private Const DT_LEFT           As Long = &H0
Private Const DT_RIGHT          As Long = &H2
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_TOP            As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_WORDBREAK      As Long = &H10
Private Const DT_WORD_ELLIPSIS  As Long = &H40000
Private Const PS_SOLID          As Long = 0
Private Const SW_SHOWNORMAL     As Long = 1
Private Const Version           As String = "SOfficeButon 1.0.3 By HACKPRO TM"

'* API's Windows Call.
Private Declare Function CreatePen Lib "gdi32" ( _
        ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DrawState Lib "user32" _
        Alias "DrawStateA" ( _
        ByVal hDC As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal Flags As Long) As Long
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
Private Declare Function FrameRect Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" ( _
        lpRect As RECT, _
        ByVal x As Long, _
        ByVal y As Long) As Long
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
Private Declare Function RoundRect Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long, _
        ByVal X3 As Long, _
        ByVal Y3 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" ( _
        ByVal xPoint As Long, _
        ByVal yPoint As Long) As Long

'* Public Events.
Public Event Click()
Public Event ChangedTheme()

'* For Create GrayIcon --> MArio Florez.
Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" ( _
        ByVal hDC As Long, _
        ByVal xLeft As Long, _
        ByVal yTop As Long, _
        ByVal hIcon As Long, _
        ByVal cxWidth As Long, _
        ByVal cyWidth As Long, _
        ByVal istepIfAniCur As Long, _
        ByVal hbrFlickerFreeDraw As Long, _
        ByVal diFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" ( _
        ByVal hIcon As Long, _
        ByRef piconinfo As ICONINFO) As Long
Private Declare Function GetPixel Lib "gdi32.dll" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32.dll" _
        Alias "GetObjectA" ( _
        ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' Type - GetObjectAPI.lpObject

Private Type BITMAP
    bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
    bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width
    '   must be greater than zero.
    bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height
    '   must be greater than zero.
    bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value
    '   must be divisible by 2, because Windows assumes that the bit values of a bitmap form an
    '   array
    '   that is word aligned.
    bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
    bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color
    '   of a pixel.
    bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The
    '   bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

' Type - CreateIconIndirect / GetIconInfo

Private Type ICONINFO
    fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A
    '   value of TRUE specifies an icon; FALSE specifies a cursor.
    xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor’s hot spot. If this
    '   structure defines an icon, the hot spot is always in the center of the icon, and this member
    '   is ignored.
    yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor’s hot spot. If this
    '   structure defines an icon, the hot spot is always in the center of the icon, and this member
    '   is ignored.
    hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a
    '   black and white icon, this bitmask is formatted so that the upper half is the icon AND
    '   bitmask
    '   and the lower half is the icon XOR bitmask. Under this condition, the height should be an
    '   even
    '   multiple of two. If this structure defines a color icon, this mask only defines the AND
    '   bitmask of the icon.
    hbmColor As Long 'HBITMAP // Identifies the icon color bitmap. This member can be optional if
    '   this structure defines a black and white icon. The AND bitmask of hbmMask is applied with
    '   the
    '   SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the
    '   destination by using the SRCINVERT flag.
End Type

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

Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" ( _
        lpVersionInformation As OSVERSIONINFO) As Long

Private mWindowsNT   As Boolean

'*******************************************************'
'* Public Properties.                                  *'
'*******************************************************'

'*******************************************************'
'* Private Subs and Functions.                         *'
'*******************************************************'
'* English: Paints lines in a simple and faster.


Private Sub APILine(ByVal whDC As Long, _
                    ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal X2 As Long, _
                    ByVal Y2 As Long, _
                    ByVal lColor As Long)

  Dim PT As POINTAPI
  Dim hPen As Long
  Dim hPenOld As Long

    '* Pinta líneas de forma sencilla y rápida.

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(whDC, hPen)
    Call MoveToEx(whDC, X1, Y1, PT)
    Call LineTo(whDC, X2, Y2)
    Call SelectObject(whDC, hPenOld)
    Call DeleteObject(hPen)

End Sub

Public Property Let AutoSizePicture(ByVal TheAutoSize As Boolean)

'* English: Adjusts the control to the picture size.

    '* Ajusta el control al tamaño de la imagen.

    isAutoSizePic = TheAutoSize
    Call PropertyChanged("AutoSizePicture")
    Call Refresh(isState)

End Property

Public Property Get AutoSizePicture() As Boolean

    AutoSizePicture = isAutoSizePic

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = isBackColor

End Property

Public Property Let BackColor(ByVal theColor As OLE_COLOR)

'* English: Returns/Sets the background color used to display text and graphics in an object.

    '* Devuelve ó establece el color del Usercontrol.

    isBackColor = ConvertSystemColor(theColor)
    Call PropertyChanged("BackColor")
    Call Refresh(isState)

End Property

Public Property Get BorderColor() As OLE_COLOR

    BorderColor = isBorderColor

End Property

Public Property Let BorderColor(ByVal theColor As OLE_COLOR)

'* English: Returns/Sets the color of border of the Object.

    '* Devuelve ó establece el color del borde del objeto.

    isBorderColor = ConvertSystemColor(theColor)
    Call PropertyChanged("BorderColor")
    If (isSetBorder = True) Then Call Refresh(isState)

End Property

Public Property Let ButtonShape(ByVal theButtonShape As ShapeBorder)

'* English: Returns/Sets the type of border of the control.

    '* Devuelve ó establece el tipo de borde del botón.

    isButtonShape = theButtonShape
    Call PropertyChanged("ButtonShape")
    If (isSetBorder = True) Then Call Refresh(isState)

End Property

Public Property Get ButtonShape() As ShapeBorder

    ButtonShape = isButtonShape

End Property

Public Property Let Caption(ByVal TheCaption As String)

'* English: Returns/Sets "Caption" property.

    '* Devuelve ó establece el texto del Objeto.

    isCaption = TheCaption
    Call SetAccessKey(isCaption)
    Call PropertyChanged("Caption")
    Call Refresh(isState)

End Property

Public Property Get Caption() As String

    Caption = isCaption

End Property

Public Property Let CaptionAlign(ByVal theAlign As OfficeAlign)

'* English: Returns/Sets alignment of the text.

    '* Devuelve ó establece la alineación del texto.

    isFontAlign = theAlign
    Call PropertyChanged("CaptionAlign")
    Call Refresh(isState)

End Property

Public Property Get CaptionAlign() As OfficeAlign

    CaptionAlign = isFontAlign

End Property

Private Function ConvertSystemColor(ByVal theColor As Long) As Long

'* English: Convert Long to System Color.

    '* Convierte un long en un color del sistema.

    Call OleTranslateColor(theColor, 0, ConvertSystemColor)

End Function

Private Function CreateIconFromBMP(ByVal hBMP_Mask As Long, ByVal hBMP_Image As Long) As Long

  Dim TempICONINFO As ICONINFO

    If (hBMP_Mask = 0) Or (hBMP_Image = 0) Then Exit Function
    TempICONINFO.fIcon = 1
    TempICONINFO.hbmMask = hBMP_Mask
    TempICONINFO.hbmColor = hBMP_Image
    CreateIconFromBMP = CreateIconIndirect(TempICONINFO)

End Function

Private Sub CreateToolTip()

'* Private sub used with Create and Update subs/functions.

  Dim lpRect As RECT
  Dim lWinStyle As Long

    '* If Tool Tip already made, destroy it and reconstruct.

    Call TipRemove
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    '* Create Baloon TipStyle if desired.
    If (TOOLSTYLE = StyleBalloon) Then lWinStyle = lWinStyle Or TTS_BALLOON
    '* The parent control has to be set first.

    If (UserControl.hWnd <> &H0) Then
        m_ltthWnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, _
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
        Call SendMessage(m_ltthWnd, TTM_ACTIVATE, CInt(ToolActive), TI)
        '* Make our ToolTip window a topmost window.
        Call SetWindowPos(m_ltthWnd, hWnd_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or _
            SWP_NOMOVE)
        '* Get the rectangle of the parent control.
        Call GetClientRect(UserControl.hWnd, lpRect)
        '* Now set up our ToolTip info structure.

        With TI
            '* If we want it TipCentered, then set that flag.

            If (ToolCentered = True) Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
             Else
                .lFlags = TTF_SUBCLASS
            End If

            '* Set the hWnd prop to our Parent Control's hWnd.
            .lhWnd = UserControl.hWnd
            .lId = 0
            .hInstance = App.hInstance
            .lpRect = lpRect
            .lpStr = ToolText
        End With

        '* Add the ToolTip Structure.
        Call SendMessage(m_ltthWnd, TTM_ADDTOOLA, 0&, TI)
        '* Set Max Width to 32 characters, and enable Multi Line Tool Tips.
        Call SendMessage(m_ltthWnd, TTM_SETMAXTIPWIDTH, 0&, &H20)

        If (ToolIcon <> TipNoIcon) Or (ToolTitle <> vbNullString) Then
            '* If we want a TipTitle or we want an TipIcon.
            Call SendMessage(m_ltthWnd, TTM_SETTITLE, CLng(ToolIcon), ByVal ToolTitle)
        End If

        If (ToolForeColor <> Empty) Then
            '* 0 (zero) or Null is seen by the API as the default color. See ForeColor property for
            '   more datails.
            Call SendMessage(m_ltthWnd, TTM_SETTIPTEXTCOLOR, ToolForeColor, 0&)
        End If

        If (ToolBackColor <> Empty) Then
            '* 0 (zero) or Null is seen by the API as the default color. See BackColor property for
            '   more datails.
            Call SendMessage(m_ltthWnd, TTM_SETTIPBKCOLOR, ToolBackColor, 0&)
        End If

    End If

End Sub

Public Property Let DisabledColor(ByVal theColor As OLE_COLOR)

'* English: Returns/Sets the color of the disabled text.

    '* Devuelve ó establece el color del texto deshabilitado.

    isDisabledColor = ConvertSystemColor(theColor)
    Call PropertyChanged("DisabledColor")
    Call Refresh(isState)

End Property

Public Property Get DisabledColor() As OLE_COLOR

    DisabledColor = isDisabledColor

End Property

Private Sub DrawBox(ByVal hDC As Long, _
                    ByVal Offset As Long, _
                    ByVal Radius As Long, _
                    ByVal ColorFill As Long, _
                    ByVal ColorBorder As Long, _
                    ByVal isWidth As Long, _
                    ByVal isHeight As Long)

'* English: Paints a rectangle with oval border.

  Dim pRect As RECT
  Dim hPen As Long
  Dim hBrush As Long

    '* Crea un rectángulo con border ovalados.

    On Error Resume Next
    pRect.xLeft = -4
    pRect.xRight = isWidth - IIf(isCaption = "", 1, -2)
    pRect.xTop = -3
    pRect.xBottom = isHeight - 1
    hPen = SelectObject(hDC, CreatePen(PS_SOLID, 1, ColorBorder))
    hBrush = SelectObject(hDC, CreateSolidBrush(ColorFill))
    Call InflateRect(pRect, -Offset, -Offset)
    Call RoundRect(hDC, pRect.xLeft, pRect.xTop, pRect.xRight, pRect.xBottom, Radius, Radius)
    Call InflateRect(pRect, Offset, Offset)
    Call DeleteObject(SelectObject(hDC, hPen))
    Call DeleteObject(SelectObject(hDC, hBrush))
    On Error GoTo 0

End Sub

Private Sub DrawCaption(ByVal iColor1 As Long, ByVal iColor2 As Long)

'* English: Draw the text on the Object.

  Dim lColor As Long
  Dim isFAlign As Long

    '* Dibuja el texto sobre el Objeto.

    If (isMultiLine = True) Then lColor = DT_WORDBREAK Else lColor = DT_SINGLELINE

    Select Case isFontAlign
     Case ACenter
        isFAlign = DT_CENTER Or DT_VCENTER Or lColor Or DT_WORD_ELLIPSIS

     Case ALeft
        isFAlign = DT_VCENTER Or DT_LEFT Or lColor Or DT_WORD_ELLIPSIS

     Case ARight
        isFAlign = DT_VCENTER Or DT_RIGHT Or lColor Or DT_WORD_ELLIPSIS

     Case ATop
        isFAlign = DT_CENTER Or DT_TOP Or lColor Or DT_WORD_ELLIPSIS

     Case ABottom
        isFAlign = DT_CENTER Or DT_BOTTOM Or lColor Or DT_WORD_ELLIPSIS
    End Select

    If (isState <> OfficeDisabled) Then
        lColor = iColor2
     Else
        lColor = iColor1
    End If

    If (isShadowText = True) And ((isState = &H1) Or (isState = &H2)) Then
        isTxtRect.xLeft = isTxtRect.xLeft + 1.5
        isTxtRect.xTop = isTxtRect.xTop + 1.5
        Call SetTextColor(UserControl.hDC, ShiftColorOXP(lColor))

        If (mWindowsNT = True) Then
            Call DrawTextW(UserControl.hDC, StrPtr(isCaption), Len(isCaption), isTxtRect, isFAlign)
         Else
            Call DrawTextA(UserControl.hDC, isCaption, Len(isCaption), isTxtRect, isFAlign)
        End If

        isTxtRect.xLeft = isTxtRect.xLeft - 1.5
        isTxtRect.xTop = isTxtRect.xTop - 1.5
    End If

    Call SetTextColor(UserControl.hDC, lColor)
    '*************************************************************************
    '* Draws the text with Unicode support based on OS version.              *
    '* Thanks to Richard Mewett.                                             *
    '*************************************************************************

    If (mWindowsNT = True) Then
        Call DrawTextW(UserControl.hDC, StrPtr(isCaption), Len(isCaption), isTxtRect, isFAlign)
     Else
        Call DrawTextA(UserControl.hDC, isCaption, Len(isCaption), isTxtRect, isFAlign)
    End If

End Sub

Private Sub DrawFocus()

'* English: Show focus of control.

  Dim iPos As Integer

    '* Muestra el enfoque del control.

    If (isFocus = True) And (isShowFocus = True) Then
        If (isButtonShape = &H0) Then '* Shape Rectangle.
            Call DrawFocusRect(UserControl.hDC, RectButton)
         Else

            For iPos = RectButton.xLeft + 3 To RectButton.xRight - IIf(isCaption = "", 7, 4)
                Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + 1, &H1DD6B7)
                Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + isHeight - 3, &H1DD6B7)
            Next iPos

            For iPos = RectButton.xTop + 4 To RectButton.xTop + isHeight - 5
                Call SetPixel(UserControl.hDC, RectButton.xLeft, iPos, &H1DD6B7)
                Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 4, 1), iPos, _
                    &H1DD6B7)
            Next iPos

            For iPos = RectButton.xLeft + 3 To RectButton.xRight - IIf(isCaption = "", 7, 4) Step 2
                Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + 1, &H24427A)
                Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + isHeight - 3, &H24427A)
            Next iPos

            For iPos = RectButton.xTop + 4 To RectButton.xTop + isHeight - 5 Step 2
                Call SetPixel(UserControl.hDC, RectButton.xLeft, iPos, &H24427A)
                Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 4, 1), iPos, _
                    &H24427A)
            Next iPos

            Call SetPixel(UserControl.hDC, RectButton.xLeft + 1, 2, vbBlack)
            Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 5, 2), 2, _
                vbBlack)
            Call SetPixel(UserControl.hDC, RectButton.xLeft + 1, RectButton.xTop + isHeight - 4, _
                vbBlack)
            Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 5, 2), _
                RectButton.xTop + isHeight - 4, vbBlack)
        End If

    End If

End Sub

Private Sub DrawPicture()

'* English: Draw a picture in the Object.

  Dim isType As Long
  Dim isValue As Long

    '* Crea la imagen sobre el Objeto.

    On Error Resume Next

    If Not (isPicture Is Nothing) Then
        If (Picture <> 0) Then
            Dim Ix As Long
            Dim Iy As Long

            If (isPictureSize <= 0) Then isPictureSize = 16

            Select Case isPicture.Type
             Case 1, 4: isType = DST_BITMAP
             Case 3:    isType = DST_ICON
            End Select

            If (isPictureAlign = &H0) Then
                Ix = (isWidth - isPictureSize) / 2
                Iy = (isHeight - isPictureSize) / 2
             ElseIf (isPictureAlign = &H1) Then
                Ix = isXPos
                Iy = (isHeight - isPictureSize) / 2
             ElseIf (isPictureAlign = &H2) Then
                Ix = isWidth - isPictureSize - isXPos
                Iy = (isHeight - isPictureSize) / 2
             ElseIf (isPictureAlign = &H3) Then
                Ix = (isWidth - isPictureSize) / 2
                Iy = isYPos
             ElseIf (isPictureAlign = &H4) Then
                Ix = (isWidth - isPictureSize) / 2
                Iy = isHeight - isPictureSize - isYPos
            End If

        End If

        If (isEnabled = False) Then
            isValue = DSS_DISABLED

            If (m_bGrayIcon = False) Then
                Call DrawState(UserControl.hDC, 0, 0, isPicture.handle, 0, Ix, Iy, isPictureSize, _
                    isPictureSize, isType Or isValue)
             Else
                Call RenderIconGrayscale(UserControl.hDC, isPicture.handle, Ix, Iy, isPictureSize, _
                    isPictureSize)
            End If

         Else
            isValue = DSS_NORMAL

            If (isState = OfficeHot) Then
                Ix = Ix - 1
                Iy = Iy - 1
             ElseIf (isState = OfficeHighLight) Then
                isValue = CreateSolidBrush(RGB(136, 141, 157))
                Call DrawState(UserControl.hDC, isValue, 0, isPicture.handle, 0, Ix, Iy, _
                    isPictureSize, isPictureSize, isType Or DSS_MONO)
                Ix = Ix - 2
                Iy = Iy - 2
                isValue = DSS_NORMAL
                Call DrawState(UserControl.hDC, 0, 0, isPicture.handle, 0, Ix, Iy, isPictureSize, _
                    isPictureSize, isType Or isValue)
                Call DeleteObject(isValue)
                Exit Sub
            End If

            Call RenderIconGrayscale(UserControl.hDC, isPicture.handle, Ix, Iy, isPictureSize, _
                isPictureSize, False)
        End If

    End If

End Sub

Private Sub DrawRectangle(ByVal hDC As Long, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal Width As Long, _
                          ByVal Height As Long, _
                          ByVal ColorFill As Long, _
                          ByVal ColorBorder As Long, _
                          Optional ByVal SetBackGround As Boolean = True)

'* English: Draw a rectangle area with a specific color.

  Dim hBrush As Long
  Dim TempRect As RECT

    '* Crea un área rectangular con un color específico.

    TempRect.xLeft = x
    TempRect.xTop = y
    TempRect.xRight = x + Width
    TempRect.xBottom = y + Height
    hBrush = CreateSolidBrush(ColorBorder)
    Call FrameRect(hDC, TempRect, hBrush)
    Call DeleteObject(hBrush)

    If (SetBackGround = True) Then
        TempRect.xLeft = x + 1
        TempRect.xTop = y + 1
        TempRect.xRight = x + Width - 1
        TempRect.xBottom = y + Height - 1
        hBrush = CreateSolidBrush(ColorFill)
        Call FillRect(hDC, TempRect, hBrush)
        Call DeleteObject(hBrush)
    End If

End Sub

Private Sub DrawVGradient(ByVal whDC As Long, _
                          ByVal lEndColor As Long, _
                          ByVal lStartcolor As Long, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal X2 As Long, _
                          ByVal Y2 As Long)

'* English: Draws a degraded one in vertical form.

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

    '* Dibuja un degradado en forma vertical.

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2

    For ni = 0 To Y2
        Call APILine(whDC, x, y + ni, X2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * _
            dB)))
    Next ni

End Sub

Public Property Let Enabled(ByVal TheEnabled As Boolean)

'* English: Returns/Sets the Enabled property of the control.

    '* Devuelve ó establece si el Usercontrol esta habilitado ó deshabilitado.

    isEnabled = TheEnabled
    UserControl.Enabled = isEnabled
    Call PropertyChanged("Enabled")

    If (isEnabled = True) Then
        isState = OfficeNormal
     Else
        isState = OfficeDisabled
    End If

    Call Refresh(isState)

End Property

Public Property Get Enabled() As Boolean

    Enabled = isEnabled

End Property

Public Property Get Font() As StdFont

    Set Font = g_Font

End Property

Public Property Set Font(ByVal New_Font As StdFont)

'* English: Returns/Sets the Font of the control.

    '* Devuelve ó establece el tipo de fuente del texto.

    On Error Resume Next

    With g_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With

    Call PropertyChanged("Font")
    Call Refresh(isState)

End Property

Public Property Let ForeColor(ByVal theColor As OLE_COLOR)

'* English: Use this color for drawing Normal Font.

    '* Devuelve ó establece el color de la fuente.

    isForeColor = ConvertSystemColor(theColor)
    Call PropertyChanged("ForeColor")
    Call Refresh(isState)

End Property

Public Property Get ForeColor() As OLE_COLOR

    ForeColor = isForeColor

End Property

Public Property Get GetControlVersion() As String

'* English: Control Version.

    '* Español: Version del Control.

    GetControlVersion = Version & " © " & Year(Now)

End Property

Public Function GetIconBitmaps(ByVal hIcon As Long, _
                               ByRef Return_hBmpMask As Long, _
                               ByRef Return_hBmpImage As Long) As Boolean

  Dim TempICONINFO As ICONINFO

    If (GetIconInfo(hIcon, TempICONINFO) = 0) Then Exit Function
    Return_hBmpMask = TempICONINFO.hbmMask
    Return_hBmpImage = TempICONINFO.hbmColor
    GetIconBitmaps = True

End Function

Public Property Let GrayIcon(ByVal bGrayIcon As Boolean)

    m_bGrayIcon = bGrayIcon
    Call PropertyChanged("GrayIcon")
    Call Refresh

End Property

Public Property Get GrayIcon() As Boolean

    GrayIcon = m_bGrayIcon

End Property

Public Property Let HighLightColor(ByVal theColor As OLE_COLOR)

'* English: Use this color for drawing.

    '* Color de fondo cuando el mouse pasa sobre el Objeto.

    isHighLightColor = ConvertSystemColor(theColor)
    Call PropertyChanged("HighLightColor")
    Call Refresh(isState)

End Property

Public Property Get HighLightColor() As OLE_COLOR

    HighLightColor = isHighLightColor

End Property

Public Property Get HotColor() As OLE_COLOR

    HotColor = isHotColor

End Property

Public Property Let HotColor(ByVal theColor As OLE_COLOR)

'* English: Use this color for drawing.

    '* Color de fondo cuando se tiene presionado el Objeto.

    isHotColor = ConvertSystemColor(theColor)
    Call PropertyChanged("HotColor")
    Call Refresh(isState)

End Property

Public Property Let HotTitle(ByVal theTitle As Boolean)

'* English: Use this color for drawing.

    '* Color de fondo cuando se tiene presionado el Objeto.

    isHotTitle = theTitle
    Call PropertyChanged("HotTitle")

End Property

Public Property Get HotTitle() As Boolean

    HotTitle = isHotTitle

End Property

Public Property Get hWnd() As Long

'* English: Returns a handle to the control.

    '* Devuelve el controlador del control.

    hWnd = UserControl.hWnd

End Property

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
'Determine if the passed function is supported

  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)

        If hMod Then
            bLibLoaded = True
        End If

    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If

    End If

    If bLibLoaded Then
        FreeLibrary hMod
    End If

End Function

Private Function IsMouseOver() As Boolean

'* English: Return, if the mouse is over the Object.

  Dim PT As POINTAPI

    '* Devuelve si el mouse esta sobre el objeto.

    Call GetCursorPos(PT)
    IsMouseOver = (WindowFromPoint(PT.x, PT.y) = hWnd)

End Function

Public Property Get MouseIcon() As StdPicture

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal MouseIcon As StdPicture)

'* English: Sets a custom mouse icon.

    '* Devuelve ó establece un icono de mouse personalizado.

    Set UserControl.MouseIcon = MouseIcon
    Call PropertyChanged("MouseIcon")

End Property

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal MousePointer As MousePointerConstants)

'* English: Returns/Sets the type of mouse pointer displayed when over part of an object.

    '* Devuelve ó establece el tipo de puntero a mostrar cuando el mouse pase sobre el objeto.

    UserControl.MousePointer = MousePointer
    Call PropertyChanged("MousePointer")

End Property

Public Property Let MultiLine(ByVal theMultiLine As Boolean)

'* English: Returns/Sets if the text is shown in multiple lines.

    '* Devuelve ó establece si el texto se muestra en múltiples líneas.

    isMultiLine = theMultiLine
    Call PropertyChanged("MultiLine")
    Call Refresh(isState)

End Property

Public Property Get MultiLine() As Boolean

    MultiLine = isMultiLine

End Property

Public Function OpenLink(ByVal sLink As String) As Long

'* English: Executable file or a document file.

    '* Ejecuta un archivo ó documento cualquiera.

    On Error Resume Next
    OpenLink = ShellExecute(Parent.hWnd, vbNullString, sLink, vbNullString, "C:\", SW_SHOWNORMAL)
    On Error GoTo 0

End Function

Public Property Get Picture() As StdPicture

    Set Picture = isPicture

End Property

Public Property Set Picture(ByVal thePicture As StdPicture)

'* English: Returns/Sets the image of the control.

    '* Devuelve ó establece la imagen del control.

    Set isPicture = thePicture
    Call PropertyChanged("Picture")
    Call Refresh(isState)

End Property

Public Property Let PictureAlign(ByVal theAlign As OfficeAlign)

'* English: Returns/Sets the alignment of the image.

    '* Devuelve ó establece la alineación de la imagen.

    isPictureAlign = theAlign
    Call PropertyChanged("PictureAlign")
    Call Refresh(isState)

End Property

Public Property Get PictureAlign() As OfficeAlign

    PictureAlign = isPictureAlign

End Property

Public Property Let PictureSize(ByVal theSize As Integer)

'* English: Returns/Sets the picture size.

    '* Devuelve ó establece el tamaño de la imagen.

    isPictureSize = theSize
    Call PropertyChanged("PictureSize")
    Call Refresh(isState)

End Property

Public Property Get PictureSize() As Integer

    PictureSize = isPictureSize

End Property

Public Sub Refresh(Optional ByVal State As OfficeState = 0)

'* English: Draw appearance of the control.

  Dim lColor  As Long
  Dim lBase    As Long
  Dim iColor1 As Long
  Dim iColor2 As Long
  Dim iColor3 As Long
  Dim iColor4 As Long
  Dim iColor5 As Long
  Dim iColor6 As Long
  Dim lBase1   As Integer

    '* Crea la apariencia del control.

    If (isEnabled = False) Then State = OfficeDisabled

    If (isSystemColor = False) Then
        iColor1 = isBackColor
        iColor2 = isBorderColor
        iColor3 = isDisabledColor
        iColor4 = isForeColor
        iColor5 = isHighLightColor
        iColor6 = isHotColor
     Else
        iColor1 = ConvertSystemColor(defBackColor)
        iColor2 = ConvertSystemColor(defBorderColor)
        iColor3 = ConvertSystemColor(defDisabledColor)
        iColor4 = ConvertSystemColor(defForeColor)
        iColor5 = ConvertSystemColor(defHighLightColor)
        iColor6 = ConvertSystemColor(defHotColor)
    End If

    If (isEnabled = False) Then iColor2 = iColor3

    With UserControl
        isHeight = .ScaleHeight
        isWidth = .ScaleWidth
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .Cls
        On Error Resume Next
        Set .Font = g_Font
        Call GetClientRect(.hWnd, RectButton)
        Call GetClientRect(.hWnd, isTxtRect)
        .BackColor = iColor1
        lBase = &HB0
        lBase1 = 1
        If Not (isButtonShape = &H0) Then lBase1 = 4
        'If (State > &H0) And (State < &H3) And (isSetGradient = True) Then State = &H0

        Select Case State
         Case &H0 '* Normal State.
            If (isSetGradient = True) Then Call DrawVGradient(.hDC, iColor1, ShiftColorOXP(iColor1, _
                &H72), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)

            If (isSetBorder = True) Then
                If (isButtonShape = &H0) Then
                    Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, iColor1, iColor2, _
                        IIf(isSetGradient = True, False, True))
                 Else
                    Call DrawBox(.hDC, 4, 5, iColor1, iColor2, RectButton.xRight + 2, _
                        RectButton.xBottom + 3)
                End If

             ElseIf (isSetGradient = False) Then
                Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, iColor1, iColor1)
            End If

         Case &H1, &H2 '* HighLight or Hot State.

            If (isSetHighLight = True) Then
                If (State = &H1) Then
                    lColor = ShiftColorOXP(iColor5, &H40)
                    If (isSetGradient = True) Then Call DrawVGradient(.hDC, iColor1, _
                        ShiftColorOXP(iColor5, &H122), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - _
                        lBase1)
                 Else
                    lColor = ShiftColorOXP(iColor6, &H10)
                    lBase = &H9C
                    If (isSetGradient = True) Then Call DrawVGradient(.hDC, iColor1, _
                        ShiftColorOXP(iColor6, &H40), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - _
                        lBase1)
                End If

             ElseIf (isSetBorderH = True) Then
                lColor = iColor1
                lBase = 0
            End If

            If (isSetBorderH = True) And (isButtonShape = &H0) Then
                Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, ShiftColorOXP(lColor, lBase), _
                    iColor2, IIf(isSetGradient = True, False, True))
             ElseIf (isSetBorderH = True) Then
                Call DrawBox(.hDC, 4, 5, ShiftColorOXP(lColor, lBase), iColor2, RectButton.xRight + _
                    2, RectButton.xBottom + 3)
            End If

         Case &H3 '* Disabled State.
            lColor = iColor3

            If (isSetBorder = True) Then
                If (isButtonShape = &H0) Then
                    Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, iColor1, lColor)
                 Else
                    Call DrawBox(.hDC, 4, 5, iColor1, lColor, RectButton.xRight + 2, _
                        RectButton.xBottom + 3)
                End If

            End If
        End Select
        isState = State

        If (isAutoSizePic = True) Then
            .Width = isPicture.Width
            .Height = isPicture.Height
            isHeight = .ScaleHeight
            isWidth = .ScaleWidth
        End If

        Call DrawCaption(iColor3, iColor4)
        Call DrawPicture
        If (isState <> &H3) Then Call DrawFocus
    End With

End Sub

Private Function RenderBitmapGrayscale(ByVal Dest_hDC As Long, _
                                       ByVal hBitmap As Long, _
                                       Optional ByVal Dest_X As Long, _
                                       Optional ByVal Dest_Y As Long, _
                                       Optional ByVal Srce_X As Long, _
                                       Optional ByVal Srce_Y As Long, _
                                       Optional ByVal GrayC As Boolean = True) As Boolean

'=============================================================================================================

  Dim TempBitmap As BITMAP
  Dim hScreen    As Long
  Dim hDC_Temp   As Long
  Dim hBMP_Prev    As Long
  Dim MyCounterX As Long
  Dim MyCounterY   As Long
  Dim NewColor   As Long
  Dim hNewPicture As Long
  Dim DeletePic  As Boolean

    ' Make sure parameters passed are valid

    If (Dest_hDC = 0) Or (hBitmap = 0) Then Exit Function
    ' Get the handle to the screen DC
    hScreen = GetDC(0)
    If (hScreen = 0) Then Exit Function
    ' Create a memory DC to work with the picture
    hDC_Temp = CreateCompatibleDC(hScreen)
    If (hDC_Temp = 0) Then GoTo CleanUp
    ' If the user specifies NOT to alter the original, then make a copy of it to use
    DeletePic = False
    hNewPicture = hBitmap
    ' Select the bitmap into the DC
    hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
    ' Get the height / width of the bitmap in pixels
    If (GetObjectAPI(hNewPicture, Len(TempBitmap), TempBitmap) = 0) Then GoTo CleanUp
    If (TempBitmap.bmHeight <= 0) Or (TempBitmap.bmWidth <= 0) Then GoTo CleanUp
    ' Loop through each pixel and conver it to it's grayscale equivelant

    If (GrayC = True) Then

        For MyCounterX = 0 To TempBitmap.bmWidth - 1
            For MyCounterY = 0 To TempBitmap.bmHeight - 1
                NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)

                If (NewColor <> -1) Then

                    Select Case NewColor
                        ' If the color is already a grey shade, no need to convert it
                     Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, _
                         &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, _
                         &HE0E0E0, &HF0F0F0
                        NewColor = NewColor

                     Case Else
                        NewColor = 0.33 * (NewColor Mod 256) + 0.59 * ((NewColor \ 256) Mod 256) + _
                            0.11 * ((NewColor \ 65536) Mod 256)
                        NewColor = RGB(NewColor, NewColor, NewColor)
                    End Select

                    Call SetPixel(hDC_Temp, MyCounterX, MyCounterY, NewColor)
                End If

            Next MyCounterY
        Next MyCounterX
    End If

    ' Display the picture on the specified hDC
    Call BitBlt(Dest_hDC, Dest_X, Dest_Y, TempBitmap.bmWidth, TempBitmap.bmHeight, hDC_Temp, Srce_X, _
        Srce_Y, vbSrcCopy)
    RenderBitmapGrayscale = True
CleanUp:
    Call ReleaseDC(0, hScreen): hScreen = 0
    Call SelectObject(hDC_Temp, hBMP_Prev)
    Call DeleteDC(hDC_Temp): hDC_Temp = 0

    If (DeletePic = True) Then
        Call DeleteObject(hNewPicture)
        hNewPicture = 0
    End If

End Function

Private Function RenderIconGrayscale(ByVal Dest_hDC As Long, _
                                     ByVal hIcon As Long, _
                                     Optional ByVal Dest_X As Long, _
                                     Optional ByVal Dest_Y As Long, _
                                     Optional ByVal Dest_Height As Long, _
                                     Optional ByVal Dest_Width As Long, _
                                     Optional ByVal GrayC As Boolean = True) As Boolean

' See post: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58622&lngWId=1
' Thanks MArio Florez.

  Dim hBMP_Mask As Long
  Dim hBMP_Image As Long
  Dim hBMP_Prev As Long
  Dim hIcon_Temp As Long
  Dim hDC_Temp  As Long

    ' Make sure parameters passed are valid

    If (Dest_hDC = 0) Or (hIcon = 0) Then Exit Function
    ' Extract the bitmaps from the icon
    If (GetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False) Then Exit Function
    ' Create a memory DC to work with
    hDC_Temp = CreateCompatibleDC(0)
    If (hDC_Temp = 0) Then GoTo CleanUp
    ' Make the image bitmap gradient
    If (RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0, , , GrayC) = False) Then GoTo CleanUp
    ' Extract the gradient bitmap out of the DC
    Call SelectObject(hDC_Temp, hBMP_Prev)
    ' Take the newly gradient bitmap and make a gradient icon from it
    hIcon_Temp = CreateIconFromBMP(hBMP_Mask, hBMP_Image)
    If (hIcon_Temp = 0) Then GoTo CleanUp
    ' Draw the newly created gradient icon onto the specified DC

    If (DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0) _
        Then
        RenderIconGrayscale = True
    End If

CleanUp:
    Call DestroyIcon(hIcon_Temp): hIcon_Temp = 0
    Call DeleteDC(hDC_Temp): hDC_Temp = 0
    Call DeleteObject(hBMP_Mask): hBMP_Mask = 0
    Call DeleteObject(hBMP_Image): hBMP_Image = 0

End Function

Private Sub sc_AddMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

'Add the message value to the window handle's specified callback table

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the
        '   thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is
            '   to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                              'Add the message
            '   to the before table
        End If

        If When And MSG_AFTER Then                                              'If message is to
            '   be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                              'Add the message
            '   to the after table
        End If

    End If

End Sub

Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long

'Call the original WndProc

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the
        '   thunk hasn't already released its memory
        sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) _
            'Call the original WndProc of the passed window handle parameter
    End If

End Function

Private Sub sc_DelMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

'Delete the message value from the window handle's specified callback table

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the
        '   thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                             'If the message is
            '   to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                              'Delete the
            '   message from the before table
        End If

        If When And MSG_AFTER Then                                              'If the message is
            '   to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                              'Delete the
            '   message from the after table
        End If

    End If

End Sub

Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)

'Let the subclasser lParamUser callback parameter

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the
        '   thunk hasn't already released its memory
        zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser
        '   callback parameter
    End If

End Property

Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long

'Get the subclasser lParamUser callback parameter

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the
        '   thunk hasn't already released its memory
        sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser
        '   callback parameter
    End If

End Property

Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                             Optional ByVal lParamUser As Long = 0, _
                             Optional ByVal nOrdinal As Long = 1, _
                             Optional ByVal oCallback As Object = Nothing, _
                             Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle

'-SelfSub code------------------------------------------------------------------------------------

'*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 =
    '   second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks
    '   are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only
    '   disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************

  Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
  Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate

    '   per thunk, data + code + msg tables
  Const PAGE_RWX      As Long = &H40&                                         'Allocate executable

    '   memory
  Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated

    '   memory
  Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated

    '   memory flag
  Const IDX_EBMODE    As Long = 3                                             'Thunk data index of

    '   the EbMode function address
  Const IDX_CWP       As Long = 4                                             'Thunk data index of

    '   the CallWindowProc function address
  Const IDX_SWL       As Long = 5                                             'Thunk data index of

    '   the SetWindowsLong function address
  Const IDX_FREE      As Long = 6                                             'Thunk data index of

    '   the VirtualFree function address
  Const IDX_BADPTR    As Long = 7                                             'Thunk data index of

    '   the IsBadCodePtr function address
  Const IDX_OWNER     As Long = 8                                             'Thunk data index of

    '   the Owner object's vTable address
  Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of

    '   the callback method address
  Const IDX_EBX       As Long = 16                                            'Thunk code patch

    '   index of the thunk data
  Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long

    If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window
        '   handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If

    nMyID = GetCurrentProcessId                                               'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID
    '   associated with the window handle
    If nID <> nMyID Then                                                      'Ensure that the
        '   window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If

    If oCallback Is Nothing Then                                              'If the user hasn't
        '   specified the callback owner
        Set oCallback = Me                                                      'Then it is me
    End If

    nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of
    '   the specified ordinal method
    If nAddr = 0 Then                                                         'Ensure that we've
        '   found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If

    If z_Funk Is Nothing Then                                                 'If this is the first
        '   time through, do the one-time initialization
        Set z_Funk = New Collection                                             'Create the
        '   hWnd/thunk-address collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: _
            z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: _
            z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: _
            z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = _
            &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) _
            = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: _
            z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: _
            z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = _
            &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: _
            z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = _
            &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: _
            z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = _
            &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store
        '   CallWindowProc function address in the thunk data
        z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the
        '   SetWindowLong function address in the thunk data
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the
        '   VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the
        '   IsBadCodePtr function address in the thunk data
    End If

    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable
    '   memory

    If z_ScMem <> 0 Then                                                      'Ensure the
        '   allocation succeeded
        On Error GoTo CatchDoubleSub                                            'Catch double
        '   subclassing
        z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the
        '   hWnd/thunk-address to the collection
        On Error GoTo 0

        If bIdeSafety Then                                                      'If the user wants
            '   IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode
            '   function address in the thunk data
        End If

        z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk
        '   data address
        z_Sc(IDX_hWnd) = lng_hWnd                                               'Store the window
        '   handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address
        '   of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address
        '   of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback
        '   owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback
        '   address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the
        '   lParamUser callback parameter in the thunk data

        nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new
        '   WndProc, return the address of the original WndProc
        If nAddr = 0 Then                                                       'Ensure the new
            '   WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If

        z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original
        '   WndProc address in the thunk data
        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk
        '   code/data to the allocated memory
        sc_Subclass = True                                                      'Indicate success
     Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If

    Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has
    '   failed after memory allocation, so release the memory

End Function

Private Sub sc_Terminate()

'Terminate all subclassing

  Dim i As Long

    If Not (z_Funk Is Nothing) Then                                           'Ensure that
        '   subclassing has been started

        With z_Funk

            For i = .Count To 1 Step -1                                           'Loop through the
                '   collection of window handles in reverse order
                z_ScMem = .Item(i)                                                  'Get the thunk
                '   address
                If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that
                    '   the thunk hasn't already released its memory
                    sc_UnSubclass zData(IDX_hWnd)                                     'UnSubclass
                End If

            Next i                                                                'Next member of
            '   the collection
        End With

        Set z_Funk = Nothing                                                    'Destroy the
        '   hWnd/thunk-address collection
    End If

End Sub

Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)

'UnSubclass the specified window handle

    If z_Funk Is Nothing Then                                                 'Ensure that
        '   subclassing has been started
        zError "sc_UnSubclass", "Window handle isn't subclassed"
     Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the
            '   thunk hasn't already released its memory
            zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown
            '   indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all
            '   before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after
            '   messages
        End If

        z_Funk.Remove "h" & lng_hWnd                                            'Remove the
        '   specified window handle from the collection
    End If

End Sub

Private Sub SetAccessKey(ByVal Caption As String)

'* English: Returns or sets a string that contains the keys that will act as the access keys (or hot keys for the control.)

  Dim AmperSandPos As Long
  Dim isText As String

    '* Devuelve ó establece una cadena que contiene las teclas que funcionarán como teclas de
    '   acceso (o teclas aceleradoras) del control.

    With UserControl
        .AccessKeys = ""

        If (Len(Caption) > 1) Then
            AmperSandPos = InStr(1, Caption, "&", vbTextCompare)

            If (AmperSandPos < Len(Caption)) And (AmperSandPos > 0) Then
                isText = Mid$(Caption, AmperSandPos + 1, 1)

                If (isText <> "&") Then
                    .AccessKeys = LCase$(isText)
                 Else
                    AmperSandPos = InStr(AmperSandPos + 2, Caption, "&", vbTextCompare)
                    isText = Mid$(Caption, AmperSandPos + 1, 1)
                    If (isText <> "&") Then .AccessKeys = LCase$(isText)
                End If

            End If
        End If
    End With

End Sub

Public Property Get SetBorder() As Boolean

    SetBorder = isSetBorder

End Property

Public Property Let SetBorder(ByVal theSetBorder As Boolean)

'* English: Returns/Sets if it's always shown the border.

    '* Devuelve ó establece si se muestra siempre un borde.

    isSetBorder = theSetBorder
    Call PropertyChanged("SetBorder")
    Call Refresh(isState)

End Property

Public Property Let SetBorderH(ByVal theSetBorderH As Boolean)

'* English: Returns/Sets if it's always shown the Hot border.

    '* Devuelve ó establece si se muestra siempre un borde.

    isSetBorderH = theSetBorderH
    Call PropertyChanged("SetBorderH")

End Property

Public Property Get SetBorderH() As Boolean

    SetBorderH = isSetBorderH

End Property

Public Property Get SetGradient() As Boolean

    SetGradient = isSetGradient

End Property

Public Property Let SetGradient(ByVal theSetGradient As Boolean)

'* English: Returns/Sets if the background is gradient.

    '* Devuelve ó establece si el fondo es en degradado.

    isSetGradient = theSetGradient
    Call PropertyChanged("SetGradient")
    Call Refresh(isState)

End Property

Public Property Let SetHighLight(ByVal theSetHighLight As Boolean)

'* English: Returns/Sets if the background change is shown.

    '* Devuelve ó establece si se muestra el cambio de fondo.

    isSetHighLight = theSetHighLight
    Call PropertyChanged("SetHighLight")

End Property

Public Property Get SetHighLight() As Boolean

    SetHighLight = isSetHighLight

End Property

Public Property Get ShadowText() As Boolean

    ShadowText = isShadowText

End Property

Public Property Let ShadowText(ByVal theShadowText As Boolean)

'* English: Returns/Sets if a shadow is shown in the text of the button.

    '* Devuelve ó establece si se muestra una sombra en el texto del botón.

    isShadowText = theShadowText
    Call PropertyChanged("ShadowText")

End Property

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long

'* English: Shift a color.

  Dim Red   As Long
  Dim Blue   As Long
  Dim Delta As Long
  Dim Green As Long

    '* Devuelve un Color con menos intensidad.

    Blue = ((theColor \ &H10000) Mod &H100)
    Green = ((theColor \ &H100) Mod &H100)
    Red = (theColor And &HFF)
    Delta = &HFF - Base
    Blue = Base + Blue * Delta \ &HFF
    Green = Base + Green * Delta \ &HFF
    Red = Base + Red * Delta \ &HFF
    If (Red > 255) Then Red = 255
    If (Green > 255) Then Green = 255
    If (Blue > 255) Then Blue = 255
    ShiftColorOXP = Red + 256& * Green + 65536 * Blue

End Function

Public Property Let ShowFocus(ByVal theFocus As Boolean)

'* English: Do you want to show the focus?

    '* Permite ver el enfoque del control.

    isShowFocus = theFocus
    Call PropertyChanged("ShowFocus")

End Property

Public Property Get ShowFocus() As Boolean

    ShowFocus = isShowFocus

End Property

Public Property Let SystemColor(ByVal theSystemColor As Boolean)

'* English: Take the system color.

    '* Toma los colores del Sistema.

    isSystemColor = theSystemColor
    Call PropertyChanged("SystemColor")
    Call Refresh(isState)

End Property

Public Property Get SystemColor() As Boolean

    SystemColor = isSystemColor

End Property

Public Property Let TipActive(ByVal ToolData As Boolean)

    '* If True, activate (show) ToolTip, False deactivate (hide) tool tip.
    '* Syntax: object.TipActive = True/False.

    ToolActive = ToolData
    Call PropertyChanged("TipActive")

End Property

Public Property Get TipActive() As Boolean

    '* Retrieving value of a property, Boolean responce (true/false).
    '* Syntax: BooleanVar = object.TipActive.

    TipActive = ToolActive

End Property

Public Property Get TipBackColor() As OLE_COLOR

    '* Retrieving value of a property, returns RGB as Long.
    '* Syntax: LongVar = object.BackColor.

    TipBackColor = ToolBackColor

End Property

Public Property Let TipBackColor(ByVal ToolData As OLE_COLOR)

    '* Assigning a value to the property, set RGB value as Long.
    '* Syntax: object.BackColor = RGB (as Long). Since 0 is Black (no RGB), and the API thinks 0 is
    '   the default color ("off" yellow), we need to "fudge" Black a bit (yes set bit "1" to "1",).
    '   I
    '   couldn't resist the pun!. So, in module or form code, if setting to Black, make it "1", if
    '   restoring the default color, make it "0".

    ToolBackColor = ConvertSystemColor(ToolData)
    Call PropertyChanged("TipBackColor")

End Property

Public Property Get TipCentered() As Boolean

    '* Retrieving value of a property, returns Boolean true/false.
    '* Syntax: BooleanVar = object.TipCentered.

    TipCentered = ToolCentered

End Property

Public Property Let TipCentered(ByVal ToolData As Boolean)

    '* Assigning a value to the property, Set Boolean true/false if ToolTip. Is TipCentered on the
    '   parent control.
    '* Syntax: object.TipCentered = True/False.

    ToolCentered = ToolData
    Call PropertyChanged("TipCentered")

End Property

Public Property Let TipForeColor(ByVal ToolData As OLE_COLOR)

    '* Assigning a value to the property, set RGB value as Long.
    '* Syntax: object.ForeColor = RGB(As Long).
    '* Since 0 is Black (no RGB), and the API thinks 0 is the default color ("off" yellow), we need
    '   to "fudge" Black a bit (yes set bit "1" to "1",). I couldn't resist the pun!. So, in module
    '   or
    '   form code, if setting to Black, make it "1" if restoring the default color, make it "0".
    '* Syntax: object.ForeColor = RGB(as long).

    ToolForeColor = ConvertSystemColor(ToolData)
    Call PropertyChanged("TipForeColor")

End Property

Public Property Get TipForeColor() As OLE_COLOR

    '* Retrieving value of a property, returns RGB value as Long.
    '* Syntax: LongVar = object.ForeColor.

    TipForeColor = ToolForeColor

End Property

Public Property Get TipIcon() As ToolIconType

    '* Retrieving value of a property, returns string.
    '* Syntax: StringVar = object.TipIcon.

    TipIcon = ToolIcon

End Property

Public Property Let TipIcon(ByVal ToolData As ToolIconType)

    '* Assigning a value to the property, set TipIcon TipStyle with type var.
    '* Syntax: object.TipIcon = IconStyle.
    '* TipIcon Styles are: INFO, WARNING And ERROR (TipNoIcom, TipIconInfo, TipIconWarning,
    '   TipIconError).

    ToolIcon = ToolData
    Call PropertyChanged("TipIcon")

End Property

Public Sub TipRemove()

    '* Kills Tool Tip Object.

    If (m_ltthWnd <> 0) Then Call DestroyWindow(m_ltthWnd)

End Sub

Public Property Let TipStyle(ByVal ToolData As ToolStyleEnum)

    '* Assigning a value to the property, set TipStyle param Standard or Balloon
    '* Syntax: object.TipStyle = TipStyle.

    TOOLSTYLE = ToolData
    Call PropertyChanged("TipStyle")

End Property

Public Property Get TipStyle() As ToolStyleEnum

    '* Retrieving value of a property, returns string.
    '* Syntax: StringVar = object.TipStyle.

    TipStyle = TOOLSTYLE

End Property

Public Property Let TipText(ByVal ToolData As String)

    '* Assigning a value to the property, Set as String.
    '* Syntax: object.TipText = StringVar.
    '* Multi line Tips are enabled in the Create sub.
    '* To change lines, just add a vbCrLF between text.
    '* ex. object.TipText = "Line 1 text" & vbCrLF & "Line 2 text".

    ToolText = ToolData
    Call PropertyChanged("TipText")

End Property

Public Property Get TipText() As String

    '* Retrieving value of a property, returns string..
    '* Syntax: StringVar = object.TipText.

    TipText = ToolText

End Property

Public Property Get TipTitle() As String

    '* Retrieving value of a property, returns string.
    '* Syntax: StringVar = object.TipTitle.

    TipTitle = ToolTitle

End Property

Public Property Let TipTitle(ByVal ToolData As String)

    '* Assigning a value to the property, set as string.
    '* Syntax: object.TipTitle = StringVar.

    ToolTitle = ToolData
    Call PropertyChanged("TipTitle")

End Property

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

'Track the mouse leaving the indicated window

  Dim tme As TRACKMOUSEEVENT_STRUCT

    If bTrack Then

        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hWndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            TrackMouseEvent tme
         Else
            TrackMouseEventComCtl tme
        End If

    End If

End Sub

Private Sub UpDate()

    '* Used to update tooltip parameters that require reconfiguration of subclass to envoke.

    If (ToolActive = True) Then Call CreateToolTip '* Refresh the object.

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    If (isEnabled = True) Then RaiseEvent Click

End Sub

Private Sub UserControl_Click()

    If (isHotTitle = False) Then
        Call Refresh(OfficeHighLight)
        RaiseEvent Click
    End If

End Sub

Private Sub UserControl_GotFocus()

    If (isHotTitle = False) Then
        isFocus = True
        Call Refresh(isState)
    End If

End Sub

Private Sub UserControl_Initialize()

  Dim OS As OSVERSIONINFO

    '* Get the operating system version for text drawing purposes.

    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_InitProperties()

    isAutoSizePic = False
    isBackColor = ConvertSystemColor(defBackColor)
    isBorderColor = ConvertSystemColor(defBorderColor)
    isButtonShape = defShape
    isCaption = Ambient.DisplayName
    isDisabledColor = ConvertSystemColor(defDisabledColor)
    isEnabled = True
    isFontAlign = ACenter
    isForeColor = ConvertSystemColor(defForeColor)
    isHighLightColor = ConvertSystemColor(defHighLightColor)
    isHotColor = ConvertSystemColor(defHotColor)
    isHotTitle = False
    isMultiLine = False
    isPictureAlign = ACenter
    isPictureSize = 16
    isSetBorder = False
    isSetGradient = False
    isSetHighLight = True
    isShadowText = False
    isShowFocus = False
    isSystemColor = True
    isXPos = 4
    isYPos = 4
    m_bGrayIcon = False
    Set g_Font = Ambient.Font
    Set isPicture = Nothing
    ToolActive = False
    ToolBackColor = vbInfoBackground
    ToolCentered = True
    ToolForeColor = vbInfoText
    ToolIcon = 1
    TOOLSTYLE = 1
    ToolTitle = "HACKPRO TM"
    ToolText = Extender.ToolTipText

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
     Case 13, 32 '* Enter.
        RaiseEvent Click

     Case 37, 38 '* Left Arrow and Up.
        Call SendKeys("+{TAB}")

     Case 39, 40 '* Right Arrow and Down.
        Call SendKeys("{TAB}")
    End Select

End Sub

Private Sub UserControl_LostFocus()

    If (isHotTitle = False) Then
        isFocus = False
        Call Refresh(isState)
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (isHotTitle = False) And (Button = vbLeftButton) And (isEnabled = True) Then
        Call Refresh(OfficeHot)
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim tmpState As Integer

    If (isEnabled = True) And (isHotTitle = False) Then
        If (IsMouseOver = True) Then
            Call Refresh(isState)
         Else
            tmpState = isState
            Call Refresh(OfficeNormal)
            isState = tmpState
        End If

    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next

    With PropBag
        AutoSizePicture = .ReadProperty("AutoSizePicture", False)
        BackColor = .ReadProperty("BackColor", ConvertSystemColor(defBackColor))
        BorderColor = .ReadProperty("BorderColor", ConvertSystemColor(defBorderColor))
        ButtonShape = .ReadProperty("ButtonShape", defShape)
        Caption = .ReadProperty("Caption", Ambient.DisplayName)
        CaptionAlign = .ReadProperty("CaptionAlign", &H0)
        DisabledColor = .ReadProperty("DisabledColor", ConvertSystemColor(defDisabledColor))
        Enabled = .ReadProperty("Enabled", True)
        ForeColor = .ReadProperty("ForeColor", ConvertSystemColor(defForeColor))
        GrayIcon = PropBag.ReadProperty("GrayIcon", True)
        HighLightColor = .ReadProperty("HighlightColor", ConvertSystemColor(defHighLightColor))
        HotColor = .ReadProperty("HotColor", ConvertSystemColor(defHotColor))
        HotTitle = .ReadProperty("HotTitle", False)
        MultiLine = .ReadProperty("MultiLine", False)
        PictureAlign = .ReadProperty("PictureAlign", &H0)
        PictureSize = .ReadProperty("PictureSize", 16)
        SetBorder = .ReadProperty("SetBorder", False)
        SetBorderH = .ReadProperty("SetBorderH", True)
        SetGradient = .ReadProperty("SetGradient", False)
        Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
        SetHighLight = .ReadProperty("SetHighLight", True)
        Set isPicture = .ReadProperty("Picture", Nothing)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        ShadowText = .ReadProperty("ShadowText", False)
        ShowFocus = .ReadProperty("ShowFocus", False)
        SystemColor = .ReadProperty("SystemColor", True)
        TipActive = .ReadProperty("TipActive", False)
        TipBackColor = .ReadProperty("TipBackColor", vbInfoBackground)
        TipCentered = .ReadProperty("TipCentered", True)
        TipForeColor = .ReadProperty("TipForeColor", vbInfoText)
        TipIcon = .ReadProperty("TipIcon", 1)
        TipStyle = .ReadProperty("TipStyle", 1)
        TipTitle = .ReadProperty("TipTitle", "HACKPRO TM")
        TipText = .ReadProperty("TipText", "")
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        XPosPicture = .ReadProperty("XPosPicture", 4)
        YPosPicture = .ReadProperty("YPosPicture", 4)
    End With

    If (Ambient.UserMode = True) Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not (bTrackUser32 = True) Then
            If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
                bTrack = False
            End If

        End If
        If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
            '* Start subclassing the UserControl.
            Call sc_Subclass(hWnd)
            Call sc_AddMsg(hWnd, WM_MOUSEMOVE)
            Call sc_AddMsg(hWnd, WM_MOUSELEAVE)
            Call sc_AddMsg(hWnd, WM_THEMECHANGED)
            Call sc_AddMsg(hWnd, WM_SYSCOLORCHANGE)
        End If

    End If

End Sub

Private Sub UserControl_Resize()

    If (isHotTitle = False) Then Call Refresh(isState) '* Call the Refresh Sub.

End Sub

Private Sub UserControl_Terminate()

'* The control is terminating - a good place to stop the subclasser

    On Error GoTo Catch
    Call TipRemove
    If (Ambient.UserMode = True) Then Call sc_Terminate '* Stop all subclassing.
Catch:

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next

    With PropBag
        Call .WriteProperty("AutoSizePicture", isAutoSizePic, False)
        Call .WriteProperty("BackColor", isBackColor, ConvertSystemColor(defBackColor))
        Call .WriteProperty("BorderColor", isBorderColor, ConvertSystemColor(defBorderColor))
        Call .WriteProperty("ButtonShape", isButtonShape, defShape)
        Call .WriteProperty("Caption", isCaption, Ambient.DisplayName)
        Call .WriteProperty("CaptionAlign", isFontAlign, &H0)
        Call .WriteProperty("DisabledColor", isDisabledColor, ConvertSystemColor(defDisabledColor))
        Call .WriteProperty("Enabled", isEnabled, True)
        Call .WriteProperty("Font", g_Font, Ambient.Font)
        Call .WriteProperty("ForeColor", isForeColor, ConvertSystemColor(defForeColor))
        Call .WriteProperty("GrayIcon", m_bGrayIcon, True)
        Call .WriteProperty("HighlightColor", isHighLightColor, _
            ConvertSystemColor(defHighLightColor))
        Call .WriteProperty("HotColor", isHotColor, ConvertSystemColor(defHotColor))
        Call .WriteProperty("HotTitle", isHotTitle, False)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", MousePointer, vbDefault)
        Call .WriteProperty("MultiLine", isMultiLine, False)
        Call .WriteProperty("Picture", isPicture, Nothing)
        Call .WriteProperty("PictureAlign", isPictureAlign, &H0)
        Call .WriteProperty("PictureSize", isPictureSize, 16)
        Call .WriteProperty("SetBorder", isSetBorder, False)
        Call .WriteProperty("SetBorderH", isSetBorderH, True)
        Call .WriteProperty("SetGradient", isSetGradient, False)
        Call .WriteProperty("SetHighLight", isSetHighLight, True)
        Call .WriteProperty("ShadowText", isShadowText, False)
        Call .WriteProperty("ShowFocus", isShowFocus, False)
        Call .WriteProperty("SystemColor", isSystemColor, True)
        Call .WriteProperty("TipActive", ToolActive, False)
        Call .WriteProperty("TipBackColor", ToolBackColor, vbInfoBackground)
        Call .WriteProperty("TipCentered", ToolCentered, True)
        Call .WriteProperty("TipForeColor", ToolForeColor, vbInfoText)
        Call .WriteProperty("TipIcon", ToolIcon, 1)
        Call .WriteProperty("TipStyle", TOOLSTYLE, 1)
        Call .WriteProperty("TipText", ToolText, "")
        Call .WriteProperty("TipTitle", ToolTitle, "HACKPRO TM")
        Call .WriteProperty("XPosPicture", isXPos, 4)
        Call .WriteProperty("YPosPicture", isYPos, 4)
    End With

    On Error GoTo 0

End Sub

Public Property Let XPosPicture(ByVal theXPos As Integer)

'* English: Returns/Sets the Position X of the image.

    '* Devuelve ó establece la Posición X de la imagen.

    isXPos = theXPos
    Call PropertyChanged("XPosPicture")
    Call Refresh(isState)

End Property

Public Property Get XPosPicture() As Integer

    XPosPicture = isXPos

End Property

Public Property Let YPosPicture(ByVal theYPos As Integer)

'* English: Returns/Sets the Position Y of the image.

    '* Devuelve ó establece la Posición Y de la imagen.

    isYPos = theYPos
    Call PropertyChanged("YPosPicture")
    Call Refresh(isState)

End Property

Public Property Get YPosPicture() As Integer

    YPosPicture = isYPos

End Property

Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)

'-The following routines are exclusively for the sc_ subclass routines----------------------------
'Add the message to the specified table of the window handle

  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

    nBase = z_ScMem                                                            'Remember z_ScMem so
    '   that we can restore its value on exit
    z_ScMem = zData(nTable)                                                    'Map zData() to the
    '   specified table

    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are
        '   being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table
        '   entry count to ALL_MESSAGES
     Else
        nCount = zData(0)                                                       'Get the current
        '   table entry count
        If nCount >= MSG_ENTRIES Then                                           'Check for message
            '   table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const" & _
                " MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                                     'Loop through the
            '   table entries
            If zData(i) = 0 Then                                                  'If the element
                '   is free...
                zData(i) = uMsg                                                     'Use this
                '   element
                GoTo Bail                                                           'Bail
             ElseIf zData(i) = uMsg Then                                           'If the message
                '   is already in the table...
                GoTo Bail                                                           'Bail
            End If

        Next i                                                                  'Next message table
        '   entry

        nCount = i                                                              'On drop through: i
        '   = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                    'Store the message
        '   in the appended table entry
    End If

    zData(0) = nCount                                                         'Store the new table
    '   entry count
Bail:
    z_ScMem = nBase                                                           'Restore the value of
    '   z_ScMem

End Sub

Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc

  Dim bSub  As Byte                                                         'Value we expect to

    '   find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit

    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of
    '   the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class
        '   method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form
            '   method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a
                '   UserControl method
                Exit Function                                                       'Bail...
            End If

        End If
    End If

    i = i + 4                                                                 'Bump to the next
    '   entry
    j = i + 1024                                                              'Set a reasonable
    '   limit, scan 256 vTable entries

    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address
        '   stored in this vTable entry

        If IsBadCodePtr(nAddr) Then                                             'Is the entry an
            '   invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the
            '   specified vTable entry address
            Exit Do                                                               'Bad method
            '   signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte
        '   pointed to by the vTable entry
        If bVal <> bSub Then                                                    'If the byte doesn
            '  't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the
            '   specified vTable entry address
            Exit Do                                                               'Bad method
            '   signature, quit loop
        End If

        i = i + 4                                                             'Next vTable entry
    Loop

End Function

Private Property Get zData(ByVal nIndex As Long) As Long

    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)

    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)

'Delete the message from the specified table of the window handle

  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

    nBase = z_ScMem                                                           'Remember z_ScMem so
    '   that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the
    '   specified table

    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are
        '   being deleted from the table...
        zData(0) = 0                                                            'Zero the table
        '   entry count
     Else
        nCount = zData(0)                                                       'Get the table
        '   entry count

        For i = 1 To nCount                                                     'Loop through the
            '   table entries
            If zData(i) = uMsg Then                                               'If the message
                '   is found...
                zData(i) = 0                                                        'Null the msg
                '   value -- also frees the element for re-use
                GoTo Bail                                                           'Bail
            End If

        Next i                                                                  'Next message table
        '   entry

        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If

Bail:
    z_ScMem = nBase                                                           'Restore the value of
    '   z_ScMem

End Sub

Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)

'Error handler

    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    'MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." &
    '   sRoutine

End Sub

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long

'Return the address of the specified DLL/procedure

    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified
    '   procedure address
    Debug.Assert zFnAddr                                                      'In the IDE, validate
    '   that the procedure address was located

End Function

Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long

'Map zData() to the thunk address for the specified window handle

    If z_Funk Is Nothing Then                                                 'Ensure that
        '   subclassing has been started
        zError "zMap_hWnd", "Subclassing hasn't been started"
     Else
        On Error GoTo Catch                                                     'Catch unsubclassed
        '   window handles
        z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk
        '   address
        zMap_hWnd = z_ScMem
    End If

    Exit Function                                                             'Exit returning the

    '   thunk address

Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"

End Function

Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean

'Probe at the specified start address for a method signature

  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long

    nAddr = nStart                                                            'Start address
    nLimit = nAddr + 32                                                       'Probe eight entries

    Do While nAddr < nLimit                                                   'While we've not
        '   reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable
        '   entry

        If nEntry <> 0 Then                                                     'If not an
            '   implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value
            '   pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a
                '   native or pcode method signature
                nMethod = nAddr                                                     'Store the
                '   vTable entry
                bSub = bVal                                                         'Store the
                '   found method signature
                zProbe = True                                                       'Indicate
                '   success
                Exit Function                                                       'Return
            End If

        End If

        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop

End Function

Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------

'*************************************************************************************************
    '* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
    '*              you will know unless the callback for the uMsg value is specified as
    '*              MSG_BEFORE_AFTER (both before and after the original WndProc).
    '* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
    '*              message being passed to the original WndProc and (if set to do so) the after
    '*              original WndProc callback.
    '* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
    '*              and/or, in an after the original WndProc callback, act on the return value as
    '   set
    '*              by the original WndProc.
    '* lng_hWnd   - Window handle.
    '* uMsg       - Message value.
    '* wParam     - Message related data.
    '* lParam     - Message related data.
    '* lParamUser - User-defined callback parameter
'*************************************************************************************************

    Select Case uMsg
     Case WM_MOUSEMOVE
        If (isSetHighLight = False) Then Exit Sub

        If Not (isInCtrl = True) Then
            isInCtrl = True
            Call TrackMouseLeave(lng_hWnd)
            Call Refresh(OfficeHighLight)
            Call UpDate
            RaiseEvent MouseEnter
        End If

     Case WM_MOUSELEAVE
        If (isSetHighLight = False) Then Exit Sub
        isInCtrl = False
        Call Refresh(OfficeNormal)
        RaiseEvent MouseLeave

     Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
        Call UserControl_Resize
        RaiseEvent ChangedTheme
    End Select

End Sub
