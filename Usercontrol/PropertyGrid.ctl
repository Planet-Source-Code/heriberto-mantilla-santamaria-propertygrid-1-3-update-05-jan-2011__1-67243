VERSION 5.00
Begin VB.UserControl PropertyGrid 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "PropertyGrid.ctx":0000
   Begin PropertyGridDemo.ucUpDownBox UpDown 
      Height          =   285
      Left            =   1665
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   503
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PropertyGridDemo.CoolList lstFX1 
      Height          =   1560
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   2752
      Appearance      =   0
      BorderStyle     =   0
      ScrollBarWidth  =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      ItemHeight      =   13
      ShadowColorText =   6844272
   End
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1455
      Visible         =   0   'False
      Width           =   1395
   End
   Begin PropertyGridDemo.isButton isBttAction 
      Height          =   270
      Left            =   2835
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      Style           =   0
      Caption         =   " ..."
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   1
      ToolTipType     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PropertyGridDemo.SComboBox SCmb 
      Height          =   315
      Left            =   2160
      Top             =   2355
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      AppearanceCombo =   3
      ArrowColor      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      XpAppearance    =   0
   End
   Begin PropertyGridDemo.McCalendar McCalendar 
      Height          =   3150
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Animate         =   -1  'True
      CalendarHeight  =   192
      CalendarBackCol =   16743805
      WeekDaySunCol   =   12640511
      YearGradient    =   -1  'True
      BorderColor     =   4210752
   End
End
Attribute VB_Name = "PropertyGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################'
' Title:     PropertyGrid                                                   '
' Author:    Heriberto Mantilla Santamaría                                  '
' Company:   HACKPRO TM                                                     '
' Created:   20/11/06                                                       '
' Version:   1.3 (03 January 2011)                                            '
'                                                                           '
'           Copyright © 2006 - 2011 HACKPRO TM. All rights reserved         '
'###########################################################################'

'---------------------------------------------------------------------------'
'                               PropertyGrid                                '
'---------------------------------------------------------------------------'
'  CREDITS                                                                  '
'---------------------------------------------------------------------------'
' Paul Caton                                                                '
' Steve McMahon                                                             '
' Richard Mewett                                                            '
' Paul R. Territo, Ph.D                                                     '
' Carles P.V                                                                '
' Fred.cpp                                                                  '
' Jim Jose                                                                  '
' Matthew R. Usner                                                          '
' Calendar (I don't found the name of PSC Contributor in the code           '
'           CodeId=61147, I put the name later)                             '
'---------------------------------------------------------------------------'
' DEDICATION                                                                '
'---------------------------------------------------------------------------'
' Dedicated to the prettiest girl that I have known XD, T.Q.M mibi          '
'---------------------------------------------------------------------------'
' IMPORTANT NOTE                                                            '
'---------------------------------------------------------------------------'
' Feel free to use this UC in your App's, provided ALL credits remain       '
' intact. Only dishonorable thieves download code that REAL programmers     '
' work hard to write and freely share with their programming peers, then    '
' remove the comments and claim that they wrote the code.                   '
'---------------------------------------------------------------------------'
' HISTORY                                                                   '
'---------------------------------------------------------------------------'
' (*) Fixed SComboBox Control.                                              '
' (*) Fixed (+) Treeview draw.                                              '
' (*) Fixed DrawXPTheme Function in Win98.                                  '
' (*) Now work keys in the CoolList with UC PropertyGrid.                   '
' (*) Fixed help resize.                                                    '
' (*) Fixed isButton when is pressed.                                       '
' (*) Remove inneccesary code in SComboBox.                                 '
' (*) Remove API's and Const's aren't used.                                 '
' (*) Fixed the text focus.                                                 '
' (*) Fixed call events.                                                    '
' (*) Fixed Scrollbar Visible/Hide.                                         '
' (*) Added McCalendar UC by Jim Jose.                                      '
' (*) Edit a little McCalendar.                                             '
' (*) Fixed little error's.                                                 '
' (*) Added simple click in (+)/(-).                                        '
' (*) Added Redraw if the last select don't appear well.                    '
' (*) Now is compatibility with VB 5.0 and VB 6.0.                          '
' (*) Fixed called events.                                                  '
' (*) Added Subs and edited some.                                           '
' (*) Fixed text Keypress values.                                           '
' (*) Add 2 new themes.                                                     '
' (*) Fixed font property.                                                  '
' (*) Fixed minor bugs.                                                     '
' (*) Add new Properties and functions                                      '
'---------------------------------------------------------------------------'
'                 Testing only in Win98SE and WinXP SP2                     '
'---------------------------------------------------------------------------'

Option Explicit

Private Const thisVersion = "1.3"

'----------------------------------------------------------------------------------------->
' Scrollbar's Steve McMahon (steve@dogma.demon.co.uk)
'----------------------------------------------------------------------------------------->
Private isXp                As Boolean
Private lStyle              As Long
Private m_bNoFlatScrollBars As Boolean
Private m_hWnd              As Long
Private m_lSmallChange      As Long
Private Value1              As Long

' Hack for XP Crash with VB6 controls:
Private m_hMod As Long

Private Declare Function CreateWindowEx Lib "user32" _
        Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        lpParam As Any) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        LPSCROLLINFO As SCROLLINFO, _
        ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" ( _
        ByVal hWnd As Long, _
        ByVal Index As Long, _
        ByVal newValue As Long, _
        ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        ByVal fRedraw As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal n As Long, _
        LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd _
   As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function MoveWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal bRepaint As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal n As Long, _
        lpcScrollInfo As SCROLLINFO, _
        ByVal BOOL As Boolean) As Long
Private Declare Function ShowScrollBar Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal wBar As Long, _
        ByVal bShow As Long) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hWnd As Long) As Long

'Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As String, _
        ByRef lpUsedDefaultChar As Long) As Long
'Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As String, _
        ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long) As Long

'Private Const CP_UTF8 As Long = 65001   ' UTF-8 translation

Private Const FSB_ENCARTA_MODE      As Long = 1
Private Const SB_BOTTOM             As Long = 7
Private Const SB_CTL                As Long = 2
Private Const SB_ENDSCROLL          As Long = 8
Private Const SB_LEFT               As Long = 6
Private Const SB_LINEDOWN           As Long = 1
Private Const SB_LINELEFT           As Long = 0
Private Const SB_LINERIGHT          As Long = 1
Private Const SB_LINEUP             As Long = 0
Private Const SB_PAGEDOWN           As Long = 3
Private Const SB_PAGELEFT           As Long = 2
Private Const SB_PAGERIGHT          As Long = 3
Private Const SB_PAGEUP             As Long = 2
Private Const SB_RIGHT              As Long = 7
Private Const SB_THUMBTRACK         As Long = 5
Private Const SB_TOP                As Long = 6
Private Const SB_VERT               As Long = 1
Private Const SBS_HORZ              As Long = &H0&
Private Const SIF_RANGE             As Long = &H1
Private Const SIF_PAGE              As Long = &H2
Private Const SIF_POS               As Long = &H4
Private Const SIF_TRACKPOS          As Long = &H10
Private Const SIF_ALL As Long = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const WM_VSCROLL            As Long = &H115
Private Const WS_CHILD              As Long = &H40000000
Private Const WSB_PROP_HSTYLE       As Long = &H200&
Private Const WS_VISIBLE            As Long = &H10000000

' Scroll bar stuff
Private Type SCROLLINFO
    cbSize    As Long
    fMask     As Long
    nMin      As Long
    nMax      As Long
    nPage     As Long
    nPos      As Long
    nTrackPos As Long
End Type

'----------------------------------------------------------------------------------------->

'----------------------------------------------------------------------------------------->
' uSelfSub By Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'----------------------------------------------------------------------------------------->

'-Selfsub declarations----------------------------------------------------------------------------

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
Private z_Sc(64)            As Long                                         'Thunk machine-code
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
'----------------------------------------------------------------------------------------->

Private Const WM_ACTIVATE           As Long = &H6

Private Const WM_CTLCOLORSCROLLBAR  As Long = &H137

Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_MOUSELEAVE         As Long = &H2A3

Private Const WM_MOUSEWHEEL         As Long = &H20A

Private Const WM_EXITSIZEMOVE       As Long = &H232
Private Const WM_MOVING             As Long = &H216
Private Const WM_SIZING             As Long = &H214

Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_MBUTTONDOWN        As Long = &H207
Private Const WM_NCLBUTTONDOWN      As Long = &HA1
Private Const WM_RBUTTONDOWN        As Long = &H204

Private Const WM_SYSCOLORCHANGE     As Long = &H15
Private Const WM_THEMECHANGED       As Long = &H31A

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
Private bInCtrl               As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" ( _
        lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" _
        Alias "_TrackMouseEvent" ( _
        lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'----------------------------------------------------------------------------------------->
' By Richard Mewett
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61438&lngWId=1
'----------------------------------------------------------------------------------------->

' Declares for Unicode support.
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
End Type

Private mWindowsNT     As Boolean
'----------------------------------------------------------------------------------------->

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    rLeft   As Long
    rTop    As Long
    rRight  As Long
    rBottom As Long
End Type

Private Type RGB
    Red     As Integer
    Green   As Integer
    Blue    As Integer
End Type

Public Enum PropertyItemType
    PropertyItemBool = &H0
    PropertyItemColor = &H1
    PropertyItemDate = &H2
    PropertyItemFolder = &H3
    PropertyItemFolderFile = &H4
    PropertyItemFont = &H5
    PropertyItemForm = &H6
    PropertyItemNumber = &H7
    PropertyItemPicture = &H8
    PropertyItemString = &H9
    PropertyItemStringList = &HA
    PropertyItemStringReadOnly = &HB
    PropertyItemUpDown = &HC
    PropertyItemCheckBox = &HD
End Enum

Public Enum PropertyStyleSort
    Categorized = &H0
    Alphabetical = &H1
    NoSort = &H2
End Enum

Public Enum PropertyGridStyle
    VBTheme = &H0
    NormalTheme = &H1
    OfficeTheme = &H2
    AutoDeskTheme = &H3
End Enum

Public Enum StringStyle
    ItemNormal = &H0
    ItemNumeric = &H1
    ItemPassword = &H2
    ItemLowerCase = &H3
    ItemUpperCase = &H4
End Enum

Public Enum ColorStyle
    RGB_Color = &H0
    VB_Color = &H1
    Web_Color = &H2
    C_Color = &H3
    Delphi_Color = &H4
    Java_Color = &H5
    PhotoShop_Color = &H6
End Enum

Private Const COLOR_BTNFACE    As Long = 15
Private Const DT_NOPREFIX      As Long = &H800
Private Const DT_SINGLELINE    As Long = &H20
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_WORDBREAK     As Long = &H10

Private Const GWL_EXSTYLE      As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80

Private Const SIZE_VARIANCE    As Long = 24
Private Const CURSOR_ARROW_VERTICAL_SPLITTER As String = "ARROW_VERTICAL_SPLITTER"
Private Const CURSOR_ARROW_HORIZONTAL_SPLITTER As String = "ARROW_HORIZONTAL_SPLITTER"
Private Const CURSOR_HAND As String = "HAND"

Private Type TypeCateg
    Expand        As Boolean
    Key           As String
    Title         As String
    ToolTipText   As String
End Type

Private Type TypeChildItem
    Filters       As String
    ItemValue     As String
    KeyCategory   As String
    KeyName       As String
    StyleString   As StringStyle
    theFont       As New StdFont
    Title         As String
    ToolTipText   As String
    TypeGrid      As PropertyItemType
    Value         As Variant
End Type

Private CountCatg   As Long
Private CountChild  As Long

Private R As Byte
Private G As Byte
Private B As Byte

Private m_Style As StringStyle

' Events of the UC.
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event Expand(ByVal ExpandedOnce As Boolean, ByVal KeyCategory As String)
Public Event FormClick(ByRef Value As Variant, ByVal KeyCategory As String, ByVal Title As String, ByVal X As Integer, ByVal Y As Integer)
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event PropertySortChanged(ByVal PropertySort As PropertyStyleSort)
Public Event RightClick(ByVal Position As Integer)
Public Event Scroll()
Public Event SelectionChanged(ByVal KeyCategory As String, ByVal Title As String, ByVal Index As _
    Integer)
Public Event ValueChanged(ByVal KeyCategory As String, ByVal Title As String, ByVal Value As _
    Variant, ByVal theFont As StdFont)

Private Category()               As TypeCateg
Private ChildItem()              As TypeChildItem
Private ChildOrd()               As TypeChildItem

Private b_Focus                  As Boolean
Private bChanged                 As Boolean
Private bExpand                  As Boolean
Private bRedraw                  As Boolean
Private bResize                  As Boolean
Private bResY                    As Boolean
Private bSubClass                As Boolean
Private bUserFocus               As Boolean
Private bVScroll                 As Boolean
Private lButton                  As Integer
Private lCateg                   As Long
Private lChild                   As Long
Private lDblClick                As Boolean
Private setYet                   As Boolean
Private lItemSelected            As Long
Private lLastItemSelected        As Long
Private lTCaption                As String
Private lTTitle                  As String
Private lTotalItems              As Long
Private lXPos                    As Integer
Private lYPos                    As Integer
Private RGBColor                 As RGB
Private xSplitter                As Integer
Private xSplitterY               As Integer
Private m_bEnabled               As Boolean
Private m_bFixedSplit            As Boolean
Private m_bHelpVisible           As Boolean
Private m_bTrimPath              As Boolean
Private m_iHelpHeight            As Integer
Private m_iSplitterPos           As Single
Private m_lAutoFilter            As Boolean
Private m_lBackColor             As OLE_COLOR
Private m_lFont                  As StdFont
Private m_lLineColor             As OLE_COLOR
Private m_ListIndex              As Long
Private m_iHeight                As Long
Private m_lHelpBackColor         As OLE_COLOR
Private m_lHelpForeColor         As OLE_COLOR
Private m_lViewBackColor         As OLE_COLOR
Private m_lViewCategoryForeColor As OLE_COLOR
Private m_lViewForeColor         As OLE_COLOR
Private m_sDecimalSymbol         As String
Private m_SetColorStyle          As ColorStyle
Private m_StyleButton            As isbStyle
Private m_StyleComboBox          As ComboAppearance
Private m_StylePropertyGrid      As PropertyGridStyle
Private m_vPropertySort          As PropertyStyleSort
Private NoShowList               As Boolean
Private hTheme                   As Long             '* hTheme Handle.

Private Declare Function DrawFrameControl Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal un1 As Long, _
        ByVal un2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" ( _
        ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
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
Private Declare Function FillRect Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" ( _
        ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long '* XP detection.
Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" ( _
        lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowLong Lib "user32" _
        Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function LineTo Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" ( _
        lpRect As RECT, _
        ByVal X As Long, _
        ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
        ByVal lOleColor As Long, _
        ByVal lHPalette As Long, _
        lColorRef As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" ( _
        ByVal hWndChild As Long, _
        ByVal hWndNewParent As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long

Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" ( _
        ByVal hTheme As Long, _
        ByVal lhDC As Long, _
        ByVal iPartId As Long, _
        ByVal iStateId As Long, _
        pRect As RECT, _
        pClipRect As RECT) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" ( _
        ByVal hWnd As Long, _
        ByVal hDC As Long, _
        prc As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" ( _
        ByVal hWnd As Long, _
        ByVal pszClassList As Long) As Long

'----------------------------------------------------------------------------------------->
' By Paul R. Territo, Ph.D
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63905&lngWId=1
'----------------------------------------------------------------------------------------->
Private Const WM_USER As Long = &H400

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const LF_FACESIZE = 33
Private Const MAX_PATH As Long = 4096 '260
Private Const OUT_DEFAULT_PRECIS = 0
Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONCHECK = &H0
Private Const DFCS_BUTTONRADIO = &H4
Private Const DFCS_BUTTON3STATE = &H10
Private Const DFCS_CHECKED         As Long = &H400
Private Const DFCS_INACTIVE        As Long = &H100

Public Enum ColorDialogFlags
    '   ShowColor Flags
    RGBInit = &H1
    FullOpen = &H2
    PreventFullOpen = &H4
    ShowHelp = &H8
    EnableHook = &H10
    EnableTemplate = &H20
    EnableTemplateHandle = &H40
    SolidColor = &H80
    AnyColor = &H100
    '   Custom Non-Win32 Flags which are a Combinations of Flags
    ShowColor_Default = FullOpen Or AnyColor Or RGBInit
End Enum

Public Enum FolderDialogFlags
    ReturnOnlyFSDirs = &H1
    DontGoBelowDomain = &H2
    StatusText = &H4
    ReturnFSAncestors = &H8
    EditBox = &H10
    Validate = &H20
    NewDialogStyle = &H40
    UseNewUI = (NewDialogStyle Or EditBox)
    BrowseIncludeURLs = &H80
    UahInt = &H100
    NoneWFolderButton = &H200
    NoTranslateTargets = &H400
    BrowseForComputer = &H1000
    BrowseForPrinter = &H2000
    BrowseIncludeFiles = &H4000
    Shareable = &H8000
    ShowFolder_Default = NewDialogStyle Or BrowseForComputer
End Enum

Public Enum FontDialogFlags
    '   ShowFont Flags
    ScreenFonts = &H1
    PrinterFonts = &H2
    Both = (ScreenFonts Or PrinterFonts)
    ShowHelp = &H4
    EnableHook = &H8
    EnableTemplate = &H10
    EnableTemplateHandle = &H20
    InitToLogFontStruct = &H40
    UseStyle = &H80
    Effects = &H100
    Apply = &H200
    AnsiOnly = &H400
    ScriptsOnly = AnsiOnly
    NoVectorFonts = &H800
    NoOEMFonts = NoVectorFonts
    NoSimulations = &H1000
    LimitSize = &H2000
    FixedPitchOnly = &H4000
    WYSIWYG = &H8000 '  Must Also Have Screenfonts Printerfonts
    ForceFontExist = &H10000
    ScalableOnly = &H20000
    TTonly = &H40000
    NoFaceSel = &H80000
    NoStyleSel = &H100000
    NoSizeSel = &H200000
    SelectScript = &H400000
    NoScriptSel = &H800000
    NoVertFonts = &H1000000
    '   Custom Non-Win32 Flags which are a Combinations of Flags
    ShowFont_Default = Both Or Effects Or ForceFontExist Or InitToLogFontStruct Or LimitSize
End Enum

Private Type CHOOSECOLORS
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Enum OpenSaveDialogFlags
    '   ShowOpen / ShowSave Flags
    ReadOnly = &H1
    OverwritePrompt = &H2
    HideReadOnly = &H4
    NoChangeDir = &H8
    ShowHelp = &H10
    EnableHook = &H20
    EnableTemplate = &H40
    EnableTemplateHandle = &H80
    NoValidate = &H100
    AllowMultiselect = &H200
    ExtensionDifferent = &H400
    PathMustExist = &H800
    FileMustExist = &H1000
    Createprompt = &H2000
    ShareAware = &H4000
    NoReadOnlyReturn = &H8000
    NoTestFileCreate = &H10000
    NoNetworkButton = &H20000
    NoLongNames = &H40000
    Explorer = &H80000
    LongNames = &H200000
    NoDeReferenceLinks = &H100000
    '   Custom Non-Win32 Flags Which Are A Combinations Of Flags
    ShowOpen_Default = Explorer Or LongNames Or Createprompt Or NoDeReferenceLinks Or HideReadOnly
    ShowSave_Default = Explorer Or LongNames Or OverwritePrompt Or HideReadOnly
End Enum

Private Type OPENFILENAME
    nStructSize As Long
    hWndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Private Type CHOOSEFONTS
    lStructSize As Long
    hWndOwner As Long           '  caller's window handle
    hDC As Long                 '  printer DC/IC or NULL
    lpLogFont As Long           '  ptr. to a LOGFONT struct
    iPointSize As Long          '  10 * size in points of selected font
    Flags As Long               '  enum. type flags
    rgbColors As Long           '  returned text color
    lCustData As Long           '  data passed to hook fn.
    lpfnHook As Long            '  ptr. to hook function
    lpTemplateName As String    '  custom template name
    hInstance As Long           '  instance handle of.EXE that
    lpszStyle As String         '  return the style field here
    nFontType As Integer        '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long            '  minimum pt size allowed &
    nSizeMax As Long            '  max pt size allowed if
End Type

Private Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled      As Boolean
End Type

Private Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Private Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Private Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFacename(LF_FACESIZE) As Byte
End Type

'   Custom Colors Dialog Array
Private CustomColors(0 To (16 * 4 - 1)) As Byte

'   Private Dialog Structure Definitions
Private ColorDialog   As CHOOSECOLORS
Private FileDialog    As OPENFILENAME
Private FontDialog    As CHOOSEFONTS
Private m_ColorFlags  As ColorDialogFlags
Private m_Filters     As String
Private m_FileFlags   As OpenSaveDialogFlags
Private m_FolderFlags As FolderDialogFlags
Private m_Font        As StdFont
Private m_FontColor   As OLE_COLOR
Private m_FontFlags   As FontDialogFlags
Private m_MultiSelect As Boolean
Private m_Path        As String

'   Private API Declarations
Private Declare Function ChooseColor Lib "comdlg32.dll" _
        Alias "ChooseColorA" ( _
        pChoosecolor As CHOOSECOLORS) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" _
        Alias "ChooseFontA" ( _
        pChoosefont As CHOOSEFONTS) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" ( _
        pOpenfilename As OPENFILENAME) As Long
'Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As
'   Long
Private Declare Function lStrCat Lib "kernel32" _
        Alias "lstrcatA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" ( _
        ByVal pidl As Long, _
        ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        Alias "SHBrowseForFolderA" ( _
        lpBrowseInfo As BROWSEINFO) As Long

' Matthew R. Usner.
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Const DT_CALCRECT As Long = &H400
Private Const DT_LEFT As Long = &H0

Private Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color.

'-----------------------------------------------------------------------------
' SComboBox Properties Colors
'-----------------------------------------------------------------------------
Public Sub SComboBoxColors(Optional ByVal ArrowColor As OLE_COLOR = -1, _
        Optional ByVal BackColor As OLE_COLOR = -1, _
        Optional ByVal DisabledColor As OLE_COLOR = -1, _
        Optional ByVal GradientColor1 As OLE_COLOR = -1, _
        Optional ByVal GradientColor2 As OLE_COLOR = -1, _
        Optional ByVal HighLightBorderColor As OLE_COLOR = -1, _
        Optional ByVal HighLightColorText As OLE_COLOR = -1, _
        Optional ByVal NormalBorderColor As OLE_COLOR = -1, _
        Optional ByVal NormalColorText As OLE_COLOR = -1, _
        Optional ByVal SelectBorderColor As OLE_COLOR = -1)
    
    With SCmb
        If (ArrowColor <> -1) Then .ArrowColor = ArrowColor
        If (BackColor <> -1) Then .BackColor = BackColor
        If (DisabledColor <> -1) Then .DisabledColor = DisabledColor
        If (GradientColor1 <> -1) Then .GradientColor1 = GradientColor1
        If (GradientColor2 <> -1) Then .GradientColor2 = GradientColor2
        If (HighLightBorderColor <> -1) Then .HighLightBorderColor = HighLightBorderColor
        If (HighLightColorText <> -1) Then .HighLightColorText = HighLightColorText
        If (NormalBorderColor <> -1) Then .NormalBorderColor = NormalBorderColor
        If (NormalColorText <> -1) Then .NormalColorText = NormalColorText
        If (SelectBorderColor <> -1) Then .SelectBorderColor = SelectBorderColor
    End With
    
End Sub

'-----------------------------------------------------------------------------
' isButton Properties Colors
'-----------------------------------------------------------------------------
Public Sub isButtonColors(Optional ByVal BackColor As OLE_COLOR = -1, _
        Optional ByVal FontColor As OLE_COLOR = -1, _
        Optional ByVal FontHighlightColor As OLE_COLOR = -1, _
        Optional ByVal HighlightColor As OLE_COLOR = -1, _
        Optional ByVal UseCustomColors As Boolean = True)
    
    With isBttAction
        If (BackColor <> -1) Then .BackColor = BackColor
        If (FontColor <> -1) Then .FontColor = FontColor
        If (FontHighlightColor <> -1) Then .FontHighlightColor = FontHighlightColor
        If (HighlightColor <> -1) Then .HighlightColor = HighlightColor
        .UseCustomColors = UseCustomColors
    End With
    
End Sub

Public Sub AboutBox()

    MsgBox "Developed by Heriberto Mantilla Santamaría" & vbCrLf & "PropertyGrid 1.0", _
        vbInformation + vbOKOnly, Ambient.DisplayName

End Sub

Public Sub AddCategory(ByVal Key As String, _
                       ByVal Title As String, _
                       Optional ByVal ToolTipText As String, _
                       Optional ByVal Expand As Boolean = False)

    If (LenB(Trim$(Key)) > 0) And (FindCategory(Key) = False) Then
        ReDim Preserve Category(CountCatg)
        Category(CountCatg).Expand = Expand
        Category(CountCatg).Key = Key
        Category(CountCatg).Title = Title
        Category(CountCatg).ToolTipText = ToolTipText
        CountCatg = CountCatg + 1
    Else
        Err.Raise 512, "AddCategory", "The key already exist or is empty."
    End If

End Sub

Public Sub AddChildItem(ByVal KeyCategory As String, ByVal TypeGridItem As PropertyItemType, ByVal _
                        Title As String, ByVal Value As Variant, Optional ByVal ItemValue As _
                        String = vbNullString, Optional ByVal ToolTipText As String = vbNullString, _
                        Optional ByVal StyleText As StringStyle = vbNormal, _
                        Optional ByVal sFilters As String = vbNullString, _
                        Optional KeyName As String = vbNullString)
    
    If (LenB(KeyName) = 0) Then
        KeyName = Title
    End If

    If (FindCategory(KeyCategory) = True) And (FindChild(Title, KeyCategory) = False) Then
        ReDim Preserve ChildItem(CountChild)
        ChildItem(CountChild).ItemValue = ItemValue
        ChildItem(CountChild).KeyCategory = KeyCategory
        ChildItem(CountChild).Title = Title
        ChildItem(CountChild).ToolTipText = ToolTipText
        ChildItem(CountChild).TypeGrid = TypeGridItem
        ChildItem(CountChild).Filters = sFilters
        ChildItem(CountChild).KeyName = KeyName

        If (TypeGridItem = PropertyItemFont) Or (TypeGridItem = PropertyItemPicture) Then
            Set ChildItem(CountChild).Value = Value
        Else
            ChildItem(CountChild).Value = Value
        End If
        
        If (TypeGridItem = PropertyItemNumber) Then
            ChildItem(CountChild).StyleString = ItemNumeric
        Else
            ChildItem(CountChild).StyleString = StyleText
        End If

        CountChild = CountChild + 1
        Let ChildOrd() = ChildItem()

        If (m_vPropertySort = Alphabetical) Then
            Insertion
        End If

    Else
        Err.Raise 513, "AddChildItem", "The KeyCategory doesn't exists or the key Title already exist or is empty."
    End If

End Sub

Public Sub AddToList(ByVal KeyCategory As String, ByVal Title As String, ByVal Item As String)

  Dim DimChild As Long
  Dim mUbound As Long

    If (FindCategory(KeyCategory) = True) And (FindChild(Title, KeyCategory) = True) Then
        DimChild = GetChildIndex(Title, KeyCategory)
        If (IsArray(ChildOrd(DimChild).Value) = True) And (FindInList(KeyCategory, Title, Item) = -1) Then
            mUbound = UBound(ChildOrd(DimChild).Value) + 1
            ReDim Preserve ChildOrd(DimChild).Value(mUbound)
            ChildOrd(DimChild).Value(mUbound) = Item
        Else
            Err.Raise 513, "AddToList", "The Item already exists."
        End If
        
    Else
        Err.Raise 513, "AddToList", "The KeyCategory doesn't exists or the key Title doesn't exists either."
    End If

End Sub

Private Sub APILine(ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal X2 As Long, _
                    ByVal Y2 As Long, _
                    ByVal lColor As Long)

  Dim PT As POINTAPI
  Dim hPen As Long
  Dim hPenOld As Long

On Error Resume Next
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hDC, hPen)

    MoveToEx UserControl.hDC, X1, Y1, PT
    LineTo UserControl.hDC, X2, Y2
    SelectObject hDC, hPenOld
    DeleteObject hPen
On Error GoTo 0

End Sub

Private Function APIRectangle(ByVal hDC As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal W As Long, _
                              ByVal H As Long, _
                              Optional ByVal lColor As OLE_COLOR = -1) As Long

  Dim hPen As Long
  Dim hPenOld As Long
  Dim PT   As POINTAPI
  
On Error Resume Next
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hDC, hPen)

    MoveToEx hDC, X, Y, PT

    LineTo hDC, X + W, Y
    LineTo hDC, X + W, Y + H
    LineTo hDC, X, Y + H
    LineTo hDC, X, Y

    SelectObject hDC, hPenOld
    DeleteObject hPen
On Error GoTo 0

End Function

Public Property Get AutoFilter() As Boolean
    
    '---------------------------------------------------------------------------------------
    ' Procedure : AutoFilter
    ' DateTime  : 15/06/2007 11:41
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------
    
    AutoFilter = m_lAutoFilter
    
End Property

Public Property Let AutoFilter(ByVal lFilter As Boolean)
    
    '---------------------------------------------------------------------------------------
    ' Procedure : AutoFilter
    ' DateTime  : 15/06/2007 11:41
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------
    
    m_lAutoFilter = lFilter
    
    UserControl.PropertyChanged "AutoFilter"

    If (Ambient.UserMode = False) Then
        Refresh False
    End If
        
End Property

Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : BackColor
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lBackColor = ConvertSystemColor(lBackColor)
    
    With McCalendar
        .CalendarGradientCol = m_lBackColor
        .HeaderGradientCol = m_lBackColor
        .MonthGradientCol = m_lBackColor
        .YearGradientCol = m_lBackColor
        .WeekDaySelCol = m_lBackColor
        .WeekDaySunCol = m_lBackColor
        .DaySunCol = ShiftColorOXP(m_lBackColor, &H20)
        .DayCol = ShiftColorOXP(m_lBackColor)
    End With
    
    With lstFX1
        .BackSelected = m_lBackColor
        .BackSelectedG1 = m_lBackColor
        .BackSelectedG2 = m_lBackColor
    End With

    UserControl.PropertyChanged "BackColor"

    If (Ambient.UserMode = False) Then
        UserControl_Resize
    End If

End Property

Public Property Get BackColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : BackColor
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    BackColor = m_lBackColor

End Property

Public Sub CategoryChanged(ByVal Key As String, ByVal NewKey As String, _
                           ByVal Title As String, _
                           Optional ByVal ToolTipText As String = vbNullString, _
                           Optional ByVal Expand As Boolean = False)

  Dim lKeyIndex As Long

    If (FindCategory(Key) = True) And (LenB(Trim$(NewKey)) > 0) Then
        lKeyIndex = GetCategoryIndex(Key)
        
        If (NewKey <> Key) Then
            If (FindCategory(NewKey) = True) Then
                Err.Raise 512, "CategoryChanged", "The key already exist or is empty."
                Exit Sub
            End If
            
        End If
        
        Category(lKeyIndex).Expand = Expand
        Category(lKeyIndex).Key = NewKey
        Category(lKeyIndex).Title = Title
        Category(lKeyIndex).ToolTipText = ToolTipText
    Else
        Err.Raise 512, "CategoryChanged", "The key already exist or is empty."
    End If
    
End Sub

Public Sub CategoryTitleChanged(ByVal Key As String, ByVal Title As String, _
                           Optional ByVal ToolTipText As String = vbNullString)

  Dim lKeyIndex As Long

    If (FindCategory(Key) = True) And (LenB(Trim$(Title)) > 0) Then
        lKeyIndex = GetCategoryIndex(Key)
        Category(lKeyIndex).Title = Title
        Category(lKeyIndex).ToolTipText = ToolTipText
    Else
        Err.Raise 515, "CategoryTitleChanged", "The key not exist or title is empty."
    End If
    
End Sub

Public Sub Clear()
    
    Erase Category
    Erase ChildItem
    Erase ChildOrd
    CountCatg = 0
    CountChild = 0
    lItemSelected = -1
    lButton = -1
    Refresh True
    
End Sub

Private Function CloneFont(ByVal Font As StdFont) As StdFont

On Error Resume Next
    Set CloneFont = New StdFont
    CloneFont.Name = Font.Name
    CloneFont.Size = Font.Size
    CloneFont.Bold = Font.Bold
    CloneFont.Italic = Font.Italic
    CloneFont.Underline = Font.Underline
    CloneFont.Strikethrough = Font.Strikethrough
On Error GoTo 0
    
End Function

Private Function ConvertSystemColor(ByVal theColor As Long) As Long

    OleTranslateColor theColor, 0, ConvertSystemColor

End Function

Public Sub ChildItemChanged(ByVal Index As Long, ByVal LastKeyCategory, ByVal NewKeyCategory As String, ByVal TypeGridItem As PropertyItemType, ByVal _
                        Title As String, ByVal Value As Variant, Optional ByVal ItemValue As _
                        String = vbNullString, Optional ByVal ToolTipText As String = vbNullString, _
                        Optional ByVal StyleText As StringStyle = vbNormal, Optional ByVal sFilters As String = vbNullString)

    If (FindCategory(LastKeyCategory) = True) And (FindCategory(NewKeyCategory) = True) And (FindChild(ChildItem(Index).Title, LastKeyCategory) = True) Then
        If (ChildItem(Index).Title <> Title) Then
            If (FindChild(Title, NewKeyCategory) = True) Or (LenB(Title) = 0) Then
                Err.Raise 513, "ChildItemChanged", "The KeyCategory doesn't exists or the key Title already exist or is empty."
                Exit Sub
            End If
            
        End If
        
        ChildItem(Index).ItemValue = ItemValue
        ChildItem(Index).KeyCategory = NewKeyCategory
        ChildItem(Index).Title = Title
        ChildItem(Index).ToolTipText = ToolTipText
        ChildItem(Index).TypeGrid = TypeGridItem
        ChildItem(Index).Filters = sFilters

        If (TypeGridItem = PropertyItemFont) Or (TypeGridItem = PropertyItemPicture) Then
            Set ChildItem(Index).Value = Value
        Else
            ChildItem(Index).Value = Value
        End If
        
        If (TypeGridItem = PropertyItemNumber) Then
            ChildItem(Index).StyleString = ItemNumeric
        Else
            ChildItem(Index).StyleString = StyleText
        End If

        Let ChildOrd() = ChildItem()

        If (m_vPropertySort = Alphabetical) Then
            Insertion
        End If

    Else
        Err.Raise 513, "ChildItemChanged", "The KeyCategory doesn't exists or the key Title already exist or is empty."
    End If
    
End Sub

Public Property Get DecimalSymbol() As String

    DecimalSymbol = m_sDecimalSymbol

End Property

Public Property Let DecimalSymbol(ByVal sDecimalSymbol As String)

    m_sDecimalSymbol = sDecimalSymbol

    UserControl.PropertyChanged "DecimalSymbol"

End Property

Public Sub DelAllChildToCatg(ByVal KeyCategory As String)
    
    RemoveItem KeyCategory, vbNullString, True
    
End Sub

Private Sub DrawCaption(ByVal lCaption As String, _
                        ByRef m_btnRect As RECT, _
                        Optional ByVal lColor As OLE_COLOR = &HF0, _
                        Optional ByVal lLeft As Integer = 18, _
                        Optional ByVal lRight As Integer = 0, _
                        Optional ByVal WordWrap As Boolean = False)

  Dim lAlign      As Long
  
    SetTextColor UserControl.hDC, lColor
    m_btnRect.rBottom = UserControl.ScaleHeight
    m_btnRect.rLeft = lLeft

    If (lRight = 0) Then
        m_btnRect.rRight = UserControl.ScaleWidth
    ElseIf (lRight = -1) And (bVScroll = False) Then
        m_btnRect.rRight = UserControl.ScaleWidth - 5
    ElseIf (lRight = -1) And (bVScroll = True) Then
        m_btnRect.rRight = UserControl.ScaleWidth - 19
    Else
        m_btnRect.rRight = lRight
    End If

    If (WordWrap = True) Then
        lAlign = DT_WORDBREAK
    Else
        lAlign = DT_SINGLELINE
    End If
    
    '*************************************************************************
    '* Draws the text with Unicode support based on OS version.              *
    '* Thanks to Richard Mewett.                                             *
    '*************************************************************************

    If (mWindowsNT = True) Then
        DrawTextW UserControl.hDC, StrPtr(lCaption), Len(lCaption), m_btnRect, DT_WORD_ELLIPSIS Or lAlign Or DT_NOPREFIX
    Else
        DrawTextA UserControl.hDC, lCaption, Len(lCaption), m_btnRect, DT_WORD_ELLIPSIS Or lAlign Or DT_NOPREFIX
    End If

End Sub

Private Sub DrawRectangleBorder(ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal Color As Long, _
                                Optional ByVal SetBorder As Boolean = True)

  Dim hBrush   As Long
  Dim TempRect As RECT

    TempRect.rLeft = X
    TempRect.rTop = Y
    TempRect.rRight = X + Width
    TempRect.rBottom = Y + Height
    hBrush = CreateSolidBrush(Color)

    If (SetBorder = True) Then
        FrameRect hDC, TempRect, hBrush
    Else
        FillRect hDC, TempRect, hBrush
    End If

    DeleteObject hBrush

End Sub

Private Sub DrawScrollBar()

  Dim lMajor As Long
  Dim lMinor As Long

    isXp = False
    GetWindowsVersion lMajor, lMinor

    If (lMajor > 5) Then
        isXp = True
    ElseIf (lMajor = 5) And (lMinor >= 1) Then
        isXp = True
    End If

    lStyle = WS_CHILD Or WS_VISIBLE

    If Not (m_hWnd) Then
        lStyle = lStyle Or SB_VERT And Not SBS_HORZ
        m_hWnd = CreateWindowEx(0, "SCROLLBAR", vbNullString, lStyle, UserControl.ScaleWidth - 19, 2, 17, _
            UserControl.ScaleHeight - 22, hWnd, 0, App.hInstance, ByVal 0&)

        If (isXp = True) Then
            ShowScrollBar m_hWnd, SB_CTL, 0
        Else
            FlatSB_ShowScrollBar m_hWnd, SB_CTL, False
        End If

    End If

End Sub

Private Sub DrawVGradient(ByVal lEndColor As Long, _
                            ByVal lStartcolor As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long)

    ''Draw a Vertical Gradient in the current hDC
  Dim dR As Single
  Dim dG As Single
  Dim dB As Single
  Dim sR As Single
  Dim sG As Single
  Dim sB As Single
  Dim eR As Single
  Dim eG As Single
  Dim eB As Single
  Dim ni As Long

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
        APILine X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni

End Sub

'---------------------------------------------------------------------------------------
' Function  : DrawTheme
' DateTime  : 03/08/05 13:38
' Author    : HACKPRO TM
' Purpose   : Try to open Uxtheme.dll.
'---------------------------------------------------------------------------------------
Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           rtRect As RECT, _
                           Optional ByVal CloseTheme As Boolean = False) As Boolean
  
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
    If (CloseTheme = True) Then Call CloseThemeData(hTheme)
    '* Exit the function now.
    Exit Function
    
NoXP:
    '* An Error was detected, drawing Failed.
    DrawTheme = False
On Error GoTo 0

End Function

Private Function DrawXPTheme(ByVal m_sClass As String, _
                             ByRef tW As RECT, _
                             ByVal iStateId As Long, _
                             ByVal iPartId As Long, _
                             ByVal lFocus As Boolean) As Boolean

  Dim hTheme As Long
  Dim lR As Long
  Dim bSuccess As Boolean
  Dim lDefaultColor As OLE_COLOR

On Error GoTo DrawXPThemeError
    If (m_StylePropertyGrid = &H3) Then
        tW.rTop = tW.rTop - 6
        
        lDefaultColor = vbBlack
        
        If (lFocus = True) Then lDefaultColor = vbWhite
        
        If (iStateId = 1) Then
            UserControl.Line (tW.rLeft + 9, tW.rTop + 8)-(tW.rLeft + 13, tW.rTop + 12), lDefaultColor
            UserControl.Line (tW.rLeft + 10, tW.rTop + 8)-(tW.rLeft + 13, tW.rTop + 11), lDefaultColor
            UserControl.Line (tW.rLeft + 15, tW.rTop + 8)-(tW.rLeft + 11, tW.rTop + 12), lDefaultColor
            UserControl.Line (tW.rLeft + 14, tW.rTop + 8)-(tW.rLeft + 11, tW.rTop + 11), lDefaultColor
            UserControl.Line (tW.rLeft + 9, tW.rTop + 12)-(tW.rLeft + 13, tW.rTop + 16), lDefaultColor
            UserControl.Line (tW.rLeft + 10, tW.rTop + 12)-(tW.rLeft + 13, tW.rTop + 15), lDefaultColor
            UserControl.Line (tW.rLeft + 15, tW.rTop + 12)-(tW.rLeft + 11, tW.rTop + 16), lDefaultColor
            UserControl.Line (tW.rLeft + 14, tW.rTop + 12)-(tW.rLeft + 11, tW.rTop + 15), lDefaultColor
        Else
            UserControl.Line (tW.rLeft + 9, tW.rTop + 11)-(tW.rLeft + 13, tW.rTop + 7), lDefaultColor
            UserControl.Line (tW.rLeft + 10, tW.rTop + 11)-(tW.rLeft + 13, tW.rTop + 8), lDefaultColor
            UserControl.Line (tW.rLeft + 15, tW.rTop + 11)-(tW.rLeft + 11, tW.rTop + 7), lDefaultColor
            UserControl.Line (tW.rLeft + 14, tW.rTop + 11)-(tW.rLeft + 11, tW.rTop + 8), lDefaultColor
            UserControl.Line (tW.rLeft + 9, tW.rTop + 15)-(tW.rLeft + 13, tW.rTop + 11), lDefaultColor
            UserControl.Line (tW.rLeft + 10, tW.rTop + 15)-(tW.rLeft + 13, tW.rTop + 12), lDefaultColor
            UserControl.Line (tW.rLeft + 15, tW.rTop + 15)-(tW.rLeft + 11, tW.rTop + 11), lDefaultColor
            UserControl.Line (tW.rLeft + 14, tW.rTop + 15)-(tW.rLeft + 11, tW.rTop + 12), lDefaultColor
        End If
        
        Exit Function
    End If

    bSuccess = True
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(m_sClass))

    If (hTheme) Then
        lR = DrawThemeParentBackground(UserControl.hWnd, UserControl.hDC, tW)
        lR = DrawThemeBackground(hTheme, UserControl.hDC, iPartId, iStateId, tW, tW)

        If (lR <> 0) Then
            bSuccess = False
        End If

    Else
        bSuccess = False
    End If

    DrawXPTheme = bSuccess
    CloseThemeData hTheme
    Exit Function
    
DrawXPThemeError:
   
   bSuccess = False

End Function

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled = New_Enabled
    m_bEnabled = New_Enabled
    PropertyChanged "Enabled"

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_bEnabled

End Property

Public Sub ExpandCategory(ByVal Index As Long, ByVal bExpand As Boolean)

On Error GoTo myErr
    Category(Index).Expand = bExpand
    Exit Sub
    
myErr:
    
End Sub

Public Function ExtractPath(ByVal sFileName) As String

    '------------------------------------------------------------------>
    ' By Paul R. Territo, Ph.D

    '   Extract the Path from the full filename...
  Dim lStrCnt As Long

    lStrCnt = InStrRev(sFileName, "\")

    If lStrCnt > 0 Then
        ExtractPath = Mid$(sFileName, 1, lStrCnt - 1)
    End If

End Function

Public Sub FilterItemChanged(ByVal Title As String, ByVal Filters As String, _
    Optional ByVal AutoRedraw As Boolean = False)
  
  Dim i As Long
  Dim lPos As Long

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            lPos = i
            Exit For
        End If

    Next i
    
    ChildItem(lPos).Filters = Filters
    
    For i = 0 To CountChild - 1

        If (Title = ChildOrd(i).Title) Then
            lPos = i
            Exit For
        End If

    Next i
    
    ChildOrd(lPos).Filters = Filters
    
    If (AutoRedraw = True) Then
        ReDraw False
    End If
    
End Sub

Public Function FindCategory(ByVal KeyCategory As String) As Boolean

  Dim i As Long


On Error GoTo FindCategoryErr
    FindCategory = False

    For i = 0 To CountCatg - 1

        If (KeyCategory = Category(i).Key) Then
            FindCategory = True
            Exit For
        End If

    Next i
    
    Exit Function
    
FindCategoryErr:
    
    FindCategory = False

End Function

Public Function FindChild(ByVal Title As String, ByVal KeyCategory As String) As Boolean

  Dim i As Long

On Error GoTo FindChildErr
    FindChild = False

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            If (KeyCategory = ChildItem(i).KeyCategory) Then
                FindChild = True
                Exit For
            End If
            
        End If

    Next i
    
    Exit Function
    
FindChildErr:
    
    FindChild = False

End Function

Public Function FindChildKey(ByVal Title As String, ByVal KeyName As String) As Boolean

  Dim i As Long

On Error GoTo FindChildKeyErr
    FindChildKey = False
    
    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            If (KeyName = ChildItem(i).KeyName) Then
                FindChildKey = True
                Exit For
            End If
            
        End If

    Next i
    
    Exit Function
    
FindChildKeyErr:
    
    FindChildKey = False

End Function

Public Function FindInList(ByVal KeyCategory As String, ByVal Title As String, ByVal Item As String) As Long

  Dim DimChild As Long
  Dim mUbound As Long
  Dim i As Long

    FindInList = -1
    If (FindCategory(KeyCategory) = True) And (FindChild(Title, KeyCategory) = True) Then
        
        DimChild = GetChildIndex(Title, KeyCategory)
        If (IsArray(ChildOrd(DimChild).Value) = True) Then
             mUbound = UBound(ChildOrd(DimChild).Value)
             For i = 0 To mUbound
                 If (Item = ChildOrd(DimChild).Value(i)) Then
                     FindInList = i
                     Exit For
                 End If
                 
             Next i
             
        End If
        
    End If
    
End Function

Public Function FindItem(ByVal Title As String) As String

  Dim i As Long

On Error GoTo FindItemErr

    FindItem = vbNullString

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            FindItem = ChildItem(i).KeyCategory
            Exit For
        End If

    Next i
    
    Exit Function
    
FindItemErr:
    
    FindItem = vbNullString

End Function

Public Property Let FixedSplit(ByVal New_FixedSplit As Boolean)
    
    m_bFixedSplit = New_FixedSplit
    
    UserControl.PropertyChanged "FixedSplit"
    
End Property

Public Property Get FixedSplit() As Boolean

   '---------------------------------------------------------------------------------------
   ' Procedure : FixedSplit
   ' DateTime  : 20/11/2006 21:03
   ' Author    : HACKPRO TM
   ' Purpose   :
   '---------------------------------------------------------------------------------------

   FixedSplit = m_bFixedSplit

End Property

Public Property Get Font() As StdFont

    '---------------------------------------------------------------------------------------
    ' Procedure : Font
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    Set Font = m_lFont

End Property

Public Property Set Font(ByVal New_Font As StdFont)

    With m_lFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
        
    UserControl.PropertyChanged "Font"
    
End Property

Public Property Get GetCategoryChildCount(ByVal KeyCategory As String) As Long
  
  Dim i As Long
  Dim cChild As Long
  
    cChild = 0
    
    For i = 0 To CountChild - 1
        If (ChildOrd(i).KeyCategory = KeyCategory) Then
            cChild = cChild + 1
        End If
        
    Next i
    
    GetCategoryChildCount = cChild
    
End Property

Public Property Get GetCategoryIndex(ByVal KeyCategory As String) As Long

On Error GoTo GetCategoryIndexErr
  
  Dim i As Long
  
    GetCategoryIndex = -1

    For i = 0 To CountCatg - 1

        If (KeyCategory = Category(i).Key) Then
            GetCategoryIndex = i
            Exit For
        End If

    Next i
    
    Exit Property

GetCategoryIndexErr:
  
    GetCategoryIndex = -1

End Property

Public Property Get GetCategoryKey(ByVal Index As Long) As String
    
On Error GoTo GetCategoryKeyErr
    GetCategoryKey = Category(Index).Key
    Exit Property
    
GetCategoryKeyErr:

    GetCategoryKey = vbNullString
    
End Property

Public Property Get GetChildCategory(ByVal Index As Long) As String
    
On Error GoTo GetChildCategoryErr
    GetChildCategory = ChildOrd(Index).KeyCategory
    Exit Property
    
GetChildCategoryErr:

    GetChildCategory = vbNullString
    
End Property

Public Property Get GetChildIndex(ByVal Title As String, ByVal KeyCategory As String) As Long
    
On Error GoTo GetChildIndexErr
  Dim i As Long
  
    GetChildIndex = -1

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            If (KeyCategory = ChildItem(i).KeyCategory) Then
                GetChildIndex = i
                Exit For
            End If
            
        End If

    Next i
    
    Exit Property
    
GetChildIndexErr:

    GetChildIndex = -1
    
End Property

Public Property Get GetChildKeyName(ByVal Title As String, ByVal KeyCategory As String) As String
    
On Error GoTo GetChildKeyNameErr
  Dim i As Long
  
    GetChildKeyName = vbNullString

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            If (KeyCategory = ChildItem(i).KeyCategory) Then
                GetChildKeyName = ChildItem(i).KeyName
                Exit For
            End If
            
        End If

    Next i
    
    Exit Property
    
GetChildKeyNameErr:

    GetChildKeyName = vbNullString
    
End Property

Public Property Get GetChildTitle(ByVal Index As Long) As String
    
On Error GoTo GetChildTitleErr
    GetChildTitle = ChildOrd(Index).Title
    Exit Property
    
GetChildTitleErr:

    GetChildTitle = vbNullString
    
End Property

Public Property Get GetChildType(ByVal Index As Long) As PropertyItemType
    
On Error GoTo GetChildTypeErr
    GetChildType = ChildOrd(Index).TypeGrid
    Exit Property
    
GetChildTypeErr:
    
    GetChildType = -1
    
End Property

Public Property Get GetChildValue(ByVal KeyCategory As String, ByVal Title As String) As Variant
    
  Dim i As Long
  Dim lPos As Long
  Dim lOk As Boolean

On Error Resume Next
    For i = 0 To CountChild - 1

        If (Title = ChildOrd(i).Title) And (KeyCategory = ChildOrd(i).KeyCategory) Then
            lPos = i
            lOk = True
            Exit For
        End If

    Next i
    
    If (lOk = True) Then
        If (ChildOrd(i).TypeGrid = PropertyItemFont) Or (ChildOrd(i).TypeGrid = PropertyItemPicture) Then
            Set GetChildValue = ChildOrd(i).Value
        ElseIf (ChildOrd(i).TypeGrid = PropertyItemStringList) Then
            GetChildValue = ChildOrd(i).ItemValue
        Else
            GetChildValue = ChildOrd(i).Value
        End If
        
    Else
        If (ChildOrd(i).TypeGrid = PropertyItemFont) Or (ChildOrd(i).TypeGrid = PropertyItemPicture) Then
            Set GetChildValue = Nothing
        Else
            GetChildValue = vbNullString
        End If
        
    End If
On Error GoTo 0
    
End Property

Public Property Get GetChildValueFromIndex(ByVal Index As Long) As Variant
    
  Dim i As Long
  Dim lPos As Long
  Dim lOk As Boolean

On Error GoTo GetChildValueFromIndexError
    If (ChildOrd(Index).TypeGrid = PropertyItemFont) Or _
        (ChildOrd(Index).TypeGrid = PropertyItemPicture) _
    Then
        Set GetChildValueFromIndex = ChildOrd(Index).Value
    ElseIf (ChildOrd(Index).TypeGrid = PropertyItemStringList) Then
        GetChildValueFromIndex = ChildOrd(Index).ItemValue
    Else
        GetChildValueFromIndex = ChildOrd(Index).Value
    End If
    
    Exit Property
    
GetChildValueFromIndexError:
    
    GetChildValueFromIndex = vbNullString
    
End Property

Public Property Get GetCountCategory() As Long
    
    GetCountCategory = CountCatg - 1
    
End Property

Public Property Get GetCountChild() As Long
    
    GetCountChild = CountChild - 1
    
End Property

Private Function getResourceCursor(ByVal ResCursor As String) As IPictureDisp

    Set getResourceCursor = LoadResPicture(ResCursor, vbResCursor)

End Function

Public Property Get GetControlVersion() As String
    
    
    GetControlVersion = thisVersion
    
End Property

Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, _
                              Optional ByRef lMinor = 0, _
                              Optional ByRef lRevision = 0, _
                              Optional ByRef lBuildNumber = 0)

    '---------------------------------------------------------------------------------------
    ' Procedure : GetWindowsVersion
    ' DateTime  : 09/07/05 18:00
    ' Author    : HACKPRO TM
    ' Purpose   : OS Version.
    '---------------------------------------------------------------------------------------

  Dim lR As Long

    lR = GetVersion()
    lBuildNumber = (lR And &H7F000000) \ &H1000000
    If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
    lRevision = (lR And &HFF0000) \ &H10000
    lMinor = (lR And &HFF00&) \ &H100
    lMajor = (lR And &HFF)

End Sub

Public Property Get HelpBackColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpBackColor
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    HelpBackColor = m_lHelpBackColor

End Property

Public Property Let HelpBackColor(ByVal lHelpBackColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpBackColor
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lHelpBackColor = ConvertSystemColor(lHelpBackColor)
    
    McCalendar.DaySelCol = m_lHelpBackColor

    UserControl.PropertyChanged "HelpBackColor"

    If (Ambient.UserMode = False) Then
        UserControl_Resize
    End If

End Property

Public Property Let HelpForeColor(ByVal lHelpForeColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpForeColor
    ' DateTime  : 25/11/2006 11:30
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lHelpForeColor = lHelpForeColor

    UserControl.PropertyChanged "HelpForeColor"

End Property

Public Property Get HelpForeColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpForeColor
    ' DateTime  : 25/11/2006 11:30
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    HelpForeColor = m_lHelpForeColor

End Property

Public Property Let HelpHeight(ByVal iHelpHeight As Integer)

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpHeight
    ' DateTime  : 20/11/2006 20:39
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_iHelpHeight = iHelpHeight
    
    xSplitterY = m_iHelpHeight

    UserControl.PropertyChanged "HelpHeight"

    If (Ambient.UserMode = False) Then
        UserControl_Resize
        ReDraw
    End If

End Property

Public Property Get HelpHeight() As Integer

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpHeight
    ' DateTime  : 20/11/2006 20:39
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    HelpHeight = m_iHelpHeight

End Property

Public Property Let HelpVisible(ByVal bHelpVisible As Boolean)

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpVisible
    ' DateTime  : 20/11/2006 21:11
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_bHelpVisible = bHelpVisible

    UserControl.PropertyChanged "HelpVisible"

    If (Ambient.UserMode = False) Then
        UserControl_Resize
        ReDraw
    End If

End Property

Public Property Get HelpVisible() As Boolean

    '---------------------------------------------------------------------------------------
    ' Procedure : HelpVisible
    ' DateTime  : 20/11/2006 21:11
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    HelpVisible = m_bHelpVisible

End Property

Public Function HexToRGB(ByVal HexColor As String) As String
 
    R = CByte("&H" & Mid(HexColor, 2, 2))
    G = CByte("&H" & Mid(HexColor, 4, 2))
    B = CByte("&H" & Mid(HexColor, 6, 2))
    HexToRGB = R & "," & G & "," & B
    
End Function

Public Property Get hWnd() As Long

    hWnd = UserControl.hWnd

End Property

Private Sub InitCustomColors()

  Dim i As Long

    '   Init the Custom Colors Array to White

    For i = LBound(CustomColors) To UBound(CustomColors)
        ' Sets all custom colors to white
        CustomColors(i) = 254
    Next i

    '   Convert array to Unicode Strings
    ColorDialog.lpCustColors = StrConv(CustomColors, vbUnicode)

End Sub

Private Sub Insertion()

  Dim Index    As Long
  Dim Temp    As Variant
  Dim NextV As Long
  Dim LIndex   As Long
  Dim UIndex As Long

  Dim Matrix() As TypeChildItem

    Let Matrix() = ChildOrd()

    LIndex = LBound(Matrix)
    UIndex = UBound(Matrix)

    NextV = LIndex + 1

    While (NextV <= UIndex)
        Index = NextV

        Do

            If (Index > LIndex) Then
                If (Matrix(Index).Title < Matrix(Index - 1).Title) Then
                    Temp = Matrix(Index).Filters
                    Matrix(Index).Filters = Matrix(Index - 1).Filters
                    Matrix(Index - 1).Filters = Temp
                    
                    Temp = Matrix(Index).ItemValue
                    Matrix(Index).ItemValue = Matrix(Index - 1).ItemValue
                    Matrix(Index - 1).ItemValue = Temp

                    Temp = Matrix(Index).KeyCategory
                    Matrix(Index).KeyCategory = Matrix(Index - 1).KeyCategory
                    Matrix(Index - 1).KeyCategory = Temp

                    Temp = Matrix(Index).Title
                    Matrix(Index).Title = Matrix(Index - 1).Title
                    Matrix(Index - 1).Title = Temp

                    Temp = Matrix(Index).ToolTipText
                    Matrix(Index).ToolTipText = Matrix(Index - 1).ToolTipText
                    Matrix(Index - 1).ToolTipText = Temp

                    Temp = Matrix(Index).TypeGrid
                    Matrix(Index).TypeGrid = Matrix(Index - 1).TypeGrid
                    Matrix(Index - 1).TypeGrid = Temp

                    Temp = Matrix(Index).Value
                    Matrix(Index).Value = Matrix(Index - 1).Value
                    Matrix(Index - 1).Value = Temp

                    Index = Index - 1
                Else
                    Exit Do
                End If

            Else
                Exit Do
            End If

        Loop
        NextV = NextV + 1
    Wend
    
    Erase ChildOrd
    Let ChildOrd() = Matrix()
    Erase Matrix

End Sub

Private Sub isBttAction_Click()

  Dim m_Value As Variant
  Dim psColor As SelectedColor
  Dim psFont  As SelectedFont
  Dim psFile  As SelectedFile

    ' Based in the UC ucPickBox developed by Paul R. Territo, Ph.D.
    ' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63905&lngWId=1

On Error GoTo ClickError

    bResize = True
    If (lChild >= 0) Then
        If Not (ChildOrd(lChild).TypeGrid = PropertyItemPicture) Then
            m_Value = ChildOrd(lChild).Value
        Else
        On Error Resume Next
            Set m_Value = ChildOrd(lChild).Value
        On Error GoTo 0
        End If

        Select Case ChildOrd(lChild).TypeGrid
        Case &H1 ' Color Dialog.

            '   Pick a color from the Color Dialog
            psColor = ShowColor()

            If psColor.bCanceled = False Then
                '   Get the color from the dialog
                m_Value = (CLng(psColor.oSelectedColor))
            End If

        Case &H3 ' Folder Dialog.
            Dim lWidth As Long

            m_Value = ShowFolderBrowse()
            lWidth = UserControl.ScaleWidth - 10

            If (LenB(m_Value) > 0) Then
                m_Path = QualifyPath(m_Value)
                If (m_bTrimPath = True) Then
                    m_Value = TrimPathByLen(m_Path, lWidth)
                End If
                
            Else
                m_Value = vbNullString
            End If

        Case &H6 ' Form.
            Dim p As POINTAPI
            
            ' Get the position of the cursor
            GetCursorPos p
            RaiseEvent FormClick(ChildOrd(lChild).Value, ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, p.X, p.Y)
            m_Value = ChildOrd(lChild).Value
            
        Case &H5 ' Font Dialog.

            If (m_Font Is Nothing) Then
                '   Create a Font if Missing
                Set m_Font = New StdFont
                With m_Font
                    .Bold = ChildOrd(lChild).theFont.Bold
                    .Charset = ChildOrd(lChild).theFont.Charset
                    .Italic = ChildOrd(lChild).theFont.Italic
                    .Name = ChildOrd(lChild).theFont.Name
                    .Size = ChildOrd(lChild).theFont.Size
                    .Strikethrough = ChildOrd(lChild).theFont.Strikethrough
                    .Underline = ChildOrd(lChild).theFont.Underline
                    .Weight = ChildOrd(lChild).theFont.Weight
                End With
                
            End If

            psFont = ShowFont(m_Font, m_FontColor)

            If (psFont.bCanceled = False) Then
                '   Set the Font type
                Set m_Font = New StdFont

                With m_Font
                    .Bold = psFont.bBold
                    .Italic = psFont.bItalic
                    .Name = psFont.sSelectedFont
                    .Size = psFont.nSize
                    .Strikethrough = psFont.bStrikeOut
                    .Underline = psFont.bUnderline
                End With

                m_Value = psFont.sSelectedFont
            End If

        Case &H4, &H8 ' Picture Dialog or FolderFile.
            If (ChildOrd(lChild).TypeGrid = PropertyItemFolderFile) Then
                m_Filters = ChildOrd(lChild).Filters
                FileDialog.sFile = ChildOrd(lChild).Value
            End If
            
            psFile = ShowOpen(m_Filters)

            If (psFile.bCanceled = False) And (psFile.nFilesSelected > 0) Then
                '  Concatinate the filename and path
                '  Store the qaulified path
                m_Path = QualifyPath(ExtractPath(psFile.sFiles(1)))
                m_Value = psFile.sFiles(1)
            End If

        End Select
        
        bResize = False
        ChildOrd(lChild).Value = m_Value
        bChanged = True
        
        If (ChildOrd(lChild).TypeGrid = PropertyItemFont) Then
            RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                ChildOrd(lChild).Value, m_Font)
            Set ChildOrd(lChild).theFont = m_Font
        Else
            RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                ChildOrd(lChild).Value, Nothing)
        End If
        
        ReDraw
        '   Pass the focus back the Host
        SetFocus UserControl.Parent.hWnd
    End If
    
    Exit Sub
    
ClickError:

End Sub

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean

    ' Subclassing for Paul Caton --> The Man (^~^)#
    ' Determine if the passed function is supported

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

Private Property Let LargeChange(ByVal iLargeChange As Long)

  Dim tSI As SCROLLINFO

    pGetSI tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pLetSI tSI, SIF_PAGE Or SIF_RANGE

End Property

Private Property Get LargeChange() As Long

  Dim tSI As SCROLLINFO

    pGetSI tSI, SIF_PAGE
    LargeChange = tSI.nPage

End Property

Public Property Get LineColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : LineColor
    ' DateTime  : 20/11/2006 20:45
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    LineColor = m_lLineColor

End Property

Public Property Let LineColor(ByVal lLineColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : LineColor
    ' DateTime  : 20/11/2006 20:45
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lLineColor = ConvertSystemColor(lLineColor)
    
    lstFX1.SelectBorderColor = m_lLineColor
    lstFX1.SelectListBorderColor = m_lLineColor
    McCalendar.BorderColor = m_lLineColor

    UserControl.PropertyChanged "LineColor"

    If (Ambient.UserMode = False) Then
        UserControl_Resize
    End If

End Property

Public Property Get ListCount() As Long
    
    ListCount = lstFX1.ListCount
    
End Property

Public Property Get ListIndex() As Long
    
    ListIndex = m_ListIndex
    
End Property


Public Function LongToHex(ByVal hColor As Long) As String

  Dim Red As Long
  Dim Green As Long
  Dim Blue As Long
  Dim sRed As String
  Dim sBlue As String
  Dim sGreen As String

    ' Separate the colours into their own variables
    Red = hColor And 255
    Green = (hColor And 65280) \ 256
    Blue = (hColor And 16711680) \ 65535
    
    ' Get the hex equivalents
    sRed = Hex(Red)
    sBlue = Hex(Blue)
    sGreen = Hex(Green)
    
    ' Pad each colour, to make sure it's 2 chars
    sRed = String(2 - Len(sRed), "0") & sRed
    sBlue = String(2 - Len(sBlue), "0") & sBlue
    sGreen = String(2 - Len(sGreen), "0") & sGreen

    LongToHex = sRed & sGreen & sBlue
    
End Function

Public Function LongToRGB(ByVal lColor As Long) As String

    RGBColor.Red = lColor And &HFF
    RGBColor.Green = (lColor \ &H100) And &HFF
    RGBColor.Blue = (lColor \ &H10000) And &HFF

    LongToRGB = RGBColor.Red & ", " & RGBColor.Green & ", " & RGBColor.Blue

End Function


Private Sub lstFX1_Click()

On Error GoTo ClickError

    With lstFX1
        
        If (ChildOrd(lChild).Filters <> "MS") Then
            .Visible = False
        End If
        
        If (lChild >= 0) And (ChildOrd(lChild).Filters <> "MS") Then

            m_ListIndex = .ListIndex
            Select Case ChildOrd(lChild).TypeGrid
            Case &H0 ' True/False.
                ChildOrd(lChild).Value = .ItemText(.ListIndex)
                RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                    ChildOrd(lChild).Value, Nothing)

            Case &HA ' StringList.
                ChildOrd(lChild).ItemValue = .ItemText(.ListIndex)
                RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                    ChildOrd(lChild).ItemValue, Nothing)
                
            End Select

            bChanged = True
            ReDraw
        
        ElseIf (ChildOrd(lChild).TypeGrid = &HA) Then
          Dim i As Long
                        
            ChildOrd(lChild).ItemValue = vbNullString
            For i = 0 To .ListCount - 1
                If (.ItemSelected(i) = True) Then
                    ChildOrd(lChild).ItemValue = .ItemText(i) & ", " & ChildOrd(lChild).ItemValue
                End If
                
            Next i
            
            SCmb.Text = ChildOrd(lChild).ItemValue
            
        End If

    End With
    
    If (ChildOrd(lChild).Filters <> "MS") Then
        SCmb.ClosedList
    End If
    
    Exit Sub
    
ClickError:

End Sub

Public Function MakeWebColor(ByVal hColor As Long) As String
 
    MakeWebColor = "#" & LongToHex(hColor)
    
End Function

Private Property Let Max(ByVal iMax As Long)

  Dim tSI As SCROLLINFO

    If (bResize = False) Then
        tSI.nMax = iMax + LargeChange
        tSI.nMin = Min
        pLetSI tSI, SIF_RANGE
        pRaiseEvent False
    End If

End Property

Private Property Get Max() As Long

  Dim tSI As SCROLLINFO

    If (bResize = False) Then
        pGetSI tSI, SIF_RANGE Or SIF_PAGE
        Max = tSI.nMax - tSI.nPage
    End If

End Property

Private Sub McCalendar_DateChanged()
On Error GoTo ClickError

    With McCalendar
        .Visible = False
        If (lChild >= 0) Then

            ChildOrd(lChild).Value = .Value
            RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                    ChildOrd(lChild).Value, Nothing)
            bChanged = True
            ReDraw
        End If

    End With
    
    SCmb.ClosedList
    
    Exit Sub
    
ClickError:
End Sub

Private Property Let Min(ByVal iMin As Long)

  Dim tSI As SCROLLINFO

    tSI.nMin = iMin
    tSI.nMax = Max + LargeChange
    pLetSI tSI, SIF_RANGE

End Property

Private Property Get Min() As Long

  Dim tSI As SCROLLINFO

    pGetSI tSI, SIF_RANGE
    Min = tSI.nMin

End Property

Public Sub OfficeAppearance(ByVal lOfficeAppearance As ComboOfficeAppearance)

    SCmb.OfficeAppearance = lOfficeAppearance

End Sub

Private Sub pGetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)

    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars = True) Then
        GetScrollInfo m_hWnd, SB_CTL, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, SB_CTL, tSI
    End If

End Sub

Private Sub pLetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)

    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars = True) Then
        SetScrollInfo m_hWnd, SB_CTL, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, SB_CTL, tSI, True
    End If

End Sub

Private Function pRaiseEvent(ByVal bScroll As Boolean)

  Static s_lLastValue As Long

    If (Value <> s_lLastValue) Then
        If (bScroll = True) Then
            RaiseEvent Scroll
        Else
            RaiseEvent Change
        End If

        s_lLastValue = Value
    End If

End Function

Private Function ProcessFilter(sFilter As String) As String

  Dim i As Long

    '   This routine replaces the Pipe (|) character for filter
    '   strings and pads the size to the required legnth.
    '
    '   Example:
    '   - Input (String)
    '       "Supported files|*.bmp;*.doc;*.jpg;*.rtf;*.txt;*.tif|Bitmap files (*.bmp)|*.bmp|Word
    '   files (*.doc)|*.doc|JPEG files (*.jpg)|*.jpg|RichText files (*.rtf)|*.rtf|Text files
    '   (*.txt)|*.txt"
    '   - Output (String)
    '       "Supported files *.bmp;*.doc;*.jpg;*.rtf;*.txt;*.tif Bitmap files (*.bmp) *.bmp Word
    '   files (*.doc) *.doc JPEG files (*.jpg) *.jpg RichText files (*.rtf) *.rtf Text files (*.txt)
    '   *.txt"
    '
    '   Check to see if the Filter is set....if not then use the "All Files (*.*)"

    If Len(sFilter) = 0 Then
        sFilter = "Supported Files|*.*|All Files (*.*)"
        '   Make sure to store this in the Control as well...
        m_Filters = sFilter
    End If

    '   Now Replace the Pipes in the Filter String

    For i = 1 To Len(sFilter)

        If (Mid$(sFilter, i, 1) = "|") Then
            Mid$(sFilter, i, 1) = vbNullChar
        End If

    Next i
    '   Pad the string to the correct length

    If (Len(sFilter) < MAX_PATH) Then
        sFilter = sFilter & String$(MAX_PATH - Len(sFilter), 0)
    Else
        sFilter = sFilter & Chr(0) & Chr(0)
    End If

    '   Pass the fixed filter back....
    ProcessFilter = sFilter

End Function

Public Property Let PropertySort(ByVal vPropertySort As PropertyStyleSort)

    m_vPropertySort = vPropertySort

    UserControl.PropertyChanged "PropertySort"

    If (m_vPropertySort = Alphabetical) Then
        If (CountChild > 0) Then
            Insertion
        End If
        
    ElseIf (m_vPropertySort = NoSort) Then
        Let ChildOrd() = ChildItem()
    End If

    If (Ambient.UserMode = False) Then
        ReDraw
        RaiseEvent PropertySortChanged(m_vPropertySort)
    End If

End Property

Public Property Get PropertySort() As PropertyStyleSort

    PropertySort = m_vPropertySort

End Property

Public Function QualifyPath(ByVal sPath As String) As String

  Dim lStrCnt As Long

    '   Look for the PathSep
    lStrCnt = InStrRev(sPath, "\")

    If (lStrCnt <> Len(sPath)) Or Right$(sPath, 1) <> "\" Then
        '   None, so add it...
        QualifyPath = sPath & "\"
    Else
        '   We are good, so return the value unchanged
        QualifyPath = sPath
    End If

End Function

Private Sub ReDraw(Optional ByVal Refresh As Boolean = True, _
                   Optional ByVal DblClick As Boolean = False, _
                   Optional ByVal CategoryFocus As Boolean = False)

  
  Dim i            As Long
  Dim j            As Long
  Dim m_iHeight    As Long
  Dim NeedReDraw   As Boolean
  Dim l            As Long
  Dim lCnt         As Long
  Dim lFocus       As Boolean
  Dim lHg          As Long
  Dim lItems       As Long
  Dim lTotal       As Long
  Dim lTypeB       As Integer
  Dim lValue       As Variant
  Dim sChild       As Integer
  Dim tF           As RECT
  Dim tL           As RECT
  Dim tR           As RECT
  Dim tRBt         As RECT
  Dim tW           As RECT
  Dim xTop         As Long
  Dim lHighColor   As OLE_COLOR
  Dim lHightColorT As OLE_COLOR
  Dim bColor       As OLE_COLOR
  Dim lHeg         As Integer
  Dim lJ           As Integer
  Dim lI           As Integer
  Dim bDrawTheme   As Boolean

On Error Resume Next
    ' Prepare to draw the all control.
    UserControl.Cls

    GetClientRect UserControl.hWnd, tR
    GetClientRect UserControl.hWnd, tL
    GetClientRect UserControl.hWnd, tW
    GetClientRect UserControl.hWnd, tRBt
    
    m_iHeight = UserControl.TextHeight("A") + 5

    If (m_bHelpVisible = False) Then
        lHg = 1
    Else
        lHg = xSplitterY
    End If
    
    lHighColor = ConvertSystemColor(vbHighlight)
    lHightColorT = ConvertSystemColor(vbHighlightText)
    
    If (m_StylePropertyGrid = &H3) Or (m_StylePropertyGrid = &H2) Then
        If (m_StylePropertyGrid = &H2) Then
            lHighColor = ConvertSystemColor(vb3DLight)
            lHightColorT = ConvertSystemColor(vbHighlight)
        Else
            m_lBackColor = RGB(200, 197, 180)
            m_lViewBackColor = RGB(238, 237, 232)
            txtValue.BackColor = m_lViewBackColor
            lHighColor = RGB(49, 106, 197)
            lHightColorT = vbWhite
        End If
        
        SCmb.AppearanceCombo = Office
        SCmb.OfficeAppearance = [Office 2003]
        isBttAction.Style = isbOfficeXP
        isBttAction.Height = 17
    Else
        isBttAction.Style = m_StyleButton
        SCmb.AppearanceCombo = m_StyleComboBox
        isBttAction.Height = 17
    End If

    DrawRectangleBorder UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight _
        - lHg, m_lViewBackColor, False

    tL.rTop = 2
    tW.rTop = 3
    tR.rBottom = m_iHeight
On Error GoTo 0
    
    If (Ambient.UserMode = True) Then
On Error GoTo ReDrawError
        Set isBttAction.Font = m_lFont
        Set SCmb.Font = m_lFont
        Set txtValue.Font = m_lFont
        
        With McCalendar.Font
            .Bold = m_lFont.Bold
            .Charset = m_lFont.Charset
            .Italic = m_lFont.Italic
            .Name = m_lFont.Name
            .Size = m_lFont.Size
            .Strikethrough = m_lFont.Strikethrough
            .Underline = m_lFont.Underline
            .Weight = m_lFont.Weight
        End With
        
        With UpDown.Font
            .Bold = m_lFont.Bold
            .Charset = m_lFont.Charset
            .Italic = m_lFont.Italic
            .Name = m_lFont.Name
            .Size = m_lFont.Size
            .Strikethrough = m_lFont.Strikethrough
            .Underline = m_lFont.Underline
            .Weight = m_lFont.Weight
        End With
        
        With UserControl.Font
            .Bold = m_lFont.Bold
            .Charset = m_lFont.Charset
            .Italic = m_lFont.Italic
            .Name = m_lFont.Name
            .Size = m_lFont.Size
            .Strikethrough = m_lFont.Strikethrough
            .Underline = m_lFont.Underline
            .Weight = m_lFont.Weight
        End With
        
        Value1 = Value
        lItems = 0
        lCnt = 0
        lChild = -1
        sChild = -1
        SCmb.Visible = False
        SCmb.Width = (UserControl.ScaleWidth - xSplitter) - 1
        SCmb.Height = m_iHeight + 1
        lstFX1.BorderStyle = 0
        isBttAction.Visible = False
        isBttAction.Width = 17
        isBttAction.CaptionAlign = isbCenter
        isBttAction.Caption = " ..."
        txtValue.Visible = False
        txtValue.Width = UserControl.ScaleWidth - xSplitter - 10
        txtValue.Height = m_iHeight - 5
        txtValue.PasswordChar = vbNullString
        txtValue.Tag = vbNullString
        McCalendar.Visible = False
        McCalendar.Width = SCmb.Width
        
        If (McCalendar.Width < 241) Then McCalendar.Width = 241
        lTypeB = -1
        lFocus = False
        m_iHeight = UserControl.TextHeight("A") + 5
        lTotal = Int((UserControl.ScaleHeight - lHg) / m_iHeight) ' Total visible items.
        
        With UpDown
            .Height = SCmb.Height - 2
            .Width = SCmb.Width - 5
            .Visible = False
        End With
        
        For i = 0 To CountCatg - 1
            lItems = lItems + 1
            lCnt = lCnt + 1
            lFocus = False

            ' Draw the category items.
            If (lItems >= Value) And (m_vPropertySort = Categorized) Then
                
                ' If the Item selected and DblClick event occur.
                If (Refresh = True) And (lItemSelected = lItems) And (DblClick = True) Then
                    If (Category(i).Expand = True) Then
                        Category(i).Expand = False
                    Else
                        Category(i).Expand = True
                    End If

                    lFocus = True
                    lCateg = i
                    sChild = 2
                    bExpand = True
                ElseIf (Refresh = True) And (lItemSelected = lItems) And (lButton = vbLeftButton) _
                    Then

                    ' Only changed the item.
                    lFocus = True
                    lCateg = i
                    sChild = 2
                    If (bChanged = True) Then RaiseEvent SelectionChanged(Category(i).Key, _
                        Category(i).Title, vbNullString)
                        
                End If

                ' Draw the rectangle border.
                If (m_StylePropertyGrid = &H0) Or (m_StylePropertyGrid = &H3) Then
                    
                    If (m_StylePropertyGrid = &H3) Then
                        
                        If (lFocus = False) Then
                            DrawRectangleBorder UserControl.hDC, 0, tR.rTop, UserControl.ScaleWidth - 1, _
                                m_iHeight, m_lBackColor, False
                        Else
                            DrawRectangleBorder UserControl.hDC, 0, tR.rTop, UserControl.ScaleWidth - 1, _
                                m_iHeight, lHighColor, False
                        End If
                        
                        lJ = 6
                        
                        If (i = 0) Then
                            lHeg = 7
                            lI = tR.rTop + 1
                        Else
                            lHeg = 6
                            lI = tR.rTop
                        End If
                        
                        Do While (lI <= (tR.rTop + lHeg))
                            APILine 0, lI, lJ, lI, m_lViewBackColor
                            APILine (tR.rRight - 1), lI, (tR.rRight - 1) - lJ, lI, m_lViewBackColor
                            lJ = lJ - 1
                            lI = lI + 1
                        Loop
                        
                    Else
                        DrawRectangleBorder UserControl.hDC, 0, tR.rTop, UserControl.ScaleWidth - 1, _
                            m_iHeight, m_lBackColor, False
                    End If
                    
                ElseIf (m_StylePropertyGrid = &H1) Then
                    DrawRectangleBorder UserControl.hDC, 0, tR.rTop, UserControl.ScaleWidth - 1, _
                        m_iHeight, m_lViewBackColor, False
                    
                    ' Draw the horizontal line item.
                    APIRectangle UserControl.hDC, tR.rLeft, tR.rTop, UserControl.ScaleWidth _
                        - tR.rLeft - 1, m_iHeight, ShiftColorOXP(m_lBackColor, &H40)
                Else
                    DrawRectangleBorder UserControl.hDC, 0, tR.rTop, 15, _
                        m_iHeight, ConvertSystemColor(vbHighlight), False
                    
                    DrawVGradient ShiftColorOXP(ConvertSystemColor(GetSysColor(COLOR_2NDACTIVECAPTION))), ShiftColorOXP(ConvertSystemColor(vbHighlight), 110), 15, tR.rTop - 1, _
                        tR.rRight, m_iHeight
                End If
                
                ' Draw the text.
                UserControl.Font.Bold = True
                DrawCaption Category(i).Title, tL, m_lViewCategoryForeColor
                UserControl.Font.Bold = m_lFont.Bold
                
                tF = tR

                If (lFocus = True) Or ((lItemSelected = lItems) And (CategoryFocus = True)) Or _
                   ((lItemSelected = lItems) And (b_Focus = True)) Then
                    
                    tF.rLeft = 15

                    If (tF.rTop <= 0) Then
                        tF.rTop = 1
                        tF.rBottom = m_iHeight
                    Else
                        tF.rBottom = tF.rTop + m_iHeight
                    End If

                    UserControl.Font.Bold = True
                    tF.rRight = TextWidthU(UserControl.hDC, Category(i).Title) + tF.rLeft + 6
                    UserControl.Font.Bold = m_lFont.Bold
                    
                    ' Draw the focus.
                    If (m_StylePropertyGrid = &H0) Then
                        DrawFocusRect UserControl.hDC, tF
                    ElseIf (m_StylePropertyGrid = &H1) Or (m_StylePropertyGrid = &H3) Then
                            
                        If (m_StylePropertyGrid = &H3) Then
                            DrawRectangleBorder UserControl.hDC, 0, tR.rTop, UserControl.ScaleWidth - 1, _
                                m_iHeight, lHighColor, False
                            
                            lFocus = True
                            
                            lJ = 6
                        
                            If (i = 0) Then
                                lHeg = 7
                                lI = tR.rTop + 1
                            Else
                                lHeg = 6
                                lI = tR.rTop
                            End If
                        
                            Do While (lI <= (tR.rTop + lHeg))
                                APILine 0, lI, lJ, lI, m_lViewBackColor
                                APILine (tR.rRight - 1), lI, (tR.rRight - 1) - lJ, lI, m_lViewBackColor
                                lJ = lJ - 1
                                lI = lI + 1
                            Loop
                        
                        Else
                            DrawRectangleBorder UserControl.hDC, 15, tR.rTop, UserControl.ScaleWidth - 1, _
                                m_iHeight, lHighColor, False
                        End If
                            
                        UserControl.Font.Bold = True
                        DrawCaption Category(i).Title, tL, lHightColorT
                        UserControl.Font.Bold = m_lFont.Bold
                    Else
                        DrawVGradient ShiftColorOXP(ConvertSystemColor(GetSysColor(COLOR_2NDACTIVECAPTION))), ShiftColorOXP(ConvertSystemColor(vbHighlight), 110), 15, tR.rTop - 1, _
                            tR.rRight, m_iHeight
                        
                        UserControl.Font.Bold = True
                        DrawCaption Category(i).Title, tL, lHightColorT
                        UserControl.Font.Bold = m_lFont.Bold
                    End If
                    
                    sChild = 2
                    lCateg = i
                    b_Focus = False
                End If

                tW.rLeft = -5
                tW.rBottom = tW.rTop + 13
                tW.rRight = 22
                tR.rBottom = UserControl.ScaleHeight - (lHg - 4)

                If (Category(i).Expand = True) Then
                    If (m_StylePropertyGrid <> &H3) And (DrawXPTheme("Treeview", tW, 2, 2, lFocus) = False) Then
                        ' Draw the Treeview icon (-) if theme fault.
                        UserControl.Line (4 * Screen.TwipsPerPixelY, tW.rTop + 3 * _
                            Screen.TwipsPerPixelY)-Step(8 * 1, 8 * 1), &HFFFFFF, BF
                        UserControl.Line (4 * 1, tW.rTop + 3 * 1)-Step(8 * 1, 8 * 1), &H808080, B
                        UserControl.Line (6, tW.rTop + 7)-Step(5, 0), &H0
                    End If
                
                ElseIf (m_StylePropertyGrid <> &H3) And (DrawXPTheme("Treeview", tW, 1, 2, lFocus) = False) Then
                    ' Draw the Treeview icon (+) if theme fault.
                    UserControl.Line (4 * Screen.TwipsPerPixelY, tW.rTop + 3 * _
                        Screen.TwipsPerPixelY)-Step(8 * 1, 8 * 1), &HFFFFFF, BF
                    UserControl.Line (4 * 1, tW.rTop + 3 * 1)-Step(8 * 1, 8 * 1), &H808080, B
                    UserControl.Line (6, tW.rTop + 7)-Step(5, 0), &H0
                    UserControl.Line (8, tW.rTop + 5)-Step(0, 5), &H0
                End If

                tL.rTop = tL.rTop + m_iHeight
                tW.rTop = tL.rTop
                OffsetRect tR, 0, m_iHeight
            ElseIf (m_vPropertySort <> Categorized) Then
                lItems = lItems - 1
                lCnt = lCnt - 1
            End If

            If (Category(i).Expand = True) Or (m_vPropertySort <> Categorized) Then
                xTop = tL.rTop

                For j = 0 To CountChild - 1
                    If ((Category(i).Key = ChildOrd(j).KeyCategory) And (m_vPropertySort = _
                        Categorized)) Or (m_vPropertySort <> Categorized) Then
                        lItems = lItems + 1

                        If (lItems >= Value) Then
                            lCnt = lCnt + 1
                            tR.rLeft = 15
                            tR.rBottom = UserControl.ScaleHeight - (lHg - 4)
                            
                            If (lItemSelected = lItems) Then
                                
                                If (bUserFocus = False) Then
                                    ' Draw the left rectangle border without focus.
                                    DrawRectangleBorder UserControl.hDC, tR.rLeft, tR.rTop, _
                                        xSplitter, m_iHeight, m_lBackColor, False
                                Else
                                    ' Draw the left rectangle border with focus.
                                    DrawRectangleBorder UserControl.hDC, tR.rLeft, tR.rTop, _
                                        xSplitter, m_iHeight, IIf(ChildOrd(j).TypeGrid <> PropertyItemStringReadOnly, lHighColor, ShiftColorOXP(lHighColor)), _
                                        False
                                End If

                                ' Draw the right rectangle border.
                                DrawRectangleBorder UserControl.hDC, xSplitter, tR.rTop, _
                                    UserControl.ScaleWidth - 1, m_iHeight, m_lViewBackColor, False
                                
                                lChild = j
                                tRBt.rLeft = UserControl.ScaleWidth - 36
                                tRBt.rTop = tL.rTop
                                tRBt.rBottom = tW.rTop + 13
                                tRBt.rRight = tL.rRight - 25
                                
                                If (ChildOrd(j).TypeGrid = PropertyItemStringList) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemPicture) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemFont) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemFolder) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemDate) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemColor) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemFolderFile) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemForm) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemBool) Then
                                    If (ChildOrd(j).TypeGrid = PropertyItemBool) Or _
                                        (ChildOrd(j).TypeGrid = PropertyItemDate) Or _
                                        (ChildOrd(j).TypeGrid = PropertyItemStringList) Then
                                        SCmb.NormalColorText = m_lViewForeColor
                                        SCmb.Move xSplitter, IIf(tRBt.rTop > 2, tRBt.rTop - 2, 1)

                                        If (tRBt.rTop = 2) Then
                                            SCmb.Height = m_iHeight
                                        End If

                                        If (ChildOrd(j).TypeGrid = PropertyItemBool) Then
                                            If (ChildOrd(j).Value = True) Then
                                                SCmb.Text = "True"
                                            ElseIf (ChildOrd(j).Value = False) Then
                                                SCmb.Text = "False"
                                            Else
                                                SCmb.Text = ChildOrd(j).Value
                                            End If

                                        ElseIf (ChildOrd(j).TypeGrid = PropertyItemDate) Then
                                            McCalendar.Value = ChildOrd(j).Value
                                            SCmb.Text = McCalendar.Value
                                        Else
                                            SCmb.Text = ChildOrd(j).ItemValue
                                        End If

                                        If ((SCmb.Top + isBttAction.Height - 2) < (UserControl.ScaleHeight - lHg - 1)) _
                                            Then
                                            lTypeB = 1
                                        ElseIf (NeedReDraw = False) Then
                                            NeedReDraw = True
                                            lTypeB = 0
                                        End If

                                    Else
                                        isBttAction.Move tRBt.rLeft + 18, tRBt.rTop - 1
                                        If ((isBttAction.Top + isBttAction.Height - 2) < (UserControl.ScaleHeight - lHg _
                                            - 1)) Then
                                            lTypeB = 2
                                        ElseIf (NeedReDraw = False) Then
                                            NeedReDraw = True
                                            lTypeB = 0
                                        End If

                                    End If
                                    
                                ElseIf (ChildOrd(j).TypeGrid = PropertyItemString) Or _
                                    (ChildOrd(j).TypeGrid = PropertyItemNumber) Then
                                    
                                    ' Set the style of the textbox.
                                    If (ChildOrd(j).TypeGrid = PropertyItemNumber) Then
                                        m_Style = ItemNumeric
                                    Else
                                        m_Style = ChildOrd(j).StyleString
                                    End If
                                    
                                    If (ChildOrd(j).StyleString = ItemPassword) And (ChildOrd(j).TypeGrid = PropertyItemString) Then
                                        txtValue.PasswordChar = "*"
                                        txtValue.Text = ChildOrd(j).Value
                                    ElseIf (ChildOrd(j).StyleString = ItemLowerCase) Then
                                        txtValue.Text = LCase$(ChildOrd(j).Value)
                                    ElseIf (ChildOrd(j).StyleString = ItemUpperCase) Then
                                        txtValue.Text = UCase$(ChildOrd(j).Value)
                                    ElseIf (ChildOrd(j).StyleString = ItemNumeric) Then
                                        txtValue.Text = Val(Replace$(ChildOrd(j).Value, ",", "."))
                                    Else
                                        txtValue.Text = ChildOrd(j).Value
                                    End If
                                    
                                    txtValue.Move xSplitter + 5, tRBt.rTop
                                    
                                    If (ChildOrd(j).Filters = "LK") Then
                                        txtValue.Locked = True
                                    Else
                                        txtValue.Locked = False
                                    End If
                                    
                                    If (Len(ChildOrd(j).Filters) > 3) Then
                                        If (Mid$(ChildOrd(j).Filters, 1, 3) = "MX:") Then
                                            If (IsNumeric(Mid$(ChildOrd(j).Filters, 4)) = True) Then
                                                txtValue.MaxLength = Val(Mid$(ChildOrd(j).Filters, 4))
                                            End If
                                            
                                        Else
                                            txtValue.MaxLength = 0
                                        End If
                                    
                                    End If

                                    If ((txtValue.Top + isBttAction.Height - 3) < (UserControl.ScaleHeight - lHg - 1)) _
                                        Then
                                        lTypeB = 3
                                    ElseIf (NeedReDraw = False) Then
                                        NeedReDraw = True
                                        lTypeB = 0
                                    End If

                                ElseIf (ChildOrd(j).TypeGrid = PropertyItemUpDown) Then
                                On Error GoTo isABug
                                    UpDown.Min = Mid$(ChildOrd(j).Filters, 1, InStr(1, ChildOrd(j).Filters, ":") - 1)
                                    UpDown.Max = Mid$(ChildOrd(j).Filters, InStr(1, ChildOrd(j).Filters, ":") + 1)
                                    UpDown.Value = ChildOrd(j).Value
                                    GoTo isNoBug
isABug:
                                    UpDown.Min = 1
                                    UpDown.Max = 100
                                    UpDown.Value = 1
isNoBug:
                                    UpDown.Move xSplitter + 4, IIf(tRBt.rTop > 2, tRBt.rTop - 1, 1)
                                    
                                    If ((UpDown.Top + isBttAction.Height - 3) < (UserControl.ScaleHeight - lHg - 1)) _
                                        Then
                                        lTypeB = 4
                                    ElseIf (NeedReDraw = False) Then
                                        NeedReDraw = True
                                        lTypeB = 0
                                    End If
                                    
                                Else
                                    lTypeB = -1
                                End If

                                sChild = 1
                            Else
                                ' Draw the complete rectangle border.
                                DrawRectangleBorder UserControl.hDC, tR.rLeft, tR.rTop, _
                                    UserControl.ScaleWidth - 1, m_iHeight, m_lViewBackColor, False
                            End If
                            
                            ' Draw the horizontal line item (Complete border left).
                            If (m_StylePropertyGrid = &H2) And (bUserFocus = True) _
                                   And (lItemSelected = lItems) _
                            Then
                                DrawRectangleBorder UserControl.hDC, tR.rLeft, tR.rTop, xSplitter - tR.rLeft, m_iHeight, _
                                    ShiftColorOXP(ConvertSystemColor(vbHighlight), 90)
                                
                                APIRectangle UserControl.hDC, xSplitter, tR.rTop, UserControl.ScaleWidth _
                                    - tR.rLeft - 1, m_iHeight, ShiftColorOXP(m_lBackColor, &H40)
                            Else
                                APIRectangle UserControl.hDC, tR.rLeft, tR.rTop, UserControl.ScaleWidth _
                                    - tR.rLeft - 1, m_iHeight, ShiftColorOXP(m_lBackColor, &H40)
                            End If
                            
                            If (bUserFocus = True) And (lItemSelected = lItems) Then
                                ' Draw the focus text.
                                If (m_StylePropertyGrid = &H2) Then
                                    UserControl.Font.Bold = True
                                End If
                                
                                DrawCaption ChildOrd(j).Title, tL, _
                                    lHightColorT, , xSplitter
                                If (m_StylePropertyGrid = &H2) Then
                                    UserControl.Font.Bold = m_lFont.Bold
                                End If
                                
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                If (Mid$(ChildOrd(j).Filters, 1, 3) = "FC:") Then
                                On Error GoTo noColorLine
                                    bColor = ConvertSystemColor(Mid$(ChildOrd(j).Filters, 4))
                                    GoTo OkColorLine
noColorLine:
                                    If (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                        bColor = ConvertSystemColor(vbGrayText)
                                    Else
                                        bColor = ConvertSystemColor(m_lViewForeColor)
                                    End If
OkColorLine:
                                ElseIf (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                    bColor = ConvertSystemColor(vbGrayText)
                                Else
                                    bColor = ConvertSystemColor(m_lViewForeColor)
                                End If
                                
                                ' Draw the readonly text.
                                DrawCaption ChildOrd(j).Title, tL, _
                                    bColor, , xSplitter
                            Else
                                
                                If (ChildOrd(j).TypeGrid = PropertyItemCheckBox) Then
                                    
                                    If (Mid$(ChildOrd(j).Filters, 1, 3) = "FC:") Then
                                        On Error GoTo noColorLine1
                                            bColor = ConvertSystemColor(Mid$(ChildOrd(j).Filters, 4))
                                            GoTo OkColorLine1
noColorLine1:
                                        If (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                            bColor = ConvertSystemColor(vbGrayText)
                                        Else
                                            bColor = ConvertSystemColor(m_lViewForeColor)
                                        End If
OkColorLine1:
                                    ElseIf (ChildOrd(j).ItemValue = False) Then
                                        bColor = ConvertSystemColor(vbGrayText)
                                    Else
                                        bColor = ConvertSystemColor(m_lViewForeColor)
                                    End If
                                    
                                ElseIf (Mid$(ChildOrd(j).Filters, 1, 3) = "FC:") Then
                                On Error GoTo noColorLine2
                                    bColor = ConvertSystemColor(Mid$(ChildOrd(j).Filters, 4))
                                    GoTo OkColorLine2
noColorLine2:
                                    If (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                        bColor = ConvertSystemColor(vbGrayText)
                                    Else
                                        bColor = ConvertSystemColor(m_lViewForeColor)
                                    End If
OkColorLine2:
                                ElseIf (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                    bColor = ConvertSystemColor(vbGrayText)
                                Else
                                    bColor = ConvertSystemColor(m_lViewForeColor)
                                End If
                            
                                ' Draw the normal text.
                                DrawCaption ChildOrd(j).Title, tL, _
                                    bColor, , xSplitter
                            End If
                                                        
                            If (ChildOrd(j).Filters = "WB") Then
                                bColor = &HFF0000
                                UserControl.Font.Underline = True
                            ElseIf (Len(ChildOrd(j).Filters) > 3) Then
                                If (Mid$(ChildOrd(j).Filters, 1, 3) = "FC:") Then
                                On Error GoTo noColor
                                    bColor = ConvertSystemColor(Mid$(ChildOrd(j).Filters, 4))
                                    GoTo OkColor
noColor:
                                    If (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                        bColor = ConvertSystemColor(vbGrayText)
                                    Else
                                        bColor = ConvertSystemColor(m_lViewForeColor)
                                    End If
OkColor:
                                ElseIf (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                    bColor = ConvertSystemColor(vbGrayText)
                                Else
                                    bColor = ConvertSystemColor(m_lViewForeColor)
                                End If
                                
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemStringReadOnly) Then
                                bColor = ConvertSystemColor(vbGrayText)
                            Else
                                bColor = ConvertSystemColor(m_lViewForeColor)
                            End If
                            
                            ' Redraw the left vertical border.
                            If (m_StylePropertyGrid = &H0) Then
                                For l = 1 To 14
                                    APILine l, tR.rTop, l, UserControl.ScaleHeight - lHg - 1, _
                                        m_lBackColor
                                Next l
                                
                            ElseIf (m_StylePropertyGrid = &H2) Then
                                For l = 1 To 14
                                    APILine l, tR.rTop, l, UserControl.ScaleHeight - lHg, _
                                        ConvertSystemColor(vbHighlight)
                                Next l
                                
                            End If

                            tL.rLeft = xSplitter + 5
                            
                            If (ChildOrd(j).TypeGrid = PropertyItemColor) Then
                                ' Draw the color border.
                                DrawRectangleBorder UserControl.hDC, xSplitter + 4, tR.rTop + 3, 18, _
                                    12, ConvertSystemColor(ChildOrd(j).Value), False
                                APIRectangle UserControl.hDC, xSplitter + 4, tR.rTop + 3, 18, 12, _
                                    ShiftColorOXP(ConvertSystemColor(vbButtonText), &H40)
                                tL.rLeft = tL.rLeft + 25
                                
                                Select Case m_SetColorStyle
                                Case &H0
                                    lValue = "RGB(" & LongToRGB(ChildOrd(j).Value) & ")"
                                Case &H1
                                    LongToRGB ChildOrd(j).Value
                                    lValue = RGBToHex(RGBColor.Red, RGBColor.Green, RGBColor.Blue)
                                Case &H2
                                    lValue = MakeWebColor(ChildOrd(j).Value)
                                Case &H3
                                    lValue = "0x00" & LongToHex(ChildOrd(j).Value)
                                Case &H4
                                    lValue = "$00" & LongToHex(ChildOrd(j).Value)
                                Case &H5
                                    lValue = "0x" & StrReverse(LongToHex(ChildOrd(j).Value))
                                Case &H6
                                    lValue = StrReverse(LongToHex(ChildOrd(j).Value))
                                Case Else
                                    lValue = "RGB(" & LongToRGB(ChildOrd(j).Value) & ")"
                                End Select
                                
                                If (ChildOrd(j).Filters = "RO") Then
                                    lTypeB = 5
                                ElseIf (ChildOrd(j).Filters = "RA") Then
                                    lTypeB = 5
                                    lValue = vbNullString
                                End If
                                
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemCheckBox) Then
                              
                              Dim m_Buttons As RECT
                              Dim isOpt As Integer
                              Dim isDisable As Integer
                            
                                m_Buttons.rLeft = xSplitter + 1
                                m_Buttons.rTop = tR.rTop + 2
                                m_Buttons.rBottom = tR.rTop + 18
                                m_Buttons.rRight = m_Buttons.rLeft + 18
                                
                                If (lItemSelected = lItems) And _
                                  (ChildOrd(j).ItemValue = True) Then
                                    If (DblClick = True) And (setYet = False) Then
                                        ChildOrd(j).Value = Not (ChildOrd(j).Value)
                                        setYet = True
                                    ElseIf (setYet = False) And (lXPos > xSplitter) And (lXPos < (xSplitter + 14)) Then
                                        ChildOrd(j).Value = Not (ChildOrd(j).Value)
                                        setYet = True
                                    End If
                                    
                                End If
                                
                                isOpt = (ChildOrd(j).Value * -5)
                                
                                If (ChildOrd(j).ItemValue = False) Then
                                    isDisable = DFCS_INACTIVE
                                    
                                    If (ChildOrd(j).Value = True) Then
                                        isOpt = 8
                                    Else
                                        isOpt = 4
                                    End If
                                    
                                End If
                                
                                bDrawTheme = DrawTheme("Button", 3, isOpt, m_Buttons)
                                lValue = vbNullString
                                
                                If (bDrawTheme = False) Then
                                    If (ChildOrd(j).Value = True) Then
                                        DrawFrameControl hDC, m_Buttons, DFC_BUTTON, _
                                            DFCS_BUTTONCHECK Or DFCS_CHECKED Or isDisable
                                    Else
                                        DrawFrameControl hDC, m_Buttons, DFC_BUTTON, _
                                            DFCS_BUTTONCHECK Or isDisable
                                    End If
                                    
                                End If
                            
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemBool) Then

                                If (ChildOrd(j).Value = True) Then
                                    lValue = "True"
                                ElseIf (ChildOrd(j).Value = False) Then
                                    lValue = "False"
                                Else
                                    lValue = ChildOrd(j).Value
                                End If

                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemNumber) Then

                                If (IsNumeric(ChildOrd(j).Value) = True) Then
                                    lValue = ChildOrd(j).Value
                                Else
                                    lValue = 0
                                End If

                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemDate) Then

                                If (IsDate(ChildOrd(j).Value) = True) Then
                                    McCalendar.Value = CDate(ChildOrd(j).Value)
                                    lValue = McCalendar.Value
                                Else
                                    lValue = vbNullString
                                End If

                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemFolder) Then
                                lValue = ChildOrd(j).Value
                            
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemFolderFile) Then
                                
                                lValue = ChildOrd(j).Value
                                
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemFont) Then
                                Dim lFont As New StdFont
                                
                                Set lFont = CloneFont(ChildOrd(j).theFont)
                                If Not (lFont Is Nothing) Then
                                    lValue = lFont.Name & "; " & CInt(lFont.Size) & "pt"
                                Else
                                    lValue = vbNullString
                                End If
                                
                                lFont = m_lFont
                                
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemForm) Then
                                
                                lValue = ChildOrd(j).Value
                                
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemPicture) Then
                                Dim lPict As StdPicture

                                If Not (ChildOrd(j).Value Is Nothing) Then
                                    Set lPict = ChildOrd(j).Value
                                End If

                                If Not (lPict Is Nothing) Then
                                    If (lPict.Type = 0) Then
                                        lValue = "(Nothing)"
                                    ElseIf (lPict.Type = 1) Then
                                        lValue = "(Bitmap)"
                                    ElseIf (lPict.Type = 3) Then
                                        lValue = "(Icon)"
                                    Else
                                        lValue = "(Bitmap)"
                                    End If

                                Else
                                    If Not (ChildOrd(j).Value Is Nothing) Then
                                    On Error Resume Next
                                        Set lPict = LoadPicture(ChildOrd(j).Value)
                                    On Error GoTo 0
                                    End If

                                    If Not (lPict Is Nothing) Then
                                        If (lPict.Type = 0) Then
                                            lValue = "(Nothing)"
                                        ElseIf (lPict.Type = 1) Then
                                            lValue = "(Bitmap)"
                                        ElseIf (lPict.Type = 3) Then
                                            lValue = "(Icon)"
                                        Else
                                            lValue = "(Bitmap)"
                                        End If

                                    Else
                                        lValue = "(Nothing)"
                                    End If

                                End If
                            ElseIf (ChildOrd(j).TypeGrid = PropertyItemStringList) Then
                                lValue = ChildOrd(j).ItemValue
                            ElseIf (ChildOrd(j).StyleString = ItemPassword) Then
                                lValue = String$(Len(ChildOrd(j).Value), "*")
                            Else
                                
                                Select Case ChildOrd(j).StyleString
                                Case ItemUpperCase
                                     lValue = UCase$(ChildOrd(j).Value)
                                Case ItemLowerCase
                                     lValue = LCase$(ChildOrd(j).Value)
                                Case ItemNumeric
                                     lValue = Val(ChildOrd(j).Value)
                                Case Else
                                    lValue = ChildOrd(j).Value
                                End Select
                                
                            End If

                            ' Draw the value text.
                            DrawCaption lValue, tL, bColor, tL.rLeft, -1
                            If (lItemSelected = lItems) Then
                                UpDown.ForeColor = bColor
                                txtValue.ForeColor = bColor
                                SCmb.NormalColorText = bColor
                            End If
                            
                            UserControl.Font.Underline = False
                            
                            tL.rTop = tL.rTop + m_iHeight
                            tW.rTop = tL.rTop
                            OffsetRect tR, 0, m_iHeight

                            If (lItemSelected = lItems) And (bChanged = True) Then
                              Dim kValue As Variant

                                If (ChildOrd(j).TypeGrid = PropertyItemColor) Then
                                    kValue = ChildOrd(j).Value
                                Else
                                    kValue = lValue
                                End If

                                bChanged = False
                                RaiseEvent SelectionChanged(ChildOrd(lChild).KeyCategory, _
                                    ChildOrd(lChild).Title, j)
                            End If

                        End If
                        
                    End If
                    
                Next j

                ' Draw the vertical line item (for split position +).
                APIRectangle UserControl.hDC, xSplitter, xTop - 2, 0.1, (tL.rTop - xTop), _
                    ShiftColorOXP(m_lBackColor, &H40)
            End If

            If (m_vPropertySort <> Categorized) Then Exit For
            
        Next i

    End If

ReDrawError:

On Error Resume Next
    ' Redraw the left vertical border.
    If (m_StylePropertyGrid = &H0) Then
        For l = tR.rTop To (UserControl.ScaleHeight - lHg - 1)
            APILine 1, l, UserControl.ScaleWidth - 1, l, m_lViewBackColor
        Next l
        
    End If

    If (m_StylePropertyGrid <> &H3) Then
        ' Draw the last rectangle in the grid.
        APIRectangle UserControl.hDC, tR.rLeft, tR.rTop, UserControl.ScaleWidth - tR.rLeft - 1, 0, _
            ShiftColorOXP(m_lBackColor, &H40)
    End If

    If (m_StylePropertyGrid = &H0) Then
        For l = 1 To 15
            APILine l, tR.rTop, l, tR.rTop + 1, m_lBackColor
        Next l
        
    End If

    ' Draw the external border of the control.
    APIRectangle UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - lHg, _
        m_lLineColor

    ' Draw the Rectangle if the Help is True.
    If (m_bHelpVisible = True) Then
    On Error Resume Next
        DrawRectangleBorder UserControl.hDC, 0, UserControl.ScaleHeight - lHg + 1, _
            UserControl.ScaleWidth, UserControl.ScaleHeight - (UserControl.ScaleHeight - lHg + _
            5) - 1, ConvertSystemColor(Extender.Parent.BackColor), False
        DrawRectangleBorder UserControl.hDC, 0, UserControl.ScaleHeight - lHg + 5, _
            UserControl.ScaleWidth, UserControl.ScaleHeight - (UserControl.ScaleHeight - lHg + _
            5) - 1, m_lHelpBackColor, False
        APIRectangle UserControl.hDC, 0, UserControl.ScaleHeight - lHg + 5, UserControl.ScaleWidth _
            - 1, UserControl.ScaleHeight - (UserControl.ScaleHeight - lHg + 5) - 1, m_lLineColor
    On Error GoTo 0
    End If
    
    ' If Selected the last visible item then redraw the control again.
    l = (lItemSelected - lTotal) + 1
    If ((l > Value) And (bRedraw = False)) Or ((NeedReDraw = True) And (bRedraw = False)) Then
        If (sChild = 2) Or (sChild = 1) Then
            Value = l
            bRedraw = True
            ReDraw
        End If
        
    End If

    lTotalItems = lItems
    
    ' Show/Hide the Scrollbar.
    If (lItems > lTotal) And (lItems > 0) Then
        Max = (lItems - lTotal) + 1
        ScrollVisible True
        bVScroll = True

        ' Moving the control again if the scrollbar's is show.
        If (lTypeB = 1) Then
            SCmb.Move SCmb.Left, SCmb.Top, (UserControl.ScaleWidth - xSplitter) - 19, SCmb.Height
        ElseIf (lTypeB = 2) Then
            isBttAction.Move UserControl.ScaleWidth - 36, isBttAction.Top, isBttAction.Width, _
                isBttAction.Height
        ElseIf (lTypeB = 3) Then
            txtValue.Move txtValue.Left, txtValue.Top, (UserControl.ScaleWidth - xSplitter) - 25, 13
        ElseIf (lTypeB = 4) Then
            UpDown.Move UpDown.Left, UpDown.Top, SCmb.Width - 22
        End If

    ElseIf (Value = 0) Then
        Max = lItems
        ScrollVisible False
        bVScroll = False
    End If

    ' Hide the Scrollbar if the total items show is minor than total items add.
    If (lCnt < lTotal) And (m_vPropertySort = Categorized) Then
        ScrollVisible False
        bVScroll = False
        If (bRedraw = False) Then
            Value = 0
            bRedraw = True
            ReDraw
        End If
        
    End If
        
    ' Show the special control's if the control have the focus and refresh is true.
    If ((Refresh = True) And (bUserFocus = True)) Or (CategoryFocus = True) Then
        If (lTypeB = 1) Then
            SCmb.Refresh
            SCmb.Visible = True
            SCmb.ZOrder 0
        ElseIf (lTypeB = 2) Then
            isBttAction.Visible = True
            isBttAction.ZOrder 0
        ElseIf (lTypeB = 3) Then
            txtValue.Visible = True
         On Error Resume Next
            If (DblClick = True) Then
                txtValue.SelStart = 0
                txtValue.SelLength = Len(txtValue.Text)
                txtValue.SetFocus
                txtValue.Tag = "Y"
            Else
                txtValue.SelStart = 0
                txtValue.Tag = vbNullString
            End If
            
          On Error GoTo 0
            txtValue.ZOrder 0
        ElseIf (lTypeB = 4) Then
            UpDown.Visible = True
        On Error Resume Next
            If (DblClick = True) Then
                UpDown.SetFocus
            End If
        On Error GoTo 0
            UpDown.ZOrder 0
        End If

    ElseIf (lstFX1.Visible = True) Then
        SCmb.Refresh
        SCmb.Visible = True
        SCmb.ZOrder 0
    End If

    ' Move the Scrollbar.
    MoveWindow m_hWnd, UserControl.ScaleWidth - 19, 1, 18, UserControl.ScaleHeight - lHg - 1, 1

    ' Draw ToolTipText in the select item #(^~^)# (If HelpVisible = True).
    If (m_bHelpVisible = True) Then
        If (sChild = 1) Then ' Draw in the Child Item.
            UserControl.Font.Bold = True
            tR.rTop = (UserControl.ScaleHeight - xSplitterY) + 10
            DrawCaption ChildOrd(lChild).Title, tR, m_lHelpForeColor, 5
            UserControl.Font.Bold = m_lFont.Bold
            tR.rTop = tR.rTop + UserControl.TextHeight(ChildOrd(lChild).Title) + 5
            DrawCaption ChildOrd(lChild).ToolTipText, tR, m_lHelpForeColor, 5, , True
             lTCaption = ChildOrd(lChild).ToolTipText
            lTTitle = ChildOrd(lChild).Title
        ElseIf (sChild = 2) Then ' Draw in the Parent Item.
            UserControl.Font.Bold = True
            tR.rTop = (UserControl.ScaleHeight - xSplitterY) + 10
            DrawCaption Category(lCateg).Title, tR, m_lHelpForeColor, 5
            UserControl.Font.Bold = m_lFont.Bold
            tR.rTop = tR.rTop + UserControl.TextHeight(Category(lCateg).Title) + 5
            DrawCaption Category(lCateg).ToolTipText, tR, m_lHelpForeColor, 5, , True
            lTCaption = Category(lCateg).ToolTipText
            lTTitle = Category(lCateg).Title
        ElseIf (LenB(Trim$(lTTitle)) > 0) Then
            UserControl.Font.Bold = True
            tR.rTop = (UserControl.ScaleHeight - xSplitterY) + 10
            DrawCaption lTTitle, tR, m_lHelpForeColor, 5
            UserControl.Font.Bold = m_lFont.Bold
            tR.rTop = tR.rTop + UserControl.TextHeight(Category(lCateg).Title) + 5
            DrawCaption lTCaption, tR, m_lHelpForeColor, 5, , True
        End If

    End If
    
    ' Draw AutoFilter option.
    If (SCmb.Visible = True) And (lChild > 0) And (lItemSelected > 0) And (m_lAutoFilter = True) Then
        If (ChildOrd(lChild).TypeGrid = PropertyItemStringList) Then
            
        End If
        
    End If
    
    If (sChild = 2) Then RaiseEvent Expand(Category(lCateg).Expand, Category(lCateg).Key)
    
On Error GoTo 0
    
End Sub

Public Sub Refresh(Optional Value As Boolean = True)

    ReDraw Value

End Sub

Public Sub RemoveCategory(ByVal KeyCategory As String)

  Dim i As Long
  Dim Categ() As TypeCateg
  Dim LXd As Long

    ReDim Categ(0)
    LXd = 0

    For i = 0 To CountCatg - 1

        If (KeyCategory <> Category(i).Key) Then
            ReDim Preserve Categ(LXd)
            Categ(LXd).Expand = Category(i).Expand
            Categ(LXd).Key = Category(i).Key
            Categ(LXd).Title = Category(i).Title
            Categ(LXd).ToolTipText = Category(i).ToolTipText
            LXd = LXd + 1
        End If

    Next i
    
    ReDim Category(0)
    Let Category = Categ
    CountCatg = LXd
    RemoveItem KeyCategory, vbNullString, True
    Refresh True
    
End Sub

Public Sub RemoveChildItem(ByVal KeyCategory As String, ByVal Title As String)

  RemoveItem KeyCategory, Title
    
End Sub

Private Sub RemoveItem(ByVal KeyCategory As String, Optional ByVal Title As String = vbNullString, Optional ByVal OnlyCategory As Boolean = False)

  Dim i As Long
  Dim Categ() As TypeChildItem
  Dim LXd As Long

    ReDim Categ(0)
    LXd = 0

    For i = 0 To CountChild - 1

        If ((KeyCategory <> ChildItem(i).KeyCategory) And (Title <> ChildItem(i).Title)) Or _
            ((OnlyCategory = True) And (KeyCategory <> ChildItem(i).KeyCategory)) _
        Then
            ReDim Preserve Categ(LXd)
            Categ(LXd).Filters = ChildItem(i).Filters
            Categ(LXd).ItemValue = ChildItem(i).ItemValue
            Categ(LXd).KeyCategory = ChildItem(i).KeyCategory
            Categ(LXd).StyleString = ChildItem(i).StyleString
            Categ(LXd).theFont = ChildItem(i).theFont
            Categ(LXd).Title = ChildItem(i).Title
            Categ(LXd).ToolTipText = ChildItem(i).ToolTipText
            Categ(LXd).TypeGrid = ChildItem(i).TypeGrid
            Categ(LXd).Value = ChildItem(i).Value
            LXd = LXd + 1
        End If

    Next i
    
    ReDim ChildItem(0)
    Let ChildItem = Categ
    CountChild = LXd
    If (OnlyCategory = False) Then Refresh True

End Sub

Public Function RGBToColor(ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
 
    RGBToColor = RGB(R, G, B)
    
End Function

Public Function RGBToHex(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As String

  Dim Hex1 As String
  Dim Hex2 As String
  Dim Hex3 As String
 
    If (R < 16) Then
        Hex1 = 0 & Hex$(R)
    Else
        Hex1 = Hex$(R)
    End If
    
    If (R < 16) Then
        Hex2 = 0 & Hex$(G)
    Else
        Hex2 = Hex$(G)
    End If
 
    If (B < 16) Then
        Hex3 = 0 & Hex$(B)
    Else
        Hex3 = Hex$(B)
    End If
    
    RGBToHex = "&H00" & Hex1 & Hex2 & Hex3 & "&"
    
End Function

Private Sub SCmb_Click()

  Dim iList   As Long, lIndexA  As Long
  Dim oRect   As RECT, sHeight  As String
  Dim mSplit  As Variant, jList As Long
                                
On Error GoTo ClickError
    ' Show the list.
    McCalendar.Visible = False
    lstFX1.Visible = False
    lstFX1.Clear

    GetWindowRect SCmb.hWnd, oRect
    If (lChild >= 0) Then

        Select Case ChildOrd(lChild).TypeGrid
        Case &H0 ' True/False.

            With lstFX1
                Set .Font = SCmb.Font
                .AddItem CStr("True")
                .AddItem CStr("False")
                
                If (NoShowList = False) Then
                    lIndexA = .FindFirst(IIf(CBool(ChildOrd(lChild).Value) = True, "True", "False"), _
                        , True)
                    If (lIndexA >= 0) Then
                        .ListIndex = lIndexA
                    Else
                        lIndexA = 0
                    End If
                
                End If
                
                .Width = (25 + TextWidthU(hDC, "False")) * ScaleY(Screen.TwipsPerPixelY, vbTwips, vbPixels) + _
                    (.Width - .Width * ScaleX(Screen.TwipsPerPixelX, vbTwips, vbPixels))
                    
                If (.Width < SCmb.Width) Then
                    .Width = SCmb.Width
                    .Move oRect.rLeft, oRect.rTop + 20
                Else
                    .Move (oRect.rLeft + SCmb.Width) - .Width, oRect.rTop + 20
                End If
                
                .VisibleRows = 2
                .Refresh
                
                If (NoShowList = False) Then
                    .Visible = True
                    .ZOrder 0
                    SetFocus hWnd
                End If
            
            End With

        Case &H2 ' Date.
            If (NoShowList = False) Then
              With McCalendar
                  .Value = CDate(ChildOrd(lChild).Value)
                  .Move (oRect.rLeft + SCmb.Width) - .Width, oRect.rTop + 20
                  .Visible = True
                  .ZOrder 0
              End With
              
            End If

        Case &HA ' StringList.

            If (IsArray(ChildOrd(lChild).Value) = True) Then
                
                sHeight = vbNullString
                With lstFX1
                    Set .Font = SCmb.Font

                    For iList = 0 To UBound(ChildOrd(lChild).Value)
                        .AddItem CStr(ChildOrd(lChild).Value(iList))
                        If (Len(ChildOrd(lChild).Value(iList)) > Len(sHeight)) Then
                            sHeight = ChildOrd(lChild).Value(iList)
                        End If
                    
                    Next iList
                    
                    If (ChildOrd(lChild).Filters = "MS") Then
                        .SelectMode = &H1
                        mSplit = Split(ChildOrd(lChild).ItemValue, ", ")
                        
                        If (UBound(mSplit) > 0) Then
                            jList = UBound(mSplit)
                            For iList = 0 To jList
                                lIndexA = .FindFirst(mSplit(iList), , True)
                                If (lIndexA >= 0) Then
                                    .ItemSelected(lIndexA) = True
                                End If
                                
                            Next iList
                            
                        Else
                            lIndexA = .FindFirst(ChildOrd(lChild).ItemValue, , True)
                            If (lIndexA >= 0) Then
                                .ItemSelected(lIndexA) = True
                            End If
                            
                        End If
                        
                    Else
                        .SelectMode = &H0
                    End If

                    If (NoShowList = False) Then
                        lIndexA = .FindFirst(ChildOrd(lChild).ItemValue, , True)
                        If (lIndexA >= 0) Then
                            .ListIndex = lIndexA
                        Else
                            lIndexA = 0
                        End If
                    
                    End If
                                        
                    .Width = (25 + TextWidthU(hDC, sHeight)) * ScaleY(Screen.TwipsPerPixelY, vbTwips, _
                        vbPixels) + (.Width - .Width * ScaleX(Screen.TwipsPerPixelX, vbTwips, _
                        vbPixels))
                        
                    If (.Width < SCmb.Width) Then
                        .Width = SCmb.Width
                        .Move oRect.rLeft, oRect.rTop + 20
                    ElseIf (.Width > 300) Then
                        .Width = 300
                        .Move (oRect.rLeft + SCmb.Width) - .Width, oRect.rTop + 20
                    Else
                        .Move (oRect.rLeft + SCmb.Width) - .Width, oRect.rTop + 20
                    End If
                    
                    .VisibleRows = IIf(.ListCount < 8, .ListCount, 8)
                    .Refresh
                    
                    If (NoShowList = False) Then
                        .Visible = True
                        .ZOrder 0
                        SetFocus hWnd
                    End If
                    
                End With

            End If
            
        End Select
        
    End If
    
    Exit Sub
    
ClickError:

End Sub

Private Sub SCmb_CloseList()

    If (lstFX1.Visible = True) Then
        lstFX1.Visible = False
    ElseIf (McCalendar.Visible = True) Then
        McCalendar.Visible = False
    End If

End Sub

Private Sub ScrollVisible(ByVal bState As Boolean)

    '---------------------------------------------------------------------------------------
    ' Procedure : ScrollVisible
    ' DateTime  : 10/07/05 09:24
    ' Author    : HACKPRO TM
    ' Purpose   : Set visible the Scrollbar's.
    '---------------------------------------------------------------------------------------

  Dim lR As Long

    If (m_hWnd = 0) Then Exit Sub

    UserControl.BackColor = ShiftColorOXP(ConvertSystemColor(vbScrollBars))

    If (m_bNoFlatScrollBars = True) Then
        If (Err.Number <> 0) Then m_bNoFlatScrollBars = False
        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, FSB_ENCARTA_MODE, True)
    End If

    If (m_bNoFlatScrollBars = False) Then
        ShowScrollBar m_hWnd, SB_CTL, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, SB_CTL, Abs(bState)
    End If

End Sub

' Subclassing for Paul Caton --> The Man (^~^)#
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

Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal newValue As Long)

    'Let the subclasser lParamUser callback parameter

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the
        '   thunk hasn't already released its memory
        zData(IDX_PARM_USER) = newValue                                         'Set the lParamUser
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

    '-SelfSub
    '   code------------------------------------------------------------------------------------

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

Public Sub Selected(ByVal Index As Integer)

    lItemSelected = Index

    ReDraw

End Sub

Public Property Get SetColorStyle() As ColorStyle

    SetColorStyle = m_SetColorStyle

End Property

Public Property Let SetColorStyle(ByVal eSetColorStyle As ColorStyle)

    m_SetColorStyle = eSetColorStyle

    UserControl.PropertyChanged "SetColorStyle"

End Property

Public Sub SetKeyChild(ByVal KeyCategory As String, ByVal Title As String, _
    ByVal KeyName As String)

On Error GoTo SetKeyChildErr
    If (FindChild(Title, KeyCategory) = True) And (FindChildKey(Title, KeyName) = False) Then
        ChildOrd(GetChildIndex(Title, KeyCategory)).KeyName = KeyName
        ChildItem(GetChildIndex(Title, KeyCategory)).KeyName = KeyName
    End If
    
    Exit Sub
    
SetKeyChildErr:
    
    Err.Raise 513, "AddChildItem", "Error trying set Key of this Child Item."
    
End Sub

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

Private Function ShowColor() As SelectedColor

  Dim lRet As Long

    '   Color Common Dialog Controls
    With ColorDialog
        .hWndOwner = UserControl.Parent.hWnd
        .lStructSize = Len(ColorDialog)

        If m_ColorFlags <> 0 Then
            .Flags = m_ColorFlags
        Else
            .Flags = ShowColor_Default
        End If
        
    End With

    lRet = ChooseColor(ColorDialog)

    If lRet Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If

End Function

Private Function ShowFolderBrowse(Optional Title As String = "Please Select a Folder:") As String

  Dim bInf As BROWSEINFO
  Dim RetVal As Long
  Dim PathID As Long
  Dim RetPath As String
  Dim Offset As Integer
  Dim lpSelPath As Long

    'Set the properties of the folder dialog
    With bInf
        .hWndOwner = UserControl.Parent.hWnd
        '   Set the Root Folder to be the DeskTop
        .pIDLRoot = 0
        '   Pass our Title
        .lpszTitle = lStrCat(Title, vbNullString)
        
        '   What style? Use the Flasgs set by the properties dialog
        If m_FolderFlags <> 0 Then
            .ulFlags = m_FolderFlags
        Else
            .ulFlags = ShowFolder_Default
        End If

        '   If the Path is not set then set it as the location of the App
        If (LenB(m_Path) = 0) Then
            m_Path = App.Path & vbNullChar
        Else
            m_Path = m_Path & vbNullChar
        End If

        '   Note: This is under development.....
        '   Get address of function.
'        If bSubClass Then
'            .lpfnCallback = zAddressOf(UserControl, 1)   'zData(0)
'        End If
        '   Now the fun part. Allocate some memory for the dialog's
        '   selected folder path (sSelPath), blast the string into
        '   the allocated memory, and set the value of the returned
        '   pointer to lParam. Note: VB's StrPtr function won't
        '   work here because a variable's memory address goes out
        '   of scope when passed to SHBrowseForFolder.
        'lpSelPath = LocalAlloc(LPTR, Len(m_Path))
        
        '   Did our string accolation work?
        If lpSelPath Then
            '   Copy this to a pointer
            'RtlMoveMemory ByVal lpSelPath, ByVal m_Path, Len(m_Path)
            '   Pass this to the structure
            '.lParam = lpSelPath
        End If

    End With
    
    '   Show the Browse For Folder dialog
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)

    If RetVal Then
        'Trim off the null chars ending the path
        'and display the returned folder
        Offset = InStr(RetPath, Chr$(0))
        ShowFolderBrowse = Left$(RetPath, Offset - 1)
        'Free memory allocated for PIDL
        CoTaskMemFree PathID
    Else
        ShowFolderBrowse = vbNullString
    End If

End Function

Private Function ShowFont(ByVal oFont As StdFont, ByVal lFontColor As OLE_COLOR) As SelectedFont

  Dim lRet As Long
  Dim lfLogFont As LOGFONT
  Dim i As Integer
  Dim StartingFontName As String

    '   Font Common Dialog Controls
    '   Note: This has been modified to allow the caller to pass
    '         previous instance data to the Dialogs (i.e. FontName, PoitSize, Color...)
    With lfLogFont
        .lfHeight = 0                           ' determine default height
        .lfWidth = 0                            ' determine default width
        .lfEscapement = 0                       ' angle between baseline and escapement vector
        .lfOrientation = 0                      ' angle between baseline and orientation vector
        .lfCharSet = oFont.Charset              ' use default character set
        .lfOutPrecision = OUT_DEFAULT_PRECIS    ' default precision mapping
        .lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
        .lfQuality = DEFAULT_QUALITY            ' default quality setting
        .lfPitchAndFamily = DEFAULT_PITCH       ' default pitch, proportional with serifs
        .lfItalic = oFont.Italic
        .lfStrikeOut = oFont.Strikethrough
        .lfUnderline = oFont.Underline
        .lfWeight = oFont.Weight

    End With

    With FontDialog

        If m_FontFlags <> 0 Then
            .Flags = m_FontFlags
        Else
            .Flags = ShowFont_Default
        End If

        .hDC = UserControl.Parent.hDC
        .hWndOwner = UserControl.Parent.hWnd
        .iPointSize = oFont.Size * 10 '   10pt
        .lCustData = 0
        .lpfnHook = 0
        .lpLogFont = VarPtr(lfLogFont)
        .lpTemplateName = Space$(2048)
        .lStructSize = Len(FontDialog)
        .nFontType = Screen.FontCount
        .nSizeMax = 72
        .nSizeMin = 8
        .rgbColors = lFontColor
    End With

    StartingFontName = oFont.Name

    For i = 0 To Len(StartingFontName) - 1
        lfLogFont.lfFacename(i) = Asc(Mid(StartingFontName, i + 1, 1))
    Next i

    lRet = ChooseFont(FontDialog)

    If lRet Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        m_FontColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10

        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFacename(i))
        Next i

        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, _
            Chr(0)) - 1)
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If

End Function

Private Function ShowOpen(sFilter As String) As SelectedFile

  Dim lRet As Long
  Dim Count As Integer
  Dim LastCharacter As Integer
  Dim NewCharacter As Integer
  Dim tempFiles(1 To 200) As String

    'Dim lhWnd As Long
    '   Open Common Dialog Controls
    '   Note: This has been modified to allow the user to select either
    '         a Single or Mutliple Files...In either case the data is sent
    '         back to the caller as part of the SelectedFile data structure
    '         which has been modified to allow for Array of strings in the
    '         sFiles section.
    m_MultiSelect = False

    With FileDialog
        .nStructSize = Len(FileDialog)
        .hWndOwner = UserControl.Parent.hWnd
        .sFileTitle = Space$(2048)
        .nTitleSize = Len(FileDialog.sFileTitle)
        .sFile = .sFile & Space$(2047) & Chr$(0)
        .nFileSize = Len(FileDialog.sFile)

        If m_FileFlags <> 0 Then
            .Flags = m_FileFlags
        Else
            .Flags = ShowOpen_Default
        End If

        If m_MultiSelect Then
            .Flags = .Flags Or AllowMultiselect
        End If

        '   Init the File Names
        .sFile = .sFile & Space$(2047) & Chr$(0)
        '   Process the Filter string to replace the
        '   pipes and fix the len to correct dims
        sFilter = ProcessFilter(sFilter)
        '   Set the Filter for Use...
        .sFilter = sFilter
    End With

    '   Open the Common Dialog via API Calls
    lRet = GetOpenFileName(FileDialog)

    If lRet Then
        '   Retry Flag
GoAgain:

        If (FileDialog.nFileOffset = 0) Then
            '   This is a first time through, so the Offset will be zero. This is the
            '   case when MultiSelect = False and this is our first file selected.
            '   For cases where this is not our first time, then see "Else" notes below.
            '
            '   Extract the single Filename and pass it back....
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, _
                FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        ElseIf (InStr(FileDialog.nFileOffset, FileDialog.sFile, Chr$(0)) = FileDialog.nFileOffset) _
            Then
            '   See if we have an offset by the dialog and see if this matches the position of
            '   the (Chr$(0)) character. If this is the case, then we have Mulplitple files selected
            '   in the FileDialog.sFile array. The GetOpenFileName passes back (Chr$(0)) delimited
            '   filenames
            '   when we are in Multipile File selection mode, and the stripping of the names needs
            '   to be handled
            '   differently than when there is simply one....
            '
            '   Extract all of the files selected and pass them back in an array.
            LastCharacter = 0
            Count = 0

            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)

                If Count > 0 Then
                    tempFiles(Count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - _
                        LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter _
                        - LastCharacter - 1)
                End If

                Count = Count + 1

                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = _
                    InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) _
                    Then
                    tempFiles(Count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + _
                        1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If

                LastCharacter = NewCharacter
            Wend

            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)

            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next Count

        Else
            '   This is the case where we have MutliSelect = False, but this is our
            '   Second through "n" times through...To fix this case we simlply set the
            '   FileOffset like it is our first time and then re-run the routine....
            '   The net effect is that the sub acts as if this were the first time and
            '   yeilds the name and path correctly.
            FileDialog.nFileOffset = 0
            GoTo GoAgain
        End If

        ShowOpen.bCanceled = False
        Exit Function
    Else
        '   The Cancel Button was pressed
        ShowOpen.sLastDirectory = vbNullString
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If

End Function

Private Property Let SmallChange(ByVal lSmallChange As Long)

    m_lSmallChange = lSmallChange

End Property

Private Property Get SmallChange() As Long

    SmallChange = m_lSmallChange

End Property

Public Property Get SplitterPos() As Single

    '---------------------------------------------------------------------------------------
    ' Procedure : SplitterPos
    ' DateTime  : 24/11/2006 20:39
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    SplitterPos = m_iSplitterPos

End Property

Public Property Let SplitterPos(ByVal iSplitterPos As Single)

    '---------------------------------------------------------------------------------------
    ' Procedure : SplitterPos
    ' DateTime  : 24/11/2006 20:39
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    If (IsNumeric(iSplitterPos) = True) Then
        If (iSplitterPos > 0) Then
            m_iSplitterPos = iSplitterPos
        Else
            m_iSplitterPos = 0.5
        End If

        xSplitter = m_iSplitterPos * 220

        UserControl.PropertyChanged "SplitterPos"

    Else
        Err.Raise 514, "SplitterPos", Err.Description
    End If

End Property

Public Property Get StyleButton() As isbStyle

    StyleButton = m_StyleButton

End Property

Public Property Let StyleButton(ByVal mStyleButton As isbStyle)

    m_StyleButton = mStyleButton
    isBttAction.Style = m_StyleButton

    UserControl.PropertyChanged "StyleButton"

End Property

Public Property Get StyleComboBox() As ComboAppearance

    StyleComboBox = m_StyleComboBox

End Property

Public Property Let StyleComboBox(ByVal mStyleComboBox As ComboAppearance)

    m_StyleComboBox = mStyleComboBox
    SCmb.AppearanceCombo = m_StyleComboBox

    UserControl.PropertyChanged "StyleComboBox"

End Property

Public Property Get StylePropertyGrid() As PropertyGridStyle

    StylePropertyGrid = m_StylePropertyGrid

End Property

Public Property Let StylePropertyGrid(ByVal mStylePropertyGrid As PropertyGridStyle)

    m_StylePropertyGrid = mStylePropertyGrid

    UserControl.PropertyChanged "StylePropertyGrid"

End Property

Private Function TextWidthU(ByVal hDC As Long, sString As String) As Long

'*************************************************************************
'* A better altenative to the VB method .TextWidth.  Thanks LaVolpe!     *
'*************************************************************************

  Dim Flags    As Long
  Dim TextRect As RECT
   
    SetRect TextRect, 0, 0, 0, 0
    Flags = DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
    DrawTextA hDC, sString, -1, TextRect, Flags
    TextWidthU = TextRect.rRight + 1

End Function

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

Public Property Let TrimPath(ByVal lTrimPath As Boolean)
    
    m_bTrimPath = lTrimPath
    
    UserControl.PropertyChanged "TrimPath"
    
End Property

Public Property Get TrimPath() As Boolean

    TrimPath = m_bTrimPath
 
End Property

Public Function TrimPathByLen(ByVal sInput As String, ByVal iTextWidth As Integer, Optional ByVal _
                              sReplaceString As String = "...", Optional ByVal sFont As String = _
                              "MS Sans Serif", Optional ByVal iFontSize As Integer = 8) As String

    '**************************************************************************
    'Function TrimPathByLen
    '
    'Inputs:
    'sInput As String :         the path to alter
    'iTextWidth as Integer :    the desired length of the inputted path in twips
    'sReplaceString as String : the string which is interted for missing text.  Default "..."
    'sFont as String :          the font being used for display.  Default "MS Sans Serif"
    'iFontSize as Integer :     the font size being used for display.  Default "8"

    'Output:
    'TrimPathByLen intellengently cuts the input (sInput) to a string that fits
    'within the desired Width.
    '
    '**************************************************************************

  Dim iInputLen As Integer
  Dim sBeginning As String
  Dim sEnd As String
  Dim aBuffer() As String
  Dim bAddedTrailSlash As Boolean
  Dim iIndex As Integer
  Dim iArrayCount As Integer  ', sBuffer As String

    'setup font attributes
    Printer.Font = sFont$
    Printer.FontSize = iFontSize%
    Printer.ScaleMode = vbTwips

    'get length of input string in twips
    iInputLen% = Printer.TextWidth(sInput$)

    'let's be reasonable here on the TextWidth
    If iTextWidth% < 200 Then Exit Function
    iTextWidth% = iTextWidth% - 400

    'make sure the desired text Width is smaller than
    'the length of the current string

    If iTextWidth < iInputLen% Then

        'now that we know how much to trim, we need to
        'determine the path type: local, network, or URL

        If InStr(1, sInput$, "\") > 0 Then 'LOCAL
            'add trailing slash if there is none

            If Right$(sInput$, 1) <> "\" Then
                bAddedTrailSlash = True
                sInput$ = sInput$ & "\"
            End If

            'throw path into an array
            aBuffer() = Split(sInput$, "\")

            If UBound(aBuffer()) > LBound(aBuffer()) Then

                iArrayCount% = UBound(aBuffer()) - 1 'the last element is blank
                sBeginning$ = aBuffer(0) & "\" & aBuffer(1) & "\"
                sEnd$ = "\" & aBuffer(iArrayCount%)

                If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + _
                    UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                    'if the total outputed string is too big then sTop
                    sBeginning$ = aBuffer(0) & "\"

                    If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) _
                        + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                        TrimPathByLen$ = sReplaceString$ & sEnd$
                    Else
                        TrimPathByLen$ = sBeginning$ & sReplaceString$ & sEnd$
                    End If

                Else

                    For iIndex% = iArrayCount% - 1 To 1 Step -1 'go throug the remaing elements to
                        '   get the best fit
                        sEnd$ = "\" & aBuffer(iIndex%) & sEnd$

                        If (UserControl.TextWidth(sBeginning$) + _
                            UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) _
                            > iTextWidth% Then
                            'if the total outputed string is too big then sTop
                            TrimPathByLen$ = sBeginning$ & sReplaceString$ & Mid$(sEnd$, _
                                Len(aBuffer(iIndex%)) + 2)
                            Exit For
                        End If

                        DoEvents
                    Next iIndex%

                End If
            Else
                'there is only one array element: bad.
                TrimPathByLen$ = sInput$
            End If

            Exit Function

        ElseIf InStr(1, sInput$, "/") > 0 Then
            If InStr(1, sInput$, ":") > 0 Then 'URL
                'start by triming off the extra params
                If InStr(1, sInput$, "?") > 0 Then sInput$ = Mid$(sInput$, 1, InStr(1, sInput$, _
                    "?") - 1)

                'add trailing slash if there is none

                If Right$(sInput$, 1) <> "/" Then
                    bAddedTrailSlash = True
                    sInput$ = sInput$ & "/"
                End If

                'throw path into an array
                aBuffer() = Split(sInput$, "/")

                If UBound(aBuffer()) > LBound(aBuffer()) Then

                    iArrayCount% = UBound(aBuffer()) - 1 'the last element is blank
                    sBeginning$ = aBuffer(0) & "/" & aBuffer(1) & "/"
                    sEnd$ = "/" & aBuffer(iArrayCount%)

                    If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) _
                        + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                        'if the total outputed string is too big then sTop
                        sBeginning$ = aBuffer(0) & "/"

                        If (UserControl.TextWidth(sBeginning$) + _
                            UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) _
                            > iTextWidth% Then
                            TrimPathByLen$ = sReplaceString$ & sEnd$
                        Else
                            TrimPathByLen$ = sBeginning$ & sReplaceString$ & sEnd$
                        End If

                    Else

                        For iIndex% = iArrayCount% - 1 To 1 Step -1 'go throug the remaing elements
                            '   to get the best fit
                            sEnd$ = "/" & aBuffer(iIndex%) & sEnd$

                            If (UserControl.TextWidth(sBeginning$) + _
                                UserControl.TextWidth(sReplaceString$) + _
                                UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                                'if the total outputed string is too big then sTop
                                TrimPathByLen$ = sBeginning$ & sReplaceString$ & Mid$(sEnd$, _
                                    Len(aBuffer(iIndex%)) + 2)
                                Exit For
                            End If

                            DoEvents
                        Next iIndex%

                    End If
                Else
                    'there is only one array element: bad.
                    TrimPathByLen$ = sInput$
                End If

            Else ' NETWORK

                'add trailing slash if there is none

                If Right$(sInput$, 1) <> "/" Then
                    bAddedTrailSlash = True
                    sInput$ = sInput$ & "/"
                End If

                'throw path into an array
                aBuffer() = Split(sInput$, "/")

                If UBound(aBuffer()) > LBound(aBuffer()) Then

                    iArrayCount% = UBound(aBuffer()) - 1 'the last element is blank
                    sBeginning$ = aBuffer(0) & "/" & aBuffer(1) & "/"
                    sEnd$ = "/" & aBuffer(iArrayCount%)

                    If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) _
                        + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                        'if the total outputed string is too big then sTop
                        sBeginning$ = aBuffer(0) & "/"

                        If (UserControl.TextWidth(sBeginning$) + _
                            UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) _
                            > iTextWidth% Then
                            TrimPathByLen$ = sReplaceString$ & sEnd$
                        Else
                            TrimPathByLen$ = sBeginning$ & sReplaceString$ & sEnd$
                        End If

                    Else

                        For iIndex% = iArrayCount% - 1 To 1 Step -1 'go throug the remaing elements
                            '   to get the best fit
                            sEnd$ = "/" & aBuffer(iIndex%) & sEnd$

                            If (UserControl.TextWidth(sBeginning$) + _
                                UserControl.TextWidth(sReplaceString$) + _
                                UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                                'if the total outputed string is too big then sTop
                                TrimPathByLen$ = sBeginning$ & sReplaceString$ & Mid$(sEnd$, _
                                    Len(aBuffer(iIndex%)) + 2)
                                Exit For
                            End If

                            DoEvents
                        Next iIndex%

                    End If
                Else
                    'there is only one array element: bad.
                    TrimPathByLen$ = sInput$
                End If

            End If
        Else
            'um, yeah.
        End If

    Else
        'we can return the value since it's already small enough
        TrimPathByLen$ = sInput$
    End If

End Function

Private Sub txtValue_Change()

On Error GoTo ChangeError

    If (lChild >= 0) And (txtValue.Visible = True) Then
        If (txtValue.Text <> ChildOrd(lChild).Value) Then
            ChildOrd(lChild).Value = txtValue.Text
            RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                ChildOrd(lChild).Value, Nothing)
        End If

    End If

    Exit Sub
    
ChangeError:

End Sub

Private Sub txtValue_GotFocus()

    txtValue.Tag = "Y"

End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)

  Dim tPoints As Variant
    
    Select Case m_Style
    Case ItemNumeric
        tPoints = Split(txtValue.Text, m_sDecimalSymbol)
        
        If (KeyAscii = 8) Then Exit Sub
        
        If ((KeyAscii < 48) Or (KeyAscii > 57)) Then
            If (KeyAscii <> 8) Then
                If (UBound(tPoints) >= 1) Then
                    KeyAscii = 0
                    Beep
                ElseIf (KeyAscii <> Asc(m_sDecimalSymbol)) Then
                    KeyAscii = 0
                    Beep
                End If
                
            End If
            
        End If
    
    Case ItemLowerCase
        
        If (KeyAscii = 8) Then Exit Sub
        
        If ((KeyAscii >= 65) And (KeyAscii <= 90)) Then
            If (KeyAscii <> 8) Or (KeyAscii <> 32) Then
                KeyAscii = 0
                Beep
            End If
            
        End If
    
    Case ItemUpperCase
        
        If (KeyAscii = 8) Then Exit Sub
        
        If ((KeyAscii >= 97) And (KeyAscii <= 122)) Then
            If (KeyAscii <> 8) Or (KeyAscii <> 32) Then
                KeyAscii = 0
                Beep
            End If
            
        End If
        
    End Select

End Sub

Public Sub TypeChildItemChanged(ByVal Title As String, _
                                ByVal TypeGridItem As PropertyItemType, _
                                Optional ByVal AutoRedraw As Boolean = False)
    
  Dim i As Long
  Dim lPos As Long

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) Then
            lPos = i
            Exit For
        End If

    Next i
    
    ChildItem(lPos).TypeGrid = TypeGridItem
    
    For i = 0 To CountChild - 1

        If (Title = ChildOrd(i).Title) Then
            lPos = i
            Exit For
        End If

    Next i
    
    ChildOrd(lPos).TypeGrid = TypeGridItem
    
    If (AutoRedraw = True) Then
        ReDraw False
    End If
    
End Sub

Private Sub UpDown_Change()
    
On Error Resume Next
    ChildOrd(lChild).Value = UpDown.Value
On Error GoTo 0
    
End Sub

Private Sub UserControl_DblClick()

On Error GoTo DblClickError
    setYet = False
    ReDraw , True
    
    If (SCmb.Visible = True) And (lstFX1.ListCount > 0) Then
        Dim lValIndex As Long
    
        lValIndex = lstFX1.FindFirst(CStr(SCmb.Text), , True)
        If (lValIndex < 0) Then lValIndex = 0
        lstFX1.ListIndex = lValIndex
        If ((lstFX1.ListIndex + 1) >= lstFX1.ListCount) Then
            lstFX1.ListIndex = 0
        Else
            lstFX1.ListIndex = lstFX1.ListIndex + 1
        End If
        
        If (lChild >= 0) Then
            Select Case ChildOrd(lChild).TypeGrid
            Case &H0 ' True/False.
                ChildOrd(lChild).Value = lstFX1.ItemText(lstFX1.ListIndex)
                RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                    ChildOrd(lChild).Value, Nothing)

            Case &HA ' StringList.
                ChildOrd(lChild).ItemValue = lstFX1.ItemText(lstFX1.ListIndex)
                RaiseEvent ValueChanged(ChildOrd(lChild).KeyCategory, ChildOrd(lChild).Title, _
                    ChildOrd(lChild).ItemValue, Nothing)
            End Select
            
            SCmb.Text = lstFX1.ItemText(lstFX1.ListIndex)
        End If
    
    ElseIf (isBttAction.Visible = True) Then
        isBttAction_Click
    Else
        ReDraw False, True, True
        If (bExpand = False) Then
            RaiseEvent DblClick
        End If
        bExpand = False
        
    End If
    
    Exit Sub
    
DblClickError:

End Sub

Private Sub UserControl_EnterFocus()

    ReDraw False, , True
    UserControl_GotFocus

End Sub

Private Sub UserControl_ExitFocus()

    'lItemSelected = -1
    lButton = -1
    SCmb.ClosedList
    bUserFocus = False
    bRedraw = False
    UserControl_LostFocus

End Sub

Private Sub UserControl_GotFocus()

    bUserFocus = True

End Sub

Private Sub UserControl_Initialize()

  Dim OS As OSVERSIONINFO

    ' Get the operating system version for text drawing purposes.
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    lItemSelected = -1
    lButton = -1
    InitCustomColors

    ' Hack for XP Crash under VB6
    m_hMod = LoadLibraryA("shell32.dll")
    InitCommonControls

End Sub

Private Sub UserControl_InitProperties()

    m_bEnabled = True
    m_bFixedSplit = False
    m_bHelpVisible = True
    m_bTrimPath = False
    m_Filters = "Supported files|*.*|All Files (*.*)"
    m_iHelpHeight = 58
    m_iSplitterPos = 0.5
    m_lAutoFilter = False
    m_lBackColor = ConvertSystemColor(GetSysColor(COLOR_BTNFACE))
    m_lHelpBackColor = ConvertSystemColor(vbButtonFace)
    m_lHelpForeColor = ConvertSystemColor(vbButtonText)
    m_lLineColor = ConvertSystemColor(&H808080)
    m_ListIndex = -1
    m_lViewBackColor = ConvertSystemColor(vbWindowBackground)
    m_lViewCategoryForeColor = ConvertSystemColor(vbButtonText)
    m_lViewForeColor = ConvertSystemColor(vbButtonText)
    m_sDecimalSymbol = "."
    m_SetColorStyle = RGB_Color
    m_StyleButton = &H6
    m_StyleComboBox = &H3
    m_StylePropertyGrid = &H0
    m_vPropertySort = Categorized
    xSplitter = m_iSplitterPos * 220
    
    With SCmb
        .AppearanceCombo = m_StyleComboBox
        .OfficeAppearance = &H3
        .XpAppearance = &H0
    End With
    
    With isBttAction
        .Style = m_StyleButton
        .NonThemeStyle = &H0
    End With

    Set m_lFont = Ambient.Font

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim lVisibleItems As Long
  Dim lNewPosition  As Long

    If (txtValue.Visible = True) And (txtValue.Tag = "Y") Then
        Exit Sub
    ElseIf (lstFX1.Visible = True) Then
        Select Case KeyCode
        Case vbKeyUp '{Up arrow}
            If (lstFX1.ListIndex > 0) Then
               lstFX1.ListIndex = lstFX1.ListIndex - 1
            End If
         
        Case vbKeyDown '{Down arrow}
            If (lstFX1.ListIndex < lstFX1.ListCount - 1) Then
               lstFX1.ListIndex = lstFX1.ListIndex + 1
            End If

        Case vbKeyPageDown '{PageDown}
            If (lstFX1.ListIndex < lstFX1.ListCount - lstFX1.VisibleRows - 1) Then
                lstFX1.ListIndex = lstFX1.ListIndex + lstFX1.VisibleRows
            Else
                lstFX1.ListIndex = lstFX1.ListCount - 1
            End If
        
        Case vbKeyPageUp '{PageUp}
            If (lstFX1.ListIndex > lstFX1.VisibleRows) Then
                lstFX1.ListIndex = lstFX1.ListIndex - lstFX1.VisibleRows
            Else
                lstFX1.ListIndex = 0
            End If
            
        Case vbKeyHome '{Start}
            lstFX1.ListIndex = 0
        
        Case vbKeyEnd '{End}
            lstFX1.ListIndex = lstFX1.ListCount - 1
        
        Case vbKeyReturn '{Enter}
            lstFX1_Click
        
        End Select
        
        Exit Sub
    End If
    
    lVisibleItems = (ScaleHeight \ (UserControl.TextHeight("A") + 5)) - 2

    Select Case KeyCode
    Case vbKeyF2 ' Action in the grid.
        If (SCmb.Visible = True) Then ' Show/Hide Combobox.

            If (lstFX1.Visible = True) Then
                SCmb.ClosedList
            Else
                SCmb_Click
            End If

        ElseIf (isBttAction.Visible = True) Then ' Open a dialog Color, Font, Folder.
            isBttAction_Click
        End If

        Exit Sub

    Case vbKeySpace, vbKeyF3 ' Expand/Collapsed.
        lNewPosition = lItemSelected
        ReDraw , True
        
        RaiseEvent DblClick
    
    Case vbKeyLeft ' Move position or (-) Collapsed.
        
        ' in wait...
        Exit Sub
    
    Case vbKeyRight ' Move position or (+) Expand.
        
        ' in wait...
        Exit Sub
    
    Case vbKeyUp '{Up arrow}

        If (lItemSelected > 1) Then
            lNewPosition = lItemSelected - 1
            If (Value > lNewPosition) And (bVScroll = True) Then Value = Value - 2
        Else
            lNewPosition = lItemSelected
            If (bVScroll = True) Then Value = 0
        End If

    Case vbKeyDown '{Down arrow}

        If (lItemSelected < lTotalItems) Then
            lNewPosition = lItemSelected + 1
            If (lNewPosition >= lVisibleItems) And (bVScroll = True) Then Value = Value + 2
        Else
            lNewPosition = lItemSelected
        End If

    Case vbKeyPageUp '{PageUp}

        If (lItemSelected > 1) Then
            lNewPosition = lItemSelected - lVisibleItems
            If (lNewPosition <= 0) Then lNewPosition = 1
            If (bVScroll = True) Then Value = lNewPosition - 1
        Else
            lNewPosition = IIf(lItemSelected = 0, 1, lItemSelected)

        End If

    Case vbKeyPageDown '{PageDown}

        If (lItemSelected < lTotalItems) Then
            lNewPosition = lItemSelected + lVisibleItems
            If (lNewPosition > lTotalItems) Then lNewPosition = lTotalItems
            If (bVScroll = True) Then Value = lNewPosition - 1
        Else
            lNewPosition = lItemSelected
        End If

    Case vbKeyHome '{Start}
        lNewPosition = 1
        If (bVScroll = True) Then Value = 0

    Case vbKeyEnd '{End}
        lNewPosition = lTotalItems
        If (bVScroll = True) Then Value = Max + 1

    Case Else
        Exit Sub
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)

    lItemSelected = lNewPosition

    ReDraw , , True

End Sub

Private Sub UserControl_LostFocus()

    ReDraw
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    bRedraw = False
    bUserFocus = True
    SCmb.Visible = False
    SCmb.ClosedList
    isBttAction.Visible = False
    txtValue.Visible = False
    McCalendar.Visible = False
    
    ' Calculate row from mouse 'Y'.
    m_iHeight = UserControl.TextHeight("A") + 5
    If (Value1 > 0) Then Value1 = Value1 - 1

    If (Button <> vbLeftButton) Then
        lItemSelected = Value1 + Int(Y / m_iHeight) + 1
        RaiseEvent RightClick(lItemSelected)
        Exit Sub
    End If

    SetCapture UserControl.hWnd
    lButton = Button

    lXPos = X
    lYPos = Y
    setYet = False

    If (((X - 1) + SIZE_VARIANCE) = (xSplitter + SIZE_VARIANCE)) Then
        If ((Y < (UserControl.ScaleHeight - xSplitterY - 2)) And (m_bHelpVisible = True)) Or _
            (m_bHelpVisible = False) Then
            
            bResize = True
            lItemSelected = lLastItemSelected
        End If
        
    ElseIf ((Y < (UserControl.ScaleHeight - xSplitterY - 2)) And (m_bHelpVisible = True)) Or _
            (m_bHelpVisible = False) Then
        
        lItemSelected = Value1 + Int(Y / m_iHeight) + 1
        lDblClick = False
        
        If (X < 17) Then
            lDblClick = True
        End If
        
        ReDraw , lDblClick, lDblClick
        lLastItemSelected = lItemSelected
                
        If (SCmb.Visible = True) Then
            NoShowList = True
            SCmb_Click
            NoShowList = False
        ElseIf (lDblClick = False) Then
            RaiseEvent Click
        End If
        
    ElseIf (m_bHelpVisible = True) And (((Y - 1) + SIZE_VARIANCE) = ((UserControl.ScaleHeight - (xSplitterY - 1)) + SIZE_VARIANCE)) Then
        bResize = True
        lItemSelected = lLastItemSelected
    Else
        lLastItemSelected = -1
        ReDraw
    End If

    lButton = -1

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo MouseMoveError
    If (bResize = False) And (m_bFixedSplit = False) Then
        If (((X - 1) + SIZE_VARIANCE) = (xSplitter + SIZE_VARIANCE)) Then
             UserControl.MousePointer = 99
             UserControl.MouseIcon = getResourceCursor(CURSOR_ARROW_VERTICAL_SPLITTER)
             bResY = False
        ElseIf (((Y - 1) + SIZE_VARIANCE) = ((UserControl.ScaleHeight - (xSplitterY - 1)) + SIZE_VARIANCE)) And (m_bHelpVisible = True) Then
             UserControl.MousePointer = 99
             UserControl.MouseIcon = getResourceCursor(CURSOR_ARROW_HORIZONTAL_SPLITTER)
             bResY = True
        Else
            UserControl.MousePointer = vbDefault
            Set UserControl.MouseIcon = Nothing
        End If
        
    ElseIf (Button = vbLeftButton) And (m_bFixedSplit = False) Then
        ' We are resizing the items.
        If (bResY = False) Then
            If (X > lXPos) Then
                If (xSplitter < UserControl.ScaleWidth - 75) Then
                    xSplitter = xSplitter + (X - lXPos)
                    If (xSplitter > UserControl.ScaleWidth) Then
                        xSplitter = UserControl.ScaleWidth
                    End If
                
                End If

            ElseIf (xSplitter > 30) Then
                xSplitter = xSplitter - (lXPos - X)
                If (xSplitter < 30) Then
                    xSplitter = 30
                End If
            
            End If
            
            lXPos = X
        Else ' Resize help ^/^;
            
            If (Y < lYPos) Then
                If (xSplitterY < (UserControl.ScaleHeight - 35)) Then
                    xSplitterY = xSplitterY + (lYPos - Y)
                    If (xSplitterY > UserControl.ScaleHeight - 35) Then
                        xSplitterY = UserControl.ScaleHeight - 35
                    End If
                    
                End If
                
            ElseIf (xSplitterY > UserControl.TextHeight("A") + 25) Then
                    xSplitterY = xSplitterY - (Y - lYPos)
                    If (xSplitterY < UserControl.TextHeight("A") + 25) Then
                        xSplitterY = UserControl.TextHeight("A") + 25
                    End If
                    
            End If
            
            lYPos = Y
        End If
        ' Changed bResize = False when redraw the control.
        bResize = False
        b_Focus = True
        ReDraw False
        ' Put again resize mode for prevent Error #7 Out of memory.
        bResize = True
        UserControl.Refresh
    Else
        UserControl.MousePointer = vbDefault
        Set UserControl.MouseIcon = Nothing
    End If
    
    Exit Sub
    
MouseMoveError:

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (Button = vbLeftButton) Then
        ReleaseCapture
        bResize = False
        UserControl.MousePointer = vbDefault
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag

        AutoFilter = .ReadProperty("AutoFilter", False)
        BackColor = .ReadProperty("BackColor", ConvertSystemColor(GetSysColor(COLOR_BTNFACE)))
        DecimalSymbol = .ReadProperty("DecimalSymbol", ".")
        Enabled = .ReadProperty("Enabled", True)
        FixedSplit = .ReadProperty("FixedSplit", False)
        HelpBackColor = .ReadProperty("HelpBackColor", ConvertSystemColor(vbButtonFace))
        HelpForeColor = .ReadProperty("HelpForeColor", ConvertSystemColor(vbButtonText))
        HelpHeight = .ReadProperty("HelpHeight", "58")
        HelpVisible = .ReadProperty("HelpVisible", True)
        LineColor = .ReadProperty("LineColor", ConvertSystemColor(&H808080))
        PropertySort = .ReadProperty("PropertySort", Categorized)
        Set m_lFont = .ReadProperty("Font", Ambient.Font)
        SetColorStyle = .ReadProperty("SetColorStyle", RGB_Color)
        SplitterPos = .ReadProperty("SplitterPos", 0.5)
        StyleButton = .ReadProperty("StyleButton", &H6)
        StyleComboBox = .ReadProperty("StyleComboBox", &H3)
        StylePropertyGrid = .ReadProperty("StylePropertyGrid", &H0)
        TrimPath = .ReadProperty("TrimPath", False)
        ViewBackColor = .ReadProperty("ViewBackColor", ConvertSystemColor(vbWindowBackground))
        ViewCategoryForeColor = .ReadProperty("ViewCategoryForeColor", _
            ConvertSystemColor(vbButtonText))
        ViewForeColor = .ReadProperty("ViewForeColor", ConvertSystemColor(vbButtonText))
        xSplitter = m_iSplitterPos * 220

    End With

    If Ambient.UserMode Then                                                              'If we're
        ' Not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If

        End If

        If bTrack Then
            ' OS supports mouse leave, so let's subclass for it
            If (m_hWnd = 0) Then DrawScrollBar

            With UserControl
                ' Subclass the UserControl
                sc_Subclass .hWnd
                sc_AddMsg .hWnd, BFFM_INITIALIZED
                sc_AddMsg .hWnd, WM_MOUSEWHEEL
                sc_AddMsg .hWnd, WM_MOUSEMOVE
                sc_AddMsg .hWnd, WM_MOUSELEAVE
                sc_AddMsg .hWnd, WM_CTLCOLORSCROLLBAR
                sc_AddMsg .hWnd, WM_VSCROLL
                sc_AddMsg .hWnd, WM_SYSCOLORCHANGE

                If (isXp = True) Then
                    sc_AddMsg .hWnd, WM_THEMECHANGED
                End If

            End With
            ' Subclass the parent form.

            With Extender.Parent
                sc_Subclass .hWnd
                sc_AddMsg .hWnd, WM_MOUSEMOVE
                sc_AddMsg .hWnd, WM_MOUSELEAVE
                sc_AddMsg .hWnd, WM_MOVING
                sc_AddMsg .hWnd, WM_SIZING
                sc_AddMsg .hWnd, WM_EXITSIZEMOVE
                sc_AddMsg .hWnd, WM_LBUTTONDOWN
                sc_AddMsg .hWnd, WM_RBUTTONDOWN
                sc_AddMsg .hWnd, WM_MBUTTONDOWN
                sc_AddMsg .hWnd, WM_ACTIVATE
                sc_AddMsg .hWnd, WM_NCLBUTTONDOWN
            End With
            bSubClass = True
        End If
    End If

End Sub

Private Sub UserControl_Resize()

    If (Ambient.UserMode = False) Then
        Dim lHg As Long
    
        UserControl.Cls
        UserControl.BackColor = m_lViewBackColor
        lHg = 58
        
        ' Draw the external border of the control.
        APIRectangle UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - lHg, _
            m_lLineColor

        ' Draw the Rectangle if the Help is True.
    On Error Resume Next
        DrawRectangleBorder UserControl.hDC, 0, UserControl.ScaleHeight - lHg + 1, _
            UserControl.ScaleWidth, UserControl.ScaleHeight - (UserControl.ScaleHeight - lHg + _
            5) - 1, ConvertSystemColor(Extender.Parent.BackColor), False
        DrawRectangleBorder UserControl.hDC, 0, UserControl.ScaleHeight - lHg + 5, _
            UserControl.ScaleWidth, UserControl.ScaleHeight - (UserControl.ScaleHeight - lHg + _
            5) - 1, m_lHelpBackColor, False
        APIRectangle UserControl.hDC, 0, UserControl.ScaleHeight - lHg + 5, UserControl.ScaleWidth _
            - 1, UserControl.ScaleHeight - (UserControl.ScaleHeight - lHg + 5) - 1, m_lLineColor
    On Error GoTo 0
    
    End If

End Sub

Private Sub UserControl_Show()

  Dim lResult As Long

On Error Resume Next
    lResult = GetWindowLong(lstFX1.hWnd, GWL_EXSTYLE)
    SetWindowLong lstFX1.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW
    SetWindowPos lstFX1.hWnd, lstFX1.hWnd, -1, 0, 0, 0, 2 Or 1
    SetWindowLong lstFX1.hWnd, -8, Parent.hWnd
    SetParent lstFX1.hWnd, 0
    
    lResult = GetWindowLong(McCalendar.hWnd, GWL_EXSTYLE)
    SetWindowLong McCalendar.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW
    SetWindowPos McCalendar.hWnd, McCalendar.hWnd, -1, 0, 0, 0, 2 Or 1
    SetWindowLong McCalendar.hWnd, -8, Parent.hWnd
    SetParent McCalendar.hWnd, 0
    
    With lstFX1
        .ItemHeightAuto = True
        .ScrollBarWidth = 18
        .HoverSelection = False
        .WordWrap = False
    End With
    
    Min = 0
    SmallChange = 5
    LargeChange = 5
    
    If (Ambient.UserMode = True) Then
        ReDraw False
    End If
On Error GoTo 0

End Sub

Private Sub UserControl_Terminate()

    Erase Category
    Erase ChildItem
    Erase ChildOrd
        
    ' The control is terminating - a good place to sTop the subclasser
    sc_Terminate ' Terminate all subclassing

    If Not (m_hMod = 0) Then
        FreeLibrary m_hMod
        m_hMod = 0
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag

        .WriteProperty "AutoFilter", m_lAutoFilter, False
        .WriteProperty "BackColor", m_lBackColor, ConvertSystemColor(GetSysColor(COLOR_BTNFACE))
        .WriteProperty "DecimalSymbol", m_sDecimalSymbol, "."
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "FixedSplit", m_bFixedSplit, False
        .WriteProperty "Font", m_lFont, Ambient.Font
        .WriteProperty "HelpBackColor", m_lHelpBackColor, ConvertSystemColor(vbButtonFace)
        .WriteProperty "HelpForeColor", m_lHelpForeColor, ConvertSystemColor(vbButtonText)
        .WriteProperty "HelpHeight", m_iHelpHeight, "58"
        .WriteProperty "HelpVisible", m_bHelpVisible, True
        .WriteProperty "LineColor", m_lLineColor, ConvertSystemColor(vbButtonText)
        .WriteProperty "PropertySort", m_vPropertySort, Categorized
        .WriteProperty "SetColorStyle", m_SetColorStyle, RGB_Color
        .WriteProperty "SplitterPos", m_iSplitterPos, 0.5
        .WriteProperty "StyleButton", m_StyleButton, &H6
        .WriteProperty "StyleComboBox", m_StyleComboBox, &H3
        .WriteProperty "StylePropertyGrid", m_StylePropertyGrid, &H0
        .WriteProperty "TrimPath", m_bTrimPath, False
        .WriteProperty "ViewBackColor", m_lViewBackColor, ConvertSystemColor(vbWindowBackground)
        .WriteProperty "ViewCategoryForeColor", m_lViewCategoryForeColor, _
            ConvertSystemColor(vbButtonText)
        .WriteProperty "ViewForeColor", m_lViewForeColor, ConvertSystemColor(vbButtonText)

    End With

End Sub

Private Property Let Value(ByVal iValue As Long)

  Dim tSI As SCROLLINFO
  Dim lPercent As Long

    If (iValue <> Value) And (bResize = False) Then
        tSI.nPos = iValue
        pLetSI tSI, SIF_POS
        If (Max > 0) Then lPercent = iValue * 100 \ Max
        pRaiseEvent False
    End If

End Property

Private Property Get Value() As Long

  Dim tSI As SCROLLINFO

    If (bResize = False) Then
        pGetSI tSI, SIF_POS
        Value = tSI.nPos
    End If

End Property

Public Sub ValueChildItemChanged(ByVal Title As String, _
                                 ByVal KeyCategory As String, _
                                 ByVal Value As Variant, _
                                 Optional ByVal ItemValue As String = vbNullString, _
                                 Optional ByVal AutoRedraw As Boolean = True)
     
  Dim i As Long
  Dim lPos As Long
  Dim lPos1 As Long
  Dim lOk As Boolean
  Dim lOk1 As Boolean

    For i = 0 To CountChild - 1

        If (Title = ChildItem(i).Title) And (KeyCategory = ChildItem(i).KeyCategory) Then
            lPos = i
            lOk = True
            Exit For
        End If

    Next i
    
    For i = 0 To CountChild - 1

        If (Title = ChildOrd(i).Title) And (KeyCategory = ChildOrd(i).KeyCategory) Then
            lPos1 = i
            lOk1 = True
            Exit For
        End If

    Next i
    
    If (lOk = False) And (lOk1 = False) Then
        Exit Sub
    End If
     
    If (LenB(ItemValue) > 0) Then
        ChildOrd(lPos1).ItemValue = ItemValue
        ChildItem(lPos).ItemValue = ItemValue
    End If
    
    If (ChildItem(lPos).TypeGrid = PropertyItemFont) Or (ChildItem(lPos).TypeGrid = PropertyItemPicture) Then
        Set ChildItem(lPos).Value = Value
        Set ChildOrd(lPos1).Value = Value
    Else
        ChildItem(lPos).Value = Value
        ChildOrd(lPos1).Value = Value
    End If
    
    If (AutoRedraw = True) Then
        ReDraw False
    End If
    
End Sub

Public Property Let ViewBackColor(ByVal lViewBackColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : ViewBackColor
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lViewBackColor = ConvertSystemColor(lViewBackColor)
    
    lstFX1.BackNormal = m_lViewBackColor
    SCmb.BackColor = m_lViewBackColor
    txtValue.BackColor = m_lViewBackColor
    UpDown.BackColor = m_lViewBackColor
    
    With McCalendar
        .CalendarBackCol = m_lViewBackColor
        .HeaderBackCol = m_lViewBackColor
        .MonthBackCol = m_lViewBackColor
        .YearBackCol = m_lBackColor
        .WeekDayCol = m_lBackColor
    End With

    UserControl.PropertyChanged "ViewBackColor"

    If (Ambient.UserMode = False) Then
        UserControl_Resize
    End If

End Property

Public Property Get ViewBackColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : ViewBackColor
    ' DateTime  : 20/11/2006 21:03
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    ViewBackColor = m_lViewBackColor

End Property

Public Property Let ViewCategoryForeColor(ByVal lViewCategoryForeColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : ViewCategoryForeColor
    ' DateTime  : 25/11/2006 11:30
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lViewCategoryForeColor = lViewCategoryForeColor

    UserControl.PropertyChanged "ViewCategoryForeColor"

End Property

Public Property Get ViewCategoryForeColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : ViewCategoryForeColor
    ' DateTime  : 25/11/2006 11:30
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    ViewCategoryForeColor = m_lViewCategoryForeColor

End Property

Public Property Let ViewForeColor(ByVal lViewForeColor As OLE_COLOR)

    '---------------------------------------------------------------------------------------
    ' Procedure : ViewForeColor
    ' DateTime  : 25/11/2006 11:30
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    m_lViewForeColor = lViewForeColor
    SCmb.NormalColorText = m_lViewForeColor
    txtValue.ForeColor = m_lViewForeColor
    isBttAction.FontColor = m_lViewForeColor
    McCalendar.ForeColor = m_lViewForeColor
    UpDown.ForeColor = m_lViewForeColor
    With lstFX1
        .FontNormal = m_lViewForeColor
        .FontSelected = m_lViewForeColor
    End With

    UserControl.PropertyChanged "ViewForeColor"

End Property

Public Property Get ViewForeColor() As OLE_COLOR

    '---------------------------------------------------------------------------------------
    ' Procedure : ViewForeColor
    ' DateTime  : 25/11/2006 11:30
    ' Author    : HACKPRO TM
    ' Purpose   :
    '---------------------------------------------------------------------------------------

    ViewForeColor = m_lViewForeColor

End Property

Public Sub XpAppearance(ByVal lAppearance As ComboXpAppearance)

    SCmb.XpAppearance = lAppearance

End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)

    '-The following routines are exclusively for the sc_ subclass
    '   routines----------------------------
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

    'Return the address of the specified ordinal method on the oCallback object, 1 = last private
    '   method, 2 = second last private method, etc

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
    'MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & _
        sRoutine

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

    '-Subclass callback, usually ordinal #1, the last method in this source
    '   file----------------------

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
  Dim lSC As Long
  Dim lScrollcode As Long
  Dim tSI As SCROLLINFO
  Dim lV     As Long

    ' If resize mode for prevent Error #7 Out of memory I exit of Sub.
    If (bResize = True) Then Exit Sub
    
    Select Case lng_hWnd
    Case UserControl.hWnd

        Select Case uMsg
        Case BFFM_INITIALIZED
            '   BrowseForFolder Module has Initialized, so set the Starting Path
            Call SendMessage(lng_hWnd, BFFM_SETSELECTIONA, True, ByVal m_Path)

        Case WM_MOUSEMOVE

            If Not bInCtrl Then
                bInCtrl = True
                TrackMouseLeave lng_hWnd
                RaiseEvent MouseEnter
            End If

        Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
            ReDraw

        Case WM_MOUSELEAVE
            bInCtrl = False
            Set UserControl.MouseIcon = Nothing
            UserControl.MousePointer = vbNormal
            RaiseEvent MouseLeave

        Case WM_MOUSEWHEEL
                        
            If (McCalendar.Visible = False) And (lstFX1.Visible = False) And (bVScroll = True) Then
                If (wParam = &H780000) Then
                    Value = Value - 1
                ElseIf (wParam = &HFF880000) Then
                    Value = Value + 1
                End If

                b_Focus = True
                ReDraw
            End If

        Case WM_VSCROLL '* Steven
            UserControl.MousePointer = vbDefault
            Set UserControl.MouseIcon = Nothing
            
            lScrollcode = (wParam And &HFFFF&)

            Select Case lScrollcode
            Case SB_THUMBTRACK
                '* Is vertical/horizontal?
                pGetSI tSI, SIF_TRACKPOS
                Value = tSI.nTrackPos
                pRaiseEvent True

            Case SB_LEFT, SB_BOTTOM
                Value = Min
                pRaiseEvent False

            Case SB_RIGHT, SB_TOP
                Value = Max
                pRaiseEvent False

            Case SB_LINELEFT, SB_LINEUP
                lV = Value
                lSC = m_lSmallChange

                If (lV - lSC <= Min) Then
                    Value = Min
                Else
                    Value = lV - lSC
                End If

                pRaiseEvent False

            Case SB_LINERIGHT, SB_LINEDOWN
                lV = Value
                lSC = m_lSmallChange

                If (lV + lSC >= Max) Then
                    Value = Max + 1
                Else
                    Value = lV + lSC
                End If

                pRaiseEvent False

            Case SB_PAGELEFT, SB_PAGEUP
                Value = Value - LargeChange
                pRaiseEvent False

            Case SB_PAGERIGHT, SB_PAGEDOWN
                Value = Value + LargeChange
                pRaiseEvent False

            Case SB_ENDSCROLL
                pRaiseEvent False
             
            End Select
            
            Value1 = Value
            lstFX1.Visible = False
            McCalendar.Visible = False
            SCmb.ClosedList
            b_Focus = True
            ReDraw
            
        End Select

    Case Extender.Parent.hWnd

        Select Case uMsg
        Case WM_MOUSELEAVE, WM_ACTIVATE
            SCmb.Visible = False
            McCalendar.Visible = False
            lstFX1.Visible = False
            
        Case WM_MOVING, WM_SIZING, WM_EXITSIZEMOVE, WM_RBUTTONDOWN, _
             WM_MBUTTONDOWN, WM_LBUTTONDOWN, WM_NCLBUTTONDOWN

            SCmb.ClosedList
            txtValue.Visible = False
            isBttAction.Visible = False
            bInCtrl = False
            
        End Select

    End Select

End Sub
