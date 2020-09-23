VERSION 5.00
Begin VB.UserControl McCalendar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   FillColor       =   &H00257A4B&
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "McCalendar.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   172
   ToolboxBitmap   =   "McCalendar.ctx":030A
End
Attribute VB_Name = "McCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^GTECH CREATIONS^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^¶¶^^^^^¶¶^^^^^^^^¶¶¶¶¶^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^¶¶^^^^¶¶¶¶^^^$
'$^^¶¶¶^^^¶¶¶^^^^^^^¶¶^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^¶¶¶^^^^^^^¶¶¶^^^¶¶^^^^^^$
'$^^¶¶¶¶^¶¶¶¶^^¶¶¶¶^¶¶^^^^^^¶¶¶¶^^¶¶^^¶¶¶¶^^¶¶¶¶¶^^^¶¶¶¶¶^^¶¶¶¶^^¶¶^¶^^^^^^¶¶^^^^^^^^¶¶^^^¶¶^^^^^^$
'$^^¶^¶¶¶¶^¶¶^¶¶^^^^¶¶^^^^^^^^^¶¶^¶¶^¶¶^^¶¶^¶¶^^¶¶^¶¶^^¶¶^^^^^¶¶^¶¶¶¶^^^^^^¶¶^^^^^^^^¶¶^^^¶¶¶¶¶^^^$
'$^^¶^^¶¶^^¶¶^¶¶^^^^¶¶^^^^^^¶¶¶¶¶^¶¶^¶¶¶¶¶¶^¶¶^^¶¶^¶¶^^¶¶^^¶¶¶¶¶^¶¶^^^^^^^^¶¶^^^^^^^^¶¶^^^¶¶^^¶¶^^$
'$^^¶^^^^^^¶¶^¶¶^^^^¶¶^^^^^¶¶^^¶¶^¶¶^¶¶^^^^^¶¶^^¶¶^¶¶^^¶¶^¶¶^^¶¶^¶¶^^^^^^^^¶¶^^^^^^^^¶¶^^^¶¶^^¶¶^^$
'$^^¶^^^^^^¶¶^¶¶^^^^¶¶^^^^^¶¶^^¶¶^¶¶^¶¶^^^^^¶¶^^¶¶^¶¶^^¶¶^¶¶^^¶¶^¶¶^^^^^^^^¶¶^^^¶¶^^^¶¶^^^¶¶^^¶¶^^$
'$^^¶^^^^^^¶¶^^¶¶¶¶^^¶¶¶¶¶^^¶¶¶¶¶^¶¶^^¶¶¶¶¶^¶¶^^¶¶^^¶¶¶¶¶^^¶¶¶¶¶^¶¶^^^^^^^¶¶¶¶^^¶¶^^¶¶¶¶^^^¶¶¶¶^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^CODED BY : JIM JOSE^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'--------------------------------------------------------------------------------------------------
' Source Code   : McCalendar ActiveX Control
' Auther        : Jim Jose
' eMail         : jimjosev33@yahoo.com
' Purpose       : An Ownerdrawn, Sizable, Stylish, List/DropDown Calendar
' BasicWorking  : The PopDown mode 'Exports' the'm_picCalendar' to the parent form
'               : using 'SetParent' API call and subclass it to receive events
' Last Updated  : 01-07-2005
'---------------------------------------------------------------------------------------------------
'
'About The Control
'-----------------
'       I coded this control as a complete solution for fully customizable
' and resizable calendar control. Normally all the Calendar controls we can
' found in PSC are a huge stack of controls and control arrays, and thus the
' control size will be extreamly high. I used only one picturebox for building
' this control. I think it could be justified because we need an hWnd container
' to export the calendar to the parent form.
'
'       The calendar is Owner draw which enables it to draw calenders of any size
' and it also offers high performance and speed.
'
'       The calender is provided with two MODES. You can use this control as a
' 'dropdown' calendar and as a 'Listed' calendar. The calendar colors are fully
' customizable and also provided with eight different themes.
'
'       A unique animation technique is used in this control(i already posted that
' a an another submission and the response was inspiring ).
'
'Upgrade Destinations
'--------------------
'       I think the control is error free now. If there is any hidden bugs or
' any operatingtime errors please inform me, so that I can correct it soon.
'
'       If you need any aditional properties/options feel free to inform me.
' You can also use the email, if you have any doubts(even if the code is fully
' commented)


'Thanks
'------

'   Paul Caton      : for his amazing self subclassing work
'   Carles P.V.     : for his fast gradient routine
'   Dana Seaman     : for his unicode supporting DrawText code
'   Duncan          : for his calendar positioning technique
'   Ken Foster      : for the orginal idea (not rented a bit of code)
'   Ben Vonk        : for the orginal idea (not rented a bit of code)
'   PSC             : I learned fully from there

'History:
'--------

' # Vesion 1.01
' Submitted to PSC 15-6-2005
'
' # Update Vesion 1.02
' Resubmitted the control with full range of color options.
' Improved themes.
'
' # Updated Version 1.03
' Added additional features sugested by Ken Forster. This verion have
' three modes 1.List 2.PopUp 3.PopDown. Also clicking on Today Region
' will set calenar back to current day.
'
' # Updated Version 1.04
' Added additional features sugested by Ruturag. Added two more
' properties, 'Sensitive'(for popup/down mode) and 'SkipEnabled'.
' The calendar will close when clicking on cross-filled dates if
' 'Sensitive' is true. The popdown/up arrow will reverse (as Guturaj
' suggested) only if 'Sensitive' is false.
' If 'SkipEnabled' is true, then the calendar will skip into next or
' last month (according to the back or end of days clicked)
'
' # Updated Version 1.05 (Modified in Ver 1.13)
' Added DatePicker Compatable mode as suggested by Dennis.
' This version have two more modes. 1.Datepicker PopUp
' 2. DatePicker PopDown. These two modes only show the
' left popDown buttons. You can place it near the TextBox
' to which the dates must send (see the sample).
' I used this method to get the functionality, bcose otherwise
' we needs to add a additional textbox into the control only
' for this purpose. The usercontrol is not resized to the
' popDown button size. This is bcose u can use the Usercontrol
' width to adjust the calendar width, otherwise  u may need to use a
' property, the earlier is more sensible.
' This Version also contains one more property 'Header Height', which is
' needed to adjust the header height to the DatePicker's TextBox Height
'
' # Updated Version 1.06
' Added one more event function. DbClick on Days will close the Calendar
' (except ListMode)
'
' # Updated Version 1.07 (Modified in Ver 1.10)
' Added Property special days. You can add days to this property as shon bellow
' 21-1,26-1,19-2,8-3,24-3,25-3,14-4,21-4,13-8,15-8,20-8,26-8,14-9,15-9,16-9,17-9,21-9,11-10,12-10,1-11,3-11,10-12,25-12
' This is the complete list of holydays for my country(India)(some of them are only for my state)
' The days are added as 'day-month' and seperated by ",". These days are applicable to all
' years. The Special days will be indicated by a Rectangle halfly-filled on it's cell.
' See JAN - 26 (republic day)
'
' # Updated Version 1.08
' Language problem solved. Now the calendar will load the MonthNames and the WeekDayNames
' according to the current language selection of the user. Thanks to Cote for his attension on this part.
'
' # Updated Version 1.09
' Highly worked version. 'FirstDayOfWeek' Property added, which realy required a long time fix.
' Thanks to 'Tassilo' for this suggesssion.
'
' # Updated Version 1.10
' Worked more on 'Special Days'
'       >In this version you can also view 'Speciality' of the day in ToolTip
' [Check tooltip for Dec 25, X'Mas]
'       >The new format it adding the 'SpecialDays' information as follows... (it is very simple!)
' 25-12>X'mas,26-1>Republic Day,19-2>Muharam,8-3>ShivRatri,25-3>Good Friday,14-4>Vishu,15-8>Indipendance Day,15-9>Onam,1-11>Deepavali,3-11>Rumsaan
' [It is recomented that u build this string in NotePad and Copy to property window]
'       >The 'Special Days' will be indicated by a 'Downward arrow'
' draw on the top-right of the cell.
'
' # Updated Version 1.11
' Added property 'DaySelRectangle' which draws a Rounded rectangle for selected
' day cell. Implimented LoadLibrary API. Also added MonthName for day tooltip text
'
' # Updated Version 1.12
' Highly upgraded version with unicode support.
' This version of McCalendar will support unicode languages like Chines, Japanese etc
' Thanks to 'Dana Seaman' for his support in solving this and for his excellent code which
' is directly implimented to the control. His great contribution realy raised the range of this
' control. Thanks again
'
' # Updated Version 1.13
' More work on 'DatePicker' operation.
' The constrained datepicker options are removed from Mode Enum.
' The new technique to show the datepicker is 1.Hide the header 2.Stimulate Popup Externally
' This way is more suitable in specifing the calendar position
' To hide the header a new propery 'HeaderVisible' is added. The new Sub 'PopUpcalendar'
' can stimulate the calendar Pop up or down according to mode selected.
'
' # Updated Version 1.14
' Added two more properties 'BorderColor' and 'ForeColor' as suggested by 'Dennis'
'
' # Updated Version 1.15
' Updated with better calendar positioning technique using SetwindowPos API,
' Code based on Duncan_DatePicker. Thanks to Duncan for sharing that code, which
' realy made the calendar positioning more easy and accurate.
'
' # Updated Version 1.16 [ I think, <THE LAST VERSION> ]
' After all these long revisions and updating, I think now the time to
' advance to the dead-line( the SUBCLASSING )
' This version is 99% more recomented, bcose it can now operate even external
' to the form with amazing/crash-free subclassing by Paul caton. The parent form
' is subclassed to get the moving and sizing events
'
' # Updated Version 1.17 [ Ooooooooops ]
' Realy sorry for the endless updating. Actualy necessary upgrade and new ideas
' forced me to do this.
' 1) Vb's tooltip will not support Unicode languages.( thanks to Dana Seaman for informing this )
'    So I made a simple unicode supportng tooltips localy - See 'ShowToolTip'
' 2) See the usercontrol design mode, there is no picCalendar.
'    Needed controls are added dynamically. See 'CreateControls'
' [ The tooltip code is extreamly simpler than the same with creating new window ]
'
'---------------------------------------------------------------------------------------------------
' THANKS TO ALL THE CODERS WHO'S COMMENTS AND SUGGESSIONS MADE THIS CONTROL REALY USEFUL & UNIQUE
'---------------------------------------------------------------------------------------------------
' VOTES & SUGGESSIONS : APPRECIATED,       COMMENTS : ALWAYS WELCOMED
'---------------------------------------------------------------------------------------------------

Option Explicit

'[Apis]
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

' for subclassing
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'[Enums]
Public Enum GradientDirectionCts
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_Vertical] = 2
    [Fill_DownwardDiagonal] = 3
    [Fill_UpwardDiagonal] = 4
End Enum

Public Enum AppearanceConstants
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum BorderStyleEnum
    [None] = 0
    [Fixel Single] = 1
End Enum

Public Enum CalendarMode
    [List Mode] = 0
    [PopDown Mode] = 1
    [PopUp Mode] = 2
End Enum

Public Enum CalendarTheme
    [Cal_Attraction] = 1
    [Cal Blue] = 2
    [Cal Green] = 3
    [Cal Orange] = 4
    [Cal Purple] = 5
    [Cal Red] = 6
    [Cal Silver] = 7
    [Cal Yellow] = 8
End Enum

Private Enum ArrowDir
    [Arw_Left] = 0
    [Arw_Right] = 1
    [Arw_Up] = 2
    [Arw_Down] = 3
End Enum

Public Enum FirstDayOfWeekEnum
    [SunDay] = 1
    [MonDay] = 2
    [TuesDay] = 3
    [WednesDay] = 4
    [ThursDay] = 5
    [FriDay] = 6
    [SaturDay] = 7
End Enum

Private Enum AnimeEventEnum
    aUnload = 0
    aload = 1
End Enum

Private Enum AnimeEffectEnum
    eAppearFromLeft = 0
    eAppearFromRight = 1
    eAppearFromTop = 2
    eAppearFromBottom = 3
End Enum

Public Enum DateFormatEnum
    [dd-mm-yyyy] = 0
    [mm-dd-yyyy] = 1
    [yyyy-mm-dd] = 2
End Enum

' for subclassing
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

'[Types]
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

' For gardient fill
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

' Check for OS
Private Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformId         As Long
   szCSDVersion         As String * 128 ' Maintenance string
End Type

Private Type tSubData                                                                   'Subclass data type
    hWnd          As Long                                            'Handle of the window being subclassed
    nAddrSub      As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig     As Long                                            'The address of the pre-existing WndProc
    nMsgCntA      As Long                                            'Msg after table entry count
    nMsgCntB      As Long                                            'Msg before table entry count
    aMsgTblA()    As Long                                            'Msg after table array
    aMsgTblB()    As Long                                            'Msg Before table array
End Type
                                
Private Type TRACKMOUSEEVENT_STRUCT
  cbSize          As Long
  dwFlags         As TRACKMOUSEEVENT_FLAGS
  hWndTrack       As Long
  dwHoverTime     As Long
End Type

'[Property Variables:]
Private m_SelYear           As Long
Private m_SelMonth          As Long
Private m_SelDay            As Long
Private m_Appearance        As Integer
Private m_BorderStyle       As Integer
Private m_Enabled           As Boolean
Private m_Font              As Font
Private m_Theme             As CalendarTheme
Private m_Animate           As Boolean
Private m_Mode              As CalendarMode
Private m_CalendarHeight    As Long
Private m_Curvature         As Long
Private m_CalendarGradient  As Boolean
Private m_CalendarBackCol   As OLE_COLOR
Private m_SkipEnabled       As Boolean
Private m_Sensitive         As Boolean
Private m_SpecialDays       As String

Private m_TrackWidth        As Long
Private m_iHeight           As Double
Private m_iWidth            As Double
Private m_HeaderHeight      As Long
Private m_WeekDaysHeight    As Long
Private m_Right2Left        As Boolean
Private m_MouseX            As Single
Private m_MouseY            As Single
Private m_FirstDay          As Long
Private m_MonthDays         As Long
Private m_MonthPopupMode    As Boolean
Private m_MonthPopWidth     As Double
Private m_Poped             As Boolean
Private m_hMode             As Long
Private m_SpecialDayStack() As String
Private m_SpecialDayString  As String
Private m_bIsNT             As Boolean

Private m_PicToolTip        As Object
Private m_PicCalendar       As Object
Private m_TimerElsp         As Long
Private m_ToolTipText       As String

Private m_MonthBackCol      As OLE_COLOR
Private m_DayCol            As OLE_COLOR
Private m_DaySunCol         As OLE_COLOR
Private m_WeekDaySunCol     As OLE_COLOR
Private m_WeekDayCol        As OLE_COLOR
Private m_ArrowCol          As OLE_COLOR
Private m_WeekDaySelCol     As OLE_COLOR
Private m_DaySelCol         As OLE_COLOR
Private m_DateFormat        As DateFormatEnum
Private m_YearBackCol       As OLE_COLOR
Private m_YearGradient      As Boolean
Private m_YearGradientCol   As OLE_COLOR
Private m_HeaderGradientCol As OLE_COLOR
Private m_MonthGradientCol  As OLE_COLOR
Private m_CalendarGradientCol As OLE_COLOR
Private m_MonthGradient     As Boolean
Private m_HeaderGradient    As Boolean
Private m_HeaderBackCol     As OLE_COLOR
Private m_ForeColor         As OLE_COLOR
Private m_BorderColor       As OLE_COLOR
Private m_HeaderVisible     As Boolean
Private m_FirstDayOfWeek    As FirstDayOfWeekEnum

'[Default Property Values:]
Private Const m_def_Appearance = [Flat]
Private Const m_def_BorderStyle = [Fixel Single]
Private Const m_def_Enabled = True
Private Const m_def_Theme = Cal_Attraction
Private Const m_def_Animate = False
Private Const m_def_Mode = [List Mode]
Private Const m_def_CalendarHeight = 125
Private Const m_def_Curvature = 0
Private Const m_Def_CalendarGradient = True
Private Const m_def_CalendarBackCol = &HFFFFFF
Private Const m_def_SkipEnabled = False
Private Const m_def_Sensitive = True
Private Const m_def_HeaderHeight = 18

Private Const m_def_DateFormat = 0
Private Const m_def_WeekDayCol = &HFF9A35
Private Const m_def_DayCol = &HFDDBAC
Private Const m_def_DaySelCol = &HC4F9F9
Private Const m_def_WeekDaySelCol = &H59B4CA
Private Const m_def_DaySunCol = &HCAB7FD
Private Const m_def_WeekDaySunCol = &H8080FF
Private Const m_def_MonthGradient = True
Private Const m_def_HeaderGradient = True
Private Const m_def_MonthBackCol = &HFF7D7D
Private Const m_def_HeaderBackCol = &HFF7D7D
Private Const m_def_HeaderGradientCol = &HFFFFFF
Private Const m_def_MonthGradientCol = &HFFFFFF
Private Const m_def_CalendarGradientCol = &HFFFFFF
Private Const m_def_YearBackCol = &HFF7D7D
Private Const m_def_YearGradient = False
Private Const m_def_YearGradientCol = &HFFFFFF
Private Const m_def_SpecialDays = ""

Private Const m_Months = 12
Private Const m_HeaderDays = 7
Private Const m_RowDays = 5
Private Const m_def_ForeColor = 0
Private Const m_def_BorderColor = 0
Private Const m_def_HeaderVisible = True
Private Const m_def_FirstDayOfWeek = SunDay

' Constants for form animation
Private Const RGN_AND           As Long = 1
Private Const RGN_OR            As Long = 2
Private Const RGN_XOR           As Long = 3
Private Const RGN_COPY          As Long = 5
Private Const RGN_DIFF          As Long = 4

Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const DIB_RGB_ColS      As Long = 0
Private Const VER_PLATFORM_WIN32_NT  As Long = 2
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_TOOLWINDOW  As Long = &H80&

' for subclassing
Private Const WM_GETMINMAXINFO      As Long = &H24
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_WINDOWPOSCHANGING  As Long = &H46
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_SIZE               As Long = &H5
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_MOVE               As Long = &H3
Private Const WM_TIMER              As Long = &H113
Private Const WM_MOUSELEAVE         As Long = &H2A3

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private sc_aSubData()                As tSubData
Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean                                'Subclass data array

'[Event Declarations:]
Public Event DateChanged()


'[ Subclassed events receiver ]
'------------------------------
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

    ' Events for the calendar
    If lng_hWnd = m_PicCalendar.hWnd Then
    Select Case uMsg
        
        ' Mouse moving over the calendar
        Case WM_MOUSEMOVE
        
            ' Send event to picCalendar_MouseMove_SubClassed
            If m_TimerElsp = 10 Then m_TimerElsp = 0: Exit Sub
            picCalendar_MouseMove_SubClassed vbLeftButton, 999, WordLo(lParam), WordHi(lParam)
            SetTimer m_PicCalendar.hWnd, 1, 1, 0
            m_TimerElsp = 0
            m_PicToolTip.Visible = False
            
            ' Track mouse leave
            If Not bInCtrl Then
               bInCtrl = True
               Call TrackMouseLeave(lng_hWnd)
            End If
            
        ' Left button click on calendar
        Case WM_LBUTTONDOWN
            picCalendar_MouseDown_SubClassed vbLeftButton, 999, WordLo(lParam), WordHi(lParam)
        
        ' Right button click on calendar
        Case WM_RBUTTONDOWN
            picCalendar_MouseDown_SubClassed vbRightButton, 999, WordLo(lParam), WordHi(lParam)
        
        ' DbClick on calendar
        Case WM_LBUTTONDBLCLK
            picCalendar_DblClick_SubClassed
            
        ' The timer callback
        Case WM_TIMER
    
            m_TimerElsp = m_TimerElsp + 1
            If m_TimerElsp = 10 Then ' After 1 Sec
                KillTimer m_PicCalendar.hWnd, 1
                If bInCtrl Then ShowToolTip
            End If
            
        ' Mouse Leave
        Case WM_MOUSELEAVE
            bInCtrl = False
            
    End Select
    Exit Sub
    End If
    
    
    ' Events for the parent form
    If m_Mode = [List Mode] Or Not m_Poped Then Exit Sub
    Select Case uMsg
    
        ' Window resizing / clicking
        Case WM_SIZE, WM_LBUTTONDOWN, WM_RBUTTONDOWN
            m_PicCalendar.Visible = False
            m_Poped = False
        
        ' Change in position
        Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED
            If m_Mode = [PopDown Mode] Then ExportCalendar True, True
            If m_Mode = [PopUp Mode] And m_Poped Then ExportCalendar False, True

        ' Trying to minimize or maximize
        Case WM_GETMINMAXINFO
            If UserControl.Parent.WindowState = 1 Then
                m_PicCalendar.Visible = False
            Else
                m_PicCalendar.Visible = True
            End If
            
    End Select
    
End Sub


'[ Apply Calendar themes ]
'-------------------------
Private Sub ApplyTheme(ByVal ThemeIndex As CalendarTheme)

'Debug.Print "Applying new theme "
Select Case ThemeIndex
    Case [Cal_Attraction]
        m_HeaderBackCol = &HFF7D7D
        m_ArrowCol = &H257A4B
        m_MonthBackCol = &HFF7D7D
        m_DayCol = &HFDDBAC
        m_DaySunCol = &HCAB7FD
        m_WeekDayCol = &HFF9A35
        m_WeekDaySunCol = &H8080FF
        m_WeekDaySelCol = &H59B4CA
        m_DaySelCol = &HC4F9F9
        m_BorderColor = 0
        
    Case [Cal Blue]
        m_HeaderBackCol = &HDABAA8
        m_ArrowCol = &HDCC1AD
        m_MonthBackCol = &HEDC5A7
        m_DayCol = &HFCF4EF
        m_DaySunCol = &HAEA6F8
        m_WeekDayCol = &HFAE8DC
        m_WeekDaySunCol = &H8080FF
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HD8E5C8
        m_BorderColor = 0 ' &H864E02
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Green]
        m_HeaderBackCol = &HB1CB92
        m_ArrowCol = &H213B00
        m_MonthBackCol = &HB1CB92
        m_DayCol = &HE1EBD5
        m_DaySunCol = &HAEA6F8
        m_WeekDaySunCol = &H8080FF
        m_WeekDayCol = &HFAE8DC
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HDABAA8
        m_BorderColor = 0 ' &H185232
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Orange]
        m_HeaderBackCol = &HD2E2FD
        m_ArrowCol = &H16366D
        m_MonthBackCol = &HD2E2FD
        m_DayCol = &HEFF5FE
        m_DaySunCol = &HE3E3D6
        m_WeekDaySunCol = &HD5B0BF
        m_WeekDayCol = &HE3E3D6
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HB1CB92
        m_BorderColor = 0 '&H80FF&
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Purple]
        m_HeaderBackCol = &HD5B0BF
        m_ArrowCol = &H46202F
        m_MonthBackCol = &HD5B0BF
        m_DayCol = &HF7F1F3
        m_DaySunCol = &HB1CB92
        m_WeekDaySunCol = &HD5B0BF
        m_WeekDayCol = &HD1A9B9
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        m_BorderColor = 0 '&HE4616D
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Red]
        m_HeaderBackCol = &HAEA6F8
        m_ArrowCol = &H1D156A
        m_MonthBackCol = &HA79EF7
        m_DayCol = &HFFFFFF
        m_DaySunCol = &HD6D2FB
        m_WeekDaySunCol = &HD6D2FB
        m_WeekDayCol = &HAEA6F8
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        m_BorderColor = 0 '&HC0&
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Silver]
        m_HeaderBackCol = &HD9D6D3
        m_ArrowCol = &H4A4744
        m_MonthBackCol = &HD9D6D3
        m_DayCol = &HFFFFFF
        m_DaySunCol = &HD6D2FB
        m_WeekDaySunCol = &HD6D2FB
        m_WeekDayCol = &HD9D6D3
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        m_BorderColor = 0 '&H808080
        m_HeaderBackCol = m_MonthBackCol
        
    Case [Cal Yellow]
        m_HeaderBackCol = &HB9EEF4
        m_ArrowCol = &H66D5E1
        m_MonthBackCol = &HB9EEF4
        m_DayCol = &HFFFFFF
        m_DaySunCol = &HD6D2FB
        m_WeekDaySunCol = &HD6D2FB
        m_WeekDayCol = &HB9EEF4
        m_WeekDaySelCol = &HF1F5EB
        m_DaySelCol = &HE3E3D6
        m_BorderColor = 0 ' &H57C9E
        m_HeaderBackCol = m_MonthBackCol
        
End Select

m_MonthGradientCol = m_HeaderGradientCol
m_CalendarGradientCol = m_HeaderGradientCol
m_CalendarBackCol = m_MonthBackCol
m_YearBackCol = m_MonthBackCol
m_CalendarGradientCol = m_MonthGradientCol
m_ForeColor = 0

End Sub


'[ All the properties this usercontrol got ]
'-------------------------------------------
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    m_Appearance = New_Appearance
    UserControl.Appearance = New_Appearance
    PropertyChanged "Appearance"
    RedrawControl
End Property


Public Property Get BorderStyle() As BorderStyleEnum
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
    m_BorderStyle = New_BorderStyle
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    RedrawControl
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = New_Enabled
End Property


Public Property Get CalendarGradient() As Boolean
    CalendarGradient = m_CalendarGradient
End Property

Public Property Let CalendarGradient(ByVal vNewValue As Boolean)
    m_CalendarGradient = vNewValue
    PropertyChanged "CalendarGradient"
    RedrawControl
End Property


Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
    Set m_Font = New_Font
    Set UserControl.Font = New_Font
    Set m_PicCalendar.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property


Public Property Get SpecialDays() As String
    SpecialDays = m_SpecialDays
End Property

Public Property Let SpecialDays(ByVal New_SpecialDays As String)
    
    ' This property is read only at runtime
    If Ambient.UserMode Then Err.Raise 382
    m_SpecialDays = New_SpecialDays
    PropertyChanged "SpecialDays"
    If Not m_SpecialDays = vbNullString Then m_SpecialDayStack = Split(m_SpecialDays, ",")
    RedrawControl

End Property

Public Property Get Sensitive() As Boolean
    Sensitive = m_Sensitive
End Property

Public Property Let Sensitive(ByVal New_Sensitive As Boolean)
    m_Sensitive = New_Sensitive
    PropertyChanged "Sensitive"
End Property


Public Property Get SkipEnabled() As Boolean
    SkipEnabled = m_SkipEnabled
End Property

Public Property Let SkipEnabled(ByVal New_SkipEnabled As Boolean)
    m_SkipEnabled = New_SkipEnabled
    PropertyChanged "SkipEnabled"
End Property


Public Property Get Theme() As CalendarTheme
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As CalendarTheme)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    ApplyTheme New_Theme
    RedrawControl
End Property


Public Property Get Animate() As Boolean
    Animate = m_Animate
End Property

Public Property Let Animate(ByVal New_Animate As Boolean)
    m_Animate = New_Animate
    PropertyChanged "Animate"
End Property


Public Property Get CalendarHeight() As Long
    CalendarHeight = m_CalendarHeight
End Property

Public Property Let CalendarHeight(ByVal vNewValue As Long)
    m_CalendarHeight = vNewValue
    If m_Mode = [List Mode] Then Height = (m_CalendarHeight + m_HeaderHeight) * Screen.TwipsPerPixelY
    PropertyChanged "CalendarHeight"
    RedrawControl
End Property


Public Property Get Caption(ByVal LongDate As Boolean) As String
    'Format with User settings in Regional Configurations
    If LongDate Then
        Caption = FormatDateTime$(Value, vbLongDate)
    Else
        Caption = FormatDateTime$(Value, vbShortDate)
    End If
End Property


Public Property Get Curvature() As Long
    Curvature = m_Curvature
End Property

Public Property Let Curvature(ByVal vNewValue As Long)
    m_Curvature = vNewValue
    PropertyChanged "Curvature"
    RedrawControl
End Property


Public Property Get Value() As Date
   Value = DateSerial(m_SelYear, m_SelMonth, m_SelDay)
End Property

Public Property Let Value(ByVal vNewValue As Date)
   m_SelYear = Year(vNewValue)
   m_SelMonth = Month(vNewValue)
   m_SelDay = Day(vNewValue)
   RedrawControl
End Property


Public Property Get DateX() As String
Attribute DateX.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute DateX.VB_UserMemId = 0
Attribute DateX.VB_MemberFlags = "200"
    DateX = Format(DateSerial(m_SelYear, m_SelMonth, m_SelDay), GetFormat)
End Property

Public Property Let YearX(ByVal vNewValue As Long)
    m_SelYear = vNewValue
    PropertyChanged "YearX"
    LoadDay (999)
    RedrawControl
End Property


Public Property Get YearX() As Long
    YearX = m_SelYear
End Property

Public Property Get MonthX() As Long
    MonthX = m_SelMonth
End Property

Public Property Let MonthX(ByVal vNewValue As Long)
    m_SelMonth = vNewValue
    PropertyChanged "MonthX"
    LoadDay (999)
    RedrawControl
End Property


Public Property Get DayX() As Long
    DayX = m_SelDay
End Property

Public Property Let DayX(ByVal vNewValue As Long)
    m_SelDay = vNewValue
    If m_SelDay > m_MonthDays Then m_SelDay = m_MonthDays
    If m_SelDay < 1 Then m_SelDay = 1
    LoadDay vNewValue
    PropertyChanged "DayX"
    RedrawControl
End Property


Public Property Get Mode() As CalendarMode
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal vNewValue As CalendarMode)
    m_Mode = vNewValue
    PropertyChanged "Mode"
    If Not m_Mode = [List Mode] Then
        Height = m_HeaderHeight * Screen.TwipsPerPixelX
        m_PicCalendar.Visible = False
    Else
        Height = (m_HeaderHeight + m_CalendarHeight) * Screen.TwipsPerPixelX
        ImportCalendar
    End If
    UserControl_Resize
End Property


Public Property Get CalendarBackCol() As OLE_COLOR
    CalendarBackCol = m_CalendarBackCol
End Property

Public Property Let CalendarBackCol(ByVal vNewValue As OLE_COLOR)
    m_CalendarBackCol = vNewValue
    PropertyChanged "CalendarBackCol"
    RedrawControl
End Property


Public Property Get MonthGradient() As Boolean
    MonthGradient = m_MonthGradient
End Property

Public Property Let MonthGradient(ByVal New_MonthGradient As Boolean)
    m_MonthGradient = New_MonthGradient
    PropertyChanged "MonthGradient"
    RedrawControl
End Property


Public Property Get HeaderGradient() As Boolean
    HeaderGradient = m_HeaderGradient
End Property

Public Property Let HeaderGradient(ByVal New_HeaderGradient As Boolean)
    m_HeaderGradient = New_HeaderGradient
    PropertyChanged "HeaderGradient"
    RedrawControl
End Property


Public Property Get MonthBackCol() As OLE_COLOR
    MonthBackCol = m_MonthBackCol
End Property

Public Property Let MonthBackCol(ByVal New_MonthBackCol As OLE_COLOR)
    m_MonthBackCol = New_MonthBackCol
    PropertyChanged "MonthBackCol"
    RedrawControl
End Property


Public Property Get HeaderBackCol() As OLE_COLOR
    HeaderBackCol = m_HeaderBackCol
End Property

Public Property Let HeaderBackCol(ByVal New_HeaderBackCol As OLE_COLOR)
    m_HeaderBackCol = New_HeaderBackCol
    PropertyChanged "HeaderBackCol"
    RedrawControl
End Property


Public Property Get WeekDayCol() As OLE_COLOR
    WeekDayCol = m_WeekDayCol
End Property

Public Property Let WeekDayCol(ByVal New_WeekDayCol As OLE_COLOR)
    m_WeekDayCol = New_WeekDayCol
    PropertyChanged "WeekDayCol"
    RedrawControl
End Property


Public Property Get DayCol() As OLE_COLOR
    DayCol = m_DayCol
End Property

Public Property Let DayCol(ByVal New_DayCol As OLE_COLOR)
    m_DayCol = New_DayCol
    PropertyChanged "DayCol"
    RedrawControl
End Property


Public Property Get DaySelCol() As OLE_COLOR
    DaySelCol = m_DaySelCol
End Property

Public Property Let DaySelCol(ByVal New_DaySelCol As OLE_COLOR)
    m_DaySelCol = New_DaySelCol
    PropertyChanged "DaySelCol"
    RedrawControl
End Property


Public Property Get WeekDaySelCol() As OLE_COLOR
    WeekDaySelCol = m_WeekDaySelCol
End Property

Public Property Let WeekDaySelCol(ByVal New_WeekDaySelCol As OLE_COLOR)
    m_WeekDaySelCol = New_WeekDaySelCol
    PropertyChanged "WeekDaySelCol"
    RedrawControl
End Property


Public Property Get HeaderHeight() As Long
    HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeaderHeight As Long)
    m_HeaderHeight = New_HeaderHeight
    If m_HeaderHeight < m_PicCalendar.TextHeight("A") Then m_HeaderHeight = m_PicCalendar.TextHeight("A")
    PropertyChanged "HeaderHeight"
    UserControl_Resize
End Property


Public Property Get DaySunCol() As OLE_COLOR
    DaySunCol = m_DaySunCol
End Property

Public Property Let DaySunCol(ByVal New_DaySunCol As OLE_COLOR)
    m_DaySunCol = New_DaySunCol
    PropertyChanged "DaySunCol"
    RedrawControl
End Property


Public Property Get WeekDaySunCol() As OLE_COLOR
    WeekDaySunCol = m_WeekDaySunCol
End Property

Public Property Let WeekDaySunCol(ByVal New_WeekDaySunCol As OLE_COLOR)
    m_WeekDaySunCol = New_WeekDaySunCol
    PropertyChanged "WeekDaySunCol"
    RedrawControl
End Property


Public Property Get HeaderGradientCol() As OLE_COLOR
    HeaderGradientCol = m_HeaderGradientCol
End Property

Public Property Let HeaderGradientCol(ByVal New_HeaderGradientCol As OLE_COLOR)
    m_HeaderGradientCol = New_HeaderGradientCol
    PropertyChanged "HeaderGradientCol"
    RedrawControl
End Property


Public Property Get MonthGradientCol() As OLE_COLOR
    MonthGradientCol = m_MonthGradientCol
End Property

Public Property Let MonthGradientCol(ByVal New_MonthGradientCol As OLE_COLOR)
    m_MonthGradientCol = New_MonthGradientCol
    PropertyChanged "MonthGradientCol"
    RedrawControl
End Property


Public Property Get CalendarGradientCol() As OLE_COLOR
    CalendarGradientCol = m_CalendarGradientCol
End Property

Public Property Let CalendarGradientCol(ByVal New_CalendarGradientCol As OLE_COLOR)
    m_CalendarGradientCol = New_CalendarGradientCol
    PropertyChanged "CalendarGradientCol"
    RedrawControl
End Property


Public Property Get YearBackCol() As OLE_COLOR
    YearBackCol = m_YearBackCol
End Property

Public Property Let YearBackCol(ByVal New_YearBackCol As OLE_COLOR)
    m_YearBackCol = New_YearBackCol
    PropertyChanged "YearBackCol"
    RedrawControl
End Property


Public Property Get YearGradient() As Boolean
    YearGradient = m_YearGradient
End Property

Public Property Let YearGradient(ByVal New_YearGradient As Boolean)
    m_YearGradient = New_YearGradient
    PropertyChanged "YearGradient"
    RedrawControl
End Property


Public Property Get YearGradientCol() As OLE_COLOR
    YearGradientCol = m_YearGradientCol
End Property

Public Property Let YearGradientCol(ByVal New_YearGradientCol As OLE_COLOR)
    m_YearGradientCol = New_YearGradientCol
    PropertyChanged "YearGradientCol"
    RedrawControl
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    RedrawControl
End Property


Public Property Get DateFormat() As DateFormatEnum
    DateFormat = m_DateFormat
End Property

Public Property Let DateFormat(ByVal New_DateFormat As DateFormatEnum)
    m_DateFormat = New_DateFormat
    PropertyChanged "DateFormat"
    RedrawControl
End Property


Public Property Get FirstDayOfWeek() As FirstDayOfWeekEnum
    FirstDayOfWeek = m_FirstDayOfWeek
End Property

Public Property Let FirstDayOfWeek(ByVal New_FirstDayOfWeek As FirstDayOfWeekEnum)
    m_FirstDayOfWeek = New_FirstDayOfWeek
    PropertyChanged "FirstDayOfWeek"
    RedrawControl
End Property


Public Property Get HeaderVisible() As Boolean
    HeaderVisible = m_HeaderVisible
End Property

Public Property Let HeaderVisible(ByVal New_HeaderVisible As Boolean)
    m_HeaderVisible = New_HeaderVisible
    PropertyChanged "HeaderVisible"
    DoEvents
    UserControl_Resize
End Property

Public Property Get hWnd() As Long

    '* English: Returns a handle to a form or control.
    hWnd = UserControl.hWnd

End Property

'-----------------------------------------------------------------------------------------------------------
' Events on m_picCalendar are handled here :
' Note that these are orginally created without subclassing
' Now this codes handles EVENTs send from the subclass event receiver
'-----------------------------------------------------------------------------------------------------------

Private Sub picCalendar_DblClick_SubClassed()
    picCalendar_MouseDown_SubClassed vbLeftButton, 111, m_MouseX, m_MouseY
End Sub


Private Sub picCalendar_MouseDown_SubClassed(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nDay As Long

' This is the case when the user try to select a month from the
' popuped month list.
If m_MonthPopupMode Then

    If X < m_TrackWidth Then
        ' Not moving over the list. Nothing is to do here
    Else
        ' Calculate the Month that the User selected
        m_SelMonth = Int((X - m_TrackWidth) / m_MonthPopWidth + 1)
    End If
    
    ' Disable popup mode and draw new month
    m_MonthPopupMode = False
    LoadDay (999)
    RedrawControl
    Exit Sub
    
End If

' Moving over the Calendarn whaever may happen

' Moving through the left Month Display region
If X < m_TrackWidth Then
    
    ' Clecking over the Month Up button
    If Y < m_TrackWidth Then
        
        ' Select the last Month
        If m_SelMonth = 1 Then      'Back of year
            m_SelMonth = 12
            m_SelYear = m_SelYear - 1
        Else                         'No problem preceed
            m_SelMonth = m_SelMonth - 1
        End If
        
        ' Load the day, 999 checks the new Month posses Days more
        ' than selected date
        LoadDay (999)
        
    ' Clicking over Month Down Button
    ElseIf Y > m_PicCalendar.ScaleHeight - m_TrackWidth Then
    
        ' Select the last Month
        If m_SelMonth = 12 Then     ' End of year
            m_SelMonth = 1
            m_SelYear = m_SelYear + 1
        Else                        ' No problem proceed
            m_SelMonth = m_SelMonth + 1
        End If
        
        ' Load the day, 999 checks the new Month posses Days more
        ' than selected date
        LoadDay (999)
        
    'Moving throgh left TrackRegion( Month Show).
    Else
    
        ' Clicking for popuping Month list
        If X < m_TrackWidth And X > 0.75 * m_TrackWidth Then
            PopupMonthList
            DrawBody
            Exit Sub
        End If
        
    End If

' Not throgh month display region.
' To the Right of that
Else
    
    ' Moving through Header ( Week days display )
    If Y < m_WeekDaysHeight Then
        ' No events added till this version
    
    ' Moving through "DAYS'
    Else
        
        ' Trying to select a Day
        If Y < m_PicCalendar.ScaleHeight - m_iHeight Then
        
            ' Calculate Day\Load it
            If Shift = 111 Then ' Event was send from DbClick
                If Not m_Mode = [List Mode] Then CollapseCalendar
            Else
            
            ' Calculate the day
            nDay = Int((Y - m_WeekDaysHeight) / m_iHeight) * m_HeaderDays + (Int((X - m_TrackWidth) / m_iWidth)) + 2 - m_FirstDay + (m_FirstDayOfWeek - 1)
            LoadDay nDay
            End If
            
        'Year selection region
        Else
            
            'Next Year Selecting Button
            If X > m_PicCalendar.ScaleWidth - m_iWidth + 10 Then
                
                ' Load Next year . RightClick will jump FIVE
                If Button = vbLeftButton Then m_SelYear = m_SelYear + 1 Else m_SelYear = m_SelYear + 5
                
                ' Load the day, 999 checks the new Month in new year posses Days more
                ' than selected date
                LoadDay (999)
                
            'Last Year Selecting Button
            ElseIf X > m_iWidth * 4 + m_TrackWidth And X < m_iWidth * 4 + m_TrackWidth + m_iWidth - 10 Then
                
                ' Load Lastyear
                If Button = vbLeftButton Then m_SelYear = m_SelYear - 1 Else m_SelYear = m_SelYear - 5
                
                ' Load the day, 999 checks the new Month in new year posses Days more
                ' than selected date
                LoadDay (999)
                
            Else
                ' Over the Today Region
                m_SelDay = Day(Date)
                m_SelMonth = Month(Date)
                m_SelYear = Year(Date)
                LoadDay m_SelDay
            End If
            
        End If
    End If
End If

RedrawControl

End Sub

Private Sub picCalendar_MouseMove_SubClassed(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nDay As String

' Store the pos fot tooltip
m_MouseX = X
m_MouseY = Y
    
' Moving through popuped month list
If m_MonthPopupMode Then

    If X < m_TrackWidth Then
        m_ToolTipText = MonthName(m_SelMonth)
        m_PicCalendar.MousePointer = vbNormal
    Else
        m_ToolTipText = MonthName(Int((X - m_TrackWidth) / m_MonthPopWidth) + 1)
        m_PicCalendar.MousePointer = vbCustom
    End If
    
Exit Sub
End If


If X <= m_TrackWidth Then

    ' Month Selection Up
    If Y < m_TrackWidth Then
    
        m_ToolTipText = "Last Month"
        m_PicCalendar.MousePointer = vbCustom

    ' Month Selection Down
    ElseIf Y > m_PicCalendar.ScaleHeight - m_TrackWidth Then
    
        m_ToolTipText = "Next Month"
        m_PicCalendar.MousePointer = vbCustom
        
    ' b/w both
    Else
    
        ' Popup button
        If X < m_TrackWidth And X > 0.75 * m_TrackWidth Then
        
            m_PicCalendar.MousePointer = vbCustom
            m_ToolTipText = "Popup Month List >"
            
        Else
        
            m_PicCalendar.MousePointer = vbNormal
            m_ToolTipText = MonthName(m_SelMonth)
            
        End If
        
    End If
    
Else

    ' week days
    If Y < m_WeekDaysHeight Then
    
        nDay = Int((X - 1 - m_TrackWidth) / m_iWidth) + m_FirstDayOfWeek
        If nDay > 7 Then nDay = nDay - 7
        m_ToolTipText = WeekdayName(nDay)
        m_PicCalendar.MousePointer = vbNormal
    
    ' Bellow
    Else
    
        'Through days
        If Y < m_PicCalendar.ScaleHeight - m_iHeight Then
            
            ' Calculate the day
            nDay = Int((Y - m_WeekDaysHeight) / m_iHeight) * m_HeaderDays + (Int((X - m_TrackWidth) / m_iWidth)) + 2 - m_FirstDay + (m_FirstDayOfWeek - 1)
            If m_FirstDay < m_FirstDayOfWeek Then nDay = nDay - 7
            
            If nDay < 0 Then
                nDay = (35 - m_FirstDay) + (m_FirstDay + nDay)
            End If

            ' Clicking on Diagonal croseed days will unload Calendar if Sensitive=True
            If nDay > m_MonthDays Then
                 If m_Sensitive And Not m_Mode = [List Mode] Then m_ToolTipText = "Close": m_PicCalendar.MousePointer = vbCustom: Exit Sub
                 If m_SkipEnabled Then m_ToolTipText = "Skip Next Month" Else m_ToolTipText = ""
                 
            ElseIf nDay <= 0 Then
                 If m_Sensitive And Not m_Mode = [List Mode] Then m_ToolTipText = "Close": m_PicCalendar.MousePointer = vbCustom: Exit Sub
                 If m_SkipEnabled Then m_ToolTipText = "Skip Last Month" Else m_ToolTipText = ""
                 
            Else
                If IsSpecial(nDay, m_SelMonth) Then
                    m_ToolTipText = " " & MonthName(m_SelMonth) & " " & nDay & ", " & m_SpecialDayString
                Else
                    m_ToolTipText = " " & MonthName(m_SelMonth) & " " & nDay & " "
                End If
                
            End If
            m_PicCalendar.MousePointer = vbCustom
            
        Else    ' footer
        
            ' Year selecting region
            If X > m_iWidth * 4 + m_TrackWidth Then
            
                'Last Year Button
                If X < m_iWidth * 4 + m_TrackWidth + m_iWidth - 10 Then
                
                    m_ToolTipText = "Last Year"
                    m_PicCalendar.MousePointer = vbCustom
                    
                ' Next Year Button
                ElseIf X > m_PicCalendar.ScaleWidth - m_iWidth + 10 Then
                
                    m_ToolTipText = "Next Year"
                    m_PicCalendar.MousePointer = vbCustom
                    
                Else 'Middle
                
                    m_PicCalendar.MousePointer = vbNormal
                    m_ToolTipText = "Year " & m_SelYear
                    
                End If
                
            Else 'Through Date Display
            
                m_PicCalendar.MousePointer = vbCustom
                m_ToolTipText = "Today " & Format$(Date$, GetFormat)
            
            End If
        End If
    End If

End If

End Sub


'-----------------------------------------------------------------------------------------------------------
' Events on the Usercontrol: (Not subclassed)
'-----------------------------------------------------------------------------------------------------------

Private Sub UserControl_DblClick()
    m_Right2Left = False: DrawBody
End Sub


Private Sub UserControl_Initialize()

    'Debug.Print vbCrLf & "--------------------------------------" & vbCrLf & "New Compile" & vbCrLf & "--------------------------------------"
    ' Used to prevent crashes on XP
    m_hMode = LoadLibrary("shell32.dll")
    CreateControls
    
    m_SelDay = Day(Date)
    m_SelMonth = Month(Date)
    m_SelYear = Year(Date)
    m_Right2Left = True
    LoadDay m_SelDay
    
End Sub

Private Sub UserControl_InitProperties()

    Me.Appearance = m_def_Appearance
    Me.BorderStyle = m_def_BorderStyle
    Me.Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_Theme = m_def_Theme
    m_Animate = m_def_Animate
    m_CalendarHeight = m_def_CalendarHeight
    m_Curvature = m_def_Curvature
    m_Mode = m_def_Mode
    m_SelDay = Day(Date)
    m_SelMonth = Month(Date)
    m_SelYear = Year(Date)
    m_CalendarBackCol = m_def_CalendarBackCol
    m_MonthGradient = m_def_MonthGradient
    m_MonthBackCol = m_def_MonthBackCol
    m_HeaderBackCol = m_def_HeaderBackCol
    m_WeekDayCol = m_def_WeekDayCol
    m_DayCol = m_def_DayCol
    m_DaySelCol = m_def_DaySelCol
    m_WeekDaySelCol = m_def_WeekDaySelCol
    m_DaySunCol = m_def_DaySunCol
    m_WeekDaySunCol = m_def_WeekDaySunCol
    m_HeaderGradientCol = m_def_HeaderGradientCol
    m_CalendarGradientCol = m_def_CalendarGradientCol
    m_YearBackCol = m_def_YearBackCol
    m_YearGradient = m_def_YearGradient
    m_YearGradientCol = m_def_YearGradientCol
    
    ApplyTheme Cal_Attraction
    m_MonthGradientCol = m_def_MonthGradientCol
    m_HeaderGradient = m_def_HeaderGradient
    m_CalendarGradient = m_Def_CalendarGradient
    m_DateFormat = m_def_DateFormat
    m_Sensitive = m_def_Sensitive
    m_SkipEnabled = m_def_SkipEnabled
    m_HeaderHeight = m_def_HeaderHeight
    m_SpecialDays = m_def_SpecialDays
    m_FirstDayOfWeek = m_def_FirstDayOfWeek

    m_HeaderVisible = m_def_HeaderVisible
    m_ForeColor = m_def_ForeColor
    m_BorderColor = m_def_BorderColor
    
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Clicking on PopDown button
    If X > ScaleWidth - (m_HeaderHeight + 10) And m_Mode = [PopDown Mode] Then

        If Not m_Poped Then
            m_Right2Left = False
            DrawBody
            
            ' Export Calendar up/down
            ExportCalendar True
            DrawCalendar
        Else
            CollapseCalendar
            DrawBody
        End If
        
    ' Clicking on PopUp button
    ElseIf X < (m_HeaderHeight + 10) And m_Mode = [PopUp Mode] Then
        
        If Not m_Poped Then
            m_Right2Left = False
            DrawBody
            
            ' Export Calendar
            ExportCalendar False
            DrawCalendar
        Else
            CollapseCalendar
            DrawBody
        End If
        
    End If

Exit Sub
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Over the button
    If X > ScaleWidth - (m_HeaderHeight + 10) And m_Mode = [PopDown Mode] Then
        UserControl.MousePointer = vbCustom
    ElseIf X < (m_HeaderHeight + 10) And (m_Mode = [PopUp Mode]) Then
        UserControl.MousePointer = vbCustom
    Else
        UserControl.MousePointer = vbNormal
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If X > ScaleWidth - (m_HeaderHeight + 10) Or X < (m_HeaderHeight + 10) Then
        
        m_Right2Left = True
        DrawBody
        
    End If
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'Debug.Print "Reading Properties "
    
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    m_Animate = PropBag.ReadProperty("Animate", m_def_Animate)
    m_CalendarHeight = PropBag.ReadProperty("CalendarHeight", m_def_CalendarHeight)
    m_Curvature = PropBag.ReadProperty("Curvature", m_def_Curvature)
    m_Mode = PropBag.ReadProperty("Mode", m_def_Mode)
    
    m_CalendarGradient = PropBag.ReadProperty("CalendarGradient", m_Def_CalendarGradient)
    m_CalendarBackCol = PropBag.ReadProperty("CalendarBackCol", m_def_CalendarBackCol)
    m_MonthGradient = PropBag.ReadProperty("MonthGradient", m_def_MonthGradient)
    m_HeaderGradient = PropBag.ReadProperty("HeaderGradient", m_def_HeaderGradient)
    m_MonthBackCol = PropBag.ReadProperty("MonthBackCol", m_def_MonthBackCol)
    m_HeaderBackCol = PropBag.ReadProperty("HeaderBackCol", m_def_HeaderBackCol)
    m_WeekDayCol = PropBag.ReadProperty("WeekDayCol", m_def_WeekDayCol)
    m_DayCol = PropBag.ReadProperty("DayCol", m_def_DayCol)
    m_DaySelCol = PropBag.ReadProperty("DaySelCol", m_def_DaySelCol)
    m_WeekDaySelCol = PropBag.ReadProperty("WeekDaySelCol", m_def_WeekDaySelCol)
    m_DaySunCol = PropBag.ReadProperty("DaySunCol", m_def_DaySunCol)
    m_WeekDaySunCol = PropBag.ReadProperty("WeekDaySunCol", m_def_WeekDaySunCol)
    m_HeaderGradientCol = PropBag.ReadProperty("HeaderGradientCol", m_def_HeaderGradientCol)
    m_MonthGradientCol = PropBag.ReadProperty("MonthGradientCol", m_def_MonthGradientCol)
    m_CalendarGradientCol = PropBag.ReadProperty("CalendarGradientCol", m_def_CalendarGradientCol)
    m_YearBackCol = PropBag.ReadProperty("YearBackCol", m_def_YearBackCol)
    m_YearGradient = PropBag.ReadProperty("YearGradient", m_def_YearGradient)
    m_YearGradientCol = PropBag.ReadProperty("YearGradientCol", m_def_YearGradientCol)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)

    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    m_Sensitive = PropBag.ReadProperty("Sensitive", m_def_Sensitive)
    m_SkipEnabled = PropBag.ReadProperty("SkipEnabled", m_def_SkipEnabled)
    m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", m_def_HeaderHeight)
    m_SpecialDays = PropBag.ReadProperty("SpecialDays", m_def_SpecialDays)
    m_FirstDayOfWeek = PropBag.ReadProperty("FirstDayOfWeek", m_def_FirstDayOfWeek)
    m_HeaderVisible = PropBag.ReadProperty("HeaderVisible", m_def_HeaderVisible)

    UserControl.Appearance = m_Appearance
    UserControl.BorderStyle = m_BorderStyle
    UserControl.Enabled = m_Enabled
    Set UserControl.Font = m_Font
    Set m_PicCalendar.Font = m_Font
    If Not m_SpecialDays = vbNullString Then m_SpecialDayStack = Split(m_SpecialDays, ",")
    ImportCalendar
    UserControl_Resize
    
    ' We have to subclass the parent in runtime
    If Ambient.UserMode Then
    
    bTrack = True
    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  
    If Not bTrackUser32 Then
      If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
        bTrack = False
      End If
    End If
    
    If Not bTrack Then Exit Sub
    
    With UserControl.Parent
        
        ' Start subclassing the parent form
        Call Subclass_Start(.hWnd)
        
        ' Adding the messages we need to track
        Call Subclass_AddMsg(.hWnd, WM_WINDOWPOSCHANGING, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_WINDOWPOSCHANGED, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_GETMINMAXINFO, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_SIZE, MSG_AFTER)
        
    End With
    
        With m_PicCalendar
        
        ' Start subclassing our calendar
        Call Subclass_Start(.hWnd)
        
        ' Adding the messages we need to track
        Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_LBUTTONDBLCLK, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_RBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_TIMER, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)

    End With
    
    End If
    
End Sub

Private Sub UserControl_Resize()
On Error GoTo Handle
    
    ' set the height
    If Not m_Mode = [List Mode] Then
        Height = m_HeaderHeight * Screen.TwipsPerPixelX
    Else
        m_CalendarHeight = Height / Screen.TwipsPerPixelY - m_HeaderHeight
    End If
    m_PicCalendar.Width = Width / Screen.TwipsPerPixelX
    m_PicCalendar.Height = m_CalendarHeight
    
    RedrawControl
    
Handle:
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
    'Stop all subclassing
    Call Subclass_Stop(UserControl.Parent.hWnd)
    Call Subclass_Stop(m_PicCalendar.hWnd)
    Call Subclass_StopAll
    FreeLibrary m_hMode
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Debug.Print "Writing Properties "
    
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
    Call PropBag.WriteProperty("Animate", m_Animate, m_def_Animate)
    Call PropBag.WriteProperty("CalendarHeight", m_CalendarHeight, m_def_CalendarHeight)
    Call PropBag.WriteProperty("Curvature", m_Curvature, m_def_Curvature)
    Call PropBag.WriteProperty("Mode", m_Mode, m_def_Mode)
    Call PropBag.WriteProperty("CalendarGradient", m_CalendarGradient, m_Def_CalendarGradient)
    Call PropBag.WriteProperty("CalendarBackCol", m_CalendarBackCol, m_def_CalendarBackCol)
    
    Call PropBag.WriteProperty("MonthGradient", m_MonthGradient, m_def_MonthGradient)
    Call PropBag.WriteProperty("HeaderGradient", m_HeaderGradient, m_def_HeaderGradient)
    Call PropBag.WriteProperty("MonthBackCol", m_MonthBackCol, m_def_MonthBackCol)
    Call PropBag.WriteProperty("HeaderBackCol", m_HeaderBackCol, m_def_HeaderBackCol)
    Call PropBag.WriteProperty("WeekDayCol", m_WeekDayCol, m_def_WeekDayCol)
    Call PropBag.WriteProperty("DayCol", m_DayCol, m_def_DayCol)
    Call PropBag.WriteProperty("DaySelCol", m_DaySelCol, m_def_DaySelCol)
    Call PropBag.WriteProperty("WeekDaySelCol", m_WeekDaySelCol, m_def_WeekDaySelCol)
    Call PropBag.WriteProperty("DaySunCol", m_DaySunCol, m_def_DaySunCol)
    Call PropBag.WriteProperty("WeekDaySunCol", m_WeekDaySunCol, m_def_WeekDaySunCol)
    Call PropBag.WriteProperty("MonthGradientCol", m_MonthGradientCol, m_def_MonthGradientCol)
    Call PropBag.WriteProperty("CalendarGradientCol", m_CalendarGradientCol, m_def_CalendarGradientCol)

    Call PropBag.WriteProperty("YearBackCol", m_YearBackCol, m_def_YearBackCol)
    Call PropBag.WriteProperty("YearGradient", m_YearGradient, m_def_YearGradient)
    Call PropBag.WriteProperty("YearGradientCol", m_YearGradientCol, m_def_YearGradientCol)
    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    Call PropBag.WriteProperty("Sensitive", m_Sensitive, m_def_Sensitive)
    Call PropBag.WriteProperty("SkipEnabled", m_SkipEnabled, m_def_SkipEnabled)
    Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, m_def_HeaderHeight)
    Call PropBag.WriteProperty("SpecialDays", m_SpecialDays, m_def_SpecialDays)
    Call PropBag.WriteProperty("FirstDayOfWeek", m_FirstDayOfWeek, m_def_FirstDayOfWeek)

    Call PropBag.WriteProperty("HeaderVisible", m_HeaderVisible, m_def_HeaderVisible)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
End Sub


'-----------------------------------------------------------------------------------------------------------
' Our private procedures, which are the CORE
'-----------------------------------------------------------------------------------------------------------

'[ Redraw the contol ]
'---------------------
Private Sub RedrawControl()
Dim hrgn As Long

    ' DatePicker does not need the Header
    If m_HeaderVisible Then
        hrgn = CreateRectRgn(0, 0, ScaleWidth + 2, ScaleHeight + 2)
        SetWindowRgn UserControl.hWnd, hrgn, True
    Else 'Header not visible
        hrgn = CreateRectRgn(0, m_HeaderHeight, ScaleWidth + 2, ScaleHeight + 2)
        SetWindowRgn UserControl.hWnd, hrgn, True
    End If
    
    If m_Mode = [List Mode] Then ImportCalendar
    DrawCalendar
    DrawBody

End Sub


'[ Draw the header on Usercontrol ]
'----------------------------------
Private Sub DrawBody()
Dim Rct As RECT
Dim vStrDate As String
    
    'Debug.Print "Drawing the body "

    ' Define the RECT
    Rct.Left = 0
    Rct.Right = ScaleWidth
    Rct.Top = (m_HeaderHeight - TextHeight("A")) / 2
    Rct.Bottom = m_HeaderHeight
    
    ' Get the Selectecd Day string
    vStrDate = Me.Caption(True)
    
    ' Resize to fit of needed
    If TextWidth("A") * Len(vStrDate) > ScaleWidth - m_WeekDaysHeight Then vStrDate = Me.Caption(False)
    If m_MonthPopupMode Then vStrDate = "Select Month  "
    
    ' Darw Gradients and Caption
    UserControl.Cls
    UserControl.FontBold = True
    UserControl.BackColor = m_HeaderBackCol
    If m_HeaderGradient Then PaintGradient UserControl.hDC, 0, 0, ScaleWidth, m_HeaderHeight, m_HeaderBackCol, m_HeaderGradientCol, Fill_Vertical, m_Right2Left
    DrawText UserControl.hDC, vStrDate, -1, Rct, 1
    
    ' Draw the arrow buttons
    FillStyle = vbSolid: FillColor = m_ArrowCol
    
    If m_Mode = [PopDown Mode] Then
        If m_Poped And Not m_Sensitive Then
            DrawArrow hDC, ScaleWidth - (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.6, m_HeaderHeight * 0.7, Arw_Up
        Else
            DrawArrow hDC, ScaleWidth - (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.2, m_HeaderHeight * 0.7, Arw_Down
        End If
    End If
    
    If m_Mode = [PopUp Mode] Then
        If m_Poped And Not m_Sensitive Then
            DrawArrow hDC, (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.2, m_HeaderHeight * 0.7, Arw_Down
        Else
            DrawArrow hDC, (m_HeaderHeight + 10) / 2, m_HeaderHeight * 0.6, m_HeaderHeight * 0.7, Arw_Up
        End If
    End If
    UserControl.Refresh
    
End Sub


'[ Draw the entire calendar on picCalendar ]
'-------------------------------------------
Private Sub DrawCalendar()
On Error GoTo Handle
Dim X As Long
Dim Y As Long
Dim Rct As RECT
Dim vMonth As String
Dim vStrDate As String
Dim tmpValue As Double
Dim vSelWeekDay As Long

    'Debug.Print "Drawing the Calendar"
    m_PicCalendar.Cls
    
    ' Load some the neccessary values
    m_WeekDaysHeight = m_CalendarHeight / 10

    m_TrackWidth = 0.1 * ScaleWidth
    m_iHeight = (m_PicCalendar.ScaleHeight - m_WeekDaysHeight) / (m_RowDays + 1)
    m_iWidth = (m_PicCalendar.ScaleWidth - m_TrackWidth - 1) / m_HeaderDays
    vSelWeekDay = Weekday(DateSerial(m_SelYear, m_SelMonth, m_SelDay))
    
    ' Fill the background Gradient
    m_PicCalendar.BackColor = m_CalendarBackCol
    If m_CalendarGradient Then PaintGradient m_PicCalendar.hDC, 0, 0, ScaleWidth, m_PicCalendar.ScaleHeight, m_CalendarBackCol, m_CalendarGradientCol, Fill_UpwardDiagonal, True
    
    '-------------------------------
    '| Draw Month Selection Region |
    '-------------------------------
    m_PicCalendar.FontBold = True
    m_PicCalendar.FillColor = m_MonthBackCol
    
    ' draw the gradient
    RoundRect m_PicCalendar.hDC, 0, 0, m_TrackWidth, m_PicCalendar.ScaleHeight, 0, 0
    If m_MonthGradient Then PaintGradient m_PicCalendar.hDC, 1, 1, m_TrackWidth - 2, m_PicCalendar.ScaleHeight - 2, m_MonthBackCol, m_MonthGradientCol, Fill_Horizontal, False
    
    ' Get the Month Name
    vMonth = MonthName(m_SelMonth)
    tmpValue = m_PicCalendar.TextHeight("A") * Len(vMonth)
    If tmpValue > m_PicCalendar.ScaleHeight - m_TrackWidth Then vMonth = Left$(vMonth, 3): tmpValue = m_PicCalendar.TextHeight("A") * 3
    Rct.Top = (m_PicCalendar.ScaleHeight - tmpValue) / 2
    Rct.Bottom = m_PicCalendar.ScaleHeight: Rct.Right = m_TrackWidth
    
    ' Sort Downward ( downward month print )
    For X = 1 To Len(vMonth)
        vStrDate = vStrDate & Mid$(vMonth, X, 1) & vbCrLf
    Next X
    
    ' Draw it
    DrawText m_PicCalendar.hDC, UCase$(vStrDate), -1, Rct, 1
    
    ' Draw Month Selecting Arrows
    DrawArrow m_PicCalendar.hDC, m_TrackWidth / 2, m_TrackWidth * 0.55, m_TrackWidth * 0.5, Arw_Up
    DrawArrow m_PicCalendar.hDC, m_TrackWidth / 2, m_PicCalendar.ScaleHeight - m_TrackWidth * 0.55, m_TrackWidth * 0.5, Arw_Down
    
    ' Draw the month popuping arrow
    DrawArrow m_PicCalendar.hDC, m_TrackWidth * 0.8, m_PicCalendar.ScaleHeight / 2, m_TrackWidth, Arw_Right, m_TrackWidth * 0.08
    
    '-----------------------------------
    '|Draw Horizontal Header Week Days |
    '-----------------------------------
    m_PicCalendar.FontBold = False
    tmpValue = m_TrackWidth + 1
    Rct.Top = (m_WeekDaysHeight - TextHeight("A")) / 2: Rct.Bottom = m_WeekDaysHeight
    

    Dim wDay As Long
    wDay = m_FirstDayOfWeek - 1
    
    ' Through weekdays
    For X = 1 To m_HeaderDays
        
        ' Get weekday name
        wDay = wDay + 1
        If wDay >= 8 Then wDay = 1
        vStrDate = Mid$(WeekdayName(wDay), 1, 3)
        Rct.Left = Int(tmpValue): Rct.Right = Int(tmpValue + m_iWidth - 1)
        If X = m_HeaderDays Then Rct.Right = m_PicCalendar.ScaleWidth
        
        
        Select Case wDay
            Case vSelWeekDay ' Put selected weekday Col
                m_PicCalendar.FillColor = m_WeekDaySelCol
                
            Case vbSunday ' Put sunday header Col
                m_PicCalendar.FillColor = m_WeekDaySunCol
                

            Case Else   ' Put normal week-day Col
                m_PicCalendar.FillColor = m_WeekDayCol
                
        End Select
        
        ' Draw the Week days name
        RoundRect m_PicCalendar.hDC, tmpValue, 0, Rct.Right, m_WeekDaysHeight, 0, 0
        DrawText m_PicCalendar.hDC, vStrDate, -1, Rct, 1
        tmpValue = tmpValue + m_iWidth
        
    Next X
    
    
    '------------------------------
    '|Draw Each Days in the Month |
    '------------------------------
    Dim vDayCell As Long
    Dim CurrDay As Long
    X = 1: Y = 0: wDay = 0
    If m_FirstDay < m_FirstDayOfWeek Then Y = 1: wDay = m_FirstDayOfWeek - m_FirstDay
        
    ' Through days
    For vDayCell = 1 To 35 + (m_FirstDay - m_FirstDayOfWeek) + wDay

        ' Some ordering
        If X > m_HeaderDays Then X = 1: Y = Y + 1
        If vDayCell = 36 Then X = 1: Y = 0
        If Y > 4 Then Y = 0

        ' Define rect
        Rct.Left = Int(m_TrackWidth + (X - 1) * m_iWidth + 1)
        Rct.Top = m_WeekDaysHeight + Y * m_iHeight + 1
        Rct.Bottom = Rct.Top + m_iHeight - 1
        Rct.Right = Int(Rct.Left + m_iWidth - 1)
        If X = m_HeaderDays Then Rct.Right = m_PicCalendar.ScaleWidth

        ' One additional check in applying the property 'FirstDayOfWeek'
        CurrDay = vDayCell - m_FirstDay + 1 + (m_FirstDayOfWeek - 1)
        If CurrDay >= 36 Then CurrDay = CurrDay - 35

        ' Set day Cols
        If CurrDay = m_SelDay Then

            ' Put selected day Col
            m_PicCalendar.FillColor = m_DaySelCol
        Else

            ' Check if sunday then put DaySunCol
            If (m_FirstDayOfWeek = SunDay And X = 1) Or X = (9 - m_FirstDayOfWeek) Mod 8 Then
                m_PicCalendar.FillColor = m_DaySunCol
            Else    ' Normal day Col
                m_PicCalendar.FillColor = m_DayCol
            End If
        End If

        ' Draw back
        RoundRect m_PicCalendar.hDC, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, m_Curvature, m_Curvature
        
        ' Draw Rectangle for seleted Cell
        If CurrDay = m_SelDay Then m_PicCalendar.ForeColor = m_DaySunCol: RoundRect m_PicCalendar.hDC, Rct.Left + 1, Rct.Top + 1, Rct.Right - 1, Rct.Bottom - 1, m_Curvature, m_Curvature

        ' The Day now we are drawing is opted as a 'Special Day'
        ' Represent it by an Down-Arrow in the right-top of the cell
        If IsSpecial(CurrDay, m_SelMonth) = True Then m_PicCalendar.FillColor = m_DaySunCol: DrawArrow m_PicCalendar.hDC, Rct.Left + m_iWidth * 0.75, Rct.Top + m_iHeight * 0.15, m_iHeight * 0.25, Arw_Down

        ' Days are of this month
        If Month(DateSerial(m_SelYear, m_SelMonth, CurrDay)) = m_SelMonth Then

            ' Draw it
            Rct.Top = m_WeekDaysHeight + (Y) * m_iHeight + (m_iHeight - TextHeight("A")) / 2
            DrawText m_PicCalendar.hDC, CurrDay, -1, Rct, 1
        
        ' Days are not of this month
        Else
        
            ' Fill cross
            m_PicCalendar.FillStyle = vbDiagonalCross
            If (m_FirstDayOfWeek = SunDay And X = 1) Or X = (9 - m_FirstDayOfWeek) Mod 8 Then m_PicCalendar.FillColor = m_DayCol Else m_PicCalendar.FillColor = m_DaySunCol
            RoundRect m_PicCalendar.hDC, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, m_Curvature, m_Curvature
            m_PicCalendar.FillStyle = vbSolid
            
        End If
        X = X + 1

    Next vDayCell

    '----------------------------------
    '| Draw The Year Selection Region |
    '----------------------------------
    Rct.Top = m_PicCalendar.ScaleHeight - m_iHeight + 1
    Rct.Bottom = m_PicCalendar.ScaleHeight
    Rct.Left = m_PicCalendar.ScaleWidth - 3 * m_iWidth + 1
    m_PicCalendar.FillColor = m_YearBackCol
    RoundRect m_PicCalendar.hDC, Rct.Left, Rct.Top, m_PicCalendar.ScaleWidth, Rct.Bottom, 0, 0
    If m_YearGradient Then PaintGradient m_PicCalendar.hDC, Rct.Left + 1, Rct.Top + 1, 3 * m_iWidth - 3, (Rct.Bottom - Rct.Top) - 2, m_YearBackCol, m_YearGradientCol, Fill_Vertical, True

    ' Define Rect
    m_PicCalendar.FontBold = True
    Rct.Left = m_TrackWidth + 1
    Rct.Right = m_TrackWidth + 4 * m_iWidth
    Rct.Top = m_PicCalendar.ScaleHeight - m_iHeight + (m_iHeight - TextHeight("A")) / 2
    
    ' Draw Today
    DrawText m_PicCalendar.hDC, "Today " & Format$(Date$, GetFormat), -1, Rct, 1
    
    ' Draw year
    Rct.Left = m_PicCalendar.ScaleWidth - 3 * m_iWidth + 1
    Rct.Right = m_PicCalendar.ScaleWidth
    DrawText m_PicCalendar.hDC, Format$(m_SelYear, "0000"), -1, Rct, 1
    
    ' Draw year selecting arrows
    DrawArrow m_PicCalendar.hDC, Rct.Right - m_iWidth / 2, Rct.Bottom - m_iHeight / 2, m_iHeight * 0.5, Arw_Right
    DrawArrow m_PicCalendar.hDC, Rct.Left + m_iWidth / 2, Rct.Bottom - m_iHeight / 2, m_iHeight * 0.5, Arw_Left

Handle:
End Sub


'[ Export the calendar to the parent form ]
'------------------------------------------
Private Sub ExportCalendar(ByVal vDown As Boolean, Optional vFromClass As Boolean = False)
Dim Rct1 As RECT
Dim Rct2 As RECT
Dim PicPos As POINTAPI
Dim hParent As Long

    'Debug.Print "Exporting calendar "
    
    'Check Mode
    If m_Mode = [List Mode] Then ImportCalendar: Exit Sub
    If Not Ambient.UserMode Then Exit Sub
    m_Poped = True
    hParent = GetParent(UserControl.Parent.hWnd)
    
    ' Get usercontrol Rect
    GetWindowRect UserControl.hWnd, Rct1
    PicPos.X = Rct1.Left
    If vDown Then PicPos.Y = Rct1.Bottom - 1 Else PicPos.Y = Rct1.Top - m_CalendarHeight

    If UserControl.Parent.MDIChild Then
        GetWindowRect hParent, Rct2
        PicPos.X = PicPos.X - Rct2.Left - 2
        PicPos.Y = PicPos.Y - Rct2.Top - 2
        hParent = GetParent(hParent)
    End If

    ' export the calendar
    SetParent m_PicCalendar.hWnd, hParent
    
    ' do the move
    SetWindowPos m_PicCalendar.hWnd, UserControl.Parent.hWnd, PicPos.X, PicPos.Y, Rct1.Right - Rct1.Left, m_CalendarHeight, SWP_SHOWWINDOW
        
    m_PicCalendar.Visible = True
    If vFromClass Then Exit Sub
    If m_Animate Then
        If m_Mode = [PopDown Mode] Then
            AnimateForm m_PicCalendar, aload, eAppearFromRight, 10, 22
        Else
            AnimateForm m_PicCalendar, aload, eAppearFromLeft, 10, 22
        End If
    Else
        ' This is necessary to redefine the Calendar region to full size
        AnimateForm m_PicCalendar, aload, eAppearFromLeft, 0, 1
    End If
    
End Sub


'[ Import the exported calendar to Control on ListMode ]
'-------------------------------------------------------
Private Sub ImportCalendar()

    'Debug.Print "Importing calendar "
    
    ' Set the new parent and Bring on top
    SetParent m_PicCalendar.hWnd, UserControl.hWnd
    
    m_PicCalendar.Move -1, m_HeaderHeight, ScaleWidth + 2, m_CalendarHeight
    
    m_PicCalendar.Visible = True
    
End Sub


'[ Close the calendar with animation ]
'-------------------------------------
Public Sub CollapseCalendar(Optional vAnimate As Boolean = True)

    If m_Mode = [List Mode] Then Exit Sub
    If m_Animate And vAnimate Then
        If m_Mode = [PopDown Mode] Then
            AnimateForm m_PicCalendar, aUnload, eAppearFromBottom, 10, 22
        Else
            AnimateForm m_PicCalendar, aUnload, eAppearFromTop, 10, 22
        End If
    End If
    m_PicCalendar.Visible = False
    
    ' This is necessary to redefine the Calendar region to full size
    AnimateForm m_PicCalendar, aload, eAppearFromLeft, 0, 1
    m_Poped = False
    
End Sub


'[ Localy made function to Draw arrows ]
'---------------------------------------
Private Sub DrawArrow(hDC As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        ByVal vSize As Long, _
                        ByVal vArrow As ArrowDir, _
                        Optional vThickness As Long = -1)
Dim Pnts(2) As POINTAPI
    
    ' Nothing here, define a point arry, fill it
    m_PicCalendar.FillColor = m_ArrowCol
    If vThickness = -1 Then
        vThickness = vSize / 2
    Else
        ' Special case of popup month button
        m_PicCalendar.FillColor = m_HeaderBackCol:
    End If
    
    ' Self explonatory
    If vArrow = Arw_Left Or vArrow = Arw_Right Then
        Pnts(0).X = X: Pnts(0).Y = Y - vSize / 2
        Pnts(1).X = X: Pnts(1).Y = Y + vSize / 2
        Pnts(2).Y = Y
        If vArrow = Arw_Left Then Pnts(2).X = X - vThickness Else Pnts(2).X = X + vThickness
    Else
        Pnts(0).X = X - vSize / 2: Pnts(0).Y = Y
        Pnts(1).X = X + vSize / 2: Pnts(1).Y = Y
        Pnts(2).X = X
        If vArrow = Arw_Down Then Pnts(2).Y = Y + vThickness Else Pnts(2).Y = Y - vThickness
    End If
    
    ' draw it
    Polygon hDC, Pnts(0), 3
    
End Sub


'[ Checks if the Day is opted as a special day ]
'-----------------------------------------------
Private Function IsSpecial(ByVal vDay As Long, ByVal vMonth As Long) As Boolean
Dim X As Long
Dim xMax As Long
Dim vDayID As String
    
    On Error GoTo Handle
    
    ' This function is used to check whether the give day is
    ' loaded as a special day. The special days are already loaded to m_SpecialdaysStack
    If m_SpecialDays = vbNullString Or vDay = -1 Then IsSpecial = False: Exit Function
    vDayID = vDay & "-" & vMonth
    xMax = UBound(m_SpecialDayStack)
    m_SpecialDayString = vbNullString
    
    ' Loop through all the days in specialday stack
    For X = 0 To xMax
        If InStr(1, m_SpecialDayStack(X), ">") = 0 Then
            If m_SpecialDayStack(X) = vDayID Then IsSpecial = True: Exit Function
        Else
            If Split(m_SpecialDayStack(X), ">")(0) = vDayID Then IsSpecial = True: m_SpecialDayString = Split(m_SpecialDayStack(X), ">")(1): Exit Function
        End If
    Next X
    
Handle:
    IsSpecial = False
    
End Function


'[ Get the format string for the specified DateFormat ]
'------------------------------------------------------
Private Function GetFormat() As String
Select Case m_DateFormat
    Case 0
        GetFormat = "dd-mm-yyyy"
    Case 1
        GetFormat = "mm-dd-yyyy"
    Case 2
        GetFormat = "yyyy-mm-dd"
End Select
End Function


'[ All the checks for the new day calulated from mouseclick event ]
'------------------------------------------------------------------
Private Sub LoadDay(ByVal nDay As Long)
Dim dDate As String, lnDay As String

    lnDay = nDay
    If m_SelMonth < 1 Then
        m_SelMonth = 12
        m_SelYear = m_SelYear - 1
    ElseIf m_SelMonth > 12 Then
        m_SelMonth = 1
        m_SelYear = m_SelYear + 1
    End If
    
    ' calculate Days in the month
    dDate = DateSerial(m_SelYear, m_SelMonth, 1)
    m_MonthDays = DateDiff("d", dDate, DateAdd("m", 1, dDate))
    m_FirstDay = Weekday(DateSerial(m_SelYear, m_SelMonth, 1))
    
    ' Special case Currently selected date is over the daycount
    If nDay = 999 Then
        If m_SelDay >= m_MonthDays Then m_SelDay = m_MonthDays: Exit Sub Else nDay = m_SelDay
    End If
    
    ' A new line added to fix 'FirstDayOfWeek' property
    If m_FirstDay < m_FirstDayOfWeek Then nDay = nDay - 7
    
    ' Special case moving last day (Drawn first)
    ' Take July of 2005 and look 31 is on top
    If nDay < 0 Then
        nDay = (35 - m_FirstDay) + (m_FirstDay + nDay)
    End If
    
    ' Cross filled cell was selected. Collapse if Sensitive else Skip to another Month
    If (nDay <= 0 Or nDay > m_MonthDays) Then
        If m_Sensitive And Not m_Mode = [List Mode] Then CollapseCalendar: Exit Sub
        
        If m_SkipEnabled Then
            If nDay <= 0 Then m_SelMonth = m_SelMonth - 1: LoadDay (999): Exit Sub
            If nDay > m_MonthDays Then m_SelMonth = m_SelMonth + 1: LoadDay (999): Exit Sub
        Else
            Exit Sub
        End If
    
    End If
    

    ' nO PROBLEM Load the day
    m_SelDay = nDay
    If (lnDay <> 999) Then
        RaiseEvent DateChanged
    End If
End Sub


'[ Prepare the control for selecting a month from a monthlist ]
'--------------------------------------------------------------
Private Sub PopupMonthList()
Dim X As Long
Dim Rct As RECT
Dim mName As String
Dim Y As Long
Dim vText As String
    
    ' Enable Popup mode\Define rect
    m_MonthPopupMode = True
    m_MonthPopWidth = (m_PicCalendar.ScaleWidth - m_TrackWidth) / 12
    Rct.Bottom = m_PicCalendar.ScaleHeight: Rct.Top = 0
    m_PicCalendar.FontBold = False
    m_PicCalendar.FillStyle = vbSolid
    m_PicCalendar.FillColor = m_HeaderBackCol
    
    ' through months
    For X = 1 To 12
    
        ' Get month name
        vText = vbNullString
        mName = UCase$(MonthName(X))
        
        ' Sort it Downward
        For Y = 1 To Len(mName)
            vText = vText & Mid$(mName, Y, 1) & vbCrLf
        Next Y
        
        ' Draw it
        Rct.Left = m_TrackWidth + Int((X - 1) * m_MonthPopWidth) - 1
        Rct.Right = Int(Rct.Left + m_MonthPopWidth) + 2
        RoundRect m_PicCalendar.hDC, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, 0, 0
        DrawText m_PicCalendar.hDC, vText, -1, Rct, 1
        
    Next X
    m_PicCalendar.Refresh
    
End Sub


'[ Stimulate Popup calendar externaly ]
'-------------------------------------------------------------------------------------------------------------------------

Public Sub PopUpCalendar()
    If m_Mode = [PopDown Mode] Then ExportCalendar True
    If m_Mode = [PopUp Mode] Then ExportCalendar False
End Sub


'[ Create controls dynamically ]
'-------------------------------
Private Sub CreateControls()

    ' Add the Calendar
    Set m_PicCalendar = UserControl.Controls.Add("vb.picturebox", "PicCalendar")
    m_PicCalendar.AutoRedraw = True
    m_PicCalendar.Visible = True
    m_PicCalendar.MouseIcon = UserControl.MouseIcon
    m_PicCalendar.ScaleMode = vbPixels
    m_PicCalendar.Appearance = Flat
    m_PicCalendar.BorderStyle = None
    
    ' Hide the calendar from taskbar
    SetWindowLong m_PicCalendar.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    
    ' Add the tooltip
    Set m_PicToolTip = UserControl.Controls.Add("vb.picturebox", "PicToolTip")
    m_PicToolTip.AutoRedraw = True
    m_PicToolTip.ScaleMode = vbPixels
    m_PicToolTip.Appearance = Flat
    m_PicToolTip.BorderStyle = 1
    m_PicToolTip.Enabled = False
    
    ' I set the usercontrol backcolor to TooltipColor
    ' So we can get it directly
    m_PicToolTip.BackColor = UserControl.BackColor
    
    ' Hide the tooltip from taskbar
    SetWindowLong m_PicToolTip.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    
End Sub


'-------------------------------------------------------------------------------------------------------------------------
' Procedure : ShowToolTip
' Auther    : Jim Jose
' Input     : None
' OutPut    : None
' Purpose   : We have to have local tooltip generation bcose Vb's tooltips
'           : will not support Unicode languages
'-------------------------------------------------------------------------------------------------------------------------

Private Sub ShowToolTip()
Dim vWidth As Long
Dim vHeight As Long
Dim hParent As Long
Dim Rct   As RECT
Dim mPos   As POINTAPI

    Exit Sub
    'Debug.Print "Show Tip"
    If m_ToolTipText = vbNullString Then Exit Sub
    If Not Ambient.UserMode Then Exit Sub
    
    ' Get parent
    If UserControl.Parent.MDIChild Then
        hParent = GetParent(GetParent(UserControl.Parent.hWnd))
    Else
        hParent = GetParent(UserControl.Parent.hWnd)
    End If
    
    ' Export tooltip
    SetParent m_PicToolTip.hWnd, hParent
    
    ' Get size
    vHeight = m_PicToolTip.TextHeight("A") * 1.2
    vWidth = m_PicToolTip.TextWidth(m_ToolTipText)

    ' Get mose pos
    GetCursorPos mPos

    ' Move tooltip
    SetWindowPos m_PicToolTip.hWnd, UserControl.Parent.hWnd, mPos.X, mPos.Y + 20, vWidth + 15, vHeight, SWP_SHOWWINDOW

    ' Draw the tooltip text
    Rct.Left = 5: Rct.Top = 0
    Rct.Right = vWidth + 15: Rct.Bottom = vHeight
    m_PicToolTip.Cls

    m_PicToolTip.ForeColor = vbBlack
    DrawText m_PicToolTip.hDC, m_ToolTipText, -1, Rct, 1
    DrawArrow m_PicToolTip.hDC, 3, vHeight / 2.4, 9, Arw_Right, 4
    
    m_PicToolTip.ForeColor = m_DaySunCol
    RoundRect m_PicToolTip.hDC, 0, 0, vWidth + 15 - 2, vHeight - 2, 0, 0
    
    ' Show it
    m_PicToolTip.Refresh
    m_PicToolTip.Visible = True
        
    m_ToolTipText = vbNullString
    
End Sub


'-------------------------------------------------------------------------------------------------------------------------
' Procedure : DrawText
' Auther    : Dana Seaman (only slight modification to suit this project)
' Input     : hDC + Parameters
' OutPut    : None
' Purpose   : DrawText with unicode support
'-------------------------------------------------------------------------------------------------------------------------

Private Sub DrawText(ByVal hDC As Long, _
                        ByVal lpStr As String, ByVal nCount As Long, _
                        ByRef lpRect As RECT, ByVal wFormat As Long)
    ' Set to forecolor
    m_PicCalendar.ForeColor = m_ForeColor
    UserControl.ForeColor = m_ForeColor
    
    ' Draw the text
    If IsNT Then
       DrawTextW hDC, StrPtr(lpStr), nCount, lpRect, wFormat
    Else
       DrawTextA hDC, lpStr, nCount, lpRect, wFormat
    End If
    
    ' Set to Bordercolor. Because RoundRect Api uses forecolor as bordercolor
    m_PicCalendar.ForeColor = m_BorderColor
    UserControl.ForeColor = m_BorderColor
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procedure : IsNT
' Auther    : Dana Seaman
' Input     : None
' OutPut    : NT?
' Purpose   : Check for the NT Platform
'-------------------------------------------------------------------------------------------------------------------------

Private Function IsNT() As Boolean
   Static m_bInit As Boolean
   Dim udtVer           As OSVERSIONINFO
   
   On Error Resume Next
   'Cache m_bIsNT on first execution
   If Not m_bInit Then
      m_bInit = True
      udtVer.dwOSVersionInfoSize = Len(udtVer)
      If GetVersionEx(udtVer) Then
         If udtVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            m_bIsNT = True
         End If
      End If
   End If
   IsNT = m_bIsNT
   
End Function


'-------------------------------------------------------------------------------------------------------------------------
' Procedure : PaintGradient
' Auther    : Carls P.V.
' Input     : hDC + Parameters
' OutPut    : None
' Purpose   : DIB solution for fast gradients
'-------------------------------------------------------------------------------------------------------------------------

Private Sub PaintGradient(ByVal hDC As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionCts, _
                         Optional Right2Left As Boolean = True)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  Dim tmpCol  As Long
  
  
    '-- A minor check
    If GradientDirection = Fill_None Then Exit Sub
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    If Right2Left Then
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    '-- Decompose Cols
    Col1 = Col1 And &HFFFFFF
    R1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    G1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    B1 = Col1 Mod &H100&
    Col2 = Col2 And &HFFFFFF
    R2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    G2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    B2 = Col2 Mod &H100&
    
    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To Width - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [Fill_Vertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDC, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procedure : AnimateForm
' Auther    : Jim Jose
' Input     : Animating Object + Parameters
' OutPut    : None
' Purpose   : Animate the hWndObject with different animation effects.
'-------------------------------------------------------------------------------------------------------------------------

Private Function AnimateForm(hWndObject As Object, ByVal aEvent As AnimeEventEnum, _
                            Optional ByVal aEffect As AnimeEffectEnum = 11, _
                            Optional ByVal FrameTime As Long = 1, _
                            Optional ByVal FrameCount As Long = 33) As Boolean
On Error GoTo Handle
Dim X1 As Long, Y1 As Long
Dim hrgn As Long, tmpRgn As Long
Dim XValue As Long, YValue As Long
Dim XIncr As Double, YIncr As Double
Dim wHeight As Long, wWidth As Long

    wWidth = hWndObject.Width / Screen.TwipsPerPixelX
    wHeight = hWndObject.Height / Screen.TwipsPerPixelY
'    hWndObject.Visible = True
    
    Select Case aEffect
    
        Case eAppearFromLeft
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                XValue = X1 * XIncr
                hrgn = CreateRectRgn(0, 0, XValue, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hWndObject.hWnd, hrgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eAppearFromRight
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
                
                ' Define the size of current frame/Create it
                XValue = wWidth - X1 * XIncr
                hrgn = CreateRectRgn(XValue, 0, wWidth, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hWndObject.hWnd, hrgn, True:  DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eAppearFromTop
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                YValue = Y1 * YIncr
                hrgn = CreateRectRgn(0, 0, wWidth, YValue)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hWndObject.hWnd, hrgn, True:   DoEvents
                Sleep FrameTime
                
            Next Y1
            
        Case eAppearFromBottom
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                YValue = wHeight - Y1 * YIncr
                hrgn = CreateRectRgn(0, YValue, wWidth, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hrgn, hrgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hWndObject.hWnd, hrgn, True: DoEvents
                Sleep FrameTime
                
            Next Y1
    End Select

    AnimateForm = True
    
Exit Function
Handle:
    AnimateForm = False
End Function


'----------------------------------------------------------------------------------------------------------------------------
' The following bytes are donated exclusively for Paul Caton's Subclassing
' We need this to track the movement information of the m_picCalendar and
' sizing/positioning of parent form
'----------------------------------------------------------------------------------------------------------------------------
' Auther    : Paul Caton
' Purpose   : Advanced subclassing for usercontrols (Self subclasser)
' Comment   : Ooooh maan!!! How could u made this???
'           : Thanks a Billion for this ever green piece of code on subclassing
'----------------------------------------------------------------------------------------------------------------------------

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'Return the upper 16 bits of the passed 32 bit value
Private Function WordHi(lngValue As Long) As Long
  If (lngValue And &H80000000) = &H80000000 Then
    WordHi = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
  Else
    WordHi = (lngValue And &HFFFF0000) \ &H10000
  End If
End Function

'Return the lower 16 bits of the passed 32 bit value
Private Function WordLo(lngValue As Long) As Long
  WordLo = (lngValue And &HFFFF&)
End Function

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
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
    Call FreeLibrary(hMod)
  End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hWndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub
