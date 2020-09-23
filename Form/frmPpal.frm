VERSION 5.00
Begin VB.Form frmPpal 
   BackColor       =   &H00DAD4CD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PropertyGrid Demo"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6765
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmPpal.frx":038A
   ScaleHeight     =   6375
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddToList 
      Caption         =   "Add To List"
      Height          =   360
      Left            =   4185
      TabIndex        =   6
      Top             =   135
      Width           =   1080
   End
   Begin PropertyGridDemo.PropertyGrid PropertyGrid 
      Height          =   3330
      Left            =   105
      TabIndex        =   0
      Top             =   435
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5874
      AutoFilter      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineColor       =   8421504
      SetColorStyle   =   5
      StylePropertyGrid=   1
      ViewForeColor   =   8421504
   End
   Begin PropertyGridDemo.PropertyGrid PropertyGrid1 
      Height          =   2190
      Left            =   120
      TabIndex        =   5
      Top             =   4125
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3863
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HelpVisible     =   0   'False
      LineColor       =   8421504
      StylePropertyGrid=   3
      ViewCategoryForeColor=   255
      ViewForeColor   =   4210752
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dedicated to the prettiest girl that I have known XD, T.Q.M mibi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   675
      Index           =   3
      Left            =   4155
      TabIndex        =   4
      Top             =   5370
      Width           =   2475
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Based in the PropertyGrid Of Visual Basic 6.0. Any comment is wellcome and votes is great for me :)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   870
      Index           =   2
      Left            =   4155
      TabIndex        =   3
      Top             =   4410
      Width           =   2475
   End
   Begin VB.Image imgCatg 
      Height          =   255
      Index           =   1
      Left            =   465
      MouseIcon       =   "frmPpal.frx":CA46
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":CD50
      Top             =   135
      Width           =   225
   End
   Begin VB.Image imgCatg 
      Height          =   255
      Index           =   0
      Left            =   150
      MouseIcon       =   "frmPpal.frx":CDE1
      MousePointer    =   99  'Custom
      Picture         =   "frmPpal.frx":D0EB
      Top             =   135
      Width           =   240
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPpal.frx":D17B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1080
      Index           =   1
      Left            =   4155
      TabIndex        =   2
      Top             =   3270
      Width           =   2475
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result after of select in the First PropertyGrid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3870
      Width           =   3855
   End
   Begin VB.Image imgLogo 
      Height          =   2010
      Left            =   3930
      Picture         =   "frmPpal.frx":D209
      Top             =   450
      Width           =   2985
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  Dim Credits(9) As String
  Dim PSCContrib(28) As String

Private Sub cmdAddToList_Click()
     
     PropertyGrid.ValueChildItemChanged "PSC Community", "Key3", PSCContrib, PSCContrib(0) & ", " & PSCContrib(4)
     PropertyGrid.AddToList "Key3", "PSC Community", "Nuevo Valor"
     PropertyGrid.Refresh True
     
End Sub

Private Sub Form_Load()
  
  Dim lRandom        As String
  Dim xSort(2)       As String
  Dim lFont          As New StdFont
    
    Show
    Randomize
    lRandom = Int((UBound(PSCContrib) * Rnd(Timer)))  ' Random Number between 0 To
    '   UBound(PSCContrib).
    xSort(0) = "Categorized"
    xSort(1) = "Alphabetical"
    xSort(2) = "NoSort"

    PSCContrib(0) = "IVan"
    PSCContrib(1) = "Fred.cpp"
    PSCContrib(2) = "Option Explicit"
    PSCContrib(3) = "Richard Mewett"
    PSCContrib(4) = "Jim Jose"
    PSCContrib(5) = "Paul Caton"
    PSCContrib(6) = "Paul R. Territo"
    PSCContrib(7) = "Carles P.V"
    PSCContrib(8) = "Habin"
    PSCContrib(9) = "Steppenwolfe"
    PSCContrib(10) = "Kelvin C. Perez"
    PSCContrib(11) = "Juan Carlos San Román"
    PSCContrib(12) = "D. Rijmenants"
    PSCContrib(13) = "Gary Noble"
    PSCContrib(14) = "Dreamvb"
    PSCContrib(15) = "Selftaught"
    PSCContrib(16) = "Light Templer"
    PSCContrib(17) = "Luthfi"
    PSCContrib(18) = "Mario Flores"
    PSCContrib(19) = "Mario Villanueva"
    PSCContrib(20) = "rm_code"
    PSCContrib(21) = "Agustin Rodriguez"
    PSCContrib(22) = "Robert Rayment"
    PSCContrib(23) = "Ulli's"
    PSCContrib(24) = "LaVolpe"
    PSCContrib(25) = "Dean Camera"
    PSCContrib(26) = "Evan Toder"
    PSCContrib(27) = "Kinex"
    PSCContrib(28) = "int21"

    Credits(0) = PSCContrib(5) & " CodeId=54117"
    Credits(1) = "Steve McMahon" & " www.vbaccelerator.com"
    Credits(2) = PSCContrib(3) & " CodeId=61438"
    Credits(3) = PSCContrib(6) & " CodeId=63905"
    Credits(4) = PSCContrib(7) & " CodeId=29586"
    Credits(5) = PSCContrib(1) & " CodeId=61476"
    Credits(6) = PSCContrib(4) & " CodeId=60901"
    Credits(7) = PSCContrib(4) & " CodeId=61435"
    Credits(8) = "Calendar CodeId=61147"
    Credits(9) = PSCContrib(2) & " CodeId=60849"

    With PropertyGrid
        .AddCategory "Key2", "Properties", "Properties of this control", True
        .AddCategory "Key3", "About", "", True

        .AddChildItem "Key3", PropertyItemStringList, "PSC Community", PSCContrib, _
           PSCContrib(lRandom), "People that contribute in this community. #(^~^)#", , "MS"

        .AddChildItem "Key2", PropertyItemStringReadOnly, "(Name)", .Name
        .AddChildItem "Key2", PropertyItemColor, "BackColor", .BackColor, , , , "RA"
        .AddChildItem "Key2", PropertyItemFont, "Font", .Font
        .AddChildItem "Key2", PropertyItemColor, "HelpBackColor", .HelpBackColor
        .AddChildItem "Key2", PropertyItemColor, "HelpForeColor", .HelpForeColor
        .AddChildItem "Key2", PropertyItemNumber, "HelpHeight", .HelpHeight
        .AddChildItem "Key2", PropertyItemBool, "HelpVisible", .HelpVisible
        .AddChildItem "Key2", PropertyItemColor, "LineColor", .LineColor, , , , "FC:" & &H8080&
        .AddChildItem "Key2", PropertyItemStringList, "PropertySort", xSort, xSort(.PropertySort), , , "FC:" & vbBlue
        .AddChildItem "Key2", PropertyItemNumber, "SplitterPos", .SplitterPos, , , , "MX:4"
        .AddChildItem "Key2", PropertyItemColor, "ViewBackColor", .ViewBackColor
        .AddChildItem "Key2", PropertyItemColor, "ViewCategoryForeColor", .ViewCategoryForeColor
        .AddChildItem "Key2", PropertyItemColor, "ViewForeColor", .ViewForeColor
        .AddChildItem "Key2", PropertyItemCheckBox, "CheckBox1", True, True

        .AddChildItem "Key3", PropertyItemForm, "Author", "Heriberto Mantilla Santamaría", _
            , "All rights Reserved © HACKPRO TM 2006", , "WB"
        .AddChildItem "Key3", PropertyItemStringList, "Credits", Credits, Credits(0), "Credits And" & _
            " Thanks"
        .AddChildItem "Key3", PropertyItemStringReadOnly, "Dedication", "Mibi", , "Dedicated to the" & _
            " prettiest girl that I have known XD, T.Q.M mibi", , "FC:" & vbRed
            
        .AddChildItem "Key3", PropertyItemUpDown, "Version", "1", , , , Year(Now) - 10 & ":" & Year(Now) + 30
        .AddChildItem "Key3", PropertyItemFolderFile, "Folder File", ""
        .AddChildItem "Key3", PropertyItemFolder, "Folder", "C:\"
        
        .ValueChildItemChanged "Version", "Key3", Year(Now)
        
        .Refresh

    End With

    With PropertyGrid1

        .AddCategory "Key1", "PSC Community", "", False
        .AddCategory "Key2", "Properties", "Properties of this control", False
        .AddCategory "Key3", "PropertyGrid History", "", True

        .AddChildItem "Key1", PropertyItemStringList, "PSC Community", PSCContrib, _
            PSCContrib(lRandom), "People that contribute in this community. #(^~^)#"
        
        .AddChildItem "Key3", PropertyItemDate, "Update", CDate("12/12/2006"), , "Date Last Update"
        '.AddChildItem "Key3", PropertyItemDate, "Update1", CDate("14/12/2006"), , "Date Last Update"
        .AddChildItem "Key3", PropertyItemStringReadOnly, "Added", "New Theme, Style ComboBox and Button, McCalendar.", , "Fixed Bugs and new things."
        .AddChildItem "Key3", PropertyItemString, "Control Version", .GetControlVersion, , "Actual Version", , "LK"
        .AddChildItem "Key3", PropertyItemStringReadOnly, "Testing in", "Win98SE & WinXP SP2", , "Debugging OS"
        .AddChildItem "Key3", PropertyItemStringReadOnly, "Compatibility", "Now work in VB 5.0 & VB 6.0", , "Working in VB Language."
        .AddChildItem "Key3", PropertyItemStringReadOnly, "Dedication", "Mibi", , "Dedicated to the" & _
            " prettiest girl that I have known XD, T.Q.M mibi", , "FC:" & vbRed
        .AddChildItem "Key3", PropertyItemCheckBox, "CheckBox2", True, False, "CheckBox Item"

        .AddChildItem "Key2", PropertyItemStringReadOnly, "(Name)", .Name
        .AddChildItem "Key2", PropertyItemColor, "BackColor", .BackColor
        .AddChildItem "Key2", PropertyItemFont, "Font", .Font
        .AddChildItem "Key2", PropertyItemColor, "HelpBackColor", .HelpBackColor
        .AddChildItem "Key2", PropertyItemColor, "HelpForeColor", .HelpForeColor
        .AddChildItem "Key2", PropertyItemNumber, "HelpHeight", .HelpHeight, "", "", , "FC:" & vbBlue
        .AddChildItem "Key2", PropertyItemBool, "HelpVisible", .HelpVisible
        .AddChildItem "Key2", PropertyItemColor, "LineColor", .LineColor
        .AddChildItem "Key2", PropertyItemStringList, "PropertySort", xSort, xSort(.PropertySort)
        .AddChildItem "Key2", PropertyItemNumber, "SplitterPos", .SplitterPos
        .AddChildItem "Key2", PropertyItemColor, "ViewBackColor", .ViewBackColor
        .AddChildItem "Key2", PropertyItemColor, "ViewCategoryForeColor", .ViewCategoryForeColor
        .AddChildItem "Key2", PropertyItemColor, "ViewForeColor", .ViewForeColor
        
        
        .HelpForeColor = &H40C0&
        .HelpBackColor = &HC0FFFF
        .LineColor = &H808080
        .ViewBackColor = &HFFFFFF
        .ViewForeColor = &H404040
        .ViewCategoryForeColor = &H404040
        .isButtonColors .ViewBackColor, .ViewForeColor, .ViewForeColor, .ViewCategoryForeColor
        
        ' Sets the font used to display text in the PropertyGrid to "Tahoma"
        With lFont
            .Name = "Tahoma"
            .Bold = False
        End With
        
        Set .Font = lFont
        
        .ChildItemChanged .GetChildIndex("HelpVisible", "Key2"), "Key2", "Key2", PropertyItemBool, "HelpVisible1", True
        
        '.DelAllChildToCatg "Key2"

        .Refresh
        
    End With
        
End Sub

Private Sub imgCatg_Click(Index As Integer)

    With PropertyGrid

        If (.PropertySort <> Index) Then
            .PropertySort = Index
            .Refresh
        End If

    End With
    
    'PropertyGrid.ValueChildItemChanged "(Name)", "Key2", ""

End Sub

Private Sub PropertyGrid_FormClick(Value As Variant, ByVal KeyCategory As String, _
    ByVal Title As String, ByVal X As Integer, ByVal Y As Integer)
    
    Value = "Demo Form"
    
End Sub

Private Sub PropertyGrid_ValueChanged(ByVal KeyCategory As String, _
                                      ByVal Title As String, _
                                      ByVal Value As Variant, _
                                      ByVal theFont As StdFont)

    If (KeyCategory = "Key2") Then

        With PropertyGrid1

            Select Case Title
            Case "BackColor"
                .BackColor = Value

            Case "Font"
                Set .Font = theFont

            Case "HelpBackColor"
                .HelpBackColor = Value

            Case "HelpForeColor"
                .HelpForeColor = Value

            Case "HelpHeight"
                .HelpHeight = Value

            Case "HelpVisible"
                .HelpVisible = Value

            Case "LineColor"
                .LineColor = Value

            Case "PropertySort"
                Dim XValue As Integer

                If (Value = "Categorized") Then
                    XValue = &H0
                ElseIf (Value = "Alphabetical") Then
                    XValue = &H1
                Else
                    XValue = &H2
                End If

                .PropertySort = XValue

            Case "SplitterPos"
                .SplitterPos = Value

            Case "ViewBackColor"
                .ViewBackColor = Value

            Case "ViewCategoryForeColor"
                .ViewCategoryForeColor = Value

            Case "ViewForeColor"
                .ViewForeColor = Value
            End Select

            .Refresh
            
            
        End With

    Else
        'MsgBox PropertyGrid.GetChildValue(KeyCategory, Title)
    End If

End Sub
