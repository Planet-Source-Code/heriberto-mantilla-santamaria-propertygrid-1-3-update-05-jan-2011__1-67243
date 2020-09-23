VERSION 5.00
Begin VB.Form frmTemp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Demo"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin PropertyGridDemo.SComboBox SComboBox1 
      Height          =   345
      Left            =   1920
      Top             =   3825
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
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
   Begin PropertyGridDemo.PropertyGrid PropertyGrid1 
      Height          =   3105
      Left            =   225
      TabIndex        =   0
      Top             =   405
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   5477
      BackColor       =   8421504
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
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
     With PropertyGrid1

        .AddCategory "Key1", "Datos Generales", "Contiene la información del proyecto.", True
        .AddCategory "Key2", "Porcentajes", "Valores usados para los Unitarios."
        .AddCategory "Key3", "Seguridad", "Permite proteger la licitación del uso de terceros."

        .AddChildItem "Key1", PropertyItemString, "NIT ó C.C. Nº", "", , "NIT de la" & _
            " empresa.", ItemUpperCase
        .AddChildItem "Key1", PropertyItemString, "Empresa Contratista", "", , "Empresa" & _
            " licitante.", ItemUpperCase
        .AddChildItem "Key1", PropertyItemString, "Nombre del Proyecto", "", , "Nombre de" & _
            " la licitación", ItemUpperCase

        .AddChildItem "Key1", PropertyItemString, "Número del Proyecto", "", , "Número" & _
                " asignado al proyecto.", ItemUpperCase
        
        .AddChildItem "Key2", PropertyItemNumber, "Administración (%)", "", , "Este valor" & _
            " corresponde al porcentaje de administración con respecto al A.I.U.", ItemNumeric
        .AddChildItem "Key2", PropertyItemNumber, "Imprevistos (%)", "", , "Este valor" & _
            " corresponde al porcentaje de imprevistos con respecto al A.I.U.", ItemNumeric
        .AddChildItem "Key2", PropertyItemNumber, "Utilidad (%)", "", , "Este valor" & _
            " corresponde al porcentaje de utilidad con respecto al A.I.U.", ItemNumeric
        .AddChildItem "Key2", PropertyItemNumber, "IVA (%)", "", , "Este valor corresponde" & _
            " al IVA", ItemNumeric

        .AddChildItem "Key3", PropertyItemString, "Contraseña", "", , "Contraseña asignada" & _
            " al proyecto.", ItemPassword
        
        .HelpVisible = True
        .SplitterPos = 0.75
        '.FixedSplit = True
        .StylePropertyGrid = NormalTheme

        .Refresh

    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    PropertyGrid1.Move 0, 0, ScaleWidth, ScaleHeight
    PropertyGrid1.Refresh
On Error GoTo 0
End Sub
