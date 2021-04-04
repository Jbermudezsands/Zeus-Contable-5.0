VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fcalendario 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   Icon            =   "fcalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   3465
      Left            =   -30
      TabIndex        =   0
      Top             =   -150
      Width           =   3165
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1155
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   3075
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   2220
            Top             =   1080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "fcalendario.frx":030A
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   390
            Left            =   2730
            TabIndex        =   66
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList1"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Limpiar Fecha"
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Ene"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   420
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Feb"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   870
            TabIndex        =   13
            Top             =   420
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Mar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   1650
            TabIndex        =   12
            Top             =   420
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Abr"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   2400
            TabIndex        =   11
            Top             =   420
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "May"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   660
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Jun"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   870
            TabIndex        =   9
            Top             =   660
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Jul"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   1650
            TabIndex        =   8
            Top             =   660
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Ago"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   2400
            TabIndex        =   7
            Top             =   660
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Sep"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   6
            Top             =   900
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Oct"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   870
            TabIndex        =   5
            Top             =   900
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Nov"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   1650
            TabIndex        =   4
            Top             =   900
            Width           =   615
         End
         Begin VB.OptionButton mes 
            BackColor       =   &H00C0C000&
            Caption         =   "Dic"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   2400
            TabIndex        =   3
            Top             =   900
            Width           =   615
         End
         Begin VB.ComboBox ano 
            BackColor       =   &H00C0C000&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   855
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            Height          =   765
            Left            =   30
            Top             =   390
            Width           =   3045
         End
         Begin VB.Label ftext 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   900
            TabIndex        =   15
            Top             =   30
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2385
         Left            =   30
         TabIndex        =   16
         Top             =   1050
         Width           =   3075
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   57
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   930
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   1770
            MultiLine       =   -1  'True
            TabIndex        =   54
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   2190
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   6
            Left            =   2610
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   930
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   1770
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   2190
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   13
            Left            =   2610
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   810
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   930
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   1770
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   2190
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   20
            Left            =   2610
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   1110
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   22
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   23
            Left            =   930
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   25
            Left            =   1770
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   26
            Left            =   2190
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   27
            Left            =   2610
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1410
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   930
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   32
            Left            =   1770
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   2190
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   34
            Left            =   2610
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1710
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   35
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   2010
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   36
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   2010
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   37
            Left            =   930
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   2010
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   38
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   2010
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   39
            Left            =   1770
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   2010
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   40
            Left            =   2190
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   2010
            Width           =   405
         End
         Begin VB.TextBox dias 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   41
            Left            =   2610
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   2010
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Lun"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   0
            Left            =   90
            TabIndex        =   65
            Top             =   270
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mar"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   1
            Left            =   510
            TabIndex        =   64
            Top             =   270
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mié"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   2
            Left            =   930
            TabIndex        =   63
            Top             =   270
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Jue"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   3
            Left            =   1350
            TabIndex        =   62
            Top             =   270
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vie"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   4
            Left            =   1770
            TabIndex        =   61
            Top             =   270
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sáb"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   5
            Left            =   2190
            TabIndex        =   60
            Top             =   270
            Width           =   405
         End
         Begin VB.Label diastext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dom"
            ForeColor       =   &H00000000&
            Height          =   165
            Index           =   6
            Left            =   2610
            TabIndex        =   59
            Top             =   270
            Width           =   405
         End
         Begin VB.Line Line1 
            X1              =   30
            X2              =   3060
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            Height          =   2115
            Left            =   30
            Top             =   240
            Width           =   3045
         End
      End
   End
End
Attribute VB_Name = "fcalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valdia As String
Dim ini As Integer
Dim wmes As Integer
Dim wano As Integer
Dim wdia As Integer
Dim formatofecha As String
Private Sub ano_Change()
    On Local Error Resume Next
    If Len(Trim(ano.Text)) <> 4 Then Exit Sub
    wano = ano.Text
    wfecha = Trim(Str(wdia)) + "/" + Trim(Str(wmes)) + "/" + Trim(Str(wano))
    ini = 0
    Call InicializaDiasMes(wfecha)
End Sub
Private Sub ano_Click()
    Call ano_Change
End Sub
Private Sub dias_Change(Index As Integer)
    If ini = 1 Then
        dias.item(Index).Text = valdia
        dias.item(Index).SelStart = 0
        dias.item(Index).SelLength = Len(dias.item(Index).Text)
    End If
End Sub
Private Sub dias_Click(Index As Integer)
    wdia = dias.item(Index).Text
    dias.item(Index).SetFocus
    dias.item(Index).SelStart = 0
    dias.item(Index).SelLength = Len(dias.item(Index).Text)
    valdia = dias.item(Index).Text
    wfecha = IIf(IsDate(Trim(Str(wdia)) + "/" + Trim(Str(wmes)) + "/" + Trim(Str(wano))), Trim(Str(wdia)) + "/" + Trim(Str(wmes)) + "/" + Trim(Str(wano)), Trim(Str(wdia - 1)) + "/" + Trim(Str(wmes)) + "/" + Trim(Str(wano)))
    'wfecha = Trim(Str(wdia)) + "/" + Trim(Str(wmes)) + "/" + Trim(Str(wano))
    ftext.Caption = Format(wfecha, formatofecha)
    X = InStr(1, wtextc, ".")
    If X > 0 Then
        wforma.Controls(Mid(wtextc, 1, X - 1)).item(Mid(wtextc, X + 1, Len(wtextc))).Text = Format(wfecha, "dd/mm/yyyy")
    Else
        wforma.Controls(wtextc).Text = Format(wfecha, "dd/mm/yyyy")
    End If
    fcalendario.Hide
End Sub
Private Sub dias_KeyPress(Index As Integer, KeyAscii As Integer)
    On Local Error Resume Next
    dias.item(Index).Text = valdia
End Sub
Private Sub Form_Activate()
    Call Form_Load
End Sub
Private Sub Form_Load()
    Toolbar1.Enabled = whabfe
    formatofecha = "ddd dd, mmm yyyy"
    ini = 0
    
    If Len(wfecha) = 0 Or IsNull(wfecha) Or Not IsDate(wfecha) Then wfecha = Date
    wfecha = Format(wfecha, "dd/mm/yyyy")
    ano.Clear
    cont = 0
    For i = Year(wfecha) - 10 To Year(wfecha) + 10
        ano.AddItem Str(i), cont
        ano.ItemData(cont) = i
    Next i
    wdia = Day(wfecha)
    wmes = Month(wfecha)
    wano = Year(wfecha)
    Call IndiceAno(Year(wfecha))
    mes.item(Month(wfecha) - 1).Value = True
End Sub
Private Sub IndiceAno(A As Integer)
    For i = 0 To ano.ListCount - 1
        If ano.ItemData(i) = A Then
            ano.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub
Private Function PrimerDia(d As String) As Integer
    Select Case UCase(Mid(d, 1, 1)) & Mid(d, 2, Len(d))
        Case "Lunes", "Monday":
            PrimerDia = 0
        Case "Martes", "Tuesday":
            PrimerDia = 1
        Case "Miércoles", "Wednesday":
            PrimerDia = 2
        Case "Jueves", "Thursday":
            PrimerDia = 3
        Case "Viernes", "Friday":
            PrimerDia = 4
        Case "Sábado", "Saturday":
            PrimerDia = 5
        Case "Domingo", "Sunday":
            PrimerDia = 6
    End Select
End Function
Private Function UltimoDia(A As Integer, m As Integer) As Integer
    Dim f As Date
    f = Format("27/" + Str(m) + "/" + Str(A), "dd/mm/yyyy")
    For i = 27 To 32
        Err = 0
        t = Format(f, "dddd")
        If Err <> 0 Or Month(f) <> m Then GoTo salida
        f = DateAdd("d", 1, f)
    Next
salida:
    UltimoDia = i - 1
End Function
Private Sub InicializaDiasMes(f As Date)
    For i = 0 To 41
        dias.item(i).Text = ""
        dias.item(i).Enabled = False
    Next i
    Dim d As String
    d = Format("1/" + Trim(Str(Month(f))) + "/" + Trim(Str(Year(f))), "dddd")
    Inicio = PrimerDia(d)
    finaliza = UltimoDia(Year(f), Month(f))
    cont = 1
    For i = Inicio To finaliza + Inicio - 1
        dias.item(i).Text = cont
        dias.item(i).Enabled = True
        If cont = Day(wfecha) Then
            dias.item(i).SetFocus
            dias.item(i).SelStart = 0
            dias.item(i).SelLength = Len(dias.item(i).Text)
            wdia = dias.item(i).Text
            wfecha = dias.item(i).Text + "/" + Trim(Str(wmes)) + "/" + Trim(ano.Text)
            ftext.Caption = Format(wfecha, formatofecha)
        End If
        cont = cont + 1
    Next i
    ini = 1
End Sub
Private Sub mes_Click(Index As Integer)
    On Local Error Resume Next
    Dim d As String
    wmes = Index + 1
    ini = 0
    d = "1/" + Trim(Str(Index + 1)) + "/" + Trim(ano.Text)
    Call InicializaDiasMes(Format(d, "dd/mm/yyyy"))
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        X = InStr(1, wtextc, ".")
        If X > 0 Then
            wforma.Controls(Mid(wtextc, 1, X - 1)).item(Mid(wtextc, X + 1, Len(wtextc))).Text = ""
        Else
            wforma.Controls(wtextc).Text = ""
        End If
        fcalendario.Hide
    End If
End Sub
