VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "*\A..\..\..\..\SOURCE~1\AURORA~1\OCX\CAINBU~1\cainObjects.vbp"
Begin VB.Form frmTest 
   BackColor       =   &H00FFF2ED&
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   844
   StartUpPosition =   2  'Bildschirmmitte
   Begin cainObjects.cainHarmonizer cainHarmonizer1 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
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
   Begin cainObjects.cainLabel cainLabel2 
      Height          =   1335
      Index           =   0
      Left            =   9720
      TabIndex        =   28
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2355
      ImageIndex      =   1
      Caption         =   "cainLabel2"
      Hyperlink       =   "www.planet-source-code.com"
      TextWrap        =   -1  'True
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
   Begin cainObjects.cainMonthview cainMonthview1 
      Height          =   2310
      Left            =   6240
      TabIndex        =   25
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4075
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
   Begin cainObjects.cainDropDown cainDropDown1 
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   20
      Top             =   4800
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
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
   Begin cainObjects.cainScrollBar cainScrollBar1 
      Height          =   255
      Left            =   6240
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      Orientation     =   1
   End
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   15
      Top             =   1680
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   450
      Caption         =   "Text boxes"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainTextBox cainTextBox1 
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "All chars"
   End
   Begin cainObjects.cainProgressBar cainProgressBar2 
      Height          =   255
      Left            =   2520
      Top             =   7200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      Caption         =   ""
   End
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   13
      Top             =   6360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   450
      Caption         =   "Progressbar"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   12
      Top             =   5040
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   450
      Caption         =   "Slidebar"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainProgressBar cainProgressBar1 
      Height          =   375
      Left            =   2520
      Top             =   6720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
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
   Begin cainObjects.cainButton cainButton2 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "tickstyles"
   End
   Begin cainObjects.cainSlider cainSlider1 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   5400
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   16311010
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   100
   End
   Begin cainObjects.cainButton cainButton1 
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Red"
      TabSelectColor  =   49344
      ImageIndex      =   4
   End
   Begin cainObjects.cainButton cainButton1 
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Black"
      ImageIndex      =   3
   End
   Begin cainObjects.cainButton cainButton1 
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Green"
      ImageIndex      =   2
   End
   Begin cainObjects.cainButton cainButton1 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Blue"
      ImageIndex      =   1
   End
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   3840
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      Caption         =   "Menu"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainPUMenu cainPUMenu1 
      Height          =   705
      Left            =   2520
      TabIndex        =   4
      Top             =   4200
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
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
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   450
      Caption         =   "Listbox"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainListbox cainListbox1 
      Height          =   1575
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
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
   Begin cainObjects.cainToolbar cainToolbar2 
      Align           =   1  'Oben ausrichten
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   714
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
   Begin cainObjects.cainToolbar cainToolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   714
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":23E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":40EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4686
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   16
      Left            =   1320
      Top             =   5520
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7002
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":78DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":81B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8A90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin cainObjects.cainTextBox cainTextBox1 
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   16
      Top             =   2400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "text only"
      DataFormat      =   1
   End
   Begin cainObjects.cainTextBox cainTextBox1 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   17
      Top             =   2760
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "12345"
      DataFormat      =   2
   End
   Begin cainObjects.cainTextBox cainTextBox1 
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   18
      Top             =   3120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "UCASED TEXT ONLY"
      DataFormat      =   7
      SelectionColor  =   49152
   End
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   19
      Top             =   3600
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   450
      Caption         =   "Scrollbar"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   21
      Top             =   4440
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   450
      Caption         =   "Dropdown"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainDropDown cainDropDown1 
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   22
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ComboType       =   3
      PrinterName     =   "Microsoft Office Document Image Writer"
   End
   Begin cainObjects.cainDropDown cainDropDown1 
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   23
      Top             =   5520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ComboType       =   4
   End
   Begin cainObjects.cainDropDown cainDropDown1 
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   24
      Top             =   5880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ComboType       =   5
   End
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   26
      Top             =   6360
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   450
      Caption         =   "Monthview"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainLabel cainLabel1 
      Height          =   255
      Index           =   8
      Left            =   9720
      TabIndex        =   27
      Top             =   1680
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   450
      ImageIndex      =   1
      Caption         =   "Label"
      AutoSize        =   -1  'True
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
   Begin cainObjects.cainLabel cainLabel2 
      Height          =   1335
      Index           =   2
      Left            =   9720
      TabIndex        =   29
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2355
      Caption         =   "cainLabel3"
      Hyperlink       =   "www.planet-source-code.com"
      TextWrap        =   -1  'True
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
   Begin cainObjects.cainLabel cainLabel2 
      Height          =   1335
      Index           =   1
      Left            =   9720
      TabIndex        =   30
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2355
      ImageIndex      =   4
      Caption         =   "cainLabel4"
      TextWrap        =   -1  'True
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
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim ts As Integer

Private Const TEST_TEXT As String = "You 've got to press it on you. You chose to pick it. That 's what you do, baby. Hold it down, dare." & vbCrLf & vbCrLf & "Jump with the moon and move it. Jump back And forth. It feels like you where there yourself. Work it out."

Private Sub cainButton1_Click(Index As Integer)

    Dim i As Integer
    Dim lngCol As Long
    Dim lngTabCol As Long
    Dim lngSelCol As Long
    
    Select Case Index
    Case 0:        lngCol = &HC00000:   lngTabCol = &H96E7&:        lngSelCol = &H80FF&
    Case 1:        lngCol = &H8000&:    lngTabCol = &H96E7&:        lngSelCol = &H80FF&
    Case 2:        lngCol = &H0&:       lngTabCol = &H80FF&:        lngSelCol = &HC000&
    Case 3:        lngCol = &HC0&:      lngTabCol = &H80FF&:        lngSelCol = &HC0C0&
    End Select
        
    cainHarmonizer1.ForeColor = lngCol
    cainHarmonizer1.SelectionColor = lngSelCol
    cainHarmonizer1.TabSelectColor = lngTabCol
    cainHarmonizer1.Font.Name = cainDropDown1(3).FontName
    cainHarmonizer1.Harmonize
    
End Sub

Private Sub cainButton2_Click()
    
    ts = ts + 1
    If ts = 5 Then ts = 0
    
    cainSlider1.TickStyle = ts
    
End Sub

Private Sub cainLabel2_HyperlinkClick(Index As Integer, sHyperlink As String, Cancel As Integer)

    If Index = 2 Then
        Cancel = 1
        MsgBox "My Hyperlink: " & sHyperlink
    End If

End Sub

Private Sub cainPUMenu1_ItemClick(ItemIndex As Integer, ItemKey As String)
    'If ItemIndex = 3 Then
        If cainPUMenu1.menuitem(ItemIndex).Checked = True Then
            cainPUMenu1.menuitem(ItemIndex).Checked = False
        Else
            cainPUMenu1.menuitem(ItemIndex).Checked = True
        End If
    'End If

End Sub


Private Sub cainSlider1_valuechange()
    cainProgressBar1.Value = cainSlider1.Value
End Sub

Private Sub cainToolbar2_ItemClicked(ItemIndex As Integer, ItemKey As String)
    If ItemIndex = 2 Then
        If cainToolbar2.ToolBarItems.Item("a").Checked = False Then
            cainToolbar2.ToolBarItems.Item("a").Checked = True
        Else
            cainToolbar2.ToolBarItems.Item("a").Checked = False
        End If
    End If

End Sub

Private Sub Form_Initialize()
    cainButton1_Click 0
End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    For i = 0 To cainButton1.Count - 1
        
        Set cainButton1(i).ImageList = ImageList1
        
    Next i
    
    For i = 0 To cainLabel2.Count - 1
        Set cainLabel2(i).ImageList = ImageList1
    Next i
    
    For i = 0 To cainLabel1.Count - 1
        Set cainLabel1(i).ImageList = ImageList2
    Next i
    
    cainLabel2(0).Caption = TEST_TEXT
    cainLabel2(1).Caption = TEST_TEXT
    cainLabel2(2).Caption = TEST_TEXT
    
    Set cainPUMenu1.ImageList = ImageList2
    
    cainPUMenu1.menuitem.Add "a", "Cain Menu Demonstration", 1, , , , itTitle, , True
    cainPUMenu1.menuitem.Add "b", "Money Maker", 2, , , , itCheck, , True
    cainPUMenu1.menuitem.Add "c", "Green Barets", 3, , , , itCheck, True
    cainPUMenu1.menuitem.Add "d", "Martini", 4, , , , itCheck
    cainPUMenu1.menuitem.Add "e", , , , , , itPlaceholder
    cainPUMenu1.menuitem.Add "f", "Create Baby"
    cainPUMenu1.menuitem.Add "g", "Back to the Future", , , , , itTitle
    cainPUMenu1.menuitem.Add "h", "Futurama", 3
    
    Set cainToolbar1.ImageList = ImageList1
    Set cainToolbar1.MenuImageList = ImageList2
    Set cainToolbar2.ImageList = ImageList2
    
    cainToolbar1.ToolBarItems.Add "a", "Datei", 1, mitMenu
    cainToolbar1.ToolBarItems.Add "we", "Bearbeiten", 2, mitMenu2
    cainToolbar1.ToolBarItems.Add "d", "", , mitSeparator
    cainToolbar1.ToolBarItems.Add "r", "Ansicht", , mitNormalButton, , False
    
    cainToolbar2.ToolBarItems.Add "e3", "Debuggen", 4, mitNormalButton, , True
    cainToolbar2.ToolBarItems.Add "a", "Projekt", , mitCheckButton, True
    cainToolbar2.ToolBarItems.Add "e4", , , mitPlaceholder, , True
    cainToolbar2.ToolBarItems.Add "e1", , 3, mitPlaceholder
    cainToolbar2.ToolBarItems.Add "e22", "Cain Elements Demo", , mitCaption
    
    cainToolbar1.ToolBarItems.Item(1).ToolbarMenuItems.Add "d4", "Neues Projekt", 2
    cainToolbar1.ToolBarItems.Item(1).ToolbarMenuItems.Add "d34", "Projekt öffnen", 4
    cainToolbar1.ToolBarItems.Item(1).ToolbarMenuItems.Add "d43", "Projekt speichern", 1
    cainToolbar1.ToolBarItems.Item(1).ToolbarMenuItems.Add "dd", , , , , , itPlaceholder
    cainToolbar1.ToolBarItems.Item(1).ToolbarMenuItems.Add "drt", "Schließen"
    
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "dew", "Ausschneiden"
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "dwa", "Kopieren"
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "dsa", "Löschen"
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "dww", , , , , , itPlaceholder
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "ki", "Alles Auswählen", 3
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "kiu", , , , , , itPlaceholder
    cainToolbar1.ToolBarItems.Item(2).ToolbarMenuItems.Add "kh", "Log", 2, , , , itCheck

    Set cainListbox1.ImageList = ImageList2
    cainListbox1.ListboxItems.Sorted = True
    cainListbox1.ListboxItems.Sorting = sd_Ascending
    cainListbox1.ListboxItems.Add "a", "Hola", "Hola", 3
    cainListbox1.ListboxItems.Add "b", "Chica", "Hola chica", 4
    cainListbox1.ListboxItems.Add "c", "Bonita", "Hola chica bonita", 1
    cainListbox1.ListboxItems.Add "d", "Amigas", "Hola amigas"
    cainListbox1.ListboxItems.Add "e", "Tambien", "Espaniol", 2
    cainListbox1.ListboxItems.Add "f", "Sotros", "Rachel", 1
    cainListbox1.ListboxItems.Add "g", "Gracias", "Green", 1
    cainListbox1.ListboxItems.Add "h", "Muchas gracias", "Friends", 4
    cainListbox1.ListboxItems.Add "i", "Donde eres", "Mulan", 3
    cainListbox1.ListboxItems.Add "j", "Padre", "Madre"
    cainListbox1.ListboxItems.Add "k", "Der Brockhaus von A-Z", "Joey", 3
    cainListbox1.ListboxItems.Add "l", "Wir können alles schaffen... Müssen nur wollen...", "Mi", 1

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then cainPUMenu1.CreateMenu 0, 0
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    cainLabel2(0).Width = Me.ScaleWidth - cainLabel2(0).Left
    cainLabel2(1).Width = Me.ScaleWidth - cainLabel2(1).Left
    cainLabel2(2).Width = Me.ScaleWidth - cainLabel2(2).Left

End Sub

Private Sub Timer1_Timer()
    
    X = X + 2
    If X > 100 Then X = 0
    
    cainProgressBar2.Value = X
    cainToolbar2.ToolBarItems.Item("e22").Caption = X & "% Cain Elements Demo"
    cainToolbar2.ToolBarItems.Refresh
    

End Sub
