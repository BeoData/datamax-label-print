VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MILENIUM PRINT"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   600
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   500
      Left            =   2040
      TabIndex        =   3
      Text            =   "Phone: +381(0)11 3660488"
      Top             =   3000
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      Height          =   500
      Left            =   2040
      TabIndex        =   2
      Text            =   "0123456789Qw"
      Top             =   1080
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      Height          =   500
      Left            =   2040
      TabIndex        =   1
      Text            =   "MILENIUM SISTEMI D.O.O."
      Top             =   240
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   3255
      Left            =   7200
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Treci red"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Drugi red /Bar code"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Naslov/ prvi red"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "IDAHC39M Code 39 Barcode"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Function GenerateCode128(Str As String, XPos As Single, YPos As Single, Optional BarWidth As Integer = 1) As Single
    Dim Code128 As New clsCode128
    Dim BarCodeWidth As Long
    
    Me.picBarCode.Cls
    Me.picBarCode.Width = 1
    BarCodeWidth = Code128.Code128_Print(Str, Me.picBarCode, BarWidth, True)
    Me.PaintPicture Me.picBarCode.Image, XPos, YPos, Me.picBarCode.ScaleWidth, Me.picBarCode.ScaleHeight, 0, 0, Me.picBarCode.ScaleWidth, Me.picBarCode.ScaleHeight
    
    Me.CurrentX = XPos + BarCodeWidth / 2 - Me.TextWidth(Str) / 2
    Me.CurrentY = YPos + Me.picBarCode.ScaleHeight + 2
    Me.Print Str
    
    GenerateCode128 = Me.CurrentY
End Function


Private Sub Command1_Click()
Printer.FontName = "Helv"
 ' velicina fonta Printer.FontSize = "12"
 Printer.FontSize = "9"
'1,440 twips equals one inch
 Printer.Height = 1417       ' (6480 )4.5 inches in twips
 Printer.Width = 5670 '(5760 )4 inches in twips
 
 
  Printer.CurrentY = 200 '2 inches (row position)
 Printer.CurrentX = 100 '1 inch (column position)

 Printer.Print Text1.Text
 
  Printer.CurrentY = 200 '2 inches (row position)
  Printer.CurrentX = 3000   '1 inch (column position)

' Printer.Print "IME FIRME D.O.O."
  Printer.Print Text1.Text
 
 
 Printer.CurrentY = 450 '2 inches (row position)
 Printer.CurrentX = 100
 
Printer.FontName = "IDAHC39M Code 39 Barcode"
  ' velicina fonta Printer.FontSize = "9"
 Printer.Print Text2.Text '"IME FIRME 123456"
 
 
  Printer.CurrentY = 450 '2 inches (row position)
 Printer.CurrentX = 3000


Printer.FontName = "IDAHC39M Code 39 Barcode"
 
 'Printer.Print "IME FIRME 123456"
 Printer.Print Text2.Text
 
 Printer.CurrentY = 1200 '2 inches (row position)
 Printer.CurrentX = 100
 
Printer.FontName = "helv"
 ' velicina fonta Printer.FontSize = "10"
 'Printer.Print "Neki Naziv"
 Printer.Print Text3.Text
  Printer.CurrentY = 1200 '2 inches (row position)
 Printer.CurrentX = 3000
 
Printer.FontName = "helv"
 Printer.Print Text3.Text
' Printer.Print "Neki naziv"
 
 Printer.EndDoc
End Function

Private Sub Form_Load()
Dim YPos As Single

YPos = GenerateCode128("0123456789", 1, YPos + 1, 1)

Label1.Caption = Text2.Text
End Sub

Private Sub Text2_Change()
Label1.Caption = Text2.Text


End Sub
