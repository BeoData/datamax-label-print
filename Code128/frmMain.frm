VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Code128 Test"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   716
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Snimi sliku "
      Height          =   855
      Left            =   8520
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   500
      Left            =   2160
      TabIndex        =   4
      Text            =   "MILENIUM SISTEMI D.O.O."
      Top             =   1800
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      Height          =   500
      Left            =   2160
      TabIndex        =   3
      Text            =   "0123456789Qw"
      Top             =   2640
      Width           =   5000
   End
   Begin VB.TextBox Text3 
      Height          =   500
      Left            =   2160
      TabIndex        =   2
      Text            =   "Phone: +381(0)11 3660488"
      Top             =   4560
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   855
      Left            =   8520
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   360
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Naslov/ prvi red"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Drugi red /Bar code"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Treci red"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
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
' Printer.Print Text2.Text '"IME FIRME 123456"
 Printer.Print picBarCode.Image
 
 
 
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
End Sub

 

Private Sub Command2_Click()
 
   SavePicture picBarCode.Image, "slika.bmp"
End Sub

 

Private Sub Form_Load()
    Dim YPos As Single
    Me.FontBold = True
    Label1.Caption = Text2.Text
    ' Using CodeB
   ' YPos = GenerateCode128("Testing Code 128", 10, 10, 1)
    
    ' Using CodeC (Notice that the string is much longer than previous, but actual output is shorter)
   ' YPos = GenerateCode128("123456789012345678901234567890", 10, YPos + 15, 1)
    
    ' Mixing CodeC with CodeB
   ' YPos = GenerateCode128("1234567890123abcdefAAAAAAAAAAA", 10, YPos + 15, 1)
    
    ' Mixing CodeB with CodeA and CodeC (chaged bar width too for testing)
  '  YPos = GenerateCode128("1234aaaaaaabbbb" & Chr(1) & Chr(200) & Chr(255) & "567890998822334", 10, YPos + 15, 2)
    
    ' Bar Width 2
   ' YPos = GenerateCode128("ABCDEFGHIJ", 10, YPos + 15, 2)
    
    ' Bar Width 3
   YPos = GenerateCode128("0123456789", 1, YPos + 1, 1)
    
    ' Bar Width 4
  '  YPos = GenerateCode128("11098432135468798", 10, YPos + 15, 4)
End Sub
