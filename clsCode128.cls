VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCode128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



' ***    Made By Michael Ciurescu (CVMichael)   ***



' References:
' http://www.barcodeman.com/info/c128.php3

' Value,  Code Set A,  Code Set B,  Code Set C,  Bar/Space Pattern B S B S B S
'
' 0  SP  SP  00  2 1 2 2 2 2
' 1  !  !  01  2 2 2 1 2 2
' 2  "  "  02  2 2 2 2 2 1
' 3  #  #  03  1 2 1 2 2 3
' 4  $  $  04  1 2 1 3 2 2
' 5  %  %  05  1 3 1 2 2 2
' 6  &  &  06  1 2 2 2 1 3
' 7  '  '  07  1 2 2 3 1 2
' 8  (  (  08  1 3 2 2 1 2
' 9  )  )  09  2 2 1 2 1 3
' 10  *  *  10  2 2 1 3 1 2
' 11  +  +  11  2 3 1 2 1 2
' 12  ,  ,  12  1 1 2 2 3 2
' 13  -  -  13  1 2 2 1 3 2
' 14  .  .  14  1 2 2 2 3 1
' 15  /  /  15  1 1 3 2 2 2
' 16  0  0  16  1 2 3 1 2 2
' 17  1  1  17  1 2 3 2 2 1
' 18  2  2  18  2 2 3 2 1 1
' 19  3  3  19  2 2 1 1 3 2
' 20  4  4  20  2 2 1 2 3 1
' 21  5  5  21  2 1 3 2 1 2
' 22  6  6  22  2 2 3 1 1 2
' 23  7  7  23  3 1 2 1 3 1
' 24  8  8  24  3 1 1 2 2 2
' 25  9  9  25  3 2 1 1 2 2
' 26  :  :  26  3 2 1 2 2 1
' 27  ;  ;  27  3 1 2 2 1 2
' 28  <  <  28  3 2 2 1 1 2
' 29  =  =  29  3 2 2 2 1 1
' 30  >  >  30  2 1 2 1 2 3
' 31  ?  ?  31  2 1 2 3 2 1
' 32  @  @  32  2 3 2 1 2 1
' 33  A  A  33  1 1 1 3 2 3
' 34  B  B  34  1 3 1 1 2 3
' 35  C  C  35  1 3 1 3 2 1
' 36  D  D  36  1 1 2 3 1 3
' 37  E  E  37  1 3 2 1 1 3
' 38  F  F  38  1 3 2 3 1 1
' 39  G  G  39  2 1 1 3 1 3
' 40  H  H  40  2 3 1 1 1 3
' 41  I  I  41  2 3 1 3 1 1
' 42  J  J  42  1 1 2 1 3 3
' 43  K  K  43  1 1 2 3 3 1
' 44  L  L  44  1 3 2 1 3 1
' 45  M  M  45  1 1 3 1 2 3
' 46  N  N  46  1 1 3 3 2 1
' 47  O  O  47  1 3 3 1 2 1
' 48  P  P  48  3 1 3 1 2 1
' 49  Q  Q  49  2 1 1 3 3 1
' 50  R  R  50  2 3 1 1 3 1
' 51  S  S  51  2 1 3 1 1 3
' 52  T  T  52  2 1 3 3 1 1
' 53  U  U  53  2 1 3 1 3 1
' 54  V  V  54  3 1 1 1 2 3
' 55  W  W  55  3 1 1 3 2 1
' 56  X  X  56  3 3 1 1 2 1
' 57  Y  Y  57  3 1 2 1 1 3
' 58  Z  Z  58  3 1 2 3 1 1
' 59  [  [  59  3 3 2 1 1 1
' 60  \  \  60  3 1 4 1 1 1
' 61  ]  ]  61  2 2 1 4 1 1
' 62  ^  ^  62  4 3 1 1 1 1
' 63  _  _  63  1 1 1 2 2 4
' 64  NUL  `  64  1 1 1 4 2 2
' 65  SOH  a  65  1 2 1 1 2 4
' 66  STX  b  66  1 2 1 4 2 1
' 67  ETX  c  67  1 4 1 1 2 2
' 68  EOT  d  68  1 4 1 2 2 1
' 69  ENQ  e  69  1 1 2 2 1 4
' 70  ACK  f  70  1 1 2 4 1 2
' 71  BEL  g  71  1 2 2 1 1 4
' 72  BS  h  72  1 2 2 4 1 1
' 73  HT  i  73  1 4 2 1 1 2
' 74  LF  j  74  1 4 2 2 1 1
' 75  VT  k  75  2 4 1 2 1 1
' 76  FF  I  76  2 2 1 1 1 4
' 77  CR  m  77  4 1 3 1 1 1
' 78  SO  n  78  2 4 1 1 1 2
' 79  SI  o  79  1 3 4 1 1 1
' 80  DLE  p  80  1 1 1 2 4 2
' 81  DC1  q  81  1 2 1 1 4 2
' 82  DC2  r  82  1 2 1 2 4 1
' 83  DC3  s  83  1 1 4 2 1 2
' 84  DC4  t  84  1 2 4 1 1 2
' 85  NAK  u  85  1 2 4 2 1 1
' 86  SYN  v  86  4 1 1 2 1 2
' 87  ETB  w  87  4 2 1 1 1 2
' 88  CAN  x  88  4 2 1 2 1 1
' 89  EM  y  89  2 1 2 1 4 1
' 90  SUB  z  90  2 1 4 1 2 1
' 91  ESC  {  91  4 1 2 1 2 1
' 92  FS  |  92  1 1 1 1 4 3
' 93  GS  }  93  1 1 1 3 4 1
' 94  RS  ~  94  1 3 1 1 4 1
' 95  US  DEL  95  1 1 4 1 1 3
' 96  FNC 3  FNC 3  96  1 1 4 3 1 1
' 97  FNC 2  FNC 2  97  4 1 1 1 1 3
' 98  SHIFT  SHIFT  98  4 1 1 3 1 1
' 99  CODE C  CODE C  99  1 1 3 1 4 1
' 100  CODE B  FNC 4  CODE B  1 1 4 1 3 1
' 101  FNC 4  CODE A  CODE A  3 1 1 1 4 1
' 102  FNC 1  FNC 1  FNC 1  4 1 1 1 3 1
' 103  Start A  Start A  Start A  2 1 1 4 1 2
' 104  Start B  Start B  Start B  2 1 1 2 1 4
' 105  Start C  Start C  Start C  2 1 1 2 3 2
' 106  Stop  Stop  Stop  2 3 3 1 1 1 2

Private Enum eCode128Type
    eCode128_CodeSetA = 1
    eCode128_CodeSetB = 2
    eCode128_CodeSetC = 3
End Enum

Private Type tCode
    ASet As String
    BSet As String
    CSet As String
    BarSpacePattern As String
End Type

Private CodeArr() As tCode

Private Sub Class_Initialize()
    ReDim CodeArr(106)
    
    AddEntry 0, " ", " ", "00", "2 1 2 2 2 2"
    AddEntry 1, "!", "!", "01", "2 2 2 1 2 2"
    AddEntry 2, """", """", "02", "2 2 2 2 2 1"
    AddEntry 3, "#", "#", "03", "1 2 1 2 2 3"
    AddEntry 4, "$", "$", "04", "1 2 1 3 2 2"
    AddEntry 5, "%", "%", "05", "1 3 1 2 2 2"
    AddEntry 6, "&", "&", "06", "1 2 2 2 1 3"
    AddEntry 7, "'", "'", "07", "1 2 2 3 1 2"
    AddEntry 8, "(", "(", "08", "1 3 2 2 1 2"
    AddEntry 9, ")", ")", "09", "2 2 1 2 1 3"
    AddEntry 10, "*", "*", "10", "2 2 1 3 1 2"
    AddEntry 11, "+", "+", "11", "2 3 1 2 1 2"
    AddEntry 12, ",", ",", "12", "1 1 2 2 3 2"
    AddEntry 13, "-", "-", "13", "1 2 2 1 3 2"
    AddEntry 14, ".", ".", "14", "1 2 2 2 3 1"
    AddEntry 15, "/", "/", "15", "1 1 3 2 2 2"
    AddEntry 16, "0", "0", "16", "1 2 3 1 2 2"
    AddEntry 17, "1", "1", "17", "1 2 3 2 2 1"
    AddEntry 18, "2", "2", "18", "2 2 3 2 1 1"
    AddEntry 19, "3", "3", "19", "2 2 1 1 3 2"
    AddEntry 20, "4", "4", "20", "2 2 1 2 3 1"
    AddEntry 21, "5", "5", "21", "2 1 3 2 1 2"
    AddEntry 22, "6", "6", "22", "2 2 3 1 1 2"
    AddEntry 23, "7", "7", "23", "3 1 2 1 3 1"
    AddEntry 24, "8", "8", "24", "3 1 1 2 2 2"
    AddEntry 25, "9", "9", "25", "3 2 1 1 2 2"
    AddEntry 26, ":", ":", "26", "3 2 1 2 2 1"
    AddEntry 27, ";", ";", "27", "3 1 2 2 1 2"
    AddEntry 28, "<", "<", "28", "3 2 2 1 1 2"
    AddEntry 29, "=", "=", "29", "3 2 2 2 1 1"
    AddEntry 30, ">", ">", "30", "2 1 2 1 2 3"
    AddEntry 31, "?", "?", "31", "2 1 2 3 2 1"
    AddEntry 32, "@", "@", "32", "2 3 2 1 2 1"
    AddEntry 33, "A", "A", "33", "1 1 1 3 2 3"
    AddEntry 34, "B", "B", "34", "1 3 1 1 2 3"
    AddEntry 35, "C", "C", "35", "1 3 1 3 2 1"
    AddEntry 36, "D", "D", "36", "1 1 2 3 1 3"
    AddEntry 37, "E", "E", "37", "1 3 2 1 1 3"
    AddEntry 38, "F", "F", "38", "1 3 2 3 1 1"
    AddEntry 39, "G", "G", "39", "2 1 1 3 1 3"
    AddEntry 40, "H", "H", "40", "2 3 1 1 1 3"
    AddEntry 41, "I", "I", "41", "2 3 1 3 1 1"
    AddEntry 42, "J", "J", "42", "1 1 2 1 3 3"
    AddEntry 43, "K", "K", "43", "1 1 2 3 3 1"
    AddEntry 44, "L", "L", "44", "1 3 2 1 3 1"
    AddEntry 45, "M", "M", "45", "1 1 3 1 2 3"
    AddEntry 46, "N", "N", "46", "1 1 3 3 2 1"
    AddEntry 47, "O", "O", "47", "1 3 3 1 2 1"
    AddEntry 48, "P", "P", "48", "3 1 3 1 2 1"
    AddEntry 49, "Q", "Q", "49", "2 1 1 3 3 1"
    AddEntry 50, "R", "R", "50", "2 3 1 1 3 1"
    AddEntry 51, "S", "S", "51", "2 1 3 1 1 3"
    AddEntry 52, "T", "T", "52", "2 1 3 3 1 1"
    AddEntry 53, "U", "U", "53", "2 1 3 1 3 1"
    AddEntry 54, "V", "V", "54", "3 1 1 1 2 3"
    AddEntry 55, "W", "W", "55", "3 1 1 3 2 1"
    AddEntry 56, "X", "X", "56", "3 3 1 1 2 1"
    AddEntry 57, "Y", "Y", "57", "3 1 2 1 1 3"
    AddEntry 58, "Z", "Z", "58", "3 1 2 3 1 1"
    AddEntry 59, "[", "[", "59", "3 3 2 1 1 1"
    AddEntry 60, "\", "\", "60", "3 1 4 1 1 1"
    AddEntry 61, "]", "]", "61", "2 2 1 4 1 1"
    AddEntry 62, "^", "^", "62", "4 3 1 1 1 1"
    AddEntry 63, "_", "_", "63", "1 1 1 2 2 4"
    AddEntry 64, Chr(0), "`", "64", "1 1 1 4 2 2" ' Null
    AddEntry 65, Chr(1), "a", "65", "1 2 1 1 2 4" ' SOH
    AddEntry 66, Chr(2), "b", "66", "1 2 1 4 2 1" ' STX
    AddEntry 67, Chr(3), "c", "67", "1 4 1 1 2 2" ' ETX
    AddEntry 68, Chr(4), "d", "68", "1 4 1 2 2 1" ' EOT
    AddEntry 69, Chr(5), "e", "69", "1 1 2 2 1 4" ' ENQ
    AddEntry 70, Chr(6), "f", "70", "1 1 2 4 1 2" ' ACK
    AddEntry 71, Chr(7), "g", "71", "1 2 2 1 1 4" ' BEL
    AddEntry 72, Chr(8), "h", "72", "1 2 2 4 1 1" ' BS
    AddEntry 73, Chr(9), "i", "73", "1 4 2 1 1 2" ' HT
    AddEntry 74, Chr(10), "j", "74", "1 4 2 2 1 1" ' LF
    AddEntry 75, Chr(11), "k", "75", "2 4 1 2 1 1" ' VT
    AddEntry 76, Chr(12), "I", "76", "2 2 1 1 1 4" ' FF
    AddEntry 77, Chr(13), "m", "77", "4 1 3 1 1 1" ' CR
    AddEntry 78, Chr(14), "n", "78", "2 4 1 1 1 2" ' SO
    AddEntry 79, Chr(15), "o", "79", "1 3 4 1 1 1" ' SI
    AddEntry 80, Chr(16), "p", "80", "1 1 1 2 4 2" ' DLE
    AddEntry 81, Chr(17), "q", "81", "1 2 1 1 4 2" ' DC1
    AddEntry 82, Chr(18), "r", "82", "1 2 1 2 4 1" ' DC2
    AddEntry 83, Chr(19), "s", "83", "1 1 4 2 1 2" ' DC3
    AddEntry 84, Chr(20), "t", "84", "1 2 4 1 1 2" ' DC4
    AddEntry 85, Chr(21), "u", "85", "1 2 4 2 1 1" ' NAK
    AddEntry 86, Chr(22), "v", "86", "4 1 1 2 1 2" ' SYN
    AddEntry 87, Chr(23), "w", "87", "4 2 1 1 1 2" ' ETB
    AddEntry 88, Chr(24), "x", "88", "4 2 1 2 1 1" ' CAN
    AddEntry 89, Chr(25), "y", "89", "2 1 2 1 4 1" ' EM
    AddEntry 90, Chr(26), "z", "90", "2 1 4 1 2 1" ' SUB
    AddEntry 91, Chr(27), "{", "91", "4 1 2 1 2 1" ' ESC
    AddEntry 92, Chr(28), "|", "92", "1 1 1 1 4 3" ' FS
    AddEntry 93, Chr(29), "}", "93", "1 1 1 3 4 1" ' GS
    AddEntry 94, Chr(30), "~", "94", "1 3 1 1 4 1" ' RS
    AddEntry 95, Chr(31), Chr(127), "95", "1 1 4 1 1 3" ' US, DEL
    AddEntry 96, "FNC 3", "FNC 3", "96", "1 1 4 3 1 1"
    AddEntry 97, "FNC 2", "FNC 2", "97", "4 1 1 1 1 3"
    AddEntry 98, "SHIFT", "SHIFT", "98", "4 1 1 3 1 1"
    AddEntry 99, "CODE C", "CODE C", "99", "1 1 3 1 4 1"
    AddEntry 100, "CODE B", "FNC 4", "CODE B", "1 1 4 1 3 1"
    AddEntry 101, "FNC 4", "CODE A", "CODE A", "3 1 1 1 4 1"
    AddEntry 102, "FNC 1", "FNC 1", "FNC 1", "4 1 1 1 3 1"
    AddEntry 103, "Start A", "Start A", "Start A", "2 1 1 4 1 2"
    AddEntry 104, "Start B", "Start B", "Start B", "2 1 1 2 1 4"
    AddEntry 105, "Start C", "Start C", "Start C", "2 1 1 2 3 2"
    AddEntry 106, "Stop", "Stop", "Stop", "2 3 3 1 1 1 2"
End Sub

Private Sub AddEntry(ByVal Index As Integer, ASet As String, BSet As String, CSet As String, BarSpacePattern As String)
    With CodeArr(Index)
        .ASet = ASet
        .BSet = BSet
        .CSet = CSet
        .BarSpacePattern = Replace(BarSpacePattern, " ", "")
    End With
End Sub

Public Function Code128_Str(ByVal Str As String)
    Code128_Str = Replace(BuildStr(Str), " ", "")
End Function

Private Function BuildStr(ByVal Str As String) As String
    Dim SCode As eCode128Type, PrevSCode As eCode128Type
    Dim CurrChar As String, ArrIndex As Integer, CharIndex As Long
    Dim CheckDigit As Integer, CCodeIndex As Integer, TotalSum As Long
    
    SCode = eCode128_CodeSetB
    If Str Like "##*" Then SCode = eCode128_CodeSetC
    
    TotalSum = 0
    CharIndex = 1
    
    Select Case SCode
    Case eCode128_CodeSetA
        TotalSum = TotalSum + (103 * CharIndex)
        BuildStr = Trim(BuildStr) & " 2 1 1 4 1 2"
    Case eCode128_CodeSetB
        TotalSum = TotalSum + (104 * CharIndex)
        BuildStr = Trim(BuildStr) & " 2 1 1 2 1 4"
    Case eCode128_CodeSetC
        TotalSum = TotalSum + (105 * CharIndex)
        BuildStr = Trim(BuildStr) & " 2 1 1 2 3 2"
    End Select
    
    PrevSCode = SCode
    
    Do Until Len(Str) = 0
        If Str Like "####*" Then SCode = eCode128_CodeSetC
        
        If SCode = eCode128_CodeSetC And Mid(Str, 1, 2) Like "##" Then
            CurrChar = Mid(Str, 1, 2)
        Else
            CurrChar = Mid(Str, 1, 1)
        End If
        
        ArrIndex = GetCharIndex(CurrChar, SCode, True)
        
        If ArrIndex <> -1 Then
            If CodeArr(ArrIndex).BSet = CurrChar And ((SCode = eCode128_CodeSetC And CodeArr(ArrIndex).CSet <> CurrChar) Or (SCode = eCode128_CodeSetC And CodeArr(ArrIndex).ASet <> CurrChar)) Then
                SCode = eCode128_CodeSetB
            ElseIf CodeArr(ArrIndex).ASet = CurrChar And CodeArr(ArrIndex).BSet <> CurrChar Then
                SCode = eCode128_CodeSetA
            ElseIf CodeArr(ArrIndex).CSet = CurrChar Then
                SCode = eCode128_CodeSetC
            End If
            
            If PrevSCode <> SCode Then
                Select Case SCode
                Case eCode128_CodeSetA
                    CCodeIndex = GetCharIndex("CODE A", PrevSCode, False)
                Case eCode128_CodeSetB
                    CCodeIndex = GetCharIndex("CODE B", PrevSCode, False)
                Case eCode128_CodeSetC
                    CCodeIndex = GetCharIndex("CODE C", PrevSCode, False)
                End Select
                
                TotalSum = TotalSum + (CCodeIndex * CharIndex)
                BuildStr = Trim(BuildStr) & " " & CodeArr(CCodeIndex).BarSpacePattern
                
                CharIndex = CharIndex + 1
                PrevSCode = SCode
            End If
            
            BuildStr = Trim(BuildStr) & " " & CodeArr(ArrIndex).BarSpacePattern
            
            TotalSum = TotalSum + (ArrIndex * CharIndex)
            CharIndex = CharIndex + 1
        End If
        
        If SCode = eCode128_CodeSetC Then
            Str = Mid(Str, 3)
        Else
            Str = Mid(Str, 2)
        End If
    Loop
    
    CheckDigit = TotalSum Mod 103
    
    BuildStr = Trim(BuildStr) & " " & CodeArr(CheckDigit).BarSpacePattern
    BuildStr = Trim(BuildStr) & " 2 3 3 1 1 1 2"
End Function

Private Function GetCharIndex(ByVal Char As String, ByVal CodeType As eCode128Type, ByVal Recurse As Boolean) As Integer
    Dim K As Long
    
    Select Case CodeType
    Case eCode128_CodeSetA
        For K = 0 To UBound(CodeArr)
            If Char = CodeArr(K).ASet Then Exit For
        Next K
    Case eCode128_CodeSetB
        For K = 0 To UBound(CodeArr)
            If Char = CodeArr(K).BSet Then Exit For
        Next K
    Case eCode128_CodeSetC
        For K = 0 To UBound(CodeArr)
            If Char = CodeArr(K).CSet Then Exit For
        Next K
    End Select
    
    If K = UBound(CodeArr) + 1 Then
        If Not Recurse Then
            GetCharIndex = -1
        Else
            Select Case CodeType
            Case eCode128_CodeSetA
                GetCharIndex = GetCharIndex(Char, eCode128_CodeSetC, False)
            Case eCode128_CodeSetB
                GetCharIndex = GetCharIndex(Char, eCode128_CodeSetA, False)
            Case eCode128_CodeSetC
                GetCharIndex = GetCharIndex(Char, eCode128_CodeSetB, False)
            End Select
            
            If GetCharIndex = -1 Then
                Select Case CodeType
                Case eCode128_CodeSetA
                    GetCharIndex = GetCharIndex(Char, eCode128_CodeSetB, False)
                Case eCode128_CodeSetB
                    GetCharIndex = GetCharIndex(Char, eCode128_CodeSetC, False)
                Case eCode128_CodeSetC
                    GetCharIndex = GetCharIndex(Char, eCode128_CodeSetA, False)
                End Select
            End If
        End If
    Else
        GetCharIndex = K
    End If
End Function

Public Function Code128_GetWidth(ByVal Str As String, Optional ByVal BarWidth As Integer = 1) As Long
    Dim K As Long, Width As Long
    
    Str = Replace(Code128_Str(Str), " ", "")
    Debug.Print Str
    For K = 1 To Len(Str)
        Width = Width + Val(Mid(Str, K, 1))
    Next K
    
    Code128_GetWidth = Width * BarWidth + (28 * BarWidth)
End Function

Public Function Code128_Print(ByVal Str As String, Pic As PictureBox, Optional ByVal BarWidth As Integer = 1, Optional AutoEnlarge As Boolean = False) As Long
    Dim Ret As String, X As Long, CurrChar As String, K As Long, Q As Long
    Dim LineWidth As Integer
    
    If BarWidth < 1 Then BarWidth = 1
    Pic.ScaleMode = vbPixels
    
    Ret = Replace(Code128_Str(Str), " ", "")
    Debug.Print Ret
    X = 14 * BarWidth
    Do
        CurrChar = Left(Ret, 2)
        Ret = Mid(Ret, 3)
        
        LineWidth = Val(Left(CurrChar, 1))
        For K = 1 To LineWidth
            For Q = 0 To BarWidth - 1
                Pic.Line (X + Q, 0)-(X + Q, Pic.ScaleHeight), Pic.ForeColor
            Next Q
            
            X = X + BarWidth
        Next K
        
        If AutoEnlarge Then EnlargePic Pic, X + (10 * BarWidth)
        X = X + (Val(Mid(CurrChar & "0", 2, 1)) * BarWidth)
    Loop Until Len(Ret) = 0
    
    X = X + (14 * BarWidth)
    If AutoEnlarge Then EnlargePic Pic, X
    
    Code128_Print = X
End Function

Private Sub EnlargePic(Pic As PictureBox, ByVal Width As Long)
    If Pic.ScaleWidth < Width Then
        If Pic.Parent.ScaleMode = vbPixels Then
            Pic.Width = Width
        ElseIf Pic.Parent.ScaleMode = vbTwips Then
            Pic.Width = Width * Screen.TwipsPerPixelX
        End If
    End If
End Sub

