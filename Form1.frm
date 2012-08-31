VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Девятка"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "закончить"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "начать"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   11
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   10
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   9
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   8
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   7
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   6
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   5
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   11
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   10
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   9
      Left            =   10080
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   8
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   7
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   6
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   5
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   11
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   10
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   9
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   8
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   7
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   6
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   5
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   4
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   4
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   10080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape kar3 
      Height          =   3255
      Left            =   8880
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   4
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape kar2 
      Height          =   3255
      Left            =   5880
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape kar1 
      Height          =   4815
      Left            =   7440
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Image stol 
      BorderStyle     =   1  'Fixed Single
      Height          =   6375
      Left            =   120
      Top             =   120
      Width           =   5535
   End
   Begin VB.Menu m_file 
      Caption         =   "&Файл"
      Begin VB.Menu start 
         Caption         =   "Начать"
      End
      Begin VB.Menu end 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu m_nastr 
      Caption         =   "&Настройки"
   End
   Begin VB.Menu m_spravka 
      Caption         =   "&Помощь"
      Begin VB.Menu m_help 
         Caption         =   "Справка"
      End
      Begin VB.Menu m_info 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim filerub As String
Dim karti As String
Dim ind As Integer
Dim pole(1 To 9, 1 To 4) As Integer
Dim vibor As String
Dim nhod As Integer
Dim chprikup As Integer
Dim vibr1 As Integer
Dim vibr2 As Integer
Dim vibr3 As Integer
Dim nk As Integer
Dim nvibkarti As Integer
Dim nkx As Integer
Dim kolvo As Integer
Dim n As Integer, n1 As Integer
Dim chhod As Integer
Dim tip As Integer

Private Sub Command1_Click()
Call start_Click
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub end_Click()
If MsgBox("Вы действительно хотите выйти?", vbYesNo, "Выход") = vbYes Then End
End Sub

Private Sub Form_Load()
Call razdrub
For ind = 0 To 11
karta1(ind).MouseIcon = LoadPicture(App.Path + "\h_point.cur")
Next ind
'ReDim pole(1 To 9, 1 To 4) As Integer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Вы действительно хотите выйти?", vbYesNo, "Выход") = vbNo Then
Cancel = 1
End If
End Sub

Private Sub karta1_Click(Index As Integer)
If vibr1 > -1 And vibr1 < 12 Then GoTo en
If vibr1 <> 12 And Index <> 12 Then
vibr1 = Index
Else
GoTo sh
End If


For n1 = 1 To Len(vibor) Step 2
If Val(Mid(vibor, n1, 2)) = Index Then
vibor = "mozhno"
Exit For
End If
Next n1

If vibor <> "mozhno" Then GoTo en
'karta1(Index).Move stol.Left + (karta1(Index).Width - 20) - 700, stol.Top + (karta1(Index).Height - 20) - 700
karta1(Index).Move stol.Left + stol.Width / 9 * (Val(Mid(karta1(Index).Tag, 1, 1)) - 1), stol.Top + stol.Height / 4 * (Val(Mid(karta1(Index).Tag, 2, 1)) - 1)
karta1(Index).Enabled = False
If Val(Mid(karta1(Index).Tag, 1, 1)) < 4 Then
karta1(Index).ZOrder (1)
Else
karta1(Index).ZOrder (0)
End If
stol.Tag = stol.Tag + "1" + Mid(karta1(Index).Tag, 1, 2)
pole(Val(Mid(karta1(Index).Tag, 1, 1)), Val(Mid(karta1(Index).Tag, 2, 1))) = Val(karta1(Index).Tag)
kar1.Tag = Val(kar1.Tag) + 1

sh:
chhod = chhod + 1
If chhod = 4 Then chhod = 1
If chhod = 2 Then Call games
If chhod = 3 Then Call games
en:
    vibr1 = -1
End Sub

Private Sub karta1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
karta1(Index).MousePointer = 99
End Sub

Private Sub start_Click()
chhod = 1
start.Enabled = False
vibr1 = -1
vibr2 = -1
vibr3 = -1
kar1.Tag = 0
kar2.Tag = 0
kar3.Tag = 0
'stol.ZOrder (0)
kar1.Tag = ""
start.Enabled = False
Call razd 'раздача карт
'Call analiz
Call games
End Sub

Public Sub razd()
nhod = 0
nk = 36
'karti = "6б6в6к6ч7б7в7к7ч8б8в8к8ч9б9в9к9чабавакачвбвввквчдбдвдкдчкбквкккчтбтвтктч"
karti = "111213142122232431323334414243445152535461626364717273748182838491929394"
For nkx = 0 To 11
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
karta1(nkx).Tag = Mid(karti, nvibkarti, 2)
karta1(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
karta2(nkx).Tag = Mid(karti, nvibkarti, 2)
 karta2(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
karta3(nkx).Tag = Mid(karti, nvibkarti, 2)
 karta3(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
Next nkx

End Sub

Public Sub razdrub()
filerub = "b01_c.bmp" 'временная строка, имя файла текущей рубашки загружать из спец.файла из директории с рубашками
For nkx = 0 To 11
karta1(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
karta2(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
karta3(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
Next nkx
End Sub

Public Sub hod2()
If vibr2 <> -1 Then GoTo en
nach:
Call prvibr2
If vibr2 = 12 Then GoTo en
If karta2(vibr2).Enabled = False Then GoTo nach
'karta2(vibr2).Move 20 + stol.Left + (karta2(vibr2).Width - 20), 20 + stol.Top + (karta2(vibr2).Height - 20)
karta2(vibr2).Move stol.Left + stol.Width / 9 * (Val(Mid(karta2(vibr2).Tag, 1, 1)) - 1), stol.Top + stol.Height / 4 * (Val(Mid(karta2(vibr2).Tag, 2, 1)) - 1)
karta2(vibr2).Enabled = False
If Val(Mid(karta2(vibr2).Tag, 1, 1)) < 4 Then
karta2(vibr2).ZOrder (1)
Else
karta2(vibr2).ZOrder (0)
End If
stol.Tag = stol.Tag + "2" + Mid(karta2(vibr2).Tag, 1, 2)
pole(Val(Mid(karta2(vibr2).Tag, 1, 1)), Val(Mid(karta2(vibr2).Tag, 2, 1))) = Val(karta2(vibr2).Tag)
kar2.Tag = Val(kar2.Tag) + 1

en:
chhod = chhod + 1
If chhod = 4 Then chhod = 1
Call games
End Sub

Public Sub hod3()
If vibr3 <> -1 Then GoTo en
nach:
Call prvibr3
If vibr3 = 12 Then GoTo en
If karta3(vibr3).Enabled = False Then GoTo nach
'karta3(vibr3).Move 400 + stol.Left + (karta3(vibr3).Width - 20), 400 + stol.Top + (karta3(vibr3).Height - 20)
karta3(vibr3).Move stol.Left + stol.Width / 9 * (Val(Mid(karta3(vibr3).Tag, 1, 1)) - 1), stol.Top + stol.Height / 4 * (Val(Mid(karta3(vibr3).Tag, 2, 1)) - 1)
karta3(vibr3).Enabled = False
If Val(Mid(karta3(vibr3).Tag, 1, 1)) < 4 Then
karta3(vibr3).ZOrder (1)
Else
karta3(vibr3).ZOrder (0)
End If
stol.Tag = stol.Tag + "3" + Mid(karta3(vibr3).Tag, 1, 2)
pole(Val(Mid(karta3(vibr3).Tag, 1, 1)), Val(Mid(karta3(vibr3).Tag, 2, 1))) = Val(karta3(vibr3).Tag)
kar3.Tag = Val(kar3.Tag) + 1

en:
chhod = chhod + 1
If chhod = 4 Then chhod = 1
    Call games
End Sub

Public Sub games()
nach:
vibr1 = -1
vibr2 = -1
vibr3 = -1

If Val(kar1.Tag) = 12 Then
'все карты выложены выиграл 1 игрок

Exit Sub
End If

If Val(kar2.Tag) = 12 Then
'все карты выложены выиграл 2 игрок

Exit Sub
End If

If Val(kar3.Tag) = 12 Then
'все карты выложены выиграл 3 игрок

Exit Sub
End If

Select Case chhod
Case 1
'Call hod1

vibor = ""
For n = 0 To 11
If karta1(n).Enabled = False Then GoTo nxt
Select Case Val(Mid(karta1(n).Tag, 1, 1))
Case 1 To 3
If pole(Val(Mid(karta1(n).Tag, 1, 1)) + 1, Val(Mid(karta1(n).Tag, 2, 1))) <> 0 Then vibor = vibor + Right("0" + Str(n), 2)
Case 4
'vibor = vibor + karta2(n).Tag
vibor = vibor + Right("0" + Str(n), 2)
Case 5 To 9
If pole(Val(Mid(karta1(n).Tag, 1, 1)) - 1, Val(Mid(karta1(n).Tag, 2, 1))) <> 0 Then vibor = vibor + Right("0" + Str(n), 2)
End Select
nxt:
Next n
If vibor = "" Then
'нет карт чтобы сходить
'chhod = chhod + 1
'If chhod = 4 Then chhod = 1
'If chhod = 2 Then Call games
'If chhod = 3 Then Call games
'GoTo nach
    vibr1 = 12
    Call karta1_Click(12)
    Exit Sub
End If

Case 2
Call hod2
Case 3
Call hod3
End Select

End Sub

Public Sub prvibr2()
vibor = ""
For n = 0 To 11
If karta2(n).Enabled = False Then GoTo nxt
Select Case Val(Mid(karta2(n).Tag, 1, 1))
Case 1 To 3
If pole(Val(Mid(karta2(n).Tag, 1, 1)) + 1, Val(Mid(karta2(n).Tag, 2, 1))) <> 0 Then vibor = vibor + Right("0" + Str(n), 2)
Case 4
'vibor = vibor + karta2(n).Tag
vibor = vibor + Right("0" + Str(n), 2)
Case 5 To 9
If pole(Val(Mid(karta2(n).Tag, 1, 1)) - 1, Val(Mid(karta2(n).Tag, 2, 1))) <> 0 Then vibor = vibor + Right("0" + Str(n), 2)
End Select
nxt:
Next n
If vibor = "" Then
'нет карт чтобы сходить
vibr2 = 12
'call games
Exit Sub
End If

'For n1 = 1 To Len(vibor)
'выбор лучшей карты из списка доступных номеров
vibr2 = Val(Mid(vibor, 1, 2))
'Next n1

'vibr2 = Int(Rnd * 12)
End Sub

Public Sub prvibr3()
vibor = ""
For n = 0 To 11
If karta3(n).Enabled = False Then GoTo nxt
Select Case Val(Mid(karta3(n).Tag, 1, 1))
Case 1 To 3
If pole(Val(Mid(karta3(n).Tag, 1, 1)) + 1, Val(Mid(karta3(n).Tag, 2, 1))) <> 0 Then vibor = vibor + Right("0" + Str(n), 2)
Case 4
'vibor = vibor + karta3(n).Tag
vibor = vibor + Right("0" + Str(n), 2)
Case 5 To 9
If pole(Val(Mid(karta3(n).Tag, 1, 1)) - 1, Val(Mid(karta3(n).Tag, 2, 1))) <> 0 Then vibor = vibor + Right("0" + Str(n), 2)
End Select
nxt:
Next n
If vibor = "" Then
'нет карт чтобы сходить
vibr3 = 12
'call games
Exit Sub
End If

'For n1 = 1 To Len(vibor)
'выбор лучшей карты из списка доступных номеров
vibr3 = Val(Mid(vibor, 1, 2))
'Next n1


'vibr3 = Int(Rnd * 12)
End Sub
