VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalenderForm 
   Caption         =   "カレンダー"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3720
   OleObjectBlob   =   "CalenderForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CalenderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize() 'Formが開くとき
  Dim i As Integer
   
  For i = -3 To 3 '前後3年分の年を登録
    Me.ComboBox1.AddItem CStr((Year(clndr_date)) + i)
  Next i
  For i = 1 To 12 '月を登録
    Me.ComboBox2.AddItem CStr(i)
  Next i
    
  Me.ComboBox1 = Year(clndr_date) '年を指定
  Me.ComboBox2 = Month(clndr_date) '月を指定
End Sub
 
Private Sub ComboBox1_Change() '年が変更されたとき
  Call clndr_set
End Sub
 
Private Sub ComboBox2_Change() '月が変更されたとき
  Call clndr_set
End Sub
 
Private Sub clndr_set() 'カレンダーの作成と表示
  Dim yy As Integer, mm As Integer, i As Integer, n As Integer, endDay As Integer
   
  If Me.ComboBox1 = "" Or Me.ComboBox2 = "" Then Exit Sub '年か月どちらか入ってなければ中止
  yy = Me.ComboBox1 '年
  mm = Me.ComboBox2 '月
    
  For i = 1 To 42 'ラベルの初期化
    Me("Label" & i).Caption = ""
    Me("Label" & i).BackColor = Me.BackColor
  Next
   
  n = Weekday(yy & "/" & mm & "/" & 1) - 1 'その月の1日の曜日番号に、マイナス1したもの
  endDay = Day(DateAdd("d", -1, DateAdd("m", 1, yy & "/" & mm & "/" & "1"))) '月末日の算出
  For i = 1 To endDay
    'If (i <= n) Then Me("Label" & i).Caption = CInt(Day(DateAdd("d", "-" & Day(DateSerial(yy, mm, 0) + (n - i)), DateAdd("m", 1, yy & "/" & "/" & "1"))))
    If (i <= n) Then
        Me("Label" & i).Caption = CInt(Day(DateAdd("d", "-" & (n - (i - 1)), yy & "/" & mm & "/" & "1")))
        Me("Label" & i).ForeColor = RGB(169, 169, 169)
    End If
    Me("Label" & i + n).Caption = i '日を入れる
    Select Case Weekday(yy & "/" & mm & "/" & i)
    Case vbSunday
        Me("Label" & i + n).ForeColor = RGB(255, 0, 0)
    Case vbSaturday
        Me("Label" & i + n).ForeColor = RGB(0, 0, 255)
    Case Else
        Me("Label" & i + n).ForeColor = RGB(105, 105, 105)
    End Select
    If CDate(yy & "/" & mm & "/" & i) = clndr_date Then
    Me("Label" & i + n).BackColor = RGB(200, 200, 200) 'TextBoxの日と同じなら色をつける
    Else
    Me("Label" & i + n).BackColor = Me.BackColor
    End If
    
  Next i
  For i = 1 To 42
    If i + n <= (42 - endDay) Then
    Me("Label" & i + endDay + n).Caption = i
    Me("Label" & i + endDay + n).ForeColor = RGB(169, 169, 169)
    End If
    Next i
End Sub


Private Sub SpinButton1_SpinUp() 'ひと月戻る
  If Me.ComboBox2 = 1 Then '1月だったら
    Me.ComboBox1 = Me.ComboBox1 - 1 '年-1
    Me.ComboBox2 = 12 '12月へ
  Else
    Me.ComboBox2 = Me.ComboBox2 - 1
  End If
End Sub
 
Private Sub SpinButton1_SpinDown() 'ひと月進む
  If Me.ComboBox2 = 12 Then '12月だったら
    Me.ComboBox1 = Me.ComboBox1 + 1 '年+1
    Me.ComboBox2 = 1 '1月へ
  Else
    Me.ComboBox2 = Me.ComboBox2 + 1
  End If
End Sub

Private Sub LabelClick(ByVal i As Integer)
    Dim cnt As Integer
    cnt = 1
  If Me("Label" & i).Caption = "" Then Exit Sub 'ラベルが空だったら中止
  If (Me("Label" & i).Caption <= 31 And Me("Label" & i).Caption >= 25) And i <= 6 Then
    While (cnt <= 42)
        If Me("Label" & i).Caption = Me("Label" & cnt + i).Caption Then
            clndr_date = DateAdd("m", -1, Me.ComboBox1 & "/" & Me.ComboBox2 & "/" & Me("Label" & i).Caption)
            clndr_flg = True
            Me.ComboBox1 = Year(clndr_date)
            Me.ComboBox2 = Month(clndr_date)
            Me.Repaint
            Exit Sub
        ElseIf Day(DateSerial(Me.ComboBox1, Me.ComboBox2 + 1, 1) - 1) <= 30 And Day(DateSerial(Me.ComboBox1, Me.ComboBox2 + 1, 1) - 1) >= 28 Then
        clndr_date = Me.ComboBox1 & "/" & Me.ComboBox2 - 1 & "/" & Me("Label" & i).Caption
        clndr_flg = True
        Me.ComboBox1 = Year(clndr_date)
            Me.ComboBox2 = Month(clndr_date)
            Me.Repaint
            Exit Sub
        End If
        cnt = cnt + 1
        Wend
    ElseIf (Me("Label" & i).Caption >= 1 And Me("Label" & i).Caption <= 13) And i >= 32 Then
        While (cnt <= 42)
            If Me("Label" & i).Caption = Me("Label" & 42 - (i - cnt)).Caption Then
            clndr_date = DateAdd("m", 1, Me.ComboBox1 & "/" & Me.ComboBox2 & "/" & Me("Label" & i).Caption)
        clndr_flg = True
        Me.ComboBox1 = Year(clndr_date)
            Me.ComboBox2 = Month(clndr_date)
            Me.Repaint
            Exit Sub
        End If
        cnt = cnt + 1
        Wend
    End If
  clndr_date = Me.ComboBox1 & "/" & Me.ComboBox2 & "/" & Me("Label" & i).Caption '日付を生成して変数に格納
  clndr_flg = True 'フラグを立てる
  Me("Label" & i).BackColor = RGB(200, 200, 200)
  Dim cnt2 As Integer
  cnt2 = 1
  While (cnt2 <= 42)
  If cnt2 = i Then GoTo continue
    Me("Label" & cnt2).BackColor = Me.BackColor
continue:
cnt2 = cnt2 + 1
Wend
End Sub
 
Private Sub Label1_Click(): Call LabelClick(1): End Sub
Private Sub Label2_Click(): Call LabelClick(2): End Sub
Private Sub Label3_Click(): Call LabelClick(3): End Sub
Private Sub Label4_Click(): Call LabelClick(4): End Sub
Private Sub Label5_Click(): Call LabelClick(5): End Sub
Private Sub Label6_Click(): Call LabelClick(6): End Sub
Private Sub Label7_Click(): Call LabelClick(7): End Sub
Private Sub Label8_Click(): Call LabelClick(8): End Sub
Private Sub Label9_Click(): Call LabelClick(9): End Sub
Private Sub Label10_Click(): Call LabelClick(10): End Sub
Private Sub Label11_Click(): Call LabelClick(11): End Sub
Private Sub Label12_Click(): Call LabelClick(12): End Sub
Private Sub Label13_Click(): Call LabelClick(13): End Sub
Private Sub Label14_Click(): Call LabelClick(14): End Sub
Private Sub Label15_Click(): Call LabelClick(15): End Sub
Private Sub Label16_Click(): Call LabelClick(16): End Sub
Private Sub Label17_Click(): Call LabelClick(17): End Sub
Private Sub Label18_Click(): Call LabelClick(18): End Sub
Private Sub Label19_Click(): Call LabelClick(19): End Sub
Private Sub Label20_Click(): Call LabelClick(20): End Sub
Private Sub Label21_Click(): Call LabelClick(21): End Sub
Private Sub Label22_Click(): Call LabelClick(22): End Sub
Private Sub Label23_Click(): Call LabelClick(23): End Sub
Private Sub Label24_Click(): Call LabelClick(24): End Sub
Private Sub Label25_Click(): Call LabelClick(25): End Sub
Private Sub Label26_Click(): Call LabelClick(26): End Sub
Private Sub Label27_Click(): Call LabelClick(27): End Sub
Private Sub Label28_Click(): Call LabelClick(28): End Sub
Private Sub Label29_Click(): Call LabelClick(29): End Sub
Private Sub Label30_Click(): Call LabelClick(30): End Sub
Private Sub Label31_Click(): Call LabelClick(31): End Sub
Private Sub Label32_Click(): Call LabelClick(32): End Sub
Private Sub Label33_Click(): Call LabelClick(33): End Sub
Private Sub Label34_Click(): Call LabelClick(34): End Sub
Private Sub Label35_Click(): Call LabelClick(35): End Sub
Private Sub Label36_Click(): Call LabelClick(36): End Sub
Private Sub Label37_Click(): Call LabelClick(37): End Sub
Private Sub Label38_Click(): Call LabelClick(38): End Sub
Private Sub Label39_Click(): Call LabelClick(39): End Sub
Private Sub Label40_Click(): Call LabelClick(40): End Sub
Private Sub Label41_Click(): Call LabelClick(41): End Sub
Private Sub Label42_Click(): Call LabelClick(42): End Sub

