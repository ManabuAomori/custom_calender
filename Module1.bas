Attribute VB_Name = "Module1"
Option Explicit '変数の宣言を強制する

'---カレンダー用変数
Public clndr_date As Date 'テキストボックスの値を格納する変数
Public clndr_flg As Boolean 'カレンダーがクリックされたか判定するフラグ

Sub start()
  MainForm.Show 'メインフォームを開く
End Sub
