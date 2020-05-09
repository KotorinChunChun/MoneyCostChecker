Attribute VB_Name = "kccFuncCore_Partial"
Rem --------------------------------------------------
Rem
Rem @module kccFuncCore_Partial
Rem
Rem @description
Rem    必須関数だけを集めたモジュール
Rem　　　から抽出したもの
Rem
Rem --------------------------------------------------
Option Explicit

Rem 数字か否か
Rem  @param base_char   一文字
Rem  @param enable_dot  ドットを数字の一部と見なすか
Rem  @return As Boolean True:=数字,False:=数字以外
Rem  @note IsNumericは指数表記とかもTrueを返すため使えないため必要
Rem
Private Function IsNumChar(base_char, Optional enable_dot As Boolean = True) As Boolean
    If Len(base_char) <> 1 Then Err.Raise 9999, "IsNumChar", "使える文字列は1文字だけです。"
    IsNumChar = (base_char Like "#" Or (enable_dot And base_char = "."))
End Function

'文字列が数字か否か
'IsNumericは指数表記とかもTrueを返す。コンマは自由に無視するため使えない。
'オーバーフロー検査はしていない
' @param base_str   文字列
' @param enable_dot ドットを数字の一部と見なすか
Public Function IsNumStr(base_str, Optional enable_dot As Boolean = True) As Boolean
    Dim i As Long
    IsNumStr = False
    If Len(base_str) = 0 Then Exit Function
    If Left(base_str, 1) = "." Then Exit Function
    For i = 1 To Len(base_str)
        If Not IsNumChar(Mid(base_str, i, 1), enable_dot) Then Exit Function
    Next
    IsNumStr = True
End Function

