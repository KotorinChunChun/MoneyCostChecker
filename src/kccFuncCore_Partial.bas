Attribute VB_Name = "kccFuncCore_Partial"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncCore_Partial
Rem
Rem  @description   必須関数だけを集めたモジュールから抽出したもの
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
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

Rem RGBのフェードイン
Rem
Rem  @param before_rgb  進捗  0のときのRGB
Rem  @param after_rgb   進捗100のときのRGB
Rem  @param step_per    進捗率(0～100)
Rem
Public Function GetFadeColor( _
        before_rgb As Long, _
        after_rgb As Long, _
        step_percent As Long) As Long
    If 0 > step_percent Or step_percent > 100 Then Err.Raise 9999, , "GetFadeinColor:0～100を指定して下さい"
    Dim DiffR: DiffR = (CDbl(RgbToRed(after_rgb)) - RgbToRed(before_rgb)) * step_percent / 100
    Dim DiffG: DiffG = (CDbl(RgbToGreen(after_rgb)) - RgbToGreen(before_rgb)) * step_percent / 100
    Dim DiffB: DiffB = (CDbl(RgbToBlue(after_rgb)) - RgbToBlue(before_rgb)) * step_percent / 100
    GetFadeColor = RGB( _
        RgbToRed(before_rgb) + DiffR, _
        RgbToGreen(before_rgb) + DiffG, _
        RgbToBlue(before_rgb) + DiffB)
End Function

Public Function RgbToRed(ByVal rgb_value As Long) As Byte
    RgbToRed = &HFF& And rgb_value
End Function

Public Function RgbToGreen(ByVal rgb_value As Long) As Byte
    RgbToGreen = (&HFF00& And rgb_value) \ 256
End Function

Public Function RgbToBlue(ByVal rgb_value As Long) As Byte
    RgbToBlue = (&HFF0000 And rgb_value) \ 65536
End Function
