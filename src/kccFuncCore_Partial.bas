Attribute VB_Name = "kccFuncCore_Partial"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncCore_Partial
Rem
Rem  @description   �K�{�֐��������W�߂����W���[�����璊�o��������
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem �������ۂ�
Rem  @param base_char   �ꕶ��
Rem  @param enable_dot  �h�b�g�𐔎��̈ꕔ�ƌ��Ȃ���
Rem  @return As Boolean True:=����,False:=�����ȊO
Rem  @note IsNumeric�͎w���\�L�Ƃ���True��Ԃ����ߎg���Ȃ����ߕK�v
Rem
Private Function IsNumChar(base_char, Optional enable_dot As Boolean = True) As Boolean
    If Len(base_char) <> 1 Then Err.Raise 9999, "IsNumChar", "�g���镶�����1���������ł��B"
    IsNumChar = (base_char Like "#" Or (enable_dot And base_char = "."))
End Function

'�����񂪐������ۂ�
'IsNumeric�͎w���\�L�Ƃ���True��Ԃ��B�R���}�͎��R�ɖ������邽�ߎg���Ȃ��B
'�I�[�o�[�t���[�����͂��Ă��Ȃ�
' @param base_str   ������
' @param enable_dot �h�b�g�𐔎��̈ꕔ�ƌ��Ȃ���
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

Rem RGB�̃t�F�[�h�C��
Rem
Rem  @param before_rgb  �i��  0�̂Ƃ���RGB
Rem  @param after_rgb   �i��100�̂Ƃ���RGB
Rem  @param step_per    �i����(0�`100)
Rem
Public Function GetFadeColor( _
        before_rgb As Long, _
        after_rgb As Long, _
        step_percent As Long) As Long
    If 0 > step_percent Or step_percent > 100 Then Err.Raise 9999, , "GetFadeinColor:0�`100���w�肵�ĉ�����"
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
