Attribute VB_Name = "kccFuncCore_Partial"
Rem --------------------------------------------------
Rem
Rem @module kccFuncCore_Partial
Rem
Rem @description
Rem    �K�{�֐��������W�߂����W���[��
Rem�@�@�@���璊�o��������
Rem
Rem --------------------------------------------------
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

