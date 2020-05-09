Attribute VB_Name = "AppMain"
Rem
Rem @appname MoneyCostChecker - �������R�X�g�`�F�b�J�[
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/02/15 : �����
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "�������R�X�g�`�F�b�J�["
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.01"
Public Const APP_UPDATE = "2020/05/09"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/money-cost-checker"

'--------------------------------------------------
'�A�h�C���@�\���s��
Sub AddinStart()
    FormMoneyCostChecker.Show
End Sub

'�A�h�C���@�\��~��
Sub AddinStop()
End Sub

'�A�h�C���ݒ�\��
Sub AddinConfig()
End Sub

'�A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'�A�h�C���I����
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

'--------------------------------------------------
