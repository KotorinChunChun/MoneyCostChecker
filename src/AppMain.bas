Attribute VB_Name = "AppMain"
Rem
Rem @appname MoneyCostChecker - お賃金コストチェッカー
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/02/15 : 初回版
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "お賃金コストチェッカー"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.01"
Public Const APP_UPDATE = "2020/05/09"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/money-cost-checker"

'--------------------------------------------------
'アドイン機能実行時
Sub AddinStart()
    FormMoneyCostChecker.Show
End Sub

'アドイン機能停止時
Sub AddinStop()
End Sub

'アドイン設定表示
Sub AddinConfig()
End Sub

'アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'アドイン終了時
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

'--------------------------------------------------
