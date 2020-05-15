Attribute VB_Name = "AppMain"
Rem --------------------------------------------------------------------------------
Rem
Rem  @appname MoneyCostChecker - お賃金コストチェッカー
Rem
Rem  @history
Rem     2020/05/09 : 初回版
Rem     2020/05/15 : 64bit対応版
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        AppMain
Rem
Rem  @description   アドインの実行コマンド群
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit
Option Private Module

Public Const APP_NAME = "お賃金コストチェッカー"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.02"
Public Const APP_UPDATE = "2020/05/15"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/20200509-money-cost-checker"

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

Property Get ThisObject() As Object
#If DEF_EXCEL Then
    Set ThisObject = ThisWorkbook
#ElseIf DEF_WORD Then
    Set ThisObject = ThisDocument
#Else
    Err.Raise 9999, , "未定義のVBAプロジェクト"
#End If
End Property

'アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisObject.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisObject.Path & vbLf & _
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
Sub AddinEnd(): ThisObject.Close False: End Sub

'--------------------------------------------------


