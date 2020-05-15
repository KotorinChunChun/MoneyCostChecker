Attribute VB_Name = "WorksheetFunction"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        WorksheetFunction
Rem
Rem  @description   Excel以外のVBAでもExcel.WorksheetFunctionと同じ関数が使える関数群
Rem
Rem  @update        2020/05/15
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

'指定した配列のうち最大の値を返す
Public Function Max(ParamArray Nums() As Variant) As Variant
    Dim Num As Variant
    Max = Nums(0)
    For Each Num In Nums
        If Max < Num Then Max = Num
    Next
End Function
