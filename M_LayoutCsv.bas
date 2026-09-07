Attribute VB_Name = "M_LayoutCsv"
'==============================================================================
' M_LayoutCsv.bas
' 責務  : 取込 CSV のレイアウト定義
' 依存  : なし
'
' CSV の列構成が変わった場合、修正はこのモジュールのみで完結させること。
' 業務ロジック側に列番号を直接記述しない。
'==============================================================================
Option Explicit

'--- 列インデックス(0 起算)---
Public Enum ColCsv
    ColCsv_品番 = 0
    ColCsv_数量 = 1
    ColCsv_納期 = 2
    ColCsv_備考 = 3
End Enum

'--- 想定する列数 ---
Public Const CSV_COL_COUNT As Long = 4

'--- ヘッダ行の有無 ---
Public Const CSV_HAS_HEADER As Boolean = True

'--- ヘッダ検証用の文字列(1 列目のみ確認する)---
Public Const CSV_HEADER_品番 As String = "品番"

'--- 文字コード ---
Public Const CSV_CHARSET As String = "Shift_JIS"
