Attribute VB_Name = "Svc_CsvImport"
'==============================================================================
' Svc_CsvImport.bas
' 責務  : CSV 取込の業務処理(読込 → 検証 → 登録)
' 依存  : Util_Csv, Util_Convert, Dao_Common, Dao_Import, M_LayoutCsv
' 禁止  : MsgBox の表示、シート操作
' エラー方式 : 例外(Err.Raise)。呼び出し元は上位で捕捉すること
'
' 取込ポリシー:
'   1 行でも検証に失敗した場合、全件をロールバックする。
'   部分登録を許容する運用に変更する場合は 取込実行 の分岐を修正すること。
'==============================================================================
Option Explicit

' 取込結果
Public Type ImportResult
    成功件数 As Long
    エラー行 As String        ' "3, 15, 22" 形式。空文字なら正常終了
End Type

'------------------------------------------------------------------------------
' CSV ファイルを取り込む
'   filePath : 取込対象のフルパス
'   戻り値   : 取込結果
'------------------------------------------------------------------------------
Public Function 取込実行(ByVal filePath As String) As ImportResult

    Dim cn As ADODB.Connection
    Dim inTran As Boolean
    Dim errNo As Long, errSrc As String, errMsg As String

    On Error GoTo ErrHandler

    '--- 1. 読み込み ---
    Dim data As Variant
    data = Util_Csv.読込(filePath)
    If IsEmpty(data) Then
        Err.Raise vbObjectError + 610, "Svc_CsvImport.取込実行", _
                  "ファイルにデータがありません。"
    End If

    '--- 2. 形式検証(列数・ヘッダ)---
    Call 形式検証(data)

    '--- 3. 登録 ---
    Set cn = Dao_Common.接続を開く()
    Dao_Common.トランザクション開始 cn
    inTran = True

    Dim errRows As Collection
    Set errRows = New Collection

    Dim r As Long, startRow As Long
    startRow = IIf(CSV_HAS_HEADER, 1, 0)

    For r = startRow To UBound(data, 1)
        If Not 空行判定(data, r) Then
            If 行検証(data, r) Then
                Call Dao_Import.登録(cn, data, r)
                取込実行.成功件数 = 取込実行.成功件数 + 1
            Else
                errRows.Add r + 1                 ' 1 起算の行番号で記録
            End If
        End If
    Next r

    '--- 4. 確定 / 取消 ---
    If errRows.Count > 0 Then
        Dao_Common.ロールバック cn
        inTran = False
        取込実行.成功件数 = 0
        取込実行.エラー行 = 行番号連結(errRows)
    Else
        Dao_Common.コミット cn
        inTran = False
    End If

CleanUp:
    On Error Resume Next                          ' 後始末中のエラーは無視する
    If inTran Then Dao_Common.ロールバック cn
    Dao_Common.接続を閉じる cn
    On Error GoTo 0

    If errNo <> 0 Then Err.Raise errNo, errSrc, errMsg
    Exit Function

ErrHandler:
    errNo = Err.Number
    errSrc = Err.Source
    errMsg = Err.Description
    Resume CleanUp
End Function

'------------------------------------------------------------------------------
' 形式検証:ファイル全体の構造を確認する
'   不正な場合は例外を送出する(処理継続不能のため)
'------------------------------------------------------------------------------
Private Sub 形式検証(ByVal data As Variant)

    If UBound(data, 2) + 1 < CSV_COL_COUNT Then
        Err.Raise vbObjectError + 611, "Svc_CsvImport.形式検証", _
                  "列数が不足しています。想定 " & CSV_COL_COUNT & " 列、" & _
                  "実際 " & (UBound(data, 2) + 1) & " 列"
    End If

    If CSV_HAS_HEADER Then
        If Trim$(data(0, ColCsv_品番)) <> CSV_HEADER_品番 Then
            Err.Raise vbObjectError + 612, "Svc_CsvImport.形式検証", _
                      "ヘッダの形式が想定と異なります。" & vbCrLf & _
                      "1 列目: " & Trim$(data(0, ColCsv_品番))
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' 行検証:1 行分の入力値を確認する
'   戻り値 : True = 正常 / False = 不正
'   ※ 想定内の分岐のため戻り値で返す(例外にしない)
'------------------------------------------------------------------------------
Private Function 行検証(ByVal data As Variant, ByVal r As Long) As Boolean

    '--- 品番:必須 ---
    If Util_Convert.IsBlank(data(r, ColCsv_品番)) Then Exit Function

    '--- 数量:数値・1 以上 ---
    Dim qty As Long
    If Not Util_Convert.TryToLong(data(r, ColCsv_数量), qty) Then Exit Function
    If qty <= 0 Then Exit Function

    '--- 納期:入力があれば日付として解釈できること ---
    If Not Util_Convert.IsBlank(data(r, ColCsv_納期)) Then
        If Not IsDate(data(r, ColCsv_納期)) Then Exit Function
    End If

    行検証 = True
End Function

'------------------------------------------------------------------------------
' 空行判定:全列が未入力なら True
'   末尾の空行を取込対象外とするために使用する
'------------------------------------------------------------------------------
Private Function 空行判定(ByVal data As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 0 To UBound(data, 2)
        If Not Util_Convert.IsBlank(data(r, c)) Then Exit Function
    Next c
    空行判定 = True
End Function

'------------------------------------------------------------------------------
' 行番号を "3, 15, 22" 形式へ連結する
'------------------------------------------------------------------------------
Private Function 行番号連結(ByVal col As Collection) As String
    Dim s As String, v As Variant
    For Each v In col
        If s <> "" Then s = s & ", "
        s = s & v
    Next v
    行番号連結 = s
End Function
