Attribute VB_Name = "Dao_Import"
'==============================================================================
' Dao_Import.bas
' 責務  : T_IMPORT テーブルへの登録
' 依存  : ADODB, Dao_Common, M_LayoutCsv
' 禁止  : Form / Service / シートの参照、MsgBox の表示
' エラー方式 : 例外(Err.Raise)。呼び出し元は上位で捕捉すること
'
' Connection は引数で受け取る。所有者は呼び出し元(Service)であり、
' 当モジュールでは開閉しない。
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' CSV の 1 行を登録する
'   cn   : 呼び出し元が開いた接続(トランザクション内で呼ばれる想定)
'   data : Util_Csv.読込 が返した 2 次元配列
'   r    : 対象行のインデックス(0 起算)
'------------------------------------------------------------------------------
Public Sub 登録(ByVal cn As ADODB.Connection, _
                ByVal data As Variant, _
                ByVal r As Long)

    Dim cmd As ADODB.Command
    Dim errNo As Long, errSrc As String, errMsg As String

    On Error GoTo ErrHandler

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO T_IMPORT " & _
                      "(HINBAN, SURYO, NOUKI, BIKO) " & _
                      "VALUES (?, ?, ?, ?)"

    ' パラメータは ? の出現順に Append すること(Access は名前解決しない)
    Dao_Common.AddStr cmd, "pHinban", 20, data(r, ColCsv_品番)
    Dao_Common.AddLong cmd, "pSuryo", data(r, ColCsv_数量)
    Dao_Common.AddDate cmd, "pNouki", data(r, ColCsv_納期)
    Dao_Common.AddStr cmd, "pBiko", 255, data(r, ColCsv_備考)

    Dim affected As Long
    cmd.Execute affected, , adExecuteNoRecords

    If affected <> 1 Then
        Err.Raise vbObjectError + 620, "Dao_Import.登録", _
                  "登録件数が不正です(" & affected & " 件 / 行: " & r + 1 & ")"
    End If

CleanUp:
    On Error Resume Next
    If Not cmd Is Nothing Then
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
    End If
    On Error GoTo 0

    If errNo <> 0 Then Err.Raise errNo, errSrc, errMsg
    Exit Sub

ErrHandler:
    errNo = Err.Number
    errSrc = Err.Source
    errMsg = Err.Description
    Resume CleanUp
End Sub

'------------------------------------------------------------------------------
' 取込先テーブルを全件削除する
'   毎回全件を入れ替える運用の場合に、取込前に呼び出す
'   ※ 呼び出し元のトランザクション内で実行すること
'------------------------------------------------------------------------------
Public Sub 全件削除(ByVal cn As ADODB.Connection)
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = "DELETE FROM T_IMPORT"
    cmd.Execute , , adExecuteNoRecords

    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
End Sub
