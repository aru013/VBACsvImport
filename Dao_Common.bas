Attribute VB_Name = "Dao_Common"
'==============================================================================
' Dao_Common.bas(抜粋:CSV 取込で使用する部分)
' 責務  : 接続・トランザクション制御、パラメータ生成の共通処理
' 依存  : ADODB, M_Config, Util_Convert
' 禁止  : Form / Service / シートの参照、MsgBox の表示
' エラー方式 : 例外(Err.Raise)。後始末系(ロールバック / 接続を閉じる)のみ
'              意図的にエラーを無視する
'
' ※ 既存の Dao_Common に以下を追加 / 統合すること。
'    接続を開く / 接続を閉じる / トランザクション系が既にある場合は重複させない。
'==============================================================================
Option Explicit

'==============================================================================
' 接続・トランザクション
'==============================================================================

'------------------------------------------------------------------------------
' 接続を開いて返す
'   所有者は呼び出し元。エラーは捕捉せず、そのまま呼び出し元へ伝播させる
'------------------------------------------------------------------------------
Public Function 接続を開く() As ADODB.Connection
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.Open M_Config.AccessConnectionString()
    Set 接続を開く = cn
End Function

'------------------------------------------------------------------------------
' 接続を閉じる。既に閉じている場合は何もしない
'------------------------------------------------------------------------------
Public Sub 接続を閉じる(ByRef cn As ADODB.Connection)
    On Error Resume Next                  ' 後始末のため意図的に無視する
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' トランザクションを開始する
'------------------------------------------------------------------------------
Public Sub トランザクション開始(ByVal cn As ADODB.Connection)
    If cn Is Nothing Then
        Err.Raise vbObjectError + 400, "Dao_Common.トランザクション開始", _
                  "接続が確立されていません。"
    End If
    cn.BeginTrans
End Sub

Public Sub コミット(ByVal cn As ADODB.Connection)
    cn.CommitTrans
End Sub

'------------------------------------------------------------------------------
' ロールバックする
'   未開始の場合のエラー(3246)は意図的に無視する
'------------------------------------------------------------------------------
Public Sub ロールバック(ByVal cn As ADODB.Connection)
    On Error Resume Next
    If Not cn Is Nothing Then cn.RollbackTrans
    On Error GoTo 0
End Sub

'==============================================================================
' パラメータ生成
'   Access(ACE OLEDB)は名前付きパラメータを解決しない。
'   SQL 中の ? の出現順と Append の順序を必ず一致させること。
'   name は可読性のためのラベルであり、解決には使用されない。
'==============================================================================

'------------------------------------------------------------------------------
' 文字列パラメータ。未入力は Null として渡す
'   size : 列定義以上の長さを指定すること(文字列型では必須)
'------------------------------------------------------------------------------
Public Sub AddStr(ByVal cmd As ADODB.Command, _
                  ByVal name As String, _
                  ByVal size As Long, _
                  ByVal value As Variant)
    Dim p As ADODB.Parameter
    Set p = cmd.CreateParameter(name, adVarWChar, adParamInput, size)

    If Util_Convert.IsBlank(value) Then
        p.value = Null
    Else
        p.value = Left$(Trim$(CStr(value)), size)
    End If

    cmd.Parameters.Append p
End Sub

'------------------------------------------------------------------------------
' 数値(Long)パラメータ。変換できない場合は Null として渡す
'------------------------------------------------------------------------------
Public Sub AddLong(ByVal cmd As ADODB.Command, _
                   ByVal name As String, _
                   ByVal value As Variant)
    Dim p As ADODB.Parameter
    Set p = cmd.CreateParameter(name, adInteger, adParamInput)

    Dim v As Long
    If Util_Convert.TryToLong(value, v) Then
        p.value = v
    Else
        p.value = Null
    End If

    cmd.Parameters.Append p
End Sub

'------------------------------------------------------------------------------
' 日付パラメータ。変換できない場合は Null として渡す
'   文字列で受けた日付をここで型に変換するため、SQL 側の書式依存を避けられる
'------------------------------------------------------------------------------
Public Sub AddDate(ByVal cmd As ADODB.Command, _
                   ByVal name As String, _
                   ByVal value As Variant)
    Dim p As ADODB.Parameter
    Set p = cmd.CreateParameter(name, adDate, adParamInput)

    If Util_Convert.IsBlank(value) Then
        p.value = Null
    ElseIf IsDate(value) Then
        p.value = CDate(value)
    Else
        p.value = Null
    End If

    cmd.Parameters.Append p
End Sub
