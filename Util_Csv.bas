Attribute VB_Name = "Util_Csv"
'==============================================================================
' Util_Csv.bas
' 責務  : CSV ファイルの読み込みとパース
' 依存  : ADODB.Stream(遅延バインディング)
' 禁止  : MsgBox の表示、DB アクセス、業務判断
' 文字コード : Shift-JIS 既定(引数で変更可)
'
' RFC4180 に準拠し、引用符で囲まれた値に含まれるカンマ・改行を正しく扱う。
'==============================================================================
Option Explicit

Private Const DELIM As String = ","

'------------------------------------------------------------------------------
' CSV ファイルを 2 次元配列(0 起算)で返す
'   filePath : 読み込むファイルのフルパス
'   charset  : 文字コード。既定は Shift_JIS
'   戻り値   : 2 次元配列。データが無い場合は Empty
'------------------------------------------------------------------------------
Public Function 読込(ByVal filePath As String, _
                     Optional ByVal charset As String = "Shift_JIS") As Variant

    If Trim$(filePath) = "" Then
        Err.Raise vbObjectError + 600, "Util_Csv.読込", _
                  "ファイルパスが指定されていません。"
    End If
    If Dir(filePath) = "" Then
        Err.Raise vbObjectError + 601, "Util_Csv.読込", _
                  "ファイルが見つかりません: " & filePath
    End If

    Dim text As String
    text = ファイル読込(filePath, charset)
    If Len(text) = 0 Then Exit Function

    読込 = パース(text)
End Function

'------------------------------------------------------------------------------
' ファイル全体を文字列として読み込む
'------------------------------------------------------------------------------
Private Function ファイル読込(ByVal filePath As String, _
                              ByVal charset As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")

    st.Type = 2                      ' adTypeText
    st.charset = charset
    st.Open
    st.LoadFromFile filePath
    ファイル読込 = st.ReadText
    st.Close
    Set st = Nothing
End Function

'------------------------------------------------------------------------------
' CSV 文字列を 2 次元配列へ分解する
'   引用符内のカンマ・改行はデータとして扱う。
'   引用符のエスケープ("")にも対応する。
'------------------------------------------------------------------------------
Private Function パース(ByVal text As String) As Variant
    Dim rows As Collection
    Set rows = New Collection

    Dim fields As Collection
    Set fields = New Collection

    Dim buf As String
    Dim inQuote As Boolean
    Dim i As Long, n As Long
    Dim c As String, nextC As String

    n = Len(text)

    For i = 1 To n
        c = Mid$(text, i, 1)

        If inQuote Then
            '--- 引用符の内側 ---
            If c = """" Then
                nextC = Mid$(text, i + 1, 1)
                If nextC = """" Then
                    buf = buf & """"            ' エスケープされた引用符
                    i = i + 1
                Else
                    inQuote = False             ' 引用終了
                End If
            Else
                buf = buf & c                   ' カンマ・改行もそのまま保持
            End If

        Else
            '--- 引用符の外側 ---
            Select Case c
                Case """"
                    inQuote = True

                Case DELIM
                    fields.Add buf
                    buf = ""

                Case vbCr
                    ' CRLF の CR は読み飛ばし、次の LF で行を確定する
                    If Mid$(text, i + 1, 1) <> vbLf Then
                        fields.Add buf
                        buf = ""
                        rows.Add コレクションを配列化(fields)
                        Set fields = New Collection
                    End If

                Case vbLf
                    fields.Add buf
                    buf = ""
                    rows.Add コレクションを配列化(fields)
                    Set fields = New Collection

                Case Else
                    buf = buf & c
            End Select
        End If
    Next i

    '--- 最終行(改行で終わっていない場合)---
    If Len(buf) > 0 Or fields.Count > 0 Then
        fields.Add buf
        rows.Add コレクションを配列化(fields)
    End If

    If rows.Count = 0 Then Exit Function
    パース = 二次元化(rows)
End Function

'------------------------------------------------------------------------------
' Collection を 1 次元配列(0 起算)へ変換する
'------------------------------------------------------------------------------
Private Function コレクションを配列化(ByVal col As Collection) As Variant
    Dim arr() As String
    ReDim arr(0 To col.Count - 1)

    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i

    コレクションを配列化 = arr
End Function

'------------------------------------------------------------------------------
' 行コレクションを 2 次元配列へ変換する
'   列数は全行の最大値に合わせ、不足分は空文字で埋める
'------------------------------------------------------------------------------
Private Function 二次元化(ByVal rows As Collection) As Variant
    Dim maxCol As Long, i As Long
    For i = 1 To rows.Count
        If UBound(rows(i)) + 1 > maxCol Then maxCol = UBound(rows(i)) + 1
    Next i
    If maxCol = 0 Then Exit Function

    Dim result() As String
    ReDim result(0 To rows.Count - 1, 0 To maxCol - 1)

    Dim r As Long, c As Long
    For r = 1 To rows.Count
        For c = 0 To UBound(rows(r))
            result(r - 1, c) = rows(r)(c)
        Next c
    Next r

    二次元化 = result
End Function
