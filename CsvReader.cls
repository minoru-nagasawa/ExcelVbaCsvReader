VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CSVファイル読み込みクラス
' 以下のように使用する。
'   Dim reader as CsvReader
'   Set reader = new CsvReader
'   reader.OpenCsv("C:\xx\yy.csv")
'   Debug.Print reader.At(1,1)

' 読み込んだデータ。Scripting.Dictionary型
' 行と列の値をカンマでつなげた文字をキーにしている。
' 例えば以下
'  Key: "1,1"       Value: "xx"
'  Key: "1,2"       Value: "yy"
'  Key: "1,3"       Value: "zz"
'  Key: "2,1"       Value: "XX"
Private m_allData

' 読み込んだCSVの行数
Private m_maxRow As Long

' 読み込んだCSVの列数
Private m_maxColumn As Long

' CSVファイルを読み込んでメンバに格納する
Public Sub OpenCsv(fullPath As String)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(fullPath) Then
        Call Err.Raise(1, "OpenCsv", "ファイルが見つかりません")
    End If
    
    ' ファイルを全部読み込む
    Dim file
    Set file = fso.GetFile(fullPath)
    Dim stream
    Set stream = file.OpenAsTextStream
    
    Dim text As String
    text = stream.ReadAll
    
    stream.Close
    Set file = Nothing
    Set fso = Nothing
    
    ' CSV読み込み用の正規表現を生成
    Dim regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.MultiLine = True
    regex.Pattern = "(,|\r?\n|^)(""(?:(?:"""")*[^""]*)*""|[^"",\r\n]*)" '参考にした: https://stackoverflow.com/questions/18144431/regex-to-split-a-csv
                                                                        '確認用：    https://regex101.com/r/qpitgR/1
    
    ' 正規表現でマッチングして、結果を格納
    Set m_allData = CreateObject("Scripting.Dictionary")
    Dim row As Long, column As Long
    row = 1
    column = 1
    
    Dim matches
    Set matches = regex.Execute(text)
    Dim mc As Long
    For mc = 0 To matches.count - 1
        Dim sm
        Set sm = matches(mc).SubMatches
        
'        ' 改行を見つけたら次の行として解釈する
        If sm(0) = vbLf Or sm(0) = vbCrLf Then
            If m_maxColumn < column - 1 Then
                m_maxColumn = column - 1
            End If
            row = row + 1
            column = 1
        End If

        Dim key As String
        key = row & "," & column
        If sm(1) <> "" Then
            Dim t As String
            t = sm(1)
            ' ExcelでCSVを開いたときと同じ結果になるように、
            ' 前後のダブルクォーテーションを削除して、
            ' 2つ連続のダブルクォーテーションを1つにする。
            If InStr(1, t, """") <> 0 Then
                If Left(t, 1) = """" Then
                    t = Mid(t, 2, Len(t) - 2)
                End If
                t = Replace(t, """""", """")
            End If
            Call m_allData.Add(key, t)
        Else
            Call m_allData.Add(key, "")
        End If
        column = column + 1
    Next
    m_maxRow = row
End Sub

' CSVの指定した場所の値を返す
' 見つからなかったら空文字を返す
Public Function At(rowIndex As Long, columnIndex As Long) As String
    Dim key As String
    key = rowIndex & "," & columnIndex
    If m_allData.Exists(key) Then
        At = m_allData.Item(key)
    Else
        At = ""
    End If
End Function

' 行数を返す
Public Function RowCount() As Long
    RowCount = m_maxRow
End Function

' 列行数を返す
Public Function ColumnCount() As Long
    ColumnCount = m_maxColumn
End Function

