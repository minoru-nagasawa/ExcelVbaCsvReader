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

' CSV�t�@�C���ǂݍ��݃N���X
' �ȉ��̂悤�Ɏg�p����B
'   Dim reader as CsvReader
'   Set reader = new CsvReader
'   reader.OpenCsv("C:\xx\yy.csv")
'   Debug.Print reader.At(1,1)

' �ǂݍ��񂾃f�[�^�BScripting.Dictionary�^
' �s�Ɨ�̒l���J���}�łȂ����������L�[�ɂ��Ă���B
' �Ⴆ�Έȉ�
'  Key: "1,1"       Value: "xx"
'  Key: "1,2"       Value: "yy"
'  Key: "1,3"       Value: "zz"
'  Key: "2,1"       Value: "XX"
Private m_allData

' �ǂݍ���CSV�̍s��
Private m_maxRow As Long

' �ǂݍ���CSV�̗�
Private m_maxColumn As Long

' CSV�t�@�C����ǂݍ���Ń����o�Ɋi�[����
Public Sub OpenCsv(fullPath As String)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(fullPath) Then
        Call Err.Raise(1, "OpenCsv", "�t�@�C����������܂���")
    End If
    
    ' �t�@�C����S���ǂݍ���
    Dim file
    Set file = fso.GetFile(fullPath)
    Dim stream
    Set stream = file.OpenAsTextStream
    
    Dim text As String
    text = stream.ReadAll
    
    stream.Close
    Set file = Nothing
    Set fso = Nothing
    
    ' CSV�ǂݍ��ݗp�̐��K�\���𐶐�
    Dim regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.MultiLine = True
    regex.Pattern = "(,|\r?\n|^)(""(?:(?:"""")*[^""]*)*""|[^"",\r\n]*)" '�Q�l�ɂ���: https://stackoverflow.com/questions/18144431/regex-to-split-a-csv
                                                                        '�m�F�p�F    https://regex101.com/r/qpitgR/1
    
    ' ���K�\���Ń}�b�`���O���āA���ʂ��i�[
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
        
'        ' ���s���������玟�̍s�Ƃ��ĉ��߂���
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
            ' Excel��CSV���J�����Ƃ��Ɠ������ʂɂȂ�悤�ɁA
            ' �O��̃_�u���N�H�[�e�[�V�������폜���āA
            ' 2�A���̃_�u���N�H�[�e�[�V������1�ɂ���B
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

' CSV�̎w�肵���ꏊ�̒l��Ԃ�
' ������Ȃ�������󕶎���Ԃ�
Public Function At(rowIndex As Long, columnIndex As Long) As String
    Dim key As String
    key = rowIndex & "," & columnIndex
    If m_allData.Exists(key) Then
        At = m_allData.Item(key)
    Else
        At = ""
    End If
End Function

' �s����Ԃ�
Public Function RowCount() As Long
    RowCount = m_maxRow
End Function

' ��s����Ԃ�
Public Function ColumnCount() As Long
    ColumnCount = m_maxColumn
End Function
