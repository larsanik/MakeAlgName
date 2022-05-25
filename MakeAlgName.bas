Attribute VB_Name = "MakeAlgName"
Sub Make_AlgName(WRKSH As String, PRM As String, DateType As String, RangeOut As String)

Dim str As String '������������ ��������� �� IO_List
Dim words() As String ' ������ ��� ������ ���������
Dim algName() As String '������ � ������� ���������������� ����� ????????
Dim wordsLib() As String '������ � ������� ������� �� Lib
Dim strF As String
Dim loArray, hiArray As Integer ' ����������� � ������������ �������� ������� ������������� �������
Dim countLib As Integer '���������� �������� ����� �� ����� Lib ��� ������������ ��������� (������� A)
Dim countLibAlgName As Integer '���������� �������� ����� �� ����� Lib ��� ������������ ��������������� ���� (������� D)
Dim countPar As Integer '���������� ���������� �� �����
Dim countRes As Integer '������� ��������

countRes = 0 '�������� ������� ���������
countPar = WorksheetFunction.CountA(Worksheets(WRKSH).Columns("F")) '���������� ���������� ����� �� ����� ���������
countLib = WorksheetFunction.CountA(Worksheets("Lib").Columns("A")) '�������� ���������� �������� ����� �� ����� Lib
countLibAlgName = WorksheetFunction.CountA(Worksheets("Lib").Columns("D")) '�������� ���������� �������� ����� �� ����� Lib

'������� ������ �������� � ������� �����
For i = 4 To countPar + 2
Worksheets(WRKSH).Range(RangeOut & CStr(i)).Value = ""
Worksheets(WRKSH).Range(RangeOut & CStr(i)).Interior.Color = -4142
Next i

For iiii = 4 To countPar + 2 ' ����� �� ����������
str = Trim(Worksheets(WRKSH).Range("F" & CStr(iiii)).Value) ' ��������� ������������ ��������� � ����������
words = Split(str) '������� ����� � ������������ ������
algName = words
loArray = LBound(words) '������ ������� ������� ��������� �������
hiArray = UBound(words) '������� ������� ������� ��������� �������

For i = loArray To hiArray ' ���� ���� � ������������� ���������
 For ii = 1 To countLib ' ���� ��������� � Lib
   strF = Worksheets("Lib").Range("A" & ii).Value ' ��������� ������� �� Lib
   wordsLib = Split(strF, ",") '������� ����� �� ������� Lib � ������������ ������
   For iii = LBound(wordsLib) To UBound(wordsLib) ' ���� ���� � ������ A Lib
    If InStr(1, words(i), wordsLib(iii), vbBinaryCompare) Then
    algName(i) = Worksheets("Lib").Range("B" & CStr(ii)).Value
    End If
   Next iii
 Next ii
Next i
a = Join(algName, "_") '������ ����������
 ' �������������� ���������� �����
 iRussian$ = "��������������������������������.,/():;'""- "
    iTranslit = Array("", "a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "jj", "k", _
                      "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ch", _
                      "sh", "zch", "", "y", "", "eh", "ju", "ja", "", "", "", "", "", "", "", "", "", "", "")
    For iCount% = 1 To 43
        a = Replace(a, Mid(iRussian$, iCount%, 1), iTranslit(iCount%), , , vbTextCompare)
    Next
    If InStr(1, str, "������", 1) Then
    a = a + "_" & CStr(countRes)
    countRes = countRes + 1
    End If
'����������� ���������������� �����
For i = 1 To countLibAlgName
a = Replace(a, Worksheets("Lib").Range("D" & CStr(i)).Value, Worksheets("Lib").Range("E" & CStr(i)).Value, , , vbBinaryCompare)
Next i
    
Worksheets(WRKSH).Range(RangeOut & CStr(iiii)).Value = PRM & a
'���� �������� ������ ��� �������� � ��� ������ ������
'Worksheets(WRKSH).Range("R" & CStr(iiii)).Value = DateType
'Worksheets(WRKSH).Range("S" & CStr(iiii)).Value = "0"
'Worksheets(WRKSH).Range("T" & CStr(iiii)).Value = Worksheets(WRKSH).Range("F" & CStr(iiii))

Next iiii

End Sub
