Attribute VB_Name = "MakeAlgName"
Sub Make_AlgName(WRKSH As String, PRM As String, DateType As String, RangeOut As String)

Dim str As String 'наименование параметра из IO_List
Dim words() As String ' массив сос ловами параметра
Dim algName() As String 'массив с част€ми алгоритмического имени ????????
Dim wordsLib() As String 'массив с част€ми строчки из Lib
Dim strF As String
Dim loArray, hiArray As Integer ' минимальное и максимальное значение индекса динамического массива
Dim countLib As Integer 'количество непустых строк на листе Lib дл€ наименовани€ параметра (столбец A)
Dim countLibAlgName As Integer 'количество непустых строк на листе Lib дл€ коректировки алгоритмических имен (столбец D)
Dim countPar As Integer 'количество параметров на листе
Dim countRes As Integer '—четчик резервов

countRes = 0 'обнул€ем счетчик резервово
countPar = WorksheetFunction.CountA(Worksheets(WRKSH).Columns("F")) 'определ€ем количество строк на листе параметра
countLib = WorksheetFunction.CountA(Worksheets("Lib").Columns("A")) 'получаем количество непустых строк на листе Lib
countLibAlgName = WorksheetFunction.CountA(Worksheets("Lib").Columns("D")) 'получаем количество непустых строк на листе Lib

'”бираем старые значени€ и заливку €чеек
For i = 4 To countPar + 2
Worksheets(WRKSH).Range(RangeOut & CStr(i)).Value = ""
Worksheets(WRKSH).Range(RangeOut & CStr(i)).Interior.Color = -4142
Next i

For iiii = 4 To countPar + 2 ' бежим по параметрам
str = Trim(Worksheets(WRKSH).Range("F" & CStr(iiii)).Value) ' записывем наименование параметра в переменную
words = Split(str) 'заносим слова в динамический массив
algName = words
loArray = LBound(words) 'нижн€€ граница индекса элементов массива
hiArray = UBound(words) 'верхн€€ граница индекса элементов массива

For i = loArray To hiArray ' цикл слов в напименовании параметра
 For ii = 1 To countLib ' цикл выражений в Lib
   strF = Worksheets("Lib").Range("A" & ii).Value ' записывем строчку из Lib
   wordsLib = Split(strF, ",") 'заносим слова из строчки Lib в динамический массив
   For iii = LBound(wordsLib) To UBound(wordsLib) ' цикл слов в €чейке A Lib
    If InStr(1, words(i), wordsLib(iii), vbBinaryCompare) Then
    algName(i) = Worksheets("Lib").Range("B" & CStr(ii)).Value
    End If
   Next iii
 Next ii
Next i
a = Join(algName, "_") 'сборка переменной
 ' транслитераци€ оставшейс€ части
 iRussian$ = "абвгдеЄжзийклмнопрстуфхцчшщъыьэю€.,/():;'""- "
    iTranslit = Array("", "a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "jj", "k", _
                      "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ch", _
                      "sh", "zch", "", "y", "", "eh", "ju", "ja", "", "", "", "", "", "", "", "", "", "", "")
    For iCount% = 1 To 43
        a = Replace(a, Mid(iRussian$, iCount%, 1), iTranslit(iCount%), , , vbTextCompare)
    Next
    If InStr(1, str, "резерв", 1) Then
    a = a + "_" & CStr(countRes)
    countRes = countRes + 1
    End If
'корекировка алгоритмического имени
For i = 1 To countLibAlgName
a = Replace(a, Worksheets("Lib").Range("D" & CStr(i)).Value, Worksheets("Lib").Range("E" & CStr(i)).Value, , , vbBinaryCompare)
Next i
    
Worksheets(WRKSH).Range(RangeOut & CStr(iiii)).Value = PRM & a
'блок создани€ данных дл€ переноса в тип данных —онаты
'Worksheets(WRKSH).Range("R" & CStr(iiii)).Value = DateType
'Worksheets(WRKSH).Range("S" & CStr(iiii)).Value = "0"
'Worksheets(WRKSH).Range("T" & CStr(iiii)).Value = Worksheets(WRKSH).Range("F" & CStr(iiii))

Next iiii

End Sub
