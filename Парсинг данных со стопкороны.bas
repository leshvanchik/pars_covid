Attribute VB_Name = "Module1"
Sub Grade()
Dim XMLHTTP As Object
Dim myBook, loadBook As String
Dim r As Double
Dim k, m, n, g, h, j, t As Long
Dim i As Date
i = Date
Dim URL, Kurs, Purs, Txt, a, b, c, d, y, z As String

z = ThisWorkbook.Worksheets("Свод").Cells(4, 10)
y = ThisWorkbook.Worksheets("Свод").Cells(8, 10)
a = Dir(z & "*РПН*.xlsx")
b = z & "Оперативный мониторинг готовности региональных систем здравоохранения к госпитализации больных пневмонией " & i - 2 & ".xlsx"
c = Dir(z & "*за*.xlsx")
d = z & "Доступность лабораторий и тестов " & i - 1 & ".xlsx"

t = DateDiff("s", "01.01.1970 03:00:00", Now)
URL = "https://yastatic.net/s3/milab/2020/podomam/data/index_data.json?ts=" & t
Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")


    myBook = ThisWorkbook.Name
    loadBook = Dir(z & a)
    GetObject (z & a)
    k = ThisWorkbook.Worksheets("Летал_Темп_Заболеваемость СПб").Cells(Rows.Count, 1).End(xlUp).Row
    With Workbooks(myBook).Worksheets("Летал_Темп_Заболеваемость СПб")
    .Range("B" & k + 1) = Workbooks(loadBook).Worksheets(1).Range("C14").Value
    .Range("C" & k + 1) = Workbooks(loadBook).Worksheets(1).Range("B14").Value
    .Range("D" & k + 1) = Workbooks(loadBook).Worksheets(1).Range("D14").Value
    .Range("E" & k + 1) = Workbooks(loadBook).Worksheets(1).Range("E14").Value

    .Range("A" & k + 1) = i
    .Range("F" & k & ":J" & k).Copy ThisWorkbook.Worksheets("Летал_Темп_Заболеваемость СПб").Range("F" & k & ":J" & k).Offset(1, 0)
    .Range("B" & k + 1).Copy ThisWorkbook.Worksheets("Rt").Range("C" & k)
    End With
    Workbooks(loadBook).Close (False)
    
XMLHTTP.Open "GET", URL, False
'XMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.101 Safari/537.36"
XMLHTTP.SEND
If XMLHTTP.Status = 200 Then
    Txt = XMLHTTP.responseText
    
    n = InStr(1, Txt, "Екатеринбург") - 19
    g = InStr(n - 18, Txt, ",""diff")
    Kurs = Mid(Txt, n, g - n)
    Kurs = Replace(Replace(Kurs, ".", ","), ":", "")
    r = Kurs
    ThisWorkbook.Worksheets("Rt").Range("E" & k) = r

    h = InStr(1, Txt, "Екатеринбург") - 41
    j = InStr(h, Txt, ",""value")
    Purs = Mid(Txt, h, j - h)
    Purs = Replace(Replace(Purs, ":", ""), """", "")
    MsgBox "Индекс самоизоляции вставлен за " & Purs
Else
    MsgBox "Отсутствует интернет-соединение..."
End If
Set XMLHTTP = Nothing

ThisWorkbook.Worksheets("Rt").Range("A" & k) = i
ThisWorkbook.Worksheets("Rt").Range("B" & k) = 1
ThisWorkbook.Worksheets("Rt").Range("D" & k - 1).Copy ThisWorkbook.Worksheets("Rt").Range("D" & k - 1).Offset(1, 0)

    loadBook = Dir(b)
    GetObject (b)
    k = ThisWorkbook.Worksheets("СКФ").Cells(Rows.Count, 1).End(xlUp).Row
    With Workbooks(myBook).Worksheets("СКФ")
    .Range("B" & k + 1 & ":U" & k + 1) = Workbooks(loadBook).Worksheets(1).Range("A37:T37").Value
    .Range("V" & k & ":Z" & k).Copy ThisWorkbook.Worksheets("СКФ").Range("V" & k + 1 & ":Z" & k + 1)
    .Range("A" & k + 1) = i - 2
    End With
    Workbooks(loadBook).Close (False)

    loadBook = Dir(z & c)
    GetObject (z & c)
    k = ThisWorkbook.Worksheets("ОТ СПб").Cells(Rows.Count, 1).End(xlUp).Row
    With Workbooks(myBook).Worksheets("ОТ СПб")
    .Range("P" & k + 1) = Workbooks(loadBook).Worksheets(1).Range("V5").Value
    .Range("A" & k + 1) = i
    .Range("V" & k & ":W" & k).Copy ThisWorkbook.Worksheets("ОТ СПб").Range("V" & k + 1 & ":W" & k + 1)
    .Range("R" & k + 1).Formula = "=SUM(R[-1]C, RC[-2])"
    End With
    Workbooks(loadBook).Close (False)

    loadBook = Dir(d)
    GetObject (d)
    m = ThisWorkbook.Worksheets("ОТ РФ").Cells(Rows.Count, 1).End(xlUp).Row
    Workbooks(myBook).Worksheets("ОТ СПб").Range("B" & k & ":U" & k) = Workbooks(loadBook).Worksheets(1).Range("A35:T35").Value
    With Workbooks(myBook).Worksheets("ОТ РФ")
    .Range("A" & m + 1) = i - 1
    .Range("B" & m + 1 & ":U" & m + 1) = Workbooks(loadBook).Worksheets(1).Range("A4:T4").Value
    .Range("V" & m & ":W" & m).Copy ThisWorkbook.Worksheets("ОТ РФ").Range("V" & m + 1 & ":W" & m + 1)
    End With
    Workbooks(loadBook).Close (False)

    Workbooks.Add
    ActiveWorkbook.Worksheets(1).Range("A1:O10") = ThisWorkbook.Worksheets("ЗАГРУЗОЧНЫЙ").Range("B1:P10").Value
    ActiveWorkbook.SaveAs Filename:=y & " " & i & ".xlsx"
    ActiveWorkbook.Close (True)

End Sub
