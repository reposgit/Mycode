Attribute VB_Name = "CalcInteropModule1"
Private calc As Object
Private ei As Object
'версия модуля'
Public Function getversion() As String
    getversion = 47
End Function
'Возвращает класс чистоты по индексу загрязненности по ГОСТ 17216
Public Function CalcClassChist2(index As Double)
    CalcClassChist2 = GetEI().GetCalculations().CalcClassChist2(index)
End Function
'Интерполяция
Public Function Interpol(x As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    Interpol = GetEI().Interpol(x, x1, y1, x2, y2)
End Function
'Округляет число до необходимого количества значащих цифр, n - число, Nznach - количество значащих цифр'
Public Function OkrToZnach(n As String, Nznach As Integer)
    OkrToZnach = GetEI().GetCalculations().OkrToZnach(n, Nznach)
End Function
'Расчет объема по высоте адсорбента для ГОСТ 11382, docid - шифр документа с данными для расчета, x - показатель'
Public Function CalcVolume(docid As String, x As Double) As Double
    CalcVolume = GetEI().GetCalculations().CalcVolume(docid, x)
End Function
'Округляет число до необходимого количества значащих цифр, Val - число, Count - количество значащих цифр'
'Не работает на данный момент'
Public Function Substring(Val As String, count As Integer)
    Substring = GetEI().GetCalculations().Substring(Val, count, count)
End Function
'возвращает по номеру пробы или иного документа значение указанного показателя из этого документа; аргументы: batchno- номер пробы, key - ключ (это может быть дата анализа, результат анализа, исполнитель, погешность и т.д. из указанного документа)'
Public Function getBatchDocByKey(batchno As String, key As String) As String
    getBatchDocByKey = GetWI().getBatchDocByKey(batchno, key)
End Function
'преобразует строку в массив согласно делителю и возвращает элемент массива; str - строка, index - номер элемента массива, delimeter - разделитель'
Public Function SplitData(str As String, index As Long, delimeter As String) As String
    Dim a As Variant
    a = Split(str, delimeter)
    SplitData = a(index)
End Function
'Возвращает метаданные в формате даты по номеру документа; аргументы: docid - номер документа, key - ключ, defval - формат даты'
Public Function getDocMetadataAsDate(docid As String, key As String, defval As Date) As Date
    getDocMetadataAsDate = tryToDate(GetBatchMetaData(docid, docid, key), defval)
End Function
'Возвращает метаданные в формате числа; аргументы: docid - номер документа, batchno - номер пробы, key - ключ, defval - значение по умолчанию'
Public Function getBatchMetadataAsDouble(docid As String, batchno As String, key As String, defval As Double) As Double
    getBatchMetadataAsDouble = tryToNumber(GetBatchMetaData(docid, batchno, key), defval)
End Function
'Возвращает метаданные в формате числа по номеру документа; аргументы: docid - номер документа, key - ключ, defval - значение по умолчанию'
Public Function getDocBatchMetadataAsDouble(docid As String, key As String, defval As Double) As Double
    getDocBatchMetadataAsDouble = getBatchMetadataAsDouble(SplitData(docid, 0, "|"), SplitData(docid, 1, "|"), key, defval)
End Function
'Возвращает метаданные в формате числа по номеру документа; аргументы: docid - номер документа, key - ключ, defval - значение по умолчанию'
Public Function getDocMetadataAsDouble(docid As String, key As String, defval As Double) As Double
    getDocMetadataAsDouble = tryToNumber(GetBatchMetaData(docid, docid, key), defval)
End Function
'Возвращает информацию из базы данных; аргументы: docid - номер документа, batchno - номер пробы, key - ключ'
Public Function GetBatchMetaData(docid As String, batchno As String, key As String) As String
    GetBatchMetaData = GetEI().GetCalculations().GetBatchMetaData(docid, batchno, key)
End Function
'Возвращает список показателей по пробе; аргументы: batchno - номер пробы
Public Function GetTestNamesByBatchNo(batchno As String) As String
    GetTestNamesByBatchNo = GetWI().GetTestNamesByBatchNo(batchno)
End Function
'Возвращает метаданные для методов с множеством показателей; аргументы: dodcid - номер документа, batchno - номер пробы, key - ключ, methodname - имя метода, testname - имя показателя, opr - номер определения (в формате 1опр и 2опр)'
Public Function GetBatchMDtoMultiMethod(docid As String, batchno As String, key As String, methodname As String, testname As String, opr As String) As String
    GetBatchMDtoMultiMethod = GetEI().GetCalculations().GetBatchMDtoMultiMethod(docid, batchno, key, methodname, testname, opr)
End Function
'Возвращает значение оптической плотности; аргументы: docid - номер документа, batchno - номер пробы, key - ключ, methodname - имя метода, testname - имя показателя, opr - номер определения, kuv - толщина кюветы'
Public Function GetBatcOptPtoMultiMethod(docid As String, batchno As String, key As String, methodname As String, testname As String, opr As String, kuv As String) As String
    GetBatcOptPtoMultiMethod = GetEI().GetCalculations().GetBatcOptPtoMultiMethod(docid, batchno, key, methodname, testname, opr, kuv)
End Function
'Возвращает метаданные по ключу; docbatchid - шифр документа и пробы, разделенные символом |, key - ключ'
Public Function GetDocBatchMetaData(docbatchid As String, key As String) As String
    GetDocBatchMetaData = GetEI().GetCalculations().GetDocBatchMetaData(docbatchid, key)
End Function
'преобразует дату из текстого формата в формат даты; аргументы: strdate - дата в текстовом формате; defdateval - значение, выводимое по умолчанию'
Public Function tryToDate(strdate As String, defdateval As Date) As Date
    tryToDate = GetEI().GetCalculations().tryToDate(strdate, defdateval)
End Function
'возвращает дату в текстовом формате; аргументы: str - дата в текстовом формате, format - формат даты'
Public Function tryToDateStr(str As String, format As String) As String
    tryToDateStr = GetEI().GetCalculations().tryToDateStr(str, format)
End Function
'возвращает число; аргументы: strnum - число в текстовом формате, defnnum - значение по умолчанию'
Public Function tryToNumber(strnum As String, defnumber As Double) As Double
    tryToNumber = GetEI().GetCalculations().tryToNumber(strnum, defnumber)
End Function
'возвращает объект; strnum - число в строковом представлении, с - строка, возвращаемая когда не удалось преобразовать строку в число'
Public Function trytodigit(strnum As String, C As String) As Variant
    trytodigit = GetEI().GetCalculations().trytodigit(strnum, C)
End Function
'выводит результат сравнения двух чисел; аргументы: r - число 1, rk - число 2, strformat - формат числа в строковом представлении'
Public Function nervo(r As Double, rk As Double, strformat As String) As String
    nervo = GetEI().GetCalculations().nervo(r, rk, strformat)
End Function
'округление неопределенности, погрешности по МР'
Public Function RoundNeopr(str As String) As String
    RoundNeopr = GetEI().GetCalculations().RoundNeopr(str)
End Function
'Округление неопределённости, погрешности по ГОСТ 31371.7'
Public Function RoundNeoprGOSTGas(str As String) As String
    RoundNeoprGOSTGas = GetEI().GetCalculations().RoundNeoprGOSTGas(str)
End Function
'возвращает дату и время отбора пробы; аргументы: dodcid - номер документа, batchno - номер пробы, includeTime - отображать время или нет'
Public Function ДатаВремяОтбораПробы(docid As String, batchno As String, includeTime As Boolean) As String
    ДатаВремяОтбораПробы = GetEI().GetCalculations().BatchSamplingDateTime(docid, batchno, includeTime)
End Function
'возвращает дату и время анализа пробы; batchno - номер пробы'
Public Function ДатаВремяАнализаПробы(batchno As String) As String
    ДатаВремяАнализаПробы = GetEI().GetCalculations().GetDateOfAnalysis(batchno)
End Function
'возвращает результат округленный по погрешности, если не задана погрешность, то округляет по количеству знаков после запятой: resvalue - результат, pogrvalue - погрешность, precstring - точность'
Public Function GetResultString(resvalue As String, pogrvalue As String, precstring As String) As String
    GetResultString = GetEI().GetCalculations().GetResultString(resvalue, pogrvalue, precstring)
End Function
'возвращает значение результата в числовом формате с учетом погрешности и количества знаков после запятой; resvalue - результат, pogrvalue - погрешность, precstring - точность'
Public Function GetResultDigit(resvalue As String, pogrvalue As String, precstring As String) As Variant
    GetResultDigit = GetEI().GetCalculations().GetResultDigit(resvalue, pogrvalue, precstring)
End Function
'возвращает результат в формате результат  плюс минус погрешность; resvalue - результат, pogrvalue - погрешность, precstring - количество знаков в формате 0.000'
Public Function GetResultWithPogr(resvalue As String, pogrvalue As String, precstring As String) As String
    GetResultWithPogr = GetEI().GetCalculations().GetResultWithPogr(resvalue, pogrvalue, precstring)
End Function
'возвращает номер пробы; docid - номер документа,batchno - номер пробы '
Public Function GetProbNum(docid As String, batchno As String) As String
    ProbNum = GetEI().GetCalculations().GetBatchMetaData(docid, batchno, "Номер пробы")
    If ProbNum = "" Then
        GetProbNum = GetEI().GetCalculations().GetBatchMetaData(docid, batchno, "Шифр (номер) пробы")
    Else
        GetProbNum = ProbNum
    End If
End Function
'отображает формулу; r - ячейка с вычислениями'
Public Function GetFormula(r As Range) As String
    GetFormula = GetEI().GetCalculations().GetFormuls(r)
End Function
'Проходит по Range из первого аргумента и собирает все уникальные строки из него, разделяя их вторым аргументом'
Public Function SumDistinctStringsRange(r As Range, delimetr As String) As String
Dim result As String
result = GetEI().GetCalculations().SumDistinctStringsRange(r, delimetr)
If result = "-2146826273" Then
    SumDistinctStringsRange = "—"
Else
    SumDistinctStringsRange = result
End If
End Function
'возвращает шифр документа с поправками для барометра; ZavNumber - заводской номер, ToDate - дата, до которой будет производиться поиск документа'
Public Function GetBarPopDocId(ZavNumber As String, ToDate As String, r As Range) As String
    GetBarPopDocId = GetEI().GetBarPopDocId(ZavNumber, ToDate, r)
End Function
'возвращает шифр документа с поправками для термометра; ZavNumber - заводской номер, ToDate - дата, до которой будет производиться поиск документа'
Public Function GetTempPopDocId(ZavNumber As String, ToDate As String, r As Range) As String
    GetTempPopDocId = GetEI().GetTempPopDocId(ZavNumber, ToDate, r)
End Function
'возвращает значение поправки для температуры; docid - шифр документа с поправкой для термометра, znach - значение температуры'
Public Function GetPopr(docid As String, znach As Double) As Double
    GetPopr = GetEI().GetPopr(docid, znach)
End Function
'возвращает документ с условиями окружающей среды; mesto - место измерения, lab - лаборатория, DateTo - дата, до которой искать, r - любая ячейка на листе'
Public Function GetUslOkrDOc(mesto As String, lab As String, DateTo As String, r As Range) As String
    GetUslOkrDOc = GetEI().GetUslOkrDOc(mesto, lab, DateTo, r)
End Function
'убирает шифр пробы, str - строка с шифром, disp=true, если шифр пробы для отображения, если false, то системный'
Public Function RemoveIDStr(str As String, disp As Boolean) As String
    RemoveIDStr = GetWI().RemoveIDStr(str, disp)
End Function
'извлекает шифр пробы, str - строка с шифром, disp=true, если шифр пробы для отображения, если false, то системный'
Public Function GetIDFromStr(str As String, disp As Boolean) As String
    GetIDFromStr = GetWI().GetIDFromStr(str, disp)
End Function
'возвращает шифр калибровочного графика; r - любая ячейка на листе, lab - лаборатория, kuveta1 - толщина кюветы 1, optplotn1 - значение оптической плотности 1-ой кюветы, kuveta2 - толшина кюветы 2, optplotn2 - значение оптической плотности 2-ой кюветы, prodnumber - заводской номер фотометра, samplingplace - место отбора'
Function GetCalibrGrafDocid(r As Range, lab As String, kuveta1 As String, optplotn1 As String, kuveta2 As String, optplotn2 As String, prodnumber As String, samplingplace As String, Optional test As String = "") As String
    GetCalibrGrafDocid = GetEI().GetCalculations().GetCalibrGrafDocid(r, lab, kuveta1, optplotn1, kuveta2, optplotn2, prodnumber, samplingplace, test)
End Function
'возвращает коэффициент A калибровочного графика; docid - шифр документа калибровочного графика'
Function getcoefA(docid As String) As Variant
    If InStr(1, docid, "|") <> 0 Then
        getcoefA = getBatchMetadataAsDouble(docid, docid, "Коэффициент A", 0)
    Else
        getcoefA = ""
    End If
End Function
'возвращает коэффициент B калибровочного графика; docid - шифр документа калибровочного графика'
Function getcoefB(docid As String) As Variant
    If InStr(1, docid, "|") <> 0 Then
        getcoefB = getBatchMetadataAsDouble(docid, docid, "Коэффициент B", 1)
    Else
        getcoefB = ""
    End If
End Function
'возвращает толщину кюветы  измерения оптической плотности; docid - шифр документа калибровочного графика'
Function getKuveta(docid As String) As String
    getKuveta = GetBatchMetaData(docid, docid, "Кювета")
End Function
'используется для поиска заводских номеров термометров, барометров и т.п. в области, которая заполняется при выборе СИ; r - диапазон при выборе средств измерений, what - что искать, searchcolindex - в каком столбике искать, colindex - из какого столбика возвращать'
Function GetWorkNumber(r As Range, what As String, Optional n As Long = 1, Optional searchcolindex As Long = 1, Optional colindex As Long = 2) As String
    GetWorkNumber = GetEI().GetCalculations().GetWorkNumber(r, what, n, searchcolindex, colindex)
End Function
'Возвращает что ищем из диапазона, r - диапазон, what - объект для поиска, isInstr - инструмент или нет'
Function GetSomethingInRange(r As Range, what As String, isInstr As Boolean, n As Long, searchcolindex As Long, colindex As Long) As String
    GetSomethingInRange = GetEI().GetCalculations().GetSomethingInRange(r, what, isInstr, n, searchcolindex, colindex)
End Function
'находит документы, заданные аргументами, сортирует в обратном порядке и возвращает первый из отсортированных, r - любая ячейка на листе, lab - лаборатория, docdescription - имя документа, key1 - ключ 1, value1 - значение 1, sortmetadataname - имя метаданных для сортировки, найденного в обратном порядке, DateTo - до какой даты искать'
Function GetSomeDocidBeforeDate(r As Range, lab As String, docdescription As String, key1 As String, Value1 As String, sortmetadataname As String, DateTo As String) As String
    GetSomeDocidBeforeDate = GetEI().GetSomeDocidBeforeDateForCurUser(r, lab, docdescription, key1, Value1, sortmetadataname, DateTo)
End Function
'Возвращает вместимость кислородной колбы; r - любая ячейка на листе, num - номер колбы, lab - лаборатория, DateTo - дата, до которой производится поиск'
Function GetKislKolbVmest(r As Range, num As String, lab As String, DateTo As String) As Variant
    GetKislKolbVmest = GetEI().GetKislKolbVmest(r, num, lab, DateTo)
End Function
'возвращает температурную поправку термометра; termdocid - номер документа с поправками, temp - температура'
Public Function GetTermTempPopr(termdocid As String, temp As Double) As Variant
    GetTermTempPopr = GetEI().GetPopr(termdocid, temp)
End Function
'возвращает значение барометрической поправки; docid - шифр документа с поправкой, davl - величина давления, temp - температура'
Public Function GetBarPopr(docid As String, davl As Double, temp As Double) As Double
    GetBarPopr = GetEI().GetPopr(docid, davl)
End Function
'возвращает шифр документа с поправками для термометра; ZavNumber - заводской номер термометра, r - любая ячейка на листе , lab - лаборатория, DateTo - дата, до которой осуществлять поиск документа'
Public Function GetTermPoprDocId(ZavNumber As String, r As Range, lab As String, DateTo As String) As String
    GetTermPoprDocId = GetSomePovDocId("Термометр", ZavNumber, r, lab, DateTo)
End Function
'ищет документы для поправок, what - Объект поиска, ZavNumber - заводской номер, r - диапазон для поиска, lab - лаборатория, DateTo - дата, до которой искать'
Public Function GetSomePovDocId(what As String, ZavNumber As String, r As Range, lab As String, DateTo As String) As String
   On Error Resume Next
    Application.Cursor = xlWait
    GetSomePovDocId = GetSomeDocidBeforeDate(r, lab, what, "Заводской номер", CStr(ZavNumber), "Дата поверки", DateTo)
    Application.Cursor = xlDefault
End Function
'Функция для поиска по интервалам x в таблице tabledataformethods в Oracle - вызывает соответствующую PL\SQL функцию; x - интервал, description - название таблицы, notFoundValue - значение, выводимое по умолчанию'
Function TableFind(x As Double, description As String, notFoundValue As Double) As Double
On Error GoTo errorhandler
    Dim wi As Object
    Application.Cursor = xlWait
    Set wi = GetWI
    TableFind = wi.TableFind(description, x, notFoundValue)
errorhandler:
    Application.Cursor = xlDefault
End Function
'Возвращает рассчитанное табличное значение по двум параметрам; x1 - значение параметра 1, x2 - значение параметра 2, description - название таблицы, notFoundValue - значение,выводимое по умолчанию'
Function TableCalc(x1 As Double, x2 As Double, description As String, notFoundValue As Double) As Double
    On Error GoTo errorhandler
    Dim wi As Object
    Application.Cursor = xlWait
    Set wi = GetWI
    TableCalc = wi.TableCalc(description, x1, x2, notFoundValue)
errorhandler:
    Application.Cursor = xlDefault
End Function
'По показателю и парметру возвращает значение второго параметра; x1 - значение параметра 1, Y - показатель 2, description - название таблицы, notFoundValue - значение,выводимое по умолчанию'
Function TableCalcObr(x1 As Double, Y As Double, description As String, notFoundValue As Double) As Double
On Error GoTo errorhandler
    Dim wi As Object
    Application.Cursor = xlWait
    Set wi = GetWI
    TableCalcObr = wi.TableCalcObr(description, x1, Y, notFoundValue)
errorhandler:
    Application.Cursor = xlDefault
End Function
'возвращает табличное значение плотности по температуре и давлению; description - название таблицы, t - температура, p - давление'
Public Function DensCalc(description As String, t As Double, p As Double) As Double
    DensCalc = TableCalc(t, p, description, -999)
End Function
'возвращает значение температуры; description - название таблицы, t - температура, p20 - значение плотности при 20'
Public Function DensCalcobr(description As String, t As Double, p20 As Double) As Double
    DensCalcobr = TableCalcObr(t, p20, description, -999)
End Function
'возвращает табличное значение плотности по температуре и давлению; description - название таблицы, t - температура, p - давление по методу МИ 2153 при 20 градусах'
Public Function MI2153_20(description As String, t As Double, p As Double) As Double
    MI2153_20 = TableCalc(t, p, description, -999) '"МИ 2153 пл. при 20 гр. 20 Б1"
End Function
'возвращает табличное значение плотности по температуре и давлению; description - название таблицы, t - температура, p - давление по методу МИ 2153 при 15 градусах'
Public Function MI2153_15(description As String, t As Double, p As Double) As Double
    MI2153_15 = TableCalc(t, p, description, -999) '"МИ 2153 пл. при 15 гр. 20 Б2"
End Function
'см. TableCalc'
Public Function GetCorrection(description As String, t As Double, p As Double) As Double
    GetCorrection = TableCalc(t, p, description, -999)
End Function
'возвращает поправку для метода РД 52.24.495; description - название таблицы, t - значение температуры'
Function CalcPoprRD52_24_495(description As String, t As Double) As Double
    Dim vs As Double
    vs = TableCalc(t * 10 \ 10, t - t * 10 \ 10, description, -999)
    CalcPoprRD52_24_495 = -999999
    If vs <> -1 Then
        CalcPoprRD52_24_495 = vs
    End If
End Function
'Возвращает табличное значение поправки для метода ПНД Ф 14.1:2:3:4.123; аргументы: description - название таблицы, t - значение температуры'
Function CalcPoprPNDF_123(description As String, t As Double) As Double
    Dim vs As Double

    vs = TableCalc(t * 10 \ 10, t - t * 10 \ 10, description, -1)
    CalcPoprPNDF_123 = 1
    If vs <> -1 Then
        CalcPoprPNDF_123 = vs
    End If
End Function
'см. TableCalc'
Public Function getValueA(description As String, t As Double, p As Double) As Double
    Dim wi As Object
    Set wi = GetWI
    getValueA = wi.TableCalcInterval(description, p, t, -999999)
End Function
'Вычисляет класс чистоты масел, description - название таблицы, razmer и count - параметры для расчета'
Function CalcClassChist(description As String, razmer As String, count As Double) As Variant
    Dim wi As Object
    Set wi = GetWI
    CalcClassChist = wi.TableCalcInterval(description, razmer, count, -999999)
End Function
'возвращает значение поправки для нулевой точки;termdocid - шифр документа с поправкой на термометр, zeropoint - положение нулевой точки'
Function GetTermZeroPopr(termdocid As String, zeropoint As Variant) As Double
    Dim zero As Double
    zero = getBatchMetadataAsDouble(termdocid, termdocid, "Пол-е нулевой точки", -999)
    If zero <> -999 And zeropoint <> "Не определялся" Then
        GetTermZeroPopr = zero - zeropoint
    Else
        GetTermZeroPopr = 0
    End If
End Function
'Возвращает поправку на давление; termdocid - шифр документа с температурной поправкой, davl - давление, unit - единица измерения давления'
Function GetTermAtmPopr(termdocid As String, davl As Double, unit As String) As Double
    Dim Be As Double
    Dim unitterm As String
    Be = getBatchMetadataAsDouble(termdocid, termdocid, "Be", -999)
    unitterm = unit
    If Be <> -999 Then
        If unit <> "кПа" Then
            If unitterm <> "кПа" Then
                GetTermAtmPopr = Be * (760 - davl)
            Else
                GetTermAtmPopr = Be * (101.325 - davl / 0.00750064 / 1000)
            End If
        Else
            If unitterm <> "кПа" Then
                GetTermAtmPopr = Be * (760 - davl * 0.00750064 * 1000)
            Else
                GetTermAtmPopr = Be * (101.325 - davl)
            End If
        End If
    Else
        GetTermAtmPopr = 0
    End If
End Function
'Возвращает поправку на давление для метода 2177, t - значение температуры, P - значение давления, isformula -  isformula, то считает по формуле из ГОСТ 2177, иначе  по таблицам интерполяцией, unit - единица измерения'
Function CalcPopr2177(t As Double, p As Double, isFormula As Boolean, unit As String)
    Dim vs As Double
    If isFormula Then
        If unit <> "кПа" Then
            CalcPopr2177 = 0.00012 * (760 - p) * (273 + t)
        Else
            CalcPopr2177 = 0.00009 * (101.3 - p) * (273 + t)
        End If
        Exit Function
    End If
    If p < 750 Or p > 770 Then
        vs = TableFind(t, "ГОСТ 2177 Поправки", -1)
        CalcPopr2177 = -999999
        If vs <> -1 Then
            ' неправильный вариант. правильно отнимать от 760, но все просят делать так
            If p < 750 Then
                CalcPopr2177 = Round((750 - p) / 10, 2) * vs
            End If
            If p > 770 Then
                CalcPopr2177 = Round((770 - p) / 10, 2) * vs
            End If
        End If
    Else
        CalcPopr = 0
    End If
End Function
'возвращает значение поправки по методу ГОСТ 21534'
Public Function Calc21534(v As Range, vx As Double, t As Double, a As Double, vpr As Double) As Double
    Dim C As Long
    Dim sumv As Double
    Dim count As Long
    sumv = 0
    count = 0
    For C = 1 To v.Columns.count
        If v.Cells(1, C).value <> 0 Then
            sumv = sumv + v.Cells(1, C).value
            count = count + 1
        Else
            Exit For
        End If
    Next C
    Calc21534 = (sumv - count * vx) * t * 1000 * a / vpr
End Function
'Возвращает строку с наибольшей длиной, strings - диапазон строк'
Public Function MaxStr(strings As Range) As String
    MaxStr = GetEI().GetCalculations().MaxLenStr(strings)
End Function
'возвращает интервал времени, r - диапазон значений'
Public Function GetDateInterval(r As Range) As String
    GetDateInterval = GetEI().GetCalculations().GetDateInterval(r)
End Function
'возвращает номер документа по трем ключам; r - любая ячейка на листе, lab - лаборатория, docdescription - описание документа, key1 - ключ 1, Value1 - значение ключа 1, key2 - ключ 2, Value2 - значение ключа 2, key3 - ключ 3, Value3 - значение ключа 3, sortmetadataname - имя метаданных для сортировки, найденного в обратном порядке, DateTo - дата, до которой производить поиск'
Function GetSomeDocidBy3Keys(r As Range, lab As String, docdescription As String, key1 As String, Value1 As String, key2 As String, Value2 As String, key3 As String, Value3 As String, sortmetadataname As String, DateTo As String)
    GetSomeDocidBy3Keys = GetEI().GetCalculations().GetSomeDocidBy3Keys(r, lab, docdescription, key1, Value1, key2, Value2, key3, Value3, sortmetadataname, DateTo)
End Function
'возвращает условия окружающей среды, r - '
Public Function GetUslByIntervals(r As Range) As String
    GetUslByIntervals = GetEI().GetCalculations().GetUslByIntervals(r)
End Function
'находит минимальную и максимальныую даты и возвращает их в виде строки мин - макс, если разница между ними меньше заданного в минутах интервала, то возвращает мин - мин+интервал, r - диапазон, minutes - интервал'
Public Function GetDateIntervalMinMinutes(r As Range, minutes As Double) As String
    GetDateIntervalMinMinutes = GetEI().GetCalculations().GetDateIntervalMinMinutes(r, minutes)
End Function
'возвращает норматив спецификации, batchno - номер пробы, orderatrtibute - параметр для сортировки'
Public Function GetSpecNorm(batchno As String, orderattribute As String) As String
    GetSpecNorm = GetEI().GetCalculations().GetSpecNorm(batchno, orderattribute)
End Function
'возвращает норму для расчета из спецификации, batchno - номер пробы, orderatrtibute - параметр для сортировки'
Public Function GetSpecRaschNorm(batchno As String, orderattribute As String) As String
    GetSpecRaschNorm = GetEI().GetCalculations().GetSpecRaschNorm(batchno, orderattribute)
End Function
'возвращает количество знаков из спецификации, batchno - номер пробы, orderattribute - показатель'
Public Function GetSpecNumPrec(batchno As String, orderattribute As String) As String
    GetSpecNumPrec = GetEI().GetCalculations().GetSpecNumPrec(batchno, orderattribute)
End Function
'возвращает номер спецификации для пробы, batchno - номер пробы, orderattribute - показатель'
Public Function GetSpecNo(batchno As String, orderattribute As String) As String
    GetSpecNo = GetEI().GetCalculations().GetSpecNo(batchno, orderattribute)
End Function
'Расчет концентрации, docid - шифр документа с данными для расчета, x - показатель'
Public Function CalcConcentration(docid As String, x As Double) As Double
    CalcConcentration = GetEI().GetCalculations().CalcConcentration(docid, x)
End Function
'Заменяет цифры в строке на строковое значение; str - строка, repl - строковое значение'
Public Function ReplaceDigitsTo(str As String, repl As String) As String
    ReplaceDigitsTo = GetEI().GetCalculations().ReplaceDigitsTo(str, repl)
End Function
'Создает строку формата по числу; str - число в строковом выражении'
Public Function GetNumFormatByNumber(str As String) As String
    GetNumFormatByNumber = GetEI().GetCalculations().GetNumFormatByNumber(str)
End Function
'Для получения данных по взвешиваниям в рабочих журналах по гравиметрическим методам; docid - необязательный аргумент, batchno - шифр пробы, key - часть ключа без номера взвешивания, methodname - метод, testname - показатель, opr - номер определения (1опр для первого), n - номер взвешивания с конца (1 для последнего, 2 для предпоследнего и т.д.)'
Public Function GetBatcMasBSMultiMethod(docid As String, batchno As String, key As String, methodname As String, testname As String, opr As String, n As String) As Variant
    Set ei = GetEI
    GetBatcMasBSMultiMethod = ei.GetCalculations().GetBatcMasBSMultiMethod(docid, batchno, key, methodname, testname, opr, n)
End Function
'Возвращает требуемое значение из докуммента; r - , docdescription - описание документа, key1 - ключ 1, value 1 - значение 1, whattoget - ключ, sortmetadata - имя метаданных по значениям которых будет произведена сортировка найденных шифров в обратном порядке (чаще всего Дата регистрации), DateTo - дата, до которой осуществлять поиск'
Public Function GetSomething(r As Range, lab As String, docdescription As String, key1 As String, Value1 As String, whattoget As String, sortmetadataname As String, DateTo As String) As Variant
    GetSomething = GetEI().GetCalculations().GetSomething(r, lab, docdescription, key1, Value1, whattoget, sortmetadataname, DateTo)
End Function
'Возвращает вмеситмость пикнометра; r - , num - номер пикнометра, lab - лаборатория, DateTo - дата, до которой осуществлять поиск'
Function GetPiknVmest(r As Range, num As String, lab As String, DateTo As String) As Variant
On Error Resume Next
    GetPiknVmest = GetSomething(r, lab, "Определение вместимости пикнометра", "Номер пикнометра", num, "Вместимость", "Дата составления", DateTo)
End Function
'Возвращает поправку на гири; docid -номер документа с поправками на гири, nomrange - диапазон с номиналами'
Function GetGiriPopr(docid As String, nomrange As Range) As Double
    GetGiriPopr = GetEI().GetCalculations().GetGiriPopr(docid, nomrange)
End Function
'Расчет точки росы ГОСТ 20060
Function TRossCalc(description As String, x1 As Double, x2 As Double, defaultval As Double) As Double
    TRossCalc = GetWI().TRossCalc(description, x1, x2, defaultval)
End Function
'Расчет точки росы'
Function TRosCalc(x As Double, xr As Range, yr As Range) As Double
    TRosCalc = GetEI().GetCalculations().TRosCalc(x, xr, yr)
End Function
'Расчёт точки росы по ГОСТ 53763 по таблицам в Oracle; x1 - среднее по двум определениям давление газа в приборе, x2 - температура точки росы по влаге, description - "ГОСТ 53763" таблица ТТР", notfoundvalue - значение, выводимое когда ничего не найдено'
Function TRos40Calc(x1 As Double, x2 As Double, description As String, notFoundValue As Double) As Double
    On Error GoTo errorhandler
    Application.Cursor = xlWait
    TRos40Calc = GetWI().TRos40CalcStr(description, x1, x2, notFoundValue)
errorhandler:
    Application.Cursor = xlDefault
End Function
'Строковый расчёт точки росы по ГОСТ 53763 по таблицам в Oracle; x1 - среднее по двум определениям давление газа в приборе, x2 - температура точки росы по влаге, description - "ГОСТ 53763" таблица ТТР", notfoundvalue - значение, выводимое когда ничего не найдено'
Function TRos40CalcStr(x1 As Double, x2 As Double, description As String, notFoundValue As Double) As Double
    On Error GoTo errorhandler
    Application.Cursor = xlWait
    TRos40CalcStr = GetWI().TRos40CalcStr(description, x1, x2, notFoundValue)
errorhandler:
    Application.Cursor = xlDefault
End Function
'возвращает средства измерения по пробе в формате описание оборудования (заводской номер); batchno - номер пробы, methodname - имя метода, testname - имя показателя, instrtype - тип интрумента, instrdelimeter - разделитель'
Function GetBatchInstruments(batchno As String, methodname As String, testname As String, instrtype As String, instrdelimeter As String) As String
    GetBatchInstruments = GetEI().GetCalculations.GetBatchInstruments(batchno, methodname, testname, instrtype, instrdelimeter)
End Function
'возвращает информацию по используемым инструментам; batchno - номер пробы, orderattribute - код основного показателя в формате ГОСТ_показатель, instrtype - тип инструмента, instrdelimeter - разделитель'
Function GetBatchInstrumentsByOrderattr(batchno As String, orderattribute As String, instrtype As String, instrdelimeter As String) As String
    GetBatchInstrumentsByOrderattr = GetEI().GetCalculations.GetBatchInstrumentsByOrderattr(batchno, orderattribute, instrtype, instrdelimeter)
End Function
'возвращает информацию по инструментам в пробе, batchno - шифр пробы, methoddescr - описание метода, testmeta - ид основного показателя, instrtype - тип оборудования, instrdelimeter - разделитель оборудования при нескольких, formatstr - формат вывода (пр: {Description}{ProdNumber}, в фигурных скобках могут быть любые поля таблицы Instruments))'
Function GetBatchInstrumentsInfo(batchno As String, methoddescr As String, testmeta As String, instrtype As String, instrdelimeter As String, formatstr As String) As String
    GetBatchInstrumentsInfo = GetEI().GetCalculations.GetBatchInstrumentsInfo(batchno, methoddescr, testmeta, instrtype, instrdelimeter, formatstr)
End Function
'Возвращает количество определений в пробе; batchno - номер пробы, orderattribute - показатель'
Function GetEnteredOpredCount(batchno As String, orderattribute As String) As String
    GetEnteredOpredCount = GetEI().GetCalculations.GetEnteredOpredCount(batchno, orderattribute)
End Function
'возвращает информацию по инструментам в пробе, batchno - шифр пробы, orderattribute - код основного показателя в формате ГОСТ_показатель, instrtype - тип оборудования, instrdelimeter - разделитель оборудования при нескольких, formatstr - формат вывода (пр: {Description}{ProdNumber}, в фигурных скобках могут быть любые поля таблицы Instruments))'
Function GetBatchInstrumentsInfoByOrderattr(batchno As String, orderattribute As String, instrtype As String, instrdelimeter As String, formatstr As String) As String
GetBatchInstrumentsInfoByOrderattr = GetEI().GetCalculations.GetBatchInstrumentsInfoByOrderattr(batchno, orderattribute, instrtype, instrdelimeter, formatstr)
End Function
'преобразует число из числового формата в текстовый в родительном падеже; number - число'
Function ConvertNumberToTextGenitive(number As Long)
    ConvertNumberToTextGenitive = GetEI().GetCalculations.ConvertNumberToTextGenitive(number)
End Function
'преобразует число в текст в иминительном падеже, number - число'
Function ConvertNumberToTextNominative(number As Long)
    ConvertNumberToTextNominative = GetEI().GetCalculations.ConvertNumberToTextNominative(number)
End Function
'просматривает область и если предыдущее значение отличается от текущего не более чем на Delta, возвращает его; r - область значений'
Function GetUnchanged(r As Range, Optional Delta As Double = 0) As Variant
    GetUnchanged = GetEI().GetUnchanged(r, Delta)
End Function
'возвращает среднее из диапазона; r - диапазон'
Function Getchcount(r As Range) As Double
    Getchcount = GetEI().GetCalculations.Getchcount(r)
End Function
'Расчет минимальной концентрации, docid - шифр документа с данными для расчета,x - показатель'
Function CalcMinConcentration(docid As String, x As Double) As Double
    CalcMinConcentration = GetEI().GetCalculations.CalcMinConcentration(docid, x)
End Function
'Расчет максимальной концентрации, docid - шифр документа с данными для расчета,x - показатель'
Function CalcMaxConcentration(docid As String, x As Double) As Double
    CalcMaxConcentration = GetEI().GetCalculations.CalcMaxConcentration(docid, x)
End Function
'отсутствует в c#'
Function GetMetaValueFromSheet(testname As String, methodname As String) As Double
    GetMetaValueFromSheet = GetEI().GetCalculations.GetMetaValueFromSheet(testname, methodname)
End Function
'Возвращает все показатели (названия), привязанные к методу с типом "результат"
Function GetMethodsTests(methodsids As String) As String
    GetMethodsTests = GetEI().GetCalculations.GetMethodsTests(methodsids)
End Function
'Если на входе строка пустая возвращает пробел - для отчётов, в которых не нужно скрывать столбики
Public Function GetSpaceInsteadEmpty(str As String) As String
    GetSpaceInsteadEmpty = str
    If str = "" Then
        GetSpaceInsteadEmpty = " "
    End If
End Function
'Возвращает все показатели (краткое описание), привязанные к методу с типом "результат"
Function GetMethodsTestsShortDesc(methodsids As String) As String
    GetMethodsTestsShortDesc = GetEI().GetCalculations.GetMethodsTestsShortDesc(methodsids)
End Function
'Возвращает значение Цели контроля по ID, purposeid - id цели контроля
Function GetPurposeDescriptionById(purposeid As String) As String
        GetPurposeDescriptionById = GetEI().GetCalculations.GetPurposeDescriptionById(purposeid)
End Function
'возвращает информацию по инструментам в пробе, batchno - шифр пробы, methoddescr - описание метода (не обязательно), testmeta - ид основного показателя (не обязательно), instrtype - тип оборудования (не обязательно), instrdelimeter - разделитель оборудования при нескольких, formatstr - формат вывода (пр: {Description}{ProdNumber}, в фигурных скобках могут быть любые поля таблицы Instruments))'
Function GetBatchInstrumentsInfo2(batchno As String, methoddescr As String, testmeta As String, instrtype As String, instrdelimeter As String, formatstr As String) As String
    GetBatchInstrumentsInfo2 = GetEI().GetCalculations.GetBatchInstrumentsInfo2(batchno, methoddescr, testmeta, instrtype, instrdelimeter, formatstr)
End Function
'Ищет максимальнок количество знаков из диапазона
Function FindMaxCharsInRange(r As Range) As Integer
Dim C As Range
Dim i, count As Integer
count = 0
For Each C In r
    If CStr(C.value) <> "" Then
        i = IIf(Int(C.value) = C.value, 0, Len(Split(C.value, Mid(1 / 2, 2, 1))(UBound(Split(C.value, Mid(1 / 2, 2, 1))))))
            If i > count Then count = i
            End If
Next
FindMaxCharsInRange = count
End Function
'Дата словами
Public Function ДатаСловами(d As String)
    Dim vs As String
    Dim year As String
    Dim vsd As Date
    vsd = CDate(d)
    vs = format(vsd, "Long Date")
    year = format(vsd, "yyyy")
    ДатаСловами = vs
    vs = Replace(vs, " " & year & " г.", "")
    If (Mid(vs, Len(vs), 1) = "?" Or Mid(vs, Len(vs), 1) = "ь" Or Mid(vs, Len(vs), 1) = "й") Then
        vs = Mid(vs, 1, Len(vs) - 1) & "я"
    Else
        vs = vs & "а"
    End If
    ДатаСловами = """" & Replace(LCase(vs), " ", """ ") & " " & LCase(year) & " г."
End Function
