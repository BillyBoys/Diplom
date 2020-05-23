VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportBIQTasks 
   Caption         =   "Перенос BIQ задач "
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11925
   OleObjectBlob   =   "ImportBIQTasks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportBIQTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









'==========================================================================='
'Скрипт импорта задач из оценки'
'==========================================================================='

'Названия полей в MS Project
Dim projectField_Name         As Long
Dim projectField_JirID        As Long
Dim projectField_Cost         As Long
Dim projectField_Actor        As Long
Dim projectField_DurationDays As Long
Dim projectField_Restrict     As Long
Dim projectField_JiraProjName As Long
Dim projectField_Predecessors As Long
Dim projectField_Start        As Long
Dim projectField_End          As Long
Dim projectField_ITService    As Long
Dim projectField_TypeWork     As Long
Dim projectField_Teg          As Long
Dim projectField_ResGroup     As Long
Dim projectField_ResGroupCk   As Long
Dim projectField_FuncArea1    As Long
Dim projectField_FuncArea2    As Long
Dim projectField_FuncArea3    As Long
Dim projectField_System1      As Long
Dim projectField_System2      As Long
Dim projectField_ImpDate      As Long
Dim projectField_EmpImpTask   As Long

'Кнопка создания листа справочника для Project
Private Sub CreateManual_Click()
  
  If Len(Trim(FileNameManTextBox.Text)) <> 0 Then
    PathToExc = FileNameManTextBox.Text
    Set xlobject = CreateObject("Excel.Application")
    xlobject.Workbooks.Open PathToExc
    If xlobject.ActiveWorkbook Is Nothing Then
      xlobject.Quit 'Закрытие Excel файла
      MsgBox "Некоректный путь к экспресс оценке задачи"
      Exit Sub
    End If
    xlobject.DisplayAlerts = False
    'Создание нового листа
    xlobject.ActiveWorkbook.Sheets.Add.name = "Справочник для Project"
    xlobject.ActiveWorkbook.Sheets(1).Move After:=xlobject.ActiveWorkbook.Sheets(4)
    Set ExcelSheet = xlobject.ActiveWorkbook.Sheets(4)
    'Запись в лист
    ExcelSheet.Cells(1, 1).Value = "Группа ЦК"
    ExcelSheet.Cells(2, 1).Value = "Функциональная область"
    ExcelSheet.Cells(3, 1).Value = "Тег"
    ExcelSheet.Cells(1, 2).Formula = "=Оценка!C6"
    ExcelSheet.Cells(1, 2).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(2, 2).Formula = "=Оценка!C7"
    ExcelSheet.Cells(2, 2).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(3, 2).Value = " "
    ExcelSheet.Cells(3, 2).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(1, 3).Formula = "=Оценка!C1"
    ExcelSheet.Cells(1, 3).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(2, 3).Formula = "=Оценка!C2"
    ExcelSheet.Cells(2, 3).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(2, 4).Value = "JIRACFT"
    ExcelSheet.Cells(2, 4).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(2, 5).Value = "25"
    ExcelSheet.Cells(2, 5).Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Cells(6, 1).Value = "Таблица задач"
    ExcelSheet.Cells(7, 1).Value = "Номер задачи"
    'Формат ячеек
    ExcelSheet.Range("A7:I24").Borders.LineStyle = True
    ExcelSheet.Cells(1, 3).WrapText = True
    For i = 1 To 17
      ExcelSheet.Cells(7 + i, 1).Value = i - 1
    Next i
    ExcelSheet.Range("A7:I24").Interior.Color = RGB(220, 230, 241)
    ExcelSheet.Range("A7:I7").Interior.Color = RGB(79, 129, 189)
    ExcelSheet.Columns(1).ColumnWidth = 17
    ExcelSheet.Columns(2).ColumnWidth = 60
    ExcelSheet.Columns(3).ColumnWidth = 36
    ExcelSheet.Columns(4).ColumnWidth = 20
    ExcelSheet.Columns(5).ColumnWidth = 15
    ExcelSheet.Columns(6).ColumnWidth = 15
    ExcelSheet.Columns(7).ColumnWidth = 20
    ExcelSheet.Columns(8).ColumnWidth = 20
    ExcelSheet.Columns(9).ColumnWidth = 15
    ExcelSheet.Columns(11).ColumnWidth = 20
    ExcelSheet.Columns(12).ColumnWidth = 15
    ExcelSheet.Columns(13).ColumnWidth = 15
    'Столбец Наименование работы в оценке
    ExcelSheet.Cells(7, 2).Value = "Наименование работы в оценке"
    ExcelSheet.Cells(8, 2).Value = "Поддержка написания и согласование (БТ, АВ, ТВИС и т.п.),  в.т.ч. регист-я запросов, проведение оценки"
    ExcelSheet.Cells(9, 2).Value = "Разработка ТВИСа (в т.ч. Xsd-схем)"
    ExcelSheet.Cells(10, 2).Value = "Согласование ТВИС"
    ExcelSheet.Cells(11, 2).Value = "Разработка ФТ (+ xsd при необходимости)"
    ExcelSheet.Cells(12, 2).Value = "Согласование ФТ"
    ExcelSheet.Cells(13, 2).Value = "Поддержка написания спецификации (СТП)"
    ExcelSheet.Cells(14, 2).Value = "Согласование спецификации"
    ExcelSheet.Cells(15, 2).Value = "Поддержка этапа разработки(подрядчик + собственными силами)"
    ExcelSheet.Cells(16, 2).Value = "Поддержка ПСИ (в т.ч. в случае проведения ПСИ)"
    ExcelSheet.Cells(17, 2).Value = "Управление задачей, Согласование ФТ на смежные системы"
    ExcelSheet.Cells(18, 2).Value = "Тиражирование"
    ExcelSheet.Cells(19, 2).Value = "Поддержка до окончания ПСИ (исправление ошибок, консультации, сборка пакета на пром, "
    ExcelSheet.Cells(20, 2).Value = "Поддержка написания и согласование документации(ТВИС,ФТ)"
    ExcelSheet.Cells(21, 2).Value = "Разработка (включает в себя разработку и сборку первого тестового пакета)"
    ExcelSheet.Cells(22, 2).Value = "Поддержка предварительного тестирования (в т.ч. интеграционного тестирования)"
    ExcelSheet.Cells(23, 2).Value = "Оценка тестировщика"
    ExcelSheet.Cells(24, 2).Value = "Оценка подрядчика"
    'Столбец Наименование работы в MS Project
    ExcelSheet.Cells(7, 3).Value = "Наименование работы в MS Project"
    ExcelSheet.Cells(8, 3).Value = "Поддержка написания и согласование"
    ExcelSheet.Cells(9, 3).Value = "Разработка ТВИСа"
    ExcelSheet.Cells(10, 3).Value = "Согласование ТВИС"
    ExcelSheet.Cells(11, 3).Value = "Разработка ФТ"
    ExcelSheet.Cells(12, 3).Value = "Согласование ФТ"
    ExcelSheet.Cells(13, 3).Value = "Поддержка написания спецификации"
    ExcelSheet.Cells(14, 3).Value = "Согласование спецификации"
    ExcelSheet.Cells(15, 3).Value = "Поддержка этапа разработки"
    ExcelSheet.Cells(16, 3).Value = "Поддержка ПСИ"
    ExcelSheet.Cells(17, 3).Value = "Управление задачей"
    ExcelSheet.Cells(18, 3).Value = "Тиражирование"
    ExcelSheet.Cells(19, 3).Value = "Поддержка до окончания ПСИ"
    ExcelSheet.Cells(20, 3).Value = "Поддержка написания и согласование документации"
    ExcelSheet.Cells(21, 3).Value = "Собственная разработка"
    ExcelSheet.Cells(22, 3).Value = "Поддержка предварительного тестирования"
    ExcelSheet.Cells(23, 3).Value = "Оценка тестировщика"
    ExcelSheet.Cells(24, 3).Value = "Оценка подрядчика"
    'Столбец Предшественники
    ExcelSheet.Cells(7, 4).Value = "Предшественники"
    ExcelSheet.Cells(8, 4).Value = ""
    ExcelSheet.Cells(9, 4).Value = "0"
    ExcelSheet.Cells(10, 4).Value = "1"
    ExcelSheet.Cells(11, 4).Value = "2"
    ExcelSheet.Cells(12, 4).Value = "3"
    ExcelSheet.Cells(13, 4).Value = "4"
    ExcelSheet.Cells(14, 4).Value = "5"
    ExcelSheet.Cells(15, 4).Value = "6;14#НО"
    ExcelSheet.Cells(16, 4).Value = "14"
    ExcelSheet.Cells(17, 4).Value = "'10#ОО;0#НН"
    ExcelSheet.Cells(18, 4).Value = "8"
    ExcelSheet.Cells(19, 4).Value = "13;10#ОО"
    ExcelSheet.Cells(20, 4).Value = "6#ОО;0#НН"
    ExcelSheet.Cells(21, 4).Value = "6"
    ExcelSheet.Cells(22, 4).Value = "13"
    ExcelSheet.Cells(23, 4).Value = "13;16"
    ExcelSheet.Cells(24, 4).Value = "6"
    'Столбец Тип работы
    ExcelSheet.Cells(7, 5).Value = "Тип работы"
    ExcelSheet.Cells(8, 5).Value = "495"
    ExcelSheet.Cells(9, 5).Value = "496"
    ExcelSheet.Cells(10, 5).Value = "497"
    ExcelSheet.Cells(11, 5).Value = "498"
    ExcelSheet.Cells(12, 5).Value = "499"
    ExcelSheet.Cells(13, 5).Value = "500"
    ExcelSheet.Cells(14, 5).Value = "501"
    ExcelSheet.Cells(15, 5).Value = "502"
    ExcelSheet.Cells(16, 5).Value = "504"
    ExcelSheet.Cells(17, 5).Value = "505"
    ExcelSheet.Cells(18, 5).Value = "506"
    ExcelSheet.Cells(19, 5).Value = "510"
    ExcelSheet.Cells(20, 5).Value = "508"
    ExcelSheet.Cells(21, 5).Value = "509"
    ExcelSheet.Cells(22, 5).Value = "503"
    ExcelSheet.Cells(23, 5).Value = "512"
    ExcelSheet.Cells(24, 5).Value = "511"
    'Столбец Исполнитель
    ExcelSheet.Cells(7, 6).Value = "Исполнитель"
    ExcelSheet.Cells(8, 6).Value = "Аналитик1[50 %]"
    ExcelSheet.Cells(9, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(10, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(11, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(12, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(13, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(14, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(15, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(16, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(17, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(18, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(19, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(20, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(21, 6).Value = "Аналитик1[50 %]"
    ExcelSheet.Cells(22, 6).Value = "Аналитик1[20 %]"
    ExcelSheet.Cells(23, 6).Value = "Аналитик1[50 %]"
    ExcelSheet.Cells(24, 6).Value = "Аналитик1[50 %]"
    'Столбец Часы
    ExcelSheet.Cells(7, 7).Value = "Часы"
    ExcelSheet.Cells(8, 7).Formula = "=Оценка!D11"
    ExcelSheet.Cells(9, 7).Formula = "=Оценка!D12"
    ExcelSheet.Cells(10, 7).Formula = "=Оценка!D13"
    ExcelSheet.Cells(11, 7).Formula = "=Оценка!D14"
    ExcelSheet.Cells(12, 7).Formula = "=Оценка!D15"
    ExcelSheet.Cells(13, 7).Formula = "=Оценка!D16"
    ExcelSheet.Cells(14, 7).Formula = "=Оценка!D17"
    ExcelSheet.Cells(15, 7).Formula = "=Оценка!D18"
    ExcelSheet.Cells(16, 7).Formula = "=Оценка!D20"
    ExcelSheet.Cells(17, 7).Formula = "=Оценка!D21"
    ExcelSheet.Cells(18, 7).Formula = "=Оценка!D22"
    ExcelSheet.Cells(19, 7).Formula = "=Оценка!D26"
    ExcelSheet.Cells(20, 7).Formula = "=Оценка!D24"
    ExcelSheet.Cells(21, 7).Formula = "=Оценка!D25"
    ExcelSheet.Cells(22, 7).Formula = "=Оценка!D19"
    ExcelSheet.Cells(23, 7).Value = "30"
    ExcelSheet.Cells(24, 7).Formula = "=Оценка!D29"
    'Столбец Проценты
    ExcelSheet.Cells(7, 8).Value = "Проценты"
    ExcelSheet.Cells(8, 8).Value = "50"
    ExcelSheet.Cells(9, 8).Value = "20"
    ExcelSheet.Cells(10, 8).Value = "20"
    ExcelSheet.Cells(11, 8).Value = "20"
    ExcelSheet.Cells(12, 8).Value = "20"
    ExcelSheet.Cells(13, 8).Value = "20"
    ExcelSheet.Cells(14, 8).Value = "20"
    ExcelSheet.Cells(15, 8).Value = "20"
    ExcelSheet.Cells(16, 8).Value = "20"
    ExcelSheet.Cells(17, 8).Value = "20"
    ExcelSheet.Cells(18, 8).Value = "20"
    ExcelSheet.Cells(19, 8).Value = "20"
    ExcelSheet.Cells(20, 8).Value = "20"
    ExcelSheet.Cells(21, 8).Value = "50"
    ExcelSheet.Cells(22, 8).Value = "20"
    ExcelSheet.Cells(23, 8).Value = "50"
    ExcelSheet.Cells(24, 8).Value = "50"
    'Столбец Исполнитель
    ExcelSheet.Cells(7, 9).Value = "Исполнитель"
    ExcelSheet.Cells(8, 9).Value = "Аналитик1"
    ExcelSheet.Cells(9, 9).Value = "Аналитик1"
    ExcelSheet.Cells(10, 9).Value = "Аналитик1"
    ExcelSheet.Cells(11, 9).Value = "Аналитик1"
    ExcelSheet.Cells(12, 9).Value = "Аналитик1"
    ExcelSheet.Cells(13, 9).Value = "Аналитик1"
    ExcelSheet.Cells(14, 9).Value = "Аналитик1"
    ExcelSheet.Cells(15, 9).Value = "Аналитик1"
    ExcelSheet.Cells(16, 9).Value = "Аналитик1"
    ExcelSheet.Cells(17, 9).Value = "Аналитик1"
    ExcelSheet.Cells(18, 9).Value = "Аналитик1"
    ExcelSheet.Cells(19, 9).Value = "Аналитик1"
    ExcelSheet.Cells(20, 9).Value = "Разработчик1"
    ExcelSheet.Cells(21, 9).Value = "Разработчик1"
    ExcelSheet.Cells(22, 9).Value = "Разработчик1"
    ExcelSheet.Cells(23, 9).Value = "Тестировщик1"
    ExcelSheet.Cells(24, 9).Value = "Подрядчик"
    'Таблица скоринга
    ExcelSheet.Cells(6, 11).Value = "Таблица скоринга"
    ExcelSheet.Cells(7, 11).Value = "Группа ЦК"
    ExcelSheet.Cells(7, 12).Value = "Функциональная область"
    ExcelSheet.Cells(7, 13).Value = "Тег"
    ExcelSheet.Cells(8, 11).Value = "20"
    ExcelSheet.Cells(8, 12).Value = "20"
    ExcelSheet.Cells(8, 13).Value = "20"

    xlobject.ActiveWorkbook.Save
    xlobject.DisplayAlerts = True
    xlobject.ActiveWorkbook.Close True
    xlobject.Quit 'Закрытие Excel файла
  Else
    MsgBox "Введите путь к экспресс оценке задачи"
  End If
 'If MsgBox("Открыть Экспресс оценку?", vbYesNo, "Открытие") = vbYes Then
 '  PathToExc = FileNameManTextBox.Text
 '  Set xlobject = CreateObject("Excel.Application")
 '  xlobject.Workbooks.Open PathToExc
 '  xlobject.Visible= True 'Закрытие Excel файла
 'End If
 FileNameCFTTextBox = FileNameManTextBox
End Sub

' Кнопка импортировать
Private Sub ImportButton_Click()
  TimeForSet = Timer
  'Запись времени в текстовик
  Call SetTimeForTxt(0, "Начало импорта ", True, False)
  ' Создаем задачи по оценке ЦФТ
  If Len(Trim(FileNameCFTTextBox.Text)) <> 0 Then
    If CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameCFTTextBox.Text) = False Then
      MsgBox "Задача с такой система уже была создана"
      Exit Sub
    End If
  End If

'  ' Создаем задачи по оценки БИСквит
'  If Len(Trim(FileNameManTextBox.Text)) <> 0 Then
'    If CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameManTextBox.Text) = False Then
'      MsgBox "Задача с такой система уже была создана"
'      Exit Sub
'    End If
'  End If
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "Конец импорта ", False, True)
  
  'Запись протокола работы
  Call SetProtocolJob("Импорт")
  
End Sub

'Запись протокола работы
Sub SetProtocolJob(CallFunc)
  Dim FileText As Integer
  'Получаем свободный номер для открываемого файла
  FileText = FreeFile
  'Открываем (или создаем) файл для чтения и записи
  Open ThisProject.Path & "\ProtocolJob.txt" For Append As FileText
  Print #FileText, CallFunc & " " & TBNumBIQ
  'Закрываем файл
  Close FileText

End Sub

'Запись времени
Sub SetTimeForTxt(TimeForSet As Single, CallFunc As String, FirstEntry, LastEntry)
  Dim FileText As Integer, ObjForOpen As Object
  'Получаем свободный номер для открываемого файла
  FileText = FreeFile
  'Открываем (или создаем) файл для перезаписи или дозаписи
  If FirstEntry = True Then
    Open ThisProject.Path & "\LogTime.txt" For Output As FileText
    Print #FileText, CallFunc
  Else
    Open ThisProject.Path & "\LogTime.txt" For Append As FileText
    Print #FileText, CallFunc & TimeForSet
  End If
  'Закрываем файл
  Close FileText
'  'Открываем файл для просмотра
'  If LastEntry = True Then
'    Set ObjForOpen = CreateObject("WScript.Shell")
'    ObjForOpen.Run ThisProject.Path & "\LogTime.txt"
'    Set ObjForOpen = Nothing
'  End If
  
End Sub

' Инициализация полей
Private Sub UserForm_Initialize()
  tbStartDate = Format(Date, "dd/mm/yyyy")
  tbFieldTest = "Оценка тестировщика"
  tbFieldPodr = "Разработка"
  tbFieldHoursTest = 10
  tbFieldHoursPodr = 20
  TBNumBIQ = "BIQ-5257"
  FileNameCFTTextBox = "C:\Users\Эрнест\Documents\GitHub\Diplom\test\Расшифровка ЭО BIQ5257.xlsx"
  'FileNameCFTTextBox = "d:\info\Эрнест\Diplom\test\Расшифровка ЭО BIQ5257.xlsx"
  FileNameManTextBox = "C:\Users\Эрнест\Documents\GitHub\Diplom\test\Расшифровка ЭО BIQ5257(1).xlsx"
  TBNumBIQFDelete = 5257
  
End Sub

' Создание задач по оценке
Public Function CreateTasksByExcel(NumBIQ, StartDate, ExcelFileName) As Boolean
  'Начинается отсчет времени функции
  TimeForSet = Timer
  Dim BiqTask As Task ' Для поиска задачи по BIQ
  'Получаем название полей в MS Project
  InitFieldConst
  'Открываем оценку для чтения этапов разработки
  PathToExc = ExcelFileName
  Set xlobject = CreateObject("Excel.Application")
  xlobject.Workbooks.Open PathToExc

  'Если не удалось открыть, то выходим
  If xlobject.ActiveWorkbook Is Nothing Then
    xlobject.Quit 'Закрытие Excel файла
    Exit Function
  End If
  
  'Интересует 4 лист оценки - технический лист для данного функционала
  Set ExcelSheet = xlobject.ActiveWorkbook.Sheets(4)
  
  ' Получаем данные о задаче
  BIQName = ExcelSheet.Cells(1, 3)    'Название BIQ
  SystemCode = ExcelSheet.Cells(2, 3) 'Система
  TaskType = ExcelSheet.Cells(2, 4)   'Оцениваемая система ЦФТ
  ITService = ExcelSheet.Cells(2, 5)  'ИТ-Сервис
  TaskGroupCK = ExcelSheet.Cells(1, 2) 'Группа ЦК
  FuncArea = ExcelSheet.Cells(2, 2)   'Функциональная область
  TaskTeg = ExcelSheet.Cells(3, 2)    'Тэг
  ScoreTaskGroupCK = ExcelSheet.Cells(8, 11) 'Скоринг Группа ЦК
  ScoreFuncArea = ExcelSheet.Cells(8, 12)   'Скоринг Функциональная область
  ScoreTaskTeg = ExcelSheet.Cells(8, 13)    'Скоринг Тэг
  'Пытаемся найти главную задачу по BIQ
  BiqTaskID = 0
  For Each BiqTask In ActiveProject.Tasks
    If BiqTask.GetField(FieldID:=projectField_JirID) = NumBIQ Then
      BiqTaskID = BiqTask.id
    End If
  Next BiqTask
  'Поиск задачи с одиннаковой системой
  If SearchIdentBIQ(TaskType) = True Then
    If MsgBox("Такая задача уже есть в системе, продолжить добавление?", vbYesNo, "Добавление") = vbNo Then
      xlobject.Quit 'Закрытие Excel файла
      CreateTasksByExcel = True
      Exit Function
    End If
  Else
    BiqTaskID = 0
  End If
  Index = 1
  'Если не нашли главную задачу
  If BiqTaskID = 0 Then
    'Создаем главную задачу по BIQ
    FirstTask = False
    Call AddNewTask(True, FirstTask, StartDate, NumBIQ, TaskType, BIQName, "", 0, False, "", "", "", Index, IndexTaskFirst, IndexTaskLast)
    'Создаем подзадачу для системы
    FirstTask = True
    Call AddNewTask(False, FirstTask, StartDate, "", TaskType, BIQName, "", BiqTaskID, False, ITService, "", "", Index, IndexTaskFirst, IndexTaskLast)
  Else
    'Создаем подзадачу для системы
    FirstTask = False
    Call AddNewTask(False, FirstTask, StartDate, "", TaskType, BIQName, "", BiqTaskID, False, ITService, "", "", Index, IndexTaskFirst, IndexTaskLast)
  End If
  
  FirstTask = True
  For i = 8 To 24
    'Пропускаем строчки Итого и с пустым наименованием
    If Len(Trim(ExcelSheet.Cells(i, 3))) <> 0 Then
      TypeWork = ExcelSheet.Cells(i, 5) 'Тип работ
      TaskActor = ExcelSheet.Cells(i, 6) 'Исполнитель
      TaskName = Trim(ExcelSheet.Cells(i, 3)) 'Имя задачи
      Parenthesis = InStr(1, TaskName, "(") 'Наличие круглой скобкой
      If Parenthesis Then
        TaskName = Trim(Mid(TaskName, 1, Parenthesis - 1)) 'Название задачи
      End If
      TaskHours = ExcelSheet.Cells(i, 7) 'Время задачи
      Call AddNewTask(False, FirstTask, StartDate, "", TaskType, TaskName, TaskHours, BiqTaskID, True, ITService, TypeWork, TaskActor, Index, IndexTaskFirst, IndexTaskLast)
    End If
  Next i
'  'Оценка тестировщика
'  TypeWork = 510 'Тип работ
'  TaskActor = "Тестировщик2" 'Исполнитель
'  TaskName = tbFieldTest 'Имя задачи
'  TaskHours = tbFieldHoursTest 'Время задачи
'  Call AddNewTask(False, FirstTask, StartDate, "", TaskType, TaskName, TaskHours, BiqTaskID, True, ITService, TypeWork, TaskActor, Index, IndexTaskFirst, IndexTaskLast)
'  'Оценка подрядчика
'  TypeWork = 511 'Тип работ
'  TaskActor = "Подрядчик" 'Исполнитель
'  TaskName = tbFieldPodr 'Имя задачи
'  TaskHours = tbFieldHoursPodr 'Время задачи
'  Call AddNewTask(False, FirstTask, StartDate, "", TaskType, TaskName, TaskHours, BiqTaskID, True, ITService, TypeWork, TaskActor, Index, IndexTaskFirst, IndexTaskLast)
  'функция заполнения предшественников
  Call TaskPredInPut(ExcelSheet, StartDate, IndexTaskFirst, IndexTaskLast)
  
  xlobject.Quit 'Закрытие Excel файла
  
  'Заполняем исполнителей
  Call FillResources(TaskGroupCK, FuncArea, TaskTeg, SystemCode, IndexTaskFirst, IndexTaskLast, ScoreTaskGroupCK, ScoreFuncArea, ScoreTaskTeg)
  
  'Растягивание задач для устранения перегруза
  Call ExtendTasks(IndexTaskFirst, IndexTaskLast)

  'Растягиваем даты задач
  Call StretchTasks(IndexTaskFirst, IndexTaskLast)
  
  'Дата завершения
  Call TaskDateEnd(IndexTaskFirst, IndexTaskLast)
  
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  CreateTasksByExcel: ", False, False)
  CreateTasksByExcel = True
  
End Function

'Поиск задачи второго уровня с одиннаковой системой
Public Function SearchIdentBIQ(SystemCode) As Boolean
  Dim BiqTask As Task
  For Each BiqTask In ActiveProject.Tasks
    If BiqTask.GetField(FieldID:=projectField_JiraProjName) = SystemCode Then
      SearchIdentBIQ = True
      Exit Function
    End If
  Next BiqTask
  SearchIdentBIQ = False

End Function 'SearchIdentBIQ

'запись даты завершения
Sub TaskDateEnd(IndexTaskFirst, IndexTaskLast)
  Dim BiqTask As Task
  Dim ForDate As Date
  For Each BiqTask In ActiveProject.Tasks
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      If (BiqTask.Finish > ForDate Or BiqTask.id = IndexTaskFirst) Then
        ForDate = BiqTask.Finish
      End If
    End If
  Next BiqTask
  tbEndDate.Text = ForDate
  
End Sub

'процедура растягивания задач для устранения перегруза
Sub ExtendTasks(IndexTaskFirst, IndexTaskLast)
  'Начинается отсчет времени функции
  TimeForSet = Timer
  Dim BiqTask As Task
  Dim TaskRes As Resource
  Dim resAss  As Assignment
  'Цикл по всем задачам
  For Each BiqTask In ActiveProject.Tasks
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      For Each resAss In BiqTask.Assignments
        'Бежим по всем датам задачи
        For CheckDate = BiqTask.Start To BiqTask.Finish
          HoursDayLoad = GetResLoad(CheckDate, resAss.Resource)
          HoursDayLoadBiq = GetResLoadTask(CheckDate, BiqTask, resAss.ResourceID)
          HoursDayHas = GetResAvailability(CheckDate, resAss.Resource) * 8
          'Если в день больше чем возможно
          If HoursDayHas < HoursDayLoad Then
            'Если перегруз можно снять текущей задачей
            If HoursDayLoad - HoursDayHas < HoursDayLoadBiq Then
              Percent = (HoursDayHas - (HoursDayLoad - HoursDayLoadBiq)) / HoursDayHas
              Call SetTaskResProcent(BiqTask, resAss.ResourceID, Percent)
            End If
          End If
        Next CheckDate
      Next resAss
    End If
  Next BiqTask
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  ExtendTasks: ", False, False)

End Sub

'функция получения часов в день запланированных на ресурсе
Public Function GetResLoadTask(CheckDate, BiqTask, TaskActorId) As Single
  Dim TaskRes As Resource
  Dim resAss  As Assignment
  TimePerest = 0
  'Цикл по всем задачам ресурса на обрабатываемый день
  For Each resAss In BiqTask.Assignments
    If resAss.Start <= CheckDate And resAss.Finish >= CheckDate And resAss.ResourceID = TaskActorId Then
      Set TaskTSD = resAss.TimeScaleData(CheckDate, CheckDate, TimescaleUnit:=4)
      For i = 1 To TaskTSD.Count
        If Not TaskTSD(i).Value = "" Then
          TimePerest = TimePerest + TaskTSD(i).Value / (60)  'Нагрузку часов в день
        End If
      Next i
    End If
  Next resAss
  GetResLoadTask = TimePerest

End Function 'GetResLoadTask

'функция получения часов в день запланированных на ресурсе
Public Function GetResLoad(CheckDate, CheckRes) As Single
  Dim resAss  As Assignment
  For Each resAss In CheckRes.Assignments
    If resAss.Start <= CheckDate And CheckDate <= resAss.Finish Then
      Set TaskTSD = resAss.TimeScaleData(CheckDate, CheckDate, TimescaleUnit:=4)
      For i = 1 To TaskTSD.Count
        If Not TaskTSD(i).Value = "" Then
          TimePerest = TimePerest + TaskTSD(i).Value / (60)  'Нагрузку часов в день
        End If
      Next i
    End If
  Next resAss
  GetResLoad = TimePerest

End Function 'GetResLoad

'функция получения часов за период
Public Function GetResLoadPeriod(Res, BegDate, EndDate) As Single
  Dim resAss  As Assignment
  For Each resAss In CheckRes.Assignments
    If resAss.Start <= BegDate And EndDate <= resAss.Finish Then
      Set TaskTSD = resAss.TimeScaleData(BegDate, EndDate, TimescaleUnit:=4)
      For i = 1 To TaskTSD.Count
        If Not TaskTSD(i).Value = "" Then
          TimePerest = TimePerest + TaskTSD(i).Value / (60)  'Нагрузку часов в день
        End If
      Next i
    End If
  Next resAss
  GetResLoadPeriod = TimePerest
  
End Function 'GetResLoadPeriod

'функция получения доступности 0..1
Public Function GetResAvailability(CheckDate, CheckRes) As Single
  Dim TaskAvailabity As Availability
  ResAvailability = 0
  For Each TaskAvailabity In CheckRes.Availabilities
    If TaskAvailabity.AvailableFrom < CheckDate And CheckDate < TaskAvailabity.AvailableTo Then
      ResAvailability = ResAvailability + TaskAvailabity.AvailableUnit
    End If
  Next TaskAvailabity
  GetResAvailability = ResAvailability / 100
  
End Function 'GetResAvailability

'процедура замена даты для растяжение задач с типом НН
Sub StretchTasks(IndexTaskFirst, IndexTaskLast)
  'Начинается отсчет времени функции
  TimeForSet = Timer
  Dim BiqTaskPred As Task
  Dim BiqTaskDesc As Task
  'Цикл поиска задачи с типом предшественников НН
  For Each BiqTaskDesc In ActiveProject.Tasks
    If (BiqTaskDesc.id >= IndexTaskFirst And BiqTaskDesc.id <= IndexTaskLast) Then
      TaskDesc = BiqTaskDesc.GetField(FieldID:=projectField_Predecessors)
      If (InStr(TaskDesc, "НН") <> 0) Then
        'Текущие дата конца задачи
        DateEndDesc = Mid(BiqTaskDesc.GetField(FieldID:=projectField_End), 4)
        'Поиск требуемого дата начала
        'Номер предшественника с НН
        NumPred = Left(Mid(TaskDesc, InStr(TaskDesc, ";") + 1), InStr(Mid(TaskDesc, InStr(TaskDesc, ";")), "НН") - 2)
        'Цикл поиска предшественника по номеру
        For Each BiqTaskPred In ActiveProject.Tasks
          If NumPred = BiqTaskPred.id Then
            'Дата начала задачи предшественника - требуемое дата начала
            DateStartPred = Mid(BiqTaskPred.GetField(FieldID:=projectField_Start), 4)
          End If
        Next BiqTaskPred
        DiffDateDayNeed = DateDiff("d", DateStartPred, DateEndDesc) 'нужно дней
        HoursToWork = BiqTaskDesc.GetField(FieldID:=projectField_Cost)
        HoursToWork = Left(HoursToWork, InStr(HoursToWork, " ")) 'трудозатраты
        
        AllHoursInDiff = DiffDateDayNeed / 7 * 40
        procent = HoursToWork / AllHoursInDiff * 100
        RoundProcent = WorksheetFunction.Round(procent + 0.5, 0)
        'Замена процента
        Call SetTaskResProcent(BiqTaskDesc, -1, RoundProcent / 100)

        'Текущие дата начала
        DateStartDesc = Mid(BiqTaskDesc.GetField(FieldID:=projectField_Start), 4)
        DiffDateDayNow = DateDiff("d", DateStartPred, DateStartDesc) 'сейчас столько дней'
        If DiffDateDayNow > 0 Then
          HoursToWorkNew = HoursToWork + RoundProcent / 100 * (DiffDateDayNow * 8)
          'Замена трудоемкости
          BiqTaskDesc.SetField FieldID:=projectField_Cost, Value:=HoursToWorkNew
        End If
      End If
    End If
  Next BiqTaskDesc
  
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  StretchTasks: ", False, False)

End Sub

'процедура назначения исполнителей
Sub FillResources(TaskGroupCK, FuncArea, TaskTeg, SystemCode, IndexTaskFirst, IndexTaskLast, ScoreTaskGroupCK, ScoreFuncArea, ScoreTaskTeg)
  'Начинается отсчет времени функции
  TimeForSet = Timer
  Dim BiqTask As Task
  Dim Res     As Resource
  'Бежим по всем задачам требуеющих поиска исполнителей
  For Index = 1 To 3
  MaxScoreRes = 0
    If Index = 1 Then
      TaskActor = "Аналитик"
    End If
    If Index = 2 Then
      TaskActor = "Разработчик"
    End If
    If Index = 3 Then
      TaskActor = "Тестировщик"
    End If
    NumMainTask = DefinMainTaskForRes(IndexTaskFirst, IndexTaskLast, TaskActor)
    'Ищем главную задачу для каждого типа(Аналитик,Разработчик,Тестировщик)
    For Each BiqTask In ActiveProject.Tasks
      If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
        If (BiqTask.id = NumMainTask) Then
          MaxScoreRes = 0
          For Each Res In ActiveProject.Resources
            ScoreRes = 0
            'Условия жесткие
            If ((Res.GetField(FieldID:=projectField_System1) = SystemCode) Or (Res.GetField(FieldID:=projectField_System2) = SystemCode)) Then
              If (Res.GetField(FieldID:=projectField_ResGroup) = TaskActor) Then
                'Условия мягкие
                'Проверка по группе
                If (Res.GetField(FieldID:=projectField_ResGroupCk) = TaskGroupCK) Then
                  ScoreRes = ScoreRes + ScoreTaskGroupCK
                End If
                'Проверка по тегу
                If (TaskTeg = "" Or Res.GetField(FieldID:=projectField_Teg) = TaskTeg) Then
                  ScoreRes = ScoreRes + ScoreTaskTeg
                End If
                'Проверка по Функциональной области
                If ((Res.GetField(FieldID:=projectField_FuncArea1) = FuncArea) Or (Res.GetField(FieldID:=projectField_FuncArea2) = FuncArea) Or (Res.GetField(FieldID:=projectField_FuncArea3) = FuncArea)) Then
                  ScoreRes = ScoreRes + ScoreFuncArea
                End If
              End If
            End If
            'Проверка скорринга текущего ресурса с максимальным
            If ScoreRes > MaxScoreRes Then
              Percent = BiqTask.Assignments(1).Units
              TaskActorId = Res.id
              MaxScoreRes = ScoreRes
            End If
          Next Res
          'Поиск даты на главной задаче
          CurDate = BiqTask.Start
          BiqTask.Start = SearchMainTaskStartDate(IndexTaskFirst, IndexTaskLast, TaskActorId, BiqTask.Start, BiqTask.Finish, Percent)
          If CurDate <> BiqTask.Start Then
            MsgBox "Главная задача по " & TaskActor & " " & BiqTask.name & " начинается с " & BiqTask.Start
          End If
          'Запись в главную задачу
          Call SetTaskResProcent(BiqTask, TaskActorId, Percent)
          Exit For
        End If
      End If
    Next BiqTask
    'Заполняем побочные задачи этого типа(Аналитик,Разработчик,Тестировщик)
    For Each BiqTask In ActiveProject.Tasks
      If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
        If (TaskActor = Left(BiqTask.Assignments(1).ResourceName, Len(BiqTask.Assignments(1).ResourceName) - 1)) Then
          If (BiqTask.id <> NumMainTask) Then
            Percent = BiqTask.Assignments(1).Units
            BiqTask.Start = SearchSideTaskStartDate(IndexTaskFirst, IndexTaskLast, TaskActorId, BiqTask.Start, BiqTask.Finish, Percent)
            Call SetTaskResProcent(BiqTask, TaskActorId, Percent)
          End If
        End If
      End If
    Next BiqTask
  Next Index
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  FillResources: ", False, False)

End Sub

'Поиск Даты на главной задаче
Public Function SearchMainTaskStartDate(IndexTaskFirst, IndexTaskLast, TaskActorId, StartDate, FinishDate, Percent) As Date
  Dim resAss  As Assignment
  Dim Res     As Resource
  Dim assTask As Task
  SearchMainTaskStartDate = StartDate
  'Длительность задачи в большую сторону
  DurationDays = WorksheetFunction.RoundUp(FinishDate - StartDate, 0)
  'Цикл с начала планирования до плюс 120 дней
  For CurrentDateNew = StartDate To StartDate + 120
    'Цикл проверки по длительности
    For CurrentDate = CurrentDateNew To CurrentDateNew + DurationDays
      TimePerest = 0
      For Each Res In ActiveProject.Resources
        If (Res.id = TaskActorId) Then
          For Each resAss In Res.Assignments
            Set assTask = resAss.Task
            If assTask.Start <= CurrentDate And assTask.Finish >= CurrentDate Then
              Set TaskTSD = assTask.TimeScaleData(CurrentDate, CurrentDate, TimescaleUnit:=4)
              For i = 1 To TaskTSD.Count
                If Not TaskTSD(i).Value = "" Then
                  TimePerest = TimePerest + TaskTSD(i).Value / (60)  'Нагрузку часов в день
                End If
              Next i
            End If
          Next resAss
        End If
      Next Res
      If TimePerest >= 4 Then
        Exit For
      End If
      If (CurrentDate = CurrentDateNew + DurationDays) Then
        SearchMainTaskStartDate = CurrentDateNew
        Exit Function
      End If
    Next CurrentDate
  Next CurrentDateNew
  
End Function 'SearchMainTaskStartDate

'Поиск даты на побочной задаче
Public Function SearchSideTaskStartDate(IndexTaskFirst, IndexTaskLast, TaskActorId, StartDate, FinishDate, Percent) As Date
  Dim resAss  As Assignment
  Dim Res     As Resource
  Dim assTask As Task
  SearchSideTaskStartDate = StartDate
  'Длительность задачи в большую сторону
  DurationDays = WorksheetFunction.RoundUp(FinishDate - StartDate, 0)
  'Цикл с начала планирования до плюс 120 дней
  For CurrentDateNew = StartDate To StartDate + 120
    'Цикл проверки по длительности
    For CurrentDate = CurrentDateNew To CurrentDateNew + DurationDays
      TimePerest = 0
      For Each Res In ActiveProject.Resources
        If (Res.id = TaskActorId) Then
          For Each resAss In Res.Assignments
            Set assTask = resAss.Task
            If assTask.Start <= CurrentDate And assTask.Finish >= CurrentDate Then
              Set TaskTSD = assTask.TimeScaleData(CurrentDate, CurrentDate, TimescaleUnit:=4)
              For i = 1 To TaskTSD.Count
                If Not TaskTSD(i).Value = "" Then
                  TimePerest = TimePerest + TaskTSD(i).Value / (60)  'Нагрузку часов в день
                End If
              Next i
            End If
          Next resAss
        End If
      Next Res
      If 8 - TimePerest >= Percent * 8 Then
        Exit For
      End If
      If (CurrentDate = CurrentDateNew + DurationDays) Then
        SearchSideTaskStartDate = CurrentDateNew
        Exit Function
      End If
    Next CurrentDate
  Next CurrentDateNew
  
End Function 'SearchSideTaskStartDate

'функция поиска номера главной задачи
Public Function DefinMainTaskForRes(IndexTaskFirst, IndexTaskLast, TaskActor) As Long
  TimeForSet = Timer
  Dim BiqTask As Task
  FindProcent = 0
  For Each BiqTask In ActiveProject.Tasks
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      If (Left(BiqTask.Assignments(1).ResourceName, Len(BiqTask.Assignments(1).ResourceName) - 1) = TaskActor And BiqTask.Assignments(1).Units > FindProcent) Then
        FindProcent = BiqTask.Assignments(1).Units
        NumberBiq = BiqTask.id
      End If
    End If
  Next BiqTask
  Call SetTimeForTxt(Timer - TimeForSet, "  DefinMainTaskForRes: ", False, False)
  DefinMainTaskForRes = NumberBiq
  
End Function 'DefinMainTaskForRes

'Назначение ресурса на задачу
Sub SetTaskResProcent(BiqTask, TaskActorId, Percent)
  Dim Ass As Assignment
  'Попытка обновления если ресурс уже есть на задаче
  For Each Ass In BiqTask.Assignments
    If TaskActorId = -1 Or Ass.ResourceID = TaskActorId Then
      Ass.Units = Percent
      If TaskActorId <> -1 Then
        Exit Sub
      End If
    End If
  Next Ass
  'Если не нашли создаем новый
  If TaskActorId <> -1 Then
    BiqTask.Assignments.Add BiqTask.id, TaskActorId, Percent
    If BiqTask.Assignments.Count - 1 > 0 Then
      BiqTask.Assignments(BiqTask.Assignments.Count - 1).Delete
    End If
  End If
End Sub

'функция заполнения предшественников
Sub TaskPredInPut(ExcelSheet, BiqStartDate, IndexTaskFirst, IndexTaskLast)
  'Начинается отсчет времени функции
  TimeForSet = Timer
  i = 8
  Dim BiqTask As Task
  For Each BiqTask In ActiveProject.Tasks
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      Do Until ExcelSheet.Cells(i, 3) <> ""
        i = i + 1
      Loop
      TaskPredecessors = ExcelSheet.Cells(i, 4) 'Предешественник
      If TaskPredecessors <> "" Then
        TaskPredecessors = DelPred(TaskPredecessors, IndexTaskFirst, IndexTaskLast)
      Else
        BiqTask.SetField FieldID:=projectField_Start, Value:=BiqStartDate
      End If
      BiqTask.SetField FieldID:=projectField_Predecessors, Value:=TaskPredecessors
      i = i + 1
    End If
  Next BiqTask
  'Функция замены предшественников
  Call Zerotasksdel(IndexTaskFirst, IndexTaskLast)
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  TaskPredInPut: ", False, False)
  
End Sub

'Процедура замены у потомков
Sub Zerotasksdel(IndexTaskFirst, IndexTaskLast)
  Dim BiqTask As Task
  Dim BiqTaskSecond As Task
  TempZeroTaskID = 0
  TempPredec = 0
  KeyFoLoop = False
  For Each BiqTask In ActiveProject.Tasks 'Цикл поиска задач с нулем часов
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      If BiqTask.GetField(FieldID:=projectField_Cost) = "0 ч" Then
        TempZeroTaskID = BiqTask.id 'ИД задачи с нулем часов
        TempPredec = BiqTask.GetField(FieldID:=projectField_Predecessors) 'Предшественник задачи с нулем часов
        Call RepCycPred(TempZeroTaskID, TempPredec) 'Функция замены у потомков
      End If
    End If
  Next BiqTask
  'Удаление всех задач с нулем часов
  Call DeleteAllZeroTasks(IndexTaskFirst, IndexTaskLast)
  
End Sub

'Функция замены у потомков 2
Sub RepCycPred(TempZeroTaskID, TempPredec)
  Dim BiqTask As Task
  For Each BiqTask In ActiveProject.Tasks 'Цикл замены у потомков
    TempDesc = BiqTask.GetField(FieldID:=projectField_Predecessors)
    If InStr(1, TempDesc, ";") = 0 Then
      If CStr(TempZeroTaskID) = TempDesc And TempZeroTaskID <> 0 Then
        BiqTask.SetField FieldID:=projectField_Predecessors, Value:=TempPredec
      End If
    Else
      If InStr(1, TempDesc, CStr(TempZeroTaskID)) <> 0 And TempZeroTaskID <> 0 Then
        If InStr(1, TempDesc, CStr(TempZeroTaskID)) = 1 Then
          NumTempDesc = InStr(1, TempDesc, ";")
          NewTempDesc = Mid(TempDesc, NumTempDesc + 1)
          NewTempDesc = TempPredec + ";" + NewTempDesc
          If (InStr(1, NewTempDesc, ";") = 1) Then
            NewTempDesc = Mid(NewTempDesc, 2)
          End If
          BiqTask.SetField FieldID:=projectField_Predecessors, Value:=NewTempDesc
        Else
          NumTempDesc = InStr(1, TempDesc, ";")
          NewTempDesc = Mid(TempDesc, 1, NumTempDesc - 1)
          NewTempDesc = TempPredec + ";" + NewTempDesc
          If (InStr(1, NewTempDesc, ";") = 1) Then
            NewTempDesc = Mid(NewTempDesc, 2)
          End If
          BiqTask.SetField FieldID:=projectField_Predecessors, Value:=NewTempDesc
        End If
      End If
    End If
  Next BiqTask
  
End Sub

'функция изменения сложных предшественников
Public Function DelPred(TaskPredecessors, IndexTaskFirst, IndexTaskLast) As String
  Delim = InStr(1, TaskPredecessors, ";")
  NewPredecessors = ""
  Do While Delim
    TempPredecessors = Mid(TaskPredecessors, 1, Delim - 1)
    TaskPredecessors = Mid(TaskPredecessors, Delim + 1)
    NumDelim = InStr(1, TempPredecessors, "#")
    If NumDelim Then
      NumPredecessor = Mid(TempPredecessors, 1, NumDelim - 1)
      SufPredecessor = Mid(TempPredecessors, NumDelim + 1)
    Else
      NumPredecessor = TempPredecessors
      SufPredecessor = ""
    End If
    NumPredecessor = CStr(IndexTaskFirst + CInt(NumPredecessor) + FirstTaskId)
    If NewPredecessors = "" Then
      NewPredecessors = NumPredecessor + SufPredecessor
    Else
      NewPredecessors = NewPredecessors + ";" + NumPredecessor + SufPredecessor
    End If
    Delim = InStr(1, TaskName, ";")
  Loop
  NumDelim = InStr(1, TaskPredecessors, "#")
  If NumDelim Then
    NumPredecessor = Mid(TaskPredecessors, 1, NumDelim - 1)
    SufPredecessor = Mid(TaskPredecessors, NumDelim + 1)
  Else
    NumPredecessor = TaskPredecessors
    SufPredecessor = ""
  End If
  NumPredecessor = CStr(IndexTaskFirst + CInt(NumPredecessor) + FirstTaskId)
  If NewPredecessors = "" Then
    NewPredecessors = NumPredecessor + SufPredecessor
  Else
    NewPredecessors = NewPredecessors + ";" + NumPredecessor + SufPredecessor
  End If
  DelPred = NewPredecessors

End Function 'DelPred

'Создание задачи в MS Project
Sub AddNewTask(MainTask, ByRef FirstTask, BiqStartDate, TaskJiraId, TaskType, TaskName, TaskHours, BiqTaskID, ToTaskDays, TaskTypeITService, TaskTypeWork, TaskActor, ByRef Index, ByRef IndexTaskFirst, ByRef IndexTaskLast)
  'Начинается отсчет времени функции
  TimeForSet = Timer
  ' Создаем задачу
  If BiqTaskID = 0 Then
    Set NewTask = ActiveProject.Tasks.Add(TaskName)
  Else
    Set NewTask = ActiveProject.Tasks.Add(TaskName, BiqTaskID + Index)
  End If
  Index = Index + 1
  ' Для главной задачи возвращаем отступ к единице
  If MainTask Then
    Do While NewTask.OutlineLevel > 1
      NewTask.OutlineOutdent
    Loop
  End If
  ' Для первой задачи делаем отступ
  If FirstTask Then
    NewTask.OutlineIndent
  End If
  ' Заполняем поля
  NewTask.SetField FieldID:=projectField_JirID, Value:=TaskJiraId
  'Заполняем поля,находящиеся на низшим уровне
  If ToTaskDays Then
    NewTask.SetField FieldID:=projectField_DurationDays, Value:=WorksheetFunction.RoundUp(((Val(TaskHours)) / 8), 0)
    ' Для первой задачи предшествиника не заполняем
    If FirstTask Then
      NewTask.SetField FieldID:=projectField_Start, Value:=BiqStartDate
      IndexTaskFirst = NewTask.id 'Первый индекс
    Else
      IndexTaskLast = NewTask.id 'Последний индекс
    End If
  End If

  ' Только для первой задачи делаем отступ
  If FirstTask Then
    FirstTask = False
  End If
  
  'Сохраняем поля описания задачи
  NewTask.SetField FieldID:=projectField_ITService, Value:=TaskTypeITService
  NewTask.SetField FieldID:=projectField_Cost, Value:=TaskHours
  NewTask.SetField FieldID:=projectField_JiraProjName, Value:=TaskType
  NewTask.SetField FieldID:=projectField_TypeWork, Value:=TaskTypeWork
  NewTask.SetField FieldID:=projectField_Actor, Value:=TaskActor
  NewTask.SetField FieldID:=projectField_ImpDate, Value:=Format(Date, "dd/mm/yyyy")
  NewTask.SetField FieldID:=projectField_EmpImpTask, Value:=Application.UserName
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  AddNewTask: ", False, False)
  
End Sub

'Получаем название полей в MS Project
Sub InitFieldConst()

  projectField_Name = FieldNameToFieldConstant("Название", pjProject)
  projectField_JirID = FieldNameToFieldConstant("Jira id", pjProject)
  projectField_Cost = FieldNameToFieldConstant("Трудозатраты", pjProject)
  projectField_Actor = FieldNameToFieldConstant("Названия ресурсов", pjProject)
  projectField_DurationDays = FieldNameToFieldConstant("Длительность", pjProject)
  projectField_Restrict = FieldNameToFieldConstant("Тип ограничения", pjProject)
  projectField_JiraProjName = FieldNameToFieldConstant("Имя проекта", pjProject)
  projectField_Predecessors = FieldNameToFieldConstant("Предшественники", pjProject)
  projectField_Start = FieldNameToFieldConstant("Начало", pjProject)
  projectField_End = FieldNameToFieldConstant("Окончание", pjProject)
  projectField_ITService = FieldNameToFieldConstant("ИТ-Сервис", pjProject)
  projectField_TypeWork = FieldNameToFieldConstant("Тип работ", pjProject)
  projectField_ImpDate = FieldNameToFieldConstant("Дата импорта", pjProject)
  projectField_EmpImpTask = FieldNameToFieldConstant("Сотрудник импортировавший задачу", pjProject)
  projectField_Teg = FieldNameToFieldConstant("Тэг", pjResource)
  projectField_ResGroup = FieldNameToFieldConstant("Группа", pjResource)
  projectField_ResGroupCk = FieldNameToFieldConstant("Группа ЦК", pjResource)
  projectField_FuncArea1 = FieldNameToFieldConstant("функ. Область 1", pjResource)
  projectField_FuncArea2 = FieldNameToFieldConstant("Функ. Область 2", pjResource)
  projectField_FuncArea3 = FieldNameToFieldConstant("Функ. Область 3", pjResource)
  projectField_System1 = FieldNameToFieldConstant("Система 1", pjResource)
  projectField_System2 = FieldNameToFieldConstant("Система 2", pjResource)

End Sub

'Выбор оценки по ЦФТ
Private Sub GetExcelFileCFTButton_Click()
  FileNameCFTTextBox.Text = ShowGetOpenDialog()
  If (TBNumBIQ.Text = "") Then
    TBNumBIQ.Text = GetBiqNum(FileNameCFTTextBox.Text)
  End If
End Sub

'Выбор оценки по БИСквиту
Private Sub GetExcelFileBISButton_Click()
  FileNameManTextBox.Text = ShowGetOpenDialog()
    If (TBNumBIQ.Text = "") Then
      TBNumBIQ.Text = GetBiqNum(FileNameManTextBox.Text)
    End If
End Sub

'Номер оценки из файла
Public Function GetBiqNum(FileExcelName) As String
  PathToExc = FileExcelName
  Set xlobject = CreateObject("Excel.Application")
  xlobject.Workbooks.Open PathToExc
  ' Если не удалось открыть, то выходим
  If xlobject.ActiveWorkbook Is Nothing Then
    xlobject.Quit 'Закрытие Excel файла
    Exit Function
  End If
  Set ExcelSheet = xlobject.ActiveWorkbook.Sheets(4)
  BIQName = ExcelSheet.Cells(1, 3) 'Название BIQ
  xlobject.Quit 'Закрытие Excel файла
  GetBiqNum = Left(BIQName, InStr(BIQName, " ") - 1)
        
End Function 'GetBiqNum

'Функция открытия проводника для выбора файла
Public Function ShowGetOpenDialog() As String
  Dim xlObj As Excel.Application
  Dim fd As Office.FileDialog
  Set xlObj = New Excel.Application
  xlObj.Visible = False
  Set fd = xlObj.Application.FileDialog(msoFileDialogFilePicker)
  With fd
    .Title = "Выберите необходимый файл" 'Название проводника
    .Filters.Add "Excel", "*.xls,*.xlsx" 'Фильтры для отоброжения файлов
    .AllowMultiSelect = False            'Только один файл
    If .Show = False Then
      Set xlObj = Nothing
      Exit Function
    End If
    ShowGetOpenDialog = .SelectedItems(1) 'Возврат результата
  End With
  Set xlObj = Nothing
  
End Function 'ShowGetOpenDialog

'Удаление всех задач с нулем часов
Sub DeleteAllZeroTasks(IndexTaskFirst, IndexTaskLast)
  Dim BiqTask As Task
  For Each BiqTask In ActiveProject.Tasks
    If BiqTask.GetField(FieldID:=projectField_Cost) = "0 ч" Then
      If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
        BiqTask.Delete 'Удаление BIQ-задачи
        IndexTaskLast = IndexTaskLast - 1
      End If
    End If
  Next BiqTask

End Sub

'Кнопка удаления
Private Sub DeleteButton_Click()
  Dim BiqTask As Task ' Для поиска задачи по BIQ
  InitFieldConst
  BIQNum = TBNumBIQ 'Номер BIQ-задачи
  BiqTaskID = 0
  'Бежим по всем задачам
  For Each BiqTask In ActiveProject.Tasks
    If BiqTask.GetField(FieldID:=projectField_JirID) = BIQNum Then
      BiqTaskID = BiqTask.id
      'Запрос для случая когда удаляется задача, которая импортирована не сегодня
      If BiqTask.GetField(FieldID:=projectField_ImpDate) < Format(Date, "dd/mm/yyyy") Then
        If MsgBox("Задача была импортирована не сегодня. Удаление таких задач не рекомендуется. Вы уверены, что хотите удалить задачу " & BIQNum & "?", vbYesNo, "Удаление") = vbYes Then
          BiqTask.Delete 'Удаление BIQ-задачи
          Exit Sub
        End If
      End If
      'Запрос для случая когда удаляется задача, на которой нет даты импортирования
      If BiqTask.GetField(FieldID:=projectField_ImpDate) = "" Then
        If MsgBox("Задача была создана вручную. Удаление таких задач не рекомендуется. Вы уверены, что хотите удалить задачу " & BIQNum & "?", vbYesNo, "Удаление") = vbYes Then
          BiqTask.Delete 'Удаление BIQ-задачи
          Exit Sub
        End If
      End If
      'В остальных случаях запрос стандартный
      If MsgBox("Вы уверены что хотите удалить " & BIQNum & "?", vbYesNo, "Удаление") = vbYes Then
        BiqTask.Delete 'Удаление BIQ-задачи
        Exit Sub
      End If
    End If
  Next BiqTask
  
  If BiqTaskID = 0 Then
    MsgBox ("Такой BIQ-задачи нет")
    Exit Sub
  End If
  'Запись протокола работы
  Call SetProtocolJob("Удаление")
  
End Sub

'Растягивание задачи в зависимости от загрузки ресурсов
Sub Perest()
  Dim resAss  As Assignment
  Dim Res     As Resource
  Dim SecRes  As Resource
  Dim assTask As Task
    
  For Each Res In ActiveProject.Resources
    '1
    For Each resAss In Res.Assignments
      Set assTask = resAss.Task
      DurationWorkDaysPerest = assTask.DurationText 'Длительность в рабочих днях
      StartDatePerest = Mid(Mid(assTask.StartText, 4), 1, 6) & "20" & Mid(assTask.StartText, 10) 'Дата начала задачи
      FinishDatePerest = Mid(Mid(assTask.FinishText, 4), 1, 6) & "20" & Mid(assTask.FinishText, 10) 'Дата конца задачи
      TimePerest = assTask.TimeScaleData(assTask.Start, assTask.Finish, TimescaleUnit:=4)(1).Value / (60) 'Нагрузку часов в день
      CurrenRes = -1
      CurrentStartDate = "31.12.2040"
      GroupFirstRes = Res.Group
      For Each SecRes In ActiveProject.Resources
        '2
        SecData = StartDatePerest
        i = 0
        Do
          HoursDay = assTask.TimeScaleData(SecData, SecData, TimescaleUnit:=4)(1).Value / (60)
          i = i + 1
          If i < 5 Then
              
          End If
          SecData = DateAdd("d", 1, SecData)
        Loop Until SecData < FinishDatePerest
      Next SecRes
    Next resAss
  Next Res
End Sub
