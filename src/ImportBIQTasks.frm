VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportBIQTasks 
   Caption         =   "Перенос BIQ задач "
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8325.001
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

' Кнопка импортировать
Private Sub ImportButton_Click()
  TimeForSet = Timer
  'Запись времени в текстовик
  Call SetTimeForTxt(0, "Начало импорта ", True, False)
  ' Создаем задачи по оценке ЦФТ
  If Len(Trim(FileNameCFTTextBox.Text)) <> 0 Then
    If CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameCFTTextBox.Text)=False Then
      Msgbox "Задача с такой система уже была создана"
      Exit Sub
    End If
  End If
        
  ' Создаем задачи по оценки БИСквит
  If Len(Trim(FileNameBISTextBox.Text)) <> 0 Then
    If CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameBISTextBox.Text)=False Then
      Msgbox "Задача с такой система уже была создана"
      Exit Sub
    End If
  End If
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
  TBNumBIQ = "BIQ-5257"
  FileNameCFTTextBox = "C:\Users\Эрнест\Documents\GitHub\Diplom\test\Расшифровка ЭО BIQ5257.xlsx"
  'FileNameCFTTextBox = "d:\info\Эрнест\Diplom\test\Расшифровка ЭО BIQ5257.xlsx"
  TBNumBIQFDelete = 5257
  
End Sub

' Создание задач по оценке
Public Function CreateTasksByExcel(NumBIQ, StartDate, ExcelFileName) as Boolean
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
  TaskGroupCK = ExcelSheet.Cells(1, 2)  'Группа
  FuncArea = ExcelSheet.Cells(2, 2)   'Функциональная область
  TaskTeg = ExcelSheet.Cells(3, 2)    'Тэг

  'Пытаемся найти главную задачу по BIQ
  BiqTaskID = 0
  For Each BiqTask In ActiveProject.Tasks
    If BiqTask.GetField(FieldID:=projectField_JirID) = NumBIQ Then
      BiqTaskID = BiqTask.id
    End If
  Next BiqTask
  'Поиск задачи с одиннаковой системой
  If SearchIdentBIQ(TaskType) = True Then
    xlobject.Quit 'Закрытие Excel файла
    CreateTasksByExcel=False
    Exit Function
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
  For i = 8 To 26
    'Пропускаем строчки Итого и с пустым наименованием
    If (UCase(Left(Trim(ExcelSheet.Cells(i, 3)), 5))) <> "ИТОГО" And Len(Trim(ExcelSheet.Cells(i, 3))) <> 0 Then
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
  
  'функция заполнения предшественников
  Call TaskPredInPut(ExcelSheet, StartDate, IndexTaskFirst, IndexTaskLast)
  
  xlobject.Quit 'Закрытие Excel файла
  
  'Заполняем исполнителей
  Call FillResources(TaskGroupCK, FuncArea, TaskTeg, SystemCode, IndexTaskFirst, IndexTaskLast)
  
  'Растягивание задач для устранения перегруза
  Call ExtendTasks(IndexTaskFirst, IndexTaskLast)
      
  'Растягиваем даты задач
  Call StretchTasks(IndexTaskFirst, IndexTaskLast)
  
  'Дата завершения
  Call TaskDateEnd (IndexTaskFirst, IndexTaskLast)
  
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  CreateTasksByExcel: ", False, False)
  CreateTasksByExcel=True
  
End Function

'Поиск задачи второго уровня с одиннаковой системой
Public Function SearchIdentBIQ(SystemCode) as Boolean
  Dim BiqTask As Task
  For Each BiqTask In ActiveProject.Tasks
    If BiqTask.GetField(FieldID:=projectField_JiraProjName)=SystemCode  Then
      SearchIdentBIQ=True
      Exit Function
    End if
  Next BiqTask
  SearchIdentBIQ=False
  
End Function'SearchIdentBIQ

'запись даты завершения
Sub TaskDateEnd(IndexTaskFirst, IndexTaskLast)
  Dim BiqTask As Task
  Dim ForDate As Date
  For Each BiqTask In ActiveProject.Tasks
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      If(BiqTask.Finish>ForDate or BiqTask.id = IndexTaskFirst) then
        ForDate=BiqTask.Finish
      End if
    End if
  Next BiqTask
  tbEndDate.Text=ForDate
  
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
          ' Если в день больше чем возможно
          If HoursDayHas < HoursDayLoad Then
            'Если перегруз можно снять текущей задачей
            If HoursDayLoad - HoursDayHas < HoursDayLoadBiq Then
              Percent = (HoursDayHas - (HoursDayLoad - HoursDayLoadBiq)) / HoursDayHas * 100
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
  ' Цикл по всем задачам ресурса на обрабатываемый день
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
    If resAss.Start <= CheckDate And resAss.Finish >= CheckDate Then
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
        Call SetTaskResProcent(BiqTaskDesc, -1, RoundProcent)

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
Sub FillResources(TaskGroupCK, FuncArea, TaskTeg, SystemCode, IndexTaskFirst, IndexTaskLast)
  'Начинается отсчет времени функции
  TimeForSet = Timer
  Dim BiqTask As Task
  Dim Ass     As Assignment
  Dim Res     As Resource
  'Бежим по всем задачам требуеющих поиска исполнителей
  For Each BiqTask In ActiveProject.Tasks
    If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
      ' Бежим по всем необходимым ресурсам
      For Each Ass In BiqTask.Assignments
        'Получаем группу ресурсов
        TaskActor = Ass.ResourceName
        'Бежим по всем досутпным ресурсам - ищем исполнителя
        For Each Res In ActiveProject.Resources
          If (Res.GetField(FieldID:=projectField_ResGroupCk) = TaskGroupCK) And (TaskTeg = "" Or Res.GetField(FieldID:=projectField_Teg) = TaskTeg) _
          And ((Res.GetField(FieldID:=projectField_FuncArea1) = FuncArea) Or (Res.GetField(FieldID:=projectField_FuncArea2) = FuncArea) Or (Res.GetField(FieldID:=projectField_FuncArea3) = FuncArea)) _
          And ((Res.GetField(FieldID:=projectField_System1) = SystemCode) Or (Res.GetField(FieldID:=projectField_System2) = SystemCode)) _
          And (Res.GetField(FieldID:=projectField_ResGroup) = Left(TaskActor, Len(TaskActor) - 1)) Then
            Percent = Ass.Units
            TaskActorId = Res.id
            Call SetTaskResProcent(BiqTask, TaskActorId, Percent)
            Exit For
          End If
        Next Res
      Next Ass
    End If
  Next BiqTask
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  FillResources: ", False, False)
        
End Sub

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
  ' Если не нашли создаем новый
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

' Создание задачи в MS Project
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
        
  'Запись времени в текстовик
  Call SetTimeForTxt(Timer - TimeForSet, "  AddNewTask: ", False, False)
  
End Sub

' Получаем название полей в MS Project
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
  projectField_Teg = FieldNameToFieldConstant("Тэг", pjResource)
  projectField_ResGroup = FieldNameToFieldConstant("Группа", pjResource)
  projectField_ResGroupCk = FieldNameToFieldConstant("Группа ЦК", pjResource)
  projectField_FuncArea1 = FieldNameToFieldConstant("функ. Область 1", pjResource)
  projectField_FuncArea2 = FieldNameToFieldConstant("Функ. Область 2", pjResource)
  projectField_FuncArea3 = FieldNameToFieldConstant("Функ. Область 3", pjResource)
  projectField_System1 = FieldNameToFieldConstant("Система 1", pjResource)
  projectField_System2 = FieldNameToFieldConstant("Система 2", pjResource)

End Sub

' Выбор оценки по ЦФТ
Private Sub GetExcelFileCFTButton_Click()
  FileNameCFTTextBox.Text = ShowGetOpenDialog()
  If (TBNumBIQ.Text = "") Then
    TBNumBIQ.Text = GetBiqNum(FileNameCFTTextBox.Text )
  End If
End Sub

' Выбор оценки по БИСквиту
Private Sub GetExcelFileBISButton_Click()
  FileNameBISTextBox.Text = ShowGetOpenDialog()
    If (TBNumBIQ.Text = "") Then
      TBNumBIQ.Text = GetBiqNum(FileNameBISTextBox.Text)
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
    If MsgBox("Вы уверены что хотите удалить " & BIQNum & "?", vbYesNo, "Удаление") = vbYes Then
      For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_JirID) = BIQNum Then
          BiqTaskID = BiqTask.id
          BiqTask.Delete 'Удаление BIQ-задачи
        End If
      Next BiqTask
      If BiqTaskID = 0 Then
        MsgBox ("Такой BIQ-задачи нет")
        Exit Sub
      End If  
  End If
  'Запись протокола работы
  Call SetProtocolJob("Удаление")
  
End Sub

' Растягивание задачи в зависимости от загрузки ресурсов
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
