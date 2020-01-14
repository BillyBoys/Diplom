VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportBIQTasks 
   Caption         =   "Перенос BIQ задач "
   ClientHeight    =   6750
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
Dim projectField_ITService    As Long
Dim projectField_TypeWork     As Long

Dim TSV               As TimeScaleValues
Dim pTSV              As TimeScaleValues
Dim resAssArr         As Assignments
Dim resFilteredTask() As Task
Dim assTaskLoop       As Task
Dim resAss            As Assignment
Dim assTask           As Task
Dim Res               As Resource
Dim SecRes            As Resource
Dim AllRes            As Resources
Dim taskTime          As Variant
Dim arrTime           As Variant

Dim IndexTaskFirst    As Long
Dim IndexTaskLast     As Long
Dim Index             As Long
'Dim StartDate         As Date

Private Sub HoursRes_Click()

' Определение доступности ресурса
'    Set rs = ActiveProject.Resources
'    arrTime = 0
'    For Each r In rs
'        If r.name = ResNameForHours Then
'            Set resAssArr = r.Assignments
'            For Each resAss In resAssArr
'                Set assTask = resAss.Task
'                Set TSV = assTask.TimeScaleData(CDate(StartDateForHours.Value), CDate(EndDateForHours.Value), TimescaleUnit:=4)
'                For i = 1 To TSV.Count
'                        taskTime = ""
'                    If Not TSV(i).Value = "" Then
'                        taskTime = TSV(i).Value / (60)
'                    End If
'                    If taskTime <> "" Then
'                        arrTime = arrTime + taskTime
'                    End If
'                Next i
'            Next resAss
'        End If
'    Next r
'    msgbox arrTime
Perest

End Sub

' Кнопка импортировать
Private Sub ImportButton_Click()

    ' Создаем задачи по оценке ЦФТ
    If Len(Trim(FileNameCFTTextBox.Text)) <> 0 Then
        Call CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameCFTTextBox.Text)
    End If
    ' Создаем задачи по оценки БИСквит
    If Len(Trim(FileNameBISTextBox.Text)) <> 0 Then
        Call CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameBISTextBox.Text)
    End If
    
End Sub

' Инициализация полей
Private Sub UserForm_Initialize()
    tbStartDate = Format(Date, "dd/mm/yyyy")
    TBNumBIQ = "BIQ-5257"
    'FileNameCFTTextBox = "C:\Users\Эрнест\Documents\GitHub\Diplom\Расшифровка ЭО BIQ5257.xlsx"
    FileNameCFTTextBox = "d:\info\Эрнест\git\Diplom\Расшифровка ЭО BIQ5257.xlsx"
    TBNumBIQFDelete = 5257
End Sub

' Создание задач по оценке
Sub CreateTasksByExcel(NumBIQ, StartDate, ExcelFileName)
    
    Dim BiqTask As Task ' Для поиска задачи по BIQ
    ' Получаем название полей в MS Project
    InitFieldConst
    ' Открываем оценку для чтения этапов разработки
    PathToExc = ExcelFileName
    Set xlobject = CreateObject("Excel.Application")
    xlobject.Workbooks.Open PathToExc
		
		' Если не удалось открыть, то выходим
    If xlobject.ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
		'Интересует 4 лист оценки - технический лист для данного функционала
    Set ExcelSheet = xlobject.ActiveWorkbook.Sheets(4)
    
		' Получаем данные о задаче
    BIQName   = ExcelSheet.Cells(1, 3) 'Название BIQ
		TaskType  = ExcelSheet.Cells(2, 4) 'Оцениваемая система ЦФТ
		ITService = ExcelSheet.Cells(2, 5) 'ИТ-Сервис
    
    'Пытаемся найти главную задачу по BIQ
    BIQTaskId = 0
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_JirID) = NumBIQ Then
           BIQTaskId = BiqTask.id
        End If
    Next BiqTask

    Index = 1
    If BIQTaskId = 0 Then
        'Если не нашли то создаем главную задачу по BIQ
        Call AddNewTask(True, False, StartDate, NumBIQ, TaskType, BIQName, "", 0, False, "", "", "")
        'Создаем подзадачу для системы
        Call AddNewTask(False, True, StartDate, "", TaskType, BIQName, "", 0, False, ITService, "", "")
    Else
        'Создаем подзадачу для системы
        Call AddNewTask(False, False, StartDate, "", TaskType, BIQName, "", BIQTaskId, False, ITService, "", "")
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
            Call AddNewTask(False, FirstTask, StartDate, "", TaskType, TaskName, TaskHours, BIQTaskId, True, ITService, TypeWork, TaskActor)
            ' Для первой задачи делаем отступ
            If FirstTask Then
                FirstTask = False
            End If
        End If
    Next i
		'функция заполнения предшественников
    Call TaskPredInPut(ExcelSheet, StartDate)
    xlobject.Quit 'Закрытие Excel файла
    
End Sub

'функция заполнения предшественников
Sub TaskPredInPut(ExcelSheet, BiqStartDate)  

    i = 8
    Dim BiqTask As Task
    For Each BiqTask In ActiveProject.Tasks
        If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
            Do Until ExcelSheet.Cells(i, 3) <> ""
                i = i + 1
            Loop
            TaskPredecessors = ExcelSheet.Cells(i, 4) 'Предешественник
            If TaskPredecessors <> "" Then
                TaskPredecessors = DelPred(TaskPredecessors)
            Else
                BiqTask.SetField FieldID:=projectField_Start, Value:=BiqStartDate
            End If
            BiqTask.SetField FieldID:=projectField_Predecessors, Value:=TaskPredecessors
            i = i + 1
        End If
    Next BiqTask
		'Функция замены предшественников
    Zerotasksdel 
    
End Sub

'Функция замены у потомков
Sub Zerotasksdel() 
            
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
    DeleteAllZeroTasks 
    
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
Public Function DelPred(TaskPredecessors) As String 

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

End Function

' Создание задачи в MS Project
Sub AddNewTask(MainTask, FirstTask, BiqStartDate, TaskJiraId, TaskType, TaskName, TaskHours, BIQTaskId, ToTaskDays, TaskTypeITService, TaskTypeWork, TaskActor)
    
    ' Создаем задачу
    If BIQTaskId = 0 Then
        Set NewTask = ActiveProject.Tasks.Add(TaskName)
    Else
        Set NewTask = ActiveProject.Tasks.Add(TaskName, BIQTaskId + Index)
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
		'Сохраняем поля описания задачи
    NewTask.SetField FieldID:=projectField_ITService   , Value:=TaskTypeITService
    NewTask.SetField FieldID:=projectField_Cost        , Value:=TaskHours
    NewTask.SetField FieldID:=projectField_JiraProjName, Value:=TaskType
    NewTask.SetField FieldID:=projectField_TypeWork    , Value:=TaskTypeWork
    NewTask.SetField FieldID:=projectField_Actor       , Value:=TaskActor
    
End Sub

' Получаем название полей в MS Project
Sub InitFieldConst()

    projectField_Name         = FieldNameToFieldConstant("Название", pjProject)
    projectField_JirID        = FieldNameToFieldConstant("Jira id", pjProject)
    projectField_Cost         = FieldNameToFieldConstant("Трудозатраты", pjProject)
    projectField_Actor        = FieldNameToFieldConstant("Названия ресурсов", pjProject)
    projectField_DurationDays = FieldNameToFieldConstant("Длительность", pjProject)
    projectField_Restrict     = FieldNameToFieldConstant("Тип ограничения", pjProject)
    projectField_JiraProjName = FieldNameToFieldConstant("Имя проекта", pjProject)
    projectField_Predecessors = FieldNameToFieldConstant("Предшественники", pjProject)
    projectField_Start        = FieldNameToFieldConstant("Начало", pjProject)
    projectField_ITService    = FieldNameToFieldConstant("ИТ-Сервис", pjProject)
    projectField_TypeWork     = FieldNameToFieldConstant("Тип работ", pjProject)
  
End Sub

' Выбор оценки по ЦФТ
Private Sub GetExcelFileCFTButton_Click()
    FileNameCFTTextBox.Text = ShowGetOpenDialog()
End Sub

' Выбор оценки по БИСквиту
Private Sub GetExcelFileBISButton_Click()
    FileNameBISTextBox.Text = ShowGetOpenDialog()
End Sub

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
    
End Function

'Удаление последней задачи с нулем часов
Sub DeleteLastZeroHours()
    
    Dim BiqTask As Task
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_Cost) = "0 ч" Then
            If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
                BIQTaskId = BiqTask.id
            End If
        End If
    Next BiqTask
    IndexBIQ = 0
    For Each BiqTask In ActiveProject.Tasks
        IndexBIQ = IndexBIQ + 1
        If IndexBIQ = BIQTaskId Then
            If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
                BiqTask.Delete 'Удаление BIQ-задачи
            End If
        End If
    Next BiqTask

End Sub

'Удаление всех задач с нулем часов
Sub DeleteAllZeroTasks()

    Dim BiqTask As Task
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_Cost) = "0 ч" Then
            If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
                BiqTask.Delete 'Удаление BIQ-задачи
            End If
        End If
    Next BiqTask

End Sub

'Кнопка удаления
Private Sub DeleteButton_Click()

    Dim BiqTask As Task ' Для поиска задачи по BIQ
    InitFieldConst
    BIQNum = "BIQ-" + TBNumBIQFDelete 'Номер BIQ-задачи
    BIQTaskId = 0
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_JirID) = BIQNum Then
           BIQTaskId = BiqTask.id
           BiqTask.Delete 'Удаление BIQ-задачи
        End If
    Next BiqTask
    If BIQTaskId = 0 Then
        MsgBox ("Такой BIQ-задачи нет")
    End If
    
End Sub

' Растягивание задачи в зависимости от загрузки ресурсов
Sub Perest()
  
    
    Set AllRes = ActiveProject.Resources
    For Each Res In AllRes
        Set resAssArr = Res.Assignments
        '1
        For Each resAss In resAssArr
            Set assTask = resAss.Task
            DurationWorkDaysPerest = assTask.DurationText 'Длительность в рабочих днях
            StartDatePerest = Mid(Mid(assTask.StartText, 4), 1, 6) & "20" & Mid(assTask.StartText, 10) 'Дата начала задачи
            FinishDatePerest = Mid(Mid(assTask.FinishText, 4), 1, 6) & "20" & Mid(assTask.FinishText, 10) 'Дата конца задачи
            TimePerest = assTask.TimeScaleData(assTask.Start, assTask.Finish, TimescaleUnit:=4)(1).Value / (60) 'Нагрузку часов в день
            
            CurrenRes = -1
            CurrentStartDate = "31.12.2040"
            GroupFirstRes = Res.Group
            For Each SecRes In AllRes
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
