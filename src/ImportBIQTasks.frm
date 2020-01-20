VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportBIQTasks 
   Caption         =   "������� BIQ ����� "
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

'�������� ����� � MS Project
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

'Dim IndexTaskFirst    As Long
'Dim IndexTaskLast     As Long
'Dim Index             As Long
'Dim StartDate         As Date

Private Sub HoursRes_Click()

' ����������� ����������� �������
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

' ������ �������������
Private Sub ImportButton_Click()

    ' ������� ������ �� ������ ���
    If Len(Trim(FileNameCFTTextBox.Text)) <> 0 Then
        Call CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameCFTTextBox.Text)
    End If
    ' ������� ������ �� ������ �������
    If Len(Trim(FileNameBISTextBox.Text)) <> 0 Then
        Call CreateTasksByExcel(TBNumBIQ, CDate(tbStartDate.Value), FileNameBISTextBox.Text)
    End If
    
End Sub

' ������������� �����
Private Sub UserForm_Initialize()
    tbStartDate = Format(Date, "dd/mm/yyyy")
    TBNumBIQ = "BIQ-5257"
    'FileNameCFTTextBox = "C:\Users\������\Documents\GitHub\Diplom\test\����������� �� BIQ5257.xlsx"
    FileNameCFTTextBox = "d:\info\������\Diplom\test\����������� �� BIQ5257.xlsx"
    TBNumBIQFDelete = 5257
End Sub

' �������� ����� �� ������
Sub CreateTasksByExcel(NumBIQ, StartDate, ExcelFileName)
    
    Dim BiqTask As Task ' ��� ������ ������ �� BIQ
    ' �������� �������� ����� � MS Project
    InitFieldConst
    ' ��������� ������ ��� ������ ������ ����������
    PathToExc = ExcelFileName
    Set xlobject = CreateObject("Excel.Application")
    xlobject.Workbooks.Open PathToExc
                
    ' ���� �� ������� �������, �� �������
    If xlobject.ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    '���������� 4 ���� ������ - ����������� ���� ��� ������� �����������
    Set ExcelSheet = xlobject.ActiveWorkbook.Sheets(4)
    
    ' �������� ������ � ������
    BIQName = ExcelSheet.Cells(1, 3)    '�������� BIQ
    SystemCode = ExcelSheet.Cells(2, 3) '�������
    TaskType = ExcelSheet.Cells(2, 4)   '����������� ������� ���
    ITService = ExcelSheet.Cells(2, 5)  '��-������
    TaskGroupCK = ExcelSheet.Cells(1, 2)  '������
    FuncArea = ExcelSheet.Cells(2, 2)   '�������������� �������
    TaskTeg = ExcelSheet.Cells(3, 2)    '���

    '�������� ����� ������� ������ �� BIQ
    BIQTaskId = 0
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_JirID) = NumBIQ Then
           BIQTaskId = BiqTask.id
        End If
    Next BiqTask

    Index = 1
    ' ���� �� ����� ������� ������ 
    If BIQTaskId = 0 Then
        '������� ������� ������ �� BIQ
        FirstTask = False
        Call AddNewTask(True, FirstTask, StartDate, NumBIQ, TaskType, BIQName, "", 0, False, "", "", "", Index, IndexTaskFirst, IndexTaskLast)
    End If
    
    '������� ��������� ��� �������
    FirstTask = True
    Call AddNewTask(False, FirstTask, StartDate, "", TaskType, BIQName, "", BIQTaskId, False, ITService, "", "", Index, IndexTaskFirst, IndexTaskLast)
    
    FirstTask = True
    For i = 8 To 26
       '���������� ������� ����� � � ������ �������������
        If (UCase(Left(Trim(ExcelSheet.Cells(i, 3)), 5))) <> "�����" And Len(Trim(ExcelSheet.Cells(i, 3))) <> 0 Then
            TypeWork = ExcelSheet.Cells(i, 5) '��� �����
            TaskActor = ExcelSheet.Cells(i, 6) '�����������
            TaskName = Trim(ExcelSheet.Cells(i, 3)) '��� ������
            Parenthesis = InStr(1, TaskName, "(") '������� ������� �������
            If Parenthesis Then
                TaskName = Trim(Mid(TaskName, 1, Parenthesis - 1)) '�������� ������
            End If
            TaskHours = ExcelSheet.Cells(i, 7) '����� ������
            Call AddNewTask(False, FirstTask, StartDate, "", TaskType, TaskName, TaskHours, BIQTaskId, True, ITService, TypeWork, TaskActor, Index, IndexTaskFirst, IndexTaskLast)
        End If
    Next i
    
    '������� ���������� ����������������
    Call TaskPredInPut(ExcelSheet, StartDate, IndexTaskFirst, IndexTaskLast)
    
    '��������� ������������
    Call FillResourses(TaskGroupCK, FuncArea, TaskTeg, SystemCode, IndexTaskFirst, IndexTaskLast)
        
    '����������� ���� �����
    Call StretchTasks(IndexTaskFirst, IndexTaskLast)
        
    xlobject.Quit '�������� Excel �����
    
End Sub

'������� ������ ���� ��� ���������� ����� � ����� ��
Sub StretchTasks(IndexTaskFirst, IndexTaskLast)

  Dim BiqTaskPred As Task
  Dim BiqTaskDesc As Task
  '���� ������ ������ � ����� ���������������� ��
  For Each BiqTaskDesc In ActiveProject.Tasks
    If (BiqTaskDesc.id >= IndexTaskFirst And BiqTaskDesc.id <= IndexTaskLast) Then
      TaskDesc = BiqTaskDesc.GetField(FieldID:=projectField_Predecessors)
      If (InStr(TaskDesc, "��") <> 0) Then
        '������� ���� ����� ������
        DateEndDesc = Mid(BiqTaskDesc.GetField(FieldID:=projectField_End), 4)
        '����� ���������� ���� ������
        '����� ��������������� � ��
        NumPred = Left(Mid(TaskDesc, InStr(TaskDesc, ";") + 1), InStr(Mid(TaskDesc, InStr(TaskDesc, ";")), "��") - 2)
        '���� ������ ��������������� �� ������
        For Each BiqTaskPred In ActiveProject.Tasks
          If NumPred = BiqTaskPred.id Then
            '���� ������ ������ ��������������� - ��������� ���� ������
            DateStartPred = Mid(BiqTaskPred.GetField(FieldID:=projectField_Start), 4)
          End If
        Next BiqTaskPred
        DiffDateDayNeed = DateDiff("d", DateStartPred, DateEndDesc) '����� ����
        HoursToWork = BiqTaskDesc.GetField(FieldID:=projectField_Cost)
        HoursToWork = Left(HoursToWork, InStr(HoursToWork, " ")) '������������
        
        AllHoursInDiff = DiffDateDayNeed / 7 * 40
        Procent = HoursToWork / AllHoursInDiff * 100
        RoundProcent = WorksheetFunction.Round(Procent + 0.5, 0)
        '������ ��������
        BiqTaskDesc.SetField FieldID:=projectField_Actor, Value:=Left(BiqTaskDesc.GetField(FieldID:=projectField_Actor), InStr(BiqTaskDesc.GetField(FieldID:=projectField_Actor), "[") - 1) & "[" & RoundProcent & "%]"

        '������� ���� ������
        DateStartDesc = Mid(BiqTaskDesc.GetField(FieldID:=projectField_Start), 4)
        DiffDateDayNow = DateDiff("d", DateStartPred, DateStartDesc) '������ ������� ����'
        if DiffDateDayNow > 0 Then
          HoursToWorkNew = HoursToWork + RoundProcent / 100 * (DiffDateDayNow * 8)
          '������ ������������
          BiqTaskDesc.SetField FieldID:=projectField_Cost, Value:= HoursToWorkNew
        End If
      End If
    End If
  Next BiqTaskDesc

End Sub

'������� ���������� ������������
Sub FillResourses(TaskGroupCK, FuncArea, TaskTeg, SystemCode, IndexTaskFirst, IndexTaskLast)
   
    Dim BiqTask As Task
    Set AllRes = ActiveProject.Resources
    For Each BiqTask In ActiveProject.Tasks
       If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
                            '�������� ������ ��������
            TaskActor = BiqTask.GetField(FieldID:=projectField_Actor)
            For Each Res In AllRes
                If (Res.GetField(FieldID:=projectField_ResGroupCk) = TaskGroupCK) And (TaskTeg = "" Or Res.GetField(FieldID:=projectField_Teg) = TaskTeg) _
                And ((Res.GetField(FieldID:=projectField_FuncArea1) = FuncArea) Or (Res.GetField(FieldID:=projectField_FuncArea2) = FuncArea) Or (Res.GetField(FieldID:=projectField_FuncArea3) = FuncArea)) _
                And ((Res.GetField(FieldID:=projectField_System1) = SystemCode) Or (Res.GetField(FieldID:=projectField_System2) = SystemCode)) _
                And ((Res.GetField(FieldID:=projectField_ResGroup) = Mid(TaskActor, 1, InStr(TaskActor, "[") - 2))) Then
                    TaskActor = Res.name + Mid(TaskActor, InStr(TaskActor, "["))
                End If
            Next Res
            BiqTask.SetField FieldID:=projectField_Actor, Value:=TaskActor
        End If
    Next BiqTask
    
End Sub

'������� ���������� ����������������
Sub TaskPredInPut(ExcelSheet, BiqStartDate, IndexTaskFirst, IndexTaskLast)

    i = 8
    Dim BiqTask As Task
    For Each BiqTask In ActiveProject.Tasks
        If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
            Do Until ExcelSheet.Cells(i, 3) <> ""
                i = i + 1
            Loop
            TaskPredecessors = ExcelSheet.Cells(i, 4) '���������������
            If TaskPredecessors <> "" Then
                TaskPredecessors = DelPred(TaskPredecessors, IndexTaskFirst, IndexTaskLast)
            Else
                BiqTask.SetField FieldID:=projectField_Start, Value:=BiqStartDate
            End If
            BiqTask.SetField FieldID:=projectField_Predecessors, Value:=TaskPredecessors
            i = i + 1
        End If
    Next BiqTask
                '������� ������ ����������������
    Call Zerotasksdel(IndexTaskFirst, IndexTaskLast)
    
End Sub

'������� ������ � ��������
Sub Zerotasksdel(IndexTaskFirst, IndexTaskLast)
            
    Dim BiqTask As Task
    Dim BiqTaskSecond As Task
    TempZeroTaskID = 0
    TempPredec = 0
    KeyFoLoop = False
    For Each BiqTask In ActiveProject.Tasks '���� ������ ����� � ����� �����
        If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
            If BiqTask.GetField(FieldID:=projectField_Cost) = "0 �" Then
                TempZeroTaskID = BiqTask.id '�� ������ � ����� �����
                TempPredec = BiqTask.GetField(FieldID:=projectField_Predecessors) '�������������� ������ � ����� �����
                Call RepCycPred(TempZeroTaskID, TempPredec) '������� ������ � ��������
            End If
        End If
    Next BiqTask
    '�������� ���� ����� � ����� �����
    Call DeleteAllZeroTasks(IndexTaskFirst, IndexTaskLast)
    
End Sub

'������� ������ � �������� 2
Sub RepCycPred(TempZeroTaskID, TempPredec)

    Dim BiqTask As Task

    For Each BiqTask In ActiveProject.Tasks '���� ������ � ��������
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

'������� ��������� ������� ����������������
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

End Function

' �������� ������ � MS Project
Sub AddNewTask(MainTask, ByRef FirstTask, BiqStartDate, TaskJiraId, TaskType, TaskName, TaskHours, BIQTaskId, ToTaskDays, TaskTypeITService, TaskTypeWork, TaskActor, ByRef Index, ByRef IndexTaskFirst, ByRef IndexTaskLast)
    
    ' ������� ������
    If BIQTaskId = 0 Then
        Set NewTask = ActiveProject.Tasks.Add(TaskName)
    Else
        Set NewTask = ActiveProject.Tasks.Add(TaskName, BIQTaskId + Index)
    End If
    Index = Index + 1
    ' ��� ������� ������ ���������� ������ � �������
    If MainTask Then
        Do While NewTask.OutlineLevel > 1
            NewTask.OutlineOutdent
        Loop
    End If
    ' ��� ������ ������ ������ ������
    If FirstTask Then
        NewTask.OutlineIndent
    End If
    ' ��������� ����
    NewTask.SetField FieldID:=projectField_JirID, Value:=TaskJiraId
    '��������� ����,����������� �� ������ ������
    If ToTaskDays Then
        NewTask.SetField FieldID:=projectField_DurationDays, Value:=WorksheetFunction.RoundUp(((Val(TaskHours)) / 8), 0)
        ' ��� ������ ������ �������������� �� ���������
        If FirstTask Then
            NewTask.SetField FieldID:=projectField_Start, Value:=BiqStartDate
            IndexTaskFirst = NewTask.id '������ ������
        Else
            IndexTaskLast = NewTask.id '��������� ������
        End If
    End If

    ' ������ ��� ������ ������ ������ ������
    If FirstTask Then
        FirstTask = False
    End If
    
    '��������� ���� �������� ������
    NewTask.SetField FieldID:=projectField_ITService, Value:=TaskTypeITService
    NewTask.SetField FieldID:=projectField_Cost, Value:=TaskHours
    NewTask.SetField FieldID:=projectField_JiraProjName, Value:=TaskType
    NewTask.SetField FieldID:=projectField_TypeWork, Value:=TaskTypeWork
    NewTask.SetField FieldID:=projectField_Actor, Value:=TaskActor
    
End Sub

' �������� �������� ����� � MS Project
Sub InitFieldConst()

    projectField_Name = FieldNameToFieldConstant("��������", pjProject)
    projectField_JirID = FieldNameToFieldConstant("Jira id", pjProject)
    projectField_Cost = FieldNameToFieldConstant("������������", pjProject)
    projectField_Actor = FieldNameToFieldConstant("�������� ��������", pjProject)
    projectField_DurationDays = FieldNameToFieldConstant("������������", pjProject)
    projectField_Restrict = FieldNameToFieldConstant("��� �����������", pjProject)
    projectField_JiraProjName = FieldNameToFieldConstant("��� �������", pjProject)
    projectField_Predecessors = FieldNameToFieldConstant("���������������", pjProject)
    projectField_Start = FieldNameToFieldConstant("������", pjProject)
    projectField_End = FieldNameToFieldConstant("���������", pjProject)
    projectField_ITService = FieldNameToFieldConstant("��-������", pjProject)
    projectField_TypeWork = FieldNameToFieldConstant("��� �����", pjProject)
    projectField_Teg = FieldNameToFieldConstant("���", pjResource)
    projectField_ResGroup = FieldNameToFieldConstant("������", pjResource)
    projectField_ResGroupCk = FieldNameToFieldConstant("������ ��", pjResource)
    projectField_FuncArea1 = FieldNameToFieldConstant("����. ������� 1", pjResource)
    projectField_FuncArea2 = FieldNameToFieldConstant("����. ������� 2", pjResource)
    projectField_FuncArea3 = FieldNameToFieldConstant("����. ������� 3", pjResource)
    projectField_System1 = FieldNameToFieldConstant("������� 1", pjResource)
    projectField_System2 = FieldNameToFieldConstant("������� 2", pjResource)
  
End Sub

' ����� ������ �� ���
Private Sub GetExcelFileCFTButton_Click()
    FileNameCFTTextBox.Text = ShowGetOpenDialog()
End Sub

' ����� ������ �� ��������
Private Sub GetExcelFileBISButton_Click()
    FileNameBISTextBox.Text = ShowGetOpenDialog()
End Sub

'������� �������� ���������� ��� ������ �����
Public Function ShowGetOpenDialog() As String

    Dim xlObj As Excel.Application
    Dim fd As Office.FileDialog
    Set xlObj = New Excel.Application
    xlObj.Visible = False
    Set fd = xlObj.Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "�������� ����������� ����" '�������� ����������
        .Filters.Add "Excel", "*.xls,*.xlsx" '������� ��� ����������� ������
        .AllowMultiSelect = False            '������ ���� ����
        If .Show = False Then
            Set xlObj = Nothing
            Exit Function
        End If
        ShowGetOpenDialog = .SelectedItems(1) '������� ����������
    End With
    Set xlObj = Nothing
    
End Function

'�������� ���� ����� � ����� �����
Sub DeleteAllZeroTasks(IndexTaskFirst, IndexTaskLast)

    Dim BiqTask As Task
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_Cost) = "0 �" Then
            If (BiqTask.id >= IndexTaskFirst And BiqTask.id <= IndexTaskLast) Then
                BiqTask.Delete '�������� BIQ-������
                IndexTaskLast = IndexTaskLast - 1
            End If
        End If
    Next BiqTask

End Sub

'������ ��������
Private Sub DeleteButton_Click()

    Dim BiqTask As Task ' ��� ������ ������ �� BIQ
    InitFieldConst
    BIQNum = "BIQ-" + TBNumBIQFDelete '����� BIQ-������
    BIQTaskId = 0
    For Each BiqTask In ActiveProject.Tasks
        If BiqTask.GetField(FieldID:=projectField_JirID) = BIQNum Then
           BIQTaskId = BiqTask.id
           BiqTask.Delete '�������� BIQ-������
        End If
    Next BiqTask
    If BIQTaskId = 0 Then
        MsgBox ("����� BIQ-������ ���")
    End If
    
End Sub

' ������������ ������ � ����������� �� �������� ��������
Sub Perest()
      
    Set AllRes = ActiveProject.Resources
    For Each Res In AllRes
        Set resAssArr = Res.Assignments
        '1
        For Each resAss In resAssArr
            Set assTask = resAss.Task
            DurationWorkDaysPerest = assTask.DurationText '������������ � ������� ����
            StartDatePerest = Mid(Mid(assTask.StartText, 4), 1, 6) & "20" & Mid(assTask.StartText, 10) '���� ������ ������
            FinishDatePerest = Mid(Mid(assTask.FinishText, 4), 1, 6) & "20" & Mid(assTask.FinishText, 10) '���� ����� ������
            TimePerest = assTask.TimeScaleData(assTask.Start, assTask.Finish, TimescaleUnit:=4)(1).Value / (60) '�������� ����� � ����
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
