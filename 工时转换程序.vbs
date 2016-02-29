'Option Explicit

Sub Translate(source, destine)
    Dim app,workbook,sheet, oworkbook, osheet
    Set app = WScript.CreateObject("Excel.Application")
    app.Visible = True
    Set workbook = app.WorkBooks.open(source)
    Set oworkbook = app.WorkBooks.Add
    
    Set sheet = workbook.Worksheets(1)
    Set osheet = oworkbook.Worksheets(1)
    osheet.cells(1,1).value = "姓名"
    osheet.cells(1,2).value = "日期"
    osheet.cells(1,3).value = "上班时间"
    osheet.cells(1,4).value = "下班时间"
    osheet.cells(1,5).value = "出勤时间(精确)"
    osheet.cells(1,6).value = "加班时间(精确)"
    osheet.cells(1,7).value = "出勤时间"
    osheet.cells(1,8).value = "加班时间"
    osheet.cells(1,9).value = "迟到时间"
    osheet.cells(1,10).value = "早退时间"
    osheet.cells(1,11).value = "需要乐捐"
    osheet.cells(1,12).value = "需要请假"
    osheet.cells(1,13).value = "餐补"
    osheet.Columns("B:B").ColumnWidth = 12
    osheet.Columns("E:E").ColumnWidth = 14
    osheet.Columns("F:F").ColumnWidth = 14
    'osheet.Range("A2").Select
    'osheet.ActiveWindow.FreezePanes = True

    
    'Set cell = sheet.cells(row, 2).CurrentRegion
    'Set cell = sheet.cells(2, 4)
    'Set cell1 = sheet.cells(7, 4)
    'wsh.echo(datediff("s", cell.value, cell1.value)/3600.0)
    
    Dim person, row, cell, date1, date2, date0, orow
    row = 2
    Set cell = sheet.cells(row, 2)
    
    person = Null
    date1 = Null
    date2 = Null
    date0 = Null
    orow = 2
    
    
    DO Until isEmpty(cell)
        if isNull(person) Or cell.value <> person then
            if Not(isNull(person)) and Not(isNull(date1)) then
		if (isNull(date2)) then date2 = date1 end if
                Call AddOneRecord(person, date1, date2, osheet, orow)
                orow = orow + 1
            end if
            person = cell.value
            date1 = sheet.cells(row, 4).value
            date2 = Null
        else
            date0 = sheet.cells(row, 4).value
            if year(date0) = year(date1) and month(date0) = month(date1) and day(date0) = day(date1) then
                date2 = date0
            else
                if not(isNull(date1)) then
                    if (isNull(date2)) then date2 = date1 end if
                    Call AddOneRecord(person, date1, date2, osheet, orow)
                    orow = orow + 1
                end if
                date1 = date0
                date2 = Null
            end if
        end if
        row = row + 1
        set cell = sheet.cells(row, 2)
    Loop
    
    oworkbook.saveas(destine)
    workbook.Close()
    oworkbook.Close()
    app.Quit()
End Sub

Sub AddOneRecord(person, date1, date2, osheet, orow)
    '姓名
    osheet.cells(orow, 1).value = person
    '日期
    osheet.cells(orow, 2).value = year(date1) & "-" & month(date1) & "-" & day(date1)
    '上班时间
    osheet.cells(orow, 3).value = hour(date1) & ":" & minute(date1) & ":" & second(date1)
    osheet.cells(orow, 3).NumberFormatLocal = "h:mm"
    '下班时间
    osheet.cells(orow, 4).value = hour(date2) & ":" & minute(date2) & ":" & second(date2)
    osheet.cells(orow, 4).NumberFormatLocal = "h:mm"
    '出勤时间(精确)
    osheet.cells(orow, 5).FormulaR1C1 = _
        "=MAX(0, MIN(RC[-1], TIME(17, 30, 0)) - MAX(RC[-2], TIME(8, 30, 0)) - IF(OR(RC[-1]<TIME(12,0,0), RC[-2]>TIME(13,0,0)), 0, TIME(1,0,0)))"
    osheet.cells(orow, 5).NumberFormatLocal = "h:mm"
    '加班时间(精确)：从18:30开始算加班
    osheet.cells(orow, 6).FormulaR1C1 = "=MAX(0,RC[-2]-TIME(18,30,0))"
    osheet.cells(orow, 6).NumberFormatLocal = "h:mm"
    '出勤时间
    osheet.cells(orow, 7).FormulaR1C1 = _
        "=HOUR(RC[-2])+IF(MINUTE(RC[-2])<15, 0, IF(MINUTE(RC[-2])>=45, 1, 0.5))"
    osheet.cells(orow, 7).NumberFormatLocal = "0.0_);(0.0)"
    '加班时间
    osheet.cells(orow, 8).FormulaR1C1 = "=HOUR(RC[-2])+IF(MINUTE(RC[-2])<15, 0, IF(MINUTE(RC[-2])>=45, 1, 0.5))"
    osheet.cells(orow, 8).NumberFormatLocal = "0.0_);(0.0)"
    '迟到时间
    osheet.cells(orow, 9).FormulaR1C1 = "=MAX(0, RC[-6]-TIME(8,30,0))"
    osheet.cells(orow, 9).NumberFormatLocal = "h:mm"
    '早退时间：提前下班时间
    osheet.cells(orow, 10).FormulaR1C1 = "=MAX(0, TIME(17,30,0)-RC[-6])"
    osheet.cells(orow, 10).NumberFormatLocal = "h:mm"
    '是否需要乐捐：早上8:30~9:00间到公司需要乐捐
    osheet.cells(orow, 11).FormulaR1C1 = _
        "=IF(AND(RC[-8]>TIME(8,30,59),RC[-8]<TIME(9,0,0)), ""是"","""")"
    '是否需要请假：早上9点后到公司，或17点前下班，需要请假
    osheet.cells(orow, 12).FormulaR1C1 = "=IF(OR(RC[-9]>TIME(9,0,0),RC[-8]<TIME(17,0,0)),""是"","""")"
    '餐补: 上班时间超过4小时，算1次餐补，加班时间超过1小时，再加一次餐补
    osheet.cells(orow, 13).FormulaR1C1 = _
        "=IF(RC[-6]<4,0,1)+IF(RC[-9]>TIME(19,30,0),1,0)"

End Sub

'Set fileDialog = WScript.CreateObject("SAFRCFileDlg.FileSave")
'fileDialog.FileType = ".xlsx"
'if fileDialog.OpenFileSaveDlg then
'    wsh.echo(fileDialog.FileName)
'end if

'
' Description: VBScript/VBS open file dialog
'              Compatible with most Windows platforms
' Author: wangye  <pcn88 at hotmail dot com>
' Website: http://wangye.org
'
' dir is the initial directory; if no directory is
' specified "Desktop" is used.
' filter is the file type filter; format "File type description|*.ext"
'
Public Function GetOpenFileName(dir, filter)
    Const msoFileDialogFilePicker = 3
 
    If VarType(dir) <> vbString Or dir="" Then
        dir = CreateObject( "WScript.Shell" ).SpecialFolders( "Desktop" )
    End If
 
    If VarType(filter) <> vbString Or filter="" Then
        filter = "All files|*.*"
    End If
 
    Dim i,j, objDialog, TryObjectNames
    TryObjectNames = Array( _
        "UserAccounts.CommonDialog", _
        "MSComDlg.CommonDialog", _
        "MSComDlg.CommonDialog.1", _
        "Word.Application", _
        "SAFRCFileDlg.FileOpen", _
        "InternetExplorer.Application" _
        )
 
    On Error Resume Next
    Err.Clear
 
    For i=0 To UBound(TryObjectNames)
        Set objDialog = WSH.CreateObject(TryObjectNames(i))
        If Err.Number<>0 Then
        Err.Clear
        Else
        Exit For
        End If
    Next
 
    Select Case i
        Case 0,1,2
        ' 0. UserAccounts.CommonDialog XP Only.
        ' 1.2. MSComDlg.CommonDialog MSCOMDLG32.OCX must registered.
        If i=0 Then
            objDialog.InitialDir = dir
        Else
            objDialog.InitDir = dir
        End If
        objDialog.Filter = filter
        If objDialog.ShowOpen Then
            GetOpenFileName = objDialog.FileName
        End If
        Case 3
        ' 3. Word.Application Microsoft Office must installed.
        objDialog.Visible = False
        Dim objOpenDialog, filtersInArray
        filtersInArray = Split(filter, "|")
        Set objOpenDialog = _
            objDialog.Application.FileDialog( _
                msoFileDialogFilePicker)
            With objOpenDialog
            .Title = "Open File(s):"
            .AllowMultiSelect = False
            .InitialFileName = dir
            .Filters.Clear
            For j=0 To UBound(filtersInArray) Step 2
                .Filters.Add filtersInArray(j), _
                     filtersInArray(j+1), 1
            Next
            If .Show And .SelectedItems.Count>0 Then
                GetOpenFileName = .SelectedItems(1)
            End If
            End With
            objDialog.Visible = True
            objDialog.Quit
        Set objOpenDialog = Nothing
        Case 4
        ' 4. SAFRCFileDlg.FileOpen xp 2003 only
        ' See http://www.robvanderwoude.com/vbstech_ui_fileopen.php
        If objDialog.OpenFileOpenDlg Then
           GetOpenFileName = objDialog.FileName
        End If
        Case 5
        ' 5. InternetExplorer.Application IE must installed
        objDialog.Navigate "about:blank"
        Dim objBody, objFileDialog
        Set objBody = _
            objDialog.document.getElementsByTagName("body")(0)
        objBody.innerHTML = "<input type='file' id='fileDialog'>"
        while objDialog.Busy Or objDialog.ReadyState <> 4
            WScript.sleep 10
        Wend
        Set objFileDialog = objDialog.document.all.fileDialog
            objFileDialog.click
            GetOpenFileName = objFileDialog.value
            objDialog.Quit
        Set objFileDialog = Nothing
        Set objBody = Nothing
        Case Else
        ' Sorry I cannot do that!
    End Select
 
    Set objDialog = Nothing
End Function
 
Dim strFileName, path, ext, pos
strFileName = GetOpenFileName("C:\","All files|*.*|Microsoft Excel|*.xlsx,*.xls")
    if not(isEmpty(strFileName)) then
    pos = inStrRev(strFileName, ".", len(strFileName))
    path = left(strFileName, pos-1)
    ext = right(strFileName, len(strFileName)-pos+1)
    Call Translate(strFileName, path & "-result" & ext)
end if

' Test
'Call Translate("d:\work\其他\工时统计脚本\考勤12.xlsx", "d:\work\其他\工时统计脚本\考勤123.xlsx")
