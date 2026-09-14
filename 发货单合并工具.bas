Attribute VB_Name = "合并发货单工具"
Option Explicit

'==============================================================
'  预制构件发货单一键合并工具
'  功能：选择文件夹后，自动把里面所有发货单 xlsx 合并成
'        一张 14 列的《发货明细汇总》表（含"发货时间"列）。
'  规则：
'    1. 跳过"构件编码 / PK板构件编码"等主表工作表；
'    2. 只处理行 1-8 含"项目名称/客户姓名/运输车号/出库总量"
'       标记的发货单工作表；
'    3. 列按表头文本自动匹配（板宽/板长/板厚/砼标号等顺序
'       不同也能识别），缺的列留空；
'    4. 发货时间取标题行日期（如"成品出库单丨2026年3月8日…"），
'       取不到时用"打印时间"单元格；
'    5. 自动跳过合计行、空行。
'  用法：Alt+F11 打开 VBE → 文件→导入文件 选本 .bas →
'        Alt+F8 运行宏「合并发货单」。
'==============================================================

Public Sub 合并发货单()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "选择发货单所在文件夹"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If

    ' ---- 收集待处理文件（排除输出文件和临时文件）----
    Dim files As Collection
    Set files = New Collection
    Dim f As String
    f = Dir(folderPath & "*.xlsx")
    Do While f <> ""
        If Left$(f, 2) <> "~$" And InStr(f, "合并") = 0 And InStr(f, "汇总") = 0 Then
            files.Add f
        End If
        f = Dir
    Loop
    If files.Count = 0 Then
        MsgBox "该文件夹里没有找到可处理的发货单文件。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' ---- 新建输出工作簿 ----
    Dim outWb As Workbook, outWs As Worksheet
    Set outWb = Workbooks.Add(xlWBATWorksheet)
    Set outWs = outWb.Worksheets(1)
    outWs.Name = "发货明细汇总"

    Dim headers As Variant
    headers = Array("序号", "发货时间", "构件编号", "楼栋", "楼层", "构件类型", _
                    "单块体积", "单体质量", "单块面积", _
                    "板宽(mm)", "板长(mm)", "板厚(mm)", "砼标号", "备注")
    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        outWs.Cells(1, c + 1).Value = headers(c)
    Next c

    ' ---- 逐个文件处理 ----
    Dim outRow As Long
    outRow = 2
    Dim i As Long, wb As Workbook
    For i = 1 To files.Count
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(folderPath & files(i), ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        If Not wb Is Nothing Then
            outRow = ProcessWorkbook(wb, outWs, outRow)
            wb.Close SaveChanges:=False
        End If
    Next i

    ' ---- 格式化 ----
    With outWs
        With .Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
        End With
        .Range("A1:N1").HorizontalAlignment = xlCenter
        If outRow > 2 Then
            .Range("B2:B" & outRow).NumberFormat = "yyyy-mm-dd"
            .Range("A1:N" & outRow).AutoFilter  ' 加筛选，方便按楼栋/日期筛
        End If
        .Columns.AutoFit
        .Rows(1).RowHeight = 20
        .Activate
        .Range("A2").Select
    End With

    ' ---- 保存 ----
    Dim outPath As String
    outPath = folderPath & "发货明细汇总_生成.xlsx"
    On Error Resume Next
    Application.DisplayAlerts = False
    outWb.SaveAs Filename:=outPath, FileFormat:=51   ' 51 = xlsx
    Dim savedOk As Boolean
    savedOk = (Err.Number = 0)
    On Error GoTo 0

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "合并完成！" & vbCrLf & vbCrLf & _
          "处理文件数：" & files.Count & " 个" & vbCrLf & _
          "构件明细　：" & (outRow - 2) & " 条" & vbCrLf
    If savedOk Then
        msg = msg & "已保存至：" & outPath
    Else
        msg = msg & "自动保存失败（文件可能被占用），请手动另存。"
    End If
    MsgBox msg, vbInformation, "合并发货单"
End Sub

'--------------------------------------------------------------
' 处理一个工作簿：跳过主表，逐张导入发货单工作表
'--------------------------------------------------------------
Private Function ProcessWorkbook(wb As Workbook, outWs As Worksheet, ByVal startRow As Long) As Long
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If InStr(ws.Name, "构件编码") = 0 Then          ' 主表（含 PK板构件编码）跳过
            If IsDeliverySheet(ws) Then
                startRow = ImportSheet(ws, outWs, startRow)
            End If
        End If
    Next ws
    ProcessWorkbook = startRow
End Function

'--------------------------------------------------------------
' 判断是否发货单工作表：行 1-8 是否有发货单标记
'--------------------------------------------------------------
Private Function IsDeliverySheet(ws As Worksheet) As Boolean
    Dim r As Long, c As Long, s As String
    For r = 1 To 8
        For c = 1 To 20
            s = CStr(ws.Cells(r, c).Value)
            If InStr(s, "项目名称") > 0 Or InStr(s, "客户姓名") > 0 _
               Or InStr(s, "运输车号") > 0 Or InStr(s, "出库总量") > 0 Then
                IsDeliverySheet = True
                Exit Function
            End If
        Next c
    Next r
End Function

'--------------------------------------------------------------
' 导入一张发货单工作表
'--------------------------------------------------------------
Private Function ImportSheet(ws As Worksheet, outWs As Worksheet, ByVal startRow As Long) As Long
    ImportSheet = startRow

    ' 1) 发货时间：标题行日期，回退到"打印时间"
    Dim sendDate As Variant
    sendDate = ParseSendDate(ws)

    ' 2) 定位列表头行：真正列表头通常在第 8-15 行，
    '    且同一行必须同时出现"构件编号"和"楼栋"等关键字。
    '    注意：源表单元格里"构件编号"可能是"构件"+换行+"编号"，
    '    所以判断时要去掉换行，并同时检查"构件"和"编号"两个字。
    Dim headerRow As Long, headerCol As Long
    headerRow = 0
    Dim r As Long, c As Long
    Dim hasCode As Boolean, hasBuilding As Boolean
    Dim cellText As String
    For r = 8 To 15
        hasCode = False
        hasBuilding = False
        For c = 1 To 30
            cellText = Replace(CStr(ws.Cells(r, c).Value), vbLf, "")
            cellText = Replace(cellText, vbCr, "")
            If InStr(cellText, "构件") > 0 And InStr(cellText, "编号") > 0 Then hasCode = True
            If InStr(cellText, "楼栋") > 0 Then hasBuilding = True
        Next c
        If hasCode And hasBuilding Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Function

    ' 3) 按表头文本建立 列名 -> 源列号 映射
    Dim colIdx As Object
    Set colIdx = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long, h As String, key As Variant
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol > 60 Then lastCol = 60
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If Len(h) > 0 Then
            For Each key In Array("序号", "构件编号", "楼栋", "楼层", "构件类型", _
                                  "单块体积", "单体质量", "单块面积", _
                                  "板宽", "板长", "板厚", "砼标号", "备注")
                If Not colIdx.Exists(key) Then
                    If HeaderMatches(h, CStr(key)) Then
                        colIdx.Add key, c
                        Exit For
                    End If
                End If
            Next key
        End If
    Next c
    If Not colIdx.Exists("构件编号") Then Exit Function

    ' 4) 逐行读取明细（跳过空行和合计行）
    Dim lastRow As Long, codeVal As String
    lastRow = ws.Cells(ws.Rows.Count, colIdx("构件编号")).End(xlUp).Row
    Dim outCol As Long
    For r = headerRow + 1 To lastRow
        codeVal = Trim$(CStr(ws.Cells(r, colIdx("构件编号")).Value))
        If Len(codeVal) > 0 And InStr(codeVal, "合计") = 0 Then
            outWs.Cells(startRow, 1).Value = SafeCell(ws, r, colIdx, "序号")
            outWs.Cells(startRow, 2).Value = sendDate
            outWs.Cells(startRow, 3).Value = codeVal
            outWs.Cells(startRow, 4).Value = SafeCell(ws, r, colIdx, "楼栋")
            outWs.Cells(startRow, 5).Value = SafeCell(ws, r, colIdx, "楼层")
            outWs.Cells(startRow, 6).Value = SafeCell(ws, r, colIdx, "构件类型")
            outWs.Cells(startRow, 7).Value = SafeCell(ws, r, colIdx, "单块体积")
            outWs.Cells(startRow, 8).Value = SafeCell(ws, r, colIdx, "单体质量")
            outWs.Cells(startRow, 9).Value = SafeCell(ws, r, colIdx, "单块面积")
            outWs.Cells(startRow, 10).Value = SafeCell(ws, r, colIdx, "板宽")
            outWs.Cells(startRow, 11).Value = SafeCell(ws, r, colIdx, "板长")
            outWs.Cells(startRow, 12).Value = SafeCell(ws, r, colIdx, "板厚")
            outWs.Cells(startRow, 13).Value = SafeCell(ws, r, colIdx, "砼标号")
            outWs.Cells(startRow, 14).Value = SafeCell(ws, r, colIdx, "备注")
            startRow = startRow + 1
        End If
    Next r

    ImportSheet = startRow
End Function

'--------------------------------------------------------------
' 表头文本匹配（兼容"板宽(mm)""板厚（mm）""单块体积(m3)"等写法）
'--------------------------------------------------------------
Private Function HeaderMatches(ByVal headerText As String, ByVal target As String) As Boolean
    ' 去掉换行，兼容"构件\n编号"这种单元格内换行
    Dim h As String
    h = Replace(Replace(headerText, vbLf, ""), vbCr, "")
    Select Case target
        Case "序号":     HeaderMatches = (h = "序号")
        Case "构件编号": HeaderMatches = (InStr(h, "构件") > 0 And InStr(h, "编号") > 0)
        Case "楼栋":     HeaderMatches = (InStr(h, "楼栋") > 0)
        Case "楼层":     HeaderMatches = (InStr(h, "楼层") > 0)
        Case "构件类型": HeaderMatches = (InStr(h, "构件类型") > 0)
        Case "单块体积": HeaderMatches = (InStr(h, "体积") > 0)
        Case "单体质量": HeaderMatches = (InStr(h, "质量") > 0)
        Case "单块面积": HeaderMatches = (InStr(h, "面积") > 0)
        Case "板宽":     HeaderMatches = (InStr(h, "板宽") > 0 Or InStr(h, "宽度") > 0)
        Case "板长":     HeaderMatches = (InStr(h, "板长") > 0 Or InStr(h, "长度") > 0)
        Case "板厚":     HeaderMatches = (InStr(h, "板厚") > 0 Or InStr(h, "厚度") > 0)
        Case "砼标号":   HeaderMatches = (InStr(h, "砼") > 0 Or InStr(h, "标号") > 0)
        Case "备注":     HeaderMatches = (h = "备注")
    End Select
End Function

'--------------------------------------------------------------
' 取单元格值（列不存在时返回空）
'--------------------------------------------------------------
Private Function SafeCell(ws As Worksheet, ByVal r As Long, colIdx As Object, ByVal key As String) As Variant
    If colIdx.Exists(key) Then
        SafeCell = ws.Cells(r, colIdx(key)).Value
    Else
        SafeCell = Empty
    End If
End Function

'--------------------------------------------------------------
' 解析发货时间：标题行"…2026年3月8日…"；回退"打印时间"单元格
'--------------------------------------------------------------
Private Function ParseSendDate(ws As Worksheet) As Variant
    Dim s As String
    s = CStr(ws.Cells(1, 1).Value)

    Dim pY As Long, pM As Long, pD As Long
    pY = InStr(s, "年")
    If pY >= 5 Then
        Dim y As String, m As String, d As String
        y = Mid$(s, pY - 4, 4)
        pM = InStr(pY, s, "月")
        If pM > 0 Then pD = InStr(pM, s, "日")
        If pM > 0 And pD > 0 Then
            m = Mid$(s, pY + 1, pM - pY - 1)
            d = Mid$(s, pM + 1, pD - pM - 1)
            If IsNumeric(y) And IsNumeric(m) And IsNumeric(d) Then
                ParseSendDate = DateSerial(CLng(y), CLng(m), CLng(d))
                Exit Function
            End If
        End If
    End If

    ' 回退：找"打印时间"标签，取右边一格
    Dim r As Long, c As Long
    For r = 1 To 8
        For c = 1 To 20
            If InStr(CStr(ws.Cells(r, c).Value), "打印时间") > 0 Then
                ParseSendDate = ws.Cells(r, c + 1).Value
                Exit Function
            End If
        Next c
    Next r

    ParseSendDate = Empty
End Function
