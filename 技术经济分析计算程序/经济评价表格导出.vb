Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Windows.Forms

Public Class 经济评价表格导出
    Private Sub 确定_Click(sender As Object, e As EventArgs) Handles 确定.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim mainprogram As New Com技术经济分析计算程序
        Me.Hide() '隐藏窗体
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim WordAppGetname As Word.Application '定义Word对象
        WordAppGetname = GetObject(, "Word.Application")    '当前Word对象赋值给WordApp
        Dim Wordname As String = WordAppGetname.ActiveDocument.Name
        If Wordname = Nothing Then
            '清除全部表格的批注
            If Com技术经济分析计算程序.Form11.Checkbox1.Checked = True Then
                Call mainprogram.解锁表格()
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells.Select
                ExcelApp.Selection.ClearComments
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells.Select
                ExcelApp.Selection.ClearComments
                Call mainprogram.锁定表格()
                '激活表格
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '获取当前打开的Excel文件的路径和文件名
            Dim FilePath As String = ExcelApp.ThisWorkbook.Path '路径
            Dim FileName As String = Strings.Left(ExcelApp.ActiveWorkbook.Name, Len(ExcelApp.ActiveWorkbook.Name) - 5) '文件名（不包含拓展名）
            '创建Word
            Dim WordDoc = CreateObject("Word.application") '创建Word
            Dim WordApp As Word.Application '定义Word对象
            WordApp = GetObject(, "Word.Application")    '当前Word对象赋值给WordApp
            WordDoc.Documents.Add '打开一份文档
            WordDoc.Visible = True '汇话状态打开
            '读取投资各方出资比例
            '从投资方1到5表格中读取实际设置的参数
            '出资比例
            Dim czbl1 = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(4, 38).Value * 100
            Dim czbl2 = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(4, 38).Value * 100
            Dim czbl3 = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(4, 38).Value * 100
            Dim czbl4 = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(4, 38).Value * 100
            Dim czbl5 = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(4, 38).Value * 100
            '资产处置比例
            Dim zcczbl1 = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(5, 38).Value * 100
            Dim zcczbl2 = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(5, 38).Value * 100
            Dim zcczbl3 = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(5, 38).Value * 100
            Dim zcczbl4 = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(5, 38).Value * 100
            Dim zcczbl5 = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(5, 38).Value * 100
            '利润分配比例
            Dim lrfpbl1 = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(6, 38).Value * 100
            Dim lrfpbl2 = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(6, 38).Value * 100
            Dim lrfpbl3 = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(6, 38).Value * 100
            Dim lrfpbl4 = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(6, 38).Value * 100
            Dim lrfpbl5 = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(6, 38).Value * 100
            '确定当前的投资方数量
            Dim TZFJS As Integer = 0
            Dim TZFJS_1 As Integer = 0
            Dim TZFJS_2 As Integer = 0
            Dim TZFJS_3 As Integer = 0
            '求最大值的过程量
            Dim TZFJS_a As Integer = 0
            '根据出资比例系数计算投资方数量
            If czbl1 > 0 Then
                TZFJS_1 = TZFJS_1 + 1
            End If
            If czbl2 > 0 Then
                TZFJS_1 = TZFJS_1 + 1
            End If
            If czbl3 > 0 Then
                TZFJS_1 = TZFJS_1 + 1
            End If
            If czbl4 > 0 Then
                TZFJS_1 = TZFJS_1 + 1
            End If
            If czbl5 > 0 Then
                TZFJS_1 = TZFJS_1 + 1
            End If
            '根据资产处置比例计算投资方数量
            If zcczbl1 > 0 Then
                TZFJS_2 = TZFJS_2 + 1
            End If
            If zcczbl2 > 0 Then
                TZFJS_2 = TZFJS_2 + 1
            End If
            If zcczbl3 > 0 Then
                TZFJS_2 = TZFJS_2 + 1
            End If
            If zcczbl4 > 0 Then
                TZFJS_2 = TZFJS_2 + 1
            End If
            If zcczbl5 > 0 Then
                TZFJS_2 = TZFJS_2 + 1
            End If
            '根据利润分配比例计算投资方数量
            If lrfpbl1 > 0 Then
                TZFJS_3 = TZFJS_3 + 1
            End If
            If lrfpbl2 > 0 Then
                TZFJS_3 = TZFJS_3 + 1
            End If
            If lrfpbl3 > 0 Then
                TZFJS_3 = TZFJS_3 + 1
            End If
            If lrfpbl4 > 0 Then
                TZFJS_3 = TZFJS_3 + 1
            End If
            If lrfpbl5 > 0 Then
                TZFJS_3 = TZFJS_3 + 1
            End If
            '确定最终的投资方数量结果
            TZFJS_a = Math.Max(TZFJS_1, TZFJS_2) '过程量
            TZFJS = Math.Max(TZFJS_a, TZFJS_3) '最终结果
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '导出所有的经济评价详细表格
            If Com技术经济分析计算程序.Form11.CheckBox2.Checked = True Then
                If WordDoc.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait Then '将word设置为横向
                    WordDoc.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
                Else
                    WordDoc.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait
                End If
                With WordDoc.Selection.PageSetup 'Word纸张大小设置为A3
                    .LineNumbering.Active = False
                    .Orientation = Word.WdOrientation.wdOrientLandscape
                    .TopMargin = WordApp.Application.CentimetersToPoints(3.17)
                    .BottomMargin = WordApp.Application.CentimetersToPoints(3.17)
                    .LeftMargin = WordApp.Application.CentimetersToPoints(2.54)
                    .RightMargin = WordApp.Application.CentimetersToPoints(2.54)
                    .Gutter = WordApp.Application.CentimetersToPoints(0)
                    .HeaderDistance = WordApp.Application.CentimetersToPoints(1.5)
                    .FooterDistance = WordApp.Application.CentimetersToPoints(1.75)
                    .PageWidth = WordApp.Application.CentimetersToPoints(42)
                    .PageHeight = WordApp.Application.CentimetersToPoints(29.7)
                    .FirstPageTray = Word.WdPaperTray.wdPrinterDefaultBin
                    .OtherPagesTray = Word.WdPaperTray.wdPrinterDefaultBin
                    .SectionStart = Word.WdSectionStart.wdSectionNewPage
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop
                    .SuppressEndnotes = False
                    .MirrorMargins = False
                    .TwoPagesOnOne = False
                    .BookFoldPrinting = False
                    .BookFoldRevPrinting = False
                    .BookFoldPrintingSheets = 1
                    .GutterPos = Word.WdGutterStyle.wdGutterPosLeft
                    .LinesPage = 26
                    .LayoutMode = Word.WdLayoutMode.wdLayoutModeLineGrid
                End With
                WordApp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100 '页面缩放比例设置为100%
                '————————————————————————————————————————————————————————————————————————————————————————
                '读取输入的项目计算年限
                Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                '————————————————————————————————————————————————————————————————————————————————————————
                '投资计划与资金筹措表
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Range("A1:S20").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Range("A22:S41").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '流动资金估算表
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Range("A1:T17").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Range("A19:T35").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '项目投资现金流量表
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Range("A1:S20").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Range("A22:S43").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '资本金现金流量表
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Range("A1:S25").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Range("A27:S53").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '收入税收表
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Range("A1:S22").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Range("A24:S45").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '总成本表
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Range("A1:S26").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Range("A33:S58").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '折旧摊销表
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Range("A1:S16").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Range("A18:S33").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '利润与利润分配表
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Range("A1:S19").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Range("A21:S39").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '借款还本付息计划表
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Range("A1:S29").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Range("A31:S59").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '财务计划现金流量表
                ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Range("A1:S32").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Range("A34:S65").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '资产负债表
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Range("A1:S28").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("资产负债表").Range("A30:S57").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("资产负债表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '经济增加值计算表
                ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Activate '表格激活
                '从Excel中复制为图片
                ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Range("A1:S18").Select
                ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                If jsnx > 15 Then
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Range("A20:S37").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Range("A1").Select
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                If czbl1 > 0 Or zcczbl1 > 0 Or lrfpbl1 > 0 Then
                    '投资方1现金流量表
                    ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Activate '表格激活
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Range("A1:S16").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                    If jsnx > 15 Then
                        '从Excel中复制为图片
                        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Range("A18:S35").Select
                        ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                        '粘贴到Word
                        WordDoc.Selection.Paste
                        WordDoc.Selection.TypeParagraph
                        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Range("A1").Select
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                If czbl2 > 0 Or zcczbl2 > 0 Or lrfpbl2 > 0 Then
                    '投资方2现金流量表
                    ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Activate '表格激活
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Range("A1:S16").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                    If jsnx > 15 Then
                        '从Excel中复制为图片
                        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Range("A18:S35").Select
                        ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                        '粘贴到Word
                        WordDoc.Selection.Paste
                        WordDoc.Selection.TypeParagraph
                        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Range("A1").Select
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                If czbl3 > 0 Or zcczbl3 > 0 Or lrfpbl3 > 0 Then
                    '投资方3现金流量表
                    ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Activate '表格激活
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Range("A1:S16").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                    If jsnx > 15 Then
                        '从Excel中复制为图片
                        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Range("A18:S35").Select
                        ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                        '粘贴到Word
                        WordDoc.Selection.Paste
                        WordDoc.Selection.TypeParagraph
                        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Range("A1").Select
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                If czbl4 > 0 Or zcczbl4 > 0 Or lrfpbl4 > 0 Then
                    '投资方4现金流量表
                    ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Activate '表格激活
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Range("A1:S16").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                    If jsnx > 15 Then
                        '从Excel中复制为图片
                        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Range("A18:S35").Select
                        ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                        '粘贴到Word
                        WordDoc.Selection.Paste
                        WordDoc.Selection.TypeParagraph
                        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Range("A1").Select
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                If czbl5 > 0 Or zcczbl5 > 0 Or lrfpbl5 > 0 Then
                    '投资方5现金流量表
                    ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Activate '表格激活
                    '从Excel中复制为图片
                    ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Range("A1:S16").Select
                    ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                    '根据计算年限不同，复制不同范围的表格，如果计算年限大于15年，则复制上下两个表格，否则只复制上面一个表格。
                    If jsnx > 15 Then
                        '从Excel中复制为图片
                        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Range("A18:S35").Select
                        ExcelApp.Selection.CopyPicture（Appearance:=XlPictureAppearance.xlScreen, Format:=XlCopyPictureFormat.xlPicture）
                        '粘贴到Word
                        WordDoc.Selection.Paste
                        WordDoc.Selection.TypeParagraph
                        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Range("A1").Select
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '插入下一页分节符
                WordDoc.Selection.InsertBreak(Type:=Word.WdBreakType.wdSectionBreakNextPage)
                '激活表格
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '导出财务评价指标一览表
            If Com技术经济分析计算程序.Form11.CheckBox5.Checked = True Then
                '纸张设置为A4，纵向
                With WordDoc.Selection.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = Word.WdOrientation.wdOrientLandscape
                    .TopMargin = WordApp.Application.CentimetersToPoints(3.17)
                    .BottomMargin = WordApp.Application.CentimetersToPoints(3.17)
                    .LeftMargin = WordApp.Application.CentimetersToPoints(2.54)
                    .RightMargin = WordApp.Application.CentimetersToPoints(2.54)
                    .Gutter = WordApp.Application.CentimetersToPoints(0)
                    .HeaderDistance = WordApp.Application.CentimetersToPoints(1.5)
                    .FooterDistance = WordApp.Application.CentimetersToPoints(1.75)
                    .PageWidth = WordApp.Application.CentimetersToPoints(29.7)
                    .PageHeight = WordApp.Application.CentimetersToPoints(21)
                    .FirstPageTray = Word.WdPaperTray.wdPrinterDefaultBin
                    .OtherPagesTray = Word.WdPaperTray.wdPrinterDefaultBin
                    .SectionStart = Word.WdSectionStart.wdSectionNewPage
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop
                    .SuppressEndnotes = False
                    .MirrorMargins = False
                    .TwoPagesOnOne = False
                    .BookFoldPrinting = False
                    .BookFoldRevPrinting = False
                    .BookFoldPrintingSheets = 1
                    .GutterPos = Word.WdGutterStyle.wdGutterPosLeft
                    .LinesPage = 26
                    .LayoutMode = Word.WdLayoutMode.wdLayoutModeLineGrid
                End With
                If WordDoc.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait Then '纸张设置为A4，纵向
                    WordDoc.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
                Else
                    WordDoc.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '导出财务分析指标数据一览表
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("指标数据").unProtect(Password:="wscjc") '解锁表格
                '财务评价指标一览表
                '根据投资方数量复制不同的范围
                If TZFJS = 1 Then
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B5:E39").Select
                    ExcelApp.Selection.Copy
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                End If
                If TZFJS = 2 Then
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B5:E40").Select
                    ExcelApp.Selection.Copy
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                End If
                If TZFJS = 3 Then
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B5:E41").Select
                    ExcelApp.Selection.Copy
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                End If
                If TZFJS = 4 Then
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B5:E42").Select
                    ExcelApp.Selection.Copy
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                End If
                If TZFJS = 5 Then
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B5:E43").Select
                    ExcelApp.Selection.Copy
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '插入下一页分节符
                WordDoc.Selection.InsertBreak(Type:=Word.WdBreakType.wdSectionBreakNextPage)
                '激活表格
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '导出敏感性分析表和敏感性分析图
            If Com技术经济分析计算程序.Form11.CheckBox3.Checked = True Then
                '指标数据表
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("指标数据").unProtect(Password:="wscjc") '解锁表格
                '敏感性分析表格
                '年运行小时数黄底色去除
                If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 8) > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("H137:H141").Select
                    With ExcelApp.Selection.Interior
                        .Pattern = XlPattern.xlPatternNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
                For i = 7 To 141 '隐藏为空的行
                    If (ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 12).Value = 0) Then
                        ExcelApp.ThisWorkbook.Worksheets("指标数据").Rows(i).EntireRow.Hidden = True
                    End If
                Next
                '从Excel中复制
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("G5:M141").Select
                ExcelApp.Selection.Copy()
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '取消指标数据表中的全部隐藏
                For i = 7 To 141 '隐藏为空的行
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Rows(i).EntireRow.Hidden = False
                Next
                '复制敏感性分析折线图
                ExcelApp.ActiveSheet.ChartObjects("单因素敏感性分析图").Activate
                ExcelApp.ActiveChart.ChartArea.Copy()
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '————————————————————————————————————————————————————————————————————————————————————————
                '插入下一页分节符
                WordDoc.Selection.InsertBreak(Type:=Word.WdBreakType.wdSectionBreakNextPage)
                '激活表格
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '导出盈亏平衡分析表和盈亏平衡分析图
            If Com技术经济分析计算程序.Form11.CheckBox4.Checked = True Then
                '指标数据表
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Activate '表格激活
                ExcelApp.ThisWorkbook.Worksheets("指标数据").unProtect(Password:="wscjc") '解锁表格
                '复制盈亏平衡分析表（根据选择的模式自动选择）
                If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(62, 4).Value = "总成本最大年份" Then
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B54:E60").Select
                    ExcelApp.Selection.Copy()
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                Else
                    '从Excel中复制
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("B46:E52").Select
                    ExcelApp.Selection.Copy()
                    '粘贴到Word
                    WordDoc.Selection.Paste
                    WordDoc.Selection.TypeParagraph
                End If
                '复制盈亏平衡分析图折线图
                ExcelApp.ActiveSheet.ChartObjects("盈亏平衡分析图").Activate
                ExcelApp.ActiveChart.ChartArea.Copy()
                '粘贴到Word
                WordDoc.Selection.Paste
                WordDoc.Selection.TypeParagraph
                '————————————————————————————————————————————————————————————————————————————————————————
                '插入下一页分节符
                WordDoc.Selection.InsertBreak(Type:=Word.WdBreakType.wdSectionBreakNextPage)
                '激活表格
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate
            End If
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Protect(Password:="wscjc") '锁定表格
            '————————————————————————————————————————————————————————————————————————————————————————
            '将导出的表格居中
            WordDoc.Selection.WholeStory '全选
            WordDoc.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter '居中
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("A5").Select
            '保存WORD
            WordApp.ActiveDocument.SaveAs（FilePath & "\" & FileName & "经济评价详表" & ".docx"）
            MsgBox("经济评价详表已导出至Word，并已经保存在：（" & FilePath & "）文件夹下！，文件名为：（" & FileName & "经济评价详表" & ".docx）")
        Else
            MsgBox("请保存并关闭所有已经打开的Word文档，再运行经济评价详表导出命令！")
        End If
        Me.Close()
    End Sub

    Private Sub 经济评价表格导出_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Com技术经济分析计算程序.Form11.Checkbox1.Checked = True
        Com技术经济分析计算程序.Form11.CheckBox2.Checked = True
        Com技术经济分析计算程序.Form11.CheckBox3.Checked = True
        Com技术经济分析计算程序.Form11.CheckBox4.Checked = True
    End Sub

End Class