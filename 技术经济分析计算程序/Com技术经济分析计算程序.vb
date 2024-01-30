Imports Microsoft.Office.Interop
Imports System.IO
<ComClass(Com技术经济分析计算程序.ClassId, Com技术经济分析计算程序.InterfaceId, Com技术经济分析计算程序.EventsId)>
Public Class Com技术经济分析计算程序
#Region "COM GUID"
    ' 这些 GUID 提供此类的 COM 标识 
    ' 及其 COM 接口。若更改它们，则现有的
    ' 客户端将不再能访问此类。
    Public Const ClassId As String = "de8978b4-8bff-40cb-8dd3-20d43f6a6fea"
    Public Const InterfaceId As String = "f9f114d5-c79c-462b-a470-00ede680b72f"
    Public Const EventsId As String = "a84fa039-2e81-4ccd-a62a-cc28e28fc11a"
#End Region
    ' 可创建的 COM 类必须具有一个不带参数的 Sub New() 
    ' 否则， 将不会在 
    ' COM 注册表中注册此类，且无法通过
    ' CreateObject 创建此类。
    Sub New()
        MyBase.New()
    End Sub
    Sub 一键出表()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ = MsgBox("是否将全部经济评价计算表格导出至Word文档？", vbOKCancel)
        If XZ = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————  
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Dim Form11 As New 设置表格导出内容
            Form11.ShowDialog() '窗口显示
            Form11.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
    End Sub
    Sub 确定建设期时间计划()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确定已经输入的建设期时间计划？", vbOKCancel)
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '定义局部变量
            Dim j1 As Integer = 0
            Dim j2 As Integer = 0
            Dim j3 As Integer = 0
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————     
            '添加一些报错功能
            '建设期年份总数不可以大于10
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(162, 7).Value > 10 Then
                MsgBox("建设期总年份数量不可以大于10，请重新输入建设期时间计划！")
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------------------------
            '第一次建设期
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 2).Value Then
                MsgBox("第一次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 2).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value Then
                MsgBox("第一次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 4).Value - 1
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '第二次建设期
            '第二次开始年份不可以小于第一次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value > 0 Then
                MsgBox("第二次建设期开始年份不可以小于等于第一次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value Then
                MsgBox("第二次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value Then
                MsgBox("第二次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 12).Value - 1
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '第三次建设期
            '第三次开始年份不可以小于第二次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value > 0 Then
                MsgBox("第三次建设期开始年份不可以小于等于第二次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value Then
                MsgBox("第三次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value Then
                MsgBox("第三次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 4).Value - 1
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '第四次建设期
            '第四次开始年份不可以小于第三次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value > 0 Then
                MsgBox("第四次建设期开始年份不可以小于等于第三次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value Then
                MsgBox("第四次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value Then
                MsgBox("第四次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 12).Value - 1
                End If
            End If
            '------------------------------------------------------------------------------------------------------------
            '第五次建设期
            '第五次开始年份不可以小于第四次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value > 0 Then
                MsgBox("第五次建设期开始年份不可以小于等于第四次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value Then
                MsgBox("第五次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value Then
                MsgBox("第五次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 4).Value - 1
                End If
            End If
            '------------------------------------------------------------------------------------------------------------
            '第六次建设期
            '第六次开始年份不可以小于第五次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value > 0 Then
                MsgBox("第六次建设期开始年份不可以小于等于第五次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value Then
                MsgBox("第六次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value Then
                MsgBox("第六次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 12).Value - 1
                End If
            End If
            '------------------------------------------------------------------------------------------------------------
            '第七次建设期
            '第七次开始年份不可以小于第六次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value > 0 Then
                MsgBox("第七次建设期开始年份不可以小于等于第六次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value Then
                MsgBox("第七次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value Then
                MsgBox("第七次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 4).Value - 1
                End If
            End If
            '------------------------------------------------------------------------------------------------------------
            '第八次建设期
            '第八次开始年份不可以小于第七次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value > 0 Then
                MsgBox("第八次建设期开始年份不可以小于等于第七次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value Then
                MsgBox("第八次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value Then
                MsgBox("第八次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 12).Value - 1
                End If
            End If
            '------------------------------------------------------------------------------------------------------------
            '第九次建设期
            '第九次开始年份不可以小于第八次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value > 0 Then
                MsgBox("第九次建设期开始年份不可以小于等于第八次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value Then
                MsgBox("第九次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value Then
                MsgBox("第九次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 4).Value - 1
                End If
            End If
            '------------------------------------------------------------------------------------------------------------
            '第十次建设期
            '第十次开始年份不可以小于第九次结束年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value > 0 Then
                MsgBox("第十次建设期开始年份不可以小于等于第九次建设期的结束年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value + 1
            End If
            '结束年份不可以小于开始年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value Then
                MsgBox("第十次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value + 1
            End If
            '结束年份等于开始年份，但是结束月份不可以小于开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value Then
                MsgBox("第十次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value + 1 <= 12 Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value + 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 12).Value - 1
                End If
            End If
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '将之前已有的和确定建设期时间计划相关的计算结果全部重置回默认状态，防止出错
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("G205:G214").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("O205:P214").Value = 0
            '清空逐年投产月份和建设期标记
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i).Value = 0
            Next
            '计算一次
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '进行建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            'PART1、PART2、PART3重置回默认状态
            Call 计算设置重置回默认状态_part1(ExcelApp)
            Call 计算设置重置回默认状态_part2(ExcelApp)
            Call 计算设置重置回默认状态_part3(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_默认值()
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 计算功能合并整理.确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            MsgBox("确定建设期时间计划完成！")
        End If
    End Sub
    Sub 清空时间计划表数据()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————    
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ = MsgBox("是否确定清空建设期时间计划计算出的全部数据？", vbOKCancel)
        If XZ = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————  
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————    
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("G205:G214").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("O205:P214").Value = 0
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B3:B4").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D3:D4").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J3:J4").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L3:L4").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B35:B36").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D35:D36").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J35:J36").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L35:L36").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B67:B68").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D67:D68").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J67:J68").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L67:L68").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B99:B100").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D99:D100").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J99:J100").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L99:L100").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B131:B132").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D131:D132").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J131:J132").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L131:L132").Value = 0
            '判定一下，为了兼容5.9.0~5.9.2
            Dim tmp As String = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 1).Value
            If tmp <> "" Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 2).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 10).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 2).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 10).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 2).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 10).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 2).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 10).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 2).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 10).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 12).Value = 0
            End If
            '计算一次
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            '计算一次
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '将已经输入的投资金额和各种量清空
            Dim XZ1 = MsgBox("是否需要同时清空<估算表>中输入的逐年投资额、建设期可抵扣增值税额、各种年量等数据？", vbOKCancel)
            If XZ1 = vbOK Then
                '第一到第十次投资金额
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '建设期可抵扣增值税额
                For i = 38 To 81 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '各种年量、前五次投资
                For i = 44 To 55 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '各种年量、后五次投资
                For i = 91 To 102 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
            End If
            '计算一次
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————
            'PART1、PART2、PART3重置回默认状态
            Call 计算设置重置回默认状态_part1(ExcelApp)
            Call 计算设置重置回默认状态_part2(ExcelApp)
            Call 计算设置重置回默认状态_part3(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_默认值()
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '年限系数变化后相关计算(包括了投资金额变化后计算、收入成本变化后计算)
            Call 计算功能合并整理.确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("清空建设期时间计划完成！")
        End If
    End Sub
    Sub 确定估算表参数设置()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————, xlfl_model—————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ = MsgBox("是否确定已经输入的估算表参数？", vbOKCancel)
        If XZ = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————        
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            'PART2、PART3重置回默认状态
            Call 计算设置重置回默认状态_part2(ExcelApp)
            Call 计算设置重置回默认状态_part3(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '年限系数变化后相关计算(包括了投资金额变化后计算、收入成本变化后计算)
            Call 计算功能合并整理.确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("确定估算表参数设置完成！")
        End If
    End Sub
    Sub 确定投资数据输入()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————, xlfl_model—————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ = MsgBox("是否确定已经输入的投资数据？", vbOKCancel)
        If XZ = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————        
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定投资数据输入
            Call 计算功能合并整理.确定投资数据输入(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("确定投资数据输入完成！")
        End If
    End Sub
    Sub 清空投资数据输入()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————, xlfl_model—————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ = MsgBox("是否确定已经输入的投资数据？", vbOKCancel)
        If XZ = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————        
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '将已经输入的投资金额和各种量清空
            Dim XZ1 = MsgBox("是否需要清空<估算表>中输入的逐年投资额、建设期可抵扣增值税额、各种年量等数据？", vbOKCancel)
            If XZ1 = vbOK Then
                '第一到第十次投资金额
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '建设期可抵扣增值税额
                For i = 38 To 81 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '各种年量、前五次投资
                For i = 44 To 55 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '各种年量、后五次投资
                For i = 91 To 102 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                        End If
                    Next
                Next
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '计算一次
                ExcelApp.Calculate()
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————
                'PART2、PART3重置回默认状态
                Call 计算设置重置回默认状态_part2(ExcelApp)
                Call 计算设置重置回默认状态_part3(ExcelApp)
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算模式
                Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
                Dim zbj_model = ans_mode(0)
                Dim hscz = ans_mode(1)
                Dim xlfl_cg_model = ans_mode(2)
                Dim xlfl_qt_model = ans_mode(3)
                Dim clfl_qtfl_model = ans_mode(4)
                Dim sdsl_model = ans_mode(5)
                Dim kcje_xlf_model = ans_mode(6)
                Dim kcje_clf_qtf_model = ans_mode(7)
                Dim ldzj_model = ans_mode(8)
                Dim kcje_ldzj_model = ans_mode(9)
                Dim bxf_model = ans_mode(10)
                Dim kcje_bxf_model = ans_mode(11)
                '————————————————————————————————————————————————————————————————————————————————————————
                '年限系数变化后相关计算(包括了投资金额变化后计算、收入成本变化后计算)
                Call 计算功能合并整理.确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
                '————————————————————————————————————————————————————————————————————————————————————————
                Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
                MsgBox("清空输入的投资数据完成！")
            End If
        End If
    End Sub
    Sub 确定收入成本输入()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————, xlfl_model—————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ = MsgBox("是否确定已经输入的收入成本数据？", vbOKCancel)
        If XZ = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————        
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '收入成本变化后计算
            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("确定收入成本输入完成！")
        End If
    End Sub
    Sub 清空收入成本输入()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否清空<收入&成本输入表>中收入和成本数据", vbOKCancel)
        If XZ1 = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————        
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '收入
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("E13:F25").Value = 0
            '成本
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("L14:M19").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("L21:M25").Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("L27:M27").Value = 0
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 19).Value = "直接输入" Then
                '直接输入的逐年负荷系数
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(95, 3).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(95, 4).Value = 0.35
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(95, 5).Value = 0.7
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("F95:Q95").Value = 1
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B97:Q97").Value = 1
            Else
                '逐次输入的负荷达产率
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("F119:O119").Value = Nothing
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C123:Q123").Value = Nothing
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B125:Q125").Value = Nothing
            End If
            '计算一次
            ExcelApp.Calculate()
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '收入成本变化后的操作
            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '———————————————————————————————————————————————————————————————————————————————————————— 
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("清空收入成本输入完成！")
        End If
    End Sub

    Sub 反算临界点()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否进行反算临界点计算？", vbOKCancel)
        If XZ1 = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————        
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Dim FSLJDJSCSMax As Integer = InputBox("请输入反算临界点功能的最大计算次数，输入的数字越大，计算次数越多，计算速度越慢，计算精度越高。", "输入反算临界点最大计算次数", 200)
            '实例化计算进度显示窗体
            Dim Form1 As New 计算进度显示
            '确定选择的需要反算的内部收益率类型
            '如果是资本金所得税后内部收益率，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 5).Value = "资本金所得税后内部收益率" Then
                Call 资本金税后收益率反算静态投资(ExcelApp, FSLJDJSCSMax, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                                  kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
                Call 资本金税后收益率反算收入单价(ExcelApp, FSLJDJSCSMax, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                                  kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
                Call 资本金税后收益率反算成本单价(ExcelApp, FSLJDJSCSMax, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                                  kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            End If
            '如果是全投资所得税后内部收益率，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 5).Value = "全投资所得税后内部收益率" Then
                Call 全投资税后收益率反算静态投资(ExcelApp, FSLJDJSCSMax, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                                  kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
                Call 全投资税后收益率反算收入单价(ExcelApp, FSLJDJSCSMax, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                                  kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
                Call 全投资税后收益率反算成本单价(ExcelApp, FSLJDJSCSMax, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                                  kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————
            '提醒计算完成
            Form1.Show()
            Form1.Label1.Text = "反算临界点已计算完成！"
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
    End Sub
    Sub 清空反算临界点数据()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否清空反算临界点计算的数据", vbOKCancel)
        If XZ1 = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————  
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '清空数据
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("D43:D67").ClearContents
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("G42:G67").ClearContents
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("反算临界点数据已清空！")
        End If
    End Sub
    Sub 敏感性分析计算()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否进行敏感性分析计算？", vbOKCancel)
        If XZ1 = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————       
            Dim Form10 As New 设置单因素敏感性计算内容
            '显示敏感性分析选择窗口
            Form10.ShowDialog() '窗口显示
            Form10.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————       
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
    End Sub
    Sub 盈亏平衡点计算()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否进行盈亏平衡分析计算？本计算数据采用的是总成本最大年的数据。", vbOKCancel)
        If XZ1 = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————       
            '本方法为，按照总成本最大年份进行计算
            '定义局部变量
            Dim nfxh As Integer = 0
            '寻找总成本费用最大值的年份序号，并记录下年固定成本和可变成本
            For i = 5 To 19 '前15年
                If ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(24, i).Value = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(1, 4).Value Then
                    nfxh = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4, i).Value '记录下年份序号
                    '记录下年固定成本、可变成本
                    '年固定成本
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(56, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(26, i).Value
                    '年可变成本
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(58, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(25, i).Value
                    Exit For
                End If
            Next
            For i = 4 To 19 '后16年
                If ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(56, i).Value = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(1, 4).Value Then
                    nfxh = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36, i).Value '记录下年份序号
                    '记录下年固定成本、可变成本
                    '年固定成本
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(56, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(58, i).Value
                    '年可变成本
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(58, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(57, i).Value
                    Exit For
                End If
            Next
            '记录选择的计算年份
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(55, 5).Value = nfxh
            '记录年销售收入、税金及附加
            For i = 5 To 19 '前15年
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(4, i).Value = nfxh Then
                    '年销售收入
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i).Value
                    '年税金及附加
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(59, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(17, i).Value
                End If
            Next
            '记录年销售收入、税金及附加
            For i = 4 To 19 '后16年
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(27, i).Value = nfxh Then
                    '年销售收入
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28, i).Value
                    '年税金及附加
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(59, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(40, i).Value
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————   
            '生产年份平均值计算
            '计算生产的年份数
            Dim scnfs As Integer = 0
            '前15年
            For i = 1 To 15
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i + 4).Value > 0 Then
                    scnfs += 1
                End If
            Next
            '16-31年
            For i = 16 To 31
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28, i - 12).Value > 0 Then
                    scnfs += 1
                End If
            Next
            '生产年份数写入表格
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(47, 5).Value = scnfs
            '————————————————————————————————————————————————————————————————————————————————————————  
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————       
            MsgBox("盈亏平衡点（BEP）已计算完成！")
        End If
    End Sub
    Sub 复制敏感性分析图()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否要复制已经生成的敏感性分析图？", vbOKCancel)
        If XZ1 = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Try
                '激活表格
                ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
                '复制
                ExcelApp.ActiveChart.ChartArea.Copy()
            Catch ex As Exception
                'MsgBox("单因素敏感性分析图不存在！")
            End Try
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————
            MsgBox("复制完成！")
        End If
    End Sub
    Sub 复制盈亏平衡分析图()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ1 = MsgBox("是否要复制已经生成的盈亏平衡分析图？", vbOKCancel)
        If XZ1 = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '激活表格
            ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("盈亏平衡分析图").Activate
            '复制
            ExcelApp.ActiveChart.ChartArea.Copy()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————
            MsgBox("复制完成！")
        End If
    End Sub
    Sub 清空敏感性分析和盈亏平衡分析计算数据()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '弹出对话框
        Dim XZ2 = MsgBox("是否要清空已有的全部敏感性分析和盈亏平衡分析计算数据", vbOKCancel)
        If XZ2 = vbOK Then
            '————————————————————————————————————————————————————————————————————————————————————————  
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '删除敏感性分析图表
            Try
                ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
                ExcelApp.ActiveChart.Parent.Delete
            Catch ex As Exception
                '不进行操作
            End Try
            '————————————————————————————————————————————————————————————————————————————————————————
            '将敏感性分析变化率重置回5%，格式重置回0%
            For i = 7 To 137 Step 5
                For j = 1 To 5
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).Value = 0.05 * j - 3 * 0.05
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).NumberFormatLocal = "0%"
                Next
            Next
            '清除敏感性分析数据
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("H7:H141").ClearContents
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("J7:L141").ClearContents
            '清空Excel内已有的输入
            For i = 7 To 137 Step 5
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 29).Value = Nothing
            Next
            '清空Excel内已有的输入(年运行小时数敏感性分析内容)
            '收入
            For i = 13 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 30).Value = Nothing
            Next
            '成本
            For i = 14 To 19
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            For i = 21 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 31).Value = Nothing
            '——————————————————————————————————————————————————————————————————————————————————————
            '清空盈亏平衡分析计算数据
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("E55:E59").ClearContents
            '——————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            MsgBox("清空敏感性分析和盈亏平衡分析计算数据完成！")
        End If
    End Sub
    Sub 设置所得税减免和增值税退税内容()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置所得税减免和增值税退税包含内容?", vbOKCancel)
        Dim Form6 As New 设置所得税减免和增值税退税内容
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form6.ShowDialog() '窗口显示
            Form6.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 设置收入和成本计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置部分销售收入和经营成本计算年限?", vbOKCancel)
        Dim Form13 As New 设置收入和成本计算年限
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form13.ShowDialog() '窗口显示
            Form13.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub
    Sub 设置接入费计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置接入费收入计算方式?", vbOKCancel)
        Dim Form12 As New 设置接入费计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form12.ShowDialog() '窗口显示
            Form12.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 设置补贴收入计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置补贴收入的计算年限?", vbOKCancel)
        Dim Form7 As New 设置补贴收入计算年限
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form7.ShowDialog() '窗口显示
            Form7.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 设置逐年衰减计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要修改程序内置的默认光伏发电逐年衰减效率或者蓄电池逐年衰减效率?", vbOKCancel)
        Dim Form2 As New 设置逐年衰减计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '读取输入的衰减开始年份和衰减率（%）
            Form2.ShowDialog() '窗口显示
            Form2.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“逐年衰减系数设置.确定”
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 设置修理费计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置逐年动态变化的修理费率?", vbOKCancel)
        Dim Form4 As New 设置修理费率计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '读取输入的修理费率变化率（%）
            Form4.ShowDialog() '窗口显示
            Form4.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“修理费率逐年变化设置.确定”
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub
    Sub 设置材料费和其它费计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置逐年动态变化的材料费率和其它费率?", vbOKCancel)
        Dim Form As New 设置材料费和其它费计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            '读取输入的修理费率变化率（%）
            Form.ShowDialog() '窗口显示
            Form.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“材料费率其它费率逐年变化设置.确定”
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub
    Sub 设置建设期资金运用方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置建设期资金运用方式?包括每次投资的资本金比例系数取值方式和建设期贷款利率。", vbOKCancel)
        Dim Form5 As New 设置建设期资金运用方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form5.ShowDialog() '窗口显示
            Form5.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“建设期资金运用方式设置.确定”
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
    End Sub

    Sub 设置折旧摊销计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要每次投资设置不同的固定资产折旧和无形资产摊销计算方式？", vbOKCancel)
        Dim Form8 As New 设置折旧摊销计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form8.ShowDialog() '窗口显示
            Form8.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“每次投资设置不同的折旧摊销计算方式.确定”
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
    End Sub

    Sub 设置长期贷款计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要每次投资设置不同的长期贷款还款年限和宽限年限？", vbOKCancel)
        Dim Form9 As New 设置长期贷款计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form9.ShowDialog() '窗口显示
            Form9.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“每次投资设置不同的长期贷款还款年限和宽限年限.确定”
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
    End Sub

    Sub 设置投资各方收益计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要分别设置投资各方收益计算参数？", vbOKCancel)
        Dim Form14 As New 设置投资各方计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form14.ShowDialog() '窗口显示
            Form14.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
    End Sub
    Sub 设置流动资金计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置流动资金计算方式?", vbOKCancel)
        Dim Form15 As New 设置流动资金计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form15.ShowDialog() '窗口显示
            Form15.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub
    Sub 设置保险费计算方式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否需要设置保险费率计算方式?", vbOKCancel)
        Dim Form16 As New 设置保险费率计算方式
        If XZ = vbOK Then
            Call 计算前基本处理(ExcelApp, 1)
            '————————————————————————————————————————————————————————————————————————————————————————
            Form16.ShowDialog() '窗口显示
            Form16.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        End If
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 进入维护模式()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取计数
        Dim Form3 As New 进入维护模式
        If (ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value < 3 And ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value >= 0) Then '最多只可以连续错3次。单元格Z2
            Form3.ShowDialog() '窗口显示
            Form3.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            Form3.TextBox1.Text = Nothing '清空已有的内容
        Else
            MsgBox("已超过最大尝试次数，不可以再尝试输入密码！")
            '锁定表格
            Call 锁定表格(ExcelApp)
        End If
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 打开表格自动运行()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '启动迭代计算
        ExcelApp.Application.Iteration = True
        '打开自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件触发
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        '自保护程序
        Call 自保护程序(ExcelApp)
        '验证Excel表格的更新时间
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 29).Value < 20210903 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
            '跳过投资各方收益表格操作，防止多次报错
            Exit Sub
        End If
        '投资各方收益
        Call 投资各方收益率表格操作()
        '锁定表格
        Call 锁定表格(ExcelApp)
        '保存对表格的改动
        ExcelApp.Application.DisplayAlerts = False
        ExcelApp.ThisWorkbook.Save()
        ExcelApp.Application.DisplayAlerts = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 投资各方收益率表格操作()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC As Integer = Excel版本号验证(ExcelApp)
        If ZTJC = 1 Then
            Call 锁定表格(ExcelApp)
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        Call 计算前基本处理(ExcelApp, 1)
        '————————————————————————————————————————————————————————————————————————————————————————  
        '读取输入的投资各方出资比例
        Dim TZGFCZBL(10) As Double '数组，储存投资各方出资比例
        Dim TZFJS As Integer = 0 '投资各方出资比例大于0的数量计数
        For i = 0 To 4
            TZGFCZBL(i) = ExcelApp.Worksheets("建设期时间计划表").Cells(163 + i, 2).Value
            If TZGFCZBL(i) > 0 Then
                TZFJS = TZFJS + 1
            End If
        Next
        '针对读取的投资各方出资比例添加报错功能
        '不能隔行输入
        If TZGFCZBL(0) = 0 And (TZGFCZBL(1) > 0 Or TZGFCZBL(2) > 0 Or TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            GoTo aaaaa
        End If
        If TZGFCZBL(1) = 0 And (TZGFCZBL(2) > 0 Or TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            GoTo aaaaa
        End If
        If TZGFCZBL(2) = 0 And (TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            GoTo aaaaa
        End If
        If TZGFCZBL(3) = 0 And TZGFCZBL(4) > 0 Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            GoTo aaaaa
        End If
        '根据输入的情况，自动补全后面一个出资比例，保证总出资比例为100%
        If ExcelApp.Worksheets("建设期时间计划表").Cells(168, 2).Value <= 100 Then '如果总和小于等于100
            If TZFJS = 1 Then
                ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 100 - TZGFCZBL(0)
            End If
            If TZFJS = 2 Then
                ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 100 - TZGFCZBL(1) - TZGFCZBL(0)
            End If
            If TZFJS = 3 Then
                ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 100 - TZGFCZBL(2) - TZGFCZBL(1) - TZGFCZBL(0)
            End If
            If TZFJS = 4 Then
                ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 100 - TZGFCZBL(3) - TZGFCZBL(2) - TZGFCZBL(1) - TZGFCZBL(0)
            End If
        End If
        '总和不能大于100
        If ExcelApp.Worksheets("建设期时间计划表").Cells(168, 2).Value > 100 Then
            'MsgBox("投资各方出资比例总和不可以大于100%，程序将自动重置回默认值！")
            '投资各方出资比例设置为：第一的不变，第二的变成100减去第一年，第三到第五为0            
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 100 - ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value
            GoTo aaaaa
        End If
        '如果输入到了投资方5输入比例，但是总和小于100
        If TZGFCZBL(4) > 0 And ExcelApp.Worksheets("建设期时间计划表").Cells(168, 2).Value < 100 Then
            'MsgBox("投资各方出资比例总和不可以小于100%，程序将自动重置回默认值！")
            '投资各方出资比例设置为：第一到第四不变，第五等于100减去前四个之和
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 100 - ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value - ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value - ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value - ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value
            GoTo aaaaa
        End If
aaaaa：
        '再重新读取一次各方投资比例
        For i = 0 To 4
            TZGFCZBL(i) = ExcelApp.Worksheets("建设期时间计划表").Cells(163 + i, 2).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '将输入的投资各方比例写入表格
        '投资方1
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        '投资方2
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        '投资方3
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        '投资方4
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        '投资方5
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的投资各方出资比例，自动显示或者隐藏相关工作表
        If TZGFCZBL(0) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(1) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(2) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(3) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(4) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '写入计算模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算模式
        Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
        Dim zbj_model = ans_mode(0)
        Dim hscz = ans_mode(1)
        Dim xlfl_cg_model = ans_mode(2)
        Dim xlfl_qt_model = ans_mode(3)
        Dim clfl_qtfl_model = ans_mode(4)
        Dim sdsl_model = ans_mode(5)
        Dim kcje_xlf_model = ans_mode(6)
        Dim kcje_clf_qtf_model = ans_mode(7)
        Dim ldzj_model = ans_mode(8)
        Dim kcje_ldzj_model = ans_mode(9)
        Dim bxf_model = ans_mode(10)
        Dim kcje_bxf_model = ans_mode(11)
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 计算后基本处理(ExcelApp, 0, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
    End Sub
End Class
