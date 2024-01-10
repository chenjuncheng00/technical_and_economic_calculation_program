Module 反算收益率临界点
    Sub 资本金税后收益率反算收入单价(ExcelApp As Object, FSLJDJSCSMax As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '资本金税后收益率，资本金税后收益率，资本金税后收益率，
        '反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim SRBH As Integer '收入编号
        Dim SRDJ '收入单价
        Dim SRDJWT As Double = 0 '收入单价微调（用于修正计算结果的误差）
        Dim SRZL '收入总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        '————————————————————————————————————————————————————————————————————————————————————————         
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        '计算折旧摊销
        Dim hscz As Boolean = True
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 7).Value > 0 Then '存在收入
                SRBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 27).Value
                SRDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value
                SRZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 6).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 27).Value = SRBH And SRBH > 0 And SRDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If SRDJ < 1 Then '根据收入单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf SRDJ >= 1 And SRDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf SRDJ >= 10 And SRDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf SRDJ >= 100 And SRDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf SRDJ >= 1000 And SRDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 - m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算                              
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value * (1 + 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率小于设定的临界点内部收益率，则将目前的SRDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT + WTJSBC * m '每次增加WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况， 进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率大于设定的临界点内部收益率，则将目前的SRDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT - WTJSBC * m '每次减少WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value  '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value + WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 资本金税后收益率反算成本单价(ExcelApp As Object, FSLJDJSCSMax As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '资本金税后收益率，资本金税后收益率，资本金税后收益率，
        '反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim CBBH As Integer '成本编号
        Dim CBDJ '成本单价
        Dim CBDJWT As Double = 0 '成本单价微调（用于修正计算结果的误差）
        Dim CBZL '成本总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        '————————————————————————————————————————————————————————————————————————————————————————  
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        '计算折旧摊销
        Dim hscz As Boolean = True
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 14).Value > 0 Then '存在成本
                CBBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 28).Value
                CBDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value
                CBZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 13).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 28).Value = CBBH And CBBH > 0 And CBDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If CBDJ < 1 Then '根据成本单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf CBDJ >= 1 And CBDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf CBDJ >= 10 And CBDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf CBDJ >= 100 And CBDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf CBDJ >= 1000 And CBDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 - m / JSCSMax)
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的CBDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT + WTJSBC * m '每次增加WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value  '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value - WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value * (1 - 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的CBDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT - WTJSBC * m '每次减少WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 资本金税后收益率反算静态投资(ExcelApp As Object, FSLJDJSCSMax As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim TZHH(50) As Integer '数组，储存投资所在的行号
        Dim TZHHJS As Integer = 0 '投资行号计数
        Dim TZJECC(200) '数组，储存各种投资金额
        Dim JTTZWT(200) '数组，静态投资微调
        Dim JTTZZEWT As Double = 0 '静态投资总额微调
        Dim WTJSBC As Integer '微调计算步长
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        '计算折旧摊销
        Dim hscz As Boolean = True
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        '———————————————————————————————————————————————————————————————————————————————————————— 
        For i = 30 To 110 '检索估算表中投资编号，投资编号都是1表示
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 27).Value = 1 Then
                TZHHJS = TZHHJS + 1
                TZHH(TZHHJS) = i '行号储存进数组，从下标1开始
            End If
        Next
        Dim TZJESLJS As Integer = 0 '投资金额数量计数
        '将各种投资金额储存在数组中
        For i = 1 To TZHHJS '行
            For j = 3 To 11 Step 2 '列
                TZJESLJS = TZJESLJS + 1
                TZJECC(TZJESLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(i), j).Value
            Next
        Next
        Dim TZmin = TZJECC(1）  '投资金额中的最小值，初始值为数组中的第一个数
        For i = 1 To 200
            If TZJECC(i) < TZmin And TZJECC(i) > 0 Then
                TZmin = TZJECC(i）
            End If
        Next
        WTJSBC = Math.Round(TZmin / (JSCSMax * 10), 0)
        Dim JS As Integer = 0
        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算资本金临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 - i / JSCSMax)
                        End If
                    Next
                Next
                '投资金额变化后计算
                Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的JTTZWT往上加
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算资本金临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) + WTJSBC * i '每次增加WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 + i / (JSCSMax * 25))
                            End If
                        Next
                    Next
                    '投资金额变化后计算
                    Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value  '将反算出来的总投资记在表格中 
                        Exit For
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value - WTJSBC
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / (JSCSMax * 25))
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                    End If
                Next
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
        End If
        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax * 10
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算资本金临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 + i / JSCSMax)
                        End If
                    Next
                Next
                '投资金额变化后计算
                Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / 2000), 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的JTTZWT往下减
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算资本金临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) - WTJSBC * i '每次减少WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 - i / (JSCSMax * 25))
                            End If
                        Next
                    Next
                    '投资金额变化后计算
                    Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value  '将反算出来的总投资记在表格中 
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                    End If
                Next
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
        End If
        '将反算出的静态投资结果保留为2位小数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value, 2)
        End If
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 全投资税后收益率反算收入单价(ExcelApp As Object, FSLJDJSCSMax As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '全投资税后收益率，全投资税后收益率，全投资税后收益率，
        '反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim SRBH As Integer '收入编号
        Dim SRDJ '收入单价
        Dim SRDJWT As Double = 0 '收入单价微调（用于修正计算结果的误差）
        Dim SRZL '收入总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        '————————————————————————————————————————————————————————————————————————————————————————  
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        '计算折旧摊销
        Dim hscz As Boolean = True
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 7).Value > 0 Then '存在收入
                SRBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 27).Value
                SRDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value
                SRZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 6).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 27).Value = SRBH And SRBH > 0 And SRDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If SRDJ < 1 Then '根据收入单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf SRDJ >= 1 And SRDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf SRDJ >= 10 And SRDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf SRDJ >= 100 And SRDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf SRDJ >= 1000 And SRDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 - m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value * (1 + 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率小于设定的临界点内部收益率，则将目前的SRDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT + WTJSBC * m '每次增加WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率大于设定的临界点内部收益率，则将目前的SRDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT - WTJSBC * m '每次减小WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value  '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value + WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 全投资税后收益率反算成本单价(ExcelApp As Object, FSLJDJSCSMax As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '全投资税后收益率，全投资税后收益率，全投资税后收益率，
        '反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim CBBH As Integer '成本编号
        Dim CBDJ '成本单价
        Dim CBDJWT As Double = 0 '成本单价微调（用于修正计算结果的误差）
        Dim CBZL '成本总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        '————————————————————————————————————————————————————————————————————————————————————————  
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        '计算折旧摊销
        Dim hscz As Boolean = True
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 14).Value > 0 Then '存在成本
                CBBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 28).Value
                CBDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value
                CBZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 13).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 28).Value = CBBH And CBBH > 0 And CBDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If CBDJ < 1 Then '根据成本单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf CBDJ >= 1 And CBDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf CBDJ >= 10 And CBDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf CBDJ >= 100 And CBDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf CBDJ >= 1000 And CBDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 - m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的CBDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT + WTJSBC * m '每次增加WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value - WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '收入和成本变化后相关计算
                                Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value * (1 - 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的CBDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT - WTJSBC * m '每次减小WTJSBC
                                    '收入和成本变化后相关计算
                                    Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '收入和成本变化后相关计算
                            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 全投资税后收益率反算静态投资(ExcelApp As Object, FSLJDJSCSMax As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim TZHH(50) As Integer '数组，储存投资所在的行号
        Dim TZHHJS As Integer = 0 '投资行号计数
        Dim TZJECC(200) '数组，储存各种投资金额
        Dim JTTZWT(200) '数组，静态投资微调
        Dim JTTZZEWT As Double = 0 '静态投资总额微调
        Dim WTJSBC As Integer '微调计算步长
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        '————————————————————————————————————————————————————————————————————————————————————————  
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        '计算折旧摊销
        Dim hscz As Boolean = True
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        '———————————————————————————————————————————————————————————————————————————————————————— 
        For i = 30 To 110 '检索估算表中投资编号，投资编号都是1表示
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 27).Value = 1 Then
                TZHHJS = TZHHJS + 1
                TZHH(TZHHJS) = i '行号储存进数组，从下标1开始
            End If
        Next
        Dim TZJESLJS As Integer = 0 '投资金额数量计数
        '将各种投资金额储存在数组中
        For i = 1 To TZHHJS '行
            For j = 3 To 11 Step 2 '列
                TZJESLJS = TZJESLJS + 1
                TZJECC(TZJESLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(i), j).Value
            Next
        Next
        Dim TZmin = TZJECC(1）  '投资金额中的最小值，初始值为数组中的第一个数
        For i = 1 To 200
            If TZJECC(i) < TZmin And TZJECC(i) > 0 Then
                TZmin = TZJECC(i）
            End If
        Next
        WTJSBC = Math.Round(TZmin / (JSCSMax * 10), 0)
        Dim JS As Integer = 0
        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算全投资临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 - i / JSCSMax)
                        End If
                    Next
                Next
                '投资金额变化后计算
                Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的JTTZWT往上加
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算全投资临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) + WTJSBC * i '每次增加WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 + i / (JSCSMax * 25))
                            End If
                        Next
                    Next
                    '投资金额变化后计算
                    Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value '将反算出来的总投资记在表格中 
                        Exit For
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value - WTJSBC
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / (JSCSMax * 25))
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                    End If
                Next
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
        End If
        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax * 10
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算全投资临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 + i / JSCSMax)
                        End If
                    Next
                Next
                '投资金额变化后计算
                Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / 2000), 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的JTTZWT往下减
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算全投资临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) - WTJSBC * i '每次减少WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 - i / (JSCSMax * 25))
                            End If
                        Next
                    Next
                    '投资金额变化后计算
                    Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value  '将反算出来的总投资记在表格中 
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                    End If
                Next
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
        End If
        '将反算出的静态投资结果保留为2位小数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value, 2)
        End If
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
End Module
