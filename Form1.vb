Public Class Form1
    Dim PgN As Integer = 1
    Dim AFreq As Single
    Dim PtVAC As String
    Dim PnRow As Integer = 2
    Dim PnClm As Integer = 10
    Dim ClmE As Integer = 50

    ' вычисление спектров
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Names()

        Dim EA As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EAVAC As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim k, l, m, n, o, cSign, KeyP, Rclm As Integer
        Dim Ttable(500, 10), Ky, Kd, Kp As Single
        Dim Rtable(500), Pname As String
        Dim PtMERA As String = TextBox1.Text
        Dim PtXLS As String = TextBox2.Text

        TextBox3.Text = "" : Label13.Text = "0"
        ToolStripStatusLabel1.Text = "Reading passport-file" : Me.Refresh()

        ' открытие паспорта
        Try
            EA.Workbooks.Open(PtXLS)
        Catch ex As Exception
            MsgBox("File " & PtXLS & " does not exist.")
            Exit Sub
        End Try

        Try
            EAVAC.Workbooks.Open(PtVAC)
        Catch ex As Exception
            MsgBox("File " & PtVAC & " does not exist.")
            Exit Sub
        End Try

        If CheckBox10.Checked = True Then
            EA.Visible = False
            EAVAC.Visible = False
        Else
            EA.Visible = True
            EAVAC.Visible = True
        End If

        Try
            EA.Worksheets("MAIN").Select()
        Catch ex As Exception
            MsgBox("Sheet MAIN does not exist.")
            Exit Sub
        End Try

        ' определение числа обрабатываемых параметров
        Do
            KeyP = CSng(EA.Cells(PnRow + 1, PnClm + k).Value)
            If CSng(EA.Cells(PnRow + 1, PnClm + k).Value) > 0 Then
                cSign += 1
            End If
            k += 1
        Loop While CStr(EA.Cells(PnRow, PnClm + k).Value) <> ""

        k = 0
        Do
            KeyP = CSng(EA.Cells(PnRow + 1, PnClm + k).Value)   ' ключ обработки параметра
            Pname = CStr(EA.Cells(PnRow, PnClm + k).Value)      ' имя параметра
            Rclm = CInt(EAVAC.Cells(3, 2 + k).Value)            ' номер столбика с оборотами

            If CStr(EA.Cells(2, 1).Value) <> "" Then  ' считываем параметры датчиков (ЛИ)
                Ky = CSng(EA.Cells(EA.Cells(2, 1).Value, PnClm + k).Value)
                Kd = CSng(EA.Cells(EA.Cells(2, 1).Value + 1, PnClm + k).Value)
                Kp = CSng(EA.Cells(EA.Cells(2, 1).Value + 2, PnClm + k).Value)
            End If

            If KeyP = 2 Then

                Try
                    EA.Worksheets(Pname).Select()
                Catch ex As Exception
                    MsgBox("Sheet " & Pname & " does not exist.")
                    Exit Sub
                End Try
                n = k

            End If

            If KeyP > 0 Then
                TextBox3.Text = TextBox3.Text & Pname : Me.Refresh()

                Do
                    KeyP = CSng(EA.Cells(l + PnRow + 2, PnClm + k - n).Value)        ' проверка запроса обработки по режиму
                    If KeyP = 1 Then                                                 ' ключ обработки режима
                        Ttable(m + 1, 1) = CSng(EA.Cells(PnRow + 2 + l, 3).Value)    ' время начало обработки
                        Ttable(m + 1, 2) = CSng(EA.Cells(PnRow + 2 + l, 4).Value)    ' время конца обработки
                        Ttable(m + 1, 3) = CSng(EA.Cells(PnRow + 2 + l, 5).Value)    ' уровень тяги
                        Ttable(m + 1, 4) = CSng(EA.Cells(PnRow + 2 + l, Rclm).Value) ' скорость вращения
                        Ttable(m + 1, 5) = CSng(EA.Cells(PnRow + 2 + l, 7).Value)    ' Km
                        Rtable(m + 1) = CStr(EA.Cells(PnRow + 2 + l, 1).Value)       ' шифр режима
                        m += 1
                    End If
                    l += 1
                Loop While CStr(EA.Cells(PnRow + 2 + l, 1).Value) <> ""

                ' Перегруппировка режимов по парным номерам
                'If m > 2 Then
                'Regrp(m, Ttable, Rtable)
                'End If

                ' Вычисление спектров
                If m > 0 Then
                    o += 1
                    Call SignatureTroughputProcessing(m, o, Pname, Ttable, Rtable, Ky, Kd, Kp)
                    ToolStripProgressBar1.Value = (o / cSign) * 100
                End If
            ElseIf KeyP = 0 Then
                TextBox3.Text = TextBox3.Text & Pname & vbTab & "skipped" & vbCrLf : Me.Refresh()
            End If

            m = 0
            l = 0
            n = 0
            k += 1
            EA.Worksheets("MAIN").Select()
        Loop While CStr(EA.Cells(PnRow, PnClm + k).Value) <> ""

        EA.Quit()
        EAVAC.Quit()

        If CheckBox1.Checked = True And o > 0 Then  ' если o=0 в паспорте не помечены параметры для обработки
            Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
            Dim DB As LMSTestLabAutomation.IDatabase
            DB = TL.ActiveBook.Database
            ToolStripStatusLabel1.Text = "Saving LMS file" : Me.Refresh()
            TL.ActiveBook.Save(TextBox5.Text)
        End If

        ToolStripStatusLabel1.Text = "Finished"
        ToolStripProgressBar1.Value = 0 : Me.Refresh()

        If CheckBox4.Checked = True Then
            TabControl1.SelectedTab = TabPage2
            TabPage2_Enter(sender, e)
            Button2_Click(sender, e)
        End If
    End Sub

    Private Sub Regrp(ByVal m As Integer, ByVal Ttable(,) As Single, ByVal Rtable() As String)
        Dim i, j As Integer

        For i = 1 To m - 2

            For j = i + 1 To m

                If Rtable(i) = Rtable(j) And j <> i + 1 Then
                    Dim ResT(7) As Single
                    Dim ResR As String

                    For k = 1 To 7
                        ResT(k) = Ttable(i + 1, k)
                        ResR = Rtable(i + 1)

                        Ttable(i + 1, k) = Ttable(j, k)
                        Rtable(i + 1) = Rtable(j)

                        Ttable(j, k) = ResT(k)
                        Rtable(j) = ResR
                    Next

                    Exit For
                ElseIf Rtable(i) = Rtable(j) And j = i + 1 Then
                    Exit For
                End If
            Next

        Next

    End Sub

    Private Sub SignatureTroughputProcessing(ByVal m As Integer, ByVal o As Integer, ByVal Pname As String, ByVal Ttable(,) As Single, ByVal Rtable() As String, ByVal Ky As Single, ByVal Kd As Single, ByVal Kp As Single)
        Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
        Dim DB As LMSTestLabAutomation.IDatabase

        If TL.Name <> "" And o = 1 Then
            TL.ActiveBook.Close()
            'TL.OpenProject(ComboBox7.Text & "\Project1.lms")
            TL.NewProject(ComboBox7.Text & "\Project1.lms")
        ElseIf TL.Name = "" Then
            ToolStripStatusLabel1.Text = "Opening TestLab..." : Me.Refresh()
            TL.Init("-w DesktopStandard")
        End If
        DB = TL.ActiveBook.Database

        Dim DW As LMSTestLabAutomation.DataWatch = TL.ActiveBook.FindDataWatch("Navigator_Explorer")
        Dim Ex As LMSTestLabAutomation.IExplorer = DW.Data
        Dim Br As LMSTestLabAutomation.IDataBrowser

        Dim i, k As Integer
        Dim PthMERA As String = TextBox1.Text
        Dim labs(), ids() As String
        Dim ID_sign, ID_sinus As LMSTestLabAutomation.IData
        Dim AM As LMSTestLabAutomation.AttributeMap
        Dim IB_reg, IB_regX, IB_cent_reg, IB_cent_reg_new, IB_Power, IB_aver As LMSTestLabAutomation.IBlock2
        Dim R, km, RPM, dT As String

        Try
            Br = Ex.Browser(PthMERA)
        Catch exx As Exception
            MsgBox("File " & PthMERA & " does not exist.")
            Exit Sub
        End Try

        Br.Elements("", labs, ids)  ' чтение имен замеренных параметров в мера файле

        ' чтение синуса (ЛИ)
        If RadioButton3.Checked = True And ComboBox6.Text = "РД171М" Then
            For i = 0 To UBound(ids) - 1    ' проход по каждому параметру
                If labs(i) = "sin" Then
                    ID_sinus = Br.GetItem(ids(i))
                End If
            Next
        End If

        For i = 0 To UBound(ids) - 1    ' проход по каждому параметру
            If labs(i) = Pname Then

                ToolStripStatusLabel1.Text = "Processing: " & Pname : Me.Refresh()
                ' считываем данные, свойства и конвертируем их в блок
                ID_sign = Br.GetItem(ids(i))

                ' вычисление осредненного спектра по каждому режиму
                DB.AddSection(Pname)

                Dim TX As String = TextBox3.Text
                Dim Kcoup As Integer = 1

                For k = 1 To m  ' проход по режимам

                    'If Rtable(k) = Rtable(k - 1) And Kcoup = 1 Then
                    'IB_aver = IB_Power
                    'End If

                    ' выделение режима
                    AM = TL.Factory.CreateAttributeMap
                    AM.Add("StartValueGiven", True)
                    AM.Add("StartValue", Ttable(k, 1))
                    AM.Add("NumberOfPoints", CInt((Ttable(k, 2) - Ttable(k, 1)) * CInt(NumericUpDown3.Text)))
                    AM.Add("BlockStream", ID_sign)
                    IB_reg = TL.Factory.CreateObject("LmsHq::DataModelC::ExprConcat::CBlockStreamToBlock", AM)
                    AM = Nothing

                    'DB.AddItem(Pname & "/", Pname, IB_reg, , 1)

                    ' предобработка сигнала режима (ЛИ)
                    If RadioButton3.Checked = True And ComboBox6.Text = "РД171М" Then
                        IB_reg = TransformSignal(IB_reg, ID_sinus, TL, Ttable(k, 1), CInt((Ttable(k, 2) - Ttable(k, 1)) * CInt(NumericUpDown3.Text)), Ky, Kd, Kp)
                    End If

                    'DB.AddItem(Pname & "/", Pname & "Pr", IB_reg, , 1)

                    Dim vscalar As LMSTestLabAutomation.IScalar
                    Dim StatFunc As LMSTestLabAutomation.Enumerate
                    Dim MaxSp, MinSp As Single

                    ' к сигналам запуска и останова добавление времени
                    If CSng(Rtable(k)) < 6000 And CheckBox6.Checked = True Then

                        AM = TL.Factory.CreateAttributeMap
                        AM.Add("StartValueGiven", True)

                        Dim tStart As Single
                        Dim pNumber As Integer

                        If Rtable(k) = "0" Then
                            tStart = 0
                            pNumber = CInt(Ttable(k, 2) * CInt(NumericUpDown3.Text))
                            AM.Add("StartValue", tStart)
                            AM.Add("NumberOfPoints", pNumber)
                        ElseIf Rtable(k) = "2" Then
                            tStart = Ttable(k, 1)
                            pNumber = CInt((Ttable(k, 2) + 2 - Ttable(k, 1)) * CInt(NumericUpDown3.Text))
                            AM.Add("StartValue", tStart)
                            AM.Add("NumberOfPoints", pNumber)
                        End If

                        AM.Add("BlockStream", ID_sign)
                        IB_regX = TL.Factory.CreateObject("LmsHq::DataModelC::ExprConcat::CBlockStreamToBlock", AM)
                        AM = Nothing




                        ' предобработка сигнала режима (ЛИ)
                        If RadioButton3.Checked = True And ComboBox6.Text = "РД171М" Then
                            IB_regX = TransformSignal(IB_regX, ID_sinus, TL, tStart, pNumber, Ky, Kd, Kp)
                        End If





                        ' расчет среднего значения
                        AM = TL.Factory.CreateAttributeMap
                        StatFunc = TL.Factory.CreateEnumerate("LmsHq::DataModelI::BlockStatistics::CBufferIEnumStatisticFunction", 10)
                        AM.Add("SourceBlock", IB_regX)
                        AM.Add("StatisticFunction", StatFunc)
                        vscalar = TL.Factory.CreateObject("LmsHq::DataModelC::BlockStatistics::CStatisticBlockMetric", AM)
                        AM = Nothing

                        ' центровка
                        IB_regX = TL.cmd.BLOCK_SUBTRACT_SCALAR(IB_regX, vscalar)

                        ' смещение, запуск до 5 сек останов после 5сек (для сравнения различных испытаний)
                        Dim Nt As Single = 5
                        Dim SpTimStart As Single ' поиск всплесков на старте за XX сек до выхода на режим

                        If ComboBox6.Text = "РД191" Then SpTimStart = Nt - 1.5
                        If ComboBox6.Text = "РД171М" Then SpTimStart = Nt - 1.5
                        If ComboBox6.Text = "РД180" Then SpTimStart = Nt - 1.0
                        If ComboBox6.Text = "РД180М" Then SpTimStart = Nt - 1.5
                        If ComboBox6.Text = "XXL" Then SpTimStart = Nt - 0.5

                        Dim SpTimShDown As Single = Nt + 0.5 ' поиск всплесков на останове через XXсек после начала падения оборотов

                        Dim xval(), xval_new(NumericUpDown3.Value * Nt), yval(,), yval_new(NumericUpDown3.Value * Nt, 1), yval_Real() As Double
                        Dim l As Integer
                        Dim Np As Single = NumericUpDown3.Value * Nt

                        xval = IB_regX.XValues
                        yval = IB_regX.YValues
                        yval_Real = IB_regX.UserYValues(LMSTestLabAutomation.CONST_ScaleProperty.Real_Scale)

                        MinSp = 0
                        MaxSp = 0

                        If Rtable(k) = "0" Then

                            For l = UBound(xval) To LBound(xval) Step -1

                                If Np >= 0 Then
                                    xval_new(Np) = Nt
                                    yval_new(Np, 0) = yval(l, 0)
                                End If

                                Np = Np - 1
                                Nt = Nt - 1 / NumericUpDown3.Value

                                ' поиск всплесков до роста обортов (в интервале за XX секунды до выхода на режим) и запись их в свойства
                                If Nt > SpTimStart Then
                                    If yval_Real(l) > MaxSp Then MaxSp = yval_Real(l)
                                    If yval_Real(l) < MinSp Then MinSp = yval_Real(l)
                                End If
                            Next

                        ElseIf Rtable(k) = "2" Then

                            For l = LBound(xval) To NumericUpDown3.Value * 5    'UBound(xval)
                                If l <= UBound(xval) Then yval_new(l, 0) = yval(l, 0)
                                xval_new(l) = Nt
                                Nt = Nt + 1 / NumericUpDown3.Value

                                ' поиск всплесков после падения обортов (в интервале после 1сек после останова) и запись их в свойства
                                If Nt > SpTimShDown And l <= UBound(xval) Then
                                    If yval_Real(l) > MaxSp Then MaxSp = yval_Real(l)
                                    If yval_Real(l) < MinSp Then MinSp = yval_Real(l)
                                End If
                            Next
                        End If

                        IB_regX = IB_regX.ReplaceXDoubleValues(xval_new)
                        IB_regX = IB_regX.ReplaceYComplexValues(yval_new)

                    End If

                    ' расчет среднего значения по режимам
                    AM = TL.Factory.CreateAttributeMap
                    StatFunc = TL.Factory.CreateEnumerate("LmsHq::DataModelI::BlockStatistics::CBufferIEnumStatisticFunction", 10)
                    AM.Add("SourceBlock", IB_reg)
                    AM.Add("StatisticFunction", StatFunc)
                    vscalar = TL.Factory.CreateObject("LmsHq::DataModelC::BlockStatistics::CStatisticBlockMetric", AM)
                    AM = Nothing

                    ' центровка режима
                    IB_cent_reg = TL.cmd.BLOCK_SUBTRACT_SCALAR(IB_reg, vscalar)
                    TextBox3.Text = TX & vbTab & "#" & m & "/" & k : Me.Refresh()

                    'DB.AddItem(Pname & "/", Pname & "Cent", IB_cent_reg, , 1)

                    'вычисление вероятностных характеристик режима
                    Dim yval_cent_Real() As Double
                    Dim Mx, Sx, D, Sk, Excess, N, MaxSp2, MinSp2 As Single

                    MinSp2 = 0
                    MaxSp2 = 0

                    yval_cent_Real = IB_cent_reg.UserYValues(LMSTestLabAutomation.CONST_ScaleProperty.Real_Scale)
                    N = UBound(yval_cent_Real) + 1

                    For l = LBound(yval_cent_Real) To N - 1
                        Mx = Mx + yval_cent_Real(l)
                        D = D + yval_cent_Real(l) ^ 2
                        Sk = Sk + yval_cent_Real(l) ^ 3
                        Excess = Excess + yval_cent_Real(l) ^ 4
                        If MaxSp2 < yval_cent_Real(l) And CSng(Rtable(k)) > 6000 Then MaxSp2 = yval_cent_Real(l) 'всплеск в +
                        If MinSp2 > yval_cent_Real(l) And CSng(Rtable(k)) > 6000 Then MinSp2 = yval_cent_Real(l) 'всплеск в -
                    Next

                    Mx = Mx / N
                    D = D / (N - 1)                         'Дисперсия
                    Sx = Math.Sqrt(D)                       'СКО
                    Sk = Sk / (N * Sx ^ 3)                  'Асимметрия
                    Excess = (Excess / (N * Sx ^ 4)) - 3    'Эксцесс

                    'при превышении установить ключ для записи сигнала
                    Dim WritKey As Integer = 0
                    If Math.Abs(Sk) >= Val(TextBox12.Text) And CheckBox8.Checked = True Then
                        WritKey = 1
                    ElseIf Excess >= Val(TextBox13.Text) And CheckBox9.Checked = True Then
                        WritKey = 1
                    End If

                    ' вычисление осредненного спектра
                    Dim AmpScale As LMSTestLabAutomation.Enumerate
                    Dim SpecFormat As LMSTestLabAutomation.Enumerate
                    Dim Windowing As LMSTestLabAutomation.Enumerate
                    Dim Fres As Single

                    AmpScale = TL.Factory.CreateEnumerate("LmsHq::DataModelI::DataAttributes::CBufferIEnumAmplitudeScaling", FAmplScaling(ComboBox4.Text))
                    SpecFormat = TL.Factory.CreateEnumerate("LmsHq::DataModelI::DataAttributes::CBufferIEnumSpectrumFormat", FSpecFormat(ComboBox5.Text))
                    Windowing = TL.Factory.CreateEnumerate("LmsHq::DataModelI::DataAttributes::CBufferIEnumWindowing", FWindowing(ComboBox3.Text))

                    AM = TL.Factory.CreateAttributeMap
                    AM.Add("SourceBlock", IB_cent_reg)
                    AM.Add("EnumAmplitudeScaling", AmpScale)
                    AM.Add("EnumSpectrumFormat", SpecFormat)
                    AM.Add("EnumWindowing", Windowing)
                    AM.Add("BlockSize", CInt(ComboBox1.Text))
                    AM.Add("Overlap", CInt(CSng(ComboBox1.Text) * (CSng(ComboBox2.Text) / 100)))
                    IB_Power = TL.Factory.CreateObject("LmsHq::DataModelC::SignalFunctions::CAveragedSpectrum", AM)
                    AM = Nothing

                    Fres = CSng(Val(IB_Power.Properties("Frequency resolution"))) ' взять частотное разрешение сигнала

                    'If Rtable(k) = Rtable(k - 1) Then ' осреднение со следующим интервалом времени если номер режима совпадает
                    'Dim vscal1, vscal2 As LMSTestLabAutomation.IScalar
                    'Dim vqunt As LMSTestLabAutomation.IQuantity

                    'Kcoup += 1
                    'vqunt = TL.UnitSystem.QuantityRatio
                    'vscal1 = TL.cmd.DOUBLE_TO_SCALAR(1 / (Kcoup - 1), vqunt)
                    'vscal2 = TL.cmd.DOUBLE_TO_SCALAR(Kcoup, vqunt)

                    'IB_aver = TL.cmd.BLOCK_DIVIDE_SCALAR(IB_aver, vscal1)       ' убрать осреднение
                    'IB_aver = TL.cmd.BLOCK_ADD(IB_aver, IB_Power)               ' сложить с предыдущим блоком
                    'IB_Power = TL.cmd.BLOCK_DIVIDE_SCALAR(IB_aver, vscal2)      ' осреднить
                    'IB_aver = IB_Power                                          ' сохранить для последующего возможного осреднения
                    'End If








                    If CSng(Rtable(k)) > 6000 Then  ' если режим стационарный

                        R = CStr(Math.Round(Ttable(k, 3), 1))
                        RPM = CStr(Math.Round(Ttable(k, 4)))
                        km = CStr(Math.Round(Ttable(k, 5), 2))
                        dT = CStr(Math.Round(Ttable(k, 2) - Ttable(k, 1), 1))

                        Dim IBName As String = Rtable(k) & " РП=" & R & " КПВ=" & RPM & " Km=" & km

                        DB.AddItem(Pname & "/", IBName, IB_Power, , 1)  ' запись спектра мощности
                        AM = DB.GetProperties(Pname & "/" & IBName)     ' запись дополнительных свойств
                        AM.Add("R", R)
                        AM.Add("RPM", RPM)
                        AM.Add("km", km)
                        AM.Add("dT", dT)
                        AM.Add("MaxSp", Math.Round(MaxSp2, 4))
                        AM.Add("MinSp", Math.Round(MinSp2, 4))
                        AM.Add("Sx", Math.Round(Sx, 4))
                        AM.Add("D", Math.Round(D, 4))
                        AM.Add("Sk", Math.Round(Sk, 4))
                        AM.Add("Excess", Math.Round(Excess, 4))
                        DB.AddProperties(Pname & "/" & IBName, AM, 0)
                        AM = Nothing

                        If CheckBox7.Checked = True Or WritKey = 1 Then    ' запись сигнала. если режим будет повторяться, то записан будет последний сигнал
                            DB.AddItem(Pname & "/", "Sig " & IBName, IB_cent_reg, , 1)
                            AM = DB.GetProperties(Pname & "/Sig " & IBName)
                            AM.Add("R", R)
                            AM.Add("RPM", RPM)
                            AM.Add("km", km)
                            AM.Add("dT", dT)
                            AM.Add("MaxSp", Math.Round(MaxSp2, 4))
                            AM.Add("MinSp", Math.Round(MinSp2, 4))
                            AM.Add("Sx", Math.Round(Sx, 4))
                            AM.Add("D", Math.Round(D, 4))
                            AM.Add("Sk", Math.Round(Sk, 4))
                            AM.Add("Excess", Math.Round(Excess, 4))
                            DB.AddProperties(Pname & "/Sig " & IBName, AM, 0)
                            AM = Nothing
                        End If








                        GoTo end_OMA
                        ' IBName - код режимa
                        ' Pname - имя параметра БМП
                        If CheckBox12.Checked = True Then    ' запись сигнала для обработки ОМА

                            Dim Point_ID As String = ""
                            Dim Direction As Integer

                            Select Case Pname
                                '1+X, 2-X, 3+Y, 4-Y, 5+Z, 6-Z
                                Case "21VOG"
                                    Point_ID = "rd191:gg"
                                    Direction = 3
                                Case "21VPBG"
                                    Point_ID = "rd191:bnag"
                                    Direction = 3
                                Case "21VPBO"
                                    Point_ID = "rd191:bnao"
                                    Direction = 5
                                Case "21VPG"
                                    Point_ID = "rd191:gg"
                                    Direction = 6
                                Case "21VPNG"
                                    Point_ID = "rd191:ng"
                                    Direction = 3
                                Case "21VPNO"
                                    Point_ID = "rd191:no"
                                    Direction = 6
                                Case "VONG"
                                    Point_ID = "rd191:ng"
                                    Direction = 2
                                Case "VPBG2"
                                    Point_ID = "rd191:bnag"
                                    Direction = 6
                                Case "VPBO2"
                                    Point_ID = "rd191:bnao"
                                    Direction = 4
                                Case "VPG2"
                                    Point_ID = "rd191:gg"
                                    Direction = 2
                                Case "VPNG2"
                                    Point_ID = "rd191:ng"
                                    Direction = 5
                                Case "VPNO2"
                                    Point_ID = "rd191:no"
                                    Direction = 3
                                Case "21VPRG"
                                    Point_ID = "rd191:rg"
                                    Direction = 1
                                Case "21VPDG"
                                    Point_ID = "rd191:dg"
                                    Direction = 1
                            End Select

                            If Point_ID <> "" Then
                                R = CStr(Math.Round(Ttable(k, 3), 3))

                                AM = TL.Factory.CreateAttributeMap

                                Dim Dirc As LMSTestLabAutomation.Enumerate
                                Dirc = TL.Factory.CreateEnumerate("LmsHq::DataModelI::Channel::CBufferIEnumDirections", Direction)

                                AM.Add("OriginalBlock", IB_cent_reg)
                                AM.Add("PointDirection", Dirc)
                                AM.Add("PointID", Point_ID)

                                IB_cent_reg_new = TL.Factory.CreateObject("LmsHq::DataModelC::ExprEditing::CBlockChangeAttributes", AM)

                                AM = Nothing

                                Try
                                    DB.AddSection("OMA")
                                Catch exx As Exception
                                End Try

                                Try
                                    DB.AddFolder("OMA", "Sig_R" & R)
                                Catch exx As Exception
                                End Try

                                DB.AddItem("OMA/Sig_R" & R & "/", Pname, IB_cent_reg_new, , 0)

                            End If
                        End If
end_OMA:









                    Else ' запись спектра по запуску/останову
                        ' запись дополнительных свойств по режимам запуска и останова
                        DB.AddItem(Pname & "/", Rtable(k), IB_Power, , 1)
                        AM = DB.GetProperties(Pname & "/" & Rtable(k))
                        AM.Add("MaxSp", Math.Round(MaxSp, 4))
                        AM.Add("MinSp", Math.Round(MinSp, 4))
                        AM.Add("Sx", Math.Round(Sx, 4))
                        AM.Add("D", Math.Round(D, 4))
                        AM.Add("Sk", Math.Round(Sk, 4))
                        AM.Add("Excess", Math.Round(Excess, 4))
                        DB.AddProperties(Pname & "/" & Rtable(k), AM, 0)

                        If CheckBox6.Checked = True Then ' запись сигнала
                            Dim IBName As String = "Sig " & Rtable(k) & " MaxSp=" & CStr(Math.Round(MaxSp, 1)) & " MinSp=" & CStr(Math.Round(MinSp, 1))
                            DB.AddItem(Pname & "/", IBName, IB_regX, , 1)

                            ' запись всплесков в свойства 
                            AM = DB.GetProperties(Pname & "/" & IBName)
                            AM.Add("MaxSp", Math.Round(MaxSp, 4))
                            AM.Add("MinSp", Math.Round(MinSp, 4))
                            AM.Add("Sx", Math.Round(Sx, 4))
                            AM.Add("D", Math.Round(D, 4))
                            AM.Add("Sk", Math.Round(Sk, 4))
                            AM.Add("Excess", Math.Round(Excess, 4))
                            DB.AddProperties(Pname & "/" & IBName, AM, 0)
                            AM = Nothing
                        End If

                    End If

                Next k
                TextBox3.Text = TextBox3.Text & vbCrLf
                Label13.Text = CInt(Label13.Text) + 1 : Me.Refresh()
                Exit For
            End If
        Next

        If labs(i) <> Pname Then
            TextBox3.Text = TextBox3.Text & vbTab & "absent" & vbCrLf : Me.Refresh()
        End If

    End Sub

    Function TransformSignal(ByVal IB As LMSTestLabAutomation.IBlock2, ByVal ID_sinus As LMSTestLabAutomation.IData, ByVal TL As LMSTestLabAutomation.Application, ByVal tStart As Single, ByVal pNumber As Integer, ByVal Ky As Single, ByVal Kd As Single, ByVal Kp As Single)

        Dim AM As LMSTestLabAutomation.AttributeMap
        Dim IB_sinus As LMSTestLabAutomation.IBlock2


        AM = TL.Factory.CreateAttributeMap
        AM.Add("StartValueGiven", True)
        AM.Add("StartValue", tStart)
        AM.Add("NumberOfPoints", pNumber)
        AM.Add("BlockStream", ID_sinus)
        IB_sinus = TL.Factory.CreateObject("LmsHq::DataModelC::ExprConcat::CBlockStreamToBlock", AM)
        AM = Nothing

        Dim SinMax_scal, SinMin_scal As LMSTestLabAutomation.IScalar
        Dim SinMax, SinMin, imag As Double
        Dim SinStatFunc As LMSTestLabAutomation.Enumerate

        ' расчет max min
        AM = TL.Factory.CreateAttributeMap
        SinStatFunc = TL.Factory.CreateEnumerate("LmsHq::DataModelI::BlockStatistics::CBufferIEnumStatisticFunction", 0)
        AM.Add("SourceBlock", IB_sinus)
        AM.Add("StatisticFunction", SinStatFunc)
        SinMax_scal = TL.Factory.CreateObject("LmsHq::DataModelC::BlockStatistics::CStatisticBlockMetric", AM)
        AM = Nothing

        AM = TL.Factory.CreateAttributeMap
        SinStatFunc = TL.Factory.CreateEnumerate("LmsHq::DataModelI::BlockStatistics::CBufferIEnumStatisticFunction", 3)
        AM.Add("SourceBlock", IB_sinus)
        AM.Add("StatisticFunction", SinStatFunc)
        SinMin_scal = TL.Factory.CreateObject("LmsHq::DataModelC::BlockStatistics::CStatisticBlockMetric", AM)
        AM = Nothing

        SinMax_scal.GetValue(SinMax, imag)
        SinMin_scal.GetValue(SinMin, imag)

        Dim yval_Signal() As Double
        Dim Ns As Integer

        yval_Signal = IB.UserYValues(LMSTestLabAutomation.CONST_ScaleProperty.Real_Scale)
        Ns = UBound(yval_Signal) + 1

        Dim yval(Ns - 1) As Double

        For l = LBound(yval_Signal) To Ns - 1
            yval(l) = (yval_Signal(l) - (SinMax - SinMin) / 2) * (6000 / (Ky * Kd * Kp * (SinMax - SinMin)))
        Next

        IB = IB.ReplaceYDoubleValues(yval)

        Return (IB)

    End Function

    Function FWindowing(ByVal WindT As String)
        Dim i As Integer

        Select Case WindT
            Case "Uniform"
                i = 0
            Case "Hanning"
                i = 1
            Case "Hamming"
                i = 2
            Case "Flattop"
                i = 8
        End Select

        Return i
    End Function

    Function FAmplScaling(ByVal AmScT As String)
        Dim i As Integer

        Select Case AmScT
            Case "peak"
                i = 0
            Case "RMS"
                i = 1
            Case "double sided"
                i = 2
        End Select

        Return i
    End Function

    Function FSpecFormat(ByVal SpecFrT As String)
        Dim i As Integer

        Select Case SpecFrT
            Case "linear"
                i = 0
            Case "Power"
                i = 1
            Case "PSD"
                i = 2
            Case "ESD"
                i = 3
        End Select

        Return i
    End Function

    ' расчет ВАХ
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Names()

        Dim EAdat As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EAres As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EAVAC As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EWres As Microsoft.Office.Interop.Excel.Workbook
        Dim ESres As Microsoft.Office.Interop.Excel.Worksheet
        Dim k, l, m, n, o, t, cSign, KeyP, Rclm As Integer
        Dim Ttable(500, 10) As Single
        Dim Rtable(500), VAC(40), Pname As String
        Dim PtLMS As String = TextBox1.Text
        Dim PtXLS As String = TextBox2.Text

        TextBox3.Text = "" : Label14.Text = 0
        ToolStripStatusLabel1.Text = "Reading passport-file" : Me.Refresh()

        ' открытие паспорта
        Try
            EAdat.Workbooks.Open(PtXLS)
        Catch ex As Exception
            MsgBox("File " & PtXLS & " does not exist.")
            Exit Sub
        End Try

        Try
            EAVAC.Workbooks.Open(PtVAC)
        Catch ex As Exception
            MsgBox("File " & PtVAC & " does not exist.")
            Exit Sub
        End Try

        EWres = EAres.Workbooks.Add()


        If CheckBox10.Checked = True Then
            EAdat.Visible = False
            EAres.Visible = False
        Else
            EAdat.Visible = True
            EAres.Visible = True
        End If


        Try
            EAdat.Worksheets("MAIN").Select()
        Catch ex As Exception
            MsgBox("Sheet MAIN does not exist.")
            Exit Sub
        End Try

        ' определение числа обрабатываемых параметров
        Do
            KeyP = CSng(EAdat.Cells(PnRow + 1, PnClm + k).Value)
            If CSng(EAdat.Cells(PnRow + 1, PnClm + k).Value) > 0 Then
                cSign += 1
            End If
            k += 1
        Loop While CStr(EAdat.Cells(PnRow, PnClm + k).Value) <> ""

        k = 0
        Do
            KeyP = CSng(EAdat.Cells(PnRow + 1, PnClm + k).Value)   ' ключ обработки параметра
            Pname = CStr(EAdat.Cells(PnRow, PnClm + k).Value)
            Rclm = CInt(EAVAC.Cells(3, 2 + k).Value)           ' номер столбика с оборотами

            l = 2
            t = 0
            Do
                If CStr(EAVAC.Cells(l + 2, 2 + k).Value) <> "" Then
                    VAC(t + 1) = CStr(EAVAC.Cells(l + 2, 1).Value)  ' запись ВАХ, t - количество ВАХ
                    t += 1
                End If
                l += 1
            Loop While CStr(EAVAC.Cells(l + 2, 1).Value) <> ""

            If KeyP = 2 Then

                Try
                    EAdat.Worksheets(Pname).Select()
                Catch ex As Exception
                    MsgBox("Sheet " & Pname & " does not exist.")
                    Exit Sub
                End Try

                n = k
                PnRow = CSng(EAdat.Cells(1, 1).Value)
                PnClm = CSng(EAdat.Cells(1, 2).Value)
            End If

            If KeyP > 0 Then
                TextBox3.Text = TextBox3.Text & Pname : Me.Refresh()

                l = 0
                Do
                    KeyP = CSng(EAdat.Cells(l + PnRow + 2, PnClm + k - n).Value)
                    If KeyP = 1 Then                                                    ' ключ обработки режима
                        Ttable(m + 1, 1) = CSng(EAdat.Cells(PnRow + 2 + l, 3).Value)    ' время начало обработки
                        Ttable(m + 1, 2) = CSng(EAdat.Cells(PnRow + 2 + l, 4).Value)    ' время конца обработки
                        Ttable(m + 1, 3) = CSng(EAdat.Cells(PnRow + 2 + l, 5).Value)    ' уровень тяги

                        ' скорость вращения
                        Ttable(m + 1, 4) = CSng(EAdat.Cells(PnRow + 2 + l, Rclm).Value)

                        Ttable(m + 1, 5) = CSng(EAdat.Cells(PnRow + 2 + l, 7).Value)    ' Km
                        Rtable(m + 1) = CStr(EAdat.Cells(PnRow + 2 + l, 1).Value)       ' шифр режима
                        m += 1
                    End If
                    l += 1
                Loop While CStr(EAdat.Cells(PnRow + 2 + l, 1).Value) <> ""  ' проверка конца режимов

                If m > 0 Then
                    o += 1
                    ESres = EWres.Worksheets.Add
                    ESres.Name = Pname
                    Call SpectrumsProcessing(m, o, Pname, Ttable, Rtable, ESres, VAC, t)
                    ToolStripProgressBar1.Value = (o / cSign) * 100
                End If
            ElseIf KeyP = 0 Then
                TextBox3.Text = TextBox3.Text & Pname & vbTab & "skipped" & vbCrLf : Me.Refresh()
            End If

            t = 0
            m = 0
            l = 0
            n = 0
            k += 1
            EAdat.Worksheets("MAIN").Select()
        Loop While CStr(EAdat.Cells(PnRow, PnClm + k).Value) <> ""  ' проверка конца режимов

        EAdat.Quit()
        EAVAC.Quit()

        If CheckBox1.Checked = True Then
            Try
                System.IO.File.Delete(TextBox5.Text)
            Catch ex As Exception
            End Try
            EWres.SaveAs(TextBox5.Text)
        End If

        If CheckBox2.Checked = True And o > 0 Then
            Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
            Dim DB As LMSTestLabAutomation.IDatabase
            DB = TL.ActiveBook.Database
            ToolStripStatusLabel1.Text = "Saving LMS file" : Me.Refresh()
            TL.ActiveBook.Save(TextBox1.Text)
        End If

        ToolStripStatusLabel1.Text = "Finished"
        ToolStripProgressBar1.Value = 0 : Me.Refresh()

        EAres.Quit()

        If CheckBox4.Checked = True And CheckBox12.Checked = False Then
            'EAres.Quit()
            TabControl1.SelectedTab = TabPage3
            TabPage3_Enter(sender, e)
            Button3_Click(sender, e)
        End If

    End Sub

    Private Sub SpectrumsProcessing(ByVal m As Integer, ByVal o As Integer, ByVal Pname As String, ByVal Ttable(,) As Single, ByVal Rtable() As String, ByVal ESres As Object, ByVal VAC() As String, ByVal t As Integer)
        Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
        Dim DB As LMSTestLabAutomation.IDatabase
        Dim PtXLS As String = TextBox5.Text

        If TL.Name <> "" And o = 1 Then
            TL.ActiveBook.Close()

            Try
                TL.OpenProject(TextBox1.Text)
            Catch ex As Exception
                MsgBox("File " & TextBox1.Text & " does not exist.")
                Exit Sub
            End Try

        ElseIf TL.Name = "" Then
            ToolStripStatusLabel1.Text = "Opening TestLab..." : Me.Refresh()

            Try
                TL.Init("-w DesktopStandard " & TextBox1.Text)
            Catch ex As Exception
                MsgBox("File " & TextBox1.Text & " does not exist.")
                Exit Sub
            End Try

        End If
        DB = TL.ActiveBook.Database

        Dim i, k, l, p, r, u, uu, z, Pcnt As Integer
        Dim IB As LMSTestLabAutomation.IBlock2
        Dim AM_IB, AM As LMSTestLabAutomation.AttributeMap
        Dim Fres As Single
        Dim IBname As String

        For i = 0 To DB.SectionNames.Count - 1
            If DB.SectionNames.Item(i) = Pname Then
                ToolStripStatusLabel1.Text = "Processing: " & Pname : Me.Refresh()

                Dim TX As String = TextBox3.Text

                ESres.Cells(1, 1) = Pname   ' имя параметра
                ESres.Cells(1, 2) = "Regime"
                ESres.Cells(1, 3) = "P%"
                ESres.Cells(1, 4) = "RPM"
                ESres.Cells(1, 5) = "RPS"

                For k = 1 To m ' проход по каждому режиму

                    For z = 0 To DB.ElementNames(Pname).Count   ' поиск блока с именем режима
                        IBname = DB.ElementNames(Pname).Item(z)
                        If Rtable(k) = Strings.Left(IBname, 4) Or Rtable(k) = IBname Then Exit For
                        IBname = ""
                    Next z

                    If IBname = "" Then GoTo m1 ' во избежании обработки не того режима

                    u = ClmE   ' запись роторных начать с этого столбика
                    TextBox3.Text = TX & vbTab & "#" & m & "/" & k
                    ToolStripStatusLabel1.Text = "Processing: " & Pname

                    IB = DB.GetItem(Pname & "/" & IBname)
                    AM_IB = DB.GetProperties(Pname & "/" & IBname)

                    Fres = CSng(Val(AM_IB.Item("Frequency resolution")))
                    Label31.Text = Math.Round(Fres, 3) : Me.Refresh()

                    Dim xval(), yval() As Double
                    Dim Meff, RPM, RPS, Power, Regime As Single
                    Dim Eeff, Eeff2, Emax, EHz, E, MaxSp, MinSp, Sx, Sk, dT, T0, km, Excess As Single

                    xval = IB.XValues
                    yval = IB.UserYValues(LMSTestLabAutomation.CONST_ScaleProperty.Real_Scale)

                    ESres.Cells(k + 1, 1) = Rtable(k)                           ' код режима
                    Regime = TakeRegime(Rtable(k))

                    If Rtable(k) < 7000 Then ESres.Cells(k + 1, 2) = Regime ' режим
                    If Rtable(k) >= 7000 Then ESres.Cells(k + 1, 2) = Math.Round(Ttable(k, 3), 1) / 100.0 ' режим

                    If Regime = 0 Then
                        ESres.Cells(k + 1, 2) = "0"
                    ElseIf Regime = 2 Then
                        ESres.Cells(k + 1, 2) = "2"
                    End If

                    If CSng(Rtable(k)) > 6000 Then     ' расчет только для стационарных режимов
                        Power = Math.Round(Ttable(k, 3), 1)
                        ESres.Cells(k + 1, 3) = Power       ' P%

                        RPM = Math.Round(Ttable(k, 4), 1)
                        ESres.Cells(k + 1, 4) = RPM         ' RPM

                        RPS = Math.Round(RPM / 60, 1)
                        ESres.Cells(k + 1, 5) = RPS         ' RPS

                        T0 = Math.Round(Ttable(k, 1), 1)    ' T0

                        ' расчет 0.7 роторной
                        EHz = (RPM / 60) * 0.7
                        Pcnt = Val(TextBox11.Text) / (2 * Fres)    ' поиск полюса в районе гармоники (30Hz) роторной частоты +-Pcnt (Pcnt - количество гармоник)

                        Emax = 0
                        For p = Math.Floor(EHz / Fres) - Pcnt To Math.Ceiling(EHz / Fres) + Pcnt
                            If Emax < yval(p) Then
                                Emax = yval(p)  ' максимальная амплитуда
                                r = p           ' номер гармоники
                                E = xval(p) / (RPM / 60)
                            End If
                        Next

                        Eeff = 0
                        ' расчет дисперсии полюса у максимума
                        For p = r - Math.Floor(CountLines(r, RPM) / 2) To r + Math.Floor(CountLines(r, RPM) / 2)
                            Eeff = Eeff + yval(p)   ' сумма RMS
                        Next
                        Eeff = Math.Sqrt(Eeff)

                        Eeff2 = 0
                        ' расчет дисперсии полюса четко у роторной
                        For p = Math.Floor(EHz / Fres) - Math.Floor(CountLines(r, RPM) / 2) To Math.Floor(EHz / Fres) + Math.Floor(CountLines(r, RPM) / 2)
                            Eeff2 = Eeff2 + yval(p)   ' сумма RMS
                        Next
                        Eeff2 = Math.Sqrt(Eeff2)

                        If Eeff2 > Eeff Then ' выбираем максимальное 
                            Eeff = Eeff2
                            r = Math.Floor(EHz / Fres)
                        End If

                        ESres.Cells(1, 39) = Math.Round(EHz / (RPM / 60), 1)
                        ESres.Cells(k + 1, 39) = Math.Round(Eeff, 3)         ' запись эффективного значения роторной
                        Eeff = 0 : Eeff2 = 0 : Emax = 0




                        ' поиск полюсов и расчет дисперсии полюсов на роторных частотах
                        For EHz = RPM / 60 To Val(TextBox8.Text) Step RPM / 60  ' с первой роторной до верхней границы фильтра
                            'For EHz = RPM / 60 To Val(8000) Step RPM / 60  ' с первой роторной до верхней границы фильтра


                            'For EHz = 0 To Val(TextBox8.Text) Step RPM / 60  ' с 0.7 роторной до верхней границы фильтра
                            'If EHz = 0 Then
                            'EHz = (RPM / 60) * 0.7
                            'End If


                            Pcnt = Val(TextBox11.Text) / (2 * Fres)    ' поиск полюса в районе гармоники (30Hz) роторной частоты +-Pcnt (Pcnt - количество гармоник)
                            Emax = 0
                            For p = Math.Floor(EHz / Fres) - Pcnt To Math.Ceiling(EHz / Fres) + Pcnt
                                If Emax < yval(p) Then
                                    Emax = yval(p)  ' максимальная амплитуда
                                    r = p           ' номер гармоники
                                    E = xval(p) / EHz
                                End If
                            Next

                            Eeff = 0
                            ' расчет дисперсии полюса у максимума
                            For p = r - Math.Floor(CountLines(r, RPM) / 2) To r + Math.Floor(CountLines(r, RPM) / 2)
                                Eeff = Eeff + yval(p)   ' сумма RMS
                            Next
                            Eeff = Math.Sqrt(Eeff)

                            Eeff2 = 0
                            ' расчет дисперсии полюса четко у роторной
                            For p = Math.Floor(EHz / Fres) - Math.Floor(CountLines(r, RPM) / 2) To Math.Floor(EHz / Fres) + Math.Floor(CountLines(r, RPM) / 2)
                                Eeff2 = Eeff2 + yval(p)   ' сумма RMS
                            Next
                            Eeff2 = Math.Sqrt(Eeff2)

                            If Eeff2 > Eeff Then ' выбираем максимальное 
                                Eeff = Eeff2
                                r = Math.Floor(EHz / Fres)
                            End If

                            'f EHz = (RPM / 60) * 0.7 Then
                            'ESres.Cells(1, 36) = Math.Round(EHz / (RPM / 60), 1)
                            'ESres.Cells(k + 1, 36) = Math.Round(Eeff, 3)         ' запись эффективного значения роторной
                            'ESres.Cells(k + 1, 37) = CountLines(r)               ' число линий для расчета эф. амплитуды
                            'ESres.Cells(k + 1, 38) = Math.Round(E, 2)            ' отношение частоты пика к роторной
                            'ESres.Cells(k + 1, 39) = Math.Round(xval(r), 2)      ' частота пика
                            'u += 1
                            'Else


                            ESres.Cells(1, u) = Math.Round(EHz / (RPM / 60))
                            ESres.Cells(k + 1, u) = Math.Round(Eeff, 3)         ' запись эффективного значения роторной
                            u += 1


                            'End If
                            'ESres.Cells(k + 2, u) = CountLines(r)               ' число линий для расчета эф. амплитуды
                            'ESres.Cells(k + 3, u) = Math.Round(E, 2)            ' отношение частоты пика к роторной
                            'ESres.Cells(k + 4, u) = Math.Round(xval(r), 2)      ' частота пика

                            Eeff = 0 : Eeff2 = 0 : Emax = 0
                        Next








                        Dim Asum, AnE, AF, Abl, Anoise, AA, Ads, Ads2, An, Fn As Single
                        Dim nE As Integer
                        u = 6
                        uu = ClmE
                        Abl = 0

                        For s = 1 To t  ' t - количество ДП
                            Meff = 0
                            ESres.Cells(1, u) = VAC(s)  ' запись шапки

                            If Strings.Right(VAC(s), 1) = "E" Then
                                nE = Val(Mid(VAC(s).ToString, 2, Len(VAC(s).ToString) - 2))
                                AnE = Math.Round(ESres.Cells(k + 1, uu - 1 + nE).Value, 3)
                                ESres.Cells(k + 1, u) = AnE
                                Abl = Abl + AnE * AnE

                            ElseIf Strings.Left(VAC(s), 1) = "T0" Then
                                ESres.Cells(k + 1, u) = T0

                                ' Расчет подшипниковых составляющих
                            ElseIf Strings.Left(VAC(s), 1) = "n" Then
                                Fn = RPS * Val(Strings.Right(VAC(s), Strings.Len(VAC(s)) - 1)) / 100.0
                                An = 0

                                For p = Math.Floor(Fn / Fres) - Pcnt To Math.Ceiling(Fn / Fres) + Pcnt
                                    If p <= 0 Then p = 1
                                    If An < yval(p) Then
                                        An = yval(p)  ' максимальная амплитуда
                                    End If
                                Next
                                ESres.Cells(k + 1, u) = Math.Sqrt(An)

                            Else
                                Select Case VAC(s)

                                    Case "Asum"
                                        ' расчет эффективного значения суммарного сигнала
                                        For l = Math.Round(Val(TextBox7.Text) / Fres) To Math.Round(Val(TextBox8.Text) / Fres)
                                            Meff = Meff + yval(l)
                                        Next
                                        Asum = Math.Sqrt(Meff)
                                        ESres.Cells(k + 1, 6) = Math.Round(Asum, 3)
                                        AA = Asum * Asum

                                        'запись вероятностных характеристик
                                        ESres.Cells(1, 7) = "MaxSp"
                                        MaxSp = CSng(Val(AM_IB.Item("MaxSp")))
                                        ESres.Cells(k + 1, 7) = MaxSp

                                        ESres.Cells(1, 8) = "MinSp"
                                        MinSp = CSng(Val(AM_IB.Item("MinSp")))
                                        ESres.Cells(k + 1, 8) = MinSp

                                        ESres.Cells(1, 9) = "Sx"
                                        Sx = CSng(Val(AM_IB.Item("Sx")))
                                        ESres.Cells(k + 1, 9) = Sx

                                        ESres.Cells(1, 10) = "Sk"
                                        Sk = CSng(Val(AM_IB.Item("Sk")))
                                        ESres.Cells(k + 1, 10) = Sk
                                        If Math.Abs(Sk) >= Val(TextBox12.Text) And CheckBox8.Checked = True Then ESres.Cells(k + 1, 10).Interior.Color = 200

                                        ESres.Cells(1, 11) = "Excess"
                                        Excess = CSng(Val(AM_IB.Item("Excess")))
                                        ESres.Cells(k + 1, 11) = Excess
                                        If Excess >= Val(TextBox13.Text) And CheckBox9.Checked = True Then ESres.Cells(k + 1, 11).Interior.Color = 200

                                        ESres.Cells(1, 12) = "dT"
                                        dT = CSng(Val(AM_IB.Item("dT")))
                                        ESres.Cells(k + 1, 12) = dT

                                        ESres.Cells(1, 13) = "km"
                                        km = CSng(Val(AM_IB.Item("km")))
                                        ESres.Cells(k + 1, 13) = km

                                        u = 13

                                    Case "Anoise"
                                        If AA > Abl Then
                                            AA = Math.Round(Math.Sqrt(AA - Abl), 3)
                                        Else
                                            AA = 0
                                        End If

                                        ESres.Cells(k + 1, u) = AA
                                        Anoise = AA

                                    Case "Ads"
                                        Ads = Math.Round(((Asum - Anoise) / Asum) * 100, 1)
                                        ESres.Cells(k + 1, u) = Ads

                                    Case "Abl"
                                        ESres.Cells(k + 1, u) = Math.Round(Math.Sqrt(Abl), 3)

                                        'Case "A07"
                                        'AF = Math.Round(Acust(yval, Power, RPS, VAC(s), Fres, Regime), 3)
                                        'ESres.Cells(k + 1, u) = AF
                                        'u += 1
                                        'ESres.Cells(1, u) = "F" & Mid(VAC(s), 2, 5)
                                        'ESres.Cells(k + 1, u) = Math.Round(AFreq, 1)

                                    Case "A5"
                                        ESres.Cells(k + 1, u) = Damage(yval, VAC(s), RPS, Fres, dT)

                                    Case "A10"
                                        ESres.Cells(k + 1, u) = Damage(yval, VAC(s), RPS, Fres, dT)

                                    Case "A14"
                                        ESres.Cells(k + 1, u) = Damage(yval, VAC(s), RPS, Fres, dT)

                                    Case "A33"
                                        ESres.Cells(k + 1, u) = Damage(yval, VAC(s), RPS, Fres, dT)

                                    Case "T0"
                                        ESres.Cells(k + 1, u) = T0

                                    Case Else   ' считать акустику по КС и ГГ
                                        AF = Math.Round(Acust(yval, Power, RPS, VAC(s), Fres, Regime), 3)
                                        ESres.Cells(k + 1, u) = AF
                                        AA = AA - AF * AF

                                        ' убрать из расчета Ашум составляющую А2750 двигателей РД191 и РД181
                                        If VAC(s) = "A2750" And ComboBox6.Text = "РД191" Then AA = AA + AF * AF

                                        u += 1
                                        ESres.Cells(1, u) = "F" & Mid(VAC(s), 2, 5)
                                        ESres.Cells(k + 1, u) = Math.Round(AFreq, 1)

                                End Select

                            End If

                            u += 1
                        Next ' следующий ДП







                        ' приведение блока СПМ в амплитудный
                        AM = TL.Factory.CreateAttributeMap
                        AM.Add("SourceBlock", IB)
                        IB = TL.Factory.CreateObject("LmsHq::DataModelC::MathProcessing::CSquareRootBlock", AM)
                        DB.AddItem(Pname & "/", IBname & " Aсум=" & CStr(Math.Round(Asum, 2)) & " Aшум=" & CStr(Math.Round(Anoise, 2)), IB, , 1)
                        DB.AddProperties(Pname & "/" & IBname & " Aсум=" & CStr(Math.Round(Asum, 2)) & " Aшум=" & CStr(Math.Round(Anoise, 2)), AM_IB, 0)
                        DB.Delete(Pname & "/" & IBname)
                        AM = Nothing

                        ' приведение частоты амплитудного блока в роторную
                        For l = LBound(xval) To UBound(xval)
                            xval(l) = xval(l) / (Ttable(k, 4) / 60)
                        Next

                        IB = IB.ReplaceXDoubleValues(xval)
                        DB.AddItem(Pname & "/", "E_" & IBname & " Aсум=" & CStr(Math.Round(Asum, 2)) & " Aшум=" & CStr(Math.Round(Anoise, 2)), IB, , 1)
                        DB.AddProperties(Pname & "/E_" & IBname & " Aсум=" & CStr(Math.Round(Asum, 2)) & " Aшум=" & CStr(Math.Round(Anoise, 2)), AM_IB, 0)

                    ElseIf CSng(Rtable(k)) <= 2 Then  ' обработка запуска и останова

                        Meff = 0
                        u = 6
                        ESres.Cells(1, 6) = "Asum"  ' запись шапки

                        ' расчет эффективного значения суммарного сигнала
                        For l = Math.Round(Val(TextBox7.Text) / Fres) To Math.Round(Val(TextBox8.Text) / Fres)
                            Meff = Meff + yval(l)
                        Next

                        Meff = Math.Sqrt(Meff)
                        ESres.Cells(k + 1, 6) = Math.Round(Meff, 3)

                        'запись вероятностных характеристик
                        ESres.Cells(1, 7) = "MaxSp"
                        MaxSp = CSng(Val(AM_IB.Item("MaxSp")))
                        ESres.Cells(k + 1, 7) = MaxSp

                        ESres.Cells(1, 8) = "MinSp"
                        MinSp = CSng(Val(AM_IB.Item("MinSp")))
                        ESres.Cells(k + 1, 8) = MinSp

                        ESres.Cells(1, 9) = "Sx"
                        Sx = CSng(Val(AM_IB.Item("Sx")))
                        ESres.Cells(k + 1, 9) = Sx

                        ESres.Cells(1, 10) = "Sk"
                        Sk = CSng(Val(AM_IB.Item("Sk")))
                        ESres.Cells(k + 1, 10) = Sk
                        If Math.Abs(Sk) >= Val(TextBox12.Text) And CheckBox8.Checked = True Then ESres.Cells(k + 1, 10).Interior.Color = 200

                        ESres.Cells(1, 11) = "Excess"
                        Excess = CSng(Val(AM_IB.Item("Excess")))
                        ESres.Cells(k + 1, 11) = Excess
                        If Excess >= Val(TextBox13.Text) And CheckBox9.Checked = True Then ESres.Cells(k + 1, 11).Interior.Color = 200

                        ESres.Cells(1, 12) = "dT"
                        dT = CSng(Val(AM_IB.Item("dT")))
                        ESres.Cells(k + 1, 12) = dT

                        ESres.Cells(1, 13) = "km"
                        km = CSng(Val(AM_IB.Item("km")))
                        ESres.Cells(k + 1, 13) = km

                        u = 13

                        ' приведение блока СПМ в амплитудный
                        AM = TL.Factory.CreateAttributeMap
                        AM.Add("SourceBlock", IB)
                        IB = TL.Factory.CreateObject("LmsHq::DataModelC::MathProcessing::CSquareRootBlock", AM)

                        DB.AddItem(Pname & "/", IBname & " Aсум=" & CStr(Math.Round(Meff, 2)), IB, , 1)
                        DB.Delete(Pname & "/" & IBname)
                        AM = Nothing

                    End If
m1:
                Next k

                TextBox3.Text = TextBox3.Text & vbCrLf
                Label14.Text = CInt(Label14.Text) + 1
                Exit For
            End If
        Next

        If DB.SectionNames.Item(i) <> Pname Then
            ESres.Delete()
            TextBox3.Text = TextBox3.Text & vbTab & "absent" & vbCrLf
        End If
    End Sub

    ' определение числа линий для дисперсии на роторных оборотах
    Function CountLines(ByVal r As Integer, ByVal RPM As Single)
        r = Math.Ceiling(0.01 * r + 1)
        If Strings.Right(CStr(r / 2), 2) = ",5" Or Strings.Right(CStr(r / 2), 2) = ".5" Then Return r
        Return r - 1
    End Function

    Function Damage(ByVal yval() As Double, ByVal VAC As String, ByVal RPS As Single, ByVal Fres As Single, ByVal dT As Single)
        Dim Fr_l, Emax, Fr_h, a, f, d1, d2, d3, z, ne, tz, m, nm, Ik As Single
        Dim p, r As Integer

        If ComboBox6.Text = "РД171М" Then

            Select Case VAC
                Case "A5"
                    Fr_l = 30 : Fr_h = 500

                Case "A10"
                    Fr_l = 500 : Fr_h = 10000

                Case "A14"
                    Fr_l = RPS * 14 - 100 : Fr_h = RPS * 14 + 100

                Case "A33"
                    Fr_l = RPS * 33 - 100 : Fr_h = RPS * 33 + 100

                Case Else
                    MsgBox(VAC & " was not determined")
                    End
            End Select

        End If

        Emax = 0
        For p = Fr_l / Fres To Fr_h / Fres
            If Emax < yval(p) Then
                Emax = yval(p)      ' максимальная амплитуда
                AFreq = p * Fres    ' глобальная переменная - видна везде
            End If
        Next

        ' расчет интеграла повреждаемости
        d1 = 0
        d2 = 0
        d3 = 0

        ' для широкой полосы
        For p = Fr_l / Fres To Fr_h / Fres

            f = p * Fres
            a = Math.Sqrt(yval(p)) / f

            d1 += a * a
            d2 += a * a * (f ^ 2)
            d3 += a * a * (f ^ 4)

        Next

        ' для узкой полосы
        If VAC = "A14" Or VAC = "A33" Then

            a = Math.Sqrt(Emax) / AFreq

            d1 = a * a
            d2 = a * a * (AFreq ^ 2)
            d3 = a * a * (AFreq ^ 4)

        End If

        z = Math.Sqrt(d1 * d3) / d2
        tz = (0.234 * z + 0.761) / (0.207 * z ^ 2 + 1.754 * z - 1)
        m = 4
        nm = 10 ^ (0.017 * m ^ 2 + 0.216 * m - 0.237)
        ne = 1 / 6.28 * Math.Sqrt(d3 / d2) * dT
        Ik = (1 / 2) * (d1 ^ (m / 2)) * ne * tz * nm

        Return Ik

    End Function

    Function Acust(ByVal yval() As Double, ByVal Power As Single, ByVal RPS As Single, ByVal VAC As String, ByVal Fres As Single, ByVal Regime As Single)
        Dim Fr_l, Fr_h, Emax, Meff, n1, n2 As Single
        Dim r, p As Integer
        Dim LnsCnt As Integer = CInt((TextBox10.Text - 2 * Fres) / (2 * Fres))

        Regime = Power / 100    ' чтобы быть точнее

        If ComboBox6.Text = "РД180" Then

            Select Case VAC
                Case "A1050"     ' 1 тангенциальная ГГ
                    n1 = -1.6897 * Regime + 5.7803 - 0.5
                    n2 = -1.6897 * Regime + 5.7803 + 0.5

                    Fr_l = n1 * RPS : Fr_h = n2 * RPS
                    'Fr_l = 900 : Fr_h = 1200

                Case "A1650"    ' 2 тангенциальная ГГ
                    n1 = -2.979 * Regime + 8.3396 - 0.5
                    n2 = -2.979 * Regime + 8.3396 + 0.5

                    Fr_l = n1 * RPS : Fr_h = n2 * RPS
                    'Fr_l = 1500 : Fr_h = 1800

                Case "A1850"    ' частота 1 тангенциальной по КС
                    Fr_l = 1700 : Fr_h = 2000

                Case "A2200"    ' собственная ГГ
                    Fr_l = 2140 : Fr_h = 2400

                Case "A3000"    ' собственная ГГ
                    Fr_l = 2800 : Fr_h = 3100

                Case "A100"    ' Контурная по гор до ГГ
                    Fr_l = 85 : Fr_h = 230

                Case Else
                    MsgBox(VAC & " was not determined")
                    End
            End Select

        ElseIf ComboBox6.Text = "РД171М" Then

            ' коэффициент перевода частот от РД180 к РД171М
            Dim k As Single = (135.21 * Regime + 3469) / (105.48 * Regime + 2675.7)

            Select Case VAC
                Case "A85"      ' оценка контурных колебаний
                    Fr_l = 70 : Fr_h = 100

                Case "A450"      ' что то новое
                    Fr_l = 400 : Fr_h = 500

                Case "A1050"     ' 1 тангенциальная ГГ
                    n1 = -1.6897 * Regime + 5.7803 - 0.5
                    n2 = -1.6897 * Regime + 5.7803 + 0.5

                    Fr_l = n1 * RPS * k : Fr_h = n2 * RPS * k

                Case "A1650"    ' 2 тангенциальная ГГ
                    n1 = -2.979 * Regime + 8.3396 - 0.5
                    n2 = -2.979 * Regime + 8.3396 + 0.5

                    Fr_l = n1 * RPS * k : Fr_h = n2 * RPS * k

                Case "A1850"    ' частота 1 тангенциальной по КС
                    Fr_l = 1700 : Fr_h = 2000

                Case "A2200"    ' собственная ГГ
                    Fr_l = 2140 : Fr_h = 2400

                Case "A1550"    ' РР, собственная привода на рессорах
                    Fr_l = 1450 : Fr_h = 1650

                Case Else
                    MsgBox(VAC & " was not determined")
                    End
            End Select

        ElseIf ComboBox6.Text = "РД191" Or ComboBox6.Text = "РД181" Then

            Select Case VAC
                Case "A600"     ' продольная
                    n1 = -0.742 * Regime + 2.3823
                    n2 = -1.148 * Regime + 3.7168

                    Fr_l = n1 * RPS : Fr_h = n2 * RPS

                Case "A1050"    ' 1 тангенциальная
                    n1 = -1.7935 * Regime + 4.6187
                    n2 = -1.8575 * Regime + 5.6486

                    Fr_l = n1 * RPS : Fr_h = n2 * RPS

                Case "A1650"    ' 1 смешанная 
                    n1 = -1.8575 * Regime + 5.6486 + 0.3
                    n2 = -3.2125 * Regime + 8.0057

                    Fr_l = n1 * RPS : Fr_h = n2 * RPS

                Case "A2300"    ' 2 тангенциальная
                    n1 = -3.8579 * Regime + 9.4841
                    n2 = -5.272 * Regime + 11.809

                    Fr_l = n1 * RPS : Fr_h = n2 * RPS

                Case "A2350"    ' акустическая в трубе НГ2ст - регулятор (три полуволны), только для параметра ПГПН-2
                    Fr_l = 2200 : Fr_h = 2500

                Case "A1850"    ' частота 1 тангенциальной по КС
                    Fr_l = 1700 : Fr_h = 2000

                Case "A2750"    ' частота ГГ
                    Fr_l = 2600 : Fr_h = 2900

                Case "A265" ' собственная магистрали между НГ2ст и РГ, крутильная вала ТНА 210Гц
                    LnsCnt = 6
                    n1 = 0.2557 * Regime + 0.3433
                    n2 = -0.018 * Regime + 0.7993
                    Fr_l = n1 * RPS : Fr_h = n2 * RPS

                Case "A100" ' 
                    LnsCnt = 4
                    Fr_l = 85 - 25 : Fr_h = 85 + 25

                Case "A440"
                    LnsCnt = 6
                    Fr_l = 340 : Fr_h = 540

                Case "A480" 'только в параметрах ГГ, что - неизвестно
                    LnsCnt = 6
                    Fr_l = 420 : Fr_h = 520

                Case "A720"
                    LnsCnt = 6
                    Fr_l = 600 : Fr_h = 800

                Case "A650"
                    LnsCnt = 6
                    n1 = 0.1037 * Regime + 1.4777 - 0.25
                    n2 = 0.1037 * Regime + 1.4777 + 0.25
                    Fr_l = n1 * RPS : Fr_h = n2 * RPS

                Case Else
                    MsgBox(VAC & " was not determined")
                    End
            End Select

        ElseIf ComboBox6.Text = "XXL" Then

            Select Case VAC
                Case "A300"     ' ГГ
                    Fr_l = 250 : Fr_h = 315

                Case "A700"    ' ГГ
                    Fr_l = 630 : Fr_h = 800

                Case "A2000"    ' КС
                    Fr_l = 1600 : Fr_h = 2500

                Case "A3500"    ' КС
                    Fr_l = 3200 : Fr_h = 4000

                Case Else
                    MsgBox(VAC & " was not determined")
                    End
            End Select
        End If

        ' расчет акустики
        Emax = 0
        For p = Fr_l / Fres To Fr_h / Fres
            If Emax < yval(p) Then
                Emax = yval(p)      ' максимальная амплитуда
                r = p               ' номер гармоники
                AFreq = p * Fres    ' глобальная переменная - видна везде
            End If
        Next

        ' расчет дисперсии полюса
        If r - (Fr_l / Fres) < LnsCnt Then
            For p = Fr_l / Fres To (Fr_l / Fres) + LnsCnt * 2
                Meff = Meff + yval(p)
            Next
        ElseIf (Fr_h / Fres) - r < LnsCnt Then
            For p = Fr_h / Fres - LnsCnt * 2 To Fr_h / Fres
                Meff = Meff + yval(p)
            Next
        Else
            For p = r - LnsCnt To r + LnsCnt
                Meff = Meff + yval(p)
            Next
        End If

        Meff = Math.Sqrt(Meff)
        Return Meff

    End Function

    Private Sub TabPage1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage1.Enter
        PgN = 1
        Names()
    End Sub

    Private Sub TabPage2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Enter
        PgN = 2
        Names()
    End Sub

    Private Sub TabPage3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Enter
        PgN = 3
        Names()
    End Sub

    Private Sub TabPage4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage4.Enter
        PgN = 4
        Names()
    End Sub

    Private Sub TabPage5_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.Enter
        PgN = 5
        Names()
    End Sub

    Function Names()
        Dim DiskN As String = ComboBox7.Text & "\"
        Dim TestN As String
        Dim EngN As String
        Dim BlNum As Integer

        If RadioButton4.Checked = True Then
            TestN = "\NI"
        ElseIf RadioButton5.Checked = True Then
            TestN = "\SP"
        Else
            TestN = "\FL"
        End If

        Dim suffix As String = ""

        If CheckBox12.Checked = True Then
            suffix = "_2"
        End If

        If ComboBox6.Text = "РД180" Then
            EngN = "RD180"
            BlNum = 29
            TextBox7.Text = "30"
            TextBox8.Text = "8500"
            NumericUpDown6.Value = NumericUpDown4.Value
            PtVAC = DiskN & EngN & TestN & "_RD180VAC.xls"

            If RadioButton3.Checked = True Then
                BlNum = 7
                TextBox7.Text = "30"
                TextBox8.Text = "8500"
                'NumericUpDown3.Value = 25000
                ComboBox1.Text = 2048
            End If

        ElseIf ComboBox6.Text = "XXL" Then
            EngN = "XXL"
            BlNum = 32
            TextBox7.Text = "30"
            TextBox8.Text = "8000"
            NumericUpDown6.Value = NumericUpDown4.Value
            PtVAC = DiskN & EngN & TestN & "_XXLVAC.xls"
            NumericUpDown3.Value = 40960
            ComboBox1.Text = 8192

            If RadioButton3.Checked = True Then
                BlNum = 32
                TextBox7.Text = "30"
                TextBox8.Text = "8000"
                NumericUpDown3.Value = 16000
                ComboBox1.Text = 3200
                Label19.Text = CStr(Math.Round(NumericUpDown3.Value / Val(ComboBox1.Text), 3))
            End If

        ElseIf ComboBox6.Text = "РД171М" Then
            EngN = "RD171M"
            BlNum = 33
            TextBox7.Text = "30"
            TextBox8.Text = "8500"
            NumericUpDown6.Value = NumericUpDown4.Value
            PtVAC = DiskN & EngN & TestN & "_RD171MVAC.xls"

            If RadioButton3.Checked = True Then
                BlNum = 7
                TextBox7.Text = "30"
                TextBox8.Text = "2000"
                'NumericUpDown3.Value = 8000
                ComboBox1.Text = 2048
            End If

        ElseIf ComboBox6.Text = "РД191" Then
            EngN = "RD191"
            BlNum = 31
            TextBox7.Text = "30"
            TextBox8.Text = "14000"

            If TestN = "\FL" Then
                TextBox7.Text = "200"
                TextBox8.Text = "6000"
            End If

            NumericUpDown6.Value = NumericUpDown4.Value
            PtVAC = DiskN & EngN & TestN & "_RD191VAC" & suffix & ".xls"

        ElseIf ComboBox6.Text = "РД181" Then
            EngN = "RD181"
            BlNum = 31
            TextBox7.Text = "30"
            TextBox8.Text = "14000"

            If TestN = "\FL" Then
                TextBox7.Text = "30"
                TextBox8.Text = "6000"
            End If

            NumericUpDown6.Value = NumericUpDown4.Value
            PtVAC = DiskN & EngN & TestN & "_RD181VAC.xls"
        End If

        Label2.Visible = True
        TextBox2.Visible = True
        TextBox2.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & "pas" & suffix & ".xls"
        TextBox9.Text = BlNum
        Label31.Text = "N/A"

        If PgN = 1 Then
            Label1.Text = "MERA file:"
            TextBox1.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & ".MERA"
            TextBox4.Text = DiskN & EngN & TestN & "_" & EngN & "Pro.xls"
            TextBox5.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & suffix & ".lms"
        ElseIf PgN = 2 Then
            Label1.Text = "LMS file:"
            TextBox1.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & suffix & ".lms"
            TextBox4.Text = DiskN & EngN & TestN & "_" & EngN & "Pro.xls"
            TextBox5.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & suffix & "res.xls"
        ElseIf PgN = 3 Then
            Label1.Text = "VAC file:"
            TextBox1.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & suffix & "res.xls"
            TextBox4.Text = DiskN & EngN & TestN & "_" & EngN & "Pro.xls"
            TextBox5.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & "pro.xls"
            Label2.Visible = False
            TextBox2.Visible = False
        ElseIf PgN = 4 Then
            Label1.Text = "Protocol file:"
            TextBox1.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & "pro.xls"
            TextBox4.Text = DiskN & EngN & TestN & "_" & EngN & "Pro.xls"
            TextBox5.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & "exc.lms"
            TextBox6.Text = DiskN & EngN & TestN
            Label2.Visible = False
            TextBox2.Visible = False
        ElseIf PgN = 5 Then
            Label1.Text = "Protocol file:"
            TextBox1.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & "pro.xls"
            TextBox4.Text = DiskN & EngN & TestN & "_" & EngN & "Pro.xls"
            TextBox5.Text = DiskN & EngN & TestN & NumericUpDown4.Value & TestN & NumericUpDown4.Value & "trn_exc.lms"
            TextBox6.Text = DiskN & EngN & TestN
            Label2.Visible = False
            TextBox2.Visible = False
        End If

        Label7.Text = "0"
        Label7.BackColor = Color.Empty
        Label13.Text = "0"
        Label14.Text = "0"
        Label22.Text = "0"
        Label25.Text = "0"
        TextBox3.Text = ""
        ToolStripStatusLabel1.Text = "Ready"

        Return 0
    End Function

    Private Sub NumericUpDown4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown4.ValueChanged
        Names()
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        Names()
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        Names()
    End Sub

    ' заполнение файла протокола
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Names()

        Dim EAres As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EApro As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EWres, EWpro As Microsoft.Office.Interop.Excel.Workbook
        Dim ESpro As Microsoft.Office.Interop.Excel.Worksheet

        Dim PWtbl(100) As String
        Dim DPtbl(20), Pname, PnamePrev As String

        Dim k, l, m, PwR, DPn, f, ExcKey, cSign As Integer
        Dim Exc_row As Integer = 0
        Dim Exc_clm As Integer = 1

        Dim PtRES As String = TextBox1.Text
        Dim PtPRO As String = TextBox4.Text    ' назначение файла протокола в зависимости от типа двигателя

        Label7.Text = "0"
        Label7.BackColor = Color.Empty
        Label22.Text = "0"
        TextBox3.Text = ""
        ToolStripStatusLabel1.Text = "Creating protocol" : Me.Refresh()

        ' открытие шаблона протокола
        Try
            EApro.Workbooks.Open(PtPRO)
        Catch ex As Exception
            MsgBox("File " & PtPRO & " does not exist.")
            Exit Sub
        End Try

        EWpro = EApro.ActiveWorkbook

        Try
            EApro.Worksheets("Protocol").Select()
        Catch ex As Exception
            MsgBox("Sheet Protocol does not exist.")
            Exit Sub
        End Try

        ' определение числа обрабатываемых параметров
        Do
            cSign += 1
        Loop While CStr(EApro.Cells(cSign + 3, 2).Value) <> ""

        Try
            EAres.Workbooks.Open(PtRES)
        Catch ex As Exception
            MsgBox("File " & PtRES & " does not exist.")
            Exit Sub
        End Try

        EWres = EApro.ActiveWorkbook

        If CheckBox10.Checked = True Then
            EAres.Visible = False
            EApro.Visible = False
        Else
            EAres.Visible = True
            EApro.Visible = True
        End If


        ' запись режимов 
        Do
            PWtbl(k + 1) = CStr(EApro.Cells(1, 3 + k * 2).Value)
            k += 1  ' число режимов
        Loop While CStr(EApro.Cells(1, 3 + k * 2).Value) <> ""

        l = 3
        m = 1
        Do
            Pname = CStr(EApro.Cells(l, 1).Value)   ' имя параметра
            TextBox3.Text = TextBox3.Text & Pname : Me.Refresh()

            Do
                DPtbl(m) = CStr(EApro.Cells(l, 2).Value)
                m += 1
                l += 1
            Loop While CStr(EApro.Cells(l, 1).Value) = "" And CStr(EApro.Cells(l, 2).Value) <> ""
            m -= 1    ' число ДП

            ' чтение и запись результатов
            For PwR = 1 To k            ' разбивка режимов
                ExcKey = 0              ' ключ для проверки наличия режима с превышением нормы
                Dim PwL, PwH As Single  ' верхняя и нижняя границы режима

                If PWtbl(PwR) = "0" Then
                    PwL = 0
                    PwH = 0.01   ' режим запуска
                ElseIf PWtbl(PwR) = "2" Then
                    PwL = 2
                    PwH = 2.01   ' останов
                ElseIf Strings.Left(PWtbl(PwR), 1) = "<" Then
                    PwL = 0.1
                    PwH = Mid(PWtbl(PwR), 2, 10) / 100
                ElseIf Strings.Left(PWtbl(PwR), 1) = ">" Then
                    PwL = Mid(PWtbl(PwR), 2, 10) / 100
                    PwH = 1.1
                Else
                    PwL = Math.Abs(CSng(Strings.Left(PWtbl(PwR), 4))) / 100
                    PwH = Math.Abs(CSng(Strings.Right(PWtbl(PwR), 4))) / 100
                End If

                Dim clmR, rowR As Integer
                Dim Ax As Single = 0

                For DPn = 1 To m        'проход по всем ДП

                    Try
                        EAres.Worksheets(Pname).Select()
                    Catch ex As Exception
                        'если результаты по параметру отсутствуют
                        Exit For
                    End Try

                    clmR = 0
                    Do                  'поИск в результатах соответсвующего ДП
                        clmR += 1
                        If CStr(EAres.Cells(1, clmR).Value) = DPtbl(DPn) Then ' нашел в результатах нужный ДП
                            Ax = 0
                            rowR = 1    'пропустим первую строчку

                            Dim RegDef As Integer = 0   'для подсчета дефектных интервалов при ЛИ
                            Dim RegCnt As Integer = 0

                            Do          'проход по режимам
                                rowR += 1

                                'проверка превышения вероятностных характеристик для стационарных режимов
                                'если Асимметрия превышена и стоит маркер фильтрации по её превышению
                                'И Эксцесс превышен и стоит маркер фильтрации по его превышению, то в протокол этот режим не записывать
                                Dim ExceedKey As Integer = 0
                                Dim Regime As Single = EAres.Cells(rowR, 3).Value / 100

                                If EAres.Cells(rowR, 2).Value = 2 Then Regime = 2

                                If RadioButton4.Checked = True Then 'КТИ
                                    If Math.Abs(EAres.Cells(rowR, 10).Value) >= Val(TextBox12.Text) And EAres.Cells(rowR, 11).Value >= Val(TextBox13.Text) And CheckBox11.Checked = True And CSng(EAres.Cells(rowR, 1).Value) > 6000 Then ExceedKey = 1
                                  ElseIf RadioButton3.Checked = True Then 'ЛИ
                                    If (Math.Abs(EAres.Cells(rowR, 10).Value) >= Val(TextBox12.Text) Or EAres.Cells(rowR, 11).Value >= Val(TextBox13.Text)) And CheckBox11.Checked = True And CSng(EAres.Cells(rowR, 1).Value) > 6000 Then ExceedKey = 1
                                    'бракуем запуск и останов
                                    If Math.Abs(EAres.Cells(rowR, 10).Value) >= Val(TextBox12.Text) And EAres.Cells(rowR, 11).Value >= Val(TextBox13.Text) And CheckBox11.Checked = True And CSng(EAres.Cells(rowR, 1).Value) < 6000 Then ExceedKey = 1
                                End If






                                'If CSng(EAres.Cells(rowR, 1).Value) > 6000 Then RegCnt += 1
                                RegCnt += 1
                                'подсчет дефектных интервалов при ЛИ
                                'If RadioButton3.Checked = True And DPtbl(DPn) = "Asum" And ExceedKey = 1 Then RegDef = RegDef + 1




                                'проверка на вхождение режима в интервал и поиск максимума
                                If Regime >= PwL And Regime < PwH And Math.Abs(Ax) < Math.Abs(CSng(EAres.Cells(rowR, clmR).Value)) And ExceedKey = 0 Then
                                    Ax = Math.Round(CSng(EAres.Cells(rowR, clmR).Value), 2)
                                    EApro.Cells(l - m + (DPn - 1), 1 + (PwR * 2)) = Ax

                                    ToolStripProgressBar1.Value = 100 * (l - 3) / cSign

                                    'проверка превышения норм
                                    If RadioButton1.Checked = False And Math.Abs(Math.Round(EApro.Cells(l - m + (DPn - 1), 1 + (PwR * 2) + 1).Value, 2)) < Math.Abs(Math.Round(Ax, 2)) Then

                                        EApro.Cells(l - m + (DPn - 1), 1 + (PwR * 2)).Interior.Colorindex = 17
                                        Label7.BackColor = Color.Red
                                        Label7.Text = CStr(Val(Label7.Text) + 1) : Me.Refresh()

                                        Try
                                            EApro.Worksheets("Excess").Select()
                                        Catch ex As Exception
                                            ESpro = EWres.Worksheets.Add
                                            ESpro.Name = "Excess"
                                            EApro.Worksheets("Excess").Select()
                                        End Try

                                        If PnamePrev = Pname Then
                                            For f = 2 To Exc_clm    ' проверка наличия режима в списке превышения
                                                If EApro.Cells(Exc_row, f).Value = CInt(EAres.Cells(rowR, 1).Value) Then
                                                    ExcKey = 1
                                                End If
                                            Next
                                            If ExcKey = 0 Then
                                                Exc_clm += 1
                                                EApro.Cells(Exc_row, Exc_clm) = CInt(EAres.Cells(rowR, 1).Value)
                                            End If
                                        ElseIf PnamePrev <> Pname Then
                                            PnamePrev = Pname
                                            Exc_row += 1
                                            Exc_clm = 2
                                            EApro.Cells(Exc_row, 1) = Pname
                                            EApro.Cells(Exc_row, Exc_clm) = CInt(EAres.Cells(rowR, 1).Value)
                                        End If
                                        EApro.Worksheets("Protocol").Select()
                                    End If
                                End If
                            Loop While CStr(EAres.Cells(rowR + 1, 1).Value) <> ""




                            'расчет продолжительности дефектных интервалов
                            'If RadioButton3.Checked = True And DPtbl(DPn) = "Asum" Then
                            'EApro.Cells(l - m + (DPn - 1), 14) = Math.Round(RegDef / (RegCnt / 100))
                            ' не считаем запуск и останов
                            'If Math.Round(RegDef / (RegCnt / 100)) > 100 Then
                            'EApro.Cells(l - m + (DPn - 1), 14) = 100
                            'End If
                            'End If




                        End If
                    Loop While clmR < 55 'CStr(EAres.Cells(1, clmR + 1).Value) <> ""
                Next
            Next

            m = 1
            Label22.Text = CStr(Val(Label22.Text) + 1)
            TextBox3.Text = TextBox3.Text & vbTab & "processed" & vbCrLf : Me.Refresh()
        Loop While CStr(EApro.Cells(l, 2).Value) <> ""


        ' закраска протокола
        Dim CellFrm As String
        Dim CellColr As Single

        m = 25
        Do
            l = 3
            Do
                Try
                    CellFrm = EApro.Cells(l, m).FormulaLocal
                    CellColr = EApro.Range(CellFrm).Interior.ColorIndex.ToString
                    EApro.Cells(l, m).Interior.Colorindex = CellColr
                Catch ex7 As Exception
                End Try
                l += 1
            Loop While CStr(EApro.Cells(l, m).Value) <> ""
            m += 1
        Loop While CStr(EApro.Cells(l - 1, m).Value) <> ""


        EAres.Quit()
        If CheckBox1.Checked = True Then
            Try
                System.IO.File.Delete(TextBox5.Text)
            Catch ex As Exception
            End Try
            EWpro.SaveAs(TextBox5.Text)
        End If

        ToolStripStatusLabel1.Text = "Well done!"
        EApro.Quit()
        ToolStripProgressBar1.Value = 0 : Me.Refresh()
    End Sub

    Function TakeRegime(ByVal CodeR As Integer)

        If ComboBox6.Text = "РД180" Then
            If CodeR >= 6081 And CodeR <= 6085 Then Return 0.4
            If CodeR = 6159 Then Return 0.4
            If CodeR = 6078 Then Return 0.47
            If CodeR = 6144 Then Return 0.47
            If CodeR = 6145 Then Return 0.47
            If CodeR = 6160 Then Return 0.47
            If CodeR = 6212 Then Return 0.47
            If CodeR = 6229 Then Return 0.47
            If CodeR = 6239 Then Return 0.47
            If CodeR = 6352 Then Return 0.47
            If CodeR >= 6170 And CodeR <= 6175 Then Return 0.5
            If CodeR >= 6086 And CodeR <= 6090 Then Return 0.55
            If CodeR = 6158 Then Return 0.55
            If CodeR >= 6177 And CodeR <= 6179 Then Return 0.61
            If CodeR = 6142 Then Return 0.61
            If CodeR = 6143 Then Return 0.61
            If CodeR = 6091 Then Return 0.61
            If CodeR = 6216 Then Return 0.61
            If CodeR = 6162 Then Return 0.63
            If CodeR = 6186 Then Return 0.63
            If CodeR = 6235 Then Return 0.63
            If CodeR = 6246 Then Return 0.63
            If CodeR = 6251 Then Return 0.63
            If CodeR = 6256 Then Return 0.63
            If CodeR = 6092 Then Return 0.64
            If CodeR >= 6200 And CodeR <= 6203 Then Return 0.65
            If CodeR = 6217 Then Return 0.65
            If CodeR = 6231 Then Return 0.65
            If CodeR = 6241 Then Return 0.65
            If CodeR >= 6093 And CodeR <= 6097 Then Return 0.7
            If CodeR = 6157 Then Return 0.7
            If CodeR >= 6195 And CodeR <= 6198 Then Return 0.74
            If CodeR = 6219 Then Return 0.74
            If CodeR = 6232 Then Return 0.74
            If CodeR = 6242 Then Return 0.74
            If CodeR >= 6180 And CodeR <= 6185 Then Return 0.8
            If CodeR >= 6098 And CodeR <= 6102 Then Return 0.84
            If CodeR = 6156 Then Return 0.84
            If CodeR = 6221 Then Return 0.84
            If CodeR = 6163 Then Return 0.87
            If CodeR = 6187 Then Return 0.87
            If CodeR = 6236 Then Return 0.87
            If CodeR = 6247 Then Return 0.87
            If CodeR = 6252 Then Return 0.87
            If CodeR = 6257 Then Return 0.87
            If CodeR = 6103 Then Return 0.89
            If CodeR = 6138 Then Return 0.89
            If CodeR = 6141 Then Return 0.89
            If CodeR = 6155 Then Return 0.89
            If CodeR = 6188 Then Return 0.89
            If CodeR = 6189 Then Return 0.89
            If CodeR = 6222 Then Return 0.89
            If CodeR = 6139 Then Return 0.95
            If CodeR = 6193 Then Return 0.95
            If CodeR >= 6104 And CodeR <= 6109 Then Return 1
            If CodeR = 6331 Then Return 1
            If CodeR = 6365 Then Return 1
        ElseIf ComboBox6.Text = "РД180В" Then

            Return 0
        ElseIf ComboBox6.Text = "РД171М" Then
            If CodeR >= 6141 And CodeR <= 6154 Then Return 0.4
            If CodeR >= 6158 And CodeR <= 6171 Then Return 0.45
            If CodeR >= 6176 And CodeR <= 6189 Then Return 0.5
            If CodeR >= 6193 And CodeR <= 6206 Then Return 0.55
            If CodeR >= 6211 And CodeR <= 6224 Then Return 0.6
            If CodeR >= 6228 And CodeR <= 6241 Then Return 0.65
            If CodeR >= 6246 And CodeR <= 6259 Then Return 0.7
            If CodeR >= 6263 And CodeR <= 6276 Then Return 0.75
            If CodeR >= 6281 And CodeR <= 6294 Then Return 0.8
            If CodeR >= 6298 And CodeR <= 6311 Then Return 0.85
            If CodeR >= 6316 And CodeR <= 6329 Then Return 0.9
            If CodeR >= 6333 And CodeR <= 6346 Then Return 0.95
            If CodeR >= 6351 And CodeR <= 6365 Then Return 1.0
            If CodeR >= 6368 And CodeR <= 6381 Then Return 1.05

        ElseIf ComboBox6.Text = "РД181" And RadioButton3.Checked = True Then
            If CodeR >= 6100 And CodeR <= 6199 Then Return 0.55
            If CodeR >= 6000 And CodeR <= 6099 Then Return 1.0

        Else    ' если диагностируем РД191

            If CodeR >= 6438 And CodeR <= 6451 Then Return 0.27
            If CodeR >= 6141 And CodeR <= 6154 Then Return 0.3
            If CodeR >= 6158 And CodeR <= 6171 Then Return 0.35
            If CodeR >= 6176 And CodeR <= 6189 Then Return 0.38
            If CodeR >= 6193 And CodeR <= 6206 Then Return 0.45
            If CodeR >= 6211 And CodeR <= 6224 Then Return 0.5
            If CodeR >= 6421 And CodeR <= 6434 Then Return 0.55
            If CodeR >= 6227 And CodeR <= 6241 Then Return 0.6
            If CodeR >= 6403 And CodeR <= 6416 Then Return 0.65
            If CodeR >= 6246 And CodeR <= 6259 Then Return 0.7
            If CodeR >= 6263 And CodeR <= 6276 Then Return 0.75
            If CodeR >= 6281 And CodeR <= 6294 Then Return 0.8
            If CodeR >= 6298 And CodeR <= 6311 Then Return 0.85
            If CodeR >= 6316 And CodeR <= 6329 Then Return 0.9
            If CodeR >= 6333 And CodeR <= 6346 Then Return 0.95
            If CodeR >= 6351 And CodeR <= 6364 Then Return 1
            If CodeR >= 6500 And CodeR <= 6600 Then Return 1
            If CodeR >= 6368 And CodeR <= 6381 Then Return 1.05
        End If

        Return CodeR    ' если запуск или останов

    End Function

    ' сбор статистики
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Names()

        Dim EApro As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EWpro As Microsoft.Office.Interop.Excel.Workbook

        Dim Pname(100) As String    ' имена параметров
        Dim RegCnt(100), RegNm(100, 100) As Integer ' кол. режимов по каждому параметру и режимы
        Dim i, k As Integer
        Dim PtPRO As String = TextBox1.Text    ' назначение файла протокола в зависимости от типа двигателя

        Label25.Text = "0"
        ToolStripStatusLabel1.Text = "Excesses processing" : Me.Refresh()

        ' открытие протокола
        Try
            EApro.Workbooks.Open(PtPRO)
        Catch ex As Exception
            MsgBox("File " & PtPRO & " does not exist.")
            Exit Sub
        End Try

        EWpro = EApro.ActiveWorkbook
        If CheckBox10.Checked = True Then
            EApro.Visible = False
        Else
            EApro.Visible = True
        End If

        Try
            EApro.Worksheets("Excess").Select()
        Catch ex As Exception
            MsgBox("No Excesses have been found.")
            Exit Sub
        End Try

        Do  ' заполнение массивов именами параметров и режимами
            i += 1  ' количество параметров
            Pname(i) = EApro.Cells(i, 1).Value
            k = 1
            Do
                k += 1
                RegNm(i, k - 1) = EApro.Cells(i, k).Value
            Loop While CStr(EApro.Cells(i, k + 1).Value) <> ""
            RegCnt(i) = k - 1
        Loop While CStr(EApro.Cells(i + 1, 1).Value) <> ""

        Call ExcessProcessing(i, Pname, RegCnt, RegNm)

        EApro.Quit()
        If CheckBox1.Checked = True Then
            Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
            Dim DB As LMSTestLabAutomation.IDatabase
            DB = TL.ActiveBook.Database
            ToolStripStatusLabel1.Text = "Saving LMS file" : Me.Refresh()
            TL.ActiveBook.Save(TextBox5.Text)
        End If

        ToolStripStatusLabel1.Text = "Well done!"
        ToolStripProgressBar1.Value = 0 : Me.Refresh()
        Beep()

    End Sub

    Private Sub ExcessProcessing(ByVal i As Integer, ByVal Pname() As String, ByVal RegCnt() As Integer, ByVal RegNm(,) As Integer)
        Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
        Dim DB As LMSTestLabAutomation.IDatabase

        If TL.Name <> "" Then
            TL.ActiveBook.Close()
            'TL.OpenProject("C:\LMSLocal10B\tka4\Data\Project1.lms")
            TL.NewProject(ComboBox7.Text & "\Project1.lms")
        ElseIf TL.Name = "" Then
            ToolStripStatusLabel1.Text = "Opening TestLab..." : Me.Refresh()
            TL.Init("-w DesktopStandard")
        End If
        DB = TL.ActiveBook.Database

        Dim DW As LMSTestLabAutomation.DataWatch = TL.ActiveBook.FindDataWatch("Navigator_Explorer")
        Dim Ex As LMSTestLabAutomation.IExplorer = DW.Data
        Dim Br As LMSTestLabAutomation.IDataBrowser

        Dim k, l, m, z As Integer
        Dim labs(), ids() As String
        Dim ID, ID_E, ID_Sig As LMSTestLabAutomation.IData

        For k = 1 To i  ' цикл по параметрам
            ToolStripStatusLabel1.Text = Pname(k) : Me.Refresh()

            For l = NumericUpDown5.Value To NumericUpDown6.Value   ' цикл по испытаниям
                Dim PthLMS As String = TextBox6.Text & l & "\" & Strings.Right(TextBox6.Text, 2) & l & ".lms"

                Try
                    Br = Ex.Browser(PthLMS) ' открываем испытание
                    Br.Elements(Pname(k) & "/", labs, ids)

                    For m = 1 To RegCnt(k)  ' проход по режимам
                        For z = 0 To UBound(ids) '- 1

                            Dim kk As String = Strings.Left(labs(z), 1)
                            Dim kk2 As String = CStr(RegNm(k, m))

                            If Strings.Left(labs(z), 4) = CStr(RegNm(k, m)) Or Strings.Left(labs(z), 1) = CStr(RegNm(k, m)) Then

                                ID = Br.GetItem(Pname(k) & "/" & labs(z))   ' чтение блока
                                If RegNm(k, m) > 6000 Then ID_E = Br.GetItem(Pname(k) & "/E_" & labs(z)) ' чтение спектра роторных
                                'Dim kjhg As Integer = (z - (UBound(ids) - 4) / 3 + 2)) / 2
                                'If RegNm(k, m) > 6000 Then ID_Sig = Br.GetItem(Pname(k) & "/" & labs((z - (UBound(ids) - 4) / 3 + 2) / 2)) ' чтение сигнала 
                                'если режим есть

                                Try
                                    If RegNm(k, m) < 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", l & Mid(labs(z), 2, 100), ID, , 1) ' запись спектра
                                    If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", l & Mid(labs(z), 5, 100), ID, , 1) ' добавить в базу спектр режима
                                    If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "E_" & l & Mid(labs(z), 5, 100), ID_E, , 1)
                                    'If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "Sig " & l, ID_Sig, , 1)
                                Catch ex3 As Exception
                                    DB.AddSection(Pname(k) & "_" & RegNm(k, m))                 ' создать сначала секцию в базе
                                    If RegNm(k, m) < 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", l & Mid(labs(z), 2, 100), ID, , 1)
                                    If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", l & Mid(labs(z), 5, 100), ID, , 1) ' добавить в базу спектр режима
                                    If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "E_" & l & Mid(labs(z), 5, 100), ID_E, , 1)
                                    'If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "Sig " & l, ID_Sig, , 1)
                                End Try

                            ElseIf Strings.Left(labs(z), 5) = "Sig " & CStr(RegNm(k, m)) Then ' чтение сигнала по запуску или останову
                                ID_Sig = Br.GetItem(Pname(k) & "/" & labs(z))
                                Try
                                    DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "Sig " & l & Strings.Mid(labs(z), 6, 100), ID_Sig, , 1)
                                Catch ex4 As Exception
                                    DB.AddSection(Pname(k) & "_" & RegNm(k, m))
                                    DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "Sig " & l & Strings.Mid(labs(z), 6, 100), ID_Sig, , 1)
                                End Try


                            End If









                            If Strings.Left(labs(z), 8) = "Sig " & CStr(RegNm(k, m)) Then

                                If RegNm(k, m) > 6000 Then ID_Sig = Br.GetItem(Pname(k) & "/" & labs(z)) ' чтение сигнала 
                                'если режим есть

                                Try
                                    If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "Sig " & l, ID_Sig, , 1)
                                Catch ex3 As Exception
                                    DB.AddSection(Pname(k) & "_" & RegNm(k, m))                 ' создать сначала секцию в базе
                                    If RegNm(k, m) > 6000 Then DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", "Sig " & l, ID_Sig, , 1)
                                End Try

                            End If









                        Next    ' следующий параметр из lms
                    Next

                    ' если нет параметра - берем следующее испытание
                Catch ex2 As Exception
                    ' если нет испытания - берем следующее
                End Try
            Next

            Label25.Text = k    ' счетчик обработанных параметров
            ToolStripProgressBar1.Value = 100 * k / i : Me.Refresh()
        Next

    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged
        Label19.Text = CStr(Math.Round(NumericUpDown3.Value / Val(ComboBox1.Text), 3))
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label19.Text = CStr(Math.Round(NumericUpDown3.Value / Val(ComboBox1.Text), 3))
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Me.Width = 484
        Else
            Me.Width = 283
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            NumericUpDown5.Enabled = True
            NumericUpDown6.Enabled = True
        Else
            NumericUpDown5.Enabled = False
            NumericUpDown6.Enabled = False
        End If
    End Sub

    ' goall - button
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        CheckBox4.Checked = False

        Dim i As Integer

        If ComboBox6.Text = "РД191" Then

            For i = 217 To 226 Step 1

                If i = 20 Or i = 21 Or i = 26 Then GoTo m2 ' нет паспорта
                If i = 26 Then GoTo m2 ' отсутствует
                If i = 30 Or i = 31 Then GoTo m2 ' отсутствует
                If i = 36 Then GoTo m2 ' отсутствует
                If i = 42 Or i = 127 Then GoTo m2 ' отсутствует
                'If i > 179 And i < 1001 Then GoTo m2 ' отсутствует
                If i = 113 Then GoTo m2 ' отсутствует
                'If i = 1022 Then GoTo m2 ' уже есть


                If i > 170 And i < 200 Then GoTo m2

                NumericUpDown4.Value = i
                Names()

                TabControl1.SelectedTab = TabPage1  ' Calculate spectrums
                TabPage1_Enter(sender, e)
                Button1_Click(sender, e)

                TabControl1.SelectedTab = TabPage2  ' Calculate VACs
                TabPage2_Enter(sender, e)
                Button2_Click(sender, e)

                'TabControl1.SelectedTab = TabPage3  ' Write protocol
                'TabPage3_Enter(sender, e)
                'Button3_Click(sender, e)
m2:
                Me.Refresh()
            Next

        ElseIf ComboBox6.Text = "РД180" Then

            For i = 244 To 235 Step -1
                'If i <> 275 And i <> 241 And i <> 144 And i <> 126 And i <> 140 And i <> 148 And i <> 159 Then GoTo m3

                NumericUpDown4.Value = i
                Names()

                TabControl1.SelectedTab = TabPage1  ' Calculate spectrums
                TabPage1_Enter(sender, e)
                Button1_Click(sender, e)

                TabControl1.SelectedTab = TabPage2  ' Calculate VACs
                TabPage2_Enter(sender, e)
                Button2_Click(sender, e)

                'TabControl1.SelectedTab = TabPage3  ' Write protocol
                'TabPage3_Enter(sender, e)
                'Button3_Click(sender, e)
m3:
                Me.Refresh()
            Next
        ElseIf ComboBox6.Text = "РД181" Then

            For i = 5 To 2 Step -1
                NumericUpDown4.Value = i
                Names()

                TabControl1.SelectedTab = TabPage1  ' Calculate spectrums
                TabPage1_Enter(sender, e)
                Button1_Click(sender, e)

                TabControl1.SelectedTab = TabPage2  ' Calculate VACs
                TabPage2_Enter(sender, e)
                Button2_Click(sender, e)

                TabControl1.SelectedTab = TabPage3  ' Write protocol
                TabPage3_Enter(sender, e)
                Button3_Click(sender, e)

                Me.Refresh()
            Next

        ElseIf ComboBox6.Text = "РД171М" Then

            For i = 2001 To 2016

                'If i > 999 And i < 2001 Then GoTo m4
                'If i >= 976 And i <= 978 Then GoTo m4
                'If i >= 980 And i <= 983 Then GoTo m4
                'If i >= 985 And i <= 989 Then GoTo m4

                NumericUpDown4.Value = i
                Names()

                TabControl1.SelectedTab = TabPage1  ' Calculate spectrums
                TabPage1_Enter(sender, e)
                Button1_Click(sender, e)

                TabControl1.SelectedTab = TabPage2  ' Calculate VACs
                TabPage2_Enter(sender, e)
                Button2_Click(sender, e)

                'TabControl1.SelectedTab = TabPage3  ' Write protocol
                'TabPage3_Enter(sender, e)
                'Button3_Click(sender, e)
m4:
                Me.Refresh()
            Next

        End If

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        Names()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        Names()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'сбор статистики по обработке переходных процессов
        Names()

        Dim EApro As Microsoft.Office.Interop.Excel.Application = CreateObject("Excel.Application")
        Dim EWpro As Microsoft.Office.Interop.Excel.Workbook

        Dim Pname(100) As String    ' имена параметров
        Dim RegCnt(100) As Integer ' кол. ДП по каждому параметру
        Dim RegNm(100, 100) As String ' имена ДП
        Dim i, k As Integer
        Dim PtPRO As String = TextBox1.Text    ' назначение файла протокола в зависимости от типа двигателя

        ToolStripStatusLabel1.Text = "Transient excesses processing" : Me.Refresh()

        ' открытие протокола
        Try
            EApro.Workbooks.Open(PtPRO)
        Catch ex As Exception
            MsgBox("File " & PtPRO & " does not exist.")
            Exit Sub
        End Try

        EWpro = EApro.ActiveWorkbook
        If CheckBox10.Checked = True Then
            EApro.Visible = False
        Else
            EApro.Visible = True
        End If

        Try
            EApro.Worksheets("Excess2").Select()
        Catch ex As Exception
            MsgBox("No Excesses have been found.")
            Exit Sub
        End Try

        Do  ' заполнение массивов именами параметров и списком ДП
            i += 1  ' количество параметров
            Pname(i) = EApro.Cells(i, 1).Value
            k = 1
            Do
                k += 1
                RegNm(i, k - 1) = EApro.Cells(i, k).Text
            Loop While CStr(EApro.Cells(i, k + 1).Value) <> ""
            RegCnt(i) = k - 1
        Loop While CStr(EApro.Cells(i + 1, 1).Value) <> ""

        Call ExcessProcessing2(i, Pname, RegCnt, RegNm)

        EApro.Quit()

        If CheckBox1.Checked = True Then
            Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
            Dim DB As LMSTestLabAutomation.IDatabase
            DB = TL.ActiveBook.Database
            ToolStripStatusLabel1.Text = "Saving LMS file" : Me.Refresh()
            TL.ActiveBook.Save(TextBox5.Text)
        End If

        ToolStripStatusLabel1.Text = "Well done!"
        ToolStripProgressBar1.Value = 0 : Me.Refresh()
        Beep()

    End Sub

    Private Sub ExcessProcessing2(ByVal i As Integer, ByVal Pname() As String, ByVal RegCnt() As Integer, ByVal RegNm(,) As String)
        Dim TL As LMSTestLabAutomation.Application = New LMSTestLabAutomation.Application
        Dim DB As LMSTestLabAutomation.IDatabase

        If TL.Name <> "" Then
            TL.ActiveBook.Close()
            'TL.OpenProject("C:\LMSLocal10B\tka4\Data\Project1.lms")
            TL.NewProject(ComboBox7.Text & "\Project1.lms")
        ElseIf TL.Name = "" Then
            ToolStripStatusLabel1.Text = "Opening TestLab..." : Me.Refresh()
            TL.Init("-w DesktopStandard")
        End If
        DB = TL.ActiveBook.Database

        Dim DW As LMSTestLabAutomation.DataWatch = TL.ActiveBook.FindDataWatch("Navigator_Explorer")
        Dim Ex As LMSTestLabAutomation.IExplorer = DW.Data
        Dim Br As LMSTestLabAutomation.IDataBrowser

        Dim k, l, m, z As Integer
        Dim labs(), ids() As String
        Dim ID, ID_E, ID_Sig As LMSTestLabAutomation.IData

        For k = 1 To i  ' цикл по параметрам
            ToolStripStatusLabel1.Text = Pname(k) : Me.Refresh()

            For l = NumericUpDown2.Value To NumericUpDown1.Value   ' цикл по испытаниям
                Dim PthLMS As String = TextBox6.Text & l & "\" & Strings.Right(TextBox6.Text, 2) & l & "trn.lms"





                Try
                    Br = Ex.Browser(PthLMS) ' открываем испытание
                    Br.Elements(Pname(k) & "/", labs, ids)

                    For m = 1 To RegCnt(k)  ' проход по ДП
                        For z = 0 To UBound(ids) '- 1

                            'Dim kk As String = labs(z) 'для контроля при дебагинге
                            'Dim kk2 As String = CStr(RegNm(k, m)) 'для контроля при дебагинге

                            If labs(z) = CStr(RegNm(k, m)) Then

                                ID_Sig = Br.GetItem(Pname(k) & "/" & labs(z)) ' чтение ДП 

                                Try
                                    DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", l, ID_Sig, , 1)
                                Catch ex3 As Exception
                                    DB.AddSection(Pname(k) & "_" & RegNm(k, m))                 ' создать сначала секцию в базе
                                    DB.AddItem(Pname(k) & "_" & RegNm(k, m) & "/", l, ID_Sig, , 1)
                                End Try

                            End If

                        Next    ' следующий параметр из lms
                    Next

                    ' если нет параметра - берем следующее испытание
                Catch ex2 As Exception
                    ' если нет испытания - берем следующее
                End Try





            Next

            ToolStripProgressBar1.Value = 100 * k / i : Me.Refresh()
        Next

    End Sub

    Private Sub CheckBox12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox12.CheckedChanged
        Names()
    End Sub
End Class
