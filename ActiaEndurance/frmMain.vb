Imports IniParser
Imports IniParser.Model
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.Timers
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.Automation

Public Class frmMain
    Dim ThreadLoadingApp As Thread
    Dim ThreadMain As Thread
    Dim ThreadMeasure As Thread
    Dim ThreadAlarm As Thread

    Dim configPath As String = My.Application.Info.DirectoryPath & "\Config\Config.ini"
    Dim parser As New FileIniDataParser
    Dim configData As IniData = parser.ReadFile(configPath)
    Private Sub GetUSerLevel()

        'Admin
        If UserLevel = 1 Then
            lbl_user.Text = "ADM"
            secureLvlMan.Visible = False
            secureLvlSet.Visible = False

            'Engineer
        ElseIf UserLevel = 2 Then
            lbl_user.Text = "ENG"
            secureLvlMan.Visible = False
            secureLvlSet.Visible = False

            'Operator
        ElseIf UserLevel = 3 Then
            lbl_user.Text = "OPE"
            secureLvlMan.Visible = True
            secureLvlMan.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlMan.Location = New Point(3, 0)
            secureLvlMan.Size = New Size(955, 549)
            secureLvlMan.BringToFront()

            secureLvlSet.Visible = True
            secureLvlSet.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlSet.Location = New Point(3, 0)
            secureLvlSet.Size = New Size(955, 549)
            secureLvlSet.BringToFront()

            'Quality
        ElseIf UserLevel = 4 Then
            lbl_user.Text = "QUA"
            secureLvlMan.Visible = True
            secureLvlMan.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlMan.Location = New Point(3, 0)
            secureLvlMan.Size = New Size(955, 549)
            secureLvlMan.BringToFront()

            secureLvlSet.Visible = True
            secureLvlSet.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlSet.Location = New Point(3, 0)
            secureLvlSet.Size = New Size(955, 549)
            secureLvlSet.BringToFront()
        End If
    End Sub

    Private Sub LoadingMSG(msg As String)
        Invoke(Sub()
                   MachineLog.AppendText(msg)
                   MachineLog.AppendText(Environment.NewLine)
                   MachineLog.AppendText(Environment.NewLine)
                   MachineLog.ScrollToCaret()
               End Sub)
    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ThreadLoadingApp = New Thread(AddressOf LoadingApp)
        ThreadLoadingApp.Start()

        ThreadMain = New Thread(AddressOf mainLoop)

        'ThreadMeasure = New Thread(AddressOf startMeas)

        ThreadAlarm = New Thread(AddressOf MainAlarm)


    End Sub
    Private Sub LoadingApp()
        Try
            With Config
                UserLevel = 1
                Invoke(Sub()
                           GetUSerLevel()
                       End Sub)
                LoadingMSG("Establishing connection to PLC...")
                With Config
                    .IP = configData("PLC")("IP")
                    .Port = configData("PLC")("Port")
                    Modbus.OpenPort(.IP, .Port)
                    If Modbus.ModbusConnected Then
                        'MainTab
                        ind_plc_run.BackColor = Color.Lime

                        'SettingTab
                        Invoke(Sub()
                                   lbl_plc_ind.Text = "Connected"
                                   lbl_plc_ind.BackColor = Color.LimeGreen
                               End Sub)
                    Else
                        'MainTab
                        ind_plc_run.BackColor = Color.Red

                        'SettingTab
                        Invoke(Sub()
                                   lbl_plc_ind.Text = "Disconnected"
                                   lbl_plc_ind.BackColor = Color.DarkRed
                               End Sub)
                    End If
                End With
                Thread.Sleep(500)

                LoadingMSG("Establishing connection to Keysight...")
                With Config
                    .IPSight = configData("KEYSIGHT")("IP")
                    .PortSight = configData("KEYSIGHT")("Port")
                    'ConnectToKS(.IPSight, .PortSight)
                    'If KSconnected Then
                    '    'MainTab
                    '    ind_keysight.BackColor = Color.Lime

                    '    'SettingTab
                    '    Invoke(Sub()
                    '               lbl_keysight_ind.Text = "Connected"
                    '               lbl_keysight_ind.BackColor = Color.LimeGreen
                    '           End Sub)

                    'Else
                    '    'MainTab
                    '    ind_keysight.BackColor = Color.Red

                    '    'SettingTab
                    '    Invoke(Sub()
                    '               lbl_keysight_ind.Text = "Disconnected"
                    '               lbl_keysight_ind.BackColor = Color.DarkRed
                    '           End Sub)
                    'End If
                End With

                LoadingMSG("Establishing connection to Database Server...")
                'LoadTable()
                Thread.Sleep(500)

                LoadingMSG("Reading CONFIG...")
                ReadConfig()
                Thread.Sleep(500)

                LoadingMSG("App Started")
                ind_pc_run.BackColor = Color.Lime
            End With

        Catch ex As Exception
            LoadingMSG("Error.. " + ex.Message + ", App is Closing...")
            Thread.Sleep(2000)
            End
        End Try
        ThreadAlarm.Start()
        ThreadMain.Start()
        'ThreadMeasure.Start()
        ThreadLoadingApp.Abort()
    End Sub

    Private Sub mainLoop()
        Do
            Try
                If IsConnected Then
                    'Alarm
                    With Alarm
                        'Minor Defect Extrude
                        .MinorEXT = ReadInteger(108)

                        .V101EXT = ReadBit(.MinorEXT, 0)
                        .V102EXT = ReadBit(.MinorEXT, 1)
                        .V103EXT = ReadBit(.MinorEXT, 2)
                        .V104EXT = ReadBit(.MinorEXT, 3)
                        .V105EXT = ReadBit(.MinorEXT, 4)
                        .V106EXT = ReadBit(.MinorEXT, 5)

                        'Minor Defect Retract
                        .MinorRET = ReadInteger(109)

                        .V101RET = ReadBit(.MinorRET, 0)
                        .V102RET = ReadBit(.MinorRET, 1)
                        .V103RET = ReadBit(.MinorRET, 2)
                        .V104RET = ReadBit(.MinorRET, 3)
                        .V105RET = ReadBit(.MinorRET, 4)
                        .V106RET = ReadBit(.MinorRET, 5)

                        'Major Defect
                        .Major = ReadInteger(104)

                        .AirFail = ReadBit(.Major, 0)
                        .EmgPress = ReadBit(.Major, 1)
                        .ContactorFail = ReadBit(.Major, 2)

                        'General Defect
                        .General = ReadInteger(117)

                        .MeasTO = ReadBit(.General, 0)
                        .NeedInit = ReadBit(.General, 1)
                        .NeedConf = ReadBit(.General, 2)
                    End With


                    'Manual
                    'Move CY101
                    With CY101
                        If .ExtrudeCY Then
                            'Extrude
                            Modbus.WriteInteger(1101, 1)
                            .ExtrudeCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1101, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY102
                        If .ExtrudeCY Then
                            'Extrude
                            Modbus.WriteInteger(1102, 1)
                            .ExtrudeCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1102, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY103
                        If .ExtrudeCY Then
                            'Extrude
                            Modbus.WriteInteger(1103, 1)
                            .ExtrudeCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1103, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY104
                        If .ExtrudeCY Then
                            'Extrude
                            Modbus.WriteInteger(1104, 1)
                            .ExtrudeCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1104, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY105
                        If .ExtrudeCY Then
                            'Extrude
                            Modbus.WriteInteger(1105, 1)
                            .ExtrudeCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1105, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY106
                        If .ExtrudeCY Then
                            'Extrude
                            Modbus.WriteInteger(1106, 1)
                            .ExtrudeCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1106, 2)
                            .ReturnCY = False
                        End If
                    End With

                    'Sensor
                    With CY101
                        If Modbus.ReadInteger(1201) = 1 Then
                            .MaxCY = True
                            .MinCY = False
                        ElseIf Modbus.ReadInteger(1201) = 2 Then
                            .MaxCY = False
                            .MinCY = True
                        Else
                            .MaxCY = False
                            .MinCY = False
                        End If
                    End With
                    With CY102
                        If Modbus.ReadInteger(1202) = 1 Then
                            .MaxCY = True
                            .MinCY = False
                        ElseIf Modbus.ReadInteger(1202) = 2 Then
                            .MaxCY = False
                            .MinCY = True
                        Else
                            .MaxCY = False
                            .MinCY = False
                        End If
                    End With
                    With CY103
                        If Modbus.ReadInteger(1203) = 1 Then
                            .MaxCY = True
                            .MinCY = False
                        ElseIf Modbus.ReadInteger(1203) = 2 Then
                            .MaxCY = False
                            .MinCY = True
                        Else
                            .MaxCY = False
                            .MinCY = False
                        End If
                    End With
                    With CY104
                        If Modbus.ReadInteger(1204) = 1 Then
                            .MaxCY = True
                            .MinCY = False
                        ElseIf Modbus.ReadInteger(1204) = 2 Then
                            .MaxCY = False
                            .MinCY = True
                        Else
                            .MaxCY = False
                            .MinCY = False
                        End If
                    End With
                    With CY105
                        If Modbus.ReadInteger(1205) = 1 Then
                            .MaxCY = True
                            .MinCY = False
                        ElseIf Modbus.ReadInteger(1205) = 2 Then
                            .MaxCY = False
                            .MinCY = True
                        Else
                            .MaxCY = False
                            .MinCY = False
                        End If
                    End With
                    With CY106
                        If Modbus.ReadInteger(1206) = 1 Then
                            .MaxCY = True
                            .MinCY = False
                        ElseIf Modbus.ReadInteger(1206) = 2 Then
                            .MaxCY = False
                            .MinCY = True
                        Else
                            .MaxCY = False
                            .MinCY = False
                        End If
                    End With

                    'Sensor FDUT
                    Dim FDUTAddr As Integer
                    With FDUT
                        FDUTAddr = Modbus.ReadInteger(1300)

                        If Modbus.ReadBit(FDUTAddr, 0) = 1 Then
                            .Cav1 = True
                        Else
                            .Cav1 = False
                        End If
                        If Modbus.ReadBit(FDUTAddr, 1) = 1 Then
                            .Cav2 = True
                        Else
                            .Cav2 = False
                        End If
                        If Modbus.ReadBit(FDUTAddr, 2) = 1 Then
                            .Cav3 = True
                        Else
                            .Cav3 = False
                        End If
                        If Modbus.ReadBit(FDUTAddr, 3) = 1 Then
                            .Cav4 = True
                        Else
                            .Cav4 = False
                        End If
                        If Modbus.ReadBit(FDUTAddr, 4) = 1 Then
                            .Cav5 = True
                        Else
                            .Cav5 = False
                        End If
                        If Modbus.ReadBit(FDUTAddr, 5) = 1 Then
                            .Cav6 = True
                        Else
                            .Cav6 = False
                        End If
                    End With

                    'type
                    With Type
                        Modbus.WriteInteger(10020, .NbCyc)
                    End With

                    'General Communication
                    With GeneralComm
                        'PC To PLC
                        Modbus.WriteInteger(10000, .EnVol)
                        Modbus.WriteInteger(10001, .EnCurr)
                        Modbus.WriteInteger(10002, .EnTime)
                        Modbus.WriteInteger(10003, .rdyStart)
                        Modbus.WriteInteger(10005, .FinVol)
                        Modbus.WriteInteger(10007, .FinCurr)

                        With EnCavity
                            Modbus.WriteInteger(10008, CInt(.Cav1))
                            Modbus.WriteInteger(10009, CInt(.Cav2))
                            Modbus.WriteInteger(10010, CInt(.Cav3))
                            Modbus.WriteInteger(10011, CInt(.Cav4))
                            Modbus.WriteInteger(10012, CInt(.Cav5))
                            Modbus.WriteInteger(10013, CInt(.Cav6))
                        End With

                        'PLC To PC
                        .trigVol = Modbus.ReadInteger(10004)
                        .trigSaveCurr = Modbus.ReadInteger(10006)

                        If ReadInteger(10014) = 1 Then
                            ind_cav1.BackColor = Color.Lime
                        Else
                            ind_cav1.BackColor = Color.Red
                        End If
                        If ReadInteger(10015) = 1 Then
                            ind_cav2.BackColor = Color.Lime
                        Else
                            ind_cav2.BackColor = Color.Red
                        End If
                        If ReadInteger(10016) = 1 Then
                            ind_cav3.BackColor = Color.Lime
                        Else
                            ind_cav3.BackColor = Color.Red
                        End If
                        If ReadInteger(10017) = 1 Then
                            ind_cav4.BackColor = Color.Lime
                        Else
                            ind_cav4.BackColor = Color.Red
                        End If
                        If ReadInteger(10018) = 1 Then
                            ind_cav5.BackColor = Color.Lime
                        Else
                            ind_cav5.BackColor = Color.Red
                        End If
                        If ReadInteger(10019) = 1 Then
                            ind_cav6.BackColor = Color.Lime
                        Else
                            ind_cav6.BackColor = Color.Red
                        End If

                    End With

                    'Load Manual Sensor
                    LoadSensorManual()

                    With Val
                        .Cur1 = ReadFloat(40000)
                        .Cur2 = ReadFloat(40002)
                        .Cur3 = ReadFloat(40004)
                        .Cur4 = ReadFloat(40006)
                        .Cur5 = ReadFloat(40008)
                        .Cur6 = ReadFloat(40010)

                        .Time1 = ReadInteger(10022)
                        .Time2 = ReadInteger(10023)
                        .Time3 = ReadInteger(10024)
                        .Time4 = ReadInteger(10025)
                        .Time5 = ReadInteger(10026)
                        .Time6 = ReadInteger(10027)

                        .result_1 = ReadInteger(10028)
                        .result_2 = ReadInteger(10029)
                        .result_3 = ReadInteger(10030)
                        .result_4 = ReadInteger(10031)
                        .result_5 = ReadInteger(10032)
                        .result_6 = ReadInteger(10033)

                        Invoke(Sub()
                                   TextBox1.Text = .Cur1
                                   TextBox2.Text = .Cur2
                                   TextBox3.Text = .Cur3
                                   TextBox4.Text = .Cur4
                                   TextBox5.Text = .Cur5
                                   TextBox6.Text = .Cur6

                                   TextBox13.Text = Logging.CycleCount
                                   TextBox14.Text = Logging.CycleCount
                                   TextBox15.Text = Logging.CycleCount
                                   TextBox16.Text = Logging.CycleCount
                                   TextBox17.Text = Logging.CycleCount
                                   TextBox18.Text = Logging.CycleCount

                                   TextBox19.Text = .Time1
                                   TextBox20.Text = .Time2
                                   TextBox21.Text = .Time3
                                   TextBox22.Text = .Time4
                                   TextBox23.Text = .Time5
                                   TextBox24.Text = .Time6

                                   txt_main_reference_1.Text = Logging.Reference_1
                                   txt_main_reference_2.Text = Logging.Reference_2
                                   txt_main_reference_3.Text = Logging.Reference_3
                                   txt_main_date.Text = Logging.StartDate
                                   txt_main_id.Text = Logging.OpeId
                               End Sub)
                    End With

                    Logging.CycleCount = ReadInteger(10021)

                    If GeneralComm.trigSaveCurr = 1 Then
                        Modbus.WriteInteger(10006, 0)
                        With Logging
                            ' Path to the CSV file
                            Dim filePath As String = My.Application.Info.DirectoryPath & "\Log\Log.csv"
                            ' Create a StreamWriter to write to the file
                            Using writer As New StreamWriter(filePath, append:=True)
                                ' Iterate through the list and write each row to the file
                                If EnCavity.Cav1 = "1" Then
                                    If Val.result_1 = 1 Then
                                        .FinalResult = "OK"
                                    Else
                                        .FinalResult = "NOK"
                                    End If

                                    Dim data_cav As New List(Of String()) From {
                                        New String() { .No.ToString, .StartDate, .Reference_1, .OpeId, "Cavity 1", .CycleCount.ToString, .CurrentMode, .UseVoltage, .Voltage, .UseCurrent, Val.Cur1, .UseTime, .Time, .FinalResult}
                                    }
                                    For Each row As String() In data_cav
                                        writer.WriteLine(String.Join(",", row))
                                    Next
                                End If

                                If EnCavity.Cav2 = "1" Then
                                    If Val.result_2 = 1 Then
                                        .FinalResult = "OK"
                                    Else
                                        .FinalResult = "NOK"
                                    End If

                                    Dim data_cav As New List(Of String()) From {
                                        New String() { .No.ToString, .StartDate, .Reference_1, .OpeId, "Cavity 2", .CycleCount.ToString, .CurrentMode, .UseVoltage, .Voltage, .UseCurrent, Val.Cur2, .UseTime, .Time, .FinalResult}
                                    }
                                    For Each row As String() In data_cav
                                        writer.WriteLine(String.Join(",", row))
                                    Next
                                End If

                                If EnCavity.Cav3 = "1" Then
                                    If Val.result_3 = 1 Then
                                        .FinalResult = "OK"
                                    Else
                                        .FinalResult = "NOK"
                                    End If

                                    Dim data_cav As New List(Of String()) From {
                                        New String() { .No.ToString, .StartDate, .Reference_2, .OpeId, "Cavity 3", .CycleCount.ToString, .CurrentMode, .UseVoltage, .Voltage, .UseCurrent, Val.Cur3, .UseTime, .Time, .FinalResult}
                                    }
                                    For Each row As String() In data_cav
                                        writer.WriteLine(String.Join(",", row))
                                    Next
                                End If

                                If EnCavity.Cav4 = "1" Then
                                    If Val.result_4 = 1 Then
                                        .FinalResult = "OK"
                                    Else
                                        .FinalResult = "NOK"
                                    End If

                                    Dim data_cav As New List(Of String()) From {
                                        New String() { .No.ToString, .StartDate, .Reference_2, .OpeId, "Cavity 4", .CycleCount.ToString, .CurrentMode, .UseVoltage, .Voltage, .UseCurrent, Val.Cur4, .UseTime, .Time, .FinalResult}
                                    }
                                    For Each row As String() In data_cav
                                        writer.WriteLine(String.Join(",", row))
                                    Next
                                End If

                                If EnCavity.Cav5 = "1" Then
                                    If Val.result_5 = 1 Then
                                        .FinalResult = "OK"
                                    Else
                                        .FinalResult = "NOK"
                                    End If

                                    Dim data_cav As New List(Of String()) From {
                                        New String() { .No.ToString, .StartDate, .Reference_3, .OpeId, "Cavity 5", .CycleCount.ToString, .CurrentMode, .UseVoltage, .Voltage, .UseCurrent, Val.Cur5, .UseTime, .Time, .FinalResult}
                                    }
                                    For Each row As String() In data_cav
                                        writer.WriteLine(String.Join(",", row))
                                    Next
                                End If

                                If EnCavity.Cav6 = "1" Then
                                    If Val.result_6 = 1 Then
                                        .FinalResult = "OK"
                                    Else
                                        .FinalResult = "NOK"
                                    End If

                                    Dim data_cav As New List(Of String()) From {
                                        New String() { .No.ToString, .StartDate, .Reference_3, .OpeId, "Cavity 6", .CycleCount.ToString, .CurrentMode, .UseVoltage, .Voltage, .UseCurrent, Val.Cur6, .UseTime, .Time, .FinalResult}
                                    }
                                    For Each row As String() In data_cav
                                        writer.WriteLine(String.Join(",", row))
                                    Next
                                End If
                            End Using
                        End With
                        Modbus.WriteInteger(10007, 1)
                    End If
                End If
            Catch ex As Exception
                IsConnected = False
            End Try

            Thread.Sleep(40)
        Loop
    End Sub
    Private Sub MainAlarm()
        Dim txtColor As Color = Color.DarkRed
        Dim txtBorder As BorderStyle = BorderStyle.FixedSingle

        Dim responseNo As Integer = 0
        Dim responseKSNo As Integer = 0
        Do
            With Alarm
                If IsConnected Then
                    responseNo = 0
                    'Minor Defect (Cylinder)
                    'Extrude
                    If .V101EXT = 1 Then
                        .alarmMsg = "ALARM : V101 EXT Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V102EXT = 1 Then
                        .alarmMsg = "ALARM : V102 EXT Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V103EXT = 1 Then
                        .alarmMsg = "ALARM : V103 EXT Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V104EXT = 1 Then
                        .alarmMsg = "ALARM : V104 EXT Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V105EXT = 1 Then
                        .alarmMsg = "ALARM : V105 EXT Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V106EXT = 1 Then
                        .alarmMsg = "ALARM : V106 EXT Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle

                        'Retract
                    ElseIf .V101RET = 1 Then
                        .alarmMsg = "ALARM : V101 RET Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V102RET = 1 Then
                        .alarmMsg = "ALARM : V102 RET Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V103RET = 1 Then
                        .alarmMsg = "ALARM : V103 RET Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V104RET = 1 Then
                        .alarmMsg = "ALARM : V104 RET Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V105RET = 1 Then
                        .alarmMsg = "ALARM : V105 RET Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .V106RET = 1 Then
                        .alarmMsg = "ALARM : V106 RET Discrepancy"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle


                        'Major Defect
                    ElseIf .AirFail Then
                        .alarmMsg = "ALARM : Air failure, please check air supply!"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .EmgPress Then
                        .alarmMsg = "ALARM : Emergency button pressed!"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .ContactorFail Then
                        .alarmMsg = "ALARM : Contactor 24V fail, please check!"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle

                        'Defect Alarm
                    ElseIf .MeasTO Then
                        .alarmMsg = "ALARM : Measure timeout"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .NeedInit Then
                        .alarmMsg = "ALARM : Need to initialize to start machine"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                    ElseIf .NeedConf Then
                        .alarmMsg = "ALARM : Need to set product configuration"
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle

                        'No Alarm
                    Else
                        .alarmMsg = ""
                        txtColor = Color.Transparent
                        txtBorder = BorderStyle.None

                    End If

                Else
                    .alarmMsg = "ALARM : Modbus Connection Error please check the PLC communication"

                    'MainTab
                    txtColor = Color.DarkRed
                    txtBorder = BorderStyle.FixedSingle
                    ind_plc_run.BackColor = Color.Red

                    'SettingTab
                    Invoke(Sub()
                               lbl_plc_ind.Text = "Disconnected"
                               lbl_plc_ind.BackColor = Color.DarkRed
                           End Sub)

                    If txt_alarm.Text = "ALARM : Modbus Connection Error please check the PLC communication" And responseNo = 0 Then
                        Dim response As DialogResult = MessageBox.Show("Do you want to reconnect to PLC?", "PLC COMMUNICATION ERROR", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
                        If response = DialogResult.Yes Then
                            Try
                                Modbus.OpenPort(Config.IP, Config.Port)
                                If IsConnected Then
                                    'MainTab
                                    ind_plc_run.BackColor = Color.Lime

                                    'SettingTab
                                    Invoke(Sub()
                                               lbl_plc_ind.Text = "Connected"
                                               lbl_plc_ind.BackColor = Color.LimeGreen
                                           End Sub)
                                End If
                            Catch ex As Exception
                            End Try
                        Else
                            responseNo = 1
                        End If
                    End If
                End If
                '.alarmMsg = "ALARM : KEYSIGHT Connection Error please check the KEYSIGHT communication"

                ''MainTab
                'txtColor = Color.DarkRed
                'txtBorder = BorderStyle.FixedSingle
                'ind_keysight.BackColor = Color.Red

                ''SettingTab
                'Invoke(Sub()
                '           lbl_keysight_ind.Text = "Disconnected"
                '           lbl_keysight_ind.BackColor = Color.DarkRed
                '       End Sub)
                'If txt_alarm.Text = "ALARM : KEYSIGHT Connection Error please check the KEYSIGHT communication" And responseKSNo = 0 Then
                '    Dim responseKS As DialogResult = MessageBox.Show("Do you want to reconnect to KEYSIGHT?", "KEYSIGHT COMMUNICATION ERROR", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
                '    If responseKS = DialogResult.Yes Then
                '        Try
                '            ConnectToKS(Config.IPSight, Config.PortSight)
                '            If KSconnected Then
                '                'MainTab
                '                ind_keysight.BackColor = Color.Lime

                '                'SettingTab
                '                Invoke(Sub()
                '                           lbl_keysight_ind.Text = "Connected"
                '                           lbl_keysight_ind.BackColor = Color.LimeGreen
                '                       End Sub)
                '            End If
                '        Catch ex As Exception
                '        End Try
                '    Else
                '        responseKSNo = 1
                '    End If
                'End If



                Invoke(Sub()
                           txt_alarm.Text = .alarmMsg
                           txt_alarm.BackColor = txtColor
                           txt_alarm.BorderStyle = txtBorder
                       End Sub)

            End With
            Thread.Sleep(150)
        Loop
    End Sub

    Private Sub ReadConfig()
        With Config
            txt_IP_PLC.Text = configData("PLC")("IP")
            txt_Port_PLC.Text = configData("PLC")("Port")

            txt_IP_keysight.Text = configData("KEYSIGHT")("IP")
            txt_Port_keysight.Text = configData("KEYSIGHT")("Port")

            'txt_main_reference_1.Text = configData("Config")("Reference_1")
            'txt_main_reference_2.Text = configData("Config")("Reference_2")
            'txt_main_reference_3.Text = configData("Config")("Reference_3")
            txt_cfg_date.Text = configData("Config")("StartTime")
            'txt_main_date.Text = configData("Config")("StartTime")
            txt_cfg_id.Text = configData("Config")("OpeName")
            'txt_main_id.Text = configData("Config")("OpeName")

            cb_cfg_ref_1.Text = configData("Config")("Reference_1")
            cb_cfg_ref_2.Text = configData("Config")("Reference_2")
            cb_cfg_ref_3.Text = configData("Config")("Reference_3")

            .Reference_1 = configData("Config")("Reference_1")
            .Reference_2 = configData("Config")("Reference_2")
            .Reference_3 = configData("Config")("Reference_3")
            .StartTime = configData("Config")("StartTime")
            .OperatorID = configData("Config")("OpeName")

            Logging.Reference_1 = .Reference_1
            Logging.Reference_2 = .Reference_2
            Logging.Reference_3 = .Reference_3
            Logging.StartDate = .StartTime
            Logging.OpeId = .OperatorID
        End With

        With Type
            .CurrMode = configData("Type")("CurrMode")
            .VolCtrl = configData("Type")("UseVol")
            .CurrCtrl = configData("Type")("UseCurr")
            .TimeCtrl = configData("Type")("UseTime")
            .VoltVal = configData("Type")("ValVol")
            .CurrVal = configData("Type")("ValCurr")
            .TimeVal = configData("Type")("ValTime")
            .NbCyc = configData("Type")("NbCyc")

            cb_cfg_mode.Text = .CurrMode
            txt_cfg_v_val.Text = .VoltVal
            txt_cfg_c_val.Text = .CurrVal
            txt_cfg_time_val.Text = .TimeVal
            txt_cfg_nbc.Text = .NbCyc
            Logging.CurrentMode = .CurrMode
            Logging.Voltage = .VoltVal
            Logging.Current = .CurrVal
            Logging.Time = .TimeVal

            If .VolCtrl = "1" Then
                cb_cfg_use_v.Text = "Enable"
                Logging.UseVoltage = "Use Voltage"
            Else
                cb_cfg_use_v.Text = "Disable"
                Logging.UseVoltage = "Skip Voltage"
            End If
            If .CurrCtrl Then
                cb_cfg_use_c.Text = "Enable"
                Logging.UseCurrent = "Use Current"
            Else
                cb_cfg_use_c.Text = "Disable"
                Logging.UseCurrent = "Skip Current"
            End If
            If .TimeCtrl Then
                cb_cfg_use_time.Text = "Enable"
                Logging.UseTime = "Use Time"
            Else
                cb_cfg_use_time.Text = "Disable"
                Logging.UseTime = "Skip Time"
            End If
        End With

        With UseCavity
            .Cav1 = configData("CtrlCav")("UseCav1")
            .Cav2 = configData("CtrlCav")("UseCav2")
            .Cav3 = configData("CtrlCav")("UseCav3")
            .Cav4 = configData("CtrlCav")("UseCav4")
            .Cav5 = configData("CtrlCav")("UseCav5")
            .Cav6 = configData("CtrlCav")("UseCav6")

            If .Cav1 = "1" Then
                cb_cfg_use_cav1.Text = "Enable"
            Else
                cb_cfg_use_cav1.Text = "Disable"
            End If
            If .Cav2 = "1" Then
                cb_cfg_use_cav2.Text = "Enable"
            Else
                cb_cfg_use_cav2.Text = "Disable"
            End If
            If .Cav3 = "1" Then
                cb_cfg_use_cav3.Text = "Enable"
            Else
                cb_cfg_use_cav3.Text = "Disable"
            End If
            If .Cav4 = "1" Then
                cb_cfg_use_cav4.Text = "Enable"
            Else
                cb_cfg_use_cav4.Text = "Disable"
            End If
            If .Cav5 = "1" Then
                cb_cfg_use_cav5.Text = "Enable"
            Else
                cb_cfg_use_cav5.Text = "Disable"
            End If
            If .Cav6 = "1" Then
                cb_cfg_use_cav6.Text = "Enable"
            Else
                cb_cfg_use_cav6.Text = "Disable"
            End If
        End With

        'Database Filter
        With Filter
            .CurrType = "AC"
            .UseVol = "YES"
            .UseCurr = "YES"
            .UseTime = "YES"
        End With

        With Logging
            .No = configData("Config")("Count")
        End With
    End Sub

    Private Sub LoadSensorManual()
        'CY101
        If CY101.MaxCY Then
            ind_v101_max.BackColor = Color.Lime
        Else
            ind_v101_max.BackColor = Color.Red
        End If
        If CY101.MinCY Then
            ind_v101_min.BackColor = Color.Lime
        Else
            ind_v101_min.BackColor = Color.Red
        End If

        'CY102
        If CY102.MaxCY Then
            ind_v102_max.BackColor = Color.Lime
        Else
            ind_v102_max.BackColor = Color.Red
        End If
        If CY102.MinCY Then
            ind_v102_min.BackColor = Color.Lime
        Else
            ind_v102_min.BackColor = Color.Red
        End If

        'CY103
        If CY103.MaxCY Then
            ind_v103_max.BackColor = Color.Lime
        Else
            ind_v103_max.BackColor = Color.Red
        End If
        If CY103.MinCY Then
            ind_v103_min.BackColor = Color.Lime
        Else
            ind_v103_min.BackColor = Color.Red
        End If

        'CY104
        If CY104.MaxCY Then
            ind_v104_max.BackColor = Color.Lime
        Else
            ind_v104_max.BackColor = Color.Red
        End If
        If CY104.MinCY Then
            ind_v104_min.BackColor = Color.Lime
        Else
            ind_v104_min.BackColor = Color.Red
        End If

        'CY105
        If CY105.MaxCY Then
            ind_v105_max.BackColor = Color.Lime
        Else
            ind_v105_max.BackColor = Color.Red
        End If
        If CY105.MinCY Then
            ind_v105_min.BackColor = Color.Lime
        Else
            ind_v105_min.BackColor = Color.Red
        End If

        'CY106
        If CY106.MaxCY Then
            ind_v106_max.BackColor = Color.Lime
        Else
            ind_v106_max.BackColor = Color.Red
        End If
        If CY106.MinCY Then
            ind_v106_min.BackColor = Color.Lime
        Else
            ind_v106_min.BackColor = Color.Red
        End If

        With FDUT
            If .Cav1 Then
                ind_fdut1.BackColor = Color.Lime
            Else
                ind_fdut1.BackColor = Color.Red
            End If
            If .Cav2 Then
                ind_fdut2.BackColor = Color.Lime
            Else
                ind_fdut2.BackColor = Color.Red
            End If
            If .Cav3 Then
                ind_fdut3.BackColor = Color.Lime
            Else
                ind_fdut3.BackColor = Color.Red
            End If
            If .Cav4 Then
                ind_fdut4.BackColor = Color.Lime
            Else
                ind_fdut4.BackColor = Color.Red
            End If
            If .Cav5 Then
                ind_fdut5.BackColor = Color.Lime
            Else
                ind_fdut5.BackColor = Color.Red
            End If
            If .Cav6 Then
                ind_fdut6.BackColor = Color.Lime
            Else
                ind_fdut6.BackColor = Color.Red
            End If
        End With
    End Sub
    Private Sub btn_save_config_Click(sender As Object, e As EventArgs) Handles btn_save_config.Click
        If cb_cfg_ref_1.Text = "" Or cb_cfg_ref_2.Text = "" Or cb_cfg_ref_3.Text = "" Or txt_cfg_date.Text = "" Or txt_cfg_id.Text = "" Or txt_cfg_c_val.Text = "" Or txt_cfg_v_val.Text = "" Or txt_cfg_time_val.Text = "" Or txt_cfg_nbc.Text = "" Or cb_cfg_mode.Text = "" Or cb_cfg_use_v.Text = "" Or cb_cfg_use_c.Text = "" Or cb_cfg_use_time.Text = "" Or cb_cfg_use_cav1.Text = "" Or cb_cfg_use_cav2.Text = "" Or cb_cfg_use_cav3.Text = "" Or cb_cfg_use_cav4.Text = "" Or cb_cfg_use_cav5.Text = "" Or cb_cfg_use_cav6.Text = "" Then
            MsgBox("Please fill all data!")
            Exit Sub
        Else
            'With Config
            configData("Config")("Reference_1") = cb_cfg_ref_1.Text
            Logging.Reference_1 = cb_cfg_ref_1.Text
            configData("Config")("Reference_2") = cb_cfg_ref_2.Text
            Logging.Reference_2 = cb_cfg_ref_2.Text
            configData("Config")("Reference_3") = cb_cfg_ref_3.Text
            Logging.Reference_3 = cb_cfg_ref_3.Text
            configData("Config")("StartTime") = txt_cfg_date.Text
            Logging.StartDate = txt_cfg_date.Text
            configData("Config")("OpeName") = txt_cfg_id.Text
            Logging.OpeId = txt_cfg_id.Text

            configData("Type")("CurrMode") = cb_cfg_mode.Text
            Logging.CurrentMode = cb_cfg_mode.Text
            configData("Type")("ValVol") = txt_cfg_v_val.Text
            configData("Type")("ValCurr") = txt_cfg_c_val.Text
            configData("Type")("ValTime") = txt_cfg_time_val.Text
            configData("Type")("NbCyc") = txt_cfg_nbc.Text
            If IsNumeric(txt_cfg_nbc.Text) Then
                Type.NbCyc = CInt(txt_cfg_nbc.Text)
            End If

            'With Type
            If cb_cfg_use_v.Text = "Enable" Then
                configData("Type")("UseVol") = "1"
                Logging.UseVoltage = "Use Voltage"
            Else
                configData("Type")("UseVol") = "0"
                Logging.UseVoltage = "Skip Voltage"
            End If
            If cb_cfg_use_c.Text = "Enable" Then
                configData("Type")("UseCurr") = "1"
                Logging.UseCurrent = "Use Current"
            Else
                configData("Type")("UseCurr") = "0"
                Logging.UseCurrent = "Skip Current"
            End If
            If cb_cfg_use_time.Text = "Enable" Then
                configData("Type")("UseTime") = "1"
                Logging.UseTime = "Use Time"
            Else
                configData("Type")("UseTime") = "0"
                Logging.UseTime = "Skip Time"
            End If

            'With Cavity
            If cb_cfg_use_cav1.Text = "Enable" Then
                configData("CtrlCav")("UseCav1") = "1"
            Else
                configData("CtrlCav")("UseCav1") = "0"
            End If
            If cb_cfg_use_cav2.Text = "Enable" Then
                configData("CtrlCav")("UseCav2") = "1"
            Else
                configData("CtrlCav")("UseCav2") = "0"
            End If
            If cb_cfg_use_cav3.Text = "Enable" Then
                configData("CtrlCav")("UseCav3") = "1"
            Else
                configData("CtrlCav")("UseCav3") = "0"
            End If
            If cb_cfg_use_cav4.Text = "Enable" Then
                configData("CtrlCav")("UseCav4") = "1"
            Else
                configData("CtrlCav")("UseCav4") = "0"
            End If
            If cb_cfg_use_cav5.Text = "Enable" Then
                configData("CtrlCav")("UseCav5") = "1"
            Else
                configData("CtrlCav")("UseCav5") = "0"
            End If
            If cb_cfg_use_cav6.Text = "Enable" Then
                configData("CtrlCav")("UseCav6") = "1"
            Else
                configData("CtrlCav")("UseCav6") = "0"
            End If

            'Write To PLC
            With GeneralComm
                If cb_cfg_use_v.Text = "Enable" Then
                    .EnVol = 1
                Else
                    .EnVol = 0
                End If
                If cb_cfg_use_c.Text = "Enable" Then
                    .EnCurr = 1
                Else
                    .EnCurr = 0
                End If
                If cb_cfg_use_time.Text = "Enable" Then
                    .EnTime = 1
                Else
                    .EnTime = 0
                End If
            End With

            With EnCavity
                If cb_cfg_use_cav1.Text = "Enable" Then
                    .Cav1 = 1
                Else
                    .Cav1 = 0
                End If
                If cb_cfg_use_cav2.Text = "Enable" Then
                    .Cav2 = 1
                Else
                    .Cav2 = 0
                End If
                If cb_cfg_use_cav3.Text = "Enable" Then
                    .Cav3 = 1
                Else
                    .Cav3 = 0
                End If
                If cb_cfg_use_cav4.Text = "Enable" Then
                    .Cav4 = 1
                Else
                    .Cav4 = 0
                End If
                If cb_cfg_use_cav5.Text = "Enable" Then
                    .Cav5 = 1
                Else
                    .Cav5 = 0
                End If
                If cb_cfg_use_cav6.Text = "Enable" Then
                    .Cav6 = 1
                Else
                    .Cav6 = 0
                End If
            End With
        End If
        'Write To INI File
        parser.WriteFile(configPath, configData)
        MsgBox("Save Configuration Successful")
    End Sub

    Private Sub txt_cfg_date_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txt_cfg_date.MouseDoubleClick
        txt_cfg_date.Text = Date.Now.ToShortDateString
    End Sub
    Private Sub txt_main_date_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txt_main_date.MouseDoubleClick
        txt_main_date.Text = Date.Now.ToShortDateString
    End Sub
    Private Sub current_mode_ac_Click(sender As Object, e As EventArgs) Handles current_mode_ac.Click
        current_mode_dc.Checked = False
        Filter.CurrType = "AC"
    End Sub

    Private Sub current_mode_dc_Click(sender As Object, e As EventArgs) Handles current_mode_dc.Click
        current_mode_ac.Checked = False
        Filter.CurrType = "DC"
    End Sub

    Private Sub use_vol_yes_Click(sender As Object, e As EventArgs) Handles use_vol_yes.Click
        use_vol_no.Checked = False
        Filter.UseVol = "Enable"
    End Sub

    Private Sub use_vol_no_Click(sender As Object, e As EventArgs) Handles use_vol_no.Click
        use_vol_yes.Checked = False
        Filter.UseVol = "Disable"
    End Sub

    Private Sub use_curr_yes_Click(sender As Object, e As EventArgs) Handles use_curr_yes.Click
        use_curr_no.Checked = False
        Filter.UseCurr = "Enable"
    End Sub

    Private Sub use_curr_no_Click(sender As Object, e As EventArgs) Handles use_curr_no.Click
        use_curr_yes.Checked = False
        Filter.UseCurr = "Disable"
    End Sub

    Private Sub use_time_yes_Click(sender As Object, e As EventArgs) Handles use_time_yes.Click
        use_time_no.Checked = False
        Filter.UseTime = "Enable"
    End Sub

    Private Sub use_time_no_Click(sender As Object, e As EventArgs) Handles use_time_no.Click
        use_time_yes.Checked = False
        Filter.UseTime = "Disable"
    End Sub

    Private Sub btn_search_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btn_clear_search_Click(sender As Object, e As EventArgs) Handles btn_clear_search.Click
        current_mode_ac.CheckState = CheckState.Unchecked
        current_mode_dc.CheckState = CheckState.Unchecked
        use_vol_yes.CheckState = CheckState.Unchecked
        use_vol_no.CheckState = CheckState.Unchecked
        use_curr_yes.CheckState = CheckState.Unchecked
        use_curr_no.CheckState = CheckState.Unchecked
        use_time_yes.CheckState = CheckState.Unchecked
        use_time_no.CheckState = CheckState.Unchecked
    End Sub

    Private Sub btn_v101_ext_Click(sender As Object, e As EventArgs) Handles btn_v101_ext.Click
        CY101.ExtrudeCY = True
    End Sub

    Private Sub btn_v101_ret_Click(sender As Object, e As EventArgs) Handles btn_v101_ret.Click
        CY101.ReturnCY = True
    End Sub

    Private Sub btn_v102_ext_Click(sender As Object, e As EventArgs) Handles btn_v102_ext.Click
        CY102.ExtrudeCY = True
    End Sub

    Private Sub btn_v102_ret_Click(sender As Object, e As EventArgs) Handles btn_v102_ret.Click
        CY102.ReturnCY = True
    End Sub

    Private Sub btn_v103_ext_Click(sender As Object, e As EventArgs) Handles btn_v103_ext.Click
        CY103.ExtrudeCY = True
    End Sub

    Private Sub btn_v103_ret_Click(sender As Object, e As EventArgs) Handles btn_v103_ret.Click
        CY103.ReturnCY = True
    End Sub

    Private Sub btn_v104_ext_Click(sender As Object, e As EventArgs) Handles btn_v104_ext.Click
        CY104.ExtrudeCY = True
    End Sub

    Private Sub btn_v104_ret_Click(sender As Object, e As EventArgs) Handles btn_v104_ret.Click
        CY104.ReturnCY = True
    End Sub

    Private Sub btn_v105_ext_Click(sender As Object, e As EventArgs) Handles btn_v105_ext.Click
        CY105.ExtrudeCY = True
    End Sub

    Private Sub btn_v105_ret_Click(sender As Object, e As EventArgs) Handles btn_v105_ret.Click
        CY105.ReturnCY = True
    End Sub

    Private Sub btn_v106_ext_Click(sender As Object, e As EventArgs) Handles btn_v106_ext.Click
        CY106.ExtrudeCY = True
    End Sub

    Private Sub btn_v106_ret_Click(sender As Object, e As EventArgs) Handles btn_v106_ret.Click
        CY106.ReturnCY = True
    End Sub

    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        End
    End Sub

    Private Sub btn_save_setting_Click(sender As Object, e As EventArgs) Handles btn_save_setting.Click
        configData("PLC")("IP") = txt_IP_PLC.Text
        configData("PLC")("Port") = txt_Port_PLC.Text

        configData("KEYSIGHT")("IP") = txt_IP_keysight.Text
        configData("KEYSIGHT")("Port") = txt_Port_keysight.Text

        MsgBox("Save Setting Successful")
    End Sub

    Private Sub btn_connect_plc_Click(sender As Object, e As EventArgs) Handles btn_connect_plc.Click
        Try
            Modbus.OpenPort(txt_IP_PLC.Text, txt_Port_PLC.Text)
            If IsConnected Then
                'MainTab
                ind_plc_run.BackColor = Color.Lime

                'SettingTab
                Invoke(Sub()
                           lbl_plc_ind.Text = "Connected"
                           lbl_plc_ind.BackColor = Color.LimeGreen
                       End Sub)
            End If

        Catch ex As Exception
            MessageBox.Show("Error connecting to PLC: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Invoke(Sub()
                       lbl_plc_ind.Text = "Disconnected"
                       lbl_plc_ind.BackColor = Color.DarkRed
                   End Sub)
            IsConnected = False
        End Try
    End Sub

    Private Sub btn_disconnect_plc_Click(sender As Object, e As EventArgs) Handles btn_disconnect_plc.Click
        Try
            Modbus.ClosePort()
            Invoke(Sub()
                       lbl_plc_ind.Text = "Disconnected"
                       lbl_plc_ind.BackColor = Color.DarkRed
                   End Sub)

            IsConnected = False
        Catch ex As Exception
            MessageBox.Show("Error disconnecting from PLC: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub PostGB(Page As Integer)
        If Page = 1 Then
            gb_meas1.Location = New Point(496, 15)
            gb_meas1.Size = New Size(445, 432)
            gb_meas1.BringToFront()
            gb_meas1.Visible = True
        Else
            gb_meas1.Visible = False
        End If
        If Page = 2 Then
            gb_meas2.Location = New Point(496, 15)
            gb_meas2.Size = New Size(445, 432)
            gb_meas2.BringToFront()
            gb_meas2.Visible = True
        Else
            gb_meas2.Visible = False
        End If
    End Sub
    Private Sub btn_next_Click(sender As Object, e As EventArgs) Handles btn_next.Click
        PostGB(2)
    End Sub
    Private Sub btn_back_Click(sender As Object, e As EventArgs) Handles btn_back.Click
        PostGB(1)
    End Sub
    Private Sub lbl_user_Click(sender As Object, e As EventArgs) Handles lbl_user.Click
        Hide()
        frmLogin.ShowDialog()
        GetUSerLevel()
        Show()
    End Sub

    Private Sub btn_connect_keysight_Click(sender As Object, e As EventArgs) Handles btn_connect_keysight.Click
        Try
            ConnectToKS(txt_IP_keysight.Text, txt_Port_keysight.Text)
            If KSconnected Then
                'MainTab
                ind_keysight.BackColor = Color.Lime

                'SettingTab
                Invoke(Sub()
                           lbl_keysight_ind.Text = "Connected"
                           lbl_keysight_ind.BackColor = Color.LimeGreen
                       End Sub)
            End If

        Catch ex As Exception
            MessageBox.Show("Error connecting to KEYSIGHT: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Invoke(Sub()
                       lbl_keysight_ind.Text = "Disconnected"
                       lbl_keysight_ind.BackColor = Color.DarkRed
                   End Sub)
            KSconnected = False
        End Try
    End Sub
    Private Sub btn_disconnect_keysight_Click(sender As Object, e As EventArgs) Handles btn_disconnect_keysight.Click
        Try
            DisconnectFromKS()
            Invoke(Sub()
                       lbl_plc_ind.Text = "Disconnected"
                       lbl_plc_ind.BackColor = Color.DarkRed
                   End Sub)

            KSconnected = False
        Catch ex As Exception
            MessageBox.Show("Error disconnecting from KEYSIGHT: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btn_run_MouseDown(sender As Object, e As MouseEventArgs) Handles btn_run.MouseDown
        GeneralComm.rdyStart = 1
        Logging.No += 1
        configData("Config")("Count") = Logging.No
        parser.WriteFile(configPath, configData)
    End Sub

    Private Sub btn_run_MouseUp(sender As Object, e As MouseEventArgs) Handles btn_run.MouseUp
        GeneralComm.rdyStart = 0
    End Sub
    Dim rowvalue As String
    Dim cellvalue(20) As String
    Dim columnIndex As Integer = 1
    Private Sub btn_refresh_csv_Click(sender As Object, e As EventArgs) Handles btn_search.Click
        Dim streamReader As IO.StreamReader = New IO.StreamReader(My.Application.Info.DirectoryPath & "\Log\Log.csv")
        Dim searchCurrentMode As String = ""
        Dim searchVoltage As String = ""
        Dim searchCurrent As String = ""
        Dim searchTime As String = ""

        If current_mode_ac.Checked = True Then
            searchCurrentMode = "AC"
        ElseIf current_mode_dc.Checked = True Then
            searchCurrentMode = "DC"
        End If

        If use_vol_yes.Checked = True Then
            searchVoltage = "Use Voltage"
        ElseIf use_vol_no.Checked = True Then
            searchVoltage = "Skip Voltage"
        End If

        If use_curr_yes.Checked = True Then
            searchCurrent = "Use Current"
        ElseIf use_curr_no.Checked = True Then
            searchCurrent = "Skip Current"
        End If

        If use_time_yes.Checked = True Then
            searchTime = "Use Time"
        ElseIf use_time_no.Checked = True Then
            searchTime = "Skip Time"
        End If

        DataGridView2.Rows.Clear()
        Dim rows As New List(Of String())

        While streamReader.Peek() <> -1
            rowvalue = streamReader.ReadLine()
            cellvalue = rowvalue.Split(","c)

            If cellvalue(6).Contains(searchCurrentMode) AndAlso
                cellvalue(7).Contains(searchVoltage) AndAlso
                cellvalue(9).Contains(searchCurrent) AndAlso
                cellvalue(11).Contains(searchTime) Then
                rows.Add(cellvalue)
            End If
        End While
        streamReader.Close()

        For Each row As String() In rows
            DataGridView2.Rows.Add(row)
        Next
    End Sub

    Dim _SaveFileDialog As New SaveFileDialog
    Private Sub btn_select_Click(sender As Object, e As EventArgs) Handles btn_select.Click
        _SaveFileDialog.Filter = "CSV Files|*.csv"
        _SaveFileDialog.Title = "Export-csv"
        If _SaveFileDialog.ShowDialog = DialogResult.OK Then
            txt_file_location.Text = _SaveFileDialog.FileName
        End If
    End Sub

    Private Sub btn_export_Click(sender As Object, e As EventArgs) Handles btn_export.Click
        If txt_file_location.Text <> "" Then
            If DataGridView2.RowCount > 0 Then
                Dim value As String = ""
                Dim dr As New DataGridViewRow()

                Dim swOut As StreamWriter = File.CreateText(_SaveFileDialog.FileName)

                'write header rows to csv
                For i As Integer = 0 To DataGridView2.Columns.Count - 1
                    If i > 0 Then
                        swOut.Write(",")
                    End If
                    swOut.Write(DataGridView2.Columns(i).HeaderText)
                Next

                swOut.WriteLine()

                'write DataGridView rows to csv
                ProgressBarExport.Minimum = 0
                ProgressBarExport.Maximum = DataGridView2.Rows.Count - 1
                For j As Integer = 0 To DataGridView2.Rows.Count - 1
                    If j > 0 Then
                        swOut.WriteLine()
                    End If

                    dr = DataGridView2.Rows(j)

                    For i As Integer = 0 To DataGridView2.Columns.Count - 1
                        If i > 0 Then
                            swOut.Write(",")
                        End If
                        If IsDBNull(dr.Cells(i).Value) Then
                            value = "0"
                        Else
                            value = CStr(dr.Cells(i).Value)
                        End If
                        swOut.Write(value)
                    Next
                    ProgressBarExport.Value = j
                Next

                swOut.Close()
            End If
        End If
    End Sub
End Class