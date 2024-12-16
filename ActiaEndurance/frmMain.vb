Imports IniParser
Imports IniParser.Model
Imports System.Threading
Imports System.Data.SqlClient

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

            secureLvlSet.Visible = True
            secureLvlSet.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlSet.Location = New Point(3, 0)
            secureLvlSet.Size = New Size(955, 549)

            'Quality
        ElseIf UserLevel = 4 Then
            lbl_user.Text = "QUA"
            secureLvlMan.Visible = True
            secureLvlMan.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlMan.Location = New Point(3, 0)
            secureLvlMan.Size = New Size(955, 549)

            secureLvlSet.Visible = True
            secureLvlSet.Text = "Your Security Level Is Too Low To See This Tab!!!"
            secureLvlSet.Location = New Point(3, 0)
            secureLvlSet.Size = New Size(955, 549)
        End If
    End Sub

    Private Sub LoadingMSG(msg As String)
        Invoke(Sub()
                   MachineLog.AppendText(msg)
                   MachineLog.AppendText(Environment.NewLine)
                   MachineLog.AppendText(Environment.NewLine)
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
                UserLevel = 3
                GetUSerLevel()
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
                    ConnectToKS(.IPSight, .PortSight)
                    If KSconnected Then
                        'MainTab
                        ind_keysight.BackColor = Color.Lime

                        'SettingTab
                        Invoke(Sub()
                                   lbl_keysight_ind.Text = "Connected"
                                   lbl_keysight_ind.BackColor = Color.LimeGreen
                               End Sub)

                    Else
                        'MainTab
                        ind_keysight.BackColor = Color.Red

                        'SettingTab
                        Invoke(Sub()
                                   lbl_keysight_ind.Text = "Disconnected"
                                   lbl_keysight_ind.BackColor = Color.DarkRed
                               End Sub)
                    End If
                End With

                LoadingMSG("Establishing connection to Database Server...")
                LoadTable()
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
                        If .ExtractCY Then
                            'Extract
                            Modbus.WriteInteger(1101, 1)
                            .ExtractCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1101, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY102
                        If .ExtractCY Then
                            'Extract
                            Modbus.WriteInteger(1102, 1)
                            .ExtractCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1102, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY103
                        If .ExtractCY Then
                            'Extract
                            Modbus.WriteInteger(1103, 1)
                            .ExtractCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1103, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY104
                        If .ExtractCY Then
                            'Extract
                            Modbus.WriteInteger(1104, 1)
                            .ExtractCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1104, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY105
                        If .ExtractCY Then
                            'Extract
                            Modbus.WriteInteger(1105, 1)
                            .ExtractCY = False
                        ElseIf .ReturnCY Then
                            'Retract
                            Modbus.WriteInteger(1105, 2)
                            .ReturnCY = False
                        End If
                    End With
                    With CY106
                        If .ExtractCY Then
                            'Extract
                            Modbus.WriteInteger(1106, 1)
                            .ExtractCY = False
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
                            Modbus.WriteInteger(10008, .Cav1)
                            Modbus.WriteInteger(10009, .Cav2)
                            Modbus.WriteInteger(10010, .Cav3)
                            Modbus.WriteInteger(10011, .Cav4)
                            Modbus.WriteInteger(10012, .Cav5)
                            Modbus.WriteInteger(10013, .Cav6)
                        End With

                        'PLC To PC
                        .trigVol = Modbus.ReadInteger(10004)
                        .trigCurr = Modbus.ReadInteger(10006)

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

                    If KSconnected Then

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
                        .alarmMsg = "ALARM : KEYSIGHT Connection Error please check the KEYSIGHT communication"

                        'MainTab
                        txtColor = Color.DarkRed
                        txtBorder = BorderStyle.FixedSingle
                        ind_keysight.BackColor = Color.Red

                        'SettingTab
                        Invoke(Sub()
                                   lbl_keysight_ind.Text = "Disconnected"
                                   lbl_keysight_ind.BackColor = Color.DarkRed
                               End Sub)
                        If txt_alarm.Text = "ALARM : KEYSIGHT Connection Error please check the KEYSIGHT communication" And responseKSNo = 0 Then
                            Dim responseKS As DialogResult = MessageBox.Show("Do you want to reconnect to KEYSIGHT?", "KEYSIGHT COMMUNICATION ERROR", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
                            If responseKS = DialogResult.Yes Then
                                Try
                                    ConnectToKS(Config.IPSight, Config.PortSight)
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
                                End Try
                            Else
                                responseKSNo = 1
                            End If
                        End If
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

            .Reference = configData("Config")("Reference")
            .StartTime = configData("Config")("StartTime")
            .OperatorID = configData("Config")("OpeName")

            txt_cfg_ref.Text = .Reference
            txt_cfg_date.Text = .StartTime
            txt_cfg_id.Text = .OperatorID
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

            If .VolCtrl = "1" Then
                cb_cfg_use_v.Text = "Enable"
            Else
                cb_cfg_use_v.Text = "Disable"
            End If
            If .CurrCtrl Then
                cb_cfg_use_c.Text = "Enable"
            Else
                cb_cfg_use_c.Text = "Disable"
            End If
            If .TimeCtrl Then
                cb_cfg_use_time.Text = "Enable"
            Else
                cb_cfg_use_time.Text = "Disable"
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

    End Sub
    Private Sub LoadTable()
        Try
            Call DatabaseConnection.Connect()
            Dim sc As New SqlCommand("SELECT * FROM tb_datalog", DatabaseConnection.Connection)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        Finally
        End Try

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
    End Sub
    Private Sub btn_save_config_Click(sender As Object, e As EventArgs) Handles btn_save_config.Click
        If txt_cfg_ref.Text = "" Or txt_cfg_date.Text = "" Or txt_cfg_id.Text = "" Or txt_cfg_c_val.Text = "" Or txt_cfg_v_val.Text = "" Or txt_cfg_time_val.Text = "" Or txt_cfg_nbc.Text = "" Or cb_cfg_mode.Text = "" Or cb_cfg_use_v.Text = "" Or cb_cfg_use_c.Text = "" Or cb_cfg_use_time.Text = "" Or cb_cfg_use_cav1.Text = "" Or cb_cfg_use_cav2.Text = "" Or cb_cfg_use_cav3.Text = "" Or cb_cfg_use_cav4.Text = "" Or cb_cfg_use_cav5.Text = "" Or cb_cfg_use_cav6.Text = "" Then
            MsgBox("Please fill all data!")
            Exit Sub
        Else
            'With Config
            configData("Config")("Reference") = txt_cfg_ref.Text
            configData("Config")("StartTime") = txt_cfg_date.Text
            configData("Config")("OpeName") = txt_cfg_id.Text

            configData("Type")("CurrMode") = cb_cfg_mode.Text
            configData("Type")("ValVol") = txt_cfg_v_val.Text
            configData("Type")("ValCurr") = txt_cfg_c_val.Text
            configData("Type")("ValTime") = txt_cfg_time_val.Text
            configData("Type")("NbCyc") = txt_cfg_nbc.Text

            'With Type
            If cb_cfg_use_v.Text = "Enable" Then
                configData("Type")("UseVol") = "1"
            Else
                configData("Type")("UseVol") = "0"
            End If
            If cb_cfg_use_c.Text = "Enable" Then
                configData("Type")("UseCurr") = "1"
            Else
                configData("Type")("UseCurr") = "0"
            End If
            If cb_cfg_use_time.Text = "Enable" Then
                configData("Type")("UseTime") = "1"
            Else
                configData("Type")("UseTime") = "0"
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

    Private Sub btn_start_Click(sender As Object, e As EventArgs)

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

    Private Sub btn_search_Click(sender As Object, e As EventArgs) Handles btn_search.Click
        If current_mode_ac.CheckState = CheckState.Unchecked And current_mode_dc.CheckState = CheckState.Unchecked And use_vol_yes.CheckState = CheckState.Unchecked And use_vol_no.CheckState = CheckState.Unchecked And use_curr_yes.CheckState = CheckState.Unchecked And use_curr_no.CheckState = CheckState.Unchecked And use_time_yes.CheckState = CheckState.Unchecked And use_time_no.CheckState = CheckState.Unchecked Then
            Call DatabaseConnection.Connect()
            Dim sc As New SqlCommand("SELECT * FROM tb_datalog", DatabaseConnection.Connection)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Else
            Try
                Call DatabaseConnection.Connect()
                Dim sc As New SqlCommand("SELECT * FROM tb_datalog WHERE [Current Mode] = '" & Filter.CurrType & "' AND [Voltage Control] = '" & Filter.UseVol & "' AND [Current Control] = '" & Filter.UseCurr & "' AND [Time Control] = '" & Filter.UseTime & "'", DatabaseConnection.Connection)
                Dim adapter As New SqlDataAdapter(sc)
                Dim ds As New DataSet

                adapter.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
            Catch ex As Exception
                MessageBox.Show("Error loading data: " & ex.Message)
            Finally
            End Try
        End If
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
        CY101.ExtractCY = True
    End Sub

    Private Sub btn_v101_ret_Click(sender As Object, e As EventArgs) Handles btn_v101_ret.Click
        CY101.ReturnCY = True
    End Sub

    Private Sub btn_v102_ext_Click(sender As Object, e As EventArgs) Handles btn_v102_ext.Click
        CY102.ExtractCY = True
    End Sub

    Private Sub btn_v102_ret_Click(sender As Object, e As EventArgs) Handles btn_v102_ret.Click
        CY102.ReturnCY = True
    End Sub

    Private Sub btn_v103_ext_Click(sender As Object, e As EventArgs) Handles btn_v103_ext.Click
        CY103.ExtractCY = True
    End Sub

    Private Sub btn_v103_ret_Click(sender As Object, e As EventArgs) Handles btn_v103_ret.Click
        CY103.ReturnCY = True
    End Sub

    Private Sub btn_v104_ext_Click(sender As Object, e As EventArgs) Handles btn_v104_ext.Click
        CY104.ExtractCY = True
    End Sub

    Private Sub btn_v104_ret_Click(sender As Object, e As EventArgs) Handles btn_v104_ret.Click
        CY104.ReturnCY = True
    End Sub

    Private Sub btn_v105_ext_Click(sender As Object, e As EventArgs) Handles btn_v105_ext.Click
        CY105.ExtractCY = True
    End Sub

    Private Sub btn_v105_ret_Click(sender As Object, e As EventArgs) Handles btn_v105_ret.Click
        CY105.ReturnCY = True
    End Sub

    Private Sub btn_v106_ext_Click(sender As Object, e As EventArgs) Handles btn_v106_ext.Click
        CY106.ExtractCY = True
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
End Class