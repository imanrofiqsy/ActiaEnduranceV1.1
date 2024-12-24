Module GlobalVariable
    Public Config As MachineConfig
    Public Type As ProductType
    Public UseCavity As Cavity
    Public EnCavity As Cavity
    Public FDUT As Cavity
    Public CurrentVal As Cavity
    Public Filter As FilterType
    Public Alarm As MachineAlarm
    Public Manual As ManualVar
    Public GeneralComm As GeneralCommunication
    Public Logging As ProductionData
    Public UserLevel As Integer
    Public StateNumber As Integer
    Public measResult As String
    Public trigPLC As Boolean
    Public measDone As Boolean

    Public Structure MachineConfig
        'Modbus
        Dim IP As String
        Dim Port As String

        'KeySight
        Dim IPSight As String
        Dim PortSight As String

        'Keithley
        Dim IPKeith As String
        Dim PortKeith As String

        'Product
        Dim Reference_1 As String
        Dim Reference_2 As String
        Dim Reference_3 As String
        Dim StartTime As String
        Dim OperatorID As String

    End Structure

    Public Structure MachineAlarm

        Dim alarmMsg As String

        ' Minor Defect (Cylinder)
        Dim MinorEXT As Integer

        Dim V101EXT As Integer
        Dim V102EXT As Integer
        Dim V103EXT As Integer
        Dim V104EXT As Integer
        Dim V105EXT As Integer
        Dim V106EXT As Integer

        Dim MinorRET As Integer

        Dim V101RET As Integer
        Dim V102RET As Integer
        Dim V103RET As Integer
        Dim V104RET As Integer
        Dim V105RET As Integer
        Dim V106RET As Integer

        'Major Defect
        Dim Major As Integer

        Dim AirFail As Integer
        Dim EmgPress As Integer
        Dim ContactorFail As Integer

        'Alarm Defect
        Dim General As Integer

        Dim MeasTO As Integer
        Dim NeedInit As Integer
        Dim NeedConf As Integer
    End Structure

    Public Structure ProductType
        Dim CurrMode As String
        Dim VolCtrl As String
        Dim CurrCtrl As String
        Dim TimeCtrl As String
        Dim VoltVal As String
        Dim CurrVal As String
        Dim TimeVal As String
        Dim NbCyc As String
    End Structure

    Public Structure Cavity
        Dim Cav1 As String
        Dim Cav2 As String
        Dim Cav3 As String
        Dim Cav4 As String
        Dim Cav5 As String
        Dim Cav6 As String

        Dim Cur1 As String
        Dim Cur2 As String
        Dim Cur3 As String
        Dim Cur4 As String
        Dim Cur5 As String
        Dim Cur6 As String
    End Structure

    Public Structure FilterType
        Dim CurrType As String
        Dim UseVol As String
        Dim UseCurr As String
        Dim UseTime As String
    End Structure

    'Manual Var Start
    Public CY101 As ManualVar
    Public CY102 As ManualVar
    Public CY103 As ManualVar
    Public CY104 As ManualVar
    Public CY105 As ManualVar
    Public CY106 As ManualVar

    Public Structure ManualVar
        'Move Cylinder
        Dim ExtrudeCY As Boolean
        Dim ReturnCY As Boolean

        'Cylinde Sensor
        Dim MaxCY As Boolean
        Dim MinCY As Boolean

        'Manual Var End
    End Structure

    Public Structure GeneralCommunication
        'PC To PLC
        Dim EnVol As Integer
        Dim EnCurr As Integer
        Dim EnTime As Integer

        Dim rdyStart As Integer
        Dim FinVol As Integer
        Dim FinCurr As Integer

        'PLC To PC
        Dim trigVol As Integer
        Dim trigSaveCurr As Integer

    End Structure

    Public Structure ProductionData
        Dim No As Integer
        Dim Reference_1 As String
        Dim Reference_2 As String
        Dim Reference_3 As String
        Dim StartDate As String
        Dim OpeId As String
        Dim Cavity As Integer
        Dim CycleCount As Integer
        Dim CurrentMode As String
        Dim UseVoltage As String
        Dim Voltage As String
        Dim UseCurrent As String
        Dim Current As String
        Dim UseTime As String
        Dim Time As String
        Dim FinalResult As String
    End Structure
End Module
