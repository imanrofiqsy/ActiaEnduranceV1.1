Imports EasyModbus
Module Modbus
    Dim modbusClient As ModbusClient
    Dim ConnectionError As Boolean
    Public IsConnected As Boolean

    Public Function ModbusConnected() As Boolean
        Return IsConnected
    End Function
    Public Sub OpenPort(IP As String, PORT As String)
        Try
            modbusClient = New ModbusClient(IP, Val(PORT))
            modbusClient.Connect()
            IsConnected = True
        Catch ex As Exception
            IsConnected = False
        End Try
    End Sub
    Public Sub ClosePort()
        modbusClient.Disconnect()
        IsConnected = False
    End Sub
    Public Function _ConnectionError() As Boolean
        Return ConnectionError
    End Function
    Public Function WriteBit(val As Integer, bit As Integer, sett As Integer) As Integer
        Dim binaryString As String = Convert.ToString(val, 2).PadLeft(16, "0"c)
        Dim binaryCharArray As Char() = binaryString.ToCharArray()
        Array.Reverse(binaryCharArray)
        If sett = 1 Then
            binaryCharArray(bit) = "1"
        Else
            binaryCharArray(bit) = "0"
        End If
        Array.Reverse(binaryCharArray)
        Dim modifiedBinaryString As New String(binaryCharArray)
        Dim integer_val As Integer = Convert.ToInt16(modifiedBinaryString, 2)
        Return integer_val
    End Function
    Public Function ReadBit(val As Integer, bit As Integer) As Integer
        Dim binaryString As String = Convert.ToString(val, 2).PadLeft(16, "0"c)
        binaryString = binaryString.Reverse.ToArray
        Return CInt(Char.GetNumericValue(binaryString(bit)))
    End Function
    Public Function ReadDoubleInteger(addr As Integer) As Integer
        Dim address_val() As Integer
        Dim result As Integer

        address_val = modbusClient.ReadHoldingRegisters(addr, 2)
        result = (CInt(address_val(0)) And &HFFFF) Or (CInt(address_val(1)) << 16)
        Return result
    End Function
    Public Sub WriteDoubleInteger(addr As Integer, val As Int32)
        Dim values(1) As Int32
        values(0) = CInt((val >> 16) And &HFFFF)
        values(1) = CInt(val And &HFFFF)
        Dim result(1) As Integer
        result(0) = values(1)
        result(1) = values(0)
        modbusClient.WriteMultipleRegisters(addr, result)
    End Sub
    Public Function ReadFloat(startAddress As Integer) As Double
        Dim vals() As Integer = modbusClient.ReadHoldingRegisters(startAddress, 2)
        'Console.WriteLine(vals(0).ToString + " " + vals(1).ToString)
        Return ConvertDW2Float(CShort(vals(0)), CShort(vals(1)))
    End Function
    Private Function ConvertDW2Float(reg1 As Short, reg2 As Short) As Single
        Dim bytes(3) As Byte
        Dim intBytes1 As Byte() = BitConverter.GetBytes(reg2)
        Dim intBytes2 As Byte() = BitConverter.GetBytes(reg1)

        ' Assuming Float Swap format (swapping the two 16-bit registers)
        bytes(0) = intBytes2(0)
        bytes(1) = intBytes2(1)
        bytes(2) = intBytes1(0)
        bytes(3) = intBytes1(1)

        Dim result As Single = BitConverter.ToSingle(bytes, 0)
        'Console.WriteLine(result)
        Return result
    End Function
    Public Sub WriteFloat(StartAddress As Integer, Value As Single)
        Dim values() As Integer = ConvertFloat2DW(Value)
        modbusClient.WriteMultipleRegisters(StartAddress, values)
    End Sub

    Private Function ConvertFloat2DW(fval As Single) As Integer()
        Dim byVals As Byte() = BitConverter.GetBytes(fval)

        ' Swap the byte order for "Float Swap" format
        Dim bytes(3) As Byte
        bytes(0) = byVals(2)
        bytes(1) = byVals(3)
        bytes(2) = byVals(0)
        bytes(3) = byVals(1)

        ' Convert to Short array for Modbus WriteMultipleRegisters
        Dim values(1) As Short
        values(0) = BitConverter.ToInt16(bytes, 0)
        values(1) = BitConverter.ToInt16(bytes, 2)

        ' Convert to Integer array for Modbus WriteMultipleRegisters
        Dim result(1) As Integer
        result(0) = values(1)
        result(1) = values(0)

        Return result
    End Function
    Public Function ReadInteger(address As Integer) As Integer
        Try
            Dim val() As Integer = modbusClient.ReadHoldingRegisters(address, 1)
            Return val(0)
        Catch ex As Exception
            IsConnected = False
            Return 0
        End Try
    End Function
    Public Sub WriteInteger(address As Integer, value As Integer)
        modbusClient.WriteSingleRegister(address, value)
    End Sub
End Module

