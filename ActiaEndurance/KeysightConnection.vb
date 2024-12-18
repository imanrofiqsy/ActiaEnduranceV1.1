Imports Microsoft.VisualBasic
Imports System.Text
Imports System.Net.Sockets
Imports System.Threading
Module KeysightConnection

    'Init Variable
    Dim client As TcpClient
    Dim _connected As Boolean
    Dim pingTimeout As Integer
    Dim tcpServer As Socket
    Dim tcpListeners As TcpListener
    Dim dataRX(1024) As Byte
    Dim TCPClinetStream As NetworkStream
    Dim KeyKS As String
    Public KSconnected As Boolean

    Public Sub ConnectToKS(ip As String, port As String)
rePingKS:
        Try
            client = New TcpClient(ip, port)
            WriteToKS("*IDN?")
            'Thread.Sleep(1000)
            For i As Integer = 1 To 10
                Thread.Sleep(100)
                My.Application.DoEvents()
            Next

            KeyKS = ReadFromKS()
            If client.Connected And Strings.Replace(KeyKS, Environment.NewLine, "") = "Keysight Technologies,DAQ973A,MY59012524,A.03.02-01.00-02.01-00.02-02.00-03-03" Then
                _connected = True
                KSconnected = True
            Else
                If pingTimeout < 5 Then
                    pingTimeout = pingTimeout + 1
                    GoTo rePingKS
                End If
            End If
        Catch ex As Exception
            If pingTimeout < 5 Then
                pingTimeout = pingTimeout + 1
                GoTo rePingKS
            End If
            KSconnected = False
        End Try

    End Sub

    Public Function Connected() As Boolean
        Return _connected
    End Function

    Public Sub DisconnectFromKS()
        client.Client.Close()
        If client.Connected = False Then
            _connected = False
            KSconnected = False
        End If
    End Sub

    Public Sub WriteToKS(txtData As String)
        Dim dataTX As Byte()

        Try
            If String.IsNullOrEmpty(txtData) Then
                Return
            End If

            'dataTX = Encoding.ASCII.GetBytes(txtData & vbCrLf)
            dataTX = Encoding.ASCII.GetBytes(txtData + Environment.NewLine)

            If client IsNot Nothing And client.Client.Connected Then
                client.GetStream.Write(dataTX, 0, dataTX.Length)
            End If
        Catch ex As Exception
            KSconnected = False
        End Try
    End Sub
    Dim strReceive As String
    Public Function ReadFromKS() As String

        Try
            TCPClinetStream = client.GetStream()
            While (TCPClinetStream.DataAvailable)
                TCPClinetStream.Read(dataRX, 0, dataRX.Length)
                strReceive = Encoding.ASCII.GetString(dataRX)
                strReceive = Strings.Replace(strReceive, vbLf, "").Replace(vbNullChar, "").Replace(Environment.NewLine, "")
            End While
        Catch ex As Exception
            KSconnected = False
            strReceive = ""
        End Try

        Return strReceive
    End Function
End Module
