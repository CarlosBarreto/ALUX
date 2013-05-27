Imports System.Configuration
Imports System.Net.Configuration
Imports System.Net.Configuration.MailSettingsSectionGroup
Imports System.Net
Imports System.Net.Sockets

''' <summary>
''' 
''' </summary>
''' <remarks>
''' Usage: 
''' if (!SmtpHelper.TestConnection(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)))
''' {
'''       throw new ApplicationException("The smtp connection test failed"); 
''' }
''' </remarks>
Public Class smtpHelper

    ''' <summary>
    '''  Prueba la conexion SMTP mandando un comando HELO
    ''' </summary>
    ''' <param name="config"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TestConnection(ByVal config As Configuration) As Boolean
        Dim mailSettings As MailSettingsSectionGroup = config.GetSectionGroup("system.net/mailSettings") 'as MailSettingsSectionGroup
        If (mailSettings Is Nothing) Then
            Throw New ConfigurationErrorsException("La configuración del grupo system.net/mailSettings no pudo ser leido")
        End If
        Return TestConnection(mailSettings.Smtp.Network.Host, mailSettings.Smtp.Network.Port)
    End Function

    ''' <summary>
    '''  Prueba la conexion smtp enviando un commando HELO
    ''' </summary>
    ''' <param name="smtpServerAddress"></param>
    ''' <param name="Port"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TestConnection(ByVal smtpServerAddress As String, ByVal Port As Integer)
        Dim hostEntry As IPHostEntry = Dns.GetHostEntry(smtpServerAddress)
        Dim endPoint As EndPoint = New IPEndPoint(hostEntry.AddressList(0), Port)

        Using tcpSocket As New Socket(endPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp)
            'Tratar de conectarse y probar una respuesta para el codigo 220 = success
            Try

                tcpSocket.Connect(endPoint)
                If Not CheckResponse(tcpSocket, 220) Then
                    Return False
                End If

                'Enviar HELO y probar la respuesta para el codigo 250 = proper response
                'SendData(tcpSocket, String.Format("EHLO {0}\r\n", Dns.GetHostName()))
                SendData(tcpSocket, String.Format("EHLO {0}\r\n", smtpServerAddress))
                If Not CheckResponse(tcpSocket, 250) Then
                    Return False
                End If

            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Return False
            End Try

            'Si llega asta aqui, se pudo conectar al servidor SMTP
            Return True
        End Using

    End Function

    Private Sub SendData(ByVal socket As Socket, data As String)
        Dim dataArray As Byte() = Text.Encoding.ASCII.GetBytes(data)
        socket.Send(dataArray, 0, dataArray.Length, SocketFlags.None)
    End Sub

    Private Function CheckResponse(socket As Socket, expectedCode As Integer) As Boolean
        While socket.Available = 0
            System.Threading.Thread.Sleep(100)
        End While
        Dim responseArray(1024) As Byte
        socket.Receive(responseArray, 0, socket.Available, SocketFlags.None)
        Dim responseData As String = Text.Encoding.ASCII.GetString(responseArray)
        Dim responseCode As Integer = Convert.ToInt32(responseData.Substring(0, 3))
        If responseCode = expectedCode Then
            Return True
        End If
        Return False
    End Function

End Class
