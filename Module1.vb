Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail

Module Module1
    Public mail As MailMessage
    Public smtp As SmtpClient

    Public DA As DataAccess
    Public DR As SqlDataReader
    Public Read As SqlDataReader

    Sub Main()
        DA = New DataAccess(__SERVER__, __DATABASE__, __USER__, __PASS__)
        Dim mailArray(100) As Mailing
        Dim cont As Integer = 0

        Try
            'Guardar el correo en la Base de datos
            DR = DA.ExecuteSP("sys_getMailErrorList") '
            If DA._LastErrorMessage <> "" Then
                Throw New Exception(DA._LastErrorMessage)
            Else

                While DR.Read
                    mailArray(cont) = New Mailing With {._MailID_ = DR(0), ._To_ = DR(1), ._From_ = DR(2), ._Subject_ = DR(3), ._Body_ = DR(4)}
                    If cont + 1 >= mailArray.Length Then ReDim Preserve mailArray(cont + 10)
                    cont += 1
                End While

                ReDim Preserve mailArray(cont - 1)
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            End
        End Try

        If Not DR.IsClosed Then
            DR.Close()
        End If

        For Each m As Mailing In mailArray
            Try
                mail = New MailMessage()
                mail.To.Add(m._To_)
                mail.From = New MailAddress(m._From_)
                mail.Subject = m._Subject_
                mail.Body = m._Body_
                mail.IsBodyHtml = True
                smtp = New SmtpClient()
                smtp.Host = "mail.agiotech.com"
                smtp.Port = 587
                smtp.Credentials = New NetworkCredential("carlos.barreto@agiotech.com", "Agiotech01")
                smtp.Send(mail)
                ' Si no ocurre error, cambiar el status de new a DONE
                Read = DA.ExecuteSP("sys_mailStatusCh", m._MailID_, "DONE")
            Catch ex As Exception
                'En caso de error, cambiar el estatus
                If m._MailID_ <> "" Then
                    Read = DA.ExecuteSP("sys_mailStatusCh", m._MailID_, "FAIL")
                End If
            Finally
                If Not Read.IsClosed Then
                    Read.Close()
                End If
            End Try
        Next

        DA.Dispose()

        System.Threading.Thread.Sleep(5000)
    End Sub

    Public Structure Mailing
        Public _MailID_ As String
        Public _To_ As String
        Public _From_ As String
        Public _Subject_ As String
        Public _Body_ As String
    End Structure

    Public Function ReenviarCorreo() As Boolean
        'Enviar un correo de prueba
        mail = New MailMessage
        mail.To.Add("carlosbarreto@outlook.com")
        mail.From = New MailAddress("carlos.barreto@agiotech.com")
        mail.Subject = "Prueba"
        mail.Body = "Esta es una prueba"
        mail.IsBodyHtml = False

        Dim altSmtp As New SmtpClient("smtp.gmail.com", 587)
        altSmtp.EnableSsl = True
        Dim mycred As NetworkCredential = New NetworkCredential("sys.agiotech@gmail.com", "Agiotech01")
        altSmtp.Credentials = mycred

        Try
            altSmtp.Send(mail)
            'smtp = New SmtpClient()
            'smtp.Send(mail)
            'Catch exc As SmtpFailedRecipientsException
            '   For i As Integer = 0 To exc.InnerExceptions.Length Step 1
            'en caso de error, utilizar un cliente externo
            '
            'Next
        Catch ex As SmtpException
            Console.WriteLine(ex.Message)
        End Try
    End Function
End Module

