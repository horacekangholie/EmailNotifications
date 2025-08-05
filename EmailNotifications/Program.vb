Imports System.Net.Mail
Imports System.Data.SqlClient
Imports System.Data
Imports System.Security.Cryptography
Imports System.Runtime
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Runtime.CompilerServices

Module Program
    ' Database Connection string
    Dim connectionString As String = "Server=localhost;Database=lmsportal;User ID=sa;Password=qweasdzxc1@3;Pooling=true;Trusted_Connection=false;"               ' Live server
    'Dim connectionString As String = "Server=localhost\SQLEXPRESS;Database=lmsportal;User ID=sa;Password=pass@1234;Pooling=true;Trusted_Connection=false;"        ' Test server

    ' Program Load
    Sub Main(args As String())
        GenerateMonthlyEmail("DMC Account Reminder")
        GenerateMonthlyEmail("AI Licence Reminder")
        GenerateMonthlyEmail("Termed Licence Reminder")
        'GenerateMonthlyEmail("AI Licence Billing Notifications")
    End Sub

    Sub GenerateMonthlyEmail(ByVal emailType As String)
        Dim query As String = GetSQL(emailType, Nothing)  ' Retrieve a list of recipients

        ' Generate email for each of recipient from retrieved list
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.HasRows Then
                        While reader.Read()
                            ' Email attributes
                            Dim contentTypes As String() = GetContentTypes(emailType)
                            Dim recipientID As String = reader("Recipient ID").ToString()
                            Dim recipientName As String = reader("Recipient Name").ToString()
                            Dim recipientEmail As String = reader("Recipient Email").ToString()
                            Dim ccEmail As String = reader("Cc_Email").ToString()
                            Dim bccEmail As String = reader("Bcc_Email").ToString()
                            Dim subject As String = GetEmailSubject(emailType)

                            ' Generate the body content by calling CreateEmailBody function
                            Dim body As String = CreateEmailBody(recipientName, contentTypes, recipientID, emailType)

                            ' Execute sending email if body is successfully generated
                            If Not String.IsNullOrEmpty(body) Then
                                SendEmail(recipientEmail, subject, body, ccEmail, bccEmail)
                            End If
                        End While
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub SendEmail(ByVal recipient As String, ByVal subject As String, ByVal body As String, Optional ByVal cc As String = Nothing, Optional ByVal bcc As String = Nothing)
        Try
            ' Fetch SMTP configuration from the database
            Dim smtpConfig As DataRow = GetSmtpConfigFromDb()

            If smtpConfig Is Nothing Then
                Console.WriteLine("Failed to fetch SMTP configuration from the database.")
                Return
            End If

            Dim mail As New MailMessage()
            mail.From = New MailAddress(smtpConfig("Username").ToString(), "DMC Administrator")   ' Set sender email

            ' Set recipients email
            mail.To.Add(recipient)

            ' Add CC if provided and not empty
            If Not String.IsNullOrWhiteSpace(cc) Then
                mail.CC.Add(cc)
            End If

            ' Add BCC if provided and not empty
            If Not String.IsNullOrWhiteSpace(bcc) Then
                mail.Bcc.Add(bcc)
            End If

            ' Set email subject and body
            mail.Subject = subject
            mail.Body = body
            mail.IsBodyHtml = True ' Set to False if not sending HTML email

            ' Set SMTP client using the details from the database
            Dim smtp As New SmtpClient(smtpConfig("Host").ToString())
            smtp.Credentials = New System.Net.NetworkCredential(smtpConfig("Username").ToString(), smtpConfig("Password").ToString())
            smtp.Port = Convert.ToInt32(smtpConfig("Port"))
            smtp.EnableSsl = Convert.ToBoolean(smtpConfig("SSL_Enabled"))

            ' Send email
            smtp.Send(mail)

            Console.WriteLine("Email sent successfully.")
        Catch ex As Exception
            Console.WriteLine("Error sending email: " & ex.Message)
        End Try
    End Sub



    ' Common Functions
    Private Function GetSmtpConfigFromDb() As DataRow
        Dim smtpConfig As DataTable = New DataTable()
        Using connection As New SqlConnection(connectionString)
            Try
                connection.Open()    ' Open the connection

                ' SQL query to fetch the valid SMTP configuration
                Dim query As String = "SELECT TOP 1 * FROM DB_Smtp WHERE Is_Valid = 1"

                ' Create a data adapter
                Using adapter As New SqlDataAdapter(query, connection)
                    ' Fill the DataTable with the results of the query
                    adapter.Fill(smtpConfig)
                End Using

            Catch ex As Exception
                Console.WriteLine("Error fetching SMTP configuration: " & ex.Message)
                Return Nothing
            Finally
                ' Close the connection
                connection.Close()
            End Try
        End Using

        ' Return the first row of the DataTable if available
        If smtpConfig.Rows.Count > 0 Then
            Return smtpConfig.Rows(0)
        End If

        Return Nothing
    End Function

    Private Function GetSQL(ByVal notificationType As String, Optional ByVal filter As String = Nothing) As String
        Dim query As String = Nothing

        Select Case notificationType
            Case "DMC Account Reminder"
                query = "SELECT [Recipient ID], [Recipient Name], [Recipient Email], [Cc_Email], [Bcc_Email] " &
                        "FROM _DMC_Reminder_Notifications_Email_List "

            Case "AI Licence Reminder"
                query = "SELECT [Recipient ID], [Recipient Name], [Recipient Email], [Cc_Email], [Bcc_Email] " &
                        "FROM _AI_Licence_Notifications_Email_List "

            Case "Termed Licence Reminder"
                query = "SELECT [Recipient ID], [Recipient Name], [Recipient Email], [Cc_Email], [Bcc_Email] " &
                        "FROM _Termed_Licence_Notifications_Email_List "

            Case "AI Licence Billing Notifications"
                query = "SELECT [Recipient ID], [Recipient Name], [Recipient Email], [Cc_Email], [Bcc_Email] " &
                        "FROM _AI_Licence_Billing_Notifications_Email_List "

            Case "DMC Billed Account Expiry"
                query = "SELECT [Bill Entity], [Group], [HQ Code], [HQ Name], [Store Code], [Store Name], [Created Date] " &
                        "     , [Start Date], [End Date], [Duration], [Currency], [Fee], [Status], [Account Type], [Sales Representative] " &
                        "FROM D_DMC_Billed_Account_Expired_In_2_Months " &
                        "WHERE [Sales Representative ID] = '" & filter & "' " &
                        "ORDER BY [End Date], [Bill entity], [HQ Code], [Store Code] "

            Case "DMC Trial Account Expiry"
                query = "SELECT [Customer], [Group], [HQ Code], [HQ Name], [Store Code], [Store Name], [Created Date], [End Date], [Status], [Account Type], [Sales Representative] AS [Requestor] " &
                        "FROM D_DMC_Trial_Account_Expired_In_2_Months " &
                        "WHERE [Sales Representative ID] = '" & filter & "' " &
                        "ORDER BY [End Date], [Customer], [Store Code] "

            Case "DMC Suspended Store"
                query = "SELECT [Headquarter ID], [Headquarter Name], [Store No], [Store Name], [Account Type], [Created Date], [Expiry Date], [Suspended Date], [Status], [Requestor] AS [Sales Representative], [Reason of Suspension] " &
                        "FROM R_Suspended_Stores " &
                        "WHERE [Suspended Date] = DATEADD(DAY, 1, EOMONTH(GETDATE(), -1)) " &
                        "  AND [Sales Representative ID] = '" & filter & "' " &
                        "ORDER BY [Suspended Date] DESC, [Headquarter ID] "

            Case "AI Licence (Expiring)"
                query = "SELECT [Licensee], [Application Type], [Serial No], [AI Device ID], [AI Device Serial No], [Activated Date], [Expired Date], [Licence Code] AS [Binding Key], [MAC Address] " &
                        "     , CASE WHEN CAST([Licence Term] AS int) > 1000 THEN 'No Expiry' ELSE CAST([Licence Term] as nvarchar) END AS [Licence Term] " &
                        "     , [Created Date], [Status], [Requested By] " &
                        "FROM D_Licence_With_Term  " &
                        "WHERE [Expired Date] <= DATEADD (dd, -1, DATEADD(mm, DATEDIFF(mm, 0, GETDATE()) + 10, 0))  " &
                        "  AND [Application Type] IN ('PC Scale (AI Classic)', 'PC Scale - AI (Online)', 'PC Scale - AI (Offline)')  " &
                        "  AND [Status] NOT IN ('Renew', 'Blocked', 'Expired')  " &
                        "  AND Replace([Licence Code], '-', '') NOT IN (SELECT Replace(Value_1, '-', '') FROM DB_Lookup WHERE Lookup_Name = 'Production Used Licence Key') " &
                        "  AND [Requestor ID] = '" & filter & "' " &
                        "ORDER BY [Expired Date], [Serial No] "

            Case "AI Licence (Renewed)"
                ' For license with Renew status, list them all regardless when its expiry date
                query = "SELECT [Licensee], ISNULL([Application Type] + ' (' + Activated_Module_Type + ') ', [Application Type]) AS [Application Type] " &
                        "     , [Serial No], [AI Device ID], [AI Device Serial No], [Activated Date], [Expired Date], [Licence Code] AS [Binding Key], [MAC Address] " &
                        "     , CASE WHEN CAST([Licence Term] AS int) > 1000 THEN 'No Expiry' ELSE CAST([Licence Term] as nvarchar) END AS [Licence Term] " &
                        "     , [Created Date], [Status], [Requested By] " &
                        "FROM R_LMS_Module_Licence " &
                        "LEFT JOIN LMS_Module_Licence_Activated ON LMS_Module_Licence_Activated.[Licence_Code] = REPLACE(R_LMS_Module_Licence.[Licence Code], '-', '') " &
                        "WHERE [status] IN ('Renew')  " &
                        "  AND [Requestor ID] = '" & filter & "' " &
                        "ORDER BY CAST([Expired Date] AS date) DESC "

            Case "AI Licence (Expired)"
                query = "SELECT [Licensee], [Application Type], [Serial No], [AI Device ID], [AI Device Serial No], [Activated Date], [Expired Date], [Licence Code] AS [Binding Key], [MAC Address] " &
                        "     , CASE WHEN CAST([Licence Term] AS int) > 1000 THEN 'No Expiry' ELSE CAST([Licence Term] as nvarchar) END AS [Licence Term] " &
                        "     , [Created Date], [Status], [Requested By] " &
                        "FROM D_Licence_With_Term " &
                        "WHERE [Expired Date] <= DATEADD (dd, -1, DATEADD(mm, DATEDIFF(mm, 0, GETDATE()) + 10, 0)) " &
                        "  AND [Application Type] IN ('PC Scale (AI Classic)', 'PC Scale - AI (Online)', 'PC Scale - AI (Offline)') AND [Status] IN ('Expired') " &
                        "  AND Replace([Licence Code], '-', '') NOT IN (SELECT Replace(Value_1, '-', '') FROM DB_Lookup WHERE Lookup_Name = 'Production Used Licence Key') " &
                        "  AND [Requestor ID] = '" & filter & "' " &
                        "ORDER BY [Expired Date], [Serial No] "

            Case "Termed Licence (Expiring)"
                query = "SELECT [Licensee], [Application Type] " &
                        "     , ISNULL([Serial No], '-') AS [Serial No] " &
                        "     , ISNULL([AI Device ID], '-') AS [AI Device ID] " &
                        "     , ISNULL([AI Device Serial No], '-') AS [AI Device Serial No] " &
                        "     , [Activated Date], [Expired Date], [Licence Code] AS [Binding Key], [MAC Address] " &
                        "     , CASE WHEN CAST([Licence Term] AS int) > 1000 THEN 'No Expiry' ELSE CAST([Licence Term] as nvarchar) END AS [Licence Term] " &
                        "     , [Created Date], [Status], [Requested By] " &
                        "FROM D_Licence_With_Term " &
                        "WHERE [Expired Date] BETWEEN DATEADD(mm, DATEDIFF(mm, 0, GETDATE()) - 12, 0) AND DATEADD (dd, -1, DATEADD(mm, DATEDIFF(mm, 0, GETDATE()) + 3, 0)) " &
                        "  AND [Application Type] NOT IN ('PC Scale (AI)') AND Chargeable NOT IN ('No') " &
                        "ORDER BY [Expired Date] DESC "

            Case "AI Licence Billing List"
                query = "SELECT [Code], [Distributor], [Customer], [Store] " &
                        "     , [Licence Key], [MAC Address] " &
                        "     , [Is Trial], [CZL Account], [Account Model] AS [Model] " &
                        "     , [Scale SN], [AI Activation Key], [Device Serial], [Device ID] " &
                        "     , [Mode], [Term In Month] AS [Term] " &
                        "     , [Created Date], [Registered Date] " &
                        "     , L.PO_No AS [PO No] " &
                        "     , IPS.[SO No] " &
                        "     , IPS.[Invoice No] " &
                        "     , CASE WHEN [Mode] = 'Online'  " &
                        "            THEN CASE WHEN [Renewed Date] > [Registered Date]  " &
                        "                      THEN [Renewed Date] " &
                        "                      ELSE [Registered Date] END " &
                        "            ELSE CASE WHEN [Registered Date] < DATEADD(YEAR, DATEDIFF(YEAR, [Registered Date], DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)), [Registered Date]) " &
                        "            THEN DATEADD(YEAR, DATEDIFF(YEAR, [Registered Date], DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)), [Registered Date])  " &
                        "            ELSE [Registered Date] END  " &
                        "            END AS [Registered / Renew Date]  " &
                        "     , Case When [Mode] = 'Online' " &
                        "            THEN CASE WHEN DATEDIFF(YEAR, [Registered Date], [Renewed Date]) > 0  " &
                        "                      THEN DATEDIFF(YEAR, [Registered Date], [Renewed Date]) + 2 " &
                        "                      ELSE 1 END " &
                        "            ELSE DATEDIFF(YEAR, [Registered Date], DATEADD(YEAR, DATEDIFF(YEAR, [Registered Date], DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)), [Registered Date])) + 1  " &
                        "            END AS [Bill Cycle] " &
                        "     , [Sales Rep] " &
                        "FROM R_DMC_CZL_Biling_Report R " &
                        "INNER JOIN LMS_Licence L ON L.Licence_Code = R.[Licence Key] " &
                        "INNER JOIN _Customer_Invoice_PO_SO IPS ON IPS.[Customer ID] = L.Customer_ID AND IPS.[PO No] = L.PO_No " &
                        "ORDER BY [Mode], [Distributor], [Customer], [Store], [Scale SN] "

        End Select

        Return query
    End Function

    Private Function GetEmailSubject(ByVal emailType As String) As String
        Dim subject As String = Nothing
        Select Case emailType
            Case "DMC Account Reminder"
                subject = String.Format("Upcoming DMC Account Expiry [{0} - {1}]", DateTime.Today.ToString("MMM yyyy"), DateTime.Today.AddMonths(2).ToString("MMM yyyy"))
            Case "AI Licence Reminder"
                subject = String.Format("Upcoming AI Licence expiry [{0} - {1}]", DateTime.Today.ToString("MMM yyyy"), DateTime.Today.AddMonths(2).ToString("MMM yyyy"))
            Case "Termed Licence Reminder"
                subject = String.Format("Upcoming Termed Licence expiry [{0} - {1}]", DateTime.Today.ToString("MMM yyyy"), DateTime.Today.AddMonths(2).ToString("MMM yyyy"))
            Case "AI Licence Billing Notifications"
                subject = String.Format("Activated AI Licence [{0}]", DateTime.Today.AddMonths(-1).ToString("MMM yyyy"))
        End Select
        Return subject
    End Function

    Private Function GetContentHeading(ByVal headerName As String) As String
        Dim heading As String = Nothing
        Select Case headerName
            Case "DMC Billed Account Expiry"
                heading = String.Format("<h3 Class='reportTitle'>DMC Billed Account expired as of {0}</h3>", New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(3).AddDays(-1).ToString("dd MMM yyyy"))
            Case "DMC Trial Account Expiry"
                heading = String.Format("<h3 class='reportTitle'>DMC Trial Account expired as of {0}</h3>", New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(3).AddDays(-1).ToString("dd MMM yyyy"))
            Case "DMC Suspended Store"
                heading = String.Format("<h3 class='reportTitle'>Suspended stores on {0}</h3>", New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).ToString("dd MMM yyyy"))
            Case "AI Licence (Expiring)"
                heading = "<h3 class='reportTitle'>AI Licence (Expiring)</h3><div class='alert'><div class='noteText'>Note: Following is/are expiring AI license(s), please attach SAS order for renewal.</div></div>"
            Case "AI Licence (Renewed)"
                heading = "<h3 class='reportTitle'>AI Licence (Renewed)</h3><div class='alert'><div class='noteText'>Note: Following license(s) is/are in 'Renew' status, please advise user to perform re-authentication to renew the license.</div></div>"
            Case "AI Licence (Expired)"
                heading = "<h3 class='reportTitle'>AI Licence (Expired)</h3><div class='alert'><div class='noteText'>Note: Following is/are expired AI license(s). Please advise if user wish to renew the license.</div></div>"
            Case "Termed Licence (Expiring)"
                heading = "<h3 class='reportTitle'>Termed Licence (Expiring)</h3><div class='alert'><div class='noteText'>Note: Following is/are expiring Termed license(s), please send SAS order for renewal.</div></div>"
        End Select
        Return heading
    End Function

    Private Function GetContentTypes(ByVal emailType As String) As String()
        Dim contentTypes As String() = Nothing
        Select Case emailType
            Case "DMC Account Reminder"
                contentTypes = {"DMC Billed Account Expiry", "DMC Trial Account Expiry", "DMC Suspended Store"}

            Case "AI Licence Reminder"
                contentTypes = {"AI Licence (Expiring)", "AI Licence (Renewed)", "AI Licence (Expired)"}

            Case "Termed Licence Reminder"
                contentTypes = {"Termed Licence (Expiring)"}

            Case "AI Licence Billing Notifications"
                contentTypes = {"AI Licence Billing List"}

        End Select
        Return contentTypes
    End Function

    Private Function CreateEmailBody(ByVal recipientName As String, ByVal contentTypes As String(), ByVal recipientID As String, ByVal emailType As String) As String
        Dim bodyBuilder As New System.Text.StringBuilder()
        ' Email Styling
        bodyBuilder.AppendLine("<html>")
        bodyBuilder.AppendLine("<head>")
        bodyBuilder.AppendLine("<style>")
        bodyBuilder.AppendLine(".primary-table { width: 100%; border: 1px solid #ccc; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; }")
        bodyBuilder.AppendLine(".primary-table th, .primary-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }")
        bodyBuilder.AppendLine(".primary-table th { background-color: #f2f2f2; font-weight: bold; color: #333; }")
        bodyBuilder.AppendLine(".primary-table tr:nth-child(even) { background-color: #f9f9f9; }")
        bodyBuilder.AppendLine(".primary-table tr:hover { background-color: #f1f1f1; }")
        bodyBuilder.AppendLine(".primary-table th { padding-top: 12px; padding-bottom: 12px; background-color: #b8daff; color: black; text-align: left; }")
        bodyBuilder.AppendLine(".primary-table td { text-align: left; }")
        bodyBuilder.AppendLine(".nowrap { white-space: nowrap; }")
        bodyBuilder.AppendLine(".emailGreeting { margin-bottom: 15px; }")
        bodyBuilder.AppendLine(".reportTitle { margin-top: 50px; }")
        bodyBuilder.AppendLine(".signature-text-style { font-family:'Aptos', sans-serif; font-size: 11pt; color: #9D9D9D; font-weight: normal; margin-top: 150px; }")
        bodyBuilder.AppendLine(".signature-text-style .designation { font-size: 11pt; margin-bottom: 60px; }")
        bodyBuilder.AppendLine(".signature-text-style .companyName { font-size: 16pt; }")
        bodyBuilder.AppendLine(".signature-text-style .departmentName  { font-size: 10pt; }")
        bodyBuilder.AppendLine(".signature-text-style .companyAddrress { font-size: 9pt; }")
        bodyBuilder.AppendLine(".alert { box-sizing: border-box; background-color: rgb(246, 227, 204); border: 1px solid rgb(238, 200, 153); border-radius: 5px; color: rgb(85, 47, 0); margin-bottom: 16px; position: relative; font-family: Arial, sans-serif; font-size: 11pt; font-weight: 300; line-height: 12px; text-align: start; text-size-adjust: 100%; } ")
        bodyBuilder.AppendLine(".alert .noteText { padding: 50px !important; } ")
        bodyBuilder.AppendLine("</style>")
        bodyBuilder.AppendLine("</head>")

        ' Add email greeting
        bodyBuilder.AppendLine("<body>")
        bodyBuilder.AppendLine($"<div class='emailGreeting'>Dear {recipientName},</div>")
        Select Case emailType
            Case "DMC Account Reminder"
                bodyBuilder.AppendLine("<div class='emailGreeting'>It’s the beginning of the month, kindly observe the account expiry date.<br>")
                bodyBuilder.AppendLine("Please arrange early for subscription renewal/extension before the account expired.</div>")

            Case "AI Licence Reminder"
                bodyBuilder.AppendLine("<div>Please observe the status of the following AI Licences.</div>")

            Case "Termed Licence Reminder"
                bodyBuilder.AppendLine("<div>Please observe the status of the following Termed Licences.</div>")

            Case "AI Licence Billing Notifications"
                bodyBuilder.AppendLine("<div style='margin-bottom:40px'>Following is the list of activated AI Licence.</div>")

        End Select


        ' Loop through notification types to build the email body
        bodyBuilder.AppendLine("<div>")
        For Each headerName As String In contentTypes
            Dim partialBody As String = contentStringBuilder(headerName, recipientID)
            If partialBody IsNot Nothing Then
                Dim contentTitle As String = GetContentHeading(headerName)
                bodyBuilder.AppendLine(contentTitle)
                bodyBuilder.AppendLine(partialBody)
            End If
        Next
        bodyBuilder.AppendLine("</div>")

        ' Email Signature
        bodyBuilder.AppendLine("<div class='signature-text-style'>")
        bodyBuilder.AppendLine("<div>Best regards,</div>")
        bodyBuilder.AppendLine("<div class='designation'>DMC Cloud Administrator</div>")
        bodyBuilder.AppendLine("<div class='companyName'>DIGI SINGAPORE PTE. LTD.</div>")
        bodyBuilder.AppendLine("<div class='departmentName'>Business Development Division</div>")
        bodyBuilder.AppendLine("<div class='companyAddrress'>4 Leng Kee Road, SIS Building, #06-01, Singapore 159088</div>")
        bodyBuilder.AppendLine("<div class='companyAddrress'>Phone: +65 6378 2121</div>")
        bodyBuilder.AppendLine("<div class='companyAddrress'>Mail: DMC_admin@sg.digi.inc</div>")
        bodyBuilder.AppendLine("<div class='companyAddrress'>Web: www.digisystem.com/sg/</div>")
        bodyBuilder.AppendLine("</div>")

        ' Final body content
        bodyBuilder.AppendLine("</body>")
        bodyBuilder.AppendLine("</html>")

        Return bodyBuilder.ToString()
    End Function

    Private Function contentStringBuilder(ByVal notificationType As String, ByVal recipientID As String) As String
        Dim emailBody As String = Nothing

        Try
            Dim bodyBuilder As New System.Text.StringBuilder()
            ' Dynamic body content
            bodyBuilder.AppendLine("<table class='primary-table'>")
            bodyBuilder.AppendLine("<thead>")

            ' Retrieve column headers and data dynamically from the database
            Dim query As String = GetSQL(notificationType, recipientID)

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        ' Retrieve headers
                        If reader.HasRows Then
                            If reader.Read() Then
                                bodyBuilder.AppendLine("<tr>")
                                For i As Integer = 0 To reader.FieldCount - 1
                                    bodyBuilder.AppendLine($"<th>{reader.GetName(i)}</th>")
                                Next
                                bodyBuilder.AppendLine("</tr>")
                            End If

                            bodyBuilder.AppendLine("</thead>")
                            bodyBuilder.AppendLine("<tbody>")

                            ' Retrieve data rows
                            Do
                                bodyBuilder.AppendLine("<tr>")
                                For i As Integer = 0 To reader.FieldCount - 1
                                    Dim columnName As String = reader.GetName(i)
                                    If columnName.Contains("Date") Then
                                        bodyBuilder.AppendLine($"<td class='nowrap'>{Convert.ToDateTime(reader(i)).ToString("dd MMM yyyy")}</td>")
                                    Else
                                        bodyBuilder.AppendLine($"<td>{reader(i)}</td>")
                                    End If
                                Next
                                bodyBuilder.AppendLine("</tr>")
                            Loop While reader.Read()

                            bodyBuilder.AppendLine("</tbody>")
                        Else
                            'Return Nothing
                            bodyBuilder.AppendLine("<p>There is no record this month</p>")
                        End If
                    End Using
                End Using
            End Using

            bodyBuilder.AppendLine("</table>")

            emailBody = bodyBuilder.ToString()
        Catch ex As Exception
            ' Handle any potential exceptions here (e.g., log the error)
            emailBody = Nothing
        End Try

        Return emailBody
    End Function

End Module
