Option Explicit On
Option Strict On


Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Xml


Public Class _UpdateExchangeRate

    Dim strSessionURL As String
    Dim bTest As Boolean = True
    Dim logWrite As System.IO.StreamWriter
    Dim strFileName As String



    Public Sub ExchangeRate()
        Dim objSession = New Session


        Try
            strFileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\crmod_ws.log"
            If Not FileIO.FileSystem.FileExists(strFileName) Then File.Create(strFileName) ' check if logfile exists and create if not


            UpdateStatus(" ")
            UpdateStatus(" *** Starting Exchange Rate Application... ***") ' send message to application and log file

            CRMOnDemandLogin(objSession, "Production")

            UpdateExchangeRate()

            CRMOnDemandLogout(objSession)

            UpdateStatus("---------------------------------------------------------")

        Catch ex As Exception
            UpdateStatus(ex.Message)
        End Try
    End Sub


    Private Sub CRMOnDemandLogin(ByRef objSession As Session, ByVal strLogIn As String)

        If strLogIn = "Production" Then
            objSession.username = "ASBG/Admin"
            objSession.password = "0nD3mand"
            objSession.server = "secure-ausomxjna.crmondemand.com"
        End If


        UpdateStatus("Logging In To CRMOD: " & objSession.server & " with Username: " & objSession.username)

        objSession.Establish()
        strSessionURL = objSession.GetURL
        UpdateStatus("SessionID: " & objSession.sessionId)
    End Sub

    Private Sub CRMOnDemandLogout(ByRef objSession As Session)
        UpdateStatus("Logging Out of CRMOD: " & objSession.server & " with Username: " & objSession.username)
        objSession.Destroy()
    End Sub

    Private Sub UpdateExchangeRate()


        Dim SListofExchange As New ExchangeRate.ListOfExchangeRateQuery
        Dim rtnListofExchange As New ExchangeRate.ExchangeRateQuery
        Dim rate As New ExchangeRate.queryType
        Dim currencycodeFrom As New ExchangeRate.queryType
        Dim currencycodeTo As New ExchangeRate.queryType
        Dim rateDate As New ExchangeRate.queryType
        Dim xRateID As New ExchangeRate.queryType


        Dim qryIn As New ExchangeRate.ExchangeRateQueryPage_Input
        Dim qryOut As New ExchangeRate.ExchangeRateQueryPage_Output

        Dim xRatePrxy As New ExchangeRate.ExchangeRate

        Try
            xRatePrxy.Url = strSessionURL

            'currencycodeFrom.Value = "='EUR'"
            'currencycodeTo.Value = "='USD'"

            currencycodeFrom.Value = ""
            currencycodeTo.Value = ""


            SListofExchange.recordcountneeded = True
            SListofExchange.recordcountneededSpecified = True
            SListofExchange.pagesize = "100"
            SListofExchange.startrownum = "0"

            rtnListofExchange.ToCurrencyCode = currencycodeTo
            rtnListofExchange.FromCurrencyCode = currencycodeFrom
            rtnListofExchange.ExchangeRate = rate
            rtnListofExchange.ExchangeDate = rateDate
            rtnListofExchange.Id = xRateID

            SListofExchange.ExchangeRate = rtnListofExchange
            qryIn.ListOfExchangeRate = SListofExchange

            qryOut = xRatePrxy.ExchangeRateQueryPage(qryIn)

            getCurrentExchangeRatesandUpdateQryOut_v2(qryOut.ListOfExchangeRate)


            Update_ExchangeRate(qryOut.ListOfExchangeRate)


            qryIn = Nothing
            qryOut = Nothing
            xRatePrxy = Nothing
        Catch ex As Exception
            UpdateStatus(ex.Message)
        End Try

    End Sub

    Private Sub GetExchangeRateFromWeb(ByVal strCurrency As String, ByRef strExchangeRate As String, ByRef bError As Boolean)
        Dim url As String = "https://free.currencyconverterapi.com/api/v5/convert?q=" & strCurrency & "_USD&compact=ultra"
        Dim request As WebRequest = WebRequest.Create(url)
        Dim response As WebResponse = request.GetResponse()

        ' Get the stream containing content returned by the server.
        Dim dataStream As Stream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim readerURL As New StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = readerURL.ReadToEnd()
        ' Clean up the streams and the response.
        If responseFromServer = "{}" Then
            bError = True
        Else
            bError = False
            Dim iStart As Integer = InStr(responseFromServer, ":") + 1
            Dim iEnd As Integer = InStr(responseFromServer, "}")
            strExchangeRate = Mid(responseFromServer, iStart, iEnd - iStart)
        End If
    End Sub


    Sub getCurrentExchangeRatesandUpdateQryOut_v2(ByRef qryOut As ExchangeRate.ListOfExchangeRateData)
        Dim strDate As String = DateTime.Now.ToString("MM/dd/yyyy")
        Dim strCurrencyCode As String
        Dim intNumCurrencies As Integer
        Dim strXrate As String = ""
        Dim bError As Boolean


        For intNumCurrencies = 0 To CInt(qryOut.recordcount) - 1
            strCurrencyCode = qryOut.ExchangeRate(intNumCurrencies).FromCurrencyCode
            GetExchangeRateFromWeb(strCurrencyCode, strXrate, bError)
            qryOut.ExchangeRate(intNumCurrencies).ExchangeRate = System.Convert.ToDecimal(strXrate)
            qryOut.ExchangeRate(intNumCurrencies).ExchangeDate = System.Convert.ToDateTime(strDate)


            UpdateStatus("Currency: " & strCurrencyCode & "  Rate: " & strXrate & " Date: " & strDate)
        Next


    End Sub

    Sub getCurrentExchangeRatesandUpdateQryOut(ByRef qryOut As ExchangeRate.ListOfExchangeRateData)
        Dim strDate As String
        Dim strRate As String
        Dim strName As String
        Dim strCurrencyCode As String
        Dim intNumCurrencies As Integer

        Dim url As String = "http://query.yahooapis.com/v1/public/yql?q=select * from yahoo.finance.xchange where pair in (""EURUSD"", ""CHFUSD"", ""BRLUSD"", ""CNYUSD"", ""SGDUSD"", ""MXNUSD"", ""GBPUSD"", ""JPYUSD"")&env=store://datatables.org/alltableswithkeys"
        Dim request As WebRequest = WebRequest.Create(url)
        Dim response As WebResponse = request.GetResponse()

        ' Get the stream containing content returned by the server.
        Dim dataStream As Stream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim readerURL As New StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = readerURL.ReadToEnd()
        ' Clean up the streams and the response.

        Dim reader As XmlReader = XmlReader.Create(New StringReader(responseFromServer))
        Do While reader.ReadToFollowing("rate")

            reader.MoveToFirstAttribute()
            strCurrencyCode = reader.Value

            reader.ReadToFollowing("Name")
            strName = reader.ReadElementContentAsString()

            'reader.ReadToFollowing("Rate")
            strRate = reader.ReadElementContentAsString()

            'reader.ReadToFollowing("Date")
            strDate = reader.ReadElementContentAsString()

            For intNumCurrencies = 0 To CInt(qryOut.recordcount) - 1
                If Left(strCurrencyCode, 3) = qryOut.ExchangeRate(intNumCurrencies).FromCurrencyCode Then
                    qryOut.ExchangeRate(intNumCurrencies).ExchangeRate = System.Convert.ToDecimal(strRate)
                    qryOut.ExchangeRate(intNumCurrencies).ExchangeDate = System.Convert.ToDateTime(strDate)
                End If
            Next

            UpdateStatus("Currency: " & strCurrencyCode & "  Rate: " & strRate & " Date: " & strDate)
        Loop

        readerURL.Close()
        response.Close()
        reader.Close()

    End Sub

    Private Sub Update_ExchangeRate(ByRef pobjListofExhange As ExchangeRate.ListOfExchangeRateData)
        'Create Query Update
        Dim qryUpdateIn As New ExchangeRate.ExchangeRateUpdate_Input
        Dim qryUpdateOut As ExchangeRate.ExchangeRateUpdate_Output

        Dim xRatePrxy As New ExchangeRate.ExchangeRate

        Try
            xRatePrxy.Url = strSessionURL

            qryUpdateIn.ListOfExchangeRate = pobjListofExhange
            qryUpdateOut = xRatePrxy.ExchangeRateUpdate(qryUpdateIn)

        Catch ex As Exception
            UpdateStatus(ex.Message)
        End Try

        qryUpdateIn = Nothing
        qryUpdateOut = Nothing
        xRatePrxy = Nothing

    End Sub

    Private Sub UpdateStatus(ByVal strStatusMessage As String, Optional ByVal strType As String = "")
        Dim strStatus As String = "[" & DateTime.Now & "] " & strStatusMessage

        Console.WriteLine(strStatusMessage)

        logWrite = File.AppendText(strFileName)
        logWrite.WriteLine(strStatus)
        logWrite.Close()
    End Sub

End Class