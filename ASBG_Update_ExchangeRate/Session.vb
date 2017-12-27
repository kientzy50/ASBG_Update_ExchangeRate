Imports System
Imports System.Net
Imports System.Web.Services.Protocols

Public Class Session
    Public protocol As String = "https"
    Public server As String
    Public port As String = "443"
    Public username As String
    Public password As String
    Public sessionId As String = Nothing
    Public cookie As cookie

    Public Sub Establish()

        If Not sessionId Is Nothing Then
            Destroy()

        End If

        'Check that username and password have been specified
        CheckUsernamePassword()

        ' create a container for an HTTP request
        Dim req As HttpWebRequest

        req = WebRequest.Create(GetLogInURL())

        'username and password are passed as HTTP headers
        req.Headers.Add("UserName", username)
        req.Headers.Add("Password", password)

        ' cookie container has to be added to request in order to 
        ' retrieve the cookie from the response. 
        req.CookieContainer = New CookieContainer

        ' make the HTTP call
        Dim resp As HttpWebResponse
        resp = req.GetResponse()

        If resp.StatusCode = System.Net.HttpStatusCode.OK Then

            'store cookie for later...
            cookie = resp.Cookies("JSESSIONID")

            If cookie Is Nothing Then
                Throw New Exception("No JSESSIONID cookie found in log-in response!")
            End If

            sessionId = cookie.Value

        End If

    End Sub

    Public Sub Destroy()
        ' create a container for an HTTP request
        Dim req As HttpWebRequest
        req = WebRequest.Create(GetLogOffURL())

        ' reuse the cookie that was received at Login
        req.CookieContainer = New CookieContainer
        req.CookieContainer.Add(cookie)

        ' make the HTTP call
        Dim resp As HttpWebResponse
        resp = req.GetResponse()
        If resp.StatusCode <> System.Net.HttpStatusCode.OK Then
            Throw New Exception("Logging off failed!")
        End If

        ' forget current session id
        sessionId = Nothing

    End Sub

    Public Function GetURL() As String

        If sessionId Is Nothing Then
            Throw New Exception("No session has been established!")
        End If

        CheckServerPort()
        Return protocol + "://" + server + "/Services/Integration;jsessionid=" + sessionId
    End Function

    Public Function GetURL(ByVal obj As String) As String

        If sessionId Is Nothing Then

            Throw New Exception("No session has been established!")
        End If
        CheckServerPort()
        Return protocol + "://" + server + ":" + port + "/Services/Integration/" + obj
    End Function

    Private Function GetLogInURL() As String
        CheckServerPort()
        Return protocol + "://" + server + "/Services/Integration?command=login"
    End Function

    Private Function GetLogOffURL() As String
        CheckServerPort()
        Return protocol + "://" + server + "/Services/Integration?command=logoff"
    End Function


    Private Sub CheckServerPort()

        If server Is Nothing Then
            Throw New Exception("Server not specified!")
        End If

        If port Is Nothing Then
            Throw New Exception("Port not specified!")
        End If
    End Sub
    Private Sub CheckUsernamePassword()

        If username Is Nothing Then
            Throw New Exception("Username not specified!")
        End If

        If password Is Nothing Then
            Throw New Exception("Password not specified!")
        End If
    End Sub
End Class


