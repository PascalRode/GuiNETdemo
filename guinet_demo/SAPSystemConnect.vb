Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SAP.Middleware.Connector


Class SAPSystemConnect
    Implements IDestinationConfiguration


    Public Function ChangeEventsSupported() As Boolean
        Throw New NotImplementedException()
    End Function

    Public Event ConfigurationChanged As RfcDestinationManager.ConfigurationChangeHandler

    Public Function GetParameters_obsolete(ByVal destinationName As String) As RfcConfigParameters
        Dim parms As New RfcConfigParameters()

        Dim choose = 1

        Select Case choose

            Case 1
                If "Dev".Equals(destinationName) Then

                    parms.Add(RfcConfigParameters.AppServerHost, "85.214.113.141")
                    parms.Add(RfcConfigParameters.Name, "ECC")
                    parms.Add(RfcConfigParameters.SystemNumber, "00")
                    parms.Add(RfcConfigParameters.User, "s10connect")
                    parms.Add(RfcConfigParameters.Password, "connect")
                    parms.Add(RfcConfigParameters.Client, "800")
                    parms.Add(RfcConfigParameters.Language, "DE")
                    parms.Add(RfcConfigParameters.PoolSize, "5")
                    parms.Add(RfcConfigParameters.PeakConnectionsLimit, "10")

                    parms.Add(RfcConfigParameters.ConnectionIdleTimeout, "600")
                End If

            Case 2
                If "Dev".Equals(destinationName) Then

                    parms.Add(RfcConfigParameters.AppServerHost, "/H/localhost/S/3200")
                    parms.Add(RfcConfigParameters.SystemNumber, "00")
                    parms.Add(RfcConfigParameters.Name, "ECC")
                    parms.Add(RfcConfigParameters.User, "")
                    parms.Add(RfcConfigParameters.Password, "")
                    parms.Add(RfcConfigParameters.Client, "800")
                    parms.Add(RfcConfigParameters.Language, "DE")
                    parms.Add(RfcConfigParameters.PoolSize, "5")
                    parms.Add(RfcConfigParameters.PeakConnectionsLimit, "10")

                    parms.Add(RfcConfigParameters.ConnectionIdleTimeout, "600")
                End If

        End Select

        Return parms
    End Function

    Public Function ChangeEventsSupported1() As Boolean Implements SAP.Middleware.Connector.IDestinationConfiguration.ChangeEventsSupported

    End Function

    Public Event ConfigurationChanged1(ByVal destinationName As String, ByVal args As SAP.Middleware.Connector.RfcConfigurationEventArgs) Implements SAP.Middleware.Connector.IDestinationConfiguration.ConfigurationChanged

    Public Function GetParameters1(ByVal destinationName As String) As SAP.Middleware.Connector.RfcConfigParameters Implements SAP.Middleware.Connector.IDestinationConfiguration.GetParameters

        Return Nothing

    End Function
End Class

