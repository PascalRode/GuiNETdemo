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

    Public Function ChangeEventsSupported1() As Boolean Implements SAP.Middleware.Connector.IDestinationConfiguration.ChangeEventsSupported

    End Function

    Public Event ConfigurationChanged1(ByVal destinationName As String, ByVal args As SAP.Middleware.Connector.RfcConfigurationEventArgs) Implements SAP.Middleware.Connector.IDestinationConfiguration.ConfigurationChanged

    Public Function GetParameters1(ByVal destinationName As String) As SAP.Middleware.Connector.RfcConfigParameters Implements SAP.Middleware.Connector.IDestinationConfiguration.GetParameters

        Return Nothing

    End Function
End Class

