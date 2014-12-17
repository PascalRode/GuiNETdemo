Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SAP.Middleware.Connector

Class DotNetConnector

    Protected CustomerNo As String
    Protected CustomerName As String
    Protected Address As String
    Protected City As String
    Protected StateProvince As String
    Protected CountryCode As String
    Protected PostalCode As String
    Protected Region As String
    Protected Industry As String
    Protected District As String
    Protected SalesOrg As String
    Protected DistributionChannel As String
    Protected Division As String

    Public Sub GetCustomerDetails(ByVal destination As RfcDestination)


        Try
            Dim repo As RfcRepository = destination.Repository
            Dim customerList As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_GETLIST")


            customerList.Invoke(destination)

            Dim idRange As IRfcTable = customerList.GetTable("IdRange")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "BT")
            idRange.SetValue("LOW", "0000001000")
            idRange.SetValue("HIGH", "0000002000")


            'add selection range to customerList function to search for all customers
            customerList.SetValue("idrange", idRange)


            Dim addressData As IRfcTable = customerList.GetTable("AddressData")
            customerList.Invoke(destination)

            For cuIndex As Integer = 0 To addressData.RowCount - 1

                addressData.CurrentIndex = cuIndex
                Dim customerHierachy As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_GETSALESAREAS")
                Dim customerDetail1 As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_GETDETAIL1")
                Dim customerDetail2 As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_GETDETAIL2")

                Me.CustomerNo = addressData.GetString("Customer")
                Me.CustomerName = addressData.GetString("Name")
                Me.Address = addressData.GetString("Street")
                Me.City = addressData.GetString("City")
                Me.StateProvince = addressData.GetString("Region")
                Me.CountryCode = addressData.GetString("CountryISO")
                Me.PostalCode = addressData.GetString("Postl_Cod1")

                customerDetail2.SetValue("CustomerNo", Me.CustomerNo)
                customerDetail2.Invoke(destination)
                Dim generalDetail As IRfcStructure = customerDetail2.GetStructure("CustomerGeneralDetail")

                Me.Region = generalDetail.GetString("Reg_Market")
                Me.Industry = generalDetail.GetString("Industry")


                customerDetail1.Invoke(destination)
                Dim detail1 As IRfcStructure = customerDetail1.GetStructure("PE_CompanyData")

                Me.District = detail1.GetString("District")


                customerHierachy.Invoke(destination)
                customerHierachy.SetValue("CustomerNo", Me.CustomerNo)
                customerHierachy.Invoke(destination)

                Dim otherDetail As IRfcTable = customerHierachy.GetTable("SalesAreas")

                If otherDetail.RowCount > 0 Then
                    Me.SalesOrg = otherDetail.GetString("SalesOrg")
                    Me.DistributionChannel = otherDetail.GetString("DistrChn")
                    Me.Division = otherDetail.GetString("Division")
                End If

                customerHierachy = Nothing
                customerDetail1 = Nothing
                customerDetail2 = Nothing
                GC.Collect()



                GC.WaitForPendingFinalizers()


            Next

        Catch e As RfcCommunicationException
            ' user could not logon...
        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
        Catch e As RfcAbapRuntimeException
            ' The function module returned an ABAP exception, an ABAP message
            ' or an ABAP class-based exception...
        Catch e As RfcAbapBaseException
        End Try

    End Sub

    Public Function SearchCustomers(ByVal destination As RfcDestination, ByVal searchName As String) As List(Of String())
        Dim customers As New List(Of String())


        Try
            Dim repo As RfcRepository = destination.Repository
            Dim customerList As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_FIND")

            customerList.Invoke(destination)

            customerList.SetValue("PL_HOLD", "X") 'Wildcard search on
            customerList.SetValue("MAX_CNT", "0")

            Dim idRange As IRfcTable = customerList.GetTable("SELOPT_TAB")

            idRange.Append()
            idRange.SetValue("COMP_CODE", "1000")
            idRange.SetValue("TABNAME", "KNA1")
            idRange.SetValue("FIELDNAME", "NAME1")
            idRange.SetValue("FIELDVALUE", searchName & "*")


            'add selection range to customerList function to search for all customers
            customerList.SetValue("SELOPT_TAB", idRange)


            Dim addressData As IRfcTable = customerList.GetTable("RESULT_TAB")
            customerList.Invoke(destination)

            For cuIndex As Integer = 0 To addressData.RowCount - 1

                addressData.CurrentIndex = cuIndex


                Me.CustomerNo = addressData.GetString("Customer")

                Dim detailsDic = GetCustomerDetails(destination, Me.CustomerNo)

                If detailsDic.Count > 0 Then

                    customers.Add(New String() {detailsDic.Item("Customer"), detailsDic.Item("Name")})

                End If


                GC.Collect()



                GC.WaitForPendingFinalizers()


            Next

        Catch e As RfcCommunicationException
            ' user could not logon...
        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
        Catch e As RfcAbapRuntimeException
            ' The function module returned an ABAP exception, an ABAP message
            ' or an ABAP class-based exception...
        Catch e As RfcAbapBaseException
        End Try

        Return customers


    End Function
    Public Function GetCustomerDetails(ByVal destination As RfcDestination, ByVal kunnr As String) As Dictionary(Of String, String)

        Dim returnDic = New Dictionary(Of String, String)

        Try
            Dim repo As RfcRepository = destination.Repository
            Dim customerList As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_GETLIST")


            customerList.Invoke(destination)

            If IsNumeric(kunnr) Then
                kunnr = kunnr.PadLeft(10, "0")
            End If

            Dim idRange As IRfcTable = customerList.GetTable("IdRange")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "BT")
            idRange.SetValue("LOW", kunnr)
            idRange.SetValue("HIGH", kunnr)


            'add selection range to customerList function to search for all customers
            customerList.SetValue("idrange", idRange)


            Dim addressData As IRfcTable = customerList.GetTable("AddressData")
            customerList.Invoke(destination)

            For cuIndex As Integer = 0 To addressData.RowCount - 1

                addressData.CurrentIndex = cuIndex


                returnDic.Item("Customer") = addressData.GetString("Customer")
                returnDic.Item("Name") = addressData.GetString("Name")
                returnDic.Item("Street") = addressData.GetString("Street")
                returnDic.Item("City") = addressData.GetString("City")
                returnDic.Item("Region") = addressData.GetString("Region")
                returnDic.Item("CountryISO") = addressData.GetString("CountryISO")
                returnDic.Item("Postl_Cod1") = addressData.GetString("Postl_Cod1")





                GC.Collect()



                GC.WaitForPendingFinalizers()


            Next

        Catch e As RfcCommunicationException
            ' user could not logon...
        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
        Catch e As RfcAbapRuntimeException
            ' The function module returned an ABAP exception, an ABAP message
            ' or an ABAP class-based exception...
        Catch e As RfcAbapBaseException
        End Try

        Return returnDic


    End Function


    Public Function GetCustomers(ByVal destination As RfcDestination, ByVal searchName As String) As List(Of String())


        Dim customers As New List(Of String())


        Try
            Dim repo As RfcRepository = destination.Repository
            Dim customerList As IRfcFunction = repo.CreateFunction("BAPI_CUSTOMER_GETLIST")


            customerList.Invoke(destination)

            Dim idRange As IRfcTable = customerList.GetTable("IdRange")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "BT")
            idRange.SetValue("LOW", "0000000000")
            idRange.SetValue("HIGH", "9999999999")


            'add selection range to customerList function to search for all customers
            customerList.SetValue("idrange", idRange)


            Dim addressData As IRfcTable = customerList.GetTable("AddressData")
            customerList.Invoke(destination)

            For cuIndex As Integer = 0 To addressData.RowCount - 1

                addressData.CurrentIndex = cuIndex


                Me.CustomerNo = addressData.GetString("Customer")
                Me.CustomerName = addressData.GetString("Name")
                Me.Address = addressData.GetString("Street")
                Me.City = addressData.GetString("City")
                Me.StateProvince = addressData.GetString("Region")
                Me.CountryCode = addressData.GetString("CountryISO")
                Me.PostalCode = addressData.GetString("Postl_Cod1")


                If Me.CustomerName.ToLower.Contains(searchName.ToLower) Then
                    customers.Add(New String() {Me.CustomerNo, Me.CustomerName})
                End If


                GC.Collect()



                GC.WaitForPendingFinalizers()


            Next

        Catch e As RfcCommunicationException
            MsgBox(e.Message)

        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
            MsgBox(e.Message)

        Catch e As RfcAbapRuntimeException
            ' The function module returned an ABAP exception, an ABAP message
            ' or an ABAP class-based exception...
            MsgBox(e.Message)

        Catch e As RfcAbapBaseException
            MsgBox(e.Message)

        End Try

        Return customers

    End Function

    Public Function GetMessagesToFuncLoc(ByVal destination As RfcDestination, ByVal funcloc As String) As List(Of String)


        Dim messages As New List(Of String)


        Try
            Dim repo As RfcRepository = destination.Repository
            Dim paras As IRfcFunction = repo.CreateFunction("BAPI_ALM_NOTIF_LIST_FUNCLOC")

            paras.Invoke(destination)

            paras.SetValue("FUNCLOC", funcloc)
            paras.SetValue("NOTIFICATION_DATE", "20140101")
            paras.SetValue("PARTNER", "0000000000")


            paras.Invoke(destination)


            Dim results As IRfcTable = paras.GetTable("NOTIFICATION")
            Dim returnValue = paras.GetStructure("RETURN")





            For cuIndex As Integer = 0 To results.RowCount - 1

                results.CurrentIndex = cuIndex


                Me.CustomerNo = results.GetString("DESCRIPT")

                GC.Collect()

                GC.WaitForPendingFinalizers()

                messages.Add(results.GetString("DESCRIPT"))


            Next

        Catch e As RfcCommunicationException
            ' user could not logon...
        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
        Catch e As RfcAbapRuntimeException
            ' The function module returned an ABAP exception, an ABAP message
            ' or an ABAP class-based exception...
        Catch e As RfcAbapBaseException
        End Try

        Return messages

    End Function
    Public Function GetStatusServicemsg(ByVal destination As RfcDestination, ByVal msgNo As String) As String


        Dim status As String = ""


        Try
            Dim repo As RfcRepository = destination.Repository
            Dim paras As IRfcFunction = repo.CreateFunction("/GUIXT/SELECT")

            paras.SetValue("Table", "JEST")
            paras.SetValue("Condition", "OBJNR = 'QM" & msgNo.PadLeft(12, "0") & "'")
            paras.SetValue("Fields", "STAT")

            paras.Invoke(destination)

            status = paras.GetValue("V1")


        Catch e As RfcCommunicationException

            Form1.LogError(e.Message)
        Catch e As RfcLogonException

            Form1.LogError(e.Message)
        Catch e As RfcAbapRuntimeException

            Form1.LogError(e.Message)

        Catch e As RfcAbapBaseException

            Form1.LogError(e.Message)


        End Try

        Return status


    End Function


    Public Function GetClassDataServiceMessage(ByVal destination As RfcDestination, ByVal msgNo As String) As String


        Dim returnstr As String = ""


        Try
            Dim repo As RfcRepository = destination.Repository
            Dim paras As IRfcFunction = repo.CreateFunction("CLFM_SELECT_AUSP")

            paras.SetValue("CLASSTYPE", "015")
            paras.SetValue("OBJET", msgNo.PadLeft(12, "0") & "0001")
            paras.SetValue("MAFID", "O")
            paras.SetValue("TABLE", "QMEL")

            paras.Invoke(destination)

            Dim results As IRfcTable = paras.GetTable("EXP_AUSP")


            For cuIndex As Integer = 0 To results.RowCount - 1

                results.CurrentIndex = cuIndex

                results.GetString("ATWRT")

                GC.Collect()

                GC.WaitForPendingFinalizers()

                returnstr &= results.GetString("ATWRT")


            Next



        Catch e As RfcCommunicationException
            ' user could not logon...
        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
        Catch e As RfcAbapRuntimeException
            MsgBox(e.Message)

        Catch e As RfcAbapBaseException

            MsgBox(e.Message)

        End Try

        Return returnstr


    End Function


    Public Function GetClassDataServiceOrder(ByVal destination As RfcDestination, ByVal msgNo As String) As serviceorder


        Dim so As New serviceorder



        Try
            Dim repo As RfcRepository = destination.Repository
            Dim paras As IRfcFunction = repo.CreateFunction("BAPI_ALM_ORDER_GET_DETAIL")

            msgNo = msgNo.PadLeft(13, "0")
            'MsgBox(msgNo)

            paras.SetValue("NUMBER", msgNo)
            paras.Invoke(destination)


            'MsgBox(paras.GetStructure("ES_HEADER").GetString("ORDERID"))

            Dim idReturn As IRfcTable = paras.GetTable("RETURN")

            If idReturn.RowCount > 0 Then
                MsgBox(idReturn.GetString("MESSAGE"))
            End If

            Dim operations As IRfcTable = paras.GetTable("ET_OPERATIONS")

            Dim header = paras.GetStructure("ES_HEADER")
            so.CAUFVD_STTXT = header.GetString("SYS_STATUS")

            so.CAUFVD_IWERK = header.GetString("PLANPLANT")
            so.CAUFVD_INGPR = header.GetValue("PLANGROUP")

            Dim servicedata = paras.GetStructure("ES_SRVDATA")
            so.PMSDO_MATNR = servicedata.GetString("MATERIAL")




        Catch e As RfcCommunicationException
            ' user could not logon...
        Catch e As RfcLogonException
            ' serious problem on ABAP system side...
        Catch e As RfcAbapRuntimeException
            MsgBox(e.Message)

        Catch e As RfcAbapBaseException

            MsgBox(e.Message)

        End Try


        Return so




    End Function

    Public Function GetServiceOrders(ByVal destination As RfcDestination) As String


        Try

            Dim repo As RfcRepository = destination.Repository
            Dim customerList As IRfcFunction = repo.CreateFunction("BAPI_ALM_ORDERHEAD_GET_LIST")

            customerList.Invoke(destination)
            Dim idRange As IRfcTable = customerList.GetTable("IT_RANGES")

            idRange.Append()
            idRange.SetValue("FIELD_NAME", "SHOW_DOCS_WITH_FROM_DATE")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "EQ")
            idRange.SetValue("LOW_VALUE", "00010101")

            idRange.Append()
            idRange.SetValue("FIELD_NAME", "SHOW_DOCS_WITH_TO_DATE")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "EQ")
            idRange.SetValue("LOW_VALUE", "99991231")

            idRange.Append()
            idRange.SetValue("FIELD_NAME", "SHOW_OPEN_DOCUMENTS")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "EQ")
            idRange.SetValue("LOW_VALUE", "X")

            idRange.Append()
            idRange.SetValue("FIELD_NAME", "SHOW_DOCUMENTS_IN_PROCESS")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "EQ")
            idRange.SetValue("LOW_VALUE", "X")


            idRange.Append()
            idRange.SetValue("FIELD_NAME", "SHOW_COMPLETED_DOCUMENTS")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "EQ")
            idRange.SetValue("LOW_VALUE", "X")


            idRange.Append()
            idRange.SetValue("FIELD_NAME", "SHOW_HISTORICAL_DOCUMENTS")
            idRange.SetValue("SIGN", "I")
            idRange.SetValue("OPTION", "EQ")
            idRange.SetValue("LOW_VALUE", "X")

            customerList.Invoke(destination)

            Dim addressData As IRfcTable = customerList.GetTable("ET_RESULT")

            For cuIndex As Integer = 0 To addressData.RowCount - 1

            Next

            Return addressData.RowCount.ToString


        Catch ex As Exception

            Return ex.Message

        End Try


        Return "0"

    End Function


End Class

