Imports guinet
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Imports SAP.Middleware.Connector

Imports System.Data.SqlClient
Imports Microsoft.Office.Interop



Public Class Form1
    ' connection to sap system
    Public Shared sapconnect As String
    Public Shared sapconnectclient As String
    Public Shared sapconnectuser As String
    Public Shared sapconnectpassword As String
    Public Shared sapconnectlanguage As String
    Public Shared sapgui As sapguisession = Nothing

    Public Shared parameters As RfcConfigParameters = New RfcConfigParameters

    Public service_reminder As String


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Configuration_Load()


        Dim c As Integer = 0
        For Each cp As String In My.Application.CommandLineArgs

            Select Case cp

                Case "-servicereminder"

                    service_reminder = My.Application.CommandLineArgs(c + 1)

                    Dim Timer1 As New Timer()
                    AddHandler Timer1.Tick, AddressOf Timer1_Tick
                    Timer1.Interval = 3000
                    Timer1.Start()

                Case "-minimized"

                    Me.WindowState = FormWindowState.Minimized

            End Select


        Next




      

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As EventArgs)

        sender.Stop()

        Me.WindowState = FormWindowState.Normal
        MessageBox.Show(servicemsg_getstatus(service_reminder), "Wichtiger Hinweis!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        'MessageBox.Show("Wichtiger Hinweis", "Wichtiger Hinweis!", MessageBoxButtons.OK, MessageBoxIcon.Warning)



    End Sub


    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As EventArgs)

        Dim t As Timer = sender
        t.Stop()
        MsgBox(t.Tag)


    End Sub



    Public Sub SetLogonParameters()


        sapconnect = sap_routerstring.Text
        sapconnectclient = sap_client.Text
        sapconnectuser = sap_username.Text
        sapconnectpassword = sap_password.Text
        sapconnectlanguage = sap_language.Text

        parameters.Item(RfcConfigParameters.Name) = sap_systemname.Text
        parameters.Item(RfcConfigParameters.User) = sap_username.Text
        parameters.Item(RfcConfigParameters.Password) = sap_password.Text
        parameters.Item(RfcConfigParameters.Client) = sap_client.Text
        parameters.Item(RfcConfigParameters.Language) = sap_language.Text
        parameters.Item(RfcConfigParameters.AppServerHost) = sap_appserverhost.Text
        parameters.Item(RfcConfigParameters.SystemNumber) = sap_systemnumber.Text




    End Sub


    Public Function servicemsg_getstatus(ByVal msgid As String) As String

        Configuration_Load()
        SetLogonParameters()

        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)
            sapgui.Logon("", check_sapguiscripting.Checked)

        End If



        LogInfo("Start transaction IW22 and get data")


        ' Start transaktion IW22
        sapgui.Enter("/nIW22")

        ' Enter number of service msg
        sapgui.SetField("RIWO00-QMNUM", servicemsg.Text)
        sapgui.Enter()

        Return sapgui.GetField("RIWO00-STTXT")


    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TabControl2.SelectTab(0)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        LogInfo("Logon to SAP GUI")

        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)
            sapgui.Logon("", check_sapguiscripting.Checked)
        End If

        LogInfo("Start transaction VA05 and get data")

        ' Start transaktion VA05
        sapgui.Enter("/nVA05")



        ' Enter number of productionorder
        sapgui.SetField("VBCOM-KUNDE", kunnr.Text)
        sapgui.SetField("VBCOM-VKORG", "1000")
        sapgui.SetField("VBCOM-AUDAT", "1.1.1999")

        sapgui.Enter()


        If sapgui.GetField("VBCOM-VKORG") = "?" Then
            sapgui.SetField("VBCOM-VKORG", SalesOrgTextbox.Text)
            sapgui.Enter()
        End If

        Dim dt As DataTable = New DataTable("Orders")

        dt.Columns.Add("VBELN")
        dt.Columns.Add("NETPR")
        dt.Columns.Add("ARKTX")
        dt.Columns.Add("POSNR")
        dt.Columns.Add("KWMENG")
        dt.Columns.Add("VRKME")


        LogInfo("Read SAP grid")

        sapgui.ReadGrid(dt)

        LogInfo("Got " & dt.Rows.Count.ToString & " entries")

        dgv.Columns.Clear()

        ' Add columns
        dgv.Columns.Add("VBELN", "Order no.")
        dgv.Columns.Add("NETPR", "Net value")
        dgv.Columns.Add("ARKTX", "Description")
        dgv.Columns.Add("POSNR", "ITEM")
        dgv.Columns.Add("KWMENG", "Order qty")
        dgv.Columns.Add("VRKME", "SU")


        dgv.Columns("NETPR").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        ' Add rows
        For i = 0 To dt.Rows.Count - 1

            dgv.Rows.Add(dt.Rows(i).ItemArray)

        Next

        dgv.AutoResizeColumns()


        System.Windows.Forms.Cursor.Current = Cursors.Default


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        TabControl3.SelectTab(0)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        LogInfo("Logon to SAP GUI")

        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)
            sapgui.Logon("", check_sapguiscripting.Checked)
        End If


        LogInfo("Start transaction MM60 and get data")

        ' Start transaktion MM60 
        sapgui.Enter("/nMM60")

        ' Enter material type
        sapgui.SetField("MTART-LOW", mtart.Text)

        ' Excecute with fkey F8
        ' More keys:
        ' http://www.sapdesignguild.org/resources/references/nv_fkeys_ref2_e.htm
        sapgui.Enter("/8")
        ' Alternative: Access menu via short keys (language dependant!)
        'sapgui.Enter(".PA")

        Dim dt As DataTable = New DataTable("Materials")
        dt.Columns.Add("MATNR")
        dt.Columns.Add("KTEXT")
        dt.Columns.Add("PREIS")
        dt.Columns.Add("WAERS")

        LogInfo("Read SAP grid")

        sapgui.ReadGrid(dt)

        LogInfo("Got " & dt.Rows.Count.ToString & " entries")

        dgv2.Columns.Clear()

        ' Spalten hinzufügen
        dgv2.Columns.Add("MATNR", "Material")
        dgv2.Columns.Add("KTEXT", "Material description")
        dgv2.Columns.Add("PREIS", "Price")
        dgv2.Columns.Add("WAERS", "Currency")

        dgv2.Columns("PREIS").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



        ' Zeilen hinzufügen
        For i = 0 To dt.Rows.Count - 1


            dgv2.Rows.Add(dt.Rows(i).ItemArray)

        Next

        dgv2.AutoResizeColumns()

        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)

            sapgui.Logon("", check_sapguiscripting.Checked)
        End If



        System.Windows.Forms.Cursor.Current = Cursors.Default


    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        LogInfo("Logon to SAP GUI")

        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)

            sapgui.Logon("", check_sapguiscripting.Checked)
        End If

        LogInfo("Start transaction ALO1 and get data")


        ' Start transaktion ALO1
        sapgui.Enter("/nALO1")

        sapgui.Enter("/8")
        sapgui.Enter("/8")



        Dim dt As DataTable = New DataTable("DocFlow")
        dt.Columns.Add("VBELN")
        dt.Columns.Add("ERDAT")

        LogInfo("Read SAP table")
        sapgui.ReadTable(dt)

        LogInfo("Got " & dt.Rows.Count.ToString & " entries")


        dgv3.Columns.Clear()

        ' Spalten hinzufügen
        dgv3.Columns.Add("VBELN", "Sales Activity")
        dgv3.Columns.Add("ERDAT", "Created on")


        ' Zeilen hinzufügen
        For i = 0 To dt.Rows.Count - 1

            dgv3.Rows.Add(dt.Rows(i).ItemArray)

        Next

        dgv3.AutoResizeColumns()

        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub


#Region "Trees"


    Public Function AddFullPathToTreeView(ByRef treeView As TreeView, ByVal strPath As String, ByVal nodeText As String) As Boolean

        If IsNothing(treeView) OrElse treeView.IsDisposed OrElse strPath = "" Then Return False

        Dim branches() As String = strPath.Split("\")                       ' Split the path

        Dim ParentNode As TreeNode = Nothing
        Dim CurrentPath As String = branches(0)                             ' Root path
        For NodeIndex As Integer = 0 To (branches.Length - 1)               ' Loop through all node segments
            If ParentNode IsNot Nothing Then                                ' If we have a ParentNode then set the path using this
                CurrentPath = ParentNode.FullPath & "\" & branches(NodeIndex)
            End If
            Dim foundNodes() As TreeNode = treeView.Nodes.Find(CurrentPath, True)   ' Check if Node exists
            If foundNodes.Length <= 0 Then                                  ' If Node doesn't exist then create a new one
                Dim rootNode = New TreeNode(branches(NodeIndex))           ' Add a new  node

                If ParentNode Is Nothing Then                               ' Add node to Tree or ParentNode if it exists
                    ParentNode = treeView.Nodes.Add(CurrentPath, nodeText)
                Else
                    ParentNode = ParentNode.Nodes.Add(CurrentPath, nodeText)
                End If

            Else
                ParentNode = foundNodes(0)                                '  Node created previouslly so just set it as Parent
            End If
        Next
        Return True
    End Function


#End Region

    Private Sub vbeln_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles vbeln.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kunnr.TextChanged

    End Sub


    Public Function read_order_html(ByVal orderno As String) As String

        Dim HTML As String = "<html><head><meta charset='utf-8'>"
        HTML &= table_add_css() & "</head>"
        HTML &= "<body><hr>Kurzübersicht über Bestellung " & orderno & "<br><hr><br><br>"


        Configuration_Load()
        SetLogonParameters()

        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)
            sapgui.Logon("", check_sapguiscripting.Checked)

        End If

        ' Start transaktion VA03
        sapgui.Enter("/nVA03")

        ' Enter number of productionorder
        sapgui.SetField("VBAK-VBELN", orderno)
        sapgui.Enter()

        HTML &= "Nettowert: " & sapgui.GetField("VBAK-NETWR") & " " 'Bettowert
        HTML &= sapgui.GetField("VBAK-WAERK") & "<br>" 'Währung 
        HTML &= "Auftraggeber: " & sapgui.GetField("KUAGV-TXTPA") & "<br>" 'Auftraggeber 
        HTML &= "Warenempfänger : " & sapgui.GetField("KUWEV-KUNNRR") & "<br>" 'Warenempfänger      
        HTML &= "Bestellnumme: " & sapgui.GetField("VBKD-BSTKD") & "<br>"  'Bestellnummer
        HTML &= "Bestelldatum: " & sapgui.GetField("VBKD-BSTDK") & "<br>" 'Bestelldatum

        Dim dt As DataTable = New DataTable("Orders")
        dt.Columns.Add("POSNR")
        dt.Columns.Add("MABNR")
        dt.Columns.Add("KWMENG")
        dt.Columns.Add("VRKME")
        dt.Columns.Add("ARKTX")


        HTML &= "<br> Positionen: <br><br>"

        sapgui.ReadTable(dt)

        HTML &= "<table cellspacing='0'> 	<thead>		<tr>			<th>Pos</th>			<th>Mat.No.</th>			<th>Menge</th>	<th>Einheit</th><th>Beschreibung </thead>"

        HTML &= "<tbody>"



        For i = 0 To dt.Rows.Count - 1

            HTML &= "<tr>"
            HTML &= "<td>" & dt.Rows(i)(0) & "</td>"
            HTML &= "<td>" & dt.Rows(i)(1) & "</td>"
            HTML &= "<td>" & dt.Rows(i)(2) & "</td>"
            HTML &= "<td>" & dt.Rows(i)(3) & "</td>"
            HTML &= "<td>" & dt.Rows(i)(4) & "</td>"

            HTML &= "</tr>"

        Next


        HTML &= "</tbody></table>"

        Return HTML



    End Function

    Public Function table_add_css() As String

        Return "<style>table a:link {	color: #666;	font-weight: bold;	text-decoration:none;}table a:visited {	color: #999999;	font-weight:bold;	text-decoration:none;}table a:active,table a:hover {	color: #bd5a35;	text-decoration:underline;}table {	font-family:Arial, Helvetica, sans-serif;	color:#666;	font-size:12px;	text-shadow: 1px 1px 0px #fff;	background:#eaebec;	margin:20px;	border:#ccc 1px solid;	-moz-border-radius:3px;	-webkit-border-radius:3px;	border-radius:3px;	-moz-box-shadow: 0 1px 2px #d1d1d1;	-webkit-box-shadow: 0 1px 2px #d1d1d1;	box-shadow: 0 1px 2px #d1d1d1;}table th {	padding:21px 25px 22px 25px;	border-top:1px solid #fafafa;	border-bottom:1px solid #e0e0e0;	background: #ededed;	background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));	background: -moz-linear-gradient(top,  #ededed,  #ebebeb);}table th:first-child {	text-align: left;	padding-left:20px;}table tr:first-child th:first-child {	-moz-border-radius-topleft:3px;	-webkit-border-top-left-radius:3px;	border-top-left-radius:3px;}table tr:first-child th:last-child {	-moz-border-radius-topright:3px;	-webkit-border-top-right-radius:3px;	border-top-right-radius:3px;}table tr {	text-align: center;	padding-left:20px;}table td:first-child {	text-align: left;	padding-left:20px;	border-left: 0;}table td {	padding:18px;	border-top: 1px solid #ffffff;	border-bottom:1px solid #e0e0e0;	border-left: 1px solid #e0e0e0;	background: #fafafa;	background: -webkit-gradient(linear, left top, left bottom, from(#fbfbfb), to(#fafafa));	background: -moz-linear-gradient(top,  #fbfbfb,  #fafafa);}table tr.even td {	background: #f6f6f6;	background: -webkit-gradient(linear, left top, left bottom, from(#f8f8f8), to(#f6f6f6));	background: -moz-linear-gradient(top,  #f8f8f8,  #f6f6f6);}table tr:last-child td {	border-bottom:0;}table tr:last-child td:first-child {	-moz-border-radius-bottomleft:3px;	-webkit-border-bottom-left-radius:3px;	border-bottom-left-radius:3px;}table tr:last-child td:last-child {	-moz-border-radius-bottomright:3px;	-webkit-border-bottom-right-radius:3px;	border-bottom-right-radius:3px;}table tr:hover td {	background: #f2f2f2;	background: -webkit-gradient(linear, left top, left bottom, from(#f2f2f2), to(#f0f0f0));	background: -moz-linear-gradient(top,  #f2f2f2,  #f0f0f0);	}</style>"

    End Function

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click


        TabControl2.SelectTab(1)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor


        SetLogonParameters()

        LogInfo("Logon to SAP GUI")


        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)

            sapgui.Logon("", check_sapguiscripting.Checked)
        End If

        LogInfo("Start transaction VA03 and get data")


        ' Start transaktion VA03
        sapgui.Enter("/nVA03")

        ' Enter number of productionorder
        sapgui.SetField("VBAK-VBELN", vbeln.Text)
        sapgui.Enter()



        Dim dt As DataTable = New DataTable("Orders")
        dt.Columns.Add("MABNR")
        dt.Columns.Add("KWMENG")
        dt.Columns.Add("VRKME")
        dt.Columns.Add("ARKTX")

        NETWR.Text = sapgui.GetField("VBAK-NETWR")
        WAERK.Text = sapgui.GetField("VBAK-WAERK")
        RV45A_KPRGBZ.Text = sapgui.GetField("RV45A-KPRGBZ")
        RV45A_KETDAT.Text = sapgui.GetField("RV45A-KETDAT")
        KUAGV_KUNNR.Text = sapgui.GetField("KUAGV-KUNNR")
        KUAGV_TXTPA.Text = sapgui.GetField("KUAGV-TXTPA")
        VBKD_BSTKD.Text = sapgui.GetField("VBKD-BSTKD")


        sapgui.ReadTable(dt)


        dgv_orderitems.Columns.Clear()

        ' Add columns
        dgv_orderitems.Columns.Add("MABNR", "Material no.")
        dgv_orderitems.Columns.Add("KWMENG", "Quantity")
        dgv_orderitems.Columns.Add("VRKME", "Unit")
        dgv_orderitems.Columns.Add("ARKTX", "Description")

        dgv_orderitems.Columns("KWMENG").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight


        ' Add rows
        For i = 0 To dt.Rows.Count - 1


            dgv_orderitems.Rows.Add(dt.Rows(i).ItemArray)

        Next

        dgv_orderitems.AutoResizeColumns()



        ' Start transaktion VA03
        sapgui.Enter("/nVA03")

        ' Enter number of order
        sapgui.SetField("VBAK-VBELN", vbeln.Text)
        sapgui.Enter()
        sapgui.Enter("/5")

        dt = New DataTable("docflow_orders")

        dt.Columns.Add("nodekey")
        dt.Columns.Add("nodepath")
        dt.Columns.Add("nodetext")
        dt.Columns.Add("nodelevel")
        dt.Columns.Add("nodelabel")

        LogInfo("Read SAP tree")
        sapgui.ReadTree(dt)
        LogInfo("Got " & dt.Rows.Count.ToString & " entries")

        dgv3.Columns.Clear()

        ' Add columns
        dgv3.Columns.Add("nodekey", "nodekey")
        dgv3.Columns.Add("nodepath", "nodepath")
        dgv3.Columns.Add("nodetext", "nodetext")
        dgv3.Columns.Add("nodelevel", "nodelevel")
        dgv3.Columns.Add("nodelabel", "nodelabel")

        TreeView1.Nodes.Clear()

        LogInfo("Convert table to treeview")

        ' Add rows
        For i = 0 To dt.Rows.Count - 1

            dgv3.Rows.Add(dt.Rows(i).ItemArray)

            Dim s = dt.Rows(i).ItemArray


            AddFullPathToTreeView(TreeView1, s(1), s(2))


        Next

        dgv3.AutoResizeColumns()

        System.Windows.Forms.Cursor.Current = Cursors.Default



    End Sub

    Private Sub dgv_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellContentClick



    End Sub

    Private Sub dgv_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellDoubleClick
        Dim CurrentRow = e.RowIndex
        Dim CurrentCol = e.ColumnIndex
        Dim CurrentColName = dgv.Columns(e.ColumnIndex).Name

        If CurrentColName = "VBELN" Then

            vbeln.Text = dgv.Item(CurrentCol, CurrentRow).Value.ToString()




        End If
    End Sub

    Private Sub help_tab1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub


    Private Sub Button3_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        DO_BAPI_CUSTOMER_FIND()

    End Sub

    Public Sub DO_BAPI_CUSTOMER_GETLIST()
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim customers As New List(Of String())

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: BAPI_CUSTOMER_GETLIST")

        customers = customer.GetCustomers(rfcDest, customers_search_name.Text)

        LogInfo("Got " & customers.Count.ToString & " entries")

        customers_search_resulttable.Columns.Clear()

        ' Spalten hinzufügen
        customers_search_resulttable.Columns.Add("KUNNR", "No.")
        customers_search_resulttable.Columns.Add("NAME", "Name")

        For Each c In customers

            customers_search_resulttable.Rows.Add(c)

        Next

        customers_search_resulttable.AutoResizeColumns()


        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Public Sub DO_BAPI_CUSTOMER_FIND()


        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim customers As New List(Of String())

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: BAPI_CUSTOMER_GETLIST")

        customers = customer.SearchCustomers(rfcDest, customers_search_name.Text)

        LogInfo("Got " & customers.Count.ToString & " entries")

        customers_search_resulttable.Columns.Clear()

        ' Spalten hinzufügen
        customers_search_resulttable.Columns.Add("KUNNR", "No.")
        customers_search_resulttable.Columns.Add("NAME", "Name")

        For Each c In customers

            customers_search_resulttable.Rows.Add(c)

        Next

        customers_search_resulttable.AutoResizeColumns()


        System.Windows.Forms.Cursor.Current = Cursors.Default



    End Sub

    Private Sub configuration_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles configuration_save.Click

        Dim configPath As String = Application.StartupPath

        Dim output As String = ""

        output &= "sap_appserverhost " & sap_appserverhost.Text & vbNewLine
        output &= "sap_client " & sap_client.Text & vbNewLine
        output &= "sap_language " & sap_language.Text & vbNewLine

        Dim key As String = "SP"
        key = key.PadRight(8, "X")
        Dim encryptedPassword As String = Encrypt(sap_password.Text, key)
        output &= "sap_password " & encryptedPassword & vbNewLine

        output &= "sap_routerstring " & sap_routerstring.Text & vbNewLine
        output &= "sap_systemname " & sap_systemname.Text & vbNewLine
        output &= "sap_systemnumber " & sap_systemnumber.Text & vbNewLine
        output &= "sap_username " & sap_username.Text & vbNewLine

        output &= "path_tempfiles " & path_tempfiles.Text & vbNewLine

        output &= "scripting_visible " & check_sapguiscripting.Checked.ToString & vbNewLine

        output &= "smtp_sendername " & smtp_sendername.Text & vbNewLine
        output &= "smtp_sendermail " & smtp_sendermail.Text & vbNewLine
        output &= "smtp_server " & smtp_server.Text & vbNewLine
        output &= "smtp_portno " & smtp_portno.Text & vbNewLine
        output &= "smtp_username " & smtp_username.Text & vbNewLine

        Dim encryptedSmtpPassword As String = Encrypt(smtp_password.Text, key)
        output &= "smtp_password " & encryptedSmtpPassword & vbNewLine
        output &= "smtp_usessl " & smtp_usessl.Checked.ToString & vbNewLine

        output &= "guinet_licensekey " & guinet_licensekey.Text & vbNewLine


        File.WriteAllText(configPath & "\config.ini", output)

        LogInfo("Configuration saved")


    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Configuration_Load()
    End Sub

    Public Sub Configuration_Load()

        Dim configPath As String = Application.StartupPath
        If Not File.Exists(configPath & "\config.ini") Then
            LogError("Config.ini not found in Application directory")
            Return
        End If

        For Each s As String In File.ReadAllLines(configPath & "\config.ini")

            Dim key = s.Split(" ")(0)
            Dim para = s.Substring(s.IndexOf(" ") + 1)

            Select Case key

                Case "sap_appserverhost"
                    sap_appserverhost.Text = para

                Case "sap_client"
                    sap_client.Text = para

                Case "sap_language"
                    sap_language.Text = para

                Case "sap_password"
                    Dim passphrase As String = "SP"
                    passphrase = passphrase.PadRight(8, "X")
                    sap_password.Text = Decrypt(para, passphrase)

                Case "sap_routerstring"
                    sap_routerstring.Text = para

                Case "sap_systemname"
                    sap_systemname.Text = para

                Case "sap_systemnumber"
                    sap_systemnumber.Text = para

                Case "sap_username"
                    sap_username.Text = para

                Case "path_tempfiles"
                    path_tempfiles.Text = para

                Case "scripting_visible"

                    Select Case para
                        Case "True"
                            check_sapguiscripting.Checked = True

                        Case "False"
                            check_sapguiscripting.Checked = False

                    End Select

                Case "smtp_sendername"
                    smtp_sendername.Text = para

                Case "smtp_sendermail"
                    smtp_sendermail.Text = para

                Case "smtp_server"
                    smtp_server.Text = para

                Case "smtp_portno"
                    smtp_portno.Text = para

                Case "smtp_username"
                    smtp_username.Text = para

                Case "smtp_password"
                    Dim passphrase As String = "SP"
                    passphrase = passphrase.PadRight(8, "X")
                    smtp_password.Text = Decrypt(para, passphrase)

                Case "smtp_usessl"

                    Select Case para
                        Case "True"
                            smtp_usessl.Checked = True

                        Case "False"
                            smtp_usessl.Checked = False

                    End Select

                Case "guinet_licensekey"
                    guinet_licensekey.Text = para

            End Select

        Next

        LogInfo("Settings from config.ini loaded")

    End Sub


#Region "Encryption"


    'The function used to encrypt the text
    Public Shared Function Encrypt(ByVal strText As String, ByVal strEncrKey _
             As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}

        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(strEncrKey)

            Dim des As New DESCryptoServiceProvider()
            Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(strText)
            Dim ms As New MemoryStream()
            Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Return Convert.ToBase64String(ms.ToArray())

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    'The function used to decrypt the text
    Public Shared Function Decrypt(ByVal strText As String, ByVal sDecrKey _
               As String) As String

        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Dim inputByteArray(strText.Length) As Byte

        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(sDecrKey)

            Dim des As New DESCryptoServiceProvider()
            inputByteArray = Convert.FromBase64String(strText)
            Dim ms As New MemoryStream()
            Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)

            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8

            Return encoding.GetString(ms.ToArray())

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

#End Region
#Region "Utility functions"

    Public Sub LogInfo(ByVal text As String)

        statustext.Text &= vbNewLine & "    " & text
        statustext.Select(statustext.TextLength, 0)
        statustext.ScrollToCaret()

    End Sub

    Public Sub LogError(ByVal text As String)

        statustext.Text &= vbNewLine & "    ERROR: " & text
        statustext.Select(statustext.TextLength, 0)
        statustext.ScrollToCaret()

    End Sub

#End Region

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        DO_BAPI_ALM_NOTIF_LIST_FUNCLOC()



    End Sub

    Public Sub DO_BAPI_ALM_NOTIF_LIST_FUNCLOC()

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SetLogonParameters()

        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim messages As New List(Of String)

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: BAPI_ALM_NOTIF_LIST_FUNCLOC")

        messages = customer.GetMessagesToFuncLoc(rfcDest, maintenance_funcloc.Text)

        LogInfo("Got " & messages.Count.ToString & " entries")

        customers_search_resulttable.Columns.Clear()

        ' Spalten hinzufügen
        customers_search_resulttable.Columns.Add("KUNNR", "No.")
        customers_search_resulttable.Columns.Add("NAME", "Name")

        For Each c In messages

            maintenance_text.Text = maintenance_text.Text & vbNewLine & c

        Next



        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Private Sub readservicemsg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles readservicemsg.Click


        SetLogonParameters()

        LogInfo("Logon to SAP GUI")


        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)

            sapgui.Logon("", check_sapguiscripting.Checked)
        End If

        LogInfo("Start transaction IW22 and get data")

        Try
            ' Start transaktion IW22
            sapgui.Enter("/nIW22")

            ' Enter number of service msg
            sapgui.SetField("RIWO00-QMNUM", servicemsg.Text)
            sapgui.Enter()

            servicemsg_status.Text = sapgui.GetField("RIWO00-STTXT")

            sapgui.Enter("=10\TAB02")


        Catch ex As Exception

            LogError(ex.Message)

        End Try
       



        Dim dt As DataTable = New DataTable("IHPA")
        dt.Columns.Add("PARVW")
        dt.Columns.Add("PARNR")

        sapgui.ReadTable(dt)

        DataGridView3.Columns.Clear()

        ' Spalten hinzufügen
        DataGridView3.Columns.Add("PARVW", "Rolle")
        DataGridView3.Columns.Add("PARNR", "Partner")


        ' Zeilen hinzufügen
        For i = 0 To dt.Rows.Count - 1


            DataGridView3.Rows.Add(dt.Rows(i).ItemArray)

        Next

        DataGridView3.AutoResizeColumns()


    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        LogInfo("Reading details to order with SAP GUI Scripting")

        Dim s = read_order_html(vbeln.Text)

        LogInfo("Writing HTML file to path: " & path_tempfiles.Text)
        File.WriteAllText("d:\temp\order_" & vbeln.Text & ".html", s)




    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        Dim tc As New TapiForm

        tc.ShowDialog()




    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim Timer2 As New Timer()
        AddHandler Timer2.Tick, AddressOf Timer2_Tick
        Timer2.Tag = "I am a timer"
        Dim seconds As Integer = 0

        If Integer.TryParse(timerSecondsTextbox.Text, seconds) Then
            Timer2.Interval = seconds * 1000
            Timer2.Start()
        Else
            MsgBox("Invalid input provided")
        End If




    End Sub


    Private Sub servicemsg_status_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles servicemsg_status.TextChanged

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click

        'Create ADO.NET objects.
        Dim myConn As SqlConnection
        Dim myCmd As SqlCommand
        Dim myReader As SqlDataReader
        Dim results As String = ""

        'Create a Connection object.
        myConn = New SqlConnection("Initial Catalog=coffeebar;" & _
                "Data Source=localhost;Integrated Security=SSPI;")

        'Create a Command object.
        myCmd = myConn.CreateCommand
        myCmd.CommandText = "SELECT * FROM strassen WHERE ""strassenname"" LIKE '%anken%'"

        'Open the connection.
        myConn.Open()

        Dim mysw As New Stopwatch
        mysw.Start()
        myReader = myCmd.ExecuteReader()
        mysw.Stop()

        'Concatenate the query result into a string.
        Do While myReader.Read()

            'results = results & myReader.GetString(0) '& vbTab & myReader.GetString(1) & vbLf
            results = myReader.GetString(0) '& vbTab & myReader.GetString(1) & vbLf

        Loop

        'Display results.

        MsgBox("Excecution time: " & mysw.ElapsedMilliseconds.ToString & "ms" & vbNewLine & "Found records: " & results.Count.ToString)

       
        'Close the reader and the database connection.
        myReader.Close()
        myConn.Close()

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click

        Dim ex = TestReadCellsExcel("D:\Excel\cashflow.xlsx", "Miete")


      
    End Sub

    Public Function TestReadCellsExcel(ByVal filename As String, ByVal findThis As String) As String

        Dim returncode = "1"
        Dim oXL As New Excel.Application
        Dim g As New guixt

        Try

            Dim oWB As Excel.Workbook
            Dim oSheet As Excel.Worksheet

            oXL = CreateObject("Excel.Application")
            oXL.Visible = False

            oWB = oXL.Workbooks.Open(filename)
            oSheet = oWB.ActiveSheet

            Dim range As Excel.Range = oSheet.UsedRange

            Dim c As Integer = 1

            For Each r In range.Rows

                If c Mod 500 = 0 Then
                    LogInfo(c.ToString)
                End If
                If Not oSheet.Cells(c, 1).value Is Nothing Then
                    If oSheet.Cells(c, 1).Value.ToString = findThis Then

                        ' We have found a cell containing the value of findThis (in first column!)

                        MsgBox("Found value " & oSheet.Cells(c, 1).Value.ToString & " in row " & c.ToString)
                        returncode = "0"
                        Exit For

                    End If
                End If
                c += 1
            Next

        Catch e As Exception

            MsgBox(e.Message)
            oXL.ActiveWorkbook.Close(False)
            oXL.Quit()
            Return "Exception"

        End Try

        oXL.ActiveWorkbook.Close(False)
        oXL.Quit()

        Return returncode


    End Function

    Public Function TestFillCellsExcel(ByVal rows As String) As Microsoft.Office.Interop.Excel.Application

        Dim oXL As Excel.Application

        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oXL = CreateObject("Excel.Application")
        oXL.Visible = True

        oWB = oXL.Workbooks.Add()
        oSheet = oWB.ActiveSheet
        oSheet.Columns(1).ColumnWidth = 30
        oSheet.Columns(2).ColumnWidth = 20

        For k = 1 To CInt(rows)
            oSheet.Cells(k, 1).Value = "Test"
            oSheet.Cells(k, 2).Value = "Test2"
        Next


        Return oXL

    End Function

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click


        SetLogonParameters()

        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: CLFM_SELECT_AUSP")

        Dim classData = customer.GetClassDataServiceMessage(rfcDest, servicemsg.Text)

        maintenance_text.Text = classData


    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click



        SetLogonParameters()

        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: /GUIXT/SELECT")

        Dim status = customer.GetStatusServicemsg(rfcDest, servicemsg.Text)

        servicemsg_status.Text = status


    End Sub

   




    Private Sub serviceorder_read_guinet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles serviceorder_read_guinet.Click

        SetLogonParameters()

        LogInfo("Logon to SAP GUI")


        'Logon to SAP-GUI
        If sapgui Is Nothing Then
            sapgui = New sapguisession(sapconnect, sapconnectclient, sapconnectuser, sapconnectpassword, sapconnectlanguage)
            sapgui.SetLicense(guinet_licensekey.Text)
            sapgui.Logon("", check_sapguiscripting.Checked)
        End If

        LogInfo("Start transaction IW33 and get data")


        ' Start transaktion IW33
        sapgui.Enter("/nIW33")

        ' Enter number of service order
        sapgui.SetField("CAUFVD-AUFNR", serviceorder_no.Text)
        sapgui.Enter()

        CAUFVD_KTEXT.Text = sapgui.GetField("CAUFVD-KTEXT")
        VIQMEL_AUSVN.Text = sapgui.GetField("VIQMEL-AUSVN")
        VIQMEL_AUZTV.Text = sapgui.GetField("VIQMEL-AUZTV")
        VIQMEL_AUSBS.Text = sapgui.GetField("VIQMEL-AUSBS")
        VIQMEL_AUZTB.Text = sapgui.GetField("VIQMEL-AUZTB")
        CAUFVD_STTXT.Text = sapgui.GetField("CAUFVD-STTXT")
        CAUFVD_IWERK.Text = sapgui.GetField("CAUFVD-IWERK")
        CAUFVD_INGPR.Text = sapgui.GetField("CAUFVD-INGPR")
        CAUFVD_INNAM.Text = sapgui.GetField("CAUFVD-INNAM")
        PMSDO_MATNR.Text = sapgui.GetField("PMSDO-MATNR")
        MAKT_MAKTX.Text = sapgui.GetField("MAKT-MAKTX")



    End Sub

    Private Sub serviceorder_read_bapi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles serviceorder_read_bapi.Click

        SetLogonParameters()

        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: BAPI_ALM_ORDER_GET_DETAIL")

        Dim myServiceOrder As serviceorder = customer.GetClassDataServiceOrder(rfcDest, serviceorder_no.Text)

        CAUFVD_KTEXT.Text = myServiceOrder.CAUFVD_KTEXT
        CAUFVD_STTXT.Text = myServiceOrder.CAUFVD_STTXT
        CAUFVD_INGPR.Text = myServiceOrder.CAUFVD_INGPR
        CAUFVD_IWERK.Text = myServiceOrder.CAUFVD_IWERK
        PMSDO_MATNR.Text = myServiceOrder.PMSDO_MATNR



    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click

        SetLogonParameters()
        LogInfo("Get RFC destination")

        Dim rfcDest As RfcDestination = Nothing
        rfcDest = RfcDestinationManager.GetDestination(parameters)

        Dim customer As New DotNetConnector()

        LogInfo("Call RFC: BAPI_ALM_ORDERHEAD_GET_LIST")

        MsgBox(customer.GetServiceOrders(rfcDest))


    End Sub

 
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click

        WebBrowser1.Navigate("http://www.synactive.com")

    End Sub

   
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click

        TestFillCellsExcel(20)


    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub
End Class
