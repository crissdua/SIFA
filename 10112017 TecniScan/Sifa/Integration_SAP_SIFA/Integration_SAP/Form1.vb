Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports System.Threading
Imports System.Globalization

Public Class Form1
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property
    Private Property _oCompany As SAPbobsCOM.Company

    Public Property oCompany() As SAPbobsCOM.Company
        Get
            Return _oCompany
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _oCompany = value
        End Set
    End Property

    Public Function MakeConnectionSAP() As Boolean
        Dim Connected As Boolean = False
        '' Dim cnnParam As New Settings
        Try
            Connected = -1

            'oCompany = New SAPbobsCOM.Company
            'oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            'oCompany.DbUserName = "sa"
            'oCompany.DbPassword = "12345"
            'oCompany.Server = "DESKTOP-13FMJTF"
            'oCompany.CompanyDB = "FYA"
            'oCompany.UserName = "manager"
            'oCompany.Password = "alegria"
            'oCompany.LicenseServer = "DESKTOP-13FMJTF:30000"
            'oCompany.UseTrusted = False
            'Connected = oCompany.Connect()

            oCompany = New SAPbobsCOM.Company
            oCompany.DbUserName = "USERSAP"
            oCompany.DbPassword = "$apB1T3c"
            oCompany.Server = "192.168.0.180:30015"
            oCompany.CompanyDB = "N_TECNI_SCAN"
            oCompany.UserName = "manager"
            oCompany.Password = "test"
            oCompany.LicenseServer = "192.168.0.180:40000"
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            oCompany.UseTrusted = False
            Connected = oCompany.Connect()

            If Connected <> 0 Then
                Connected = False
                MsgBox(oCompany.GetLastErrorDescription)
            Else
                'MsgBox("Conexión con SAP exitosa")
                Connected = True
            End If
            Return Connected
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Dim MyThread As Thread
    Dim myStream As Stream
    'clientes
    Dim ClientesStart As New ThreadStart(AddressOf BackgroundClientes)
    Dim CallClientes As New MethodInvoker(AddressOf Me.ClienteToma)
    'facturas
    Dim FacturasStart As New ThreadStart(AddressOf BackgroundFacturas)
    Dim CallFacturas As New MethodInvoker(AddressOf Me.FacturasToma)
    'ordenes
    Dim OrdenStart As New ThreadStart(AddressOf BackgroundOrdenes)
    Dim CallOrdenes As New MethodInvoker(AddressOf Me.OrdenesToma)
    'Nota de Credito
    'Dim NcreditoStart As New ThreadStart(AddressOf BackgroundNcredito)
    'Dim CallNcredito As New MethodInvoker(AddressOf Me.NcreditoToma)
    'Paciente
    Dim PacienteStart As New ThreadStart(AddressOf BackgroundPaciente)
    Dim CallPaciente As New MethodInvoker(AddressOf Me.PacienteToma)
    'Pagos
    Dim PagosStart As New ThreadStart(AddressOf BackgroundPagos)
    Dim CallPagos As New MethodInvoker(AddressOf Me.PagosToma)
    'Depositos
    'Dim DepositosStart As New ThreadStart(AddressOf BackgroundDepositos)
    'Dim CallDepositos As New MethodInvoker(AddressOf Me.DepositosToma)
    Dim count As Integer

    Public Sub BackgroundClientes()
        While True
            Me.BeginInvoke(CallClientes)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundFacturas()
        While True
            Me.BeginInvoke(CallFacturas)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundOrdenes()
        While True
            Me.BeginInvoke(CallOrdenes)
            Thread.Sleep(1500)
        End While
    End Sub
    'Public Sub BackgroundNcredito()
    '    While True
    '        Me.BeginInvoke(CallNcredito)
    '        Thread.Sleep(1500)
    '    End While
    'End Sub
    Public Sub BackgroundPaciente()
        While True
            Me.BeginInvoke(CallPaciente)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPagos()
        While True
            Threading.Thread.Sleep(3000)
            Me.BeginInvoke(CallPagos)
            'Thread.Sleep(1500)
        End While
    End Sub
    'Public Sub BackgroundDepositos()
    '    While True
    '        Threading.Thread.Sleep(7000)
    '        Me.BeginInvoke(CallDepositos)
    '    End While
    'End Sub
    Public Function AplicacionFuncionando() As Boolean

        Dim aProceso() As Process
        aProceso = Process.GetProcesses()
        Dim oProceso As Process
        Dim Nombre_Proceso As String
        For Each oProceso In aProceso
            Nombre_Proceso = oProceso.ProcessName.ToString()
            If Nombre_Proceso = "Integration_SAP" Then
                Me.count += 1 'Debes declarar esta variable Global 
            End If
        Next
        If count = 2 Then
            Return 1
        End If
    End Function


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If (AplicacionFuncionando() = True) Then
            MessageBox.Show("La aplicacion ya se encuentra ejecutada")
            System.Windows.Forms.Application.Exit()
        Else
            MakeConnectionSAP()

            If MakeConnectionSAP() Then
                NotifyIcon1.Text = "Integracion Conectada"
            Else
                NotifyIcon1.Text = "Integracion Desconectada, verifique"
            End If
            ' Me.Opacity = 0.0R
            Me.ShowInTaskbar = False
            Me.Visible = False
            'Thread CLIENTES
            MyThread = New Thread(ClientesStart)
            MyThread.IsBackground = True
            MyThread.Name = "MyThreadClientes"
            MyThread.Start()
            'Thread FACTURAS
            MyThread = New Thread(FacturasStart)
            MyThread.IsBackground = True
            MyThread.Name = "MyThreadFacturas"
            MyThread.Start()
            'Thread ORDENES
            MyThread = New Thread(OrdenStart)
            MyThread.IsBackground = True
            MyThread.Name = "MyThreadOrden"
            MyThread.Start()
            'Thread Ncredito
            'MyThread = New Thread(NcreditoStart)
            'MyThread.IsBackground = True
            'MyThread.Name = "MyThreadNcredito"
            'MyThread.Start()
            'Thread Paciente
            MyThread = New Thread(PacienteStart)
            MyThread.IsBackground = True
            MyThread.Name = "MyThreadPaciente"
            MyThread.Start()
            ''Thread Pagos
            MyThread = New Thread(PagosStart)
            MyThread.IsBackground = True
            MyThread.Name = "MyThreadcPagos"
            MyThread.Start()
            ''Thread Depositos
            'MyThread = New Thread(DepositosStart)
            'MyThread.IsBackground = True
            'MyThread.Name = "MyThreadcDepositos"
            'MyThread.Start()
        End If
    End Sub

    Public Sub ClienteToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\IntegracionSAP\Cliente\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeCliente() = 0 Then
                Dim entra As String = "C:\IntegracionSAP\Cliente\" + objFile.Name.ToString
                Dim integ As String = "C:\IntegracionSAP\Cliente\Integration\Cliente.xml"
                'Dim temp As String = "C:\IntegracionSAP\Cliente\temp\Cliente" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Move(entra, integ)
                'File.Move(entra, temp)
            ElseIf existeCliente() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub FacturasToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\IntegracionSAP\Factura\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeFactura() = 0 Then
                Dim entra As String = "C:\IntegracionSAP\Factura\" + objFile.Name.ToString
                Dim integ As String = "C:\IntegracionSAP\Factura\Integration\Invoice.xml"
                'Dim temp As String = "C:\IntegracionSAP\Factura\temp\Invoice" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Move(entra, integ)
                'File.Move(entra, temp)
            ElseIf existeFactura() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub OrdenesToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\IntegracionSAP\Orden\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeOrden() = 0 Then
                Dim entra As String = "C:\IntegracionSAP\Orden\" + objFile.Name.ToString
                Dim integ As String = "C:\IntegracionSAP\Orden\Integration\Order.xml"
                'Dim temp As String = "C:\IntegracionSAP\Orden\temp\Order" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Move(entra, integ)
                'File.Move(entra, temp)
            ElseIf existeOrden() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    'Public Sub NcreditoToma()
    '    Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
    '    Dim objSubFolder As Object = "C:\IntegracionSAP\NotaCredito\"
    '    Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
    '    Dim colFiles As Object = objFolder.Files

    '    For Each objFile In colFiles

    '        If existeNcredito() = 0 Then
    '            Dim entra As String = "C:\IntegracionSAP\NotaCredito\" + objFile.Name.ToString
    '            Dim integ As String = "C:\IntegracionSAP\NotaCredito\Integration\NCredito.xml"
    '            Dim temp As String = "C:\IntegracionSAP\NotaCredito\temp\NCredito" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
    '            File.Copy(entra, integ)
    '            File.Move(entra, temp)
    '        ElseIf existeNcredito() = 1 Then
    '            Exit Sub
    '        End If
    '    Next
    'End Sub
    Public Sub PacienteToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\IntegracionSAP\Paciente\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existePaciente() = 0 Then
                Dim entra As String = "C:\IntegracionSAP\Paciente\" + objFile.Name.ToString
                Dim integ As String = "C:\IntegracionSAP\Paciente\Integration\Paciente.xml"
                'Dim temp As String = "C:\IntegracionSAP\Paciente\temp\Paciente" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Move(entra, integ)
                'File.Move(entra, temp)
            ElseIf existePaciente() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PagosToma()
        Dim Connected As Boolean = False
        '' Dim cnnParam As New Settings
        Try
            Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
            Dim objSubFolder As Object = "C:\IntegracionSAP\Pagos\"
            Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
            Dim counter = My.Computer.FileSystem.GetFiles("C:\IntegracionSAP\Pagos\")
            Dim colFiles As Object = objFolder.Files

            For Each objFile In colFiles

                If existePagos() = 0 Then
                    Dim entra As String = "C:\IntegracionSAP\Pagos\" + objFile.Name.ToString
                    PagoIntegration(entra)
                    Dim integ As String = "C:\IntegracionSAP\Pagos\Integration\Payment.xml"
                    'Dim temp As String = "C:\IntegracionSAP\Pagos\temp\Payment" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                    'File.Copy(entra, temp)
                    File.Move(entra, integ)
                ElseIf existePagos() = 1 Then
                    Exit Sub
                End If
            Next

            'For Each objFile In colFiles

            '    If existePagos() = 0 Then
            '        Dim entra As String = "C:\IntegracionSAP\Pagos\" + objFile.Name.ToString
            '        Dim integ As String = "C:\IntegracionSAP\Pagos\Integration\Payment.xml"
            '        PagoIntegration()
            '        File.Move(entra, integ)

            '        'File.Delete("C:\IntegracionSAP\FacturaUpdate\Integration\Invoice.xml")
            '        'File.Move(integ, temp)
            '    ElseIf existePagos() = 1 Then
            '        Dim entra As String = "C:\IntegracionSAP\Pagos\" + objFile.Name.ToString
            '        Dim integ As String = "C:\IntegracionSAP\Pagos\Integration\Payment.xml"
            '        PagoIntegration()
            '        'File.Delete("C:\IntegracionSAP\FacturaUpdate\Integration\Invoice.xml")
            '        'File.Move(integ, temp)
            '    End If

            'Next
            ''Else
            ''    If existePagos() = 1 Then
            ''        PagoIntegration()
            ''        'File.Delete("C:\IntegracionSAP\FacturaUpdate\Integration\Invoice.xml")
            ''        'File.Move(integ, temp)

            ''    End If
            ''End If
            'Connected = 1
            'Return Connected
        Catch ex As Exception

        End Try
    End Sub
    Public Function DepositosToma() As Boolean
        Dim Connected As Boolean = False
        '' Dim cnnParam As New Settings
        Try
            Connected = -1

            Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
            Dim objSubFolder As Object = "C:\IntegracionSAP\Depositos\"
            Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
            Dim counter = My.Computer.FileSystem.GetFiles("C:\IntegracionSAP\Depositos\")
            Dim colFiles As Object = objFolder.Files
            Dim cant As Integer = counter.Count
            If cant > 0 Then
                For Each objFile In colFiles

                    If existeDepositos() = 0 Then
                        Dim entra As String = "C:\IntegracionSAP\Depositos\" + objFile.Name.ToString
                        Dim integ As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
                        File.Move(entra, integ)
                        DepositoIntegration()
                        'File.Delete("C:\IntegracionSAP\FacturaUpdate\Integration\Invoice.xml")
                        'File.Move(integ, temp)
                    ElseIf existeDepositos() = 1 Then
                        Dim entra As String = "C:\IntegracionSAP\Depositos\" + objFile.Name.ToString
                        Dim integ As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
                        DepositoIntegration()
                        'File.Delete("C:\IntegracionSAP\FacturaUpdate\Integration\Invoice.xml")
                        'File.Move(integ, temp)
                    End If

                Next
            Else
                If existeDepositos() = 1 Then
                    DepositoIntegration()
                    'File.Delete("C:\IntegracionSAP\FacturaUpdate\Integration\Invoice.xml")
                    'File.Move(integ, temp)

                End If
            End If
            Connected = 1
            Return Connected
        Catch ex As Exception

        End Try
    End Function

    Private Function existeFactura()
        If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\Factura\Integration\Invoice.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeCliente()
        If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\Cliente\Integration\Cliente.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeOrden()
        If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\Orden\Integration\Order.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function

    'Private Function existeNcredito()
    '    If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\NotaCredito\Integration\NCredito.xml") Then
    '        'ListBox1.Items.Add("Si Existe")
    '        Return 1
    '    Else
    '        'ListBox1.Items.Add("No Existe")
    '        Return 0
    '    End If
    'End Function
    Private Function existePaciente()
        If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\Paciente\Integration\Paciente.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePagos()
        If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\Pagos\Integration\Payment.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeDepositos()
        If My.Computer.FileSystem.FileExists("C:\IntegracionSAP\Depositos\Integration\Deposito.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function


    Private Function PagoIntegration(archivo As String)
        Dim termina As Integer = 0
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Dim tipodepago As String
        Dim pagoacuenta As String
        Try
            Dim entra As String = archivo
            Dim sale As String = "C:\IntegracionSAP\Pagos\temp\Payment" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("body/document")
            For Each xnDoc As XmlNode In ArticleList

                pagoacuenta = xnDoc.SelectSingleNode("pagoacuenta").InnerText
                If (pagoacuenta = "N") Then
                    If MakeConnectionSAP() Then
                        Dim oPayment As SAPbobsCOM.Payments = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                        Dim CatNodesLines As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                        For Each xnDetLines As XmlNode In CatNodesLines

                            sql = ("select top 1 " & Chr(34) & "DocEntry" & Chr(34) & " from oinv where " & Chr(34) & "U_Serie" & Chr(34) & " = '" + xnDetLines.SelectSingleNode("serie").InnerText + "' and " & Chr(34) & "U_Numero" & Chr(34) & " = " + xnDetLines.SelectSingleNode("docnum").InnerText)
                            Dim oRecordSet As SAPbobsCOM.Recordset
                            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount > 0 Then
                                If xnDetLines IsNot Nothing Then
                                    xnDetLines.ChildNodes(1).InnerText = oRecordSet.Fields.Item(0).Value
                                Else
                                    'Do whatever 
                                End If

                                Xml.Save(entra)
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            oRecordSet = Nothing
                            GC.Collect()
                        Next
                    End If


                End If
            Next
            'File.Copy(entra, sale)
            'Return termina = 1
        Catch ex As Exception
            Dim entra As String = "C:\IntegracionSAP\Pagos\Integration\Payment.xml"
            Dim sale As String = "C:\IntegracionSAP\Pagos\temp\ErrorPayment" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            File.Move(entra, sale)
        End Try
    End Function
    Private Sub DepositoIntegration()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Dim tipodepago As String
        Dim pagoacuenta As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            Dim entra As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
            Dim sale As String = "C:\IntegracionSAP\Depositos\temp\Deposito" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("deposito/document")
            For Each xnDoc As XmlNode In ArticleList
                If MakeConnectionSAP() Then


                    tipodepago = xnDoc.SelectSingleNode("tipodeposito").InnerText
#Region "Mix"
                    If (tipodepago = "") Then
                        Try
                            Dim oService As SAPbobsCOM.CompanyService = oCompany.GetCompanyService()
                            Dim dpService As SAPbobsCOM.DepositsService = oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService)
                            Dim dpsAddCash As SAPbobsCOM.Deposit = dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit)
                            dpsAddCash.DepositType = SAPbobsCOM.BoDepositTypeEnum.dtChecks
                            dpsAddCash.DepositCurrency = xnDoc.SelectSingleNode("doccurrency").InnerText
                            dpsAddCash.Series = xnDoc.SelectSingleNode("series").InnerText
                            Dim Format As String = "yyyyMMdd"
                            Dim fec As DateTime = DateTime.ParseExact(xnDoc.SelectSingleNode("depositdate").InnerText, Format, CultureInfo.CurrentCulture)
                            dpsAddCash.DepositDate = fec.ToString("yyyy-MM-dd")
                            dpsAddCash.Bank = xnDoc.SelectSingleNode("dpsbank").InnerText
                            dpsAddCash.BankBranch = xnDoc.SelectSingleNode("deposbrnch").InnerText
                            dpsAddCash.BankAccountNum = xnDoc.SelectSingleNode("deposacct").InnerText
                            dpsAddCash.BankReference = xnDoc.SelectSingleNode("ref2").InnerText
                            dpsAddCash.DepositorName = xnDoc.SelectSingleNode("dpostorname").InnerText
                            dpsAddCash.JournalRemarks = xnDoc.SelectSingleNode("memo").InnerText
                            dpsAddCash.ReconcileAfterDeposit = SAPbobsCOM.BoYesNoEnum.tYES
                            dpsAddCash.DepositAccountType = SAPbobsCOM.BoDepositAccountTypeEnum.datBankAccount

                            'sql = ("select " & Chr(34) & "AcctCode" & Chr(34) & " from oact where " & Chr(34) & "FormatCode" & Chr(34) & " = " + xnDoc.SelectSingleNode("bankacctt").InnerText)
                            'oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet.DoQuery(sql)
                            'If oRecordSet.RecordCount > 0 Then
                            '    dpsAddCash.DepositAccount = oRecordSet.Fields.Item(0).Value
                            'End If
                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            'oRecordSet = Nothing
                            'GC.Collect()
                            dpsAddCash.DepositAccount = xnDoc.SelectSingleNode("bankacctt").InnerText

                            dpsAddCash.ReconcileAfterDeposit = SAPbobsCOM.BoYesNoEnum.tYES

                            dpsAddCash.AllocationAccount = xnDoc.SelectSingleNode("allocacct").InnerText

                            dpsAddCash.TotalLC = xnDoc.SelectSingleNode("total").InnerText

                            Dim cl As CheckLine = dpsAddCash.Checks.Add()
                            sql = ("select " & Chr(34) & "CheckKey" & Chr(34) & " from ochh where " & Chr(34) & "CheckNum" & Chr(34) & " = " + xnDoc.SelectSingleNode("checkkey").InnerText)
                            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount > 0 Then
                                cl.CheckKey = oRecordSet.Fields.Item(0).Value
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            oRecordSet = Nothing
                            GC.Collect()

                            Dim dpsParamAddCash As SAPbobsCOM.DepositParams = dpService.AddDeposit(dpsAddCash)

                            'MsgBox("Deposito de Cheque Creado correctamente")
                        Catch ex As Exception
                            Dim mystring As String = ex.Message.ToString()
                            Dim NewString As String = Replace(mystring, "[", "")
                            Dim newstring2 As String = Replace(NewString, "]", "")
                            Dim newstring3 As String = Replace(newstring2, ".", "")
                            Dim newstring4 As String = Replace(newstring3, " ", "")
                            Dim newstring5 As String = Replace(newstring4, ":", "")
                            Dim entra2 As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
                            Dim sale2 As String = "C:\IntegracionSAP\Depositos\temp\ErrDeposito" + DateTime.Now.ToString("yyyyMMddHHmmss") + "" + newstring5 + ".xml"
                            File.Move(entra2, sale2)
                        End Try
                    End If

#End Region

#Region "CHQ"
                    If (tipodepago = "CHQ") Then
                        Try
                            Dim oService As SAPbobsCOM.CompanyService = oCompany.GetCompanyService()
                            Dim dpService As SAPbobsCOM.DepositsService = oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService)
                            Dim dpsAddCash As SAPbobsCOM.Deposit = dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit)
                            dpsAddCash.DepositType = SAPbobsCOM.BoDepositTypeEnum.dtChecks
                            dpsAddCash.DepositCurrency = xnDoc.SelectSingleNode("doccurrency").InnerText
                            dpsAddCash.Series = xnDoc.SelectSingleNode("series").InnerText
                            Dim Format As String = "yyyyMMdd"
                            Dim fec As DateTime = DateTime.ParseExact(xnDoc.SelectSingleNode("depositdate").InnerText, Format, CultureInfo.CurrentCulture)
                            dpsAddCash.DepositDate = fec.ToString("yyyy-MM-dd")
                            dpsAddCash.Bank = xnDoc.SelectSingleNode("dpsbank").InnerText
                            dpsAddCash.BankBranch = xnDoc.SelectSingleNode("deposbrnch").InnerText
                            dpsAddCash.BankAccountNum = xnDoc.SelectSingleNode("deposacct").InnerText
                            dpsAddCash.BankReference = xnDoc.SelectSingleNode("ref2").InnerText
                            dpsAddCash.DepositorName = xnDoc.SelectSingleNode("dpostorname").InnerText
                            dpsAddCash.JournalRemarks = xnDoc.SelectSingleNode("memo").InnerText
                            dpsAddCash.ReconcileAfterDeposit = SAPbobsCOM.BoYesNoEnum.tYES
                            dpsAddCash.DepositAccountType = SAPbobsCOM.BoDepositAccountTypeEnum.datBankAccount

                            'sql = ("select " & Chr(34) & "AcctCode" & Chr(34) & " from oact where " & Chr(34) & "FormatCode" & Chr(34) & " = " + xnDoc.SelectSingleNode("bankacctt").InnerText)
                            'oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet.DoQuery(sql)
                            'If oRecordSet.RecordCount > 0 Then
                            '    dpsAddCash.DepositAccount = oRecordSet.Fields.Item(0).Value
                            'End If
                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            'oRecordSet = Nothing
                            'GC.Collect()
                            dpsAddCash.DepositAccount = xnDoc.SelectSingleNode("bankacctt").InnerText
                            dpsAddCash.ReconcileAfterDeposit = SAPbobsCOM.BoYesNoEnum.tYES

                            Dim cl As CheckLine = dpsAddCash.Checks.Add()
                            sql = ("select " & Chr(34) & "CheckKey" & Chr(34) & " from ochh where " & Chr(34) & "CheckNum" & Chr(34) & " = " + xnDoc.SelectSingleNode("checkkey").InnerText)
                            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount > 0 Then
                                cl.CheckKey = oRecordSet.Fields.Item(0).Value
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            oRecordSet = Nothing
                            GC.Collect()

                            Dim dpsParamAddCash As SAPbobsCOM.DepositParams = dpService.AddDeposit(dpsAddCash)

                            'MsgBox("Deposito de Cheque Creado correctamente")
                        Catch ex As Exception
                            Dim mystring As String = ex.Message.ToString()
                            Dim NewString As String = Replace(mystring, "[", "")
                            Dim newstring2 As String = Replace(NewString, "]", "")
                            Dim newstring3 As String = Replace(newstring2, ".", "")
                            Dim newstring4 As String = Replace(newstring3, " ", "")
                            Dim newstring5 As String = Replace(newstring4, ":", "")
                            Dim entra2 As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
                            Dim sale2 As String = "C:\IntegracionSAP\Depositos\temp\ErrDeposito" + DateTime.Now.ToString("yyyyMMddHHmmss") + "" + newstring5 + ".xml"
                            File.Move(entra2, sale2)
                        End Try
                    End If

#End Region
#Region "EF"
                    If (tipodepago = "EF") Then
                        Try
                            Dim oService As SAPbobsCOM.CompanyService = oCompany.GetCompanyService()
                            Dim dpService As SAPbobsCOM.DepositsService = oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService)
                            Dim dpsAddCash As SAPbobsCOM.Deposit = dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit)
                            dpsAddCash.DepositType = BoDepositTypeEnum.dtCash
                            dpsAddCash.DepositCurrency = xnDoc.SelectSingleNode("doccurrency").InnerText
                            dpsAddCash.Series = xnDoc.SelectSingleNode("series").InnerText
                            Dim Format As String = "yyyyMMdd"
                            Dim fec As DateTime = DateTime.ParseExact(xnDoc.SelectSingleNode("depositdate").InnerText, Format, CultureInfo.CurrentCulture)
                            dpsAddCash.DepositDate = fec.ToString("yyyy-MM-dd")
                            dpsAddCash.Bank = xnDoc.SelectSingleNode("dpsbank").InnerText
                            dpsAddCash.BankBranch = xnDoc.SelectSingleNode("deposbrnch").InnerText
                            dpsAddCash.BankAccountNum = xnDoc.SelectSingleNode("deposacct").InnerText
                            dpsAddCash.BankReference = xnDoc.SelectSingleNode("ref2").InnerText
                            dpsAddCash.DepositorName = xnDoc.SelectSingleNode("dpostorname").InnerText
                            dpsAddCash.JournalRemarks = xnDoc.SelectSingleNode("memo").InnerText

                            'sql = ("select ifnull(" & Chr(34) & "AcctCode" & Chr(34) & ",0) from oact where " & Chr(34) & "FormatCode" & Chr(34) & " = " + xnDoc.SelectSingleNode("bankacctt").InnerText)
                            'oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet.DoQuery(sql)
                            'If oRecordSet.RecordCount > 0 Then
                            '    dpsAddCash.DepositAccount = oRecordSet.Fields.Item(0).Value
                            'End If
                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            'oRecordSet = Nothing
                            'GC.Collect()
                            '-------------------------------------------
                            dpsAddCash.DepositAccount = xnDoc.SelectSingleNode("bankacctt").InnerText


                            'sql = ("select " & Chr(34) & "AcctCode" & Chr(34) & " from oact where " & Chr(34) & "FormatCode" & Chr(34) & " = " + xnDoc.SelectSingleNode("allocacct").InnerText)
                            'oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet.DoQuery(sql)
                            'If oRecordSet.RecordCount > 0 Then
                            '    dpsAddCash.AllocationAccount = oRecordSet.Fields.Item(0).Value
                            'End If
                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            'oRecordSet = Nothing
                            'GC.Collect()
                            '------------------------------------------------------------
                            dpsAddCash.AllocationAccount = xnDoc.SelectSingleNode("allocacct").InnerText


                            dpsAddCash.TotalLC = xnDoc.SelectSingleNode("total").InnerText

                            Dim dpsParamAddCash As SAPbobsCOM.DepositParams = dpService.AddDeposit(dpsAddCash)
                            'MsgBox("Deposito de Cheque Creado correctamente")

                        Catch ex As Exception
                            Dim mystring As String = ex.Message.ToString()
                            Dim NewString As String = Replace(mystring, "[", "")
                            Dim newstring2 As String = Replace(NewString, "]", "")
                            Dim newstring3 As String = Replace(newstring2, ".", "")
                            Dim newstring4 As String = Replace(newstring3, " ", "")
                            Dim newstring5 As String = Replace(newstring4, ":", "")
                            Dim entra2 As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
                            Dim sale2 As String = "C:\IntegracionSAP\Depositos\temp\ErrDeposito" + DateTime.Now.ToString("yyyyMMddHHmmss") + "" + newstring5 + ".xml"
                            File.Move(entra2, sale2)
                        End Try
                    End If
#End Region

                End If
            Next
            File.Move(entra, sale)
        Catch ex As Exception
            Dim entra As String = "C:\IntegracionSAP\Depositos\Integration\Deposito.xml"
            Dim sale As String = "C:\IntegracionSAP\Depositos\temp\ErrDeposito" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            File.Move(entra, sale)
        End Try
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.ShowInTaskbar = True
        Me.Visible = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.ShowInTaskbar = False
        Me.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Dim Msg As MsgBoxResult
        Msg = MsgBox("Desea Cerrar la Integracion, ¿Desea salir?", vbYesNo, "Salir de la Integracion")
        If Msg = MsgBoxResult.Yes Then
            System.Windows.Forms.Application.Exit()
        Else
            Exit Sub
        End If
    End Sub
End Class
