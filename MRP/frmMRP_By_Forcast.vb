Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.DirectoryServices
Imports System.Text
Imports System.Xml

Public Class frmMRP_By_Forcast
    'Dim strCnn As String = "Server=SHPLAPSIS01\SQLEXPRESS2012; Database=SEA; User ID=sa;Password=Fernanda25"
    Dim strCnn As String = "Server=10.17.182.12\SQLEXPRESS2012;Database=SEA;User ID=sa;Password=SHPadmin14%"
    'Dim strCnn As String = "Server=10.17.182.36\SQLEXPRESS2012;Database=SEA;User ID=sa;Password=SHPadmin14%"
    'Dim strCnn As String = "Server=BMXLAPSIS06\SQLEXPRESS2017;Database=SEA;User ID=sa;Password=SHPadmin14%"
    Dim cnn As New SqlConnection(strCnn)
    Public TablaExcelForcast As New Data.DataTable
    Private tblHojasDeCalculo As New Data.DataTable
    Private TablaExcel As New Data.DataTable
    Dim ArchivoX As String
    Dim RutaX As String
    Private ConIssues As Integer
    Private ConWarning As Integer
    Dim StartDateProces, DueDateProcess, DueDateAssy, DueDateShipped, PackDueDate As String
    Dim sTempTableName As String 'variable para el nombre de las tablas
    Dim tblAUBOMWIPForecastreference As New Data.DataTable
    Dim tblRevBOMWIPForecastreference As New Data.DataTable
    Dim tblWIPBOMWIPForecastreference As New Data.DataTable
    Dim tblRevWipByAUForecastreference As New Data.DataTable
    Dim tblAUWIPForecastreference As New Data.DataTable
    Dim tblAUBOMWIP As New Data.DataTable
    Dim tblAUBOMENG As New Data.DataTable
    Dim tblRevBOMWIP As New Data.DataTable
    Dim tblRevBOMENG As New Data.DataTable
    Dim tblWIPBOMWIP As New Data.DataTable
    Dim tblPNMyTable As New Data.DataTable
    Dim tblRevSalesOrder As New Data.DataTable
    Dim tblRevWipByAU As New Data.DataTable
    Dim BanderaLogin As Integer
    Public tblPerWeek As New Data.DataTable
    Public tblPerVendor As New Data.DataTable
    Private tblWarning As New Data.DataTable
    Private tblIssues As New Data.DataTable
    Dim LaFecha As String
    Dim FechaInicial, FechaUltima As Date
    Dim OpcionCorrecta As String
    Dim Renglon As Long
    Dim iTopRow As Integer
    Dim SubPN As String
    Dim IDx As Long
    Dim PNBusquedaWip As String
    Dim FirstDayWeekBusquedaWip As String
    Dim DatoX As String

    Dim mesesEspEn As New Dictionary(Of String, String) From {
    {"ene", "Jan"}, {"feb", "Feb"}, {"mar", "Mar"},
    {"abr", "Apr"}, {"may", "May"}, {"jun", "Jun"},
    {"jul", "Jul"}, {"ago", "Aug"}, {"sep", "Sep"},
    {"oct", "Oct"}, {"nov", "Nov"}, {"dic", "Dec"}
    }

    Private Sub frmMRP_By_Forcast_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If BanderaLogin > 0 Then
            TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
            'TruncateTablaTemp("tblPurchasingTempMRP2" + sTempTableName)
            If sTempTableName <> "" Then
                DroptblPurchasingTempMRPTable(sTempTableName)
            End If
        End If
    End Sub
    Private Sub frmMRP_By_Forcast_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        ExchangeRate()
        BanderaLogin = 0
        'Dim X As String = GeneraSerialBOMWipFake("", "tblPurchasingWipFake")
        'X = GeneraSerialBOMWipFake(X, "tblPurchasingWipFake")
        'X = GeneraSerialBOMWipFake(X, "tblPurchasingWipFake")
        'X = GeneraSerialBOMWipFake(X, "tblPurchasingWipFake")
        'X = GeneraSerialBOMWipFake(X, "tblPurchasingWipFake")
        ''  ZAS = Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + "000000000"
        'X = "FW20170812A999999999999" 'FW20170812A000000000001
        'X = GeneraSerialBOMWipFake(X, "tblPurchasingWipFake")
        'X = "FW20170812Z999999999999"
        'X = GeneraSerialBOMWipFake(X, "tblPurchasingWipFake")
        'X = GeneraSerialForecastReference("", "tblCustomerServiceForecast")
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = "CFAAAA0099"
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = "CFAAAA0999"
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = "CFAAAA9999"
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        'X = "CFZZZZ9999"
        'X = GeneraSerialForecastReference(X, "tblCustomerServiceForecast")
        GeneraColumnasHojasDeExcel()
        'CreaTablaExcelGrid()
        'GroupWipSalesOrder         6,20
        'GroupBoxBudgetInformation  6,20
        'GroupBoxPurchasingOrderHistory 6,20
        'txbUserMRP.Text = "julio.gallegos"
        'txbUserMRPPassword.Text = "Fernanda25"
        GeneraColumnasTablasBOM()
        GeneraColumnas()
        lblTotal.Text = ""
        rdoViewOnly.Checked = True
        lblWeekFrom.Text = ""
        lblWeekTo.Text = ""
        'txbUser.Text = "Mario.Espinoza"
        dtpFrom.Value = Now.AddYears(-20).ToShortDateString
        dtpTo.Value = Now.AddYears(20).ToShortDateString
        lblWeekFrom.Text = Semanas(dtpFrom.Value)
        lblWeekTo.Text = Semanas(dtpTo.Value)
        rdoAllWeeks.Checked = True
        If GridMRP.Rows.Count > 0 Then
            cmbFilter.Enabled = True
            cmb10Percent.Enabled = True
        Else
            cmbFilter.Enabled = False
            cmb10Percent.Enabled = False
        End If
        cmbFilter.SelectedIndex = 0
        cmb10Percent.SelectedIndex = 0
        'CargaPOs()
        lblQty.Text = "Qty:"
        txbUserMRP.Text = Environment.UserName
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    'Funcion que regresa el tipo de cambio
    Public Sub ExchangeRate()
        Dim exchangeBMX As String = ""
        Try
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim documentoxml As New XmlDataDocument
            documentoxml.Load("http://www.banxico.org.mx/rsscb/rss?BMXC_canal=fix&BMXC_idioma=es")
            exchangeBMX = documentoxml.ChildNodes(1).ChildNodes(1).ChildNodes(8).ChildNodes(2).ChildNodes(0).InnerText
            If IsNumeric(CDec(Val(exchangeBMX))) Then
                txbExchangeRate.Text = exchangeBMX
            Else
                txbExchangeRate.Text = "19"
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub btnOpenFileForecast_Click(sender As Object, e As EventArgs) Handles btnOpenFileForecast.Click
        OpenFileDialogForecast.Title = "Select your Forecast File"
        OpenFileDialogForecast.Filter = "Excel File | *.XLSX"
        If OpenFileDialogForecast.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim Ruta As String = OpenFileDialogForecast.FileName
            RutaX = Ruta
            'GroupBoxUploadFile.Visible = True
            'GroupBoxUploadFile.Visible = False
            Dim Archivo As String = OpenFileDialogForecast.SafeFileName
            ArchivoX = Archivo
            Dim Longitud As Integer = Len(Archivo) - 4
            If Longitud > 1 Then
                BuscaHojasDeCalculo(Ruta, Archivo)
            End If
            If tblHojasDeCalculo.Rows.Count > 0 Then
                MSExcelMuestra(RutaX, cmbHojasDeCalculo.Text)
            End If
        End If
    End Sub
    Private Sub MSExcelMuestra(ByVal RutaArchivillo As String, ByVal HojaDeCalculo As String)
        Using TablaDet As New Data.DataTable("tblExcelX")
            Try
                'Private cnnExcel As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\julio.gallegos\Documents\SHPFiles\MasterCC.xlsx; Extended Properties=""Excel 12.0;HDR=YES""")    'Office 2007
                'Dim cnnExcel As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.17.0; Data Source=" + RutaArchivillo + "; Extended Properties=""Excel 12.0;HDR=YES""")    'Office 2007
                Dim cnnExcel As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & RutaArchivillo & "; Extended Properties=""Excel 12.0;HDR=YES""")

                Dim adExcel As New OleDbDataAdapter("SELECT * FROM [" + HojaDeCalculo + "$]", cnnExcel)
                Dim TablaExcelX As New Data.DataTable 'guardar el archivo de excel
                cnnExcel.Open()
                adExcel.Fill(TablaExcelX)
                Dim cmdExcel As New OleDb.OleDbCommandBuilder(adExcel)
                cnnExcel.Close()
                GridExcelForecast.DataSource = TablaExcelX
                GridExcelForecast.AutoResizeColumns()
                Dim Count As Long = TablaExcelX.Rows.Count
                lblRecordsExcelForcast.Text = "Records: " + Count.ToString
            Catch ex As Exception
                MsgBox(ex.ToString + vbNewLine + "The process can't be compleated please verify the file is close or check if the file was saved in the correct path", MsgBoxStyle.Critical, "Issues")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
    End Sub

    Private Sub btnLoginMRP_Click(sender As Object, e As EventArgs) Handles btnLoginMRP.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        LoginMRP()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GeneraColumnasHojasDeExcel()
        Try
            Dim workCol1 As DataColumn = tblHojasDeCalculo.Columns.Add("HojasDeCalculo", Type.GetType("System.String"))
        Catch ex As Exception
            MessageBox.Show(ex.ToString + vbNewLine + "Genera Columnas", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BuscaHojasDeCalculo(ByVal Ruta As String, ByVal Archivo As String)
        '
        Try
            tblHojasDeCalculo.Clear()
            Dim Aux As String
            Dim ObjExcel As Excel.Application = New Excel.Application
            Dim ObjW As Excel.Workbook = ObjExcel.Workbooks.Open(Ruta)
            'ObjExcel= New Excel.Application
            'ObjW = ObjExcel.Workbooks.Open(Ruta)
            Dim i As Integer
            'For Each sheet As Excel.Worksheet In ObjW.Worksheets
            For i = 1 To ObjW.Sheets.Count
                Dim objHojaExcel As Excel.Worksheet = CType(ObjW.Worksheets(i), Worksheet)
                'Dim A1 As DataRow = tblHojasDeCalculo.NewRow
                Aux = CStr(objHojaExcel.Name)
                'A1("HojasDeCalculo") = ObjExcel.ObjW.Sheets(i)
                tblHojasDeCalculo.Rows.Add(Aux)
            Next
            ObjExcel.DisplayAlerts = False
            ObjW.Close()
            ObjW = Nothing
            ObjExcel.Quit()
            ObjExcel = Nothing
            'releaseObject(objHojaExcel)
            releaseObject(ObjW)
            releaseObject(ObjExcel)
            With cmbHojasDeCalculo
                .DataSource = tblHojasDeCalculo
                .DisplayMember = "HojasDeCalculo"
                .ValueMember = "HojasDeCalculo"
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub MSExcelConecta(ByVal Archivo As String, ByVal Ruta As String, ByVal Hoja As String, ByVal ForecastReference As String)
        Try
            Dim Cont As Long = 0
            Dim MensarWarning As String = ""
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(Ruta)
            Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets(Hoja)
            With xlApp
                Try
                    'variables del WIP
                    'WIP, AU, Rev, PN, Qty, CreatedDate, StartDateProces, DueDateProcess, DueDateAssy, DueDateShipped, DueDateCustomer, Customer, IT, Notes, CreatedBy, KindOfAU, Family, WeekProcess, Line, ForecastReference
                    Dim WIP As String
                    Dim Week As Integer
                    'ID, ItemNumber, ItemDescription, ItemLongDescription, UOM, SupplierItemNumber, ShipTo, ShipToDescription, BlanketNumber, RowType, PastDue,Qty, DueDate, Week, AU, Rev, FileNameForecast, CreatedBy, CreatedDate, Path, ForecastReference
                    Dim BanderaPPAP, ItemNumberCol, ItemDescriptionCol, ItemLongDescriptionCol, UOMCol, SupplierItemNumberCol, ShipToCol, ShipToDescriptionCol, BlanketNumberCol, RowTypeCol, PastDueCol As Integer 'DueDateStartCol, DueDateEndCol
                    Dim PPAP, ID, ItemNumber, ItemDescription, ItemLongDescription, UOM, SupplierItemNumber, ShipTo, ShipToDescription, BlanketNumber, RowType, PastDue, DueDate, Rev As String 'FileNameForecast, Path
                    Dim AU As Long
                    Dim Qty As Decimal = 0, GeneraWip As Integer
                    Dim Contador As Integer = 0
                    Dim StartDatesCol As Integer = 0
                    Dim LastDatesCol As Integer = 0
                    Dim lastRowIndex As Integer
                    'Dim currRowIndex As Integer
                    Dim LastColuIndex As Integer
                    'Dim currRow As Range
                    Dim kk As String
                    Dim valueD As Date
                    Dim FechasForecast(40) As String
                    Dim FechasContador As Integer
                    With xlWorkSheet
                        LastColuIndex = .Cells(1, .Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                        lastRowIndex = xlWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row
                        For NM As Integer = 1 To LastColuIndex

                            kk = CType(.Cells(1, NM), Range).Value.ToString()    'revisar aqui
                            If IsDate(kk) Then
                                valueD = CDate(kk)
                                FechasContador += 1
                                If StartDatesCol = 0 Then 'identificamo cuantas columnas de fechas hay
                                    StartDatesCol = NM
                                Else
                                    LastDatesCol = NM
                                End If
                            End If
                            Select Case kk ' identificamos cuales son cada columna
                                Case "Material Number"
                                    ItemNumberCol = NM
                                Case "Material Description"
                                    ItemDescriptionCol = NM
                                Case "Material Long Description"
                                    ItemLongDescriptionCol = NM
                                Case "Unit"
                                    UOMCol = NM
                                Case "Supplier Item Number"
                                    SupplierItemNumberCol = NM
                                Case "Ship to"
                                    ShipToCol = NM
                                Case "Ship to Description"
                                    ShipToDescriptionCol = NM
                                Case "Blanket Number"
                                    BlanketNumberCol = NM
                                Case "Row Type"
                                    RowTypeCol = NM
                                Case "Past Due"
                                    PastDueCol = NM
                            End Select
                        Next
                        For JK As Integer = 1 To lastRowIndex - 1 'obtiene los valores de cada renglone
                            Console.WriteLine(JK)
                            Cont = JK
                            If JK = lastRowIndex - 1 Then
                                JK = JK
                            End If
                            If JK = 90 Then
                                JK = JK
                            End If
                            BanderaPPAP = 0
                            PPAP = ""
                            AU = 0
                            Rev = ""
                            StartDateProces = ""
                            DueDateProcess = ""
                            PackDueDate = ""
                            DueDateAssy = ""
                            DueDateShipped = ""
                            ItemDescription = ""
                            UOM = ""
                            ShipTo = ""
                            ShipToDescription = ""
                            RowType = ""
                            PastDue = "0"
                            'Alx - Error en espacios en blanco
                            Try
                                ItemNumber = CStr(CType(.Cells(JK, ItemNumberCol), Range).Value.ToString())
                                ItemDescription = CType(.Cells(JK, ItemDescriptionCol), Range).Value.ToString()
                                UOM = CType(.Cells(JK, UOMCol), Range).Value.ToString()
                                'ShipTo = CType(.Cells(JK, ShipToCol), Range).Value.ToString()
                                'ShipToDescription = CType(.Cells(JK, ShipToDescriptionCol), Range).Value.ToString()
                                RowType = CType(.Cells(JK, RowTypeCol), Range).Value.ToString()
                                PastDue = CType(.Cells(JK, PastDueCol), Range).Value.ToString()
                            Catch ex As Exception
                                'MsgBox(ex.ToString)
                            End Try

                            ItemLongDescription = "" ' CType(.Cells(JK, ItemLongDescriptionCol), Range).Value.ToString()
                            SupplierItemNumber = ""
                            BlanketNumber = ""
                            AU = BuscaAU(ItemNumber) 'busca AU en tblMaster
                            If AU > 0 Then
                                Rev = BuscaRev(ItemNumber) 'busca Rev en tblMaster
                                PPAP = BuscaPPAP(ItemNumber)  'Busque si es PPAP en tbl
                            End If
                            For WW As Integer = StartDatesCol To LastDatesCol
                                If WW = LastDatesCol - 1 Then
                                    WW = WW
                                End If
                                Qty = 0
                                DueDate = CDate(CType(.Cells(1, WW), Range).Value.ToString()).ToString("dd/MMM/yyyy")
                                Week = Semanas(DueDate)
                                Qty = CInt(Val(CType(.Cells(JK, WW), Range).Value)).ToString()
                                If Not IsDate(DueDate) Then
                                    MessageBox.Show("Por favor revise el ItemNumber " + ItemNumber, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End If
                                If Qty > 0 And IsDate(DueDate) Then
                                    Contador += 1
                                    Qty = Qty
                                    If BanderaPPAP = 0 And PPAP <> "" Then
                                        ConWarning += 1
                                        MensarWarning = "El AU " + AU.ToString + " " + PPAP + " es PPAP hay requerimiento para la fecha " + DueDate
                                        Dim Q As DataRow = tblWarning.NewRow
                                        Q.Item("Warnings") = MensarWarning
                                        Q.Item("ItemRow") = ConWarning
                                        tblWarning.Rows.Add(Q)
                                    End If
                                    'Hacemos el insert
                                    ID = BuscaNumeroDeReferencia("tblCustomerServiceForecast", "ID")
                                    ID = GeneraSerialBOMWipFake(ID, "tblCustomerServiceForecast")
                                    InsertForecast(ID, ItemNumber.TrimStart.TrimEnd, ItemDescription.TrimStart.TrimEnd, ItemLongDescription.TrimStart.TrimEnd, UOM.TrimStart.TrimEnd, SupplierItemNumber.TrimStart.TrimEnd, ShipTo.TrimStart.TrimEnd, ShipToDescription.TrimStart.TrimEnd, BlanketNumber.TrimStart.TrimEnd, RowType.TrimStart.TrimEnd, PastDue, Qty, DueDate.TrimStart.TrimEnd, Week, AU, Rev, Archivo, Ruta, ForecastReference, Contador, JK)
                                    GeneraWip = RevisaAU(ItemNumber)
                                    If GeneraWip = 0 Then
                                        If AU = 2731 Or AU = 3179 Or AU = 3598 Or AU = 3179 Then
                                            AU = AU
                                        End If
                                        'Generar un WIP falso
                                        LeadTime(AU, Rev, DueDate)
                                        WIP = WIPData(AU, Rev, ItemNumber, ForecastReference, DueDate, Qty)
                                        'Genera BOM falso
                                        If WIP <> "" Then GeneraBOM(AU, Rev, WIP, ForecastReference, DueDate, Qty)
                                    End If
                                End If
                            Next
                        Next
                    End With
                Catch ex As Exception
                    MsgBox(ex.ToString + vbNewLine + "Renglon " + CStr(Cont), MsgBoxStyle.Critical, "Issues")
                End Try
                '~~> Close workbook and quit Excel
                xlWorkBook.Close(False)
                xlApp.Quit()
                '~~> Clean Up
                releaseObject(xlWorkSheet)
                releaseObject(xlWorkBook)
                releaseObject(xlApp)
            End With
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '
    Private Sub releaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                If Runtime.InteropServices.Marshal.IsComObject(obj) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                    obj = Nothing
                End If
            End If


        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    'Llena los datos del archivo de excel en la tablaExcel
    Private Sub LLenaTablaExcel()
        Dim cadena As String = ""
        Dim str As String = ""
        Dim contador As Integer = 0
        Dim bandera As String = "OK"
        Dim i As Integer = 1
        For SW As Integer = 0 To TablaExcelForcast.Rows.Count - 1 'este ciclo es para validar hasta que renglon hay informacion
            cadena = TablaExcelForcast.Rows(SW).Item("AU").ToString.ToUpper
            'For i = 1 To Len(cadena)
            '    str = Microsoft.VisualBasic.Mid(cadena, i, 1)
            '    If str = " " Then
            '        contador += 1
            '    End If
            'Next
            'i = Len(cadena)
            'If i = contador Then
            '    bandera = "NO"
            'End If
            'If cadena = "" Then
            '    bandera = "NO"
            'End If
            If bandera = "OK" Then
                'NewRow()
                'TablaExcel.Rows(SW).Item("AU") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                ''cadena = TablaExcel.Rows(SW).Item("Au")
                'cadena = TablaExcelForcast.Rows(SW).Item("Rev").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("Rev") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("From").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("From") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("To").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("To") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("wid").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("wid") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("wire").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("wire") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("Length").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("Length") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("TermA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("TermA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("StripA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("StripA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("JoinA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("JoinA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("SpA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("SpA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("CoverA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("CoverA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("TermB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("TermB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("StripB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("StripB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("JoinB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("JoinB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("SpB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("SpB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("CoverB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("CoverB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("Ink").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("Ink") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("Notas").ToString
                'TablaExcel.Rows(SW).Item("Notas") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("EPA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("EPA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("CP").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("CP") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("EPB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("EPB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("OperA").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("WDevA") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
                'cadena = TablaExcelForcast.Rows(SW).Item("OperB").ToString.ToUpper
                'TablaExcel.Rows(SW).Item("WDevB") = LTrim(RTrim(DelApostrofe(cadena).TrimEnd.TrimStart())).ToUpper
            End If
        Next
        'BanderaTblExcel += 1
    End Sub
    'Funcion para eliminar los apostrofes 
    Private Function DelApostrofe(ByVal NewString As String) As String
        Dim p As Integer
        Dim tope As Integer
        Dim caracter As Char
        tope = (Len(NewString))
        For p = 1 To tope
            caracter = Microsoft.VisualBasic.Mid(NewString, p)
            If (caracter = "'") Then
                caracter = "/"
                Mid(NewString, p) = caracter
            End If
        Next
        Return NewString
    End Function

    Private Sub cmbHojasDeCalculo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHojasDeCalculo.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbHojasDeCalculo.SelectedIndex > -1 Then
            If cmbHojasDeCalculo.SelectedValue.ToString <> "System.Data.DataRowView" Then
                MSExcelMuestra(RutaX, cmbHojasDeCalculo.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnStartProcess_Click(sender As Object, e As EventArgs) Handles btnStartProcess.Click
        Cursor.Current = Cursors.WaitCursor
        ConIssues = 0
        ConWarning = 0
        tblIssues.Clear()
        tblWarning.Clear()
        Dim ForecastReference As String = BuscaForecastReference()
        ForecastReference = GeneraSerialForecastReference(ForecastReference, "tblCustomerServiceForecast")
        lblForecastReference.Text = ForecastReference
        lblMRPReference.Text = ForecastReference
        MSExcelConecta(ArchivoX, RutaX, cmbHojasDeCalculo.Text, ForecastReference)
        If (tblIssues.Rows.Count > 0 Or tblWarning.Rows.Count > 0) Then EnviaCorreo()
        '
        Dim IDReferenceMRP As String = BuscaNumeroDeReferenciaMRP()
        IDReferenceMRP = GeneraSerialMRP(IDReferenceMRP)
        lblMRPReference.Text = ForecastReference ' IDReferenceMRP
        BuscaPNsPrimarios()
        Dim Opcion As String = "All"
        Dim FechaInicio As String = "2000/01/01"
        Dim FechaFin As String = Now.AddYears(20)
        lblForecastReference.Text = lblForecastReference.Text ' "CFAAAA0001"
        TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
        CalculaMateriales(Opcion, FechaInicio, FechaFin, IDReferenceMRP, lblForecastReference.Text)
        GroupBoxUploadFile.Visible = False
        TabControlMRPGlobal.Visible = True
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub InsertForecast(ByVal ID As String, ByVal ItemNumber As String, ByVal ItemDescription As String, ByVal ItemLongDescription As String, ByVal UOM As String, ByVal SupplierItemNumber As String, ByVal ShipTo As String, ByVal ShipToDescription As String, ByVal BlanketNumber As String, ByVal RowType As String, ByVal PastDue As String, ByVal Qty As Integer, ByVal DueDate As String, ByVal Week As String, ByVal AU As String, ByVal Rev As String, ByVal FileNameForecast As String, ByVal Path As String, ByVal ForecastReference As String, ByVal ItemRow As Integer, ByVal ItemExcelRow As Integer)
        'ID, ItemNumber, ItemDescription, ItemLongDescription, UOM, SupplierItemNumber, ShipTo, ShipToDescription, BlanketNumber, RowType, PastDue,Qty, DueDate, Week, AU, Rev, FileNameForecast, CreatedBy, CreatedDate, Path, ForecastReference
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = "INSERT INTO tblCustomerServiceForecast (ID, ItemNumber, ItemDescription, ItemLongDescription, UOM, SupplierItemNumber, ShipTo, ShipToDescription, BlanketNumber, RowType, PastDue, Qty, DueDate, Week, AU, Rev, FileNameForecast, CreatedBy, CreatedDate, Path, ForecastReference, ItemRow, ItemExcelRow) VALUES (@ID, @ItemNumber, @ItemDescription, @ItemLongDescription, @UOM, @SupplierItemNumber, @ShipTo, @ShipToDescription, @BlanketNumber, @RowType, @PastDue, @Qty, @DueDate, @Week, @AU, @Rev, @FileNameForecast, @CreatedBy, @CreatedDate, @Path, @ForecastReference, @ItemRow, @ItemExcelRow)"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@ID", SqlDbType.NVarChar).Value = ID
            cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = Qty
            cmd.Parameters.Add("@UOM", SqlDbType.NVarChar).Value = UOM
            cmd.Parameters.Add("@ItemRow", SqlDbType.Int).Value = ItemRow
            cmd.Parameters.Add("@ItemExcelRow", SqlDbType.Int).Value = ItemExcelRow
            cmd.Parameters.Add("@ItemNumber", SqlDbType.NVarChar).Value = ItemNumber
            cmd.Parameters.Add("@ItemDescription", SqlDbType.NVarChar).Value = ItemDescription
            cmd.Parameters.Add("@ItemLongDescription", SqlDbType.NVarChar).Value = ItemLongDescription
            cmd.Parameters.Add("@SupplierItemNumber", SqlDbType.NVarChar).Value = SupplierItemNumber
            cmd.Parameters.Add("@ShipTo", SqlDbType.NVarChar).Value = ShipTo
            cmd.Parameters.Add("@ShipToDescription", SqlDbType.NVarChar).Value = ShipToDescription
            cmd.Parameters.Add("@BlanketNumber", SqlDbType.NVarChar).Value = BlanketNumber
            cmd.Parameters.Add("@RowType", SqlDbType.NVarChar).Value = RowType
            cmd.Parameters.Add("@PastDue", SqlDbType.Int).Value = CInt(Val(PastDue))
            cmd.Parameters.Add("@DueDate", SqlDbType.Date).Value = DueDate
            cmd.Parameters.Add("@Week", SqlDbType.Int).Value = Week
            cmd.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
            cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
            cmd.Parameters.Add("@FileNameForecast", SqlDbType.NVarChar).Value = FileNameForecast
            cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = txbUser.Text
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = Now
            cmd.Parameters.Add("@Path", SqlDbType.NVarChar).Value = Path
            cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error en el insert de tblCustomerServiceForecast" + " " + ItemNumber, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
            Console.WriteLine(ex.ToString())
        End Try
    End Sub

    Private Function RevisaAU(ByVal PN As String)
        Dim Edo As String = ""
        Dim Resp As Long = 0
        Using TN As New Data.DataTable 'tabla para el master
            Dim MensajeIssues As String = ""
            Dim MensarWarning As String = ""
            Using TL As New Data.DataTable 'tabla para el leadtime
                Using TB As New Data.DataTable 'tabla para el BOM
                    Dim AU As Long = 0
                    Dim Rev As String = ""
                    Dim Active As String = "False"
                    Dim KindOFAU As String = ""
                    Try
                        Dim cmd As SqlCommand
                        Dim dr As SqlDataReader
                        ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                        Dim Query As String = "SELECT * FROM tblMaster WHERE PN=@PN ORDER BY Active DESC"
                        cmd = New SqlCommand(Query, cnn)
                        cmd.CommandType = CommandType.Text
                        cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                        cnn.Open()
                        dr = cmd.ExecuteReader
                        TN.Load(dr)
                        cnn.Close()
                        If TN.Rows.Count > 0 Then
                            AU = CLng(Val(TN.Rows(0).Item("AU").ToString))
                            Rev = TN.Rows(0).Item("Rev").ToString
                            Active = TN.Rows(0).Item("Active").ToString
                            KindOFAU = TN.Rows(0).Item("KindOFAU").ToString
                            If AU = 2387 Then
                                AU = AU
                            End If
                            'revisamos los leadtime
                            Try
                                Dim Query2 As String = "SELECT * FROM tblMasterProcessLeadTime WHERE AU=@AU AND Rev=@Rev"
                                Dim cmd2 As SqlCommand = New SqlCommand(Query2, cnn)
                                Dim dr2 As SqlDataReader
                                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                                cmd2.CommandType = CommandType.Text
                                cmd2.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
                                cmd2.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                                cnn.Open()
                                dr2 = cmd2.ExecuteReader
                                TL.Load(dr2)
                                cnn.Close()
                            Catch ex As Exception
                                Edo = cnn.State.ToString
                                If Edo = "Open" Then cnn.Close()
                                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                            End Try
                            'End Using
                            'revisamos el BOM
                            Try
                                Dim Query3 As String = "SELECT * FROM tblBOM WHERE AU=@AU AND Rev=@Rev ORDER BY PN ASC"
                                Dim cmd3 As SqlCommand = New SqlCommand(Query3, cnn)
                                Dim dr3 As SqlDataReader
                                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                                cmd3.CommandType = CommandType.Text
                                cmd3.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
                                cmd3.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                                cnn.Open()
                                dr3 = cmd3.ExecuteReader
                                TB.Load(dr3)
                                cnn.Close()
                            Catch ex As Exception
                                Edo = cnn.State.ToString
                                If Edo = "Open" Then cnn.Close()
                                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                            End Try
                        End If
                    Catch ex As Exception
                        Edo = cnn.State.ToString
                        If Edo = "Open" Then cnn.Close()
                        MessageBox.Show(ex.ToString, "Error in ActualizaCambiosEnLosTags")
                        Console.WriteLine(ex.ToString())
                    End Try
                    If TN.Rows.Count > 0 Then
                        If Active.ToUpper = "FALSE" Then
                            'Resp += 1
                            ConWarning += 1
                            MensarWarning = "El AU: " + AU.ToString + " Rev: " + Rev + " no esta activo."
                            Dim Q As DataRow = tblWarning.NewRow
                            Q.Item("Warnings") = MensarWarning
                            Q.Item("ItemRow") = ConWarning
                            tblWarning.Rows.Add(Q)
                        End If
                        If TL.Rows.Count = 0 Then
                            Resp += 1
                            ConIssues += 1
                            MensajeIssues = "No hay lead time para el AU: " + AU.ToString + " Rev: " + Rev
                            Dim R As DataRow = tblIssues.NewRow
                            R.Item("Issues") = MensajeIssues
                            R.Item("ItemRow") = ConIssues
                            tblIssues.Rows.Add(R)
                        End If
                        If TB.Rows.Count = 0 Then
                            Resp += 1
                            ConIssues += 1
                            MensajeIssues = "No hay BOM para el AU: " + AU.ToString + " Rev: " + Rev
                            Dim W As DataRow = tblIssues.NewRow
                            W.Item("Issues") = MensajeIssues
                            W.Item("ItemRow") = ConIssues
                            tblIssues.Rows.Add(W)
                        End If
                    ElseIf TN.Rows.Count = 0 Then
                        Resp += 1
                        ConIssues += 1
                        MensajeIssues = "No hay un AU con este numero de parte en la base de datos " + PN
                        Dim D As DataRow = tblIssues.NewRow
                        D.Item("Issues") = MensajeIssues
                        D.Item("ItemRow") = ConIssues
                        tblIssues.Rows.Add(D)
                    End If
                End Using
            End Using
        End Using
        Return Resp
    End Function

    Private Function BuscaAU(ByVal PN As String)
        Dim Edo As String = ""
        Dim Resp As Long = 0
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                Dim Query As String = "SELECT * FROM tblMaster WHERE PN=@PN ORDER BY Active DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Resp = TN.Rows(0).Item("AU").ToString
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
        Return Resp
    End Function

    Private Function BuscaRev(ByVal PN As String)
        Dim Edo As String = ""
        Dim Resp As String = "NO"
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                Dim Query As String = "SELECT * FROM tblMaster WHERE PN=@PN ORDER BY Active DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Resp = TN.Rows(0).Item("Rev").ToString
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
        Return Resp
    End Function

    Private Function BuscaPPAP(ByVal PN As String)
        Dim Edo As String = ""
        Dim Resp As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                Dim Query As String = "SELECT * FROM tblMaster WHERE PN=@PN AND KindOfAU='PPAP'"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Resp = TN.Rows(0).Item("Rev").ToString
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
        Return Resp
    End Function
    '
    Private Function BuscaForecastReference()
        Dim Edo As String = ""
        Dim Resp As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                Dim Query As String = "SELECT TOP (1) ForecastReference FROM tblCustomerServiceForecast ORDER BY ForecastReference DESC"
                cmd = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@PN", SqlDbType.Bit).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Resp = TN.Rows(0).Item("ForecastReference").ToString
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaForecastReference")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
        Return Resp
    End Function
    'Funcion para encontrar el ultimo numero de referencia registrado en la base de datos
    Private Function BuscaNumeroDeReferencia(ByVal Tabla As String, ByVal Llave As String) As String
        Dim Edo As String = ""
        Using TN As New Data.DataTable 'Despliega los materiales 
            Dim Query As String = "SELECT TOP 1 " + Llave + " FROM " + Tabla + " ORDER BY " + Llave + " DESC "
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString.ToString + "Error loading CWO number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Console.WriteLine(ex.ToString())
            End Try
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            If TN.Rows.Count = 0 Then Edo = "" '
            'If Tabla = "tblWIP" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("WIP").ToString
            If Tabla = "tblCustomerServiceForecast" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("ID").ToString
            If Tabla = "tblPurchasingBOMWipFake" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDBOMWipFake").ToString
            If Tabla = "tblPurchasingWipFake" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("WIP").ToString
            If Tabla = "tblCustomerServiceSalesOrdersCVSFileEDIChanges" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDchange").ToString
            If Tabla = "tblCustomerServiceSalesOrders" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDQB").ToString
            If Tabla = "tblCustomerServiceSO" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("SONumber").ToString
            If Tabla = "tblCustomerServiceSalesOrdersCVSFileEDI" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDFileEDI").ToString
            If Tabla = "tblCustomerServiceSalesOrdersErrorsEDI" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDErrorEDI").ToString
            If Tabla = "tblCustomerServiceSalesOrdersTempEDI" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDTempEDI").ToString
            If Tabla = "tblQualityRMAWIP" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDRMAWIP").ToString
            If Tabla = "tblQualityRMA" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDRMA").ToString
            If Tabla = "tblQualityRMADet" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDRMADet").ToString
            If Tabla = "tblQualityRMATicketOneHistory" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDRMATicketOneHistory").ToString
            If Tabla = "tblQualityRMATicketOne" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDRMATO").ToString
            If Tabla = "tblQualityRMATicketOneDet" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDRMATicketOne").ToString
            If Tabla = "tblTicketOne" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDTicketOne").ToString
            If Tabla = "tblWIP" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("WIP").ToString
            If Tabla = "tblWipDet" Then If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("WireID").ToString
        End Using
        Return Edo
    End Function
    'Genera el Numero de serie de diferentes tablas
    Private Function GeneraSerialForecastReference(ByVal PreviousSerial As String, ByVal Tabla As String) As String
        Dim Numero, ascii1, ascii2, ascii3, ascii4 As Long
        Dim NumeroString, Letras, letra1, letra2, letra3, letra4, NewSerial As String
        NewSerial = ""
        Dim TNewSerial As String = ""
        Select Case Tabla
            Case "tblCustomerServiceForecast"
                TNewSerial = "CF"
                PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 3)
        End Select
        Try
            If PreviousSerial <> "" Then
                Letras = Microsoft.VisualBasic.Left(PreviousSerial, 4)
                Numero = Convert.ToInt64(Microsoft.VisualBasic.Right(PreviousSerial, 4))
                If Numero < 9999 Then
                    Numero = Numero + 1
                    NumeroString = Numero.ToString
                    If NumeroString.Length < 4 Then
                        For count As Integer = NumeroString.Length To 3
                            NumeroString = "0" + NumeroString
                        Next
                    End If
                    NewSerial = Letras + NumeroString
                ElseIf Numero = 9999 Then
                    NumeroString = "0001"
                    letra1 = Mid(Letras, 1, 1)
                    letra2 = Mid(Letras, 2, 1)
                    letra3 = Mid(Letras, 3, 1)
                    letra4 = Mid(Letras, 4, 1)
                    ascii1 = Asc(letra1)
                    ascii2 = Asc(letra2)
                    ascii3 = Asc(letra3)
                    ascii4 = Asc(letra4)
                    If ascii4 < 90 Then
                        ascii4 = ascii4 + 1
                    ElseIf ascii4 = 90 And ascii3 < 90 Then
                        ascii4 = 65
                        ascii3 = ascii3 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 < 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = ascii2 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 < 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = 65
                        ascii1 = ascii1 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 = 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = 65
                        ascii1 = 65
                    End If
                    letra1 = Convert.ToChar(ascii1).ToString
                    letra2 = Convert.ToChar(ascii2).ToString
                    letra3 = Convert.ToChar(ascii3).ToString
                    letra4 = Convert.ToChar(ascii4).ToString
                    Letras = letra1 + letra2 + letra3 + letra4
                    NewSerial = Letras + NumeroString
                End If
            ElseIf PreviousSerial = "" Then
                Letras = "AAAA"
                NumeroString = "0001"
                NewSerial = Letras + NumeroString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Console.WriteLine(ex.ToString())
        End Try
        TNewSerial += NewSerial
        Return TNewSerial
    End Function

    Private Function GeneraSerialBOMWipFake(ByVal PreviousSerial As String, ByVal Tabla As String) As String
        'Dim ZAS As String = Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + "000000000001"
        Dim Numero, ascii1 As Long 'ascii1, ascii2, ascii3, ascii4 As Long
        Dim NumeroString, Letras, letra1, NewSerial, TnewSerial As String  'Letras, letra1, letra2, letra3, letra4, NewSerial, TnewSerial As String
        'PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 4)
        NumeroString = ""
        NewSerial = ""
        TnewSerial = ""
        Select Case Tabla
            Case "tblPurchasingWipFake"
                PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 11) '3
                TnewSerial = "FW"
                NumeroString = PreviousSerial.ToString
                If NumeroString = "A000000000002" Then
                    NumeroString = NumeroString
                End If
            Case "tblPurchasingBOMWipFake"
                PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 11)
                TnewSerial = "BW"
            Case "tblCustomerServiceForecast"
                PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 11)
                TnewSerial = "CF"
        End Select
        Try
            If PreviousSerial <> "" Then
                Letras = Microsoft.VisualBasic.Left(PreviousSerial, 1)
                Numero = Convert.ToInt64(Microsoft.VisualBasic.Right(PreviousSerial, 12))
                If Numero < 999999999999 Then
                    Numero = Numero + 1
                    NumeroString = Numero.ToString
                    If NumeroString.Length < 12 Then
                        For count As Integer = NumeroString.Length To 11
                            NumeroString = "0" + NumeroString
                        Next
                    End If
                    NewSerial = Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + Letras + NumeroString
                ElseIf Numero = 999999999999 Then
                    NumeroString = "000000000001"
                    letra1 = Mid(Letras, 1, 1)
                    ascii1 = Asc(letra1)
                    If ascii1 < 90 Then
                        ascii1 = ascii1 + 1
                    ElseIf ascii1 = 90 Then
                        ascii1 = 65
                    End If
                    'letra2 = Mid(Letras, 2, 1)
                    'letra3 = Mid(Letras, 3, 1)
                    'letra4 = Mid(Letras, 4, 1)
                    'ascii1 = Asc(letra1)
                    'ascii2 = Asc(letra2)
                    'ascii3 = Asc(letra3)
                    'ascii4 = Asc(letra4)
                    'If ascii4 < 90 Then
                    '    ascii4 = ascii4 + 1
                    'ElseIf ascii4 = 90 And ascii3 < 90 Then
                    '    ascii4 = 65
                    '    ascii3 = ascii3 + 1
                    'ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 < 90 Then
                    '    ascii4 = 65
                    '    ascii3 = 65
                    '    ascii2 = ascii2 + 1
                    'ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 < 90 Then
                    '    ascii4 = 65
                    '    ascii3 = 65
                    '    ascii2 = 65
                    '    ascii1 = ascii1 + 1
                    'ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 = 90 Then
                    '    ascii4 = 65
                    '    ascii3 = 65
                    '    ascii2 = 65
                    '    ascii1 = 65
                    'End If
                    letra1 = Convert.ToChar(ascii1).ToString
                    'letra2 = Convert.ToChar(ascii2).ToString
                    'letra3 = Convert.ToChar(ascii3).ToString
                    'letra4 = Convert.ToChar(ascii4).ToString
                    Letras = letra1 '+ letra2 + letra3 + letra4
                    'NewSerial = Letras + NumeroString
                    NewSerial = Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + Letras + NumeroString
                End If
            ElseIf PreviousSerial = "" Then
                Letras = "A"
                NumeroString = "000000000001"
                'NewSerial = Letras + NumeroString
                NewSerial = Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + Letras + NumeroString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Console.WriteLine(ex.ToString())
        End Try
        TnewSerial += NewSerial
        Return TnewSerial
    End Function

    Private Sub InsertWIP(ByVal WIP As String, ByVal AU As Long, ByVal Rev As String, ByVal PN As String, ByVal Qty As Integer, ByVal StartDateProces As String, ByVal DueDateProcess As String, ByVal DueDateAssy As String, ByVal DueDateShipped As String, ByVal DueDateCustomer As String, ByVal Customer As String, ByVal IT As Integer, ByVal Notes As String, ByVal KindOfAU As String, ByVal Family As String, ByVal WeekProcess As Integer, ByVal Line As String, ByVal ForecastReference As String)
        'WIP, AU, Rev, PN, Qty, CreatedDate, StartDateProces, DueDateProcess, DueDateAssy, DueDateShipped, DueDateCustomer, Customer, IT, Notes, CreatedBy, KindOfAU, Family, WeekProcess, Line, ForecastReference
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = "INSERT INTO tblPurchasingWipFake (WIP, AU, Rev, PN, Qty, CreatedDate, StartDateProces, DueDateProcess, DueDateAssy, DueDateShipped, DueDateCustomer, Customer, IT, Notes, CreatedBy, KindOfAU, Family, WeekProcess, Line, ForecastReference) VALUES (@WIP, @AU, @Rev, @PN, @Qty, @CreatedDate, @StartDateProces, @DueDateProcess, @DueDateAssy, @DueDateShipped, @DueDateCustomer, @Customer, @IT, @Notes, @CreatedBy, @KindOfAU, @Family, @WeekProcess, @Line, @ForecastReference)"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@WIP", SqlDbType.NVarChar).Value = WIP
            cmd.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
            cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
            cmd.Parameters.Add("@Qty", SqlDbType.Int).Value = Qty
            cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
            cmd.Parameters.Add("@StartDateProces", SqlDbType.Date).Value = StartDateProces
            cmd.Parameters.Add("@DueDateProcess", SqlDbType.Date).Value = DueDateProcess
            cmd.Parameters.Add("@DueDateAssy", SqlDbType.Date).Value = DueDateAssy
            cmd.Parameters.Add("@DueDateShipped", SqlDbType.Date).Value = DueDateShipped
            cmd.Parameters.Add("@DueDateCustomer", SqlDbType.Date).Value = DueDateCustomer
            cmd.Parameters.Add("@Customer", SqlDbType.NVarChar).Value = Customer
            cmd.Parameters.Add("@IT", SqlDbType.Int).Value = IT
            cmd.Parameters.Add("@Notes", SqlDbType.NVarChar).Value = Notes
            cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = txbUser.Text
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = Now
            cmd.Parameters.Add("@KindOfAU", SqlDbType.NVarChar).Value = KindOfAU
            cmd.Parameters.Add("@Family", SqlDbType.NVarChar).Value = Family
            cmd.Parameters.Add("@WeekProcess", SqlDbType.Int).Value = WeekProcess
            cmd.Parameters.Add("@Line", SqlDbType.NVarChar).Value = Line
            cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error en el insert de tblPurchasingWipFake" + " " + AU, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub

    Private Sub InsertBOM(ByVal IDBOMWipFake As String, ByVal WIP As String, ByVal AU As Long, ByVal Rev As String, ByVal PN As String, ByVal Description As String, ByVal Qty As Decimal, ByVal Unit As String, ByVal MaterialGroup As String, ByVal PercentIncrease As Integer, ByVal PickList As String, ByVal Route As String, ByVal Week As Integer, ByVal LeadTime As Integer, ByVal RequieredDate As String, ByVal ProcessDate As String, ByVal FirstDayWeek As String, ByVal ForecastReference As String)
        ' IDBOMWipFake, WIP, AU, Rev, PN, Description, Qty, Unit, MaterialGroup, PercentIncrease, PickList, Route, CreatedBy, CreatedDate, Week, LeadTime, RequieredDate, ProcessDate, FirstDayWeek, ForecastReference
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = "INSERT INTO tblPurchasingBOMWipFake (IDBOMWipFake, WIP, AU, Rev, PN, Description, Qty, Unit, MaterialGroup, PercentIncrease, PickList, Route, CreatedBy, CreatedDate, Week, LeadTime, RequieredDate, ProcessDate, FirstDayWeek, ForecastReference) VALUES (@IDBOMWipFake, @WIP, @AU, @Rev, @PN, @Description, @Qty, @Unit, @MaterialGroup, @PercentIncrease, @PickList, @Route, @CreatedBy, @CreatedDate, @Week, @LeadTime, @RequieredDate, @ProcessDate, @FirstDayWeek, @ForecastReference)"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@IDBOMWipFake", SqlDbType.NVarChar).Value = IDBOMWipFake
            cmd.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
            cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
            cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
            cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = txbUser.Text
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = Now
            cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = Qty
            cmd.Parameters.Add("@WIP", SqlDbType.NVarChar).Value = WIP
            cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
            cmd.Parameters.Add("@Description", SqlDbType.NVarChar).Value = Description
            cmd.Parameters.Add("@Unit", SqlDbType.NVarChar).Value = Unit
            cmd.Parameters.Add("@MaterialGroup", SqlDbType.NVarChar).Value = MaterialGroup
            cmd.Parameters.Add("@PercentIncrease", SqlDbType.Int).Value = PercentIncrease
            cmd.Parameters.Add("@PickList", SqlDbType.NVarChar).Value = PickList
            cmd.Parameters.Add("@Route", SqlDbType.NVarChar).Value = Route
            cmd.Parameters.Add("@Week", SqlDbType.Int).Value = Week
            cmd.Parameters.Add("@LeadTime", SqlDbType.Int).Value = LeadTime
            cmd.Parameters.Add("@RequieredDate", SqlDbType.Date).Value = RequieredDate
            cmd.Parameters.Add("@ProcessDate", SqlDbType.Date).Value = ProcessDate
            cmd.Parameters.Add("@FirstDayWeek", SqlDbType.Date).Value = FirstDayWeek
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error en el insert de tblCustomerServiceForecast" + " " + PN + " WIP:" + WIP, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    '
    Private Function WIPData(ByVal AU As Long, ByVal Rev As String, ByVal PN As String, ByVal ForecastReference As String, ByVal DueDate As String, ByVal Qty As Integer)
        Dim Edo As String = ""
        Dim WIP As String = ""
        Using TN As New Data.DataTable
            Dim Customer, Notes, KindOfAU, Family, Line As String
            Dim IT, WeekProcess As Integer
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                Dim Query As String = "SELECT * FROM tblMaster WHERE AU=@AU AND Rev=@Rev"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Customer = TN.Rows(0).Item("Cust").ToString
                    IT = CInt(Val(TN.Rows(0).Item("IT").ToString))
                    Notes = TN.Rows(0).Item("Notes").ToString
                    KindOfAU = TN.Rows(0).Item("KindOfAU").ToString
                    Family = TN.Rows(0).Item("Family").ToString
                    WeekProcess = Semanas(DueDate)
                    Line = TN.Rows(0).Item("Line").ToString
                    WIP = BuscaNumeroDeReferencia("tblPurchasingWipFake", "WIP")
                    WIP = GeneraSerialBOMWipFake(WIP, "tblPurchasingWipFake")
                    InsertWIP(WIP, AU, Rev, PN, Qty, StartDateProces, DueDateProcess, DueDateAssy, DueDateShipped, DueDate, Customer, IT, Notes, KindOfAU, Family, WeekProcess, Line, ForecastReference)
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
        Return WIP
    End Function
    '
    Private Sub GeneraBOM(ByVal AU As Long, ByVal Rev As String, ByVal WIP As String, ByVal ForecastReference As String, ByVal DueDate As String, ByVal QtyWip As Integer)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try
                'variables del BOM 
                ' IDBOMWipFake, WIP, AU, Rev, PN, Description, Qty, Balance, Unit, MaterialGroup, PercentIncrease, PickList, Route, CreatedBy, CreatedDate, Week, LeadTime, RequieredDate, ProcessDate, FirstDayWeek, ForecastReference
                Dim IDBOMWipFake, PN, Description, Unit, MaterialGroup, PercentIncrease, PickList, Route, RequieredDate, ProcessDate, FirstDayWeek As String
                Dim Qty As Decimal
                Dim LeadTime, Week As Integer
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT * FROM tblBOM WHERE AU=@AU AND Rev=@Rev ORDER BY PN ASC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    For NM As Integer = 0 To TN.Rows.Count - 1
                        PN = TN.Rows(NM).Item("PN").ToString
                        If AU = 2731 Or AU = 3179 Or AU = 3598 Or AU = 3179 Then
                            AU = AU
                            If PN = "CT-1-1418448-1" Then
                                PN = PN
                            End If
                        End If
                        Description = TN.Rows(NM).Item("Description").ToString
                        Unit = TN.Rows(NM).Item("Unit").ToString
                        MaterialGroup = TN.Rows(NM).Item("MaterialGroup").ToString
                        PercentIncrease = TN.Rows(NM).Item("PercentIncrease").ToString
                        PickList = TN.Rows(NM).Item("PickList").ToString
                        Route = TN.Rows(NM).Item("Route").ToString
                        Qty = Math.Ceiling(CDec(Val(TN.Rows(NM).Item("Qty").ToString)) * QtyWip)
                        Select Case Unit
                            Case "mm"
                                Unit = "ft"
                                Qty = Math.Ceiling(Qty * 0.00328084)
                            Case "mts"
                                Unit = "ft"
                                Qty = Math.Ceiling(Qty * 3.28084)
                            Case "in"
                                Unit = "ft"
                                Qty = Math.Ceiling(Qty * 0.083333333)
                            Case "km"
                                Unit = "ft"
                                Qty = Math.Ceiling(Qty * 3280.84)
                            Case "cm"
                                Unit = "ft"
                                Qty = Math.Ceiling(Qty * 0.0328084)
                        End Select
                        'Dim StartDateProces, DueDateProcess,DueDateAssy , DueDateShipped, DueDateCustomer As Date
                        Select Case PickList
                            Case "Assembly"
                                ProcessDate = DueDateProcess
                                RequieredDate = Fechas(CDate(DueDateProcess), LeadTime, "Resta")
                                FirstDayWeek = CalculaCualEsElLunes(RequieredDate)
                                Week = Semanas(CDate(RequieredDate))
                            Case "Process"
                                ProcessDate = StartDateProces
                                RequieredDate = Fechas(CDate(StartDateProces), LeadTime, "Resta")
                                FirstDayWeek = CalculaCualEsElLunes(RequieredDate)
                                Week = Semanas(CDate(RequieredDate))
                            Case "Pack"
                                ProcessDate = DueDateAssy
                                RequieredDate = Fechas(CDate(DueDateAssy), LeadTime, "Resta")
                                FirstDayWeek = CalculaCualEsElLunes(RequieredDate)
                                Week = Semanas(CDate(RequieredDate))
                            Case Else
                                ProcessDate = StartDateProces
                                RequieredDate = Fechas(CDate(StartDateProces), LeadTime, "Resta")
                                FirstDayWeek = CalculaCualEsElLunes(RequieredDate)
                                Week = Semanas(CDate(RequieredDate))
                        End Select
                        'RequieredDate = CDate(TN.Rows(0).Item("RequieredDate").ToString).ToString("dd-MMM-yyyy")
                        'ProcessDate = CDate(TN.Rows(0).Item("ProcessDate").ToString).ToString("dd-MMM-yyyy")
                        'FirstDayWeek = CDate(TN.Rows(0).Item("FirstDayWeek").ToString).ToString("dd-MMM-yyyy")
                        IDBOMWipFake = BuscaNumeroDeReferencia("tblPurchasingBOMWipFake", "IDBOMWipFake")
                        IDBOMWipFake = GeneraSerialBOMWipFake(IDBOMWipFake, "tblPurchasingBOMWipFake")
                        InsertBOM(IDBOMWipFake, WIP, AU, Rev, PN, Description, Qty, Unit, MaterialGroup, PercentIncrease, PickList, Route, Week, LeadTime, RequieredDate, ProcessDate, FirstDayWeek, ForecastReference)
                    Next
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
    End Sub
    '
    Private Sub LeadTime(ByVal AU As Long, ByVal Rev As String, ByVal DueDate As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                ' Query = "SELECT PONUMBER AS PO, AU, ItemRevision AS Rev, ItemPartNumber AS PN, UnitPrice, QtyOrdered, DeliveryRequestedDate AS Delivery, LineItemNumber AS Line#, POBuyer AS Buyer, SEAReviewStatus AS [SEA Status], ErrQty, Changes, AssignedIdentification AS EDI_Key, PODATE, RevisionDate, RevisionNumber AS PO_Rev#, LocationNumber AS Location, PackSize, QtyEAShipped, BOLNumber, TrackingNumber, CarrierSCAC, DateShipped, TotelItemAmt, ForStoreShipTo, ForStoreBillTo, POEDI, RowEDI, SHP_RespEDI865, SHP_RespEDI865Date, EDI865, EDI865Notes, Closed, Status, IDFileEDI, FileName FROM tblCustomerServiceSalesOrdersCVSFileEDI WHERE Closed=0 AND EDI865=@EDI865 ORDER BY PONUMBER ASC, LineItemNumber ASC"
                Dim Query As String = "SELECT * FROM tblMasterProcessLeadTime WHERE AU=@AU AND Rev=@Rev"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                Dim Process, Assembly, Inspection, Shipping, TransitTime As Integer
                Dim TP, TA, TE, TT, TS As Integer
                If TN.Rows.Count > 0 Then
                    Process = CInt(Val(TN.Rows(0).Item("Process").ToString))
                    Assembly = CInt(Val(TN.Rows(0).Item("Assembly").ToString))
                    Inspection = CInt(Val(TN.Rows(0).Item("Inspection").ToString))
                    Shipping = CInt(Val(TN.Rows(0).Item("Shipping").ToString))
                    TransitTime = CInt(Val(TN.Rows(0).Item("TransitTime").ToString))
                    TT = Process + Assembly + Inspection + Shipping + TransitTime 'Tiempo total
                    TP = Assembly + Inspection + Shipping + TransitTime ' tiempo de procesos
                    TA = Inspection + Shipping + TransitTime 'Tiempo de ensamble
                    TE = Inspection + Shipping + TransitTime 'Tiempo de empaque 
                    TS = TransitTime 'Tiempo de shipping 
                    StartDateProces = Fechas(DueDate, TT, "Resta")
                    DueDateProcess = Fechas(DueDate, TP, "Resta")
                    DueDateAssy = Fechas(DueDate, TA, "Resta")
                    PackDueDate = Fechas(DueDate, TE, "Resta")
                    DueDateShipped = Fechas(DueDate, TS, "Resta")
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error in BuscaRev")
                Console.WriteLine(ex.ToString())
            End Try
        End Using
    End Sub
    'Calcula el numero de la semana del año de una fecha
    Private Function Semanas(ByVal Fecha As Date)
        Dim fechaInicio As Date = (Fecha.Year.ToString + "/1/1")
        Dim semana As Integer
        semana = DatePart(DateInterval.WeekOfYear, Fecha, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFullWeek)
        semana += 1
        Return semana.ToString
    End Function
    'Busca el LeadTime de cada numer de parte
    Private Function BuscaLeadTimeDeCadaPN(ByVal PN As String)
        Dim Edo As String = ""
        Dim Resp As Integer = 0
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT * FROM tblItemsQB WHERE (PN=@PN) ORDER BY PriOption DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@AU", SqlDbType.BigInt).Value = AU
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString.ToString + vbNewLine + "Error in Items Lead Time function.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            If TN.Rows.Count > 0 Then
                Dim AUX As Integer = 0
                For NM As Integer = 0 To TN.Rows.Count - 1
                    AUX = System.Convert.ToInt64(Val(TN.Rows(0).Item("LeadTime").ToString))
                    If ((AUX > Resp) Or (AUX = Resp)) Then
                        Resp = AUX
                    End If
                Next
            End If
        End Using
        Return Resp
    End Function
    'Calcula los dias restando un dia sin tomar en cuenta los dias habiles
    Private Function Fechas(ByVal FechaFinal As Date, ByVal CantidadDias As Integer, ByVal Task As String)
        Dim fechainicio As Date = FechaFinal
        Dim DiaDescanso As String = "", diaTemp As String = ""
        Dim XX As Integer = 0
        If Task = "Resta" Then
            While CantidadDias > 0
                If Not (fechainicio.DayOfWeek = DayOfWeek.Sunday Or fechainicio.DayOfWeek = DayOfWeek.Saturday) Then
                    CantidadDias -= 1
                    XX += 1
                    diaTemp = ""
                End If
                'funcion para identificar si es un dia de asueto
                diaTemp = fechainicio.AddDays(+1)
                'DiaDescanso = FuncionDiasDescanso(diaTemp)
                DiaDescanso = ""
                If DiaDescanso = "" Then
                    fechainicio = fechainicio.AddDays(-1)
                Else
                    CantidadDias += 1
                End If
                'fechainicio = fechainicio.AddDays(-1)
            End While
            If (fechainicio.DayOfWeek = DayOfWeek.Sunday Or fechainicio.DayOfWeek = DayOfWeek.Saturday) Then
                While (fechainicio.DayOfWeek = DayOfWeek.Sunday Or fechainicio.DayOfWeek = DayOfWeek.Saturday)
                    fechainicio = fechainicio.AddDays(-1)
                End While
            End If
        End If
        If Task = "Suma" Then
            ' Dim tope As Integer = CantidadDias - 1
            ' For NM As Integer = 0 To tope
            While CantidadDias > 0
                If Not (fechainicio.DayOfWeek = DayOfWeek.Sunday Or fechainicio.DayOfWeek = DayOfWeek.Saturday) Then
                    CantidadDias -= 1
                    XX += 1
                    diaTemp = ""
                End If
                'funcion para identificar si es un dia de asueto
                diaTemp = fechainicio.AddDays(+1)
                'DiaDescanso = FuncionDiasDescanso(diaTemp)
                DiaDescanso = ""
                If DiaDescanso = "" Then
                    fechainicio = fechainicio.AddDays(+1)
                Else
                    CantidadDias += 1
                End If
            End While
            ' Next
            If (fechainicio.DayOfWeek = DayOfWeek.Sunday Or fechainicio.DayOfWeek = DayOfWeek.Saturday) Then
                While (fechainicio.DayOfWeek = DayOfWeek.Sunday Or fechainicio.DayOfWeek = DayOfWeek.Saturday)
                    fechainicio = fechainicio.AddDays(1)
                End While
            End If
        End If
        Return fechainicio.ToString("dd/MMM/yyyy")
    End Function
    'Calcula cual es el lunes de esa semana
    Private Function CalculaCualEsElLunes(ByVal Fecha As String)
        Dim Respuesta As String = ""
        Dim FechaX As Date = CDate(Fecha)
        Dim Dia As Integer = FechaX.DayOfWeek
        Select Case Dia
            Case 1 ' "Monday"
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 2 '"Tuesday"
                Respuesta = Fechas(FechaX, 1, "Resta")
            Case 3 '"Wednesday"
                Respuesta = Fechas(FechaX, 2, "Resta")
            Case 4 '"Thursday"
                Respuesta = Fechas(FechaX, 3, "Resta")
            Case 5 '"Friday"
                Respuesta = Fechas(FechaX, 4, "Resta")
            Case 6 '"Saturday"
                Respuesta = Fechas(FechaX, 5, "Resta")
            Case 0 '"Sunday"
                Respuesta = Fechas(FechaX, 6, "Resta")
        End Select
        Return Respuesta
    End Function
    'Calcula cual es el lunes de esa semana
    Private Function CalculaCualEsElDomingo(ByVal Fecha As String)
        Dim Respuesta As String = ""
        Dim FechaX As Date = CDate(Fecha)
        Dim Dia As Integer = FechaX.DayOfWeek
        Select Case Dia
            Case 0 '"Sunday"
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 1 ' "Monday"
                FechaX = FechaX.AddDays(6) ' Fechas(FechaX, 6, "Suma")
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 2 '"Tuesday"
                FechaX = FechaX.AddDays(5) ' Fechas(FechaX, 5, "Suma")
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 3 '"Wednesday"
                FechaX = FechaX.AddDays(4) ' Fechas(FechaX, 4, "Suma")
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 4 '"Thursday"
                FechaX = FechaX.AddDays(3) ' Fechas(FechaX, 3, "Suma")
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 5 '"Friday"
                FechaX = FechaX.AddDays(2) ' Fechas(FechaX, 2, "Suma")
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 6 '"Saturday"
                FechaX = FechaX.AddDays(1) ' Fechas(FechaX, 1, "Suma")
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
            Case 7 '"Sunday"
                Respuesta = FechaX.ToString("dd/MMM/yyyy")
        End Select
        Return Respuesta
    End Function
    'Convierte unidades de medidas
    Private Function ConvierteXaY(ByVal Valor As Decimal, ByVal UMinicial As String, ByVal UMfinal As String)
        Dim Resp As String = ""
        'Referencias
        'http://www.convert-me.com/es/convert/length/
        'http://www.convertidorunidades.com/kilogramos-a-libras
        Select Case UMfinal
            Case "ea"
                'No hay conversion 
                Resp = CStr(Valor * 1)
            Case "lb"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        Resp = CStr(Valor * 1)
                    Case "ft"
                        'No hay conversion 
                        'Resp =CStr( Valor * 1)
                    Case "in"
                        'No hay conversion 
                        'Resp = CStr(Valor * 1)
                    Case "gr"
                        Resp = CStr(Valor * 453.592)
                    Case "Kg"
                        Resp = CStr(Valor * 0.453592)
                    Case "mm"
                        'No hay conversion 
                        'Resp = CStr(Valor *1)
                    Case "cm"
                        'No hay conversion 
                        'Resp = Valor * 1
                    Case "mts"
                        'No hay conversion 
                        'Resp = Valor * 1
                    Case "KM"
                        'No hay conversion 
                        'Resp = Valor * 1
                    Case "m"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "yd"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Oz"
                        Resp = CStr(Valor * 16)
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "ft"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 1)
                    Case "in"
                        Resp = CStr(Valor / 12)
                        'No hay conversion
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor / 304.8)
                    Case "cm"
                        Resp = CStr(Valor / 30.480000975359417)
                    Case "mts"
                        Resp = CStr(Valor / 0.3048)
                    Case "KM"
                        Resp = CStr(Valor * 3280.84)
                    Case "m"
                        Resp = CStr(Valor * 5280)
                    Case "yd"
                        Resp = CStr(Valor * 3)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "in"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 12)
                    Case "in"
                        Resp = CStr(Valor * 1)
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor / 25.4)
                    Case "cm"
                        Resp = CStr(Valor / 2.54)
                    Case "mts"
                        Resp = CStr(Valor / 0.0254)
                    Case "KM"
                        Resp = CStr(Valor / 0.0000254)
                    Case "m"
                        Resp = CStr(Valor * 63360)
                    Case "yd"
                        Resp = CStr(Valor / 1.0936)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "gr"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        Resp = CStr(Valor * 0.0022)
                    Case "ft"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "in"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "gr"
                        Resp = CStr(Valor * 1)
                    Case "Kg"
                        Resp = CStr(Valor * 0.001)
                    Case "mm"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "cm"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mts"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "KM"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "m"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "yd"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Oz"
                        Resp = CStr(Valor * 0.035274)
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "Kg"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        Resp = CStr(Valor * 2.20462)
                    Case "ft"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "in"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "gr"
                        Resp = CStr(Valor * 1000)
                    Case "Kg"
                        Resp = CStr(Valor * 1)
                    Case "mm"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "cm"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mts"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "KM"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "m"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "yd"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Oz"
                        Resp = CStr(Valor * 35.274)
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "mm"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion 
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 304.8)
                    Case "in"
                        Resp = CStr(Valor * 25.4)
                    Case "gr"
                        'No hay conversion 
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion 
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor * 1)
                    Case "cm"
                        Resp = CStr(Valor * 10)
                    Case "mts"
                        Resp = CStr(Valor / 0.001)
                    Case "KM"
                        Resp = CStr(Valor / 0.000001)
                    Case "m"
                        Resp = CStr(Valor / 0.0000006213712121212)
                    Case "yd"
                        Resp = CStr(Valor / 0.0010936)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "cm"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 30.48)
                    Case "in"
                        Resp = CStr(Valor * 2.54)
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor / 10)
                    Case "cm"
                        Resp = CStr(Valor * 1)
                    Case "mts"
                        Resp = CStr(Valor / 0.01)
                    Case "KM"
                        Resp = CStr(Valor / 0.00001)
                    Case "m"
                        Resp = CStr(Valor / 0.0000062137)
                    Case "yd"
                        Resp = CStr(Valor / 0.0109391)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "mts"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 0.3048)
                    Case "in"
                        Resp = CStr(Valor * 0.0254)
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor / 1000)
                    Case "cm"
                        Resp = CStr(Valor / 100)
                    Case "mts"
                        Resp = CStr(Valor / 1)
                    Case "KM"
                        Resp = CStr(Valor / 0.001)
                    Case "m"
                        Resp = CStr(Valor / 0.000621371)
                    Case "yd"
                        Resp = CStr(Valor / 1.09361)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "KM"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 0.0003048)
                    Case "in"
                        Resp = CStr(Valor * 0.0000254)
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor / 1000000)
                    Case "cm"
                        Resp = CStr(Valor / 100000)
                    Case "mts"
                        Resp = CStr(Valor / 1000)
                    Case "KM"
                        Resp = CStr(Valor * 1)
                    Case "m"
                        Resp = CStr(Valor * 1.609344)
                    Case "yd"
                        Resp = CStr(Valor / 1093.61)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "m"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 0.000189394)
                    Case "in"
                        Resp = CStr(Valor * 0.000015783)
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr((Valor / 1.609344) / 1000000)
                    Case "cm"
                        Resp = CStr(Valor * 0.0000062137)
                    Case "mts"
                        Resp = CStr(Valor * 0.00062137)
                    Case "KM"
                        Resp = CStr(Valor / 1.609344)
                    Case "m"
                        Resp = CStr(Valor * 1)
                    Case "yd"
                        Resp = CStr(Valor * 0.00056818)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "yd"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "ft"
                        Resp = CStr(Valor * 0.33333333)
                    Case "in"
                        Resp = CStr(Valor * 0.027778)
                    Case "gr"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Kg"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mm"
                        Resp = CStr(Valor * 0.0010936)
                    Case "cm"
                        Resp = CStr(Valor * 0.010936)
                    Case "mts"
                        Resp = CStr(Valor * 1.0936)
                    Case "KM"
                        Resp = CStr(Valor * 1093.613)
                    Case "m"
                        Resp = CStr(Valor * 1760)
                    Case "yd"
                        Resp = CStr(Valor * 1)
                    Case "Oz"
                        'No hay conversion 
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "Oz"
                Select Case UMinicial
                    Case "ea"
                        'No hay conversion 
                    Case "lb"
                        Resp = CStr(Valor * 0.0625)
                    Case "ft"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "in"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "gr"
                        Resp = CStr(Valor * 28.3495)
                    Case "Kg"
                        Resp = CStr(Valor * 0.0283495)
                    Case "mm"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "cm"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "mts"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "KM"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "m"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "yd"
                        'No hay conversion
                        'Resp = Valor * 1
                    Case "Oz"
                        Resp = CStr(Valor * 1)
                    Case "ton"
                        'No hay conversion 
                    Case "l"
                        'No hay conversion 
                    Case "ml"
                        'No hay conversion 
                End Select
            Case "ton"
                'No hay conversion 
            Case "l"
                'No hay conversion 
            Case "ml"
                'No hay conversion 
        End Select
        Return Resp
    End Function
    '===========================================================MRP Normal====================================================================
    '=========================================================== Controles ====================================================================
    Private Sub TabControlMRPGlobal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControlMRPGlobal.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Opcion As Integer = TabControlMRPGlobal.SelectedIndex
        Select Case Opcion
            Case 0

            Case 1
                CargaComboAUWIP()
            Case 2
                CargaComboAUENG()
            Case 3
                'CargaComboPNMyTable() Aqui no se carga
            Case 4
                CargaComboAUBOMWIPForecast()
            Case 5
                'Aqui no se carga
        End Select
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim IDReferenceMRP As String = BuscaNumeroDeReferenciaMRP()
        IDReferenceMRP = GeneraSerialMRP(IDReferenceMRP)
        lblMRPReference.Text = IDReferenceMRP
        BuscaPNsPrimarios()
        Dim Opcion As String = "All"
        Dim FechaInicio As String = "2000/01/01"
        Dim FechaFin As String = "2030/12/31"
        lblForecastReference.Text = "CFAAAA0001"
        TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
        GroupBoxUploadFile.Visible = False
        TabControlMRPGlobal.Visible = True
        CalculaMateriales(Opcion, FechaInicio, FechaFin, IDReferenceMRP, lblForecastReference.Text)
    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbExchangeRate.Text) Then
            If Val(txbExchangeRate.Text) <> 0 Then
                lblQty.Text = "Qty:"
                If LaFecha = "OK" Then
                    dtpFrom.Enabled = False
                    dtpTo.Enabled = False
                    Dim IDReferenceMRP As String = BuscaNumeroDeReferenciaMRP()
                    IDReferenceMRP = GeneraSerialMRP(IDReferenceMRP)
                    lblMRPReference.Text = IDReferenceMRP
                    Try
                        BuscaPNsPrimarios()
                        Dim Opcion As String = ""
                        Dim FechaInicio As String = dtpFrom.Value.ToShortDateString
                        Dim FechaFin As String = dtpTo.Value.ToShortDateString
                        If rdoAllWeeks.Checked = True Then
                            Opcion = "All"
                            TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
                            CalculaMateriales(Opcion, FechaInicio, FechaFin, IDReferenceMRP, lblForecastReference.Text)
                            btnCalculate.Enabled = False
                        End If
                        If rdoSpecificDates.Checked = True Then
                            Opcion = "Specific"
                            TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                            CalculaMateriales(Opcion, FechaInicio, FechaFin, IDReferenceMRP, lblForecastReference.Text)
                            btnCalculate.Enabled = False
                        End If
                        dtpFrom.Value = FechaInicio
                        dtpTo.Value = FechaFin
                        'CargaComboPNMyTable()
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString.ToString + vbNewLine + "", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            Else
                MessageBox.Show("The exchange rate can't be 0", "Important", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            MessageBox.Show("The exchange rate must be a number", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txbExchangeRate.Focus()
            txbExchangeRate.SelectAll()
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub txbUserMRPPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbUserMRPPassword.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Asc(e.KeyChar) = 13 Then
            LoginMRP()
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbUserMRPPassword_TextChanged(sender As Object, e As EventArgs) Handles txbUserMRPPassword.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilter.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbFilter.SelectedIndex > -1 And BanderaLogin > 0 Then
            MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmb10Percent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb10Percent.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmb10Percent.SelectedIndex > -1 And BanderaLogin > 0 Then
            MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoAllWeeks_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAllWeeks.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If GridMRP.RowCount > 0 Then
            MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoAllWeeks_Click(sender As Object, e As EventArgs) Handles rdoAllWeeks.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        LaFecha = "OK"
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
        dtpFrom.Value = Now.AddYears(-20).ToShortDateString
        dtpTo.Value = Now.AddYears(20).ToShortDateString
        btnCalculate.Enabled = True
        If GridMRP.RowCount > 0 Then
            MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoSpecificDates_CheckedChanged(sender As Object, e As EventArgs) Handles rdoSpecificDates.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoSpecificDates_Click(sender As Object, e As EventArgs) Handles rdoSpecificDates.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        LaFecha = "WRONG"
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        If GridMRP.Rows.Count = 0 Then
            dtpTo.Value = Now.AddDays(7).ToShortDateString
            If ckbPastDue.Checked = True Then
                Dim diaZ As Date = Now.AddDays(-14)
                dtpFrom.Value = CalculaCualEsElDomingo(diaZ.ToShortDateString.ToString)
            Else
                dtpFrom.Value = Now.ToShortDateString
            End If
        Else
            If ckbPastDue.Checked = True Then
                Dim diaZ As Date = Now.AddDays(-14)
                dtpFrom.Value = CalculaCualEsElDomingo(diaZ.ToShortDateString.ToString)
            Else
                dtpFrom.Value = Now.ToShortDateString
            End If
            dtpTo.Value = Now.AddDays(7).ToShortDateString
        End If
        btnCalculate.Enabled = True
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub dtpFrom_ValueChanged(sender As Object, e As EventArgs) Handles dtpFrom.ValueChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        lblWeekFrom.Text = Semanas(dtpFrom.Value)
        FechaInicial = dtpFrom.Value
        If GridMRP.RowCount > 0 Then
            MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub dtpTo_ValueChanged(sender As Object, e As EventArgs) Handles dtpTo.ValueChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        lblWeekTo.Text = Semanas(dtpTo.Value)
        If dtpFrom.Value > dtpTo.Value Then
            LaFecha = "WRONG"
            MessageBox.Show("Please verify your dates", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            FechaUltima = dtpTo.Value
            LaFecha = "OK"
            If GridMRP.RowCount > 0 Then
                MuestraMateriales(dtpFrom.Value, dtpTo.Value)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub txbQty_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbQty.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If e.KeyChar = ChrW(Keys.Return) Then
            If lblQty.Text = "Qty:" Then
                If OpcionCorrecta = "OK" Then If IsNumeric(txbQty.Text) Or txbQty.Text = "" Then UpdateTblPurchasingTempMRP("QtyUser", txbQty.Text, "Entero", IDx)
                MuestraMateriales(dtpFrom.Value, dtpTo.Value)
                txbQty.Text = ""
                OpcionCorrecta = "NO" '
                If Renglon > -1 Then
                    'Dim row As DataGridViewRow = Renglon
                    ' Me.GridMRP.CurrentCell = Me.GridMRP.Item(5, Renglon) 'Me.GridMRP.CurrentCell.RowIndex - 1)
                    Me.GridMRP.Rows(Renglon).Cells("QtyUser").Selected = True
                End If
                GridMRP.FirstDisplayedScrollingRowIndex = iTopRow
            ElseIf lblQty.Text = "MOQ:" Then
                If IsNumeric(txbQty.Text) Then
                    'Dim ANS As Integer = MessageBox.Show("Are you sure to change the MOQ?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    'If ANS = 6 Then GroupApproved.Visible = True
                    UpdatetblPurchasingTempMRPBySubPN("MOQ", txbQty.Text, "Decimal", SubPN)
                    UpdateItems("MOQ", txbQty.Text, "Decimal", SubPN)
                    MuestraMateriales(dtpFrom.Value, dtpTo.Value)
                    lblQty.Text = "Qty:"
                    txbQty.Text = ""
                Else
                    MessageBox.Show("This field must be a numeric data", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                GridMRP.FirstDisplayedScrollingRowIndex = iTopRow
            ElseIf lblQty.Text = " SP:" Then
                If IsNumeric(txbQty.Text) Then
                    'Dim ANS As Integer = MessageBox.Show("Are you sure to change the MOQ?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    'If ANS = 6 Then GroupApproved.Visible = True
                    UpdatetblPurchasingTempMRPBySubPN("StandarPack", txbQty.Text, "Decimal", SubPN)
                    UpdateItems("StandarPack", txbQty.Text, "Decimal", SubPN)
                    MuestraMateriales(dtpFrom.Value, dtpTo.Value)
                    lblQty.Text = "Qty:"
                    txbQty.Text = ""
                Else
                    MessageBox.Show("This field must be a numeric data", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbQty_TextChanged(sender As Object, e As EventArgs) Handles txbQty.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub ckbPastDue_CheckedChanged(sender As Object, e As EventArgs) Handles ckbPastDue.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If ckbPastDue.Checked = True And rdoSpecificDates.Checked = True Then
            Dim diaZ As Date = Now.AddDays(-14)
            dtpFrom.Value = CalculaCualEsElDomingo(diaZ.ToShortDateString.ToString)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoRequiered_CheckedChanged(sender As Object, e As EventArgs)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoRequiered_Click(sender As Object, e As EventArgs)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoNonRequiered_CheckedChanged(sender As Object, e As EventArgs)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoNonRequiered_Click(sender As Object, e As EventArgs)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoSaveReport_CheckedChanged(sender As Object, e As EventArgs) Handles rdoSaveReport.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoSaveReport_Click(sender As Object, e As EventArgs) Handles rdoSaveReport.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoViewOnly_CheckedChanged(sender As Object, e As EventArgs) Handles rdoViewOnly.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoViewOnly_Click(sender As Object, e As EventArgs) Handles rdoViewOnly.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub txbFind_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbFind.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'BuscaIDReferenceMRP(txbFind.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbFind_TextChanged(sender As Object, e As EventArgs) Handles txbFind.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        lblQty.Text = "Qty:"
        lblMRPReference.Text = ""
        'BuscaIDReferenceMRP(txbFind.Text)
        btnExportToExcel.Enabled = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnLoadMRP_Click(sender As Object, e As EventArgs) Handles btnLoadMRP.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnExportToExcel_Click(sender As Object, e As EventArgs) Handles btnExportToExcel.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim IDReferenceMRP As String = ""
        Dim NumeroDeReferencia As String = lblForecastReference.Text
        If rdoSaveReport.Checked = True Then
            IDReferenceMRP = BuscaNumeroDeReferencia()
            NumeroDeReferencia = GeneraSerial(IDReferenceMRP)
            InsertIDReferenceMRP(NumeroDeReferencia, "Forecast")
            RegistraMRP(NumeroDeReferencia)
        End If
        lblMRPReference.Text = NumeroDeReferencia
        CreaCSV(NumeroDeReferencia)
        'CreaExcel(NumeroDeReferencia)
        lblQty.Text = "Qty:"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        lblQty.Text = "Qty:"
        TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
        'TruncateTablaTemp("tblPurchasingTempMRP2" + sTempTableName)
        MuestraMateriales(dtpFrom.Value, dtpTo.Value)
        lblMRPReference.Text = ""
        btnExportToExcel.Enabled = False
        btnCalculate.Enabled = True
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        '''MuestraAyuda()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnCancelLoginEng_Click(sender As Object, e As EventArgs) Handles btnCancelLoginEng.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.Close()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridMRP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridMRP.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim cdx As Integer = e.ColumnIndex
        Dim rdx As Integer = e.RowIndex
        Dim row As DataGridViewRow = Me.GridMRP.CurrentRow
        With GridMRP
            iTopRow = .FirstDisplayedScrollingRowIndex '// get Top row.
            .FirstDisplayedScrollingRowIndex = iTopRow '// set Top row.
        End With
        'If cdx = 5 Then
        OpcionCorrecta = "OK"
        DatoX = row.Cells("QtyUser").Value.ToString
        IDx = System.Convert.ToInt64(Val(row.Cells("ID").Value.ToString))
        txbQty.Focus()
        Renglon = rdx
        lblQty.Text = "Qty:"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridMRP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridMRP.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridMRP_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridMRP.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim cdx As Integer = e.ColumnIndex
        Dim rdx As Integer = e.RowIndex
        Dim row As DataGridViewRow = Me.GridMRP.CurrentRow
        Dim Encabezado As String = GridMRP.Columns(cdx).HeaderText
        SubPN = row.Cells("SubPN").Value.ToString
        IDx = System.Convert.ToInt64(Val(row.Cells("ID").Value.ToString))
        PNBusquedaWip = row.Cells("PN").Value.ToString
        If Encabezado = "PN" Or Encabezado = "SubPN" Then
            FirstDayWeekBusquedaWip = row.Cells("FirstDayWeek").Value.ToString
            '''BuscaWips(PNBusquedaWip, FirstDayWeekBusquedaWip)
            BusquedaSalesOrders("Busca_Nada")
            GroupWipSalesOrder.Visible = True
            GridWIP.AutoResizeColumns()
        ElseIf Encabezado = "MOQ" Then
            lblQty.Text = "MOQ:"
            txbQty.Focus()
        ElseIf Encabezado = "QtyOnOrder" Then
            BusquedaDePODet(PNBusquedaWip, "Todas", Now.ToString)
            GridPurchasingOrderItemsHistory.AutoResizeColumns()
            GroupBoxPurchasingOrderHistory.Visible = True
        ElseIf Encabezado = "QtyOnOrderPerWeek" Then
            FirstDayWeekBusquedaWip = row.Cells("FirstDayWeek").Value.ToString
            BusquedaDePODet(PNBusquedaWip, "Fechas", FirstDayWeekBusquedaWip)
            GroupBoxPurchasingOrderHistory.Visible = True
            GridPurchasingOrderItemsHistory.AutoResizeColumns()
        ElseIf Encabezado = "StandarPack" Then
            lblQty.Text = " SP:"
            txbQty.Focus()
        End If
        With GridMRP
            iTopRow = .FirstDisplayedScrollingRowIndex '// get Top row.
            .FirstDisplayedScrollingRowIndex = iTopRow '// set Top row.
        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    '=========================================================== Fin Controles ============================================================================
    '=========================================================== Funciones del calculo ====================================================================
    'Genera el Numero de serie de la tabla tblPurchasingMaterialRequirementsPlanning
    Private Function GeneraSerial(ByVal PreviousSerial As String) As String
        Dim Numero, ascii1, ascii2, ascii3, ascii4 As Integer
        Dim NumeroString, Letras, letra1, letra2, letra3, letra4, NewSerial, TnewSerial As String
        NewSerial = ""
        PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 2)
        Try
            If PreviousSerial <> "" Then
                Letras = Microsoft.VisualBasic.Left(PreviousSerial, 4)
                Numero = Convert.ToInt64(Microsoft.VisualBasic.Right(PreviousSerial, 7))
                If Numero < 9999999 Then
                    Numero = Numero + 1
                    NumeroString = Numero.ToString
                    If NumeroString.Length < 7 Then
                        For count As Integer = NumeroString.Length To 6
                            NumeroString = "0" + NumeroString
                        Next
                    End If
                    NewSerial = Letras + NumeroString
                ElseIf Numero = 9999999 Then
                    NumeroString = "0000001"
                    letra1 = Mid(Letras, 1, 1)
                    letra2 = Mid(Letras, 2, 1)
                    letra3 = Mid(Letras, 3, 1)
                    letra4 = Mid(Letras, 4, 1)
                    ascii1 = Asc(letra1)
                    ascii2 = Asc(letra2)
                    ascii3 = Asc(letra3)
                    ascii4 = Asc(letra4)
                    If ascii4 < 90 Then
                        ascii4 = ascii4 + 1
                    ElseIf ascii4 = 90 And ascii3 < 90 Then
                        ascii4 = 65
                        ascii3 = ascii3 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 < 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = ascii2 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 < 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = 65
                        ascii1 = ascii1 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 = 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = 65
                        ascii1 = 65
                    End If
                    letra1 = Convert.ToChar(ascii1).ToString
                    letra2 = Convert.ToChar(ascii2).ToString
                    letra3 = Convert.ToChar(ascii3).ToString
                    letra4 = Convert.ToChar(ascii4).ToString
                    Letras = letra1 + letra2 + letra3 + letra4
                    NewSerial = Letras + NumeroString
                End If
            ElseIf PreviousSerial = "" Then
                Letras = "AAAA"
                NumeroString = "0000001"
                NewSerial = Letras + NumeroString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        TnewSerial = "R" + NewSerial
        Return TnewSerial
    End Function
    'Borra la tabla temporal de materiales para el CWO
    Private Sub TruncateTablaTemp(ByVal Tabla As String)
        Dim edo As String
        Try 'Definimos el query del insert
            Dim Query As String = "TRUNCATE TABLE " & Tabla
            Dim cmd As New SqlCommand(Query, cnn)
            cnn.Open()
            cmd.ExecuteNonQuery()
            edo = cnn.State.ToString
            If edo = "Open" Then cnn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString.ToString + "Error trying to clear " & Tabla & ", TruncateTablaTemp Function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
            Console.WriteLine(ex.ToString())
        End Try
        'Cambios(ID, "AU", "", "New Item")
        edo = cnn.State.ToString
        If edo = "Open" Then cnn.Close()
    End Sub
    '
    Private Sub GeneraColumnas()
        Dim workCol As DataColumn = tblPerWeek.Columns.Add("ItemRow", Type.GetType("System.Int64"))
        workCol.AllowDBNull = True
        tblPerWeek.Columns.Add("Date", Type.GetType("System.String"))
        tblPerWeek.Columns.Add("Week", Type.GetType("System.String"))
        tblPerWeek.Columns.Add("Amount", Type.GetType("System.Decimal"))
        Dim workCol2 As DataColumn = tblPerVendor.Columns.Add("ItemRow", Type.GetType("System.Int64"))
        workCol.AllowDBNull = True
        tblPerVendor.Columns.Add("VendorCode", Type.GetType("System.String"))
        tblPerVendor.Columns.Add("Week", Type.GetType("System.String"))
        tblPerVendor.Columns.Add("Date", Type.GetType("System.String"))
        tblPerVendor.Columns.Add("Amount", Type.GetType("System.Decimal"))
        Dim workCol3 As DataColumn = tblIssues.Columns.Add("ItemRow", Type.GetType("System.Int64"))
        tblIssues.Columns.Add("Issues", Type.GetType("System.String"))
        Dim workCol4 As DataColumn = tblWarning.Columns.Add("ItemRow", Type.GetType("System.Int64"))
        tblWarning.Columns.Add("Warnings", Type.GetType("System.String"))
    End Sub
    'Muestra los materiales que se van a comprar
    Private Sub MuestraMateriales(ByVal FechaInicio As Date, ByVal FechaFin As Date)
        Dim Edo As String = ""
        Using TN As New System.Data.DataTable 'Despliega los materiales 
            Dim PN As String = ""
            Dim Qty As Decimal = 0
            Dim UM As String = "" 'Reserved, Qty,UM AS UMOrg,
            Dim Query As String = "" '= "SELECT PN, SubPN, Qty, QtyOnHand, QtyOnOrder, QtyOnOrderPerWeek, UM as [UM Req], QtyUser, UMToBuy AS UM, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID FROM tblPurchasingTempMRPForecast" +sTempTableName
            If rdoAllWeeks.Checked = True Then Query = "SELECT PN, SubPN, QtyAcum, Qty, Difference, QtyOnOrderPerWeek, QtyOnHand, QtyOnOrder, QtyUser, UM, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
            If rdoSpecificDates.Checked = True Then Query = "SELECT PN, SubPN, QtyAcum, Difference, Qty, QtyOnOrderPerWeek, QtyOnHand, QtyOnOrder, QtyUser, UM, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE (RequieredDate BETWEEN @FechaInicio AND @FechaHasta) AND Qty>0 "
            Dim Opcion As String = cmbFilter.Text.ToString
            Select Case Opcion 'Estos queries tienen que ir igual en la funcion que crea el reporte en excel
                Case "Only Primary Without Bin Balance"
                    Query += "AND BinBalance=0 AND PriOption=1"  'AND Difference<0 
                Case "Only Primary With Bin Balance"
                    Query += "AND ((BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1)) "  'AND Difference<0 
                Case "All Without Bin Balance"
                    Query += "AND BinBalance=0 " ' AND Difference<0 
                Case "ALL"
                    Query += "AND (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
                Case "Only Bin Balance"
                    Query += "AND BinBalance=1 "  'AND Difference<0
            End Select
            Dim Opcion2 As String = cmb10Percent.Text.ToUpper
            Select Case Opcion2
                Case "ALL"
                    'No agrega Nada
                    'Query += " AND Pecent10=0"
                Case "10%"
                    'Agrega una columna al where
                    Query += " AND Pecent10=1"
            End Select
            Query += " ORDER BY SubPN ASC, FirstDayWeek ASC"
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@FechaInicio", SqlDbType.Date).Value = FechaInicio
                cmd.Parameters.Add("@FechaHasta", SqlDbType.Date).Value = FechaFin
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count > 0 Then
                Edo = Edo
            End If
            'GridMRP.DataSource = Nothing
            GridMRP.DataSource = TN
            lblRecordsMRP.Text = "Records: " + TN.Rows.Count.ToString
            If TN.Rows.Count > 0 Then
                If GridMRP.RowCount > 0 Then
                    GridMRP.Columns("QtyAcum").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("Qty").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("Difference").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("QtyOnOrderPerWeek").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("FirstDayWeek").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridMRP.Columns("Reserved").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("QtyOnHand").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("QtyOnOrder").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("QtyOnHand").DefaultCellStyle.Format = ("###,###.##")
                    'GridMRP.Columns("QtyToBuy").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("QtyUser").DefaultCellStyle.Format = ("###,###.##")
                    GridMRP.Columns("UnitPrice").DefaultCellStyle.Format = ("$###,###.##")
                    GridMRP.Columns("PackPrice").DefaultCellStyle.Format = ("$###,###.##")
                    Dim PNColumn As DataGridViewColumn = GridMRP.Columns("PN") 'QtyAcum
                    Dim SubPNColumn As DataGridViewColumn = GridMRP.Columns("SubPN")
                    Dim QtyOnHandColumn As DataGridViewColumn = GridMRP.Columns("QtyOnHand")
                    Dim QtyOnOrderColumn As DataGridViewColumn = GridMRP.Columns("QtyOnOrder")
                    Dim QtyOnOrderPerWeekColumn As DataGridViewColumn = GridMRP.Columns("QtyOnOrderPerWeek")
                    'Dim QtyToBuyColumn As DataGridViewColumn = GridMRP.Columns("QtyToBuy")
                    Dim QtyAcumColumn As DataGridViewColumn = GridMRP.Columns("QtyAcum")
                    Dim DifferenceColumn As DataGridViewColumn = GridMRP.Columns("Difference")
                    Dim QtyUserColumn As DataGridViewColumn = GridMRP.Columns("QtyUser")
                    Dim UMColumn As DataGridViewColumn = GridMRP.Columns("UM")
                    Dim QtyColumn As DataGridViewColumn = GridMRP.Columns("Qty")
                    'Dim UMReqColumn As DataGridViewColumn = GridMRP.Columns("UM Req")
                    Dim UnitPriceColumn As DataGridViewColumn = GridMRP.Columns("UnitPrice")
                    Dim PackPriceColumn As DataGridViewColumn = GridMRP.Columns("PackPrice")
                    Dim StandarPackColumn As DataGridViewColumn = GridMRP.Columns("StandarPack")
                    Dim MOQColumn As DataGridViewColumn = GridMRP.Columns("MOQ")
                    Dim LeadTimeColumn As DataGridViewColumn = GridMRP.Columns("LeadTime")
                    Dim VendorCodeColumn As DataGridViewColumn = GridMRP.Columns("VendorCode")
                    Dim DescriptionColumn As DataGridViewColumn = GridMRP.Columns("Description")
                    Dim FirstDayWeekColumn As DataGridViewColumn = GridMRP.Columns("FirstDayWeek")
                    Dim WeekColumn As DataGridViewColumn = GridMRP.Columns("Week")
                    Dim QtyInputSHPColumn As DataGridViewColumn = GridMRP.Columns("QtyInputSHP")
                    Dim KyColumn As DataGridViewColumn = GridMRP.Columns("Ky")
                    Dim IDColumn As DataGridViewColumn = GridMRP.Columns("ID")
                    PNColumn.Width = 140
                    SubPNColumn.Width = 90
                    'PNColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                    'SubPNColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                    QtyColumn.Width = 50
                    QtyOnHandColumn.Width = 70
                    QtyOnOrderColumn.Width = 70
                    QtyOnOrderPerWeekColumn.Width = 70
                    DifferenceColumn.Width = 70
                    'QtyToBuyColumn.Width = 50
                    QtyUserColumn.Width = 50
                    QtyAcumColumn.Width = 50
                    UMColumn.Width = 55
                    'UMReqColumn.Width = 25
                    UnitPriceColumn.Width = 50
                    PackPriceColumn.Width = 50
                    StandarPackColumn.Width = 50
                    MOQColumn.Width = 40
                    LeadTimeColumn.Width = 30
                    VendorCodeColumn.Width = 70
                    DescriptionColumn.Width = 70
                    FirstDayWeekColumn.Width = 70
                    WeekColumn.Width = 35
                    QtyInputSHPColumn.Width = 70
                    KyColumn.Width = 30
                    IDColumn.Width = 30
                End If
            End If
            'GridMRP.AutoResizeColumns()
            If TN.Rows.Count > 0 Then btnExportToExcel.Enabled = True
            If GridMRP.Rows.Count > 0 Then
                cmbFilter.Enabled = True
                cmb10Percent.Enabled = True
            Else
                cmbFilter.Enabled = False
                cmb10Percent.Enabled = False
            End If
            'lblTotal.Text = "Total: " & String.Format("{0:C}", CalculaTotalTotal())
            'lblTotalTotal2.Text = lblTotal.Text
            'CalculaTotalPerWeek()
            FechaInicial = "2016/01/01"
            FechaUltima = "2018/12/31"
            dtpFrom.Value = FechaInicial
            dtpTo.Value = FechaUltima
        End Using
    End Sub
    'Calcula los materiales que estan en la tabla tblPurchasingBOMWipFake dependiendo de el @ForecastReference con el que se guardo
    Private Sub CalculaMateriales(ByVal Opcion As String, ByVal FechaInicio As String, ByVal FechaFin As String, ByVal IDReferenceMRP As String, ByVal ForecastReference As String) 'tblPurchasingMaterialRequirementsPlanning
        Dim Edo As String = ""
        Dim Aprovado As Boolean = False
        Dim LunesActual As Date = CalculaCualEsElLunes(Now.ToShortDateString.ToString)
        Dim LunesAnterior As Date = CalculaCualEsElLunes(Now.AddDays(-7).ToShortDateString.ToString)
        Dim SemanaActual As String = Semanas(LunesActual)
        Dim SemanaAnterior As String = Semanas(LunesAnterior)
        'Parte Uno es donde se calcula la sumatoria de todos los materiales directos desde el BOM en los WIP abiertos
        Using TN As New System.Data.DataTable
            Dim TW As New System.Data.DataTable
            Dim Query As String = "SELECT tblPurchasingBOMWipFake.WIP, tblPurchasingBOMWipFake.AU, tblPurchasingBOMWipFake.Rev, tblPurchasingBOMWipFake.PN, tblPurchasingBOMWipFake.Description, tblPurchasingBOMWipFake.Qty AS Qty, tblPurchasingBOMWipFake.Unit AS UM, tblPurchasingBOMWipFake.PickList, tblPurchasingBOMWipFake.Week, tblPurchasingBOMWipFake.LeadTime, tblPurchasingBOMWipFake.RequieredDate, tblPurchasingBOMWipFake.ProcessDate, tblPurchasingBOMWipFake.FirstDayWeek, tblPurchasingWipFake.StartDateProces, tblPurchasingWipFake.DueDateProcess, tblPurchasingWipFake.DueDateAssy, tblPurchasingWipFake.DueDateShipped, tblPurchasingWipFake.DueDateCustomer FROM tblPurchasingBOMWipFake INNER JOIN tblPurchasingWipFake ON tblPurchasingBOMWipFake.WIP = tblPurchasingWipFake.WIP WHERE tblPurchasingWipFake.ForecastReference=@ForecastReference " ' AND tblPurchasingWipFake.KindOfAU<>'PPAP'
            Dim BalancePN As Decimal = 0
            Dim Available As String
            Dim Qty As Double = 0
            Dim FechaHasta As Date = Now.AddYears(20).ToShortDateString
            Dim WIP, AU, Rev, PN, Unit, PickList, PO, Week, LeadTime, RequieredDate, ProcessDate, FirstDayWeek As String ' Description,
            Dim StartDateProces, DueDateProcess, DueDateAssy, DueDateShipped, DueDateCustomer As String 'BalanceProcess, BalanceAssy, BalancePack, BalanceShipped, wSort, EstimatedStartDateProces,
            For JT As Integer = 0 To 1
                TN.Reset()
                Select Case JT
                    Case 0
                        Query = "SELECT tblPurchasingBOMWipFake.WIP, tblPurchasingBOMWipFake.AU, tblPurchasingBOMWipFake.Rev, tblPurchasingBOMWipFake.PN, tblPurchasingBOMWipFake.Description, tblPurchasingBOMWipFake.Qty AS Qty, tblPurchasingBOMWipFake.Unit AS UM, tblPurchasingBOMWipFake.PickList, tblPurchasingBOMWipFake.Week, tblPurchasingBOMWipFake.LeadTime, tblPurchasingBOMWipFake.RequieredDate, tblPurchasingBOMWipFake.ProcessDate, tblPurchasingBOMWipFake.FirstDayWeek, tblPurchasingWipFake.StartDateProces, tblPurchasingWipFake.DueDateProcess, tblPurchasingWipFake.DueDateAssy, tblPurchasingWipFake.DueDateShipped, tblPurchasingWipFake.DueDateCustomer FROM tblPurchasingBOMWipFake INNER JOIN tblPurchasingWipFake ON tblPurchasingBOMWipFake.WIP = tblPurchasingWipFake.WIP WHERE tblPurchasingWipFake.ForecastReference=@ForecastReference"  'AND tblPurchasingWipFake.KindOfAU<>'PPAP'
                        Try
                            Dim cmd As SqlCommand
                            Dim dr As SqlDataReader
                            cmd = New SqlCommand(Query, cnn)
                            cmd.CommandType = CommandType.Text
                            cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
                            cmd.CommandTimeout = 1000
                            'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                            cnn.Open()
                            dr = cmd.ExecuteReader
                            TN.Load(dr)
                            cnn.Close()
                            Dim Contador As Long = TN.Rows.Count
                        Catch ex As Exception
                            cnn.Close() 'cierra la conexion
                            MessageBox.Show(ex.ToString + "Error Loading PO from BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    Case 1             'Desactivado ya que solicito mario no tomar en cuenta el material que fue sacado fuera de sistema 10/10/16
                        'Query = "SELECT * FROM tblItemsMaterialRequestDet WHERE ((PO=0) AND (Status='Open'))" ' (RequieredDate BETWEEN '01/1/2000' AND @FechaFinal))"
                        ''Case 2
                        ''    Query = "SELECT * FROM tblItemsMaterialRequestDet WHERE ((PO=0) AND (RequieredDate BETWEEN @FechaInicio AND @FechaFinal))"
                        'Try
                        '    Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                        '    Dim dr As SqlDataReader
                        '    cnn.Open()
                        '    dr = cmd.ExecuteReader
                        '    TW.Load(dr)
                        '    cnn.Close()
                        'Catch ex As Exception
                        '    MessageBox.Show(ex.tostring + "Error Loading PO from BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        'End Try
                End Select
                Dim XCVZ As Integer = 0
                Dim QtyOnHand, QtyOnOrder As Decimal
                Dim QtyReq As Double = 0
                Dim UM As String = ""
                Dim SubPN As String = ""
                Dim Vendor As String = ""
                Dim VendorCode As String = ""
                Dim VendorPN As String = ""
                Dim PackPrice As String = ""
                Dim UnitPrice As String = ""
                Dim MOQ As String = ""
                Dim StandarPack As String = ""
                Dim BinBalance As String = ""
                Dim KindPurchasing As String = ""
                Dim UMInputSHP As String = ""
                Dim UMVendor As String = ""
                Dim QtyInputSHP As String = ""
                Dim ExactlyQuantity As String = ""
                Dim Ky As String = ""
                Dim Description As String = ""
                Dim Reserved As String = ""
                Dim TotalQty As Decimal = 0
                Dim Faltante As Decimal = 0
                Dim QtyOnOrderPerWeek As Double = 0
                Dim Difference As Double = 0
                Dim PriOption As String = ""
                Dim QtyAcum As Decimal = 0
                Dim Pecent10 As Boolean = False
                Dim QtyOnOrderPerPeriod As Decimal = 0
                Dim X As String
                If TN.Rows.Count > 0 Then '
                    'ByVal Qty As Double, ByVal UM As String, ByVal Task As String, ByVal SubPN As String, ByVal LeadTime As String, ByVal Vendor As String, ByVal VendorCode As String, ByVal VendorPN As String, ByVal PackPrice As String, ByVal UnitPrice As String, ByVal MOQ As String, ByVal StandarPack As String, ByVal BinBalance As String, ByVal KindPurchasing As String, ByVal UMVendor As String, ByVal UMInputSHP As String, ByVal QtyInputSHP As String, ByVal ExactlyQuantity As String, ByVal Ky As String, ByVal Description As String, ByVal QtyOnHand As Decimal, ByVal QtyOnOrder As Decimal, ByVal RequieredDate As String, ByVal FirstDayWeek As String, ByVal Week As Integer, ByVal Reserved As Double, ByVal TotalQty As Double, ByVal Faltante As Decimal, ByVal IDReferenceMRP As String, ByVal QtyOnOrderPerWeek As Double, ByVal Difference As Double
                    For NM As Long = 0 To TN.Rows.Count - 1
                        'wSort = TN.Rows(NM).Item("wSort").ToString
                        PickList = TN.Rows(NM).Item("PickList").ToString
                        'If ((PickList = "Process" Or PickList = "Assembly" Or PickList = "Pack")) Or ((wSort = "7" Or wSort = "8" Or wSort = "9" Or wSort = "10" Or wSort = "11" Or wSort = "12" Or wSort = "13" Or wSort = "14" Or wSort = "15") And (PickList = "Assembly" Or PickList = "Pack")) Or ((wSort = "16" Or wSort = "17" Or wSort = "18" Or wSort = "19") And (PickList = "Pack")) Then
                        PN = TN.Rows(NM).Item("PN").ToString
                        Qty = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Qty").ToString))
                        X = TN.Rows(NM).Item("Qty").ToString
                        UM = TN.Rows(NM).Item("UM").ToString
                        If PN = "C119" Then
                            PN = PN
                        End If
                        If JT = 0 Then
                            'Unit = TN.Rows(NM).Item("Unit").ToString
                            WIP = TN.Rows(NM).Item("WIP").ToString
                            AU = TN.Rows(NM).Item("AU").ToString
                            Rev = TN.Rows(NM).Item("Rev").ToString
                            ProcessDate = TN.Rows(NM).Item("ProcessDate").ToString
                            LeadTime = TN.Rows(NM).Item("LeadTime").ToString
                            Week = TN.Rows(NM).Item("Week").ToString
                            FirstDayWeek = TN.Rows(NM).Item("FirstDayWeek").ToString
                            'BalanceProcess = TN.Rows(NM).Item("BalanceProcess").ToString
                            'BalanceAssy = TN.Rows(NM).Item("BalanceAssy").ToString
                            'BalancePack = TN.Rows(NM).Item("BalancePack").ToString
                            'BalanceShipped = TN.Rows(NM).Item("BalanceShipped").ToString
                            'EstimatedStartDateProces = TN.Rows(NM).Item("EstimatedStartDateProces").ToString
                            StartDateProces = TN.Rows(NM).Item("StartDateProces").ToString
                            DueDateProcess = TN.Rows(NM).Item("DueDateProcess").ToString
                            DueDateAssy = TN.Rows(NM).Item("DueDateAssy").ToString
                            DueDateShipped = TN.Rows(NM).Item("DueDateShipped").ToString
                            DueDateCustomer = TN.Rows(NM).Item("DueDateCustomer").ToString
                            Available = 0 ' BuscaPNQty(PN)
                            BalancePN = 0
                        End If
                        If JT = 1 Or JT = 2 Then Unit = TN.Rows(NM).Item("UM").ToString
                        RequieredDate = TN.Rows(NM).Item("RequieredDate").ToString
                        PO = "" 'TN.Rows(NM).Item("PO").ToString
                        Week = TN.Rows(NM).Item("Week").ToString
                        LeadTime = TN.Rows(NM).Item("LeadTime").ToString
                        FirstDayWeek = CalculaCualEsElLunes(RequieredDate)
                        If CDate(FirstDayWeek) < LunesActual Then
                            FirstDayWeek = LunesAnterior.ToString("dd/MMM/yyyy")
                            Week = SemanaAnterior
                        End If
                        'Evaluamos si el PN lleva un SubBOM
                        BuscandoSubBOMs(PN, Qty, UM, FirstDayWeek, Week, RequieredDate)
                        PN += "@" + UM + "&" + FirstDayWeek + "*" + Week.ToString
                        XCVZ += 1
                        If XCVZ = 7489 Then
                            XCVZ = XCVZ
                        End If
                        InsertTablaTemp(PN, Qty, QtyReq, "", "Paso Uno", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Week), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
                        'UM,  SubPN, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, UMVendor, QtyInputSHP,  ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder Reserved
                        'End If
                    Next
                End If
            Next
        End Using
        Using TN As New System.Data.DataTable 'Calculamos la suma de los materiales
            Dim Reserved As String = ""
            Dim PN As String = ""
            Dim Qty As Double = 0
            Dim QtyReq As Double = 0
            Dim UM As String = ""
            Dim SubPN As String = ""
            Dim LeadTime As String = ""
            Dim Vendor As String = ""
            Dim VendorCode As String = ""
            Dim VendorPN As String = ""
            Dim PackPrice As String = ""
            Dim UnitPrice As String = ""
            Dim MOQ As String = ""
            Dim StandarPack As String = ""
            Dim BinBalance As String = ""
            Dim KindPurchasing As String = ""
            Dim UMVendor As String = ""
            Dim UMInputSHP As String = ""
            Dim QtyInputSHP As String = ""
            Dim ExactlyQuantity As String = ""
            Dim Description As String = ""
            Dim Ky As String = ""
            Dim QtyOnHand As Decimal = 0
            Dim QtyOnOrder As Decimal = 0
            Dim Week As String = ""
            Dim RequieredDate As String = ""
            Dim FirstDayWeek As String = ""
            Dim TotalQty As Decimal = 0
            Dim Faltante As Decimal = 0
            Dim QtyOnOrderPerWeek As Double = 0
            Dim Difference As Double = 0
            Dim PriOption As String = ""
            Dim QtyAcum As Decimal = 0
            Dim Pecent10 As Boolean = False
            Dim QtyOnOrderPerPeriod As Decimal = 0
            Dim Query As String = "SELECT PN, SUM(QTY) AS Qty FROM tblPurchasingTempMRPForecast" + sTempTableName + " GROUP BY PN"
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
            For NM As Long = 0 To TN.Rows.Count - 1
                PN = TN.Rows(NM).Item("PN").ToString
                Qty = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Qty").ToString))
                InsertTablaTemp(PN, Qty, QtyReq, UM, "Paso Uno", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Val(Week)), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
            Next
        End Using
        Using TN As New System.Data.DataTable 'Despliega los materiales ya sumados
            Dim PN As String = ""
            Dim Qty As Decimal = 0
            Dim UM As String = ""
            Dim Query As String = "SELECT PN, QTY FROM tblPurchasingTempMRPForecast" + sTempTableName
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
            Dim Tope As Integer
            Dim Bandera As Integer = 0
            Dim caracter As Char
            Dim Cadena As String
            Dim Week As String = ""
            Dim RequieredDate As String = ""
            Dim FirstDayWeek As String = ""
            Dim TotalQty As Decimal = 0
            Dim Faltante As Decimal = 0
            '
            Dim QtyReq As Double = 0
            Dim SubPN As String = ""
            Dim LeadTime As String = ""
            Dim Vendor As String = ""
            Dim VendorCode As String = ""
            Dim VendorPN As String = ""
            Dim PackPrice As String = ""
            Dim UnitPrice As String = ""
            Dim MOQ As String = ""
            Dim StandarPack As String = ""
            Dim BinBalance As String = ""
            Dim KindPurchasing As String = ""
            Dim UMVendor As String = ""
            Dim UMInputSHP As String = ""
            Dim QtyInputSHP As String = ""
            Dim ExactlyQuantity As String = ""
            Dim Description As String = ""
            Dim Ky As String = ""
            Dim QtyOnHand As Decimal = 0
            Dim QtyOnOrder As Decimal = 0
            Dim QtyOnOrderPerWeek As Double = 0
            Dim Difference As Double = 0
            Dim PriOption As String = ""
            Dim QtyAcum As Decimal = 0
            Dim Pecent10 As Boolean = False
            Dim Reserved As Double = 0
            Dim QtyOnOrderPerPeriod As Decimal = 0
            For NM As Long = 0 To TN.Rows.Count - 1
                Cadena = TN.Rows(NM).Item("PN").ToString
                Qty = TN.Rows(NM).Item("Qty").ToString
                'Descomprime PN y la Unidad de medida
                Tope = (Len(Cadena))
                For P As Integer = 1 To Tope
                    caracter = Microsoft.VisualBasic.Mid(Cadena, P)
                    If caracter = "@" Then Bandera = 1
                    If caracter = "&" Then Bandera += 1
                    If caracter = "*" Then Bandera += 1
                    If caracter <> "@" And Bandera = 0 Then PN += caracter
                    If caracter <> "@" And Bandera = 1 Then UM += caracter
                    If caracter <> "&" And Bandera = 2 Then FirstDayWeek += caracter ' FirstDayWeek += caracter
                    If caracter <> "*" And Bandera = 3 Then Week += caracter
                Next
                If PN = "CE-J301" Then
                    PN = PN
                End If
                InsertTablaTemp(PN, Qty, QtyReq, UM, "Identificacion", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ.ToString, StandarPack.ToString, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Val(Week)), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
                'BuscaSiExistePartNumberEnItemsQB(PN, Qty, UM, FirstDayWeek, Week, IDReferenceMRP)
                Bandera = 0
                UM = ""
                PN = ""
                Week = ""
                FirstDayWeek = ""
                RequieredDate = ""
            Next
        End Using
        Using TN As New System.Data.DataTable 'Despliega los materiales ya sumados
            Dim PN As String = ""
            Dim Qty As Decimal = 0
            Dim UM As String = ""
            Dim Query As String = "SELECT * FROM tblPurchasingTempMRPForecast" + sTempTableName + " ORDER BY PN ASC, FirstDayWeek ASC"
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
            Dim Bandera As Integer = 0
            Dim Week As String = ""
            Dim RequieredDate As String = ""
            Dim FirstDayWeek As String = ""
            Dim TotalQty As Decimal = 0
            Dim Faltante As Decimal = 0
            For NM As Long = 0 To TN.Rows.Count - 1
                PN = TN.Rows(NM).Item("PN").ToString
                Qty = TN.Rows(NM).Item("Qty").ToString
                UM = TN.Rows(NM).Item("UM").ToString
                FirstDayWeek = CDate(TN.Rows(NM).Item("FirstDayWeek").ToString).ToString("dd-MMM-yy")
                Week = TN.Rows(NM).Item("Week").ToString
                IDReferenceMRP = TN.Rows(NM).Item("IDReferenceMRP").ToString
                If PN = "CE-J301" Then
                    PN = PN
                End If
                BuscaSiExistePartNumberEnItemsQB(PN, Qty, UM, FirstDayWeek, Week, IDReferenceMRP)
                Bandera = 0
                UM = ""
                PN = ""
                Week = ""
                FirstDayWeek = ""
                RequieredDate = ""
            Next
        End Using
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'Se eliminaron por requerimiento nuevo de mario para ir mejorando el calculo del MRP 17/Ene/2017
        'Using TN As New System.Data.DataTable 'Calculamos la suma de los materiales
        '    Dim Reserved As String = ""
        '    Dim PN As String = ""
        '    Dim Qty As Decimal = 0
        '    Dim QtyReq As Double = 0
        '    Dim UM As String = ""
        '    Dim SubPN As String = ""
        '    Dim LeadTime As String = ""
        '    Dim Vendor As String = ""
        '    Dim VendorCode As String = ""
        '    Dim VendorPN As String = ""
        '    Dim PackPrice As String = ""
        '    Dim UnitPrice As String = ""
        '    Dim MOQ As String = ""
        '    Dim StandarPack As String = ""
        '    Dim BinBalance As String = ""
        '    Dim KindPurchasing As String = ""
        '    Dim UMVendor As String = ""
        '    Dim UMInputSHP As String = ""
        '    Dim QtyInputSHP As String = ""
        '    Dim ExactlyQuantity As String = ""
        '    Dim Description As String = ""
        '    Dim Ky As String = ""
        '    Dim QtyOnHand As Decimal = 0
        '    Dim QtyOnOrder As Decimal = 0
        '    Dim Week As String = ""
        '    Dim RequieredDate As String = ""
        '    Dim FirstDayWeek As String = ""
        '    Dim TotalQty As Decimal = 0
        '    Dim Faltante As Decimal = 0
        '    Dim QtyToBuy As String = ""
        '    Dim QtyUser As String = ""
        '    Dim UMToBuy As String = ""
        '    Dim PO As String = ""
        '    Dim QtyPO As String = ""
        '    Dim QtyOnOrderPerWeek As Double = 0
        '    Dim Difference As Double = 0
        '    'Dim IDReferenceMRP As String = ""
        '    Dim ID As String = ""
        '    Dim Query As String = "SELECT * FROM tblPurchasingTempMRPForecast"  +sTempTableName+" ORDER BY FirstDayWeek ASC, PN ASC"
        '    Try
        '        Dim dr As SqlDataReader
        '        Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
        '        'cmd.CommandType = CommandType.Text
        '        'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
        '        'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
        '        cnn.Open()
        '        dr = cmd.ExecuteReader
        '        TN.Load(dr) ''Llena la tabla
        '        Edo = cnn.State.ToString
        '        If Edo = "Open" Then cnn.Close()
        '    Catch ex As Exception
        '        MessageBox.Show(ex.tostring.ToString + "Error loading materials with requierment, BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        '    Edo = cnn.State.ToString
        '    If Edo = "Open" Then cnn.Close()
        '    'TruncateTablaTemp("tblPurchasingTempMRPForecast" +sTempTableName)
        '    For NM As Long = 0 To TN.Rows.Count - 1
        '        PN = TN.Rows(NM).Item("PN").ToString
        '        Qty = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Qty").ToString))
        '        QtyReq = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyReq").ToString))
        '        UM = TN.Rows(NM).Item("UM").ToString
        '        SubPN = TN.Rows(NM).Item("SubPN").ToString
        '        LeadTime = TN.Rows(NM).Item("LeadTime").ToString
        '        Vendor = TN.Rows(NM).Item("Vendor").ToString
        '        VendorCode = TN.Rows(NM).Item("VendorCode").ToString
        '        VendorPN = TN.Rows(NM).Item("VendorPN").ToString
        '        PackPrice = TN.Rows(NM).Item("PackPrice").ToString
        '        UnitPrice = TN.Rows(NM).Item("UnitPrice").ToString
        '        MOQ = TN.Rows(NM).Item("MOQ").ToString
        '        StandarPack = TN.Rows(NM).Item("StandarPack").ToString
        '        BinBalance = TN.Rows(NM).Item("BinBalance").ToString
        '        KindPurchasing = TN.Rows(NM).Item("KindPurchasing").ToString
        '        UMVendor = TN.Rows(NM).Item("UMVendor").ToString
        '        UMInputSHP = TN.Rows(NM).Item("UMInputSHP").ToString
        '        QtyInputSHP = TN.Rows(NM).Item("QtyInputSHP").ToString
        '        ExactlyQuantity = TN.Rows(NM).Item("ExactlyQuantity").ToString
        '        Description = TN.Rows(NM).Item("Description").ToString
        '        Ky = TN.Rows(NM).Item("Ky").ToString
        '        QtyOnHand = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnHand").ToString))
        '        QtyOnOrder = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnOrder").ToString))
        '        Week = TN.Rows(NM).Item("Week").ToString
        '        RequieredDate = TN.Rows(NM).Item("RequieredDate").ToString
        '        FirstDayWeek = TN.Rows(NM).Item("FirstDayWeek").ToString
        '        TotalQty = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("TotalQty").ToString))
        '        Faltante = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("Faltante").ToString))
        '        QtyToBuy = TN.Rows(NM).Item("QtyToBuy").ToString
        '        QtyUser = TN.Rows(NM).Item("QtyUser").ToString
        '        UMToBuy = TN.Rows(NM).Item("UMToBuy").ToString
        '        PO = TN.Rows(NM).Item("PO").ToString
        '        QtyPO = TN.Rows(NM).Item("QtyPO").ToString
        '        'IDReferenceMRP = TN.Rows(NM).Item("IDReferenceMRP").ToString
        '        ID = TN.Rows(NM).Item("ID").ToString
        '        'TotalQty = BuscaPNQty(PN, "tblPurchasingTempMRPForecast" +sTempTableName)'Se eliminaron por requerimiento nuevo de mario para ir mejorando el calculo del MRP 17/Ene/2017
        '        'ActualizandoCantidaesDeTblPurchasingTempMRP(PN) 'Se eliminaron por requerimiento nuevo de mario para ir mejorando el calculo del MRP 17/Ene/2017

        '    Next
        'End Using
        'Using TN As New System.Data.DataTable 'Calculamos la suma de los materiales
        '    Dim Reserved As String = ""
        '    Dim PN As String = ""
        '    Dim Qty As Double = 0
        '    Dim QtyReq As Double = 0
        '    Dim UM As String = ""
        '    Dim SubPN As String = ""
        '    Dim LeadTime As String = ""
        '    Dim Vendor As String = ""
        '    Dim VendorCode As String = ""
        '    Dim VendorPN As String = ""
        '    Dim PackPrice As String = ""
        '    Dim UnitPrice As String = ""
        '    Dim MOQ As String = ""
        '    Dim StandarPack As String = ""
        '    Dim BinBalance As String = ""
        '    Dim KindPurchasing As String = ""
        '    Dim UMVendor As String = ""
        '    Dim UMInputSHP As String = ""
        '    Dim QtyInputSHP As String = ""
        '    Dim ExactlyQuantity As String = ""
        '    Dim Description As String = ""
        '    Dim Ky As String = ""
        '    Dim QtyOnHand As Decimal = 0
        '    Dim QtyOnOrder As Decimal = 0
        '    Dim Week As String = ""
        '    Dim FirstDayWeek As String = ""
        '    Dim RequieredDate As String = ""
        '    Dim TotalQty As Decimal = 0
        '    Dim Faltante As Decimal = 0
        '    Dim QtyToBuy As String = ""
        '    Dim QtyUser As String = ""
        '    Dim UMToBuy As String = ""
        '    Dim PO As String = ""
        '    Dim QtyPO As String = ""
        '    Dim Difference As Double = 0
        '    Dim QtyOnOrderPerWeek As Double = 0
        '    ' Dim IDReferenceMRP As String = ""
        '    Dim ID As String = ""
        '    Dim Query As String = "SELECT * FROM tblPurchasingTempMRP2"  +sTempTableName+" ORDER BY Week ASC, PN ASC"
        '    Try
        '        Dim dr As SqlDataReader
        '        Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
        '        'cmd.CommandType = CommandType.Text
        '        'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
        '        'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
        '        cnn.Open()
        '        dr = cmd.ExecuteReader
        '        TN.Load(dr) ''Llena la tabla
        '        Edo = cnn.State.ToString
        '        If Edo = "Open" Then cnn.Close()
        '    Catch ex As Exception
        '        MessageBox.Show(ex.tostring.ToString + "Error loading materials with requierment, BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        '    Edo = cnn.State.ToString
        '    If Edo = "Open" Then cnn.Close()
        '    TruncateTablaTemp("tblPurchasingTempMRPForecast" +sTempTableName)
        '    For NM As Long = 0 To TN.Rows.Count - 1
        '        PN = TN.Rows(NM).Item("PN").ToString
        '        Qty = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Qty").ToString))
        '        QtyReq = CDbl(Val(TN.Rows(NM).Item("QtyReq").ToString))
        '        UM = TN.Rows(NM).Item("UM").ToString
        '        SubPN = TN.Rows(NM).Item("SubPN").ToString
        '        LeadTime = TN.Rows(NM).Item("LeadTime").ToString
        '        Vendor = TN.Rows(NM).Item("Vendor").ToString
        '        VendorCode = TN.Rows(NM).Item("VendorCode").ToString
        '        VendorPN = TN.Rows(NM).Item("VendorPN").ToString
        '        PackPrice = TN.Rows(NM).Item("PackPrice").ToString
        '        UnitPrice = TN.Rows(NM).Item("UnitPrice").ToString
        '        MOQ = TN.Rows(NM).Item("MOQ").ToString
        '        StandarPack = TN.Rows(NM).Item("StandarPack").ToString
        '        BinBalance = TN.Rows(NM).Item("BinBalance").ToString
        '        KindPurchasing = TN.Rows(NM).Item("KindPurchasing").ToString
        '        UMVendor = TN.Rows(NM).Item("UMVendor").ToString
        '        UMInputSHP = TN.Rows(NM).Item("UMInputSHP").ToString
        '        QtyInputSHP = TN.Rows(NM).Item("QtyInputSHP").ToString
        '        ExactlyQuantity = TN.Rows(NM).Item("ExactlyQuantity").ToString
        '        Description = TN.Rows(NM).Item("Description").ToString
        '        Ky = TN.Rows(NM).Item("Ky").ToString
        '        QtyOnHand = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnHand").ToString))
        '        QtyOnOrder = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnOrder").ToString))
        '        Week = TN.Rows(NM).Item("Week").ToString
        '        FirstDayWeek = TN.Rows(NM).Item("FirstDayWeek").ToString
        '        RequieredDate = TN.Rows(NM).Item("RequieredDate").ToString
        '        TotalQty = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("TotalQty").ToString))
        '        Faltante = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("Faltante").ToString))
        '        QtyToBuy = TN.Rows(NM).Item("QtyToBuy").ToString
        '        QtyUser = TN.Rows(NM).Item("QtyUser").ToString
        '        UMToBuy = TN.Rows(NM).Item("UMToBuy").ToString
        '        PO = TN.Rows(NM).Item("PO").ToString
        '        QtyPO = TN.Rows(NM).Item("QtyPO").ToString
        '        Difference = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Difference").ToString))
        '        'IDReferenceMRP = TN.Rows(NM).Item("IDReferenceMRP").ToString
        '        ID = TN.Rows(NM).Item("ID").ToString
        '        TotalQty = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("TotalQty").ToString))
        '        If PN = "WAB181" Then
        '            PN = PN
        '        End If
        '        InsertTablaTemp(PN, Qty, QtyReq, UM, "Paso Dos", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Val(Week)), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference)
        '    Next
        'End Using
        'TruncateTablaTemp("tblPurchasingTempMRP2" +sTempTableName)
        'CalculaMaterialAComprar()
        CargaComboAUWIPForecast()
        CargaComboAUBOMWIPForecast()
        CargaComboPNMyTable()
        MuestraMateriales(FechaInicio, FechaFin)
        dtpFrom.Value = FechaInicio
        dtpTo.Value = FechaFin
    End Sub
    'Busca si el PN tiene un SubBOM para agregarlo al requerimiento de material
    Private Sub BuscandoSubBOMs(ByVal PN As String, ByVal Qty As Decimal, ByVal UM As String, ByVal FirstDayWeek As String, ByVal Week As String, ByVal RequieredDate As String)
        Using TN As New Data.DataTable
            Dim Edo As String = ""
            Dim NuevaFecha As String = ""
            Try 'tblItemsFinantialInventoryControlTempforProductionProcess" & sTempTableName 
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT * FROM tblItemsQB WHERE PN=@PN AND SubBOM=1 ORDER BY PriOption DESC "
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Dim LeadTimeNewPN As Integer = CInt(Val(TN.Rows(0).Item("LeadTime").ToString))
                    NuevaFecha = Fechas(CDate(FirstDayWeek), LeadTimeNewPN, "Resta")
                    Agregando_SubBOMs(PN, Qty, UM, NuevaFecha, Week, RequieredDate)
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading BuscaDiaDelProceso ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
    Private Sub Agregando_SubBOMs(ByVal NewPN As String, ByVal QtyPNOriginal As Double, ByVal UMPNOriginal As String, ByVal FechaDeNewPN As String, ByVal Week As String, ByVal RequieredDate As String)
        Dim FirstDayWeek As String
        Using TN As New Data.DataTable
            Dim Edo As String = ""
            Try 'tblItemsFinantialInventoryControlTempforProductionProcess" & sTempTableName 
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT * FROM tblItemsSubBOMs WHERE NewPN=@NewPN "
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@NewPN", SqlDbType.NVarChar).Value = NewPN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Dim PN, UM As String, QtyPN, Qty As Decimal, QtyReq As Decimal = 0, LeadTime As Integer = 0
                    Dim Vendor As String = "", VendorCode As String = "", VendorPN As String = "", UMVendor As String = "", UMInputSHP As String = "", Ky As String = "", Description As String = "", IDReferenceMRP As String = "", PriOption As Integer = 0
                    Dim PackPrice As Decimal = 0, UnitPrice As Decimal = 0, MOQ As Decimal = 0, StandarPack As Decimal = 0, BinBalance As Integer = 0, KindPurchasing As Integer = 0, Reserved As Integer = 0, QtyInputSHP As Integer = 0, ExactlyQuantity As Integer = 0, QtyOnHand As Integer = 0, QtyOnOrder As Integer = 0, TotalQty As Integer = 0, Faltante As Integer = 0, QtyOnOrderPerWeek As Integer = 0, Difference As Integer = 0, QtyAcum As Decimal = 0, Pecent10 As Decimal = 0, QtyOnOrderPerPeriod As Decimal = 0
                    For NM As Integer = 0 To TN.Rows.Count - 1
                        PN = TN.Rows(NM).Item("PN").ToString
                        UM = TN.Rows(NM).Item("Unit").ToString
                        QtyPN = CDec(Val(TN.Rows(NM).Item("Qty").ToString))
                        'PN += "@" + UM + "&" + FirstDayWeek + "*" + Week.ToString
                        ' InsertTablaTemp(PN, Qty, QtyReq, "", "Paso Uno", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Week), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
                        'Buscar el leadTime de este PN
                        'LeadTime = BuscaLeadTime(PN)
                        'Agrega el leadtime a la fecha del material a retrabajar de ese numero de parte 
                        FirstDayWeek = FechaDeNewPN 'Fechas(CDate(FechaDeNewPN), LeadTime, "Resta")
                        PN += "@" + UM + "&" + FirstDayWeek + "*" + Week.ToString
                        Qty = QtyPN * QtyPNOriginal
                        InsertTablaTemp(PN, Qty, QtyReq, "", "Paso Uno", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Week), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
                    Next
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading BuscaDiaDelProceso ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
    Private Function BuscaLeadTime(ByVal PN As String)
        Dim Resp As Integer
        Using TN As New Data.DataTable
            Dim Edo As String = ""
            Try 'tblItemsFinantialInventoryControlTempforProductionProcess" & sTempTableName 
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT * FROM tblItemsQB WHERE PN=@PN ORDER BY PriOption DESC "
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    Resp = CInt(Val(TN.Rows(0).Item("LeadTime").ToString))
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading BuscaDiaDelProceso ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        Return Resp
    End Function

    'Inserta un numero de parte en la tabla temporal de los materiales
    Private Sub InsertTablaTemp(ByVal PN As String, ByVal Qty As Double, ByVal QtyReq As Double, ByVal UM As String, ByVal Task As String, ByVal SubPN As String, ByVal LeadTime As String, ByVal Vendor As String, ByVal VendorCode As String, ByVal VendorPN As String, ByVal PackPrice As String, ByVal UnitPrice As String, ByVal MOQ As String, ByVal StandarPack As String, ByVal BinBalance As String, ByVal KindPurchasing As String, ByVal UMVendor As String, ByVal UMInputSHP As String, ByVal QtyInputSHP As String, ByVal ExactlyQuantity As String, ByVal Ky As String, ByVal Description As String, ByVal QtyOnHand As Decimal, ByVal QtyOnOrder As Decimal, ByVal RequieredDate As String, ByVal FirstDayWeek As String, ByVal Week As Integer, ByVal Reserved As Double, ByVal TotalQty As Double, ByVal Faltante As Decimal, ByVal IDReferenceMRP As String, ByVal QtyOnOrderPerWeek As Double, ByVal Difference As Double, ByVal PriOption As String, ByVal QtyAcum As Decimal, ByVal Pecent10 As Boolean, ByVal QtyOnOrderPerPeriod As Decimal)
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = "INSERT INTO tblPurchasingTempMRPForecast" + sTempTableName + " (PN, Qty) VALUES (@PN, @Qty)"
            Select Case Task
                Case "Paso Uno"
                    Query = "INSERT INTO tblPurchasingTempMRPForecast" + sTempTableName + " (PN, Qty) VALUES (@PN, @Qty)"
                Case "Paso Dos"
                    'Query = "INSERT INTO tblPurchasingMaterialRequirementsPlanning (IDMaterialPurchasing, PN, SubPN, QtyOnHand, QtyOnOrder, QtyToBuy, QtyUser, UMToBuy, Qty, UM, StandarPack, UnitPrice, PackPrice, LeadTime, VendorPN, VendorCode, Vendor, BinBalance, Description, Difference, IDReferenceMRP, MOQ, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, Ky, RequieredDate, FirstDayWeek, Week, QtyOnOrderPerWeek, CreatedBy, CreatedDate, QtyAcum, Pecent10) VALUES (@IDMaterialPurchasing, @PN, @SubPN, @QtyOnHand, @QtyOnOrder, @QtyToBuy, @QtyUser, @UMToBuy, @Qty, @UM, @StandarPack, @UnitPrice, @PackPrice, @LeadTime, @VendorPN, @VendorCode, @Vendor, @BinBalance, @Description, @Difference, @IDReferenceMRP, @MOQ, @KindPurchasing, @UMVendor, @UMInputSHP, @QtyInputSHP, @Ky, @RequieredDate, @FirstDayWeek, @Week, @QtyOnOrderPerWeek, @CreatedBy, @CreatedDate, @QtyAcum, @Pecent10)"
                    Query = "INSERT INTO tblPurchasingTempMRPForecast" + sTempTableName + " (PN, Qty, QtyReq, UM, SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, Week, Reserved, TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod) VALUES (@PN, @Qty, @QtyReq, @UM, @SubPN, @LeadTime, @Vendor, @VendorCode, @VendorPN, @PackPrice, @UnitPrice, @MOQ, @StandarPack, @BinBalance, @KindPurchasing, @UMVendor, @UMInputSHP, @QtyInputSHP, @ExactlyQuantity, @Ky, @Description, @QtyOnHand, @QtyOnOrder, @RequieredDate, @FirstDayWeek, @Week, @Reserved, @TotalQty, @Faltante, @IDReferenceMRP, @QtyOnOrderPerWeek, @Difference, @PriOption, @QtyAcum, @Pecent10, @QtyOnOrderPerPeriod)"
                Case "Identificacion"
                    Query = "INSERT INTO tblPurchasingTempMRPForecast" + sTempTableName + " (PN, Qty, UM, FirstDayWeek, Week, IDReferenceMRP) VALUES (@PN, @Qty, @UM, @FirstDayWeek, @Week, @IDReferenceMRP)"
            End Select
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
            cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = Qty
            If Task = "Identificacion" Then 'PN, Qty, UM, FirstDayWeek, Week, IDReferenceMRP
                cmd.Parameters.Add("@UM", SqlDbType.NVarChar).Value = UM
                cmd.Parameters.Add("@FirstDayWeek", SqlDbType.Date).Value = CDate(FirstDayWeek)
                cmd.Parameters.Add("@Week", SqlDbType.NVarChar).Value = Week
                cmd.Parameters.Add("@IDReferenceMRP", SqlDbType.NVarChar).Value = IDReferenceMRP
            End If
            If Task = "Paso Dos" Then
                cmd.Parameters.Add("@QtyReq", SqlDbType.Float).Value = Convert.ToDouble(QtyReq)
                cmd.Parameters.Add("@UM", SqlDbType.NVarChar).Value = UM
                cmd.Parameters.Add("@SubPN", SqlDbType.NVarChar).Value = SubPN
                cmd.Parameters.Add("@LeadTime", SqlDbType.Int).Value = CInt(Val(LeadTime))
                cmd.Parameters.Add("@Vendor", SqlDbType.NVarChar).Value = Vendor
                cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
                cmd.Parameters.Add("@VendorPN", SqlDbType.NVarChar).Value = VendorPN
                cmd.Parameters.Add("@PackPrice", SqlDbType.Decimal).Value = System.Convert.ToDouble(Val(PackPrice))
                cmd.Parameters.Add("@UnitPrice", SqlDbType.Decimal).Value = System.Convert.ToDouble(Val(UnitPrice))
                cmd.Parameters.Add("@Description", SqlDbType.NVarChar).Value = Description
                cmd.Parameters.Add("@MOQ", SqlDbType.Int).Value = System.Convert.ToDouble(Val(MOQ))
                cmd.Parameters.Add("@StandarPack", SqlDbType.Int).Value = System.Convert.ToDouble(Val(StandarPack))
                cmd.Parameters.Add("@BinBalance", SqlDbType.Bit).Value = CBool(BinBalance)
                cmd.Parameters.Add("@KindPurchasing", SqlDbType.Bit).Value = CBool(KindPurchasing)
                cmd.Parameters.Add("@UMVendor", SqlDbType.NVarChar).Value = UMVendor
                cmd.Parameters.Add("@UMInputSHP", SqlDbType.NVarChar).Value = UMInputSHP
                cmd.Parameters.Add("@QtyInputSHP", SqlDbType.Decimal).Value = System.Convert.ToDouble(Val(QtyInputSHP))
                cmd.Parameters.Add("@ExactlyQuantity", SqlDbType.Bit).Value = CBool(ExactlyQuantity)
                cmd.Parameters.Add("@Ky", SqlDbType.NVarChar).Value = Ky 'RequieredDate, ProcessDate ,Week,FirstDayWeek ,LeadTime
                cmd.Parameters.Add("@QtyOnHand", SqlDbType.Decimal).Value = QtyOnHand
                cmd.Parameters.Add("@QtyOnOrder", SqlDbType.Decimal).Value = QtyOnOrder
                cmd.Parameters.Add("@Week", SqlDbType.NVarChar).Value = Week
                cmd.Parameters.Add("@RequieredDate", SqlDbType.Date).Value = CDate(RequieredDate)
                cmd.Parameters.Add("@FirstDayWeek", SqlDbType.Date).Value = CDate(FirstDayWeek)
                cmd.Parameters.Add("@Reserved", SqlDbType.Decimal).Value = Reserved
                cmd.Parameters.Add("@TotalQty", SqlDbType.Decimal).Value = TotalQty
                cmd.Parameters.Add("@Faltante", SqlDbType.Decimal).Value = Faltante
                cmd.Parameters.Add("@IDReferenceMRP", SqlDbType.NVarChar).Value = IDReferenceMRP
                cmd.Parameters.Add("@QtyOnOrderPerWeek", SqlDbType.Decimal).Value = QtyOnOrderPerWeek
                cmd.Parameters.Add("@Difference", SqlDbType.Float).Value = Difference
                cmd.Parameters.Add("@PriOption", SqlDbType.Bit).Value = CBool(PriOption)
                cmd.Parameters.Add("@QtyAcum", SqlDbType.Decimal).Value = QtyAcum
                cmd.Parameters.Add("@Pecent10", SqlDbType.Bit).Value = Pecent10
                cmd.Parameters.Add("@QtyOnOrderPerPeriod", SqlDbType.Decimal).Value = QtyOnOrderPerPeriod
            End If 'PN, Qty,"Paso Uno", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack,BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error en el insert de tblPurchasingTempMRPForecast" + sTempTableName + " SubPN: " + SubPN, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
            Console.WriteLine(ex.ToString)
        End Try
    End Sub
    'Inserta un numero de parte en la tabla temporal de los materiales
    Private Sub InsertTablaTemp2(ByVal PN As String, ByVal Qty As Double, ByVal QtyReq As Double, ByVal UM As String, ByVal Task As String, ByVal SubPN As String, ByVal LeadTime As String, ByVal Vendor As String, ByVal VendorCode As String, ByVal VendorPN As String, ByVal PackPrice As String, ByVal UnitPrice As String, ByVal MOQ As String, ByVal StandarPack As String, ByVal BinBalance As String, ByVal KindPurchasing As String, ByVal UMVendor As String, ByVal UMInputSHP As String, ByVal QtyInputSHP As String, ByVal ExactlyQuantity As String, ByVal Ky As String, ByVal Description As String, ByVal QtyOnHand As Decimal, ByVal QtyOnOrder As Decimal, ByVal RequieredDate As String, ByVal FirstDayWeek As String, ByVal Week As Integer, ByVal Reserved As Double, ByVal TotalQty As Double, ByVal Faltante As Decimal, ByVal IDReferenceMRP As String, ByVal QtyOnOrderPerWeek As Double, ByVal Difference As Double, ByVal PriOption As String, ByVal QtyAcum As Decimal, ByVal Pecent10 As Boolean, ByVal QtyOnOrderPerPeriod As Decimal)
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = "INSERT INTO tblPurchasingTempMRP2" + sTempTableName + " (PN, Qty) VALUES (@PN, @Qty)"
            Select Case Task
                Case "Paso Uno"
                    Query = "INSERT INTO tblPurchasingTempMRP2" + sTempTableName + " (PN, Qty) VALUES (@PN, @Qty)"
                Case "Paso Dos"
                    Query = "INSERT INTO tblPurchasingTempMRP2" + sTempTableName + " (PN, Qty, QtyReq, UM, SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack,BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, Week, Reserved, TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod) VALUES (@PN, @Qty, @QtyReq, @UM, @SubPN, @LeadTime, @Vendor, @VendorCode, @VendorPN, @PackPrice, @UnitPrice, @MOQ, @StandarPack, @BinBalance, @KindPurchasing, @UMVendor, @UMInputSHP, @QtyInputSHP, @ExactlyQuantity, @Ky, @Description, @QtyOnHand, @QtyOnOrder, @RequieredDate, @FirstDayWeek, @Week, @Reserved, @TotalQty, @Faltante, @IDReferenceMRP, @QtyOnOrderPerWeek, @Difference, @PriOption, @QtyAcum, @Pecent10, @QtyOnOrderPerPeriod)"
            End Select
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
            cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = Qty
            If Task = "Paso Dos" Then
                cmd.Parameters.Add("@QtyReq", SqlDbType.Float).Value = QtyReq
                cmd.Parameters.Add("@UM", SqlDbType.NVarChar).Value = UM
                cmd.Parameters.Add("@SubPN", SqlDbType.NVarChar).Value = SubPN
                cmd.Parameters.Add("@LeadTime", SqlDbType.Int).Value = CInt(Val(LeadTime))
                cmd.Parameters.Add("@Vendor", SqlDbType.NVarChar).Value = Vendor
                cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
                cmd.Parameters.Add("@VendorPN", SqlDbType.NVarChar).Value = VendorPN
                cmd.Parameters.Add("@PackPrice", SqlDbType.Decimal).Value = System.Convert.ToDouble(Val(PackPrice))
                cmd.Parameters.Add("@UnitPrice", SqlDbType.Decimal).Value = System.Convert.ToDouble(Val(UnitPrice))
                cmd.Parameters.Add("@Description", SqlDbType.NVarChar).Value = Description
                cmd.Parameters.Add("@MOQ", SqlDbType.Int).Value = System.Convert.ToDouble(Val(MOQ))
                cmd.Parameters.Add("@StandarPack", SqlDbType.Int).Value = System.Convert.ToDouble(Val(StandarPack))
                cmd.Parameters.Add("@BinBalance", SqlDbType.Bit).Value = CBool(Val(BinBalance))
                cmd.Parameters.Add("@KindPurchasing", SqlDbType.Bit).Value = CBool(KindPurchasing)
                cmd.Parameters.Add("@UMVendor", SqlDbType.NVarChar).Value = UMVendor
                cmd.Parameters.Add("@UMInputSHP", SqlDbType.NVarChar).Value = UMInputSHP
                cmd.Parameters.Add("@QtyInputSHP", SqlDbType.Decimal).Value = System.Convert.ToDouble(Val(QtyInputSHP))
                cmd.Parameters.Add("@ExactlyQuantity", SqlDbType.Bit).Value = CBool(ExactlyQuantity)
                cmd.Parameters.Add("@Ky", SqlDbType.NVarChar).Value = Ky 'RequieredDate, ProcessDate ,Week,FirstDayWeek ,LeadTime
                cmd.Parameters.Add("@QtyOnHand", SqlDbType.Decimal).Value = QtyOnHand
                cmd.Parameters.Add("@QtyOnOrder", SqlDbType.Decimal).Value = QtyOnOrder
                cmd.Parameters.Add("@Week", SqlDbType.NVarChar).Value = Week
                cmd.Parameters.Add("@RequieredDate", SqlDbType.Date).Value = CDate(RequieredDate)
                cmd.Parameters.Add("@FirstDayWeek", SqlDbType.Date).Value = CDate(FirstDayWeek)
                cmd.Parameters.Add("@Reserved", SqlDbType.Decimal).Value = Reserved
                cmd.Parameters.Add("@TotalQty", SqlDbType.Decimal).Value = TotalQty
                cmd.Parameters.Add("@Faltante", SqlDbType.Decimal).Value = Faltante
                cmd.Parameters.Add("@IDReferenceMRP", SqlDbType.NVarChar).Value = IDReferenceMRP
                cmd.Parameters.Add("@QtyOnOrderPerWeek", SqlDbType.Decimal).Value = QtyOnOrderPerWeek
                cmd.Parameters.Add("@Difference", SqlDbType.Float).Value = Difference
                cmd.Parameters.Add("@PriOption", SqlDbType.Bit).Value = CBool(PriOption)
                cmd.Parameters.Add("@QtyAcum", SqlDbType.Decimal).Value = QtyAcum
                cmd.Parameters.Add("@Pecent10", SqlDbType.Bit).Value = Pecent10
                cmd.Parameters.Add("@QtyOnOrderPerPeriod", SqlDbType.Decimal).Value = QtyOnOrderPerPeriod
            End If 'PN, Qty,"Paso Uno", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack,BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error en el insert de tblPurchasingTempMRP2" + sTempTableName + " SubPN: " + SubPN, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    '
    Private Sub BuscaSiExistePartNumberEnItemsQB(ByVal PN As String, ByVal Qty As Decimal, ByVal UM As String, ByVal FirstDayWeek As String, ByVal Week As String, ByVal IDReferenceMRP As String)
        Dim Edo As String = ""
        Using TN As New System.Data.DataTable
            Dim BanderaQtyAcumulada As Integer = 0
            Dim BQtyAcum As Decimal = 0
            Dim PriOption As String = ""
            Dim UMOriginal As String = UM
            Dim QtyOriginal As Decimal = Qty
            Dim DueDate As String = CStr(Now.ToShortDateString)
            Dim WeekDueDate As String
            Dim RequieredDate As Date
            Dim SubPN As String = ""
            Dim LeadTime As String = ""
            Dim Reserved As String = ""
            Dim Vendor As String = ""
            Dim VendorCode As String = ""
            Dim VendorPN As String = ""
            Dim PackPrice As String = ""
            Dim UnitPrice As Decimal = 0, UnitPriceMXN As Decimal = 0
            Dim MOQ As Double = 0
            Dim StandarPack As Double = 0
            Dim BinBalance As String = ""
            Dim KindPurchasing As String = ""
            Dim UMVendor As String = ""
            Dim UMInputSHP As String = ""
            Dim QtyInputSHP As String = ""
            Dim ExactlyQuantity As String = ""
            Dim Description As String = ""
            Dim Ky As String = ""
            Dim Unit As String = ""
            Dim QtyOnHand As Decimal = 0
            Dim QtyOnOrder As Decimal = 0
            Dim QtyReq As Double = 0
            Dim TotalQty As Decimal = 0
            Dim Faltante As Decimal = 0
            Dim QtyOnOrderPerWeek As Double = 0
            Dim Difference As Double = 0
            Dim Country As String
            Dim QtyAcum As Decimal = 0
            Dim DiezPorciento As Decimal = 0
            Dim Pecent10 As Boolean = False
            Dim QtyOnOrderPerPeriod As Decimal = 0
            Dim Currency As String = ""
            Dim ExchangeRate As Decimal = 0
            Dim banderaReservado As Integer
            Dim Query As String = "SELECT PN, SubPN, Reserved, Description, QtyOnHand, QtyOnOrder, Unit, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, UnitPriceMXN, ExchangeRate, Currency, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, PriOption FROM tblItemsQB WHERE PN=@PN AND Active = '1'"
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                If TN.Rows.Count > 0 Then
                    For NM As Long = 0 To TN.Rows.Count - 1
                        PriOption = TN.Rows(NM).Item("PriOption").ToString
                        Reserved = TN.Rows(NM).Item("Reserved").ToString
                        SubPN = TN.Rows(NM).Item("SubPN").ToString.ToUpper
                        LeadTime = TN.Rows(NM).Item("LeadTime").ToString
                        Unit = TN.Rows(NM).Item("Unit").ToString
                        Vendor = TN.Rows(NM).Item("Vendor").ToString.ToUpper
                        VendorCode = TN.Rows(NM).Item("VendorCode").ToString.ToUpper
                        VendorPN = TN.Rows(NM).Item("VendorPN").ToString.ToUpper
                        PackPrice = TN.Rows(NM).Item("PackPrice").ToString
                        UnitPrice = CDec(Val(TN.Rows(NM).Item("UnitPrice").ToString))
                        UnitPriceMXN = CDec(Val(TN.Rows(NM).Item("UnitPriceMXN").ToString))
                        ExchangeRate = CDec(Val(TN.Rows(NM).Item("ExchangeRate").ToString))
                        Currency = TN.Rows(NM).Item("Currency").ToString.ToUpper
                        MOQ = System.Convert.ToDouble(Val(TN.Rows(NM).Item("MOQ").ToString))
                        StandarPack = System.Convert.ToDouble(Val(TN.Rows(NM).Item("StandarPack").ToString))
                        BinBalance = TN.Rows(NM).Item("BinBalance").ToString
                        KindPurchasing = TN.Rows(NM).Item("KindPurchasing")
                        UMVendor = TN.Rows(NM).Item("UMVendor").ToString
                        UMInputSHP = TN.Rows(NM).Item("UMInputSHP").ToString
                        QtyInputSHP = TN.Rows(NM).Item("QtyInputSHP").ToString
                        ExactlyQuantity = TN.Rows(NM).Item("ExactlyQuantity").ToString
                        Description = TN.Rows(NM).Item("Description").ToString
                        QtyOnHand = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnHand").ToString))
                        QtyOnOrder = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnOrder").ToString))
                        TotalQty = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyOnHand").ToString))
                        Faltante = 0 'System.Convert.ToDecimal(Val(TN.Rows(NM).Item("Faltante").ToString))
                        Country = Vendors(VendorCode)
                        UnitPrice = UnitPrice
                        'If Currency.ToUpper <> "DLLS" Then
                        '    UnitPrice = UnitPrice '/ txbExchangeRate.Text
                        'End If
                        If Country.ToUpper = "MX" Then
                            RequieredDate = CDate(FirstDayWeek)
                        Else
                            DueDate = Fechas(FirstDayWeek, 3, "Suma")
                            RequieredDate = Fechas(FirstDayWeek, 3, "Suma")
                        End If
                        WeekDueDate = Semanas(CDate(DueDate)) 'BB-TKS
                        If BinBalance = "True" Then
                            BinBalance = BinBalance
                        End If
                        If PN = "CE-J301" Then
                            PN = PN
                        End If
                        If PN = "BB-TP2" Or PN = "BB-TP3" Or PN = "BB-TP4" Or PN = "BB-TP7" Or PN = "BB-SRS" Or PN = "BB-TP1" Or PN = "BB-TKS" Then
                            PN = PN
                        End If
                        QtyOnOrder = BuscaTotalQtyOnOrderEnTblItemsPODet(PN)
                        QtyOnOrderPerPeriod = BuscaQtyOnOrderEnTblItemsPODet(PN, RequieredDate)
                        QtyOnOrderPerWeek = BuscaPNEnPODet(PN, RequieredDate, KindPurchasing, UMInputSHP, CDbl(Val(QtyInputSHP)), MOQ, StandarPack, UMVendor, UMInputSHP)
                        If KindPurchasing = "True" Then
                            If Unit <> "ea" And Unit <> "Oz" And Unit <> "ton" And Unit <> "l" And Unit <> "ml" Then
                                If QtyInputSHP > 0 Then
                                    QtyOnHand = ConvierteXaY(QtyOnHand, UMOriginal, Unit)
                                    QtyOnOrder = ConvierteXaY(QtyOnOrder, UMOriginal, Unit)
                                    Reserved = ConvierteXaY(Reserved, UMOriginal, Unit)
                                    QtyOnHand = Math.Round(QtyOnHand / QtyInputSHP) ' Math.Ceiling(QtyOnHand / QtyInputSHP)
                                    QtyOnOrder = Math.Round(QtyOnOrder / QtyInputSHP) 'Math.Ceiling(QtyOnOrder / QtyInputSHP)
                                    Qty = QtyOriginal / QtyInputSHP ' Math.Round(QtyOriginal / QtyInputSHP) ' Math.Ceiling(QtyOriginal / QtyInputSHP)
                                    Reserved = Math.Round(Reserved / QtyInputSHP) ' Math.Ceiling(QtyOriginal / QtyInputSHP)
                                    UM = UMVendor
                                Else
                                    MessageBox.Show("Please check the confoguration of PN: " + SubPN + ". The Qty Input to SHP can't be 0 or less than 0", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End If
                            ElseIf Unit = "ea" Then
                                If QtyInputSHP > 0 Then
                                    QtyOnHand = ConvierteXaY(QtyOnHand, UMOriginal, Unit)
                                    QtyOnOrder = ConvierteXaY(QtyOnOrder, UMOriginal, Unit)
                                    Reserved = ConvierteXaY(Reserved, UMOriginal, Unit)
                                    QtyOnHand = Math.Round(QtyOnHand / QtyInputSHP) ' Math.Ceiling(QtyOnHand / QtyInputSHP)
                                    QtyOnOrder = Math.Round(QtyOnOrder / QtyInputSHP) 'Math.Ceiling(QtyOnOrder / QtyInputSHP)
                                    Qty = QtyOriginal / QtyInputSHP ' Math.Round(QtyOriginal / QtyInputSHP) ' Math.Ceiling(QtyOriginal / QtyInputSHP)
                                    Reserved = Math.Round(Reserved / QtyInputSHP) ' Math.Ceiling(QtyOriginal / QtyInputSHP)
                                    UM = UMVendor
                                End If
                            ElseIf Unit = "kg" Then
                                If QtyInputSHP > 0 Then
                                    QtyOnHand = ConvierteXaY(QtyOnHand, UMOriginal, Unit)
                                    QtyOnOrder = ConvierteXaY(QtyOnOrder, UMOriginal, Unit)
                                    Reserved = ConvierteXaY(Reserved, UMOriginal, Unit)
                                    QtyOnHand = Math.Round(QtyOnHand / QtyInputSHP) ' Math.Ceiling(QtyOnHand / QtyInputSHP)
                                    QtyOnOrder = Math.Round(QtyOnOrder / QtyInputSHP) 'Math.Ceiling(QtyOnOrder / QtyInputSHP)
                                    Qty = QtyOriginal / QtyInputSHP 'Math.Round(QtyOriginal / QtyInputSHP) ' Math.Ceiling(QtyOriginal / QtyInputSHP)
                                    Reserved = Math.Round(Reserved / QtyInputSHP) ' Math.Ceiling(QtyOriginal / QtyInputSHP)
                                    UM = UMVendor
                                End If
                            End If
                            'QtyOnHand = Fix(QtyOnHand / QtyInputSHP)
                            'QtyOnOrder = Fix(QtyOnOrder / QtyInputSHP)
                            'UM = UMVendor
                        ElseIf KindPurchasing = "False" Then
                            If Unit <> "Oz" And Unit <> "ton" And Unit <> "l" And Unit <> "ml" Then
                                QtyOnHand = ConvierteXaY(QtyOnHand, UM, Unit)
                                QtyOnOrder = ConvierteXaY(QtyOnOrder, UM, Unit)
                                Reserved = ConvierteXaY(Reserved, UM, Unit)
                            End If
                        End If
                        If NM = 0 Then
                            Ky = "*"
                        Else
                            Ky = ""
                        End If
                        'agregar validacion para para agregar el reservado o no
                        banderaReservado = BuscaPNConReservado(PN, SubPN)
                        If BanderaQtyAcumulada = 0 Then
                            BQtyAcum = BuscaRegistroEnUnaTabla(PN, "tblPurchasingTempMRPForecast" + sTempTableName, "QtyAcum", SubPN, FirstDayWeek)
                            BanderaQtyAcumulada += 1
                        End If
                        If banderaReservado = 0 Then
                            QtyAcum = BQtyAcum
                            QtyAcum += Qty
                            Difference = (QtyOnHand + QtyOnOrder - Reserved) - (QtyAcum)
                            DiezPorciento = (QtyOnHand * 0.1)
                            If ((QtyAcum <= DiezPorciento) And QtyAcum > 0) Then
                                Pecent10 = False
                            Else
                                Pecent10 = True
                            End If
                        Else
                            QtyAcum = BQtyAcum
                            QtyAcum += Qty
                            Difference = (QtyOnHand + QtyOnOrder) - (QtyAcum)
                            DiezPorciento = (QtyOnHand * 0.1)
                            If ((QtyAcum <= DiezPorciento) And QtyAcum > 0) Then
                                Pecent10 = False
                            Else
                                Pecent10 = True
                            End If
                        End If
                        If PN = "CD-3105014.017" Then
                            PN = PN
                        End If
                        If BinBalance = "False" Then 'Identifica si es Bin Balance
                            QtyReq = Qty
                            InsertTablaTemp(PN, Qty, QtyReq, UM, "Paso Dos", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ.ToString, StandarPack.ToString, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Val(Week)), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
                        ElseIf BinBalance = "True" Then
                            QtyReq = Qty
                            InsertTablaTemp(PN, Qty, QtyReq, UM, "Paso Dos", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ.ToString, StandarPack.ToString, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Val(Week)), CDec(Val(Reserved)), TotalQty, Faltante, IDReferenceMRP, QtyOnOrderPerWeek, Difference, PriOption, QtyAcum, Pecent10, QtyOnOrderPerPeriod)
                        End If
                        Pecent10 = False
                    Next
                End If
            Catch ex As Exception
                'quitar
                'MessageBox.Show(ex.ToString.ToString + "Error loading PN data, BuscaPartNumberEnItemsQB function" + vbNewLine + "Please check the configuration of PN: " + PN + " on the master list and BOM.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
        End Using
    End Sub
    'Regresa el pais del proveedor
    Private Function Vendors(ByVal VendorCode As String)
        Dim Resp As String = ""
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT * FROM tblVendors WHERE VendorCode=@VendorCode " '
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
                'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close() 'cierra la conexion
                If TN.Rows.Count > 0 Then
                    Dim Country As String = ""
                    Dim Other As Integer = 0
                    Dim Mex As Integer = 0
                    For NM As Integer = 0 To TN.Rows.Count - 1
                        Country = TN.Rows(NM).Item("Country").ToString.ToUpper
                        If Country = "MEX" Or Country = "MX" Or Country = "MEXICO" Then
                            Mex += 1
                            Resp = "MX"
                        Else
                            Other += 1
                            Resp = Country
                        End If
                        If Other > 0 Then
                            Resp = Country
                        End If
                    Next
                End If
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading tblVendors to fill combo Vendors")
            End Try
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close() 'cierra la conexion
        End Using
        Return Resp
    End Function
    '
    Private Function BuscaQtyOnOrderEnTblItemsPODet(ByVal PN As String, ByVal RequieredDate As String)
        Dim Resp As Decimal = 0
        Using TN As New Data.DataTable
            Dim Edo As String = ""
            'Dim Query As String = "SELECT Sum(QtyBalance) AS QtyBalance FROM tblItemsPOsDet WHERE PN=@PN AND JuarezReceivedDate IS NULL AND QtyBalance>0 AND DueDate<@DueDate"
            Dim Query As String = "SELECT tblItemsPOsDet.PN, SUM(tblItemsPOsDet.QtyBalance) AS QtyBalance FROM tblItemsPOs INNER JOIN tblItemsPOsDet ON tblItemsPOs.IDPO = tblItemsPOsDet.IDPO WHERE (tblItemsPOsDet.PN=@PN) AND (tblItemsPOs.Status = N'Open') AND (tblItemsPOsDet.QtyBalance > 0) AND (tblItemsPOsDet.JuarezReceivedDate IS NULL) AND (tblItemsPOsDet.DueDate<@DueDate) GROUP BY tblItemsPOsDet.PN"
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cmd.Parameters.Add("@DueDate", SqlDbType.Date).Value = CDate(RequieredDate)
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count > 0 Then
                Resp = CDec(Val(TN.Rows(0).Item("QtyBalance").ToString))
            End If
        End Using
        Return Resp
    End Function
    '
    Private Function BuscaTotalQtyOnOrderEnTblItemsPODet(ByVal PN As String)
        Dim Resp As Decimal = 0
        Using TN As New Data.DataTable
            Dim Edo As String = ""
            Dim Query As String = "SELECT tblItemsPOsDet.PN, SUM(tblItemsPOsDet.QtyBalance) AS QtyBalance FROM tblItemsPOs INNER JOIN tblItemsPOsDet ON tblItemsPOs.IDPO = tblItemsPOsDet.IDPO WHERE (tblItemsPOsDet.PN=@PN) AND (tblItemsPOs.Status = N'Open') AND (tblItemsPOsDet.QtyBalance > 0) GROUP BY tblItemsPOsDet.PN"
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                'cmd.Parameters.Add("@DueDate", SqlDbType.Date).Value = CDate(RequieredDate)
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count > 0 Then
                Resp = CDec(Val(TN.Rows(0).Item("QtyBalance").ToString))
            End If
        End Using
        Return Resp
    End Function
    '
    Private Function BuscaPNEnPODet(ByVal PN As String, ByVal RequieredDate As Date, ByVal KindPurchasing As Boolean, ByVal UMInputSHP As String, ByVal QtyInputSHP As Double, ByVal MOQ As Double, ByVal StandarPack As Double, ByVal UMVendorX As String, ByVal UMInputSHPX As String)
        Dim Respuesta As Decimal = 0
        Dim Edo As String = ""
        If PN = "CE-J301" Then
            PN = PN
        End If
        Using TN As New System.Data.DataTable
            'Dim FirstDayWeek As String = CalculaCualEsElLunes(RequieredDate)
            'Dim DueDate As String = RequieredDate
            Dim WeekDueDate As String = Semanas(CDate(RequieredDate))
            Dim FechaInicio As String = CalculaCualEsElLunes(RequieredDate)
            Dim FechaHasta As String = CalculaCualEsElDomingo(RequieredDate)
            Dim SubPN As String = ""
            Dim LeadTime As Integer = 0
            'Dim Reserved As Double = 0
            Dim Vendor As String = ""
            Dim VendorCode As String = ""
            Dim VendorPN As String = ""
            Dim PackPrice As Double = 0
            Dim UnitPrice As Double = 0
            'Dim MOQ As Double = 0
            'Dim StandarPack As Long = 0
            Dim BinBalance As Boolean = False
            Dim KindPurchasingK As Boolean = False
            Dim UMVendorK As String = ""
            Dim UMInputSHPk As String = ""
            Dim QtyInputSHPk As Double = 0
            Dim ExactlyQuantity As Boolean = False
            Dim Description As String = ""
            Dim Ky As String = ""
            Dim Unit As String = ""
            Dim QtyOrdered As Double = 0
            Dim QtyOnOrder As Double = 0
            Dim QtyBalance As Double = 0
            Dim TotalQty As Double = 0
            Dim Faltante As Double = 0
            Dim UnitK As String = ""
            Dim UM As String = ""
            'Dim Query As String = "SELECT tblItemsPOsDet.*,tblItemsQB.Unit AS UM, tblItemsQB.KindPurchasing, tblItemsQB.UMVendor, tblItemsQB.UMInputSHP, tblItemsQB.QtyInputSHP FROM tblItemsPOsDet INNER JOIN tblItemsQB ON tblItemsPOsDet.SubPN = tblItemsQB.SubPN WHERE (tblItemsPOsDet.PN = @PN) AND (tblItemsPOsDet.DueDate BETWEEN @FechaInicio AND @FechaHasta)"
            Dim Query As String = "SELECT tblItemsPOsDet.ID, tblItemsPOsDet.IDPO, tblItemsPOsDet.PN, tblItemsPOsDet.SubPN, tblItemsPOsDet.VendorPN, tblItemsPOsDet.Description, tblItemsPOsDet.VendorCode, tblItemsPOsDet.QtyOrdered, tblItemsPOsDet.QtyBalance, tblItemsPOsDet.QtyReceivedEP, tblItemsPOsDet.QtyReceivedJuarez, tblItemsPOsDet.Unit, tblItemsPOsDet.UnitPriceUSD, tblItemsPOsDet.UnitPriceMXN, tblItemsPOsDet.UnitPrice, tblItemsPOsDet.Amount, tblItemsPOsDet.DueDate, tblItemsPOsDet.EPReceivedBy, tblItemsPOsDet.EPReceivedDate, tblItemsPOsDet.EPDueDate, tblItemsPOsDet.JuarezReceivedBy, tblItemsPOsDet.JuarezReceivedDate, tblItemsPOsDet.JuarezDueDate, tblItemsPOsDet.CreatedBy, tblItemsPOsDet.CreatedDate, tblItemsPOsDet.ModifyBy, tblItemsPOsDet.ModifyDate, tblItemsPOsDet.Importation, tblItemsPOsDet.IDReferenceMRP, tblItemsPOsDet.ImportationNumber, tblItemsPOsDet.DepartmentUse, tblItemsPOsDet.Reason, tblItemsPOsDet.ItemRow, tblItemsPOsDet.MasterList, tblItemsPOsDet.Account, tblItemsPOsDet.AccountName, tblItemsPOsDet.Payment, tblItemsPOsDet.PayDate, tblItemsPOsDet.PayBy, tblItemsPOsDet.PR, tblItemsQB.Unit AS UM, tblItemsQB.KindPurchasing, tblItemsQB.UMVendor, tblItemsQB.UMInputSHP, tblItemsQB.QtyInputSHP, tblItemsPOs.Status FROM tblItemsPOsDet INNER JOIN tblItemsQB ON tblItemsPOsDet.SubPN = tblItemsQB.SubPN INNER JOIN tblItemsPOs ON tblItemsPOsDet.IDPO = tblItemsPOs.IDPO WHERE (tblItemsPOsDet.PN = @PN) AND (tblItemsPOs.Status = N'Open') AND (CAST(tblItemsPOsDet.DueDate AS DATE) BETWEEN @FechaInicio AND @FechaHasta) "
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cmd.Parameters.Add("@FechaInicio", SqlDbType.Date).Value = FechaInicio
                cmd.Parameters.Add("@FechaHasta", SqlDbType.Date).Value = FechaHasta
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                If TN.Rows.Count > 0 Then
                    For NM As Long = 0 To TN.Rows.Count - 1
                        UM = TN.Rows(NM).Item("UM").ToString
                        SubPN = TN.Rows(NM).Item("SubPN").ToString.ToUpper
                        VendorCode = TN.Rows(NM).Item("VendorCode").ToString.ToUpper
                        VendorPN = TN.Rows(NM).Item("VendorPN").ToString.ToUpper
                        UnitPrice = System.Convert.ToDouble(Val(TN.Rows(NM).Item("UnitPrice").ToString))
                        UnitK = TN.Rows(NM).Item("Unit").ToString
                        KindPurchasingK = CBool(TN.Rows(NM).Item("KindPurchasing"))
                        UMVendorK = TN.Rows(NM).Item("UMVendor").ToString
                        UMInputSHPk = TN.Rows(NM).Item("UMInputSHP").ToString
                        QtyInputSHPk = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyInputSHP").ToString))
                        Description = TN.Rows(NM).Item("Description").ToString
                        QtyOrdered += System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyOrdered").ToString))
                        QtyBalance = System.Convert.ToDecimal(Val(TN.Rows(NM).Item("QtyBalance").ToString))
                    Next
                End If
                Respuesta = QtyBalance
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading PN data, BuscaPartNumberEnItemsQB function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        Return Respuesta
    End Function
    '
    Private Function BuscaRegistroEnUnaTabla(ByVal PN As String, ByVal Tabla As String, ByVal Campo As String, ByVal SubPN As String, ByVal FirstDayWeek As String)
        Dim Resp As Decimal = 0
        Using TN As New Data.DataTable
            Dim Edo As String
            Try
                Dim Query As String = "SELECT " + Campo + " FROM " + Tabla + " WHERE PN=@PN ORDER BY ID DESC"
                Select Case Tabla
                    Case "tblPurchasingMaterialRequirementsPlanningForecast"
                        Query = "SELECT " + Campo + " FROM " + Tabla + " WHERE PN=@PN ORDER BY IDMaterialPurchasing DESC"
                    Case "tblPurchasingTempMRPForecast" + sTempTableName
                        Query = "SELECT " + Campo + " FROM " + Tabla + " WHERE PN=@PN ORDER BY ID DESC"
                    Case "tblPurchasingTempMRP2Forecast" + sTempTableName
                        Query = "SELECT " + Campo + " FROM " + Tabla + " WHERE PN=@PN ORDER BY ID DESC"
                End Select
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count > 0 Then
                Select Case Campo
                    Case "QtyAcum"
                        Resp = CDec(Val(TN.Rows(0).Item("QtyAcum").ToString))
                        If Resp > 0 Then
                            Resp = Resp
                        End If
                    Case ""

                End Select
            End If
        End Using
        Return Resp
    End Function
    'Funcion para encontrar el ultimo numero de referencia registrado en la base de datos
    Private Function BuscaNumeroDeReferenciaMRP()
        Dim Edo As String = ""
        Using TN As New System.Data.DataTable 'Despliega los materiales 
            Dim PN As String = ""
            Dim Qty As Decimal = 0
            Dim UM As String = ""
            Dim Query As String = "SELECT TOP 1 IDReferenceMRP FROM tblPurchasingReferenceSerialNumberMRP ORDER BY IDReferenceMRP DESC "
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count = 0 Then Edo = ""
            If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDReferenceMRP").ToString
        End Using
        Return Edo
    End Function
    'Genera el Numero de serie de la tabla tblPurchasingMaterialRequirementsPlanning
    Private Function GeneraSerialMRP(ByVal PreviousSerial As String) As String
        Dim Numero, ascii1, ascii2, ascii3, ascii4 As Integer
        Dim NumeroString, Letras, letra1, letra2, letra3, letra4, NewSerial, TnewSerial As String
        NewSerial = ""
        PreviousSerial = Microsoft.VisualBasic.Mid(PreviousSerial, 2)
        Try
            If PreviousSerial <> "" Then
                Letras = Microsoft.VisualBasic.Left(PreviousSerial, 4)
                Numero = Convert.ToInt64(Microsoft.VisualBasic.Right(PreviousSerial, 7))
                If Numero < 9999999 Then
                    Numero = Numero + 1
                    NumeroString = Numero.ToString
                    If NumeroString.Length < 7 Then
                        For count As Integer = NumeroString.Length To 6
                            NumeroString = "0" + NumeroString
                        Next
                    End If
                    NewSerial = Letras + NumeroString
                ElseIf Numero = 9999999 Then
                    NumeroString = "0000001"
                    letra1 = Mid(Letras, 1, 1)
                    letra2 = Mid(Letras, 2, 1)
                    letra3 = Mid(Letras, 3, 1)
                    letra4 = Mid(Letras, 4, 1)
                    ascii1 = Asc(letra1)
                    ascii2 = Asc(letra2)
                    ascii3 = Asc(letra3)
                    ascii4 = Asc(letra4)
                    If ascii4 < 90 Then
                        ascii4 = ascii4 + 1
                    ElseIf ascii4 = 90 And ascii3 < 90 Then
                        ascii4 = 65
                        ascii3 = ascii3 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 < 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = ascii2 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 < 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = 65
                        ascii1 = ascii1 + 1
                    ElseIf ascii4 = 90 And ascii3 = 90 And ascii2 = 90 And ascii1 = 90 Then
                        ascii4 = 65
                        ascii3 = 65
                        ascii2 = 65
                        ascii1 = 65
                    End If
                    letra1 = Convert.ToChar(ascii1).ToString
                    letra2 = Convert.ToChar(ascii2).ToString
                    letra3 = Convert.ToChar(ascii3).ToString
                    letra4 = Convert.ToChar(ascii4).ToString
                    Letras = letra1 + letra2 + letra3 + letra4
                    NewSerial = Letras + NumeroString
                End If
            ElseIf PreviousSerial = "" Then
                Letras = "AAAA"
                NumeroString = "0000001"
                NewSerial = Letras + NumeroString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        TnewSerial = "R" + NewSerial
        Return TnewSerial
    End Function
    'funcion para actualizar los cambios cadena tblItemsQB
    Private Sub UpdatetblPurchasingTempMRPBySubPN(ByVal Campo As String, ByVal Dato As String, ByVal Tipo As String, ByVal SubPN As String)
        Dim edo As String
        Try 'Definimos el query del update
            Dim Query As String = "UPDATE tblPurchasingTempMRPForecast" + sTempTableName + " SET " + Campo + "=@Dato  WHERE SubPN=@SubPN"
            Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@SubPN", SqlDbType.NVarChar).Value = SubPN
            If Tipo = "Cadena" Then cmd.Parameters.Add("@Dato", SqlDbType.NVarChar).Value = Dato
            If Tipo = "Decimal" Then cmd.Parameters.Add("@Dato", SqlDbType.Float).Value = System.Convert.ToDouble(Val(Dato))
            If Tipo = "Entero" Then cmd.Parameters.Add("@Dato", SqlDbType.BigInt).Value = System.Convert.ToInt64(Val(Dato))
            If Tipo = "Booleano" Then cmd.Parameters.Add("@Dato", SqlDbType.Bit).Value = System.Convert.ToBoolean(Dato)
            'cmd.Parameters.Add("@ModifyBy", SqlDbType.NVarChar).Value = txbUser.Text
            'cmd.Parameters.Add("@ModifyDate", SqlDbType.DateTime).Value = Now
            cnn.Open() 'abre la conexion
            cmd.ExecuteNonQuery() 'realiza el query
            edo = cnn.State.ToString
            If edo = "Open" Then cnn.Close() 'cierra la conexion
        Catch ex As Exception
            edo = cnn.State.ToString
            If edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error traying to update the SubPN-" + SubPN.ToUpper + " Field " + Campo + " Data " + Dato + " in SEA", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    '
    Private Sub UpdateTblPurchasingTempMRP(ByVal Campo As String, ByVal Dato As String, ByVal TipoDato As String, ByVal ID As Integer)
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = ""
            Query = "UPDATE tblPurchasingTempMRPForecast" + sTempTableName + " SET " & Campo & "=@Dato WHERE ID=@ID"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            'cmd.Parameters.Add("@IDMaterialPurchasing", SqlDbType.NVarChar).Value = IDMaterialPurchasing
            If TipoDato = "Cadena" Then cmd.Parameters.Add("@Dato", SqlDbType.NVarChar).Value = Dato
            If TipoDato = "Entero" Then cmd.Parameters.Add("@Dato", SqlDbType.BigInt).Value = System.Convert.ToInt64(Val(Dato))
            If TipoDato = "Decimal" Then cmd.Parameters.Add("@Dato", SqlDbType.Decimal).Value = System.Convert.ToDecimal(Val(Dato))
            cmd.Parameters.Add("@ID", SqlDbType.BigInt).Value = ID
            cnn.Open()
            cmd.ExecuteNonQuery()
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error in the update of tblPurchasingTempMRP" + sTempTableName + ", ID: " + ID.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    'Funcion para encontrar el ultimo numero de referencia registrado en la base de datos
    Private Function BuscaNumeroDeReferencia()
        Dim Edo As String = ""
        Using TN As New System.Data.DataTable 'Despliega los materiales 
            Dim PN As String = ""
            Dim Qty As Decimal = 0
            Dim UM As String = ""
            Dim Query As String = "SELECT TOP 1 IDReferenceMRP FROM tblPurchasingReferenceSerialNumberMRP ORDER BY IDReferenceMRP DESC "
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count = 0 Then Edo = ""
            If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDReferenceMRP").ToString
        End Using
        Return Edo
    End Function
    '
    Private Sub InsertIDReferenceMRP(ByVal IDReferenceMRP As String, ByVal KindReference As String)
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = ""
            Query = "INSERT INTO tblPurchasingReferenceSerialNumberMRP (IDReferenceMRP, KindReference, CreatedBy, CreatedDate) VALUES (@IDReferenceMRP, @KindReference, @CreatedBy, @CreatedDate)"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@KindReference", SqlDbType.NVarChar).Value = KindReference
            cmd.Parameters.Add("@IDReferenceMRP", SqlDbType.NVarChar).Value = IDReferenceMRP
            cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = txbUser.Text
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = Now
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error of insert tblPurchasingReferenceSerialNumberMRP, IDReferenceMRP: " + IDReferenceMRP, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    '
    Private Sub RegistraMRP(ByVal IDReferenceMRP As String)
        Dim Edo As String = cnn.State.ToString
        Using TN As New System.Data.DataTable
            Try
                Dim PN, SubPN, UM, UMToBuy, StandarPack, VendorPN, VendorCode, Vendor, BinBalance, Description, PO, ExactlyQuantity, UMVendor, UMInputSHP, Ky As String ' RequieredDate, ProcessDate As String
                Dim Reserved, QtyOnHand, QtyOnOrder, QtyToBuy, QtyUser, Qty, UnitPrice, PackPrice, QtyPO, Difference, MOQ, QtyInputSHP, QtyOnOrderPerWeek, QtyAcum, TotalQty, Faltante, QtyOnOrderPerPeriod As Double
                Dim LeadTime, Week As Long
                Dim RequieredDate, FirstDayWeek As Date
                Dim KindPurchasing, Pecent10 As Boolean '
                Dim Opcion As String = ""
                'Dim Query As String = "SELECT Ky, PN, SubPN, Qty, UM, StandarPack, UnitPrice, PackPrice, LeadTime, VendorCode, Vendor, BinBalance, Description, MOQ, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, QtyInputSHP, QtyOnHand, QtyOnOrder FROM tblPurchasingTempMRP" +sTempTableName
                Dim Query As String = "SELECT * FROM tblPurchasingTempMRP" + sTempTableName
                'If rdoRequiered.Checked = True Then
                '    If rdoAllWeeks.Checked = True Then Query = "SELECT PN, SubPN, QtyOnHand, QtyOnOrder, QtyToBuy, QtyUser, UMToBuy AS UM, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID FROM tblPurchasingTempMRP" +sTempTableName+" WHERE Difference<0"
                '    If rdoSpecificDates.Checked = True Then Query = "SELECT PN, SubPN, QtyOnHand, QtyOnOrder, QtyToBuy, QtyUser, UMToBuy AS UM, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID FROM tblPurchasingTempMRP" +sTempTableName +" WHERE ((Difference<0) AND (RequieredDate BETWEEN @FechaInicio AND @FechaHasta))"
                'End If
                Try
                    Dim dr As SqlDataReader
                    Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                    'cmd.CommandType = CommandType.Text
                    'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                    'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                    cnn.Open()
                    dr = cmd.ExecuteReader
                    TN.Load(dr) ''Llena la tabla
                    Edo = cnn.State.ToString
                    cnn.Close()
                Catch ex As Exception
                    Edo = cnn.State.ToString
                    If Edo = "Open" Then cnn.Close()
                    'AQUI
                    'MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                Dim X, IDMaterialPurchasing As String
                If TN.Rows.Count > 0 Then
                    For NM As Long = 0 To TN.Rows.Count - 1
                        PN = TN.Rows(NM).Item("PN").ToString
                        SubPN = TN.Rows(NM).Item("SubPN").ToString
                        Reserved = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Reserved").ToString))
                        QtyOnHand = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyOnHand").ToString))
                        QtyOnOrder = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyOnOrder").ToString))
                        QtyToBuy = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyToBuy").ToString))
                        QtyUser = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyUser").ToString))
                        Qty = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Qty").ToString))
                        UM = TN.Rows(NM).Item("UM").ToString
                        UMToBuy = TN.Rows(NM).Item("UMToBuy").ToString
                        StandarPack = System.Convert.ToDouble(Val(TN.Rows(NM).Item("StandarPack").ToString))
                        UnitPrice = System.Convert.ToDouble(Val(TN.Rows(NM).Item("UnitPrice").ToString))
                        PackPrice = System.Convert.ToDouble(Val(TN.Rows(NM).Item("PackPrice").ToString))
                        LeadTime = System.Convert.ToInt64(Val(TN.Rows(NM).Item("LeadTime").ToString))
                        VendorPN = TN.Rows(NM).Item("VendorPN").ToString
                        VendorCode = TN.Rows(NM).Item("VendorCode").ToString
                        Vendor = TN.Rows(NM).Item("Vendor").ToString
                        BinBalance = TN.Rows(NM).Item("BinBalance").ToString
                        Description = TN.Rows(NM).Item("Description").ToString
                        PO = TN.Rows(NM).Item("PO").ToString
                        QtyPO = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyPO").ToString))
                        Difference = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Difference").ToString))
                        MOQ = System.Convert.ToDouble(Val(TN.Rows(NM).Item("MOQ").ToString))
                        KindPurchasing = CBool(TN.Rows(NM).Item("KindPurchasing").ToString)
                        ExactlyQuantity = TN.Rows(NM).Item("ExactlyQuantity").ToString
                        UMVendor = TN.Rows(NM).Item("UMVendor").ToString
                        UMInputSHP = TN.Rows(NM).Item("UMInputSHP").ToString
                        QtyInputSHP = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyInputSHP").ToString))
                        Ky = TN.Rows(NM).Item("Ky").ToString
                        FirstDayWeek = System.Convert.ToDateTime(TN.Rows(NM).Item("FirstDayWeek").ToString)
                        RequieredDate = System.Convert.ToDateTime(TN.Rows(NM).Item("RequieredDate").ToString)
                        Week = System.Convert.ToInt64(Val(TN.Rows(NM).Item("Week").ToString))
                        QtyOnOrderPerWeek = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyOnOrderPerWeek").ToString)) 'QtyAcum, Pecent10
                        QtyAcum = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyAcum").ToString))
                        Pecent10 = System.Convert.ToBoolean(TN.Rows(NM).Item("Pecent10").ToString)
                        TotalQty = System.Convert.ToDouble(Val(TN.Rows(NM).Item("TotalQty").ToString))
                        Faltante = System.Convert.ToDouble(Val(TN.Rows(NM).Item("Faltante").ToString))
                        QtyOnOrderPerPeriod = System.Convert.ToDouble(Val(TN.Rows(NM).Item("QtyOnOrderPerPeriod").ToString))
                        If QtyUser > 0 Then
                            QtyUser = QtyUser
                        End If
                        X = BuscaUltimoIDMaterialPurchasing()
                        IDMaterialPurchasing = GeneraSerialMRP(X)
                        InsertTablaTblPurchasingMaterialRequirementsPlanning(IDMaterialPurchasing, PN, Qty, UM, "Paso Dos", SubPN, LeadTime, Vendor, VendorCode, VendorPN, PackPrice, UnitPrice, MOQ, StandarPack, BinBalance, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, ExactlyQuantity, Ky, Description, QtyOnHand, QtyOnOrder, RequieredDate, FirstDayWeek, CInt(Val(Week)), CDec(Val(Reserved)), IDReferenceMRP, QtyOnOrderPerWeek, Difference, QtyToBuy, QtyUser, UMToBuy, QtyAcum, Pecent10, TotalQty, Faltante, QtyOnOrderPerPeriod)
                        'InsertTablaTblPurchasingMaterialRequirementsPlanning(PN, Qty, UM, IDReferenceMRP, IDMaterialPurchasing, Week, LeadTime, RequieredDate, FirstDayWeek, StandarPack, VendorPN, VendorCode, Vendor, BinBalance, Description, MOQ, KindPurchasing, UMInputSHP, QtyInputSHP, Ky, SubPN, PackPrice, UnitPrice)
                    Next
                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
    'Funcion para encontrar el ultimo numero de referencia registrado en la base de datos
    Private Function BuscaUltimoIDMaterialPurchasing()
        Dim Edo As String = ""
        Using TN As New System.Data.DataTable 'Despliega los materiales 
            Dim PN As String = ""
            Dim Qty As Decimal = 0
            Dim UM As String = ""
            Dim Query As String = "SELECT TOP 1 IDMaterialPurchasing FROM tblPurchasingMaterialRequirementsPlanning ORDER BY IDMaterialPurchasing DESC "
            Try
                Dim dr As SqlDataReader
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@Valor1", SqlDbType.NVarChar).Value = ValorStatus
                'cmd.Parameters.Add("@Category", SqlDbType.NVarChar).Value = TipoCategoria
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr) ''Llena la tabla
                Edo = cnn.State.ToString
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            If TN.Rows.Count = 0 Then Edo = ""
            If TN.Rows.Count > 0 Then Edo = TN.Rows(0).Item("IDMaterialPurchasing").ToString
        End Using
        Return Edo
    End Function
    'Inserta un numero de parte en la tabla temporal de los materiales
    Private Sub InsertTablaTblPurchasingMaterialRequirementsPlanning(ByVal IDMaterialPurchasing As String, ByVal PN As String, ByVal Qty As Double, ByVal UM As String, ByVal Task As String, ByVal SubPN As String, ByVal LeadTime As Integer, ByVal Vendor As String, ByVal VendorCode As String, ByVal VendorPN As String, ByVal PackPrice As Double, ByVal UnitPrice As Double, ByVal MOQ As String, ByVal StandarPack As Double, ByVal BinBalance As String, ByVal KindPurchasing As Boolean, ByVal UMVendor As String, ByVal UMInputSHP As String, ByVal QtyInputSHP As Double, ByVal ExactlyQuantity As String, ByVal Ky As String, ByVal Description As String, ByVal QtyOnHand As Double, ByVal QtyOnOrder As Double, ByVal RequieredDate As Date, ByVal FirstDayWeek As Date, ByVal Week As Long, ByVal Reserved As Double, ByVal IDReferenceMRP As String, ByVal QtyOnOrderPerWeek As Double, ByVal Difference As Double, ByVal QtyToBuy As Double, ByVal QtyUser As Double, ByVal UMToBuy As String, ByVal QtyAcum As Decimal, ByVal Pecent10 As Boolean, ByVal TotalQty As Decimal, ByVal Faltante As Decimal, ByVal QtyOnOrderPerPeriod As Decimal)
        Dim Edo As String = ""
        Try 'Definimos el query del insert
            Dim Query As String = ""
            'Query = "INSERT INTO tblPurchasingMaterialRequirementsPlanning (IDMaterialPurchasing, PN, SubPN, Qty, UM, IDReferenceMRP, CreatedBy, CreatedDate, Week, LeadTime, RequieredDate, FirstDayWeek, StandarPack, VendorPN, VendorCode, Vendor, BinBalance, Description, MOQ, KindPurchasing, UMInputSHP, QtyInputSHP, Ky, PackPrice, UnitPrice) VALUES (@IDMaterialPurchasing, @PN, @SubPN, @Qty, @UM, @IDReferenceMRP, @CreatedBy, @CreatedDate, @Week, @LeadTime, @RequieredDate, @FirstDayWeek, @StandarPack, @VendorPN, @VendorCode, @Vendor, @BinBalance, @Description, @MOQ, @KindPurchasing, @UMInputSHP, @QtyInputSHP, @Ky, @PackPrice, @UnitPrice)"
            Query = "INSERT INTO tblPurchasingMaterialRequirementsPlanning (IDMaterialPurchasing, PN, SubPN, QtyOnHand, QtyOnOrder, QtyToBuy, QtyUser, UMToBuy, Qty, UM, StandarPack, UnitPrice, PackPrice, LeadTime, VendorPN, VendorCode, Vendor, BinBalance, Description, Difference, IDReferenceMRP, MOQ, KindPurchasing, UMVendor, UMInputSHP, QtyInputSHP, Ky, RequieredDate, FirstDayWeek, Week, QtyOnOrderPerWeek, CreatedBy, CreatedDate, QtyAcum, Pecent10, TotalQty, Faltante, QtyOnOrderPerPeriod) VALUES (@IDMaterialPurchasing, @PN, @SubPN, @QtyOnHand, @QtyOnOrder, @QtyToBuy, @QtyUser, @UMToBuy, @Qty, @UM, @StandarPack, @UnitPrice, @PackPrice, @LeadTime, @VendorPN, @VendorCode, @Vendor, @BinBalance, @Description, @Difference, @IDReferenceMRP, @MOQ, @KindPurchasing, @UMVendor, @UMInputSHP, @QtyInputSHP, @Ky, @RequieredDate, @FirstDayWeek, @Week, @QtyOnOrderPerWeek, @CreatedBy, @CreatedDate, @QtyAcum, @Pecent10, @TotalQty, @Faltante, @QtyOnOrderPerPeriod)"
            Dim cmd As New SqlCommand(Query, cnn)
            'IDMaterialPurchasing, PN, Qty, UM, PO, QtyPO, Difference, IDReferenceMRP, CreatedBy, CreatedDate, Week, LeadTime, RequieredDate, 
            'ProcessDate, FirstDayWeek, StandarPack, VendorPN, VendorCode, Vendor, BinBalance, Description, MOQ, KindPurchasing, UMInputSHP, 
            'QtyInputSHP, Ky, PackPrice, UnitPrice
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@IDMaterialPurchasing", SqlDbType.NVarChar).Value = IDMaterialPurchasing
            cmd.Parameters.Add("@CreatedBy", SqlDbType.NVarChar).Value = txbUser.Text
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime).Value = Now
            cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = Qty
            cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
            cmd.Parameters.Add("@SubPN", SqlDbType.NVarChar).Value = SubPN
            cmd.Parameters.Add("@QtyOnHand", SqlDbType.Float).Value = QtyOnHand
            cmd.Parameters.Add("@QtyOnOrder", SqlDbType.Float).Value = QtyOnOrder
            cmd.Parameters.Add("@QtyToBuy", SqlDbType.Float).Value = QtyToBuy
            cmd.Parameters.Add("@QtyUser", SqlDbType.Float).Value = QtyUser
            cmd.Parameters.Add("@UM", SqlDbType.NVarChar).Value = UM 'UMToBuy
            cmd.Parameters.Add("@UMToBuy", SqlDbType.NVarChar).Value = UMToBuy
            cmd.Parameters.Add("@StandarPack", SqlDbType.Float).Value = StandarPack
            cmd.Parameters.Add("@PackPrice", SqlDbType.Decimal).Value = PackPrice
            cmd.Parameters.Add("@UnitPrice", SqlDbType.Decimal).Value = UnitPrice
            cmd.Parameters.Add("@LeadTime", SqlDbType.Int).Value = LeadTime
            cmd.Parameters.Add("@Vendor", SqlDbType.NVarChar).Value = Vendor
            cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
            cmd.Parameters.Add("@VendorPN", SqlDbType.NVarChar).Value = VendorPN
            cmd.Parameters.Add("@BinBalance", SqlDbType.Bit).Value = BinBalance
            cmd.Parameters.Add("@Description", SqlDbType.NVarChar).Value = Description
            cmd.Parameters.Add("@Difference", SqlDbType.Float).Value = Difference
            cmd.Parameters.Add("@IDReferenceMRP", SqlDbType.NVarChar).Value = IDReferenceMRP
            cmd.Parameters.Add("@MOQ", SqlDbType.Float).Value = MOQ
            cmd.Parameters.Add("@KindPurchasing", SqlDbType.Bit).Value = KindPurchasing
            cmd.Parameters.Add("@UMVendor", SqlDbType.NVarChar).Value = UMVendor
            cmd.Parameters.Add("@UMInputSHP", SqlDbType.NVarChar).Value = UMInputSHP
            cmd.Parameters.Add("@QtyInputSHP", SqlDbType.Float).Value = QtyInputSHP
            cmd.Parameters.Add("@Ky", SqlDbType.NVarChar).Value = Ky
            cmd.Parameters.Add("@RequieredDate", SqlDbType.Date).Value = RequieredDate
            cmd.Parameters.Add("@FirstDayWeek", SqlDbType.Date).Value = FirstDayWeek
            cmd.Parameters.Add("@Week", SqlDbType.NVarChar).Value = Week
            cmd.Parameters.Add("@QtyOnOrderPerWeek", SqlDbType.Float).Value = QtyOnOrderPerWeek
            cmd.Parameters.Add("@QtyAcum", SqlDbType.Decimal).Value = QtyAcum
            cmd.Parameters.Add("@Pecent10", SqlDbType.Decimal).Value = Pecent10
            cmd.Parameters.Add("@TotalQty", SqlDbType.Decimal).Value = TotalQty
            cmd.Parameters.Add("@Faltante", SqlDbType.Decimal).Value = Faltante
            cmd.Parameters.Add("@QtyOnOrderPerPeriod", SqlDbType.Decimal).Value = QtyOnOrderPerPeriod
            cnn.Open()
            cmd.ExecuteNonQuery()
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error en el insert de tblPurchasingMaterialRequirementsPlanning PN: " + PN, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    'Exportar a Excel
    Private Sub CreaExcel(ByVal NumeroDeReferencia As String)
        Dim Edo As String = cnn.State.ToString
        Using TN As New System.Data.DataTable
            Dim FromDate As Date = Now
            Dim ToDate As Date = Now
            If rdoAllWeeks.Checked Then
                FromDate = Now
                ToDate = Now.AddYears(10)
            End If
            If rdoSpecificDates.Checked = True Then
                FromDate = dtpFrom.Value
                ToDate = dtpTo.Value
            End If
            If rdoViewOnly.Checked = True Then
                NumeroDeReferencia = "Report " + Now.ToString("dd-MMM-yy")
            End If
            Try
                Dim Opcion As String = ""

                'Dim Query As String = "SELECT Ky, PN, SubPN, Qty, UM, StandarPack, UnitPrice, PackPrice, LeadTime, VendorCode, Vendor, BinBalance, Description, MOQ, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, QtyInputSHP, QtyOnHand, QtyOnOrder FROM tblPurchasingTempMRP" +sTempTableName
                Dim Query As String = "SELECT * FROM tblPurchasingTempMRP" + sTempTableName + " WHERE Qty>0 "
                'If rdoRequiered.Checked = True Then
                '    If rdoAllWeeks.Checked = True Then Query = "SELECT * FROM tblPurchasingTempMRP" +sTempTableName +" WHERE Difference<0"
                '    If rdoSpecificDates.Checked = True Then Query = "SELECT * FROM tblPurchasingTempMRP" +sTempTableName +" WHERE ((Difference<0) AND (RequieredDate BETWEEN @FechaInicio AND @FechaHasta))"
                'End If
                'If rdoNonRequiered.Checked = True Then
                '    If rdoAllWeeks.Checked = True Then Query = "SELECT * FROM tblPurchasingTempMRP" +sTempTableName +" WHERE (Difference>-0 OR Difference=0)"
                '    If rdoSpecificDates.Checked = True Then Query = "SELECT * FROM tblPurchasingTempMRP" +sTempTableName +" WHERE ((Difference>-0 OR Difference=0) AND (RequieredDate BETWEEN @FechaInicio AND @FechaHasta))"
                'End If
                If rdoAllWeeks.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRP" + sTempTableName + " WHERE Qty>0 AND "
                If rdoSpecificDates.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRP" + sTempTableName + " WHERE (RequieredDate BETWEEN @FechaInicio AND @FechaHasta) AND Qty>0 AND "
                Dim OpcionFiltro = cmbFilter.Text.ToString
                Select Case OpcionFiltro
                    Case "Only Primary Without Bin Balance"
                        Query += " BinBalance=0 AND PriOption=1  " 'AND Difference<0
                    Case "Only Primary With Bin Balance"
                        Query += " ((BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1))  " 'AND Difference<0
                    Case "All Without Bin Balance"
                        Query += " BinBalance=0  " 'AND Difference<0
                    Case "ALL"
                        Query += " (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
                    Case "Only Bin Balance"
                        Query += " BinBalance=1  " 'AND Difference<0
                        'Case "Only Primary Without Bin Balance"
                        '    Query += " BinBalance=0 AND PriOption=1 "
                        'Case "Only Primary With Bin Balance"
                        '    Query += " (BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1)"
                        'Case "All Without Bin Balance"
                        '    Query += " BinBalance=0 "
                        'Case "ALL"
                        '    Query += " (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
                        'Case "Only Bin Balance"
                        '    Query += " BinBalance=1 "
                End Select
                Dim Opcion2 As String = cmb10Percent.Text.ToUpper
                Select Case Opcion2
                    Case "ALL"
                        'No agrega Nada
                        'Query += " AND Pecent10=0"
                    Case "10%"
                        'Agrega una columna al where
                        Query += " AND Pecent10=1"
                End Select
                Query += " ORDER BY SubPN ASC, FirstDayWeek ASC"
                Try
                    Dim dr As SqlDataReader
                    Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add("@FechaInicio", SqlDbType.Date).Value = dtpFrom.Value
                    cmd.Parameters.Add("@FechaHasta", SqlDbType.Date).Value = dtpTo.Value
                    cnn.Open()
                    dr = cmd.ExecuteReader
                    TN.Load(dr) ''Llena la tabla
                    Edo = cnn.State.ToString
                    cnn.Close()
                Catch ex As Exception
                    Edo = cnn.State.ToString
                    If Edo = "Open" Then cnn.Close()
                    MessageBox.Show(ex.ToString.ToString + "Error loading materials with requierment, Muestra Materiales function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                If TN.Rows.Count > 0 Then
                    Dim KindPurchasing As String
                    '================================================================================================
                    'configuracion para el MS EXCEL
                    '================================================================================================
                    Dim ObjApp As New Excel.Application
                    Dim ObjBook As Excel._Workbook = ObjApp.Workbooks.Add() 'ObjApp.Workbooks.Open(Origen) 
                    Dim ObjHoja1 As Excel._Worksheet = ObjBook.Worksheets(1)
                    With ObjHoja1
                        'Nombre de la Hoja
                        .Name = NumeroDeReferencia
                        'Orientacion de la hoja
                        .PageSetup.Orientation = XlPageOrientation.xlLandscape
                        'Tipo de letra
                        .Range("A1", "Z1000").Font.Name = "Arial"
                        'Tamaño de la letra
                        .Range("A1", "Z1000").Font.Size = 9
                        If rdoSaveReport.Checked = True Then
                            .Range("C1").Value = NumeroDeReferencia
                            .Range("C1").Interior.Color = RGB(155, 194, 230)
                            .Range("C1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C1").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        End If
                        .Range("L1").Value = "From:"
                        .Range("L2").Value = "To:"
                        .Range("L1").Interior.Color = RGB(155, 194, 230)
                        .Range("L2").Interior.Color = RGB(155, 194, 230)
                        .Range("L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                        .Range("L2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                        .Range("L1").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("L2").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("M1").Value = FromDate
                        .Range("M2").Value = ToDate
                        .Range("M1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .Range("M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                        .Range("M1").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("M2").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("M1").Value = FromDate.ToString
                        .Range("M2").Value = ToDate.ToString
                        .Range("M1").NumberFormat = "dd/MMM/yy"
                        .Range("M2").NumberFormat = "dd/MMM/yy" 'Format(ToDate, "dddd, mmm d yyyy")
                        .Range("A3").Value = "KY"
                        .Range("B3").Value = "PN"
                        .Range("C3").Value = "SubPN"
                        .Range("D3").Value = "Req Qty"
                        .Range("E3").Value = "Qty Acum"
                        .Range("F3").Value = "Diff"
                        .Range("G3").Value = "Stock Qty"
                        .Range("H3").Value = "On Order Qty"
                        .Range("I3").Value = "User Qty"
                        .Range("J3").Value = "UM"
                        .Range("K3").Value = "Standar Pack"
                        .Range("L3").Value = "Unit Price"
                        .Range("M3").Value = "Pack Price"
                        .Range("N3").Value = "Lead Time"
                        .Range("O3").Value = "Vendor"
                        .Range("P3").Value = "MOQ"
                        .Range("Q3").Value = "UM Vendor"
                        .Range("R3").Value = "UM Input SHP"
                        .Range("S3").Value = "Qty Input SHP"
                        .Range("T3").Value = "Week"
                        .Range("U3").Value = "Date Requiered"
                        .Range("V3").Value = "Date PO Requiered"
                        .Range("W3").Value = "10%"
                        '.Range("S3").Value = "Kind Purchasing"
                        '.Range("P3").Value = "Description"
                        '.Range("P3").Value = "Exactly Quantity"
                        '.Range("S3").Value = "Kind Purchasing"
                        .Range("A3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("B3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("C3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("D3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("E3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("F3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("G3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("H3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("I3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("J3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("K3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("L3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("M3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("N3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("O3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("P3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("Q3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("R3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("S3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("T3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("U3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("V3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("W3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        '.Range("X3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        '.Range("Y3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        '.Range("S3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("A3", "W3").Interior.Color = RGB(155, 194, 230)
                        .Range("A3:W3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .Range("A3:W3").BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                        .Range("A3:W3").AutoFilter(Field:=1, [Operator]:=Excel.XlAutoFilterOperator.xlFilterValues)
                        For NM As Integer = 0 To TN.Rows.Count - 1
                            .Range("A" + (NM + 4).ToString).Value = TN.Rows(NM).Item("Ky").ToString
                            .Range("B" + (NM + 4).ToString).Value = TN.Rows(NM).Item("PN").ToString
                            .Range("C" + (NM + 4).ToString).Value = TN.Rows(NM).Item("SubPN").ToString
                            .Range("D" + (NM + 4).ToString).Value = TN.Rows(NM).Item("Qty").ToString
                            .Range("D" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("E" + (NM + 4).ToString).Value = TN.Rows(NM).Item("QtyAcum").ToString
                            .Range("E" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("F" + (NM + 4).ToString).Value = TN.Rows(NM).Item("Difference").ToString
                            .Range("F" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("G" + (NM + 4).ToString).Value = TN.Rows(NM).Item("QtyOnHand").ToString
                            .Range("G" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("H" + (NM + 4).ToString).Value = TN.Rows(NM).Item("QtyOnOrder").ToString
                            .Range("H" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("I" + (NM + 4).ToString).Value = TN.Rows(NM).Item("QtyUser").ToString
                            .Range("I" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("J" + (NM + 4).ToString).Value = TN.Rows(NM).Item("UM").ToString
                            .Range("K" + (NM + 4).ToString).Value = TN.Rows(NM).Item("StandarPack").ToString
                            .Range("K" + (NM + 4).ToString).NumberFormat = "#,###,###.00"
                            .Range("L" + (NM + 4).ToString).Value = TN.Rows(NM).Item("UnitPrice").ToString
                            .Range("L" + (NM + 4).ToString).NumberFormat = "$ #,###,###.00"
                            .Range("M" + (NM + 4).ToString).Value = TN.Rows(NM).Item("PackPrice").ToString
                            .Range("M" + (NM + 4).ToString).NumberFormat = "$ #,###,###.00"
                            .Range("N" + (NM + 4).ToString).Value = TN.Rows(NM).Item("LeadTime").ToString
                            '.Range(L" + (NM + 4).ToString).Value = TN.Rows(NM).Item("VendorCode").ToString
                            .Range("O" + (NM + 4).ToString).Value = TN.Rows(NM).Item("VendorCode").ToString
                            ' Opcion = TN.Rows(NM).Item("BinBalance").ToString
                            ' If Opcion = "TRUE" Then .Range("L" + (NM + 4).ToString).Value = "Yes"
                            ' If Opcion = "FALSE" Then .Range("L" + (NM + 4).ToString).Value = ""
                            '.Range("M" + (NM + 4).ToString).Value = TN.Rows(NM).Item("Description").ToString
                            .Range("P" + (NM + 4).ToString).Value = TN.Rows(NM).Item("MOQ").ToString
                            'Opcion = TN.Rows(NM).Item("ExactlyQuantity").ToString.ToUpper
                            'If Opcion = "TRUE" Then .Range("O" + (NM + 4).ToString).Value = "No"
                            'If Opcion = "FALSE" Then .Range("O" + (NM + 4).ToString).Value = "Yes"
                            'Opcion = TN.Rows(NM).Item("KindPurchasing").ToString.ToUpper
                            'If Opcion = "TRUE" Then .Range("P" + (NM + 4).ToString).Value = "Other"
                            'If Opcion = "FALSE" Then .Range("P" + (NM + 4).ToString).Value = ""
                            KindPurchasing = TN.Rows(NM).Item("KindPurchasing").ToString
                            If KindPurchasing = "True" Then
                                .Range("Q" + (NM + 4).ToString).Value = TN.Rows(NM).Item("UMVendor").ToString
                                .Range("R" + (NM + 4).ToString).Value = TN.Rows(NM).Item("UMInputSHP").ToString
                                .Range("S" + (NM + 4).ToString).Value = TN.Rows(NM).Item("QtyInputSHP").ToString
                            End If
                            .Range("T" + (NM + 4).ToString).Value = TN.Rows(NM).Item("Week").ToString
                            .Range("U" + (NM + 4).ToString).Value = TN.Rows(NM).Item("FirstDayWeek").ToString 'Format(CDate(TN.Rows(NM).Item("FirstDayWeek").ToString), "dddd, mmm d yyyy")
                            .Range("V" + (NM + 4).ToString).Value = TN.Rows(NM).Item("RequieredDate").ToString 'Format(CDate(TN.Rows(NM).Item("RequieredDate").ToString), "dddd, mmm d yyyy")
                            .Range("U" + (NM + 4).ToString).NumberFormat = "dd/MMM/yy"
                            .Range("V" + (NM + 4).ToString).NumberFormat = "dd/MMM/yy"
                            .Range("X" + (NM + 4).ToString).Value = TN.Rows(NM).Item("Pecent10").ToString
                            '.Range("R" + (NM + 4).ToString).NumberFormat = ""
                            '.Range("Q" + (NM + 4).ToString).NumberFormat = "dd/mm/yy"
                            '.Range("R" + (NM + 4).ToString).NumberFormat = ""
                            .Range("A" + (NM + 4).ToString).BorderAround(XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("B" + (NM + 4).ToString).BorderAround(XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("C" + (NM + 4).ToString).BorderAround(XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("D" + (NM + 4).ToString).BorderAround(XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("E" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("F" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("G" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("H" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("I" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("J" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("K" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("L" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("M" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("N" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("O" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("P" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("Q" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("R" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("S" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("T" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("U" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("V" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("W" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            '.Range("S" + (NM + 4).ToString).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                            .Range("U" + (NM + 4).ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("V" + (NM + 4).ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        Next
                        .Columns.AutoFit()
                        If rdoSaveReport.Checked = True Then
                            Dim Destino As String = "\\10.17.182.22\Purchasing SHP\MRP\" + NumeroDeReferencia + ".XLSX" '
                            'Dim Destino As String = "C:\Users\julio.gallegos\Documents\PruebaCompras.XLSX"
                            If (File.Exists(Destino)) Then File.Delete(Destino)
                            'Grabamos el archivo
                            ObjBook.SaveAs(Destino)
                            'Cerramos excel y liberamos los objetos creados
                            ObjBook.Close(False)
                            ObjApp.Quit()
                        End If
                        If rdoViewOnly.Checked = True Then ObjApp.Application.Visible = True
                        releaseObject(ObjHoja1)
                        releaseObject(ObjBook)
                        releaseObject(ObjApp)
                        ObjApp = Nothing
                        ObjBook = Nothing
                        ObjHoja1 = Nothing
                    End With
                End If
                MessageBox.Show("Report Created Succesfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(ex.ToString + vbNewLine + "Error in excel function.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End Try
        End Using
    End Sub
    'Genera un archivo CSV
    Private Sub CreaCSV(ByVal NumeroDeReferencia)
        Dim Edo As String = ""
        Dim Opcion As String = ""
        'Dim Query As String = "SELECT * FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
        'If rdoAllWeeks.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
        'If rdoSpecificDates.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE (RequieredDate BETWEEN @FechaInicio AND @FechaHasta) AND Qty>0 "
        'Dim OpcionFiltro = cmbFilter.Text.ToString
        'Select Case OpcionFiltro
        '    Case "Only Primary Without Bin Balance"
        '        Query += " AND BinBalance=0 AND PriOption=1 AND Difference<0 "
        '    Case "Only Primary With Bin Balance"
        '        Query += " AND ((BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1)) AND Difference<0 "
        '    Case "All Without Bin Balance"
        '        Query += " AND BinBalance=0 AND Difference<0 "
        '    Case "ALL"
        '        Query += " AND (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
        '    Case "Only Bin Balance"
        '        Query += " AND BinBalance=1 AND Difference<0 "
        'End Select
        'Dim Opcion2 As String = cmb10Percent.Text.ToUpper
        'Select Case Opcion2
        '    Case "ALL"
        '        'No agrega Nada
        '        'Query += " AND Pecent10=0"
        '    Case "10%"
        '        'Agrega una columna al where
        '        Query += " AND Pecent10=1"
        'End Select
        'Query += " ORDER BY SubPN ASC, FirstDayWeek ASC"
        Using TFechas As New System.Data.DataTable
            Dim ContadorSemanas As Integer = 0
            Using TVendors As New System.Data.DataTable
                Using tPartNumbers As New System.Data.DataTable
                    Dim PN, ReqQty, UM, StandarPack, VendorCode, RequieredDate, FirstDayWeek As String
                    Dim BanderaRegistro As Integer = 0
                    Try
                        Dim Query2 As String = "SELECT DISTINCT FirstDayWeek FROM tblPurchasingTempMRPForecast" + sTempTableName + " ORDER BY FirstDayWeek ASC"
                        Dim cmd As SqlCommand = New SqlCommand(Query2, cnn)
                        Dim dr As SqlDataReader
                        'cmd.CommandType = CommandType.Text
                        'cmd.Parameters.Add("@ShipDate", SqlDbType.Date).Value = Fecha
                        cnn.Open()
                        dr = cmd.ExecuteReader
                        TFechas.Load(dr)
                        cnn.Close()
                    Catch ex As Exception
                        Edo = cnn.State.ToString
                        If Edo = "Open" Then cnn.Close() 'cierra la conexion
                        MessageBox.Show(ex.ToString, "Error in CreaCSV function") 'despliega un mesaje si hay un error
                    End Try
                    Try
                        Dim Query2 As String = "SELECT DISTINCT VendorCode FROM tblPurchasingTempMRPForecast" + sTempTableName + " ORDER BY VendorCode ASC"
                        Dim cmd As SqlCommand = New SqlCommand(Query2, cnn)
                        Dim dr As SqlDataReader
                        'cmd.CommandType = CommandType.Text
                        'cmd.Parameters.Add("@ShipDate", SqlDbType.Date).Value = Fecha
                        cnn.Open()
                        dr = cmd.ExecuteReader
                        TVendors.Load(dr)
                        cnn.Close()
                    Catch ex As Exception
                        Edo = cnn.State.ToString
                        If Edo = "Open" Then cnn.Close() 'cierra la conexion
                        MessageBox.Show(ex.ToString, "Error in CreaCSV function") 'despliega un mesaje si hay un error
                    End Try
                    If TFechas.Rows.Count > 0 Then
                        Dim ArchivoNombre As String = lblForecastReference.Text + ".csv"
                        Dim Path As String = "\\10.17.182.22\Purchasing SHP\MRP_Forecast\" + ArchivoNombre
                        Dim fs As FileStream = File.Create(Path)
                        Dim Cadena As String = "PN,Vendor,UM,Standar Pack,"
                        For RR As Integer = 0 To TFechas.Rows.Count - 1
                            Cadena += CDate(TFechas.Rows(RR).Item("FirstDayWeek").ToString).ToString("dd/MMM/yyyy") + ","
                        Next
                        Cadena += vbNewLine
                        Dim Titulos As Byte() = New UTF8Encoding(True).GetBytes(Cadena)
                        fs.Write(Titulos, 0, Titulos.Length)
                        For NM As Integer = 0 To TVendors.Rows.Count - 1
                            VendorCode = TVendors.Rows(NM).Item("VendorCode").ToString
                            tPartNumbers.Clear()
                            Try
                                Dim Query2 As String = "SELECT DISTINCT PN FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE VendorCode=@VendorCode "
                                Dim OpcionFiltro = cmbFilter.Text.ToString
                                Select Case OpcionFiltro
                                    Case "Only Primary Without Bin Balance"
                                        Query2 += " AND BinBalance=0 AND PriOption=1 " 'AND Difference<0 
                                    Case "Only Primary With Bin Balance"
                                        Query2 += " AND ((BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1))  " 'AND Difference<0 
                                    Case "All Without Bin Balance"
                                        Query2 += " AND BinBalance=0  "'AND Difference<0 
                                    Case "ALL"
                                        Query2 += " AND (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
                                    Case "Only Bin Balance"
                                        Query2 += " AND BinBalance=1  " 'AND Difference<0
                                End Select
                                Query2 += " ORDER BY PN ASC"
                                Dim cmd As SqlCommand = New SqlCommand(Query2, cnn)
                                Dim dr As SqlDataReader
                                cmd.CommandType = CommandType.Text
                                cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
                                cnn.Open()
                                dr = cmd.ExecuteReader
                                tPartNumbers.Load(dr)
                                cnn.Close()
                            Catch ex As Exception
                                Edo = cnn.State.ToString
                                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                                MessageBox.Show(ex.ToString, "Error in CreaCSV function") 'despliega un mesaje si hay un error
                            End Try
                            For RR As Integer = 0 To tPartNumbers.Rows.Count - 1
                                PN = tPartNumbers.Rows(RR).Item("PN").ToString
                                BanderaRegistro = 0
                                ContadorSemanas = ContadorSemanas
                                For kk As Integer = 0 To TFechas.Rows.Count - 1
                                    RequieredDate = CDate(TFechas.Rows(kk).Item("FirstDayWeek").ToString).ToString("dd/MMM/yyyy")
                                    Using Tdatos As New System.Data.DataTable
                                        If BanderaRegistro = 0 Then
                                            Try
                                                Dim Query As String = "SELECT * FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
                                                Query = "SELECT DISTINCT PN, UM, StandarPack FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE VendorCode=@VendorCode AND PN=@PN "
                                                'If rdoAllWeeks.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
                                                'If rdoSpecificDates.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE (RequieredDate BETWEEN @FechaInicio AND @FechaHasta) AND Qty>0 "
                                                Dim OpcionFiltro = cmbFilter.Text.ToString
                                                Select Case OpcionFiltro
                                                    Case "Only Primary Without Bin Balance"
                                                        Query += " AND BinBalance=0 AND PriOption=1  "'AND Difference<0
                                                    Case "Only Primary With Bin Balance"
                                                        Query += " AND ((BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1)) " 'AND Difference<0 
                                                    Case "All Without Bin Balance"
                                                        Query += " AND BinBalance=0  "'AND Difference<0
                                                    Case "ALL"
                                                        Query += " AND (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
                                                    Case "Only Bin Balance"
                                                        Query += " AND BinBalance=1  " 'AND Difference<0
                                                End Select
                                                Dim Opcion2 As String = cmb10Percent.Text.ToUpper
                                                Select Case Opcion2
                                                    Case "ALL"
                                                        'No agrega Nada
                                                        'Query += " AND Pecent10=0"
                                                    Case "10%"
                                                        'Agrega una columna al where
                                                        ' Query += " AND Pecent10=1"
                                                End Select
                                                Query += " ORDER BY PN ASC"
                                                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                                                Dim dr As SqlDataReader
                                                cmd.CommandType = CommandType.Text
                                                'cmd.Parameters.Add("@RequieredDate", SqlDbType.Date).Value = RequieredDate
                                                cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
                                                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                                                cnn.Open()
                                                dr = cmd.ExecuteReader
                                                Tdatos.Load(dr)
                                                cnn.Close()
                                            Catch ex As Exception
                                                Edo = cnn.State.ToString
                                                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                                                MessageBox.Show(ex.ToString, "Error in CreaCSV function") 'despliega un mesaje si hay un error
                                            End Try
                                        End If
                                        'If Tdatos.Rows.Count > 0 Then
                                        If BanderaRegistro = 0 And Tdatos.Rows.Count > 0 Then
                                            BanderaRegistro += 1
                                            'PN = Tdatos.Rows(0).Item("PN").ToString
                                            UM = Tdatos.Rows(0).Item("UM").ToString
                                            StandarPack = Tdatos.Rows(0).Item("StandarPack").ToString
                                            Cadena = PN + "," + VendorCode + "," + UM + "," + StandarPack + ","
                                        End If
                                        If Tdatos.Rows.Count > 0 Then
                                            ContadorSemanas = 0
                                            For PP As Integer = 0 To TFechas.Rows.Count - 1
                                                FirstDayWeek = CDate(TFechas.Rows(PP).Item("FirstDayWeek").ToString).ToString("dd/MMM/yyyy")
                                                Using TN As New System.Data.DataTable
                                                    Try
                                                        Dim Query As String = "SELECT * FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
                                                        Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 AND FirstDayWeek=@FirstDayWeek AND VendorCode=@VendorCode AND PN=@PN "
                                                        'If rdoAllWeeks.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE Qty>0 "
                                                        'If rdoSpecificDates.Checked = True Then Query = "SELECT PN, SubPN, Qty, QtyOnOrderPerWeek, UM AS [UM], QtyOnHand, QtyOnOrder, QtyUser, UMToBuy AS UMx, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID, QtyAcum, Pecent10, QtyReq, VendorPN, Vendor, BinBalance, PO, QtyPO, Difference, IDReferenceMRP, KindPurchasing, ExactlyQuantity, UMVendor, UMInputSHP, RequieredDate, TotalQty, Faltante, Orders, PriOption FROM tblPurchasingTempMRPForecast" + sTempTableName + " WHERE (RequieredDate BETWEEN @FechaInicio AND @FechaHasta) AND Qty>0 "
                                                        Dim OpcionFiltro = cmbFilter.Text.ToString
                                                        Select Case OpcionFiltro
                                                            Case "Only Primary Without Bin Balance"
                                                                Query += " AND BinBalance=0 AND PriOption=1 "'AND Difference<0 
                                                            Case "Only Primary With Bin Balance"
                                                                Query += " AND ((BinBalance=1 AND PriOption=1) OR (BinBalance=0 AND PriOption=1))  "'AND Difference<0
                                                            Case "All Without Bin Balance"
                                                                Query += " AND BinBalance=0  "'AND Difference<0
                                                            Case "ALL"
                                                                Query += " AND (BinBalance=1 OR BinBalance=0) AND (PriOption=0 OR PriOption=1) "
                                                            Case "Only Bin Balance"
                                                                Query += " AND BinBalance=1  " 'AND Difference<0
                                                        End Select
                                                        Dim Opcion2 As String = cmb10Percent.Text.ToUpper
                                                        Select Case Opcion2
                                                            Case "ALL"
                                                                'No agrega Nada
                                                                'Query += " AND Pecent10=0"
                                                            Case "10%"
                                                                'Agrega una columna al where
                                                                ' Query += " AND Pecent10=1"
                                                        End Select
                                                        Query += " ORDER BY SubPN ASC, FirstDayWeek ASC"
                                                        Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                                                        Dim dr As SqlDataReader
                                                        cmd.CommandType = CommandType.Text

                                                        'NO AGARRA LOS MESES EN ESPAÑOL / ELIMINA EL PUNTO Y CAMBIA EL NOMBRE DEL MES A INGLÉS
                                                        Dim day = FirstDayWeek.Replace(".", "")

                                                        For Each mes In mesesEspEn.Keys
                                                            If day.Contains(mes) Then
                                                                day = day.Replace(mes, mesesEspEn(mes)).Replace(".", "")
                                                                Exit For
                                                            End If
                                                        Next

                                                        'Console.WriteLine(day)
                                                        Dim fecha As DateTime = DateTime.ParseExact(day, "dd/MMM/yyyy", System.Globalization.CultureInfo.InvariantCulture)
                                                        Dim fechaSQL As String = fecha.ToString("yyyy-MM-dd")

                                                        cmd.Parameters.Add("@FirstDayWeek", SqlDbType.NVarChar).Value = fechaSQL


                                                        cmd.Parameters.Add("@VendorCode", SqlDbType.NVarChar).Value = VendorCode
                                                        cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                                                        cnn.Open()
                                                        dr = cmd.ExecuteReader
                                                        TN.Load(dr)
                                                        cnn.Close()
                                                    Catch ex As Exception
                                                        Edo = cnn.State.ToString
                                                        If Edo = "Open" Then cnn.Close() 'cierra la conexion
                                                        MessageBox.Show(ex.ToString, "Error in CreaCSV function") 'despliega un mesaje si hay un errorsssssssss

                                                    End Try
                                                    Try
                                                        ContadorSemanas += 1
                                                        If TN.Rows.Count = 0 Then
                                                            ReqQty = 0
                                                            Cadena += ReqQty + ","
                                                        ElseIf TN.Rows.Count > 0 Then
                                                            ReqQty = CStr(Val(TN.Rows(0).Item("Qty").ToString))
                                                            Cadena += ReqQty + ","
                                                        End If
                                                        'For JJ As Integer = 0 To TN.Rows.Count - 1
                                                        '    ReqQty = TN.Rows(JJ).Item("Qty").ToString
                                                        '    Cadena += ReqQty
                                                        'Next
                                                    Catch ex As Exception
                                                        MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                    End Try
                                                End Using
                                            Next
                                        End If
                                        'End If
                                    End Using
                                Next
                                Cadena += vbNewLine
                                Dim info As Byte() = New UTF8Encoding(True).GetBytes(Cadena)
                                fs.Write(info, 0, info.Length)
                            Next
                        Next
                        fs.Close()
                    End If
                    'If TN.Rows.Count > 0 Then
                    '    'Dim PN, SubPN, ReqQty, QtyAcum, Diff, StockQty, OnOrderQty, UserQty, UM, StandarPack, UnitPrice, PackPrice, LeadTime, Vendor, MOQ, UMVendor, UMInputSHP, QtyInputSHP, Week, DateRequiered, DatePORequiered, Percent10, KindPurchasing, FirstDayWeek, RequieredDate As String

                    '    'genera Archivo
                    '    'genera titulo del archivo
                    '    'PN, SubPN, ReqQty, QtyAcum, Diff, StockQty, OnOrderQty, UserQty, UM, StandarPack, UnitPrice, PackPrice, LeadTime, Vendor, MOQ, UMVendor, UMInputSHP, QtyInputSHP, Week, DateRequiered, DatePORequiered, Percent10

                    'End If
                    MessageBox.Show("The file was created into the folder \\BIMEXSERVER\Purchasing SHP\MRP_Forecast", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End Using
        End Using
    End Sub

    Private Sub BuscaWips(ByVal PN As String, ByVal RequieredDate As String, ByVal ForecastReference As String)
        Dim Edo As String = ""
        Dim Aprovado As Boolean = False
        Using TN As New System.Data.DataTable
            Dim FechaInicio As String = CalculaCualEsElLunes(RequieredDate)
            Dim FechaHasta As String = CalculaCualEsElDomingo(RequieredDate)
            Dim Query As String
            Dim BalancePN As Decimal = 0
            If ckbPastDue.Checked = True And rdoSpecificDates.Checked = True Then
                Dim EsteLunes As String = CalculaCualEsElLunes(CStr(Now))
                'Dim EsteDomingo As String = CalculaCualEsElDomingo(CStr(Now))
                If CDate(FechaInicio) < CDate(EsteLunes) Then FechaInicio = "1/Jan/1900"
            End If
            TN.Reset()
            Query = "SELECT tblBOMWIP.WIP, tblBOMWIP.AU, tblBOMWIP.Rev, tblBOMWIP.PN, tblBOMWIP.Balance, tblBOMWIP.Qty, tblBOMWIP.Unit AS UM, tblBOMWIP.PickList, tblBOMWIP.PO, tblBOMWIP.Week, tblBOMWIP.LeadTime, tblBOMWIP.RequieredDate, tblBOMWIP.ProcessDate, tblWIP.BalanceProcess, tblWIP.BalanceAssy, tblWIP.BalancePack, tblWIP.BalanceShipped, tblWIP.wSort, tblWIP.EstimatedStartDateProces, tblWIP.StartDateProces, tblWIP.DueDateProcess, tblWIP.DueDateAssy, tblWIP.DueDateShipped, tblWIP.DueDateCustomer, tblWIP.Qty AS QtyWIP, tblBOMWIP.Description FROM tblBOMWIP INNER JOIN tblWIP ON tblBOMWIP.WIP = tblWIP.WIP WHERE (((tblWIP.Status='Open') OR (tblWIP.Status='OPEN')) AND ((tblBOMWIP.PN=@PN) AND (tblBOMWIP.RequieredDate BETWEEN @FechaInicio AND @FechaHasta)))"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cmd.Parameters.Add("@FechaInicio", SqlDbType.Date).Value = FechaInicio
                cmd.Parameters.Add("@FechaHasta", SqlDbType.Date).Value = FechaHasta
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                Dim Contador As Long = TN.Rows.Count
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + "Error Loading PO from BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Dim UM As String = ""
            Dim SubPN As String = ""
            Dim Vendor As String = ""
            Dim VendorCode As String = ""
            Dim VendorPN As String = ""
            Dim PackPrice As String = ""
            Dim UnitPrice As String = ""
            Dim MOQ As String = ""
            Dim StandarPack As String = ""
            Dim BinBalance As String = ""
            Dim KindPurchasing As String = ""
            Dim UMInputSHP As String = ""
            Dim UMVendor As String = ""
            Dim QtyInputSHP As String = ""
            Dim ExactlyQuantity As String = ""
            Dim Ky As String = ""
            Dim Description As String = ""
            Dim Reserved As String = ""
            Dim TotalQty As Decimal = 0
            Dim Faltante As Decimal = 0
            Dim QtyOnOrderPerWeek As Double = 0
            Dim Difference As Double = 0
            Dim Mensaje As String = ""
            Dim BanderaMensaje As Integer = 0
            Dim UNidadDeMedida As String = ""
            GridWIP.DataSource = Nothing
            GridWIP.DataSource = TN
            lblRecordsWip.Text = "Records: " + TN.Rows.Count.ToString
            If TN.Rows.Count > 0 Then
                If GridWIP.RowCount > 0 Then
                    GridWIP.Columns("Qty").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("Balance").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("RequieredDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridWIP.Columns("ProcessDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridWip.Columns("Reserved").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("LeadTime").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("Week").DefaultCellStyle.Format = ("###,###.##")
                    'GridWip.Columns("AU").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("BalanceProcess").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("BalanceAssy").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("BalancePack").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("BalanceShipped").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("QtyWIP").DefaultCellStyle.Format = ("###,###.##")
                    GridWIP.Columns("StartDateProces").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridWIP.Columns("DueDateProcess").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridWIP.Columns("DueDateAssy").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridWIP.Columns("DueDateShipped").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridWIP.Columns("DueDateCustomer").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'Dim PNColumn As DataGridViewColumn = GridWip.Columns("PN")
                    'Dim SubPNColumn As DataGridViewColumn = GridWip.Columns("SubPN")
                    'Dim QtyOnHandColumn As DataGridViewColumn = GridWip.Columns("QtyOnHand")
                    'Dim QtyOnOrderColumn As DataGridViewColumn = GridWip.Columns("QtyOnOrder")
                    'Dim QtyOnOrderPerWeekColumn As DataGridViewColumn = GridWip.Columns("QtyOnOrderPerWeek")
                    'Dim QtyToBuyColumn As DataGridViewColumn = GridWip.Columns("QtyToBuy")
                    'Dim QtyUserColumn As DataGridViewColumn = GridWip.Columns("QtyUser")
                    'Dim UMColumn As DataGridViewColumn = GridWip.Columns("UM")
                    'Dim QtyColumn As DataGridViewColumn = GridWip.Columns("Qty")
                    'Dim UMReqColumn As DataGridViewColumn = GridWip.Columns("UM Req")
                    'Dim UnitPriceColumn As DataGridViewColumn = GridWip.Columns("UnitPrice")
                    'Dim PackPriceColumn As DataGridViewColumn = GridWip.Columns("PackPrice")
                    'Dim StandarPackColumn As DataGridViewColumn = GridWip.Columns("StandarPack")
                    'Dim MOQColumn As DataGridViewColumn = GridWip.Columns("MOQ")
                    'Dim LeadTimeColumn As DataGridViewColumn = GridWip.Columns("LeadTime")
                    'Dim VendorCodeColumn As DataGridViewColumn = GridWip.Columns("VendorCode")
                    'Dim DescriptionColumn As DataGridViewColumn = GridWip.Columns("Description")
                    'Dim FirstDayWeekColumn As DataGridViewColumn = GridWip.Columns("FirstDayWeek")
                    'Dim WeekColumn As DataGridViewColumn = GridWip.Columns("Week")
                    'Dim QtyInputSHPColumn As DataGridViewColumn = GridWip.Columns("QtyInputSHP")
                    'Dim KyColumn As DataGridViewColumn = GridWip.Columns("Ky")
                    'Dim IDColumn As DataGridViewColumn = GridWip.Columns("ID")
                    'PNColumn.Width = 90
                    'SubPNColumn.Width = 100
                    'QtyColumn.Width = 50
                    'QtyOnHandColumn.Width = 70
                    'QtyOnOrderColumn.Width = 70
                    'QtyOnOrderPerWeekColumn.Width = 70
                    'QtyToBuyColumn.Width = 50
                    'QtyUserColumn.Width = 50
                    'UMColumn.Width = 40
                    'UMReqColumn.Width = 25
                    'UnitPriceColumn.Width = 50
                    'PackPriceColumn.Width = 50
                    'StandarPackColumn.Width = 50
                    'MOQColumn.Width = 40
                    'LeadTimeColumn.Width = 30
                    'VendorCodeColumn.Width = 70
                    'DescriptionColumn.Width = 70
                    'FirstDayWeekColumn.Width = 70
                    'WeekColumn.Width = 35
                    'QtyInputSHPColumn.Width = 70
                    'KyColumn.Width = 30
                    'IDColumn.Width = 30
                End If
            End If
            GridWIP.AutoResizeColumns()
        End Using
    End Sub

    Private Sub BusquedaSalesOrders(ByVal Wipx As String)
        Dim Edo As String = ""
        Dim Aprovado As Boolean = False
        Using TN As New System.Data.DataTable
            'Dim FechaInicio As String = CalculaCualEsElLunes(RequieredDate)
            'Dim FechaHasta As String = CalculaCualEsElDomingo(RequieredDate)
            Dim Query As String
            Dim BalancePN As Decimal = 0
            TN.Reset()
            Query = "SELECT DISTINCT tblCustomerServiceSalesOrders.SONumber, tblCustomerServiceSalesOrders.AU, tblCustomerServiceSalesOrders.Rev, tblCustomerServiceSalesOrders.PO, tblCustomerServiceSalesOrders.PODate, tblCustomerServiceSalesOrders.DueDate, tblCustomerServiceSalesOrders.Qty, tblCustomerServiceSalesOrders.Balance, tblCustomerServiceSalesOrders.PackingSlipBalance, tblCustomerServiceSalesOrders.PN, tblCustomerServiceSalesOrders.Status, tblCustomerServiceSalesOrders.ItemRow, tblCustomerServiceSalesOrders.UnitPrice, tblCustomerServiceSalesOrders.Amount, tblCustomerServiceSalesOrders.CustomerCode, tblCustomerServiceSalesOrders.Customer, tblCustomerServiceSalesOrders.Description, tblCustomerServiceSalesOrders.ShipAddress1, tblCustomerServiceSalesOrders.ShipAddress2, tblCustomerServiceSalesOrders.ShipAddress3, tblCustomerServiceSalesOrders.ShipCity, tblCustomerServiceSalesOrders.ShipState, tblCustomerServiceSalesOrders.ShipCountry, tblCustomerServiceSalesOrders.ShipZip, tblCustomerServiceSalesOrders.Location FROM tblCustomerServiceSalesOrders INNER JOIN tblTicketOne ON tblCustomerServiceSalesOrders.IDQB = tblTicketOne.IDQBTemp WHERE  (tblTicketOne.WIP = @WIP)"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@WIP", SqlDbType.NVarChar).Value = Wipx
                'cmd.Parameters.Add("@FechaInicio", SqlDbType.Date).Value = FechaInicio
                'cmd.Parameters.Add("@FechaHasta", SqlDbType.Date).Value = FechaHasta
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                Dim Contador As Long = TN.Rows.Count
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + "Error Loading PO from BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            lblRecordsSalesOrder.Text = "Records: " + TN.Rows.Count.ToString
            GridSalesOrder.DataSource = TN
            If TN.Rows.Count > 0 Then
                If GridSalesOrder.RowCount > 0 Then
                    'GridSalesOrder.Columns("Qty").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("Balance").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("RequieredDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridSalesOrder.Columns("ProcessDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    ''GridSalesOrder.Columns("Reserved").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("LeadTime").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("Week").DefaultCellStyle.Format = ("###,###.##")
                    ''GridSalesOrder.Columns("AU").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("BalanceProcess").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("BalanceAssy").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("BalancePack").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("BalanceShipped").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("QtyWIP").DefaultCellStyle.Format = ("###,###.##")
                    'GridSalesOrder.Columns("StartDateProces").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridSalesOrder.Columns("DueDateProcess").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridSalesOrder.Columns("DueDateAssy").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridSalesOrder.Columns("DueDateShipped").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'GridSalesOrder.Columns("DueDateCustomer").DefaultCellStyle.Format = ("dd/MMM/yy")
                End If
            End If
            GridSalesOrder.AutoResizeColumns()
        End Using
    End Sub

    Private Sub BusquedaDePODet(ByVal PN As String, ByVal TipoBusqueda As String, ByVal DueDate As String)
        Dim Edo As String = ""
        Dim Aprovado As Boolean = False
        Using TN As New System.Data.DataTable
            Dim FechaInicio As String = CalculaCualEsElLunes(DueDate)
            Dim FechaHasta As String = CalculaCualEsElDomingo(DueDate)
            Dim Query As String = ""
            Dim BalancePN As Decimal = 0
            TN.Reset()
            If TipoBusqueda = "Todas" Then
                'Query = "SELECT tblItemsPOsDet.PN, tblItemsPOsDet.SubPN, tblItemsPOsDet.VendorPN, tblItemsPOsDet.VendorCode, tblItemsPOsDet.IDPO, tblItemsPOsDet.QtyOrdered, tblItemsPOsDet.QtyBalance, tblItemsPOsDet.QtyReceivedEP, tblItemsPOsDet.QtyReceivedJuarez, tblItemsPOsDet.Unit, UnitPrice, Amount, DueDate, EPReceivedDate, JuarezReceivedDate, EPDueDate, JuarezDueDate, Importation, IDReferenceMRP, ItemRow, Description FROM tblItemsPOsDet WHERE PN=@PN ORDER BY DueDate DESC"
                Query = "SELECT tblItemsPOsDet.PN, tblItemsPOsDet.SubPN, tblItemsPOsDet.VendorPN, tblItemsPOsDet.VendorCode, tblItemsPOsDet.IDPO, tblItemsPOsDet.QtyOrdered, tblItemsPOsDet.QtyBalance, tblItemsPOsDet.QtyReceivedEP, tblItemsPOsDet.QtyReceivedJuarez, tblItemsPOsDet.Unit, tblItemsPOsDet.UnitPrice, tblItemsPOsDet.Amount, tblItemsPOsDet.DueDate, tblItemsPOsDet.EPReceivedDate, tblItemsPOsDet.JuarezReceivedDate, tblItemsPOsDet.EPDueDate, tblItemsPOsDet.JuarezDueDate, tblItemsPOsDet.Importation, tblItemsPOsDet.IDReferenceMRP, tblItemsPOsDet.ItemRow, tblItemsPOsDet.Description, tblItemsPOs.Status FROM tblItemsPOsDet INNER JOIN tblItemsPOs ON tblItemsPOsDet.IDPO = tblItemsPOs.IDPO WHERE (tblItemsPOsDet.PN = @PN) ORDER BY tblItemsPOsDet.DueDate DESC"
            ElseIf TipoBusqueda = "Fechas" Then
                'Query = "SELECT PN, SubPN, VendorPN, VendorCode, IDPO, QtyOrdered, QtyBalance, QtyReceivedEP, QtyReceivedJuarez, Unit, UnitPrice, Amount, DueDate, EPReceivedDate, JuarezReceivedDate, EPDueDate, JuarezDueDate, Importation, IDReferenceMRP, ItemRow, Description FROM tblItemsPOsDet WHERE PN=@PN AND (DueDate BETWEEN @FechaInicio AND @FechaHasta) ORDER BY DueDate DESC"
                Query = "SELECT tblItemsPOsDet.PN, tblItemsPOsDet.SubPN, tblItemsPOsDet.VendorPN, tblItemsPOsDet.VendorCode, tblItemsPOsDet.IDPO, tblItemsPOsDet.QtyOrdered, tblItemsPOsDet.QtyBalance, tblItemsPOsDet.QtyReceivedEP, tblItemsPOsDet.QtyReceivedJuarez, tblItemsPOsDet.Unit, tblItemsPOsDet.UnitPrice, tblItemsPOsDet.Amount, tblItemsPOsDet.DueDate, tblItemsPOsDet.EPReceivedDate, tblItemsPOsDet.JuarezReceivedDate, tblItemsPOsDet.EPDueDate, tblItemsPOsDet.JuarezDueDate, tblItemsPOsDet.Importation, tblItemsPOsDet.IDReferenceMRP, tblItemsPOsDet.ItemRow, tblItemsPOsDet.Description, tblItemsPOs.Status FROM tblItemsPOsDet INNER JOIN tblItemsPOs ON tblItemsPOsDet.IDPO = tblItemsPOs.IDPO WHERE (tblItemsPOsDet.PN = @PN) AND (tblItemsPOsDet.DueDate BETWEEN @FechaInicio AND @FechaHasta) ORDER BY tblItemsPOsDet.DueDate DESC"
            End If
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                If TipoBusqueda = "Fechas" Then
                    cmd.Parameters.Add("@FechaInicio", SqlDbType.Date).Value = FechaInicio
                    cmd.Parameters.Add("@FechaHasta", SqlDbType.Date).Value = FechaHasta
                End If
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                Dim Contador As Long = TN.Rows.Count
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + "Error Loading PO from BuscaFaltantesEnLosWipActivos function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            GridPurchasingOrderItemsHistory.DataSource = Nothing
            GridPurchasingOrderItemsHistory.DataSource = TN
            lblRecordsPurchasingOrderItemsHistory.Text = "Records: " + TN.Rows.Count.ToString
            'QtyOrdered, QtyBalance, QtyReceivedEP, QtyReceivedJuarez, UnitPrice, Amount, DueDate, EPReceivedDate, JuarezReceivedDate, EPDueDate, JuarezDueDate
            If TN.Rows.Count > 0 Then
                If GridPurchasingOrderItemsHistory.RowCount > 0 Then
                    GridPurchasingOrderItemsHistory.Columns("QtyOrdered").DefaultCellStyle.Format = ("###,###.##")
                    GridPurchasingOrderItemsHistory.Columns("QtyBalance").DefaultCellStyle.Format = ("###,###.##")
                    GridPurchasingOrderItemsHistory.Columns("QtyReceivedEP").DefaultCellStyle.Format = ("###,###.##")
                    GridPurchasingOrderItemsHistory.Columns("QtyReceivedJuarez").DefaultCellStyle.Format = ("###,###.##")
                    'GridPurchasingOrderItemsHistory.Columns("AU").DefaultCellStyle.Format = ("###,###.##")
                    GridPurchasingOrderItemsHistory.Columns("UnitPrice").DefaultCellStyle.Format = ("$ ###,###.##")
                    GridPurchasingOrderItemsHistory.Columns("Amount").DefaultCellStyle.Format = ("$ ###,###.##")
                    GridPurchasingOrderItemsHistory.Columns("DueDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridPurchasingOrderItemsHistory.Columns("EPReceivedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridPurchasingOrderItemsHistory.Columns("JuarezReceivedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridPurchasingOrderItemsHistory.Columns("EPDueDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    GridPurchasingOrderItemsHistory.Columns("JuarezDueDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    'Dim PNColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("PN")
                    'Dim SubPNColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("SubPN")
                    'Dim QtyOnHandColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("QtyOnHand")
                    'Dim QtyOnOrderColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("QtyOnOrder")
                    'Dim QtyOnOrderPerWeekColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("QtyOnOrderPerWeek")
                    'Dim QtyToBuyColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("QtyToBuy")
                    'Dim QtyUserColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("QtyUser")
                    'Dim UMColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("UM")
                    'Dim QtyColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("Qty")
                    'Dim UMReqColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("UM Req")
                    'Dim UnitPriceColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("UnitPrice")
                    'Dim PackPriceColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("PackPrice")
                    'Dim StandarPackColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("StandarPack")
                    'Dim MOQColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("MOQ")
                    'Dim LeadTimeColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("LeadTime")
                    'Dim VendorCodeColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("VendorCode")
                    'Dim DescriptionColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("Description")
                    'Dim FirstDayWeekColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("FirstDayWeek")
                    'Dim WeekColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("Week")
                    'Dim QtyInputSHPColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("QtyInputSHP")
                    'Dim KyColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("Ky")
                    'Dim IDColumn As DataGridViewColumn = GridPurchasingOrderItemsHistory.Columns("ID")
                    'PNColumn.Width = 90
                    'SubPNColumn.Width = 100
                    'QtyColumn.Width = 50
                    'QtyOnHandColumn.Width = 70
                    'QtyOnOrderColumn.Width = 70
                    'QtyOnOrderPerWeekColumn.Width = 70
                    'QtyToBuyColumn.Width = 50
                    'QtyUserColumn.Width = 50
                    'UMColumn.Width = 40
                    'UMReqColumn.Width = 25
                    'UnitPriceColumn.Width = 50
                    'PackPriceColumn.Width = 50
                    'StandarPackColumn.Width = 50
                    'MOQColumn.Width = 40
                    'LeadTimeColumn.Width = 30
                    'VendorCodeColumn.Width = 70
                    'DescriptionColumn.Width = 70
                    'FirstDayWeekColumn.Width = 70
                    'WeekColumn.Width = 35
                    'QtyInputSHPColumn.Width = 70
                    'KyColumn.Width = 30
                    'IDColumn.Width = 30
                End If
            End If
            GridPurchasingOrderItemsHistory.AutoResizeColumns()
        End Using
    End Sub
    '=========================================================== Fin Funciones del calculo ====================================================================
    Private Sub LoginMRP()
        Dim entro As String = "OK"
        Try 'If entro = "OK" Then ' 
            If IsAuthenticated("SHPMFG", txbUserMRP.Text, txbUserMRPPassword.Text) Then
                Dim Autorizacion As String = AutorizacionDelUsuario(txbUserMRP.Text)
                If Autorizacion = "OK" Then
                    GroupBoxUploadFile.Visible = True
                    'TabControlMRPGlobal.Visible = True
                    GroupBoxUserMRP.Visible = False
                    txbUserMRPPassword.Text = ""
                    txbUser.Text = txbUserMRP.Text
                    Dim ExisteTabla As Integer = 0
                    Dim ExisteUsuario As Integer = 0
                    Do
                        'TruncateTablaTemp("tblPurchasingTempMRPForecast" + sTempTableName)
                        sTempTableName = GeneraNombreTabla()
                        ExisteUsuario = SelectTableDinamic(sTempTableName, txbUserMRP.Text, "MRP", "tblPurchasingTempMRPForecast" + sTempTableName, "ByUsuario") 'Revisa si existe el usuario y lo borra
                        ExisteTabla = SelectTableDinamic(sTempTableName, txbUserMRP.Text, "MRP", "tblPurchasingTempMRPForecast" + sTempTableName, "ByNombreGenerado") 'Revisa si existe la clave generada para la tabla
                    Loop Until ExisteTabla = 0
                    If ExisteTabla = 0 Then
                        CreatetblPurchasingTempMRPTable(sTempTableName)
                        InsertTableDinamic(sTempTableName, txbUser.Text, "MRP", "tblPurchasingTempMRPForecast" + sTempTableName)
                        BanderaLogin += 1
                        'TabControlMRPGlobal.Visible = True
                        'GroupBoxUserMRP.Visible = False
                        'llena el combo de los PN para la tabla temporal
                        CargaComboPNMyTable()
                        'If cmbPNMyTable.SelectedIndex > -1 Then
                        '    If cmbPNMyTable.SelectedValue.ToString <> "System.Data.DataRowView" Then
                        '        FindPNMyTable(cmbPNMyTable.SelectedValue.ToString)
                        '    End If
                        'End If
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    End If
                Else
                    txbUserMRPPassword.Text = ""
                    GroupBoxUploadFile.Visible = False
                    GroupBoxUserMRP.Visible = True
                    'TabControlMRPGlobal.Enabled = False
                End If
            Else
                MessageBox.Show("Usuario o contraseña incorrecto, por favor intente de nuevo " + vbNewLine + "User or password incorrect, Please try again.", "Authentication Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                txbUserMRPPassword.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function IsAuthenticated(ByVal domain As String, ByVal username As String, ByVal pwd As String) As Boolean
        Dim path As String = "LDAP://" & "10.17.182.22" ' domain
        Dim domainAndUsername As String = domain + "\" + username
        Dim entry As DirectoryEntry = New DirectoryEntry(path, domainAndUsername, pwd)
        Dim filterAttribute As String = ""
        Try
            'Bind to the native AdsObject to force authentication.
            Dim obj As Object = entry.NativeObject
            Dim search As DirectorySearcher = New DirectorySearcher(entry)
            search.Filter = "(SAMAccountName=" & username & ")"
            search.PropertiesToLoad.Add("cn")
            Dim result As SearchResult = search.FindOne()
            If (result Is Nothing) Then
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Function AutorizacionDelUsuario(ByVal Usuario As String)
        Dim Resp As String = "NO" 'tblItemsPOUserIDAuthorizations
        Dim Edo As String = ""
        Dim Query As String = "SELECT * FROM tblItemsPOUserIDAuthorizations WHERE UserID=@UserID AND Active = 1 AND Module = 'MRPLogin'"
        Using TN As New System.Data.DataTable
            Try
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                Dim DR As SqlDataReader
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@UserID", SqlDbType.NVarChar).Value = Usuario
                cnn.Open()
                DR = cmd.ExecuteReader
                TN.Load(DR)
                cnn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error in AutorizacionDelUsuario function") 'despliega un mesaje si hay un error
            End Try
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close() 'cierra la conexion
            If TN.Rows.Count > 0 Then
                For NM As Integer = 0 To TN.Rows.Count - 1
                    Dim Activo As String = TN.Rows(NM).Item("Active").ToString
                    Dim Modulo As String = TN.Rows(NM).Item("Module").ToString
                    If Modulo = "MRPLogin" Then
                        If Activo = "True" Then
                            Resp = "OK"
                        Else
                            MessageBox.Show("This user isn't active", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If
                Next
            End If
            If Resp = "NO" Then MessageBox.Show("This user isn't authorized to approve purchase orders", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
        Return Resp
    End Function

    Private Sub CreatetblPurchasingTempMRPTable(ByVal TemporalName As String)
        'Me.Text = "Creating Clients Table..."' 'Clients' " & _
        Dim strSQL As String =
        "USE [SEA] " & vbCrLf &
        "IF EXISTS (SELECT * FROM [SEA].dbo.sysobjects WHERE Name = 'tblPurchasingTempMRPForecast" & TemporalName & "' AND TYPE = 'u') " & vbCrLf &
        "BEGIN " & vbCrLf &
        "DROP TABLE [SEA].[dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] " & vbCrLf &
        "END " & vbCrLf &
        "/****** Object:  Table [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "]    Script Date: 3/16/2017 11:51:21 AM ******/ " & vbCrLf &
        "SET ANSI_NULLS ON " & vbCrLf &
        "SET QUOTED_IDENTIFIER ON " & vbCrLf &
        "CREATE TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "]( " & vbCrLf &
        "[PN] [nvarchar](100) NULL, " & vbCrLf &
        "[SubPN] [nvarchar](100) NULL, " & vbCrLf &
        "[Reserved] [decimal](38, 5) NULL, " & vbCrLf &
        "[QtyOnHand] [decimal](38, 5) NULL, " & vbCrLf &
        "[QtyOnOrder] [decimal](38, 5) NULL, " & vbCrLf &
        "[QtyToBuy] [decimal](38, 5) NULL, " & vbCrLf &
        "[QtyUser] [decimal](38, 5) NULL, " & vbCrLf &
        "[Qty] [decimal](38, 5) NULL, " & vbCrLf &
        "[QtyReq] [float] NULL, " & vbCrLf &
        "[UM] [nvarchar](100) NULL, " & vbCrLf &
        "[UMToBuy] [nvarchar](100) NULL, " & vbCrLf &
        "[StandarPack] [int] NULL, " & vbCrLf &
        "[UnitPrice] [decimal](38, 5) NULL, " & vbCrLf &
        "[PackPrice] [decimal](38, 5) NULL, " & vbCrLf &
        "[LeadTime] [int] NULL, " & vbCrLf &
        "[VendorPN] [nvarchar](100) NULL, " & vbCrLf &
        "[VendorCode] [nvarchar](100) NULL, " & vbCrLf &
        "[Vendor] [nvarchar](100) NULL, " & vbCrLf &
        "[BinBalance] [bit] NULL, " & vbCrLf &
        "[Description] [nvarchar](300) NULL, " & vbCrLf &
        "[PO] [bigint] NULL, " & vbCrLf &
        "[QtyPO] [decimal](38, 5) NULL, " & vbCrLf &
        "[Difference] [decimal](38, 5) NULL, " & vbCrLf &
        "[IDReferenceMRP] [nvarchar](100) NULL, " & vbCrLf &
        "[MOQ] [int] NULL, " & vbCrLf &
        "[KindPurchasing] [bit] NULL, " & vbCrLf &
        "[ExactlyQuantity] [bit] NULL, " & vbCrLf &
        "[UMVendor] [nvarchar](100) NULL, " & vbCrLf &
        "[UMInputSHP] [nvarchar](100) NULL, " & vbCrLf &
        "[QtyInputSHP] [bigint] NULL, " & vbCrLf &
        "[Ky] [nvarchar](10) NULL, " & vbCrLf &
        "[RequieredDate] [date] NULL, " & vbCrLf &
        "[FirstDayWeek] [date] NULL, " & vbCrLf &
        "[Week] [int] NULL, " & vbCrLf &
        "[ID] [int] IDENTITY(1,1) NOT NULL, " & vbCrLf &
        "[TotalQty] [decimal](38, 5) NULL, " & vbCrLf &
        "[Faltante] [decimal](38, 5) NULL, " & vbCrLf &
        "[QtyOnOrderPerWeek] [decimal](38, 5) NULL, " & vbCrLf &
        "[Orders] [bit] NULL, " & vbCrLf &
        "[PriOption] [bit] NULL, " & vbCrLf &
        "[QtyAcum] [decimal](38, 5) NULL, " & vbCrLf &
        "[Pecent10] [bit] NULL, " & vbCrLf &
        "[QtyOnOrderPerPeriod] [decimal](38, 5) NULL, " & vbCrLf &
        "CONSTRAINT [PK_tblPurchasingTempMRPForecast" & TemporalName & "] PRIMARY KEY CLUSTERED  " & vbCrLf &
        "([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_QtyOnHand]  DEFAULT ((0)) FOR [QtyOnHand] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_QtyOnOrder]  DEFAULT ((0)) FOR [QtyOnOrder] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_LeadTime_1]  DEFAULT ((0)) FOR [LeadTime] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_PO]  DEFAULT ((0)) FOR [PO] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_MOQ_1]  DEFAULT ((0)) FOR [MOQ] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_KindPurchasing_1]  DEFAULT ((0)) FOR [KindPurchasing] " & vbCrLf &
        "ALTER TABLE [dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] ADD  CONSTRAINT [DF_tblPurchasingTempMRPForecast" & TemporalName & "_QtyInputSHP_1]  DEFAULT ((0)) FOR [QtyInputSHP] " & vbCrLf
        Dim cmd As New SqlCommand(strSQL, cnn)
        cmd.CommandType = CommandType.Text
        Try
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As SqlException
            Dim Edo As String = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString, "Error")
        Finally
            cmd.Dispose()
        End Try
    End Sub

    Private Sub DroptblPurchasingTempMRPTable(ByVal TemporalName As String)
        'Me.Text = "Creating Clients Table..."' 'Clients' " & _
        Dim strSQL As String =
        "USE [SEA] " & vbCrLf &
        "IF EXISTS (SELECT * FROM [SEA].dbo.sysobjects WHERE Name = 'tblPurchasingTempMRPForecast" & TemporalName & "' AND TYPE = 'u') " & vbCrLf &
        "BEGIN " & vbCrLf &
        "DROP TABLE [SEA].[dbo].[tblPurchasingTempMRPForecast" & TemporalName & "] " & vbCrLf &
        "END " & vbCrLf
        Dim cmd As New SqlCommand(strSQL, cnn)
        cmd.CommandType = CommandType.Text
        Try
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
            DeleteUser(txbUser.Text, sTempTableName)
        Catch ex As SqlException
            Dim Edo As String = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString, "Error")
        Finally
            cmd.Dispose()
        End Try
    End Sub

    Private Function GeneraNombreTabla()
        Dim Nombre As String = ""
        Try
            Dim Posibles As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            Dim Longitud As Integer = Posibles.Length
            'Dim Letra As Char
            Dim Lugar As Integer
            Dim LongitudNuevaCadena As Integer = 0
            Do
                Randomize()
                Lugar = CInt(Int((Longitud * Rnd()) + 1))
                If Lugar = 0 Then Lugar += 1
                If Lugar = 1 Then
                    Nombre += Microsoft.VisualBasic.Left(Posibles, 1)
                    LongitudNuevaCadena += 1
                End If
                If Lugar = 36 Then
                    Nombre += Microsoft.VisualBasic.Right(Posibles, 1)
                    LongitudNuevaCadena += 1
                End If
                If Lugar > 1 And Lugar < 36 Then
                    Nombre += Mid(Posibles, Lugar, 1)
                    LongitudNuevaCadena += 1
                End If
            Loop Until LongitudNuevaCadena = 5
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return Nombre
    End Function

    Private Sub InsertTableDinamic(ByVal NombreGenerado As String, ByVal Usuario As String, ByVal Modulo As String, ByVal NombreTemporalTabla As String)
        Dim Edo As String = ""
        Try
            Dim Query As String = "INSERT INTO tblDinamicTables (NombreGenerado, Usuario, Modulo, NombreTemporalTabla, FechaDeCreado) VALUES (@NombreGenerado, @Usuario, @Modulo, @NombreTemporalTabla, @FechaDeCreado)"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@NombreGenerado", SqlDbType.NVarChar).Value = NombreGenerado
            cmd.Parameters.Add("@Usuario", SqlDbType.NVarChar).Value = Usuario
            cmd.Parameters.Add("@Modulo", SqlDbType.NVarChar).Value = Modulo
            cmd.Parameters.Add("@NombreTemporalTabla", SqlDbType.NVarChar).Value = NombreTemporalTabla
            cmd.Parameters.Add("@FechaDeCreado", SqlDbType.DateTime).Value = Now
            'cmd.Parameters.Add("@", SqlDbType.NVarChar).Value = Qty
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error in InsertTableDinamic")
        End Try
    End Sub

    Private Function SelectTableDinamic(ByVal NombreGenerado As String, ByVal Usuario As String, ByVal Modulo As String, ByVal NombreTemporalTabla As String, ByVal Busqueda As String)
        Dim Resp As Integer = 0
        Dim Edo As String = ""
        Dim Query As String = "SELECT * FROM tblDinamicTables WHERE NombreGenerado=@NombreGenerado AND Usuario, Modulo, NombreTemporalTabla"
        Select Case Busqueda
            Case "ByNombreGenerado"
                Query = "SELECT * FROM tblDinamicTables WHERE NombreGenerado=@NombreGenerado"
            Case "ByUsuario"
                Query = "SELECT * FROM tblDinamicTables WHERE Usuario=@Usuario"
        End Select
        Using TN As New System.Data.DataTable
            Try
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                Dim DR As SqlDataReader
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@NombreGenerado", SqlDbType.NVarChar).Value = NombreGenerado
                cmd.Parameters.Add("@Usuario", SqlDbType.NVarChar).Value = Usuario
                cmd.Parameters.Add("@NombreTemporalTabla", SqlDbType.NVarChar).Value = NombreTemporalTabla
                cnn.Open()
                DR = cmd.ExecuteReader
                TN.Load(DR)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error in SelectTableDinamic function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
            End Try
            Resp = TN.Rows.Count
            If Busqueda = "ByUsuario" And TN.Rows.Count > 0 Then
                Dim Codigo As String
                Dim Modulos As String
                For NM As Integer = 0 To TN.Rows.Count - 1
                    Codigo = TN.Rows(NM).Item("NombreGenerado").ToString
                    Modulos = TN.Rows(NM).Item("Modulo").ToString
                    If Codigo <> "" Then
                        If Modulos = "MRP" Then
                            DroptblPurchasingTempMRPTable(Codigo)
                            DeleteUser(Usuario, Codigo)
                        End If
                    End If
                Next
            End If
        End Using
        Return Resp
    End Function

    Private Sub DeleteUser(ByVal Usuario As String, ByVal NombreGenerado As String)
        Dim Edo As String = ""
        Try
            Dim Query As String = "DELETE FROM tblDinamicTables WHERE NombreGenerado=@NombreGenerado AND Modulo='MRP'"
            Dim cmd As New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@NombreGenerado", SqlDbType.NVarChar).Value = NombreGenerado
            'cmd.Parameters.Add("@Usuario", SqlDbType.NVarChar).Value = Usuario
            'cmd.Parameters.Add("@Modulo", SqlDbType.NVarChar).Value = Modulo
            'cmd.Parameters.Add("@NombreTemporalTabla", SqlDbType.NVarChar).Value = NombreTemporalTabla
            'cmd.Parameters.Add("@FechaDeCreado", SqlDbType.DateTime).Value = Now
            'cmd.Parameters.Add("@", SqlDbType.NVarChar).Value = Qty
            cnn.Open()
            cmd.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error in DeleteUser", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CargaComboAUWIP()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblAUBOMWIP.Clear()
        'Dim tblPNs As New DataTable
        'Dim Query As String = "SELECT DISTINCT AU FROM tblPurchasingWipFake  WHERE Forecastreference=@Forecastreference ORDER BY AU ASC"
        Dim Query As String = "SELECT DISTINCT AU FROM tblWIP WHERE Status='Open'"
        Try
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cmd = New SqlCommand(Query, cnn)
            'cmd.CommandType = CommandType.Text
            'cmd.Parameters.Add("@Forecastreference", SqlDbType.NVarChar).Value = lblForecastReference.Text
            'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
            cnn.Open()
            dr = cmd.ExecuteReader
            tblAUBOMWIP.Load(dr)
            cnn.Close()
            If tblAUBOMWIP.Rows.Count > 0 Then
                With cmbAUBOMWIP
                    .DataSource = tblAUBOMWIP
                    .DisplayMember = "AU"
                    .ValueMember = "AU"
                    ' .Text = tblItems.Rows(0).Item("ShipTo").ToString
                End With
            End If
            Dim Contador As Long = tblAUBOMWIP.Rows.Count
        Catch ex As Exception
            MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboAUWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Edo = cnn.State.ToString
        If Edo = "Open" Then cnn.Close() 'cierra la conexion
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboAUENG()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblAUBOMENG.Clear()
        'Dim tblPNs As New DataTable
        Dim Query As String = "SELECT DISTINCT AU FROM tblBOM ORDER BY AU ASC"
        Try
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cmd = New SqlCommand(Query, cnn)
            'cmd.CommandType = CommandType.Text
            'cmd.Parameters.Add("@Vendor", SqlDbType.NVarChar).Value = Vendor
            'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
            cnn.Open()
            dr = cmd.ExecuteReader
            tblAUBOMENG.Load(dr)
            cnn.Close()
            If tblAUBOMENG.Rows.Count > 0 Then
                With cmbAUBOMENG
                    .DataSource = tblAUBOMENG
                    .DisplayMember = "AU"
                    .ValueMember = "AU"
                End With
            End If
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close() 'cierra la conexion
            MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboAUENG function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboRevWIP(ByVal AU As Long)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblRevBOMWIP.Clear()
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT Rev FROM tblWIP WHERE Status='Open' AND AU=@AU ORDER BY Rev DESC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                'Dim R As DataRow = tblRevBOMWIP.NewRow
                'R.Item("Rev") = "ALL"
                'tblRevBOMWIP.Rows.Add(R)
                If TN.Rows.Count > 0 Then
                    For NM As Integer = 0 To TN.Rows.Count - 1
                        'Dim W As DataRow = tblRevBOMWIP.NewRow
                        'W.Item("Rev") = CStr(TN.Rows(NM).Item("Rev").ToString)
                        Edo = CStr(TN.Rows(NM).Item("Rev").ToString)
                        tblRevBOMWIP.Rows.Add(CStr(TN.Rows(NM).Item("Rev").ToString))
                    Next
                End If
                Dim XA As Integer = tblRevBOMWIP.Rows.Count
                'With cmbRevBOMWIP
                '    .DataSource = Nothing
                '    .DataSource = tblRevBOMWIP
                '    .DisplayMember = "Rev"
                '    .ValueMember = "Rev"
                'End With

                cmbRevBOMWIP.DataSource = Nothing
                cmbRevBOMWIP.DataSource = tblRevBOMWIP
                cmbRevBOMWIP.DisplayMember = "Rev"
                cmbRevBOMWIP.ValueMember = "Rev"
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboRevWIPForecast(ByVal AU As Long)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblRevBOMWIPForecastreference.Clear()
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT Rev FROM tblPurchasingWipFake WHERE AU=@AU AND Forecastreference=@Forecastreference ORDER BY Rev DESC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Forecastreference", SqlDbType.NVarChar).Value = lblForecastReference.Text
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                'Dim R As DataRow = tblRevBOMWIP.NewRow
                'R.Item("Rev") = "ALL"
                'tblRevBOMWIP.Rows.Add(R)
                If TN.Rows.Count > 0 Then
                    For NM As Integer = 0 To TN.Rows.Count - 1
                        'Dim W As DataRow = tblRevBOMWIP.NewRow
                        'W.Item("Rev") = CStr(TN.Rows(NM).Item("Rev").ToString)
                        Edo = CStr(TN.Rows(NM).Item("Rev").ToString)
                        tblRevBOMWIP.Rows.Add(CStr(TN.Rows(NM).Item("Rev").ToString))
                    Next
                End If
                Dim XA As Integer = tblRevBOMWIP.Rows.Count
                With cmbRevBOMForecast
                    .DataSource = Nothing
                    .DataSource = tblRevBOMWIP
                    .DisplayMember = "Rev"
                    .ValueMember = "Rev"
                End With
                'cmbRevBOMForecast.DataSource = Nothing
                'cmbRevBOMForecast.DataSource = tblRevBOMWIPForecastreference
                'cmbRevBOMForecast.DisplayMember = "Rev"
                'cmbRevBOMForecast.ValueMember = "Rev"
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboWIPForecast(ByVal AU As Long, ByVal Rev As String)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblWIPBOMWIPForecastreference.Clear()
        Dim R As DataRow = tblWIPBOMWIPForecastreference.NewRow
        R.Item("WIP") = "ALL"
        tblWIPBOMWIPForecastreference.Rows.Add(R)
        Using TN As New Data.DataTable
            Dim Complemento As String = ""
            If Rev <> "ALL" And Rev <> "" Then
                Complemento = " AND Rev=@Rev "
            End If
            Dim Query As String = "SELECT WIP FROM tblPurchasingWipFake WHERE Forecastreference=@Forecastreference AND AU=@AU " + Complemento + " ORDER BY WIP ASC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cmd.Parameters.Add("@Forecastreference", SqlDbType.NVarChar).Value = lblForecastReference.Text
                cnn.Open()
                dr = cmd.ExecuteReader
                tblWIPBOMWIPForecastreference.Load(dr)
                cnn.Close()
                'If TN.Rows.Count > 0 Then
                '    For NM As Integer = 0 To TN.Rows.Count - 1
                '        Edo = CStr(TN.Rows(NM).Item("WIP").ToString)
                '        tblWIPBOMWIP.Rows.Add(CStr(TN.Rows(NM).Item("WIP").ToString))
                '    Next
                'End If
                With cmbWIPBOMForecast 'cmbWIPBOMWIP
                    .DataSource = Nothing
                    .DataSource = tblWIPBOMWIPForecastreference
                    .DisplayMember = "WIP"
                    .ValueMember = "WIP"
                End With
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboRevENG(ByVal AU As String)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblRevBOMENG.Clear()
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT Rev FROM tblBOM WHERE AU=@AU ORDER BY Rev DESC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                cnn.Open()
                dr = cmd.ExecuteReader
                tblRevBOMENG.Load(dr)
                cnn.Close()
                'Dim R As DataRow = tblRevBOMENG.NewRow
                'R.Item("Rev") = "ALL"
                'tblRevBOMENG.Rows.Add(R)
                'If TN.Rows.Count > 0 Then
                '    For NM As Integer = 0 To TN.Rows.Count - 1
                '        tblRevBOMENG.Rows.Add(TN.Rows(NM).Item("Rev").ToString)
                '    Next
                'End If
                With cmbRevBOMENG
                    .DataSource = tblRevBOMENG
                    .DisplayMember = "Rev"
                    .ValueMember = "Rev"
                End With
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevENG function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboPNMyTable()
        'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblPNMyTable.Clear()
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT PN FROM tblPurchasingTempMRPForecast" + sTempTableName + " ORDER BY PN ASC" 'WHERE PN=@PN
            Try
                Dim R As DataRow = tblPNMyTable.NewRow
                R.Item("PN") = "ALL"
                tblPNMyTable.Rows.Add(R)
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                cnn.Open()
                dr = cmd.ExecuteReader
                tblPNMyTable.Load(dr)
                cnn.Close()
                'If TN.Rows.Count > 0 Then
                '    tblPNMyTable = TN.Copy()
                '    For NM As Integer = 0 To TN.Rows.Count - 1
                '        tblPNMyTable.Rows.Add(TN.Rows(NM).Item("PN").ToString)
                '    Next
                'End If
                With cmbPNMyTable
                    .DataSource = tblPNMyTable
                    .DisplayMember = "PN"
                    .ValueMember = "PN"
                End With
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboPNMyTable function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboRevSalesOrder(ByVal AU As Long)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblRevSalesOrder.Clear()
        Dim R As DataRow = tblRevSalesOrder.NewRow
        R.Item("Rev") = "ALL"
        tblRevSalesOrder.Rows.Add(R)
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT Rev FROM tblCustomerServiceSalesOrders WHERE AU=@AU ORDER BY Rev DESC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                cnn.Open()
                dr = cmd.ExecuteReader
                tblRevSalesOrder.Load(dr)
                cnn.Close()
                Dim XA As Integer = tblRevBOMWIP.Rows.Count
                With cmbRevSalesOrder
                    .DataSource = Nothing
                    .DataSource = tblRevSalesOrder
                    .DisplayMember = "Rev"
                    .ValueMember = "Rev"
                End With
                'cmbRevBOMWIP.DataSource = Nothing
                'cmbRevBOMWIP.DataSource = tblRevBOMWIP
                'cmbRevBOMWIP.DisplayMember = "Rev"
                'cmbRevBOMWIP.ValueMember = "Rev"
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    'WipByAU
    Private Sub CargaComboRevWipByAU(ByVal AU As Long)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblRevWipByAU.Clear()
        Dim R As DataRow = tblRevWipByAU.NewRow
        R.Item("Rev") = "ALL"
        tblRevWipByAU.Rows.Add(R)
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT Rev FROM tblWIP WHERE AU=@AU ORDER BY Rev DESC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
                cnn.Open()
                dr = cmd.ExecuteReader
                tblRevWipByAU.Load(dr)
                cnn.Close()
                Dim XA As Integer = tblRevWipByAU.Rows.Count
                With cmbRevWipByAU
                    .DataSource = Nothing
                    .DataSource = tblRevWipByAU
                    .DisplayMember = "Rev"
                    .ValueMember = "Rev"
                End With
                'cmbRevBOMWIP.DataSource = Nothing
                'cmbRevBOMWIP.DataSource = tblRevBOMWIP
                'cmbRevBOMWIP.DisplayMember = "Rev"
                'cmbRevBOMWIP.ValueMember = "Rev"
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboRevWipByAUForecastreference(ByVal AU As Long, ByVal Forecastreference As String)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblRevWipByAUForecastreference.Clear()
        Dim R As DataRow = tblRevWipByAUForecastreference.NewRow
        R.Item("Rev") = "ALL"
        tblRevWipByAUForecastreference.Rows.Add(R)
        Using TN As New Data.DataTable
            Dim Query As String = "SELECT DISTINCT Rev FROM tblPurchasingWipFake WHERE AU=@AU AND Forecastreference=@Forecastreference ORDER BY Rev DESC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Forecastreference", SqlDbType.NVarChar).Value = Forecastreference
                cnn.Open()
                dr = cmd.ExecuteReader
                tblRevWipByAUForecastreference.Load(dr)
                cnn.Close()
                Dim XA As Integer = tblRevWipByAUForecastreference.Rows.Count
                With cmbRevWIPForecast 'cmbRevWipByAu
                    .DataSource = Nothing
                    .DataSource = tblRevWipByAUForecastreference
                    .DisplayMember = "Rev"
                    .ValueMember = "Rev"
                End With
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWipByAUForecastreference function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub FindPNBOMENG(ByVal PN As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT PN, Qty, Unit,AU, Rev, Description, MaterialGroup, PercentIncrease, Reference, PickList, Route, Branch, CreatedBy, CreatedDate, ModifyBy, ModifyDate FROM tblBOM WHERE PN=@PN ORDER BY AU ASC, Rev DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMENG")
            End Try
            GridBOMENG.DataSource = TN
            GridBOMENG.AutoResizeColumns()
            lblRecordsBOMENG.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub FindPNBOMWIPForecast(ByVal PN As String, ByVal ForecastReference As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT tblPurchasingBOMWipFake.PN, tblPurchasingBOMWipFake.Qty, tblPurchasingBOMWipFake.Unit, tblPurchasingBOMWipFake.AU, tblPurchasingBOMWipFake.Rev, tblPurchasingBOMWipFake.WIP, tblPurchasingBOMWipFake.RequieredDate, tblPurchasingBOMWipFake.Description, tblPurchasingBOMWipFake.MaterialGroup, tblPurchasingBOMWipFake.PickList, tblPurchasingBOMWipFake.Route, tblPurchasingBOMWipFake.CreatedBy, tblPurchasingBOMWipFake.CreatedDate, tblPurchasingBOMWipFake.ForecastReference  FROM tblPurchasingBOMWipFake INNER JOIN tblPurchasingWipFake ON tblPurchasingBOMWipFake.WIP = tblPurchasingWipFake.WIP WHERE ((tblPurchasingWipFake.ForecastReference = @ForecastReference) AND (tblPurchasingBOMWipFake.PN = @PN)) ORDER BY tblPurchasingBOMWipFake.AU ASC, tblPurchasingBOMWipFake.Rev DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMWIP")
            End Try
            GridBOMWIP.DataSource = TN
            GridBOMWIP.AutoResizeColumns()
            lblRecordsBOMWIP.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub FindBOMENG(ByVal AU As String, ByVal Rev As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Dim ContRev As String = ""
            If Rev <> "ALL" Then
                ContRev = " AND Rev=@Rev "
            End If
            Try 'PN, Qty, Unit, AU, Rev, WIP, Description, MaterialGroup, PickList, Route, CreatedBy, CreatedDate, ModifyBy, ModifyDate
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT PN, Qty, Unit, AU, Rev, Description, MaterialGroup, Reference, PickList, Route, Branch, CreatedBy, CreatedDate, ModifyBy, ModifyDate FROM tblBOM WHERE AU=@AU " + ContRev + " ORDER BY AU ASC, Rev DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMENG")
            End Try
            GridBOMENG.DataSource = TN
            GridBOMENG.AutoResizeColumns()
            lblRecordsBOMENG.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub FindBOMWIP(ByVal AU As String, ByVal Rev As String, ByVal WIP As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try '
                Dim CompRev As String = ""
                Dim CompWip As String = ""
                If Rev <> "ALL" And Rev <> "System.Data.DataRowView" Then
                    CompRev = " AND tblBOMWIP.Rev=@Rev "
                End If
                If WIP <> "ALL" And WIP <> "System.Data.DataRowView" Then
                    CompWip = " AND tblBOMWIP.WIP=@WIP "
                End If
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader 'julio.gallegos
                Dim Query As String = "SELECT tblBOMWIP.PN, tblBOMWIP.Qty, tblBOMWIP.Balance, tblBOMWIP.Unit, tblBOMWIP.AU, tblBOMWIP.Rev, tblBOMWIP.WIP, tblBOMWIP.RequieredDate, tblBOMWIP.Description, tblBOMWIP.MaterialGroup, tblBOMWIP.PickList, tblBOMWIP.Route, tblBOMWIP.CreatedBy, tblBOMWIP.CreatedDate, tblBOMWIP.ModifyBy, tblBOMWIP.ModifyDate FROM tblBOMWIP INNER JOIN tblWIP ON tblBOMWIP.WIP = tblWIP.WIP WHERE (tblWIP.Status = N'Open') AND (tblBOMWIP.AU = @AU) " + CompRev + CompWip + "  ORDER BY tblBOMWIP.AU ASC, tblBOMWIP.Rev DESC, tblBOMWIP.WIP ASC, tblBOMWIP.PN ASC"
                'Query = "SELECT PN, Qty, Unit, AU, Rev, WIP, RequieredDate, Description, MaterialGroup, PickList, Route, CreatedBy, CreatedDate, ForecastReference FROM tblPurchasingBOMWipFake WHERE (ForecastReference =@ForecastReference) AND (AU = @AU) " + CompRev + CompWip + "  ORDER BY AU ASC, Rev DESC, WIP ASC, PN ASC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cmd.Parameters.Add("@WIP", SqlDbType.NVarChar).Value = WIP
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()

                GridBOMWIP.DataSource = TN
                GridBOMWIP.Columns("RequieredDate").DefaultCellStyle.Format = ("dd-MMM-yyyy")
                GridBOMWIP.AutoResizeColumns()
                lblRecordsBOMWIP.Text = "Records " + TN.Rows.Count.ToString
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMWIP")
            End Try

        End Using
    End Sub

    Private Sub FindBOMWIPForecast(ByVal AU As String, ByVal Rev As String, ByVal WIP As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try '
                Dim CompRev As String = ""
                Dim CompWip As String = ""
                If Rev <> "ALL" Then
                    CompRev = " AND Rev=@Rev "
                End If
                If WIP <> "ALL" Then
                    CompWip = " AND WIP=@WIP "
                End If
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader 'julio.gallegos
                Dim Query As String = "SELECT tblBOMWIP.PN, tblBOMWIP.Qty, tblBOMWIP.Balance, tblBOMWIP.Unit, tblBOMWIP.AU, tblBOMWIP.Rev, tblBOMWIP.WIP, tblBOMWIP.RequieredDate, tblBOMWIP.Description, tblBOMWIP.MaterialGroup, tblBOMWIP.PickList, tblBOMWIP.Route, tblBOMWIP.CreatedBy, tblBOMWIP.CreatedDate, tblBOMWIP.ModifyBy, tblBOMWIP.ModifyDate FROM tblBOMWIP INNER JOIN tblWIP ON tblBOMWIP.WIP = tblWIP.WIP WHERE (tblWIP.Status = N'Open') AND (tblBOMWIP.AU = @AU) " + CompRev + CompWip + "  ORDER BY tblBOMWIP.AU ASC, tblBOMWIP.Rev DESC, tblBOMWIP.WIP ASC, tblBOMWIP.PN ASC"
                Query = "SELECT PN, Qty, Unit, AU, Rev, WIP, RequieredDate, Description, MaterialGroup, PickList, Route, CreatedBy, CreatedDate, ForecastReference FROM tblPurchasingBOMWipFake WHERE (ForecastReference=@ForecastReference) AND (AU = @AU) " + CompRev + CompWip + "  ORDER BY AU ASC, Rev DESC, WIP ASC, PN ASC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cmd.Parameters.Add("@WIP", SqlDbType.NVarChar).Value = WIP
                cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = lblForecastReference.Text
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMWIP")
            End Try
            GridBOMForecast.DataSource = TN
            GridBOMForecast.Columns("RequieredDate").DefaultCellStyle.Format = ("dd-MMM-yyyy")
            GridBOMForecast.AutoResizeColumns()
            lblRecordsGridBOMForecast.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub SearchBOMWIP(ByVal ForecastReference As String)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Using TN As New Data.DataTable
            Dim B2 As String = ""
            If cmbAUBOMWIP.SelectedIndex > 0 And cmbRevBOMWIP.SelectedIndex > 0 Then
                B2 = " AND tblBOMWIP.Rev=@Rev"
            End If
            Using tb As New Data.DataTable
                Try
                    Dim cmd As SqlCommand
                    Dim dr As SqlDataReader
                    Dim Query As String = "SELECT * From tblBOM WHERE AU=@AU " + B2 + " ORDER BY PN "
                    'Query = "SELECT tblBOMWIP.PN, tblBOMWIP.Qty, tblBOMWIP.Balance, tblBOMWIP.Unit, tblBOMWIP.AU, tblBOMWIP.Rev, tblBOMWIP.WIP, tblBOMWIP.RequieredDate, tblBOMWIP.Description, tblBOMWIP.MaterialGroup, tblBOMWIP.Reference, tblBOMWIP.PickList, tblBOMWIP.Route, tblBOMWIP.CreatedBy, tblBOMWIP.CreatedDate, tblBOMWIP.ModifyBy, tblBOMWIP.ModifyDate FROM tblBOMWIP INNER JOIN tblWIP ON tblBOMWIP.WIP = tblWIP.WIP WHERE ((tblWIP.Status = N'Open') AND (tblBOMWIP.PN = @PN) AND (tblBOMWIP.Balance>0) AND (tblBOMWIP.AU) " + B2 + ") ORDER BY tblBOMWIP.AU ASC, tblBOMWIP.Rev DESC"
                    Query = "SELECT PN, Qty, Unit, AU, Rev, WIP, RequieredDate, Description, MaterialGroup, PickList, Route, CreatedBy, CreatedDate, ForecastReference FROM tblPurchasingBOMWipFake  WHERE ((PN = @PN) AND (ForecastReference=@ForecastReference) AND (AU=@AU) " + B2 + ") ORDER BY AU ASC, Rev DESC"
                    cmd = New SqlCommand(Query, cnn)
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = cmbAUBOMWIP.SelectedValue.ToString
                    cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = cmbRevBOMWIP.SelectedValue.ToString
                    cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
                    cnn.Open()
                    dr = cmd.ExecuteReader
                    tb.Load(dr)
                    cnn.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString + " Error to fill BOM ", "Critical Error")
                    cnn.Close()
                End Try
                'GridBom.DataSource = tb
                'GridBom.AutoResizeColumns()
                'lblRecords.Text = "Records: " + tb.Rows.Count.ToString
            End Using
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GeneraColumnasTablasBOM()
        Try
            Dim WC1 As DataColumn = tblRevBOMWIP.Columns.Add("Rev", Type.GetType("System.String"))
            'WC1.AllowDBNull = True
            Dim WC2 As DataColumn = tblRevBOMENG.Columns.Add("Rev", Type.GetType("System.String"))
            'WC2.AllowDBNull = True
            Dim WC3 As DataColumn = tblWIPBOMWIP.Columns.Add("WIP", Type.GetType("System.String"))
            'WC3.AllowDBNull = True
            Dim WC4 As DataColumn = tblPNMyTable.Columns.Add("PN", Type.GetType("System.String"))
            'WC4.AllowDBNull = True
            Dim WC5 As DataColumn = tblRevSalesOrder.Columns.Add("Rev", Type.GetType("System.String"))
            'WC5.AllowDBNull = True
            Dim WC6 As DataColumn = tblRevWipByAU.Columns.Add("Rev", Type.GetType("System.String"))
            'WC6.AllowDBNull = True
            Dim WC7 As DataColumn = tblRevBOMWIPForecastreference.Columns.Add("Rev", Type.GetType("System.String"))
            'WC7.AllowDBNull = True
            Dim WC8 As DataColumn = tblRevWipByAUForecastreference.Columns.Add("Rev", Type.GetType("System.String"))
            'WC8.AllowDBNull = True
            Dim WC9 As DataColumn = tblWIPBOMWIPForecastreference.Columns.Add("WIP", Type.GetType("System.String"))
            'WC9.AllowDBNull = True
            Dim WC10 As DataColumn = tblAUWIPForecastreference.Columns.Add("AU", Type.GetType("System.String"))
            'WC10.AllowDBNull = True
            'Dim WC11 As DataColumn = tblRevWipByAUForecastreference.Columns.Add("Rev", Type.GetType("System.String"))
            'WC11.AllowDBNull = True
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        'tblAUBOMWIP.Columns.Add("AU", Type.GetType("System.String"))
        'tblAUBOMENG.Columns.Add("AU", Type.GetType("System.String"))
    End Sub

    Private Sub FindPNMyTable(ByVal PN As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Dim ContPN As String = ""
            If PN <> "ALL" Then
                ContPN = " WHERE PN=@PN "
            End If
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT PN, SubPN, QtyAcum, Qty, QtyOnOrderPerWeek, QtyOnHand, QtyOnOrder, QtyUser, UM, UnitPrice, PackPrice, StandarPack, MOQ, LeadTime, VendorCode, Description, FirstDayWeek, Week, QtyInputSHP, Ky, ID FROM tblPurchasingTempMRPForecast" + sTempTableName + ContPN + " ORDER BY SubPN ASC, FirstDayWeek ASC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMWIP")
            End Try
            GridMyTable.DataSource = TN
            lblRecordsMyTable.Text = "Records: " + TN.Rows.Count.ToString
            If TN.Rows.Count > 0 Then
                If GridMRP.RowCount > 0 Then
                    With GridMyTable
                        .Columns("QtyAcum").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("Qty").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("QtyOnOrderPerWeek").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("FirstDayWeek").DefaultCellStyle.Format = ("dd/MMM/yy")
                        '.Columns("Reserved").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("QtyOnHand").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("QtyOnOrder").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("QtyOnHand").DefaultCellStyle.Format = ("###,###.##")
                        '.Columns("QtyToBuy").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("QtyUser").DefaultCellStyle.Format = ("###,###.##")
                        .Columns("UnitPrice").DefaultCellStyle.Format = ("$###,###.##")
                        .Columns("PackPrice").DefaultCellStyle.Format = ("$###,###.##")
                        'Dim PNColumn As DataGridViewColumn = GridMyTable.Columns("PN") 'QtyAcum
                        'Dim SubPNColumn As DataGridViewColumn = GridMyTable.Columns("SubPN")
                        'Dim QtyOnHandColumn As DataGridViewColumn = GridMyTable.Columns("QtyOnHand")
                        'Dim QtyOnOrderColumn As DataGridViewColumn = GridMyTable.Columns("QtyOnOrder")
                        'Dim QtyOnOrderPerWeekColumn As DataGridViewColumn = GridMyTable.Columns("QtyOnOrderPerWeek")
                        ''Dim QtyToBuyColumn As DataGridViewColumn = GridMyTable.Columns("QtyToBuy")
                        'Dim QtyAcumColumn As DataGridViewColumn = GridMyTable.Columns("QtyAcum")
                        'Dim QtyUserColumn As DataGridViewColumn = GridMyTable.Columns("QtyUser")
                        'Dim UMColumn As DataGridViewColumn = GridMyTable.Columns("UM")
                        'Dim QtyColumn As DataGridViewColumn = GridMyTable.Columns("Qty")
                        ''Dim UMReqColumn As DataGridViewColumn = GridMyTable.Columns("UM Req")
                        'Dim UnitPriceColumn As DataGridViewColumn = GridMyTable.Columns("UnitPrice")
                        'Dim PackPriceColumn As DataGridViewColumn = GridMyTable.Columns("PackPrice")
                        'Dim StandarPackColumn As DataGridViewColumn = GridMyTable.Columns("StandarPack")
                        'Dim MOQColumn As DataGridViewColumn = GridMyTable.Columns("MOQ")
                        'Dim LeadTimeColumn As DataGridViewColumn = GridMyTable.Columns("LeadTime")
                        'Dim VendorCodeColumn As DataGridViewColumn = GridMyTable.Columns("VendorCode")
                        'Dim DescriptionColumn As DataGridViewColumn = GridMyTable.Columns("Description")
                        'Dim FirstDayWeekColumn As DataGridViewColumn = GridMyTable.Columns("FirstDayWeek")
                        'Dim WeekColumn As DataGridViewColumn = GridMyTable.Columns("Week")
                        'Dim QtyInputSHPColumn As DataGridViewColumn = GridMyTable.Columns("QtyInputSHP")
                        'Dim KyColumn As DataGridViewColumn = GridMyTable.Columns("Ky")
                        'Dim IDColumn As DataGridViewColumn = GridMyTable.Columns("ID")
                        'PNColumn.Width = 90
                        'SubPNColumn.Width = 100
                        ''PNColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                        ''SubPNColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                        'QtyColumn.Width = 50
                        'QtyOnHandColumn.Width = 70
                        'QtyOnOrderColumn.Width = 70
                        'QtyOnOrderPerWeekColumn.Width = 70
                        ''QtyToBuyColumn.Width = 50
                        'QtyUserColumn.Width = 50
                        'QtyAcumColumn.Width = 50
                        'UMColumn.Width = 55
                        ''UMReqColumn.Width = 25
                        'UnitPriceColumn.Width = 50
                        'PackPriceColumn.Width = 50
                        'StandarPackColumn.Width = 50
                        'MOQColumn.Width = 40
                        'LeadTimeColumn.Width = 30
                        'VendorCodeColumn.Width = 70
                        'DescriptionColumn.Width = 70
                        'FirstDayWeekColumn.Width = 70
                        'WeekColumn.Width = 35
                        'QtyInputSHPColumn.Width = 70
                        'KyColumn.Width = 30
                        'IDColumn.Width = 30
                    End With

                End If
            End If
            GridMyTable.AutoResizeColumns()
        End Using
    End Sub
    'FindSalesOrdersByAU
    Private Sub FindSalesOrderByAU(ByVal AU As String, ByVal Rev As String) 'FindSalesOrderByAU
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try '
                Dim CompRev As String = ""
                If Rev <> "ALL" Then
                    CompRev = " AND Rev=@Rev "
                End If
                Dim CompStatus As String = ""
                'If rdoAll.Checked = True Then CompStatus = " AND Status=@"
                If rdoOpenSalesOrderByAU.Checked = True Then CompStatus = " AND Status='Open' "
                If rdoCloseSalesOrderByAU.Checked = True Then CompStatus = " AND Status='Close' "
                If rdoCancelSalesOrderByAU.Checked = True Then CompStatus = " AND Status='Cancel' "
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT AU, Rev, SONumber, PN, Qty, PackingSlipBalance, Balance, DueDate, UM, PO, PODate, Location, CustomerCode, Customer, Description, ShipAddress1, ShipAddress2, ShipAddress3, ShipCity, ShipState, ShipCountry, ShipZip, Status, CreatedBy, CreatedDate, ItemRow, UnitPrice, Amount FROM tblCustomerServiceSalesOrders WHERE AU=@AU " + CompRev + CompStatus + "  ORDER BY DueDate DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMWIP")
            End Try
            GridAUSalesOrderFind.DataSource = TN
            If GridAUSalesOrderFind.RowCount > 0 Then
                With GridAUSalesOrderFind
                    .Columns("DueDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("PODate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("CreatedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("UnitPrice").DefaultCellStyle.Format = ("$###,###.##")
                    .Columns("Amount").DefaultCellStyle.Format = ("$###,###.##")
                    '.Columns("").DefaultCellStyle.Format = ("dd/MMM/yy")
                    '.Columns("").DefaultCellStyle.Format = ("dd/MMM/yy")
                    '.Columns("").DefaultCellStyle.Format = ("dd/MMM/yy")
                End With
            End If
            GridAUSalesOrderFind.AutoResizeColumns()
            lblRecordsGridSalesOrder.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub FindWipByAU(ByVal AU As String, ByVal Rev As String, ByVal ForecastReference As String) 'FindSalesOrderByAU
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try '
                Dim CompRev As String = ""
                If Rev <> "ALL" Then
                    CompRev = " AND Rev=@Rev "
                End If
                Dim CompStatus As String = ""
                'If rdoAllWipByAU .Checked = True Then CompStatus = " AND Status=@"
                If rdoOpenWipByAU.Checked = True Then CompStatus = " AND Status='Open' "
                'If rdoCloseWipByAU.Checked = True Then CompStatus = " AND Status='Close' "
                'If rdoCancelWipByAU.Checked = True Then CompStatus = " AND Status='Cancel' "
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT WIP, AU, Rev, PN, Qty, DueDateCustomer, DueDateProcess, DueDateAssy, DueDateShipped, CreatedDate, Customer, IT, Notes, KindOfAU, Status, Line  FROM tblPurchasingWipFake WHERE ForecastReference=@ForecastReference AND AU=@AU " + CompRev + CompStatus + "  ORDER BY DueDateCustomer, Rev DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindWipByAU")
            End Try
            GridWIPForecast.DataSource = TN
            If GridWIPForecast.RowCount > 0 Then
                With GridWIPForecast
                    .Columns("DueDateCustomer").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateProcess").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateAssy").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateShipped").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("CreatedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                End With
            End If
            GridWIPForecast.AutoResizeColumns()
            lblRecordsGridWIPForecast.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub BuscaPNsPrimarios() 'ALX
        Dim Resp As String = "NO" 'tblItemsPOUserIDAuthorizations
        Dim Edo As String = ""
        Dim Query As String = "SELECT DISTINCT PN FROM tblItemsQB WHERE Active = '1'"
        Using TN As New System.Data.DataTable
            Try
                Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
                Dim DR As SqlDataReader
                'cmd.CommandType = CommandType.Text
                'cmd.Parameters.Add("@UserID", SqlDbType.NVarChar).Value = Usuario
                cnn.Open()
                DR = cmd.ExecuteReader
                TN.Load(DR)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString, "Error in BuscaPNsPrimarios function") 'despliega un mesaje si hay un error
            End Try
            If TN.Rows.Count > 0 Then
                Dim PN As String
                For NM As Integer = 0 To TN.Rows.Count - 1
                    PN = TN.Rows(NM).Item("PN").ToString
                    Using TW As New System.Data.DataTable
                        Try
                            Query = "SELECT SubPN FROM tblItemsQB WHERE PN=@PN AND PriOption='1' AND Active = '1'"
                            Dim cmd2 As SqlCommand = New SqlCommand(Query, cnn)
                            Dim DR2 As SqlDataReader
                            cmd2.CommandType = CommandType.Text
                            cmd2.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                            cnn.Open()
                            DR2 = cmd2.ExecuteReader
                            TW.Load(DR2)
                            cnn.Close()
                            If TW.Rows.Count = 0 Then
                                PN = PN
                                Using TZ As New System.Data.DataTable
                                    Try
                                        Query = "SELECT TOP(1) * FROM tblItemsQB WHERE PN=@PN AND Active = '1'"
                                        Dim cmd3 As SqlCommand = New SqlCommand(Query, cnn)
                                        Dim DR3 As SqlDataReader
                                        cmd3.CommandType = CommandType.Text
                                        cmd3.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                                        cnn.Open()
                                        DR3 = cmd3.ExecuteReader
                                        TZ.Load(DR3)
                                        cnn.Close()
                                        If TZ.Rows.Count > 0 Then
                                            Dim SubPN As String = TZ.Rows(0).Item("SubPN").ToString
                                            Dim IDItem As Long = CLng(Val(TZ.Rows(0).Item("IDItem").ToString))
                                            UpdateItems("PriOption", "1", "Entero", SubPN)
                                        End If
                                    Catch ex As Exception
                                        Edo = cnn.State.ToString
                                        If Edo = "Open" Then cnn.Close() 'cierra la conexion
                                        MessageBox.Show(ex.ToString, "Error in BuscaPNsPrimarios function") 'despliega un mesaje si hay un error
                                    End Try
                                End Using
                            End If
                        Catch ex As Exception
                            Edo = cnn.State.ToString
                            If Edo = "Open" Then cnn.Close() 'cierra la conexion
                            MessageBox.Show(ex.ToString, "Error in BuscaPNsPrimarios function") 'despliega un mesaje si hay un error
                        End Try
                    End Using
                Next
            End If
        End Using
    End Sub
    'funcion para actualizar los cambios cadena tblItemsQB
    Private Sub UpdateItems(ByVal Campo As String, ByVal Dato As String, ByVal Tipo As String, ByVal SubPN As String)
        Dim Edo As String
        Try 'Definimos el query del update
            Dim Query As String = "UPDATE tblItemsQB SET " + Campo + "=@Dato, ModifyBy=@ModifyBy, ModifyDate=@ModifyDate  WHERE SubPN=@SubPN"
            Dim cmd As SqlCommand = New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@SubPN", SqlDbType.NVarChar).Value = SubPN
            If Tipo = "Cadena" Then cmd.Parameters.Add("@Dato", SqlDbType.NVarChar).Value = Dato
            If Tipo = "Decimal" Then cmd.Parameters.Add("@Dato", SqlDbType.Float).Value = System.Convert.ToDouble(Val(Dato))
            If Tipo = "Entero" Then cmd.Parameters.Add("@Dato", SqlDbType.BigInt).Value = System.Convert.ToInt64(Val(Dato))
            If Tipo = "Booleano" Then cmd.Parameters.Add("@Dato", SqlDbType.Bit).Value = System.Convert.ToBoolean(Dato)
            cmd.Parameters.Add("@ModifyBy", SqlDbType.NVarChar).Value = txbUser.Text
            cmd.Parameters.Add("@ModifyDate", SqlDbType.DateTime).Value = Now
            cnn.Open() 'abre la conexion
            cmd.ExecuteNonQuery() 'realiza el query
            cnn.Close() 'cierra la conexion
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close() 'cierra la conexion
            MessageBox.Show(ex.ToString + vbNewLine + "Error traying to update the SubPN-" + SubPN.ToUpper + " Field " + Campo + " Data " + Dato + " in SEA", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'despliega un mesaje si hay un error
        End Try
    End Sub
    '
    Private Sub EnviaCorreo()
        Try
            Dim DestinatariosTO As String = CargaDestinatarios("MRP Forecast", "TO") ' "julio.gallegos@specializedharness.com" '"bgarcia@bitech.net, mespi@specializedharness.com, julio.gallegos@specializedharness.com"
            'Dim DestinatariosBCC As String = CargaDestinatarios("Sales Order", "BCC")
            Dim EnviadoPor As String = "shp.app@specializedharness.com"
            Dim CorreoIssues As String = ""
            Dim CorreoWarnings As String = ""
            Dim RutaWarning As String = ""
            Dim RutaIssues As String = ""
            Dim Correo As String = ""
            If tblIssues.Rows.Count > 0 Then
                CorreoIssues = "Se encontraron los siguientes Issues al momento de correr el forecast."
                CorreoIssues += vbNewLine
                RutaIssues = CreaCSVIssues()
            End If
            If tblWarning.Rows.Count > 0 Then
                CorreoWarnings += "Alertas"
                CorreoWarnings += vbNewLine
                RutaWarning = CreaCSVWarning()
            End If
            Correo += vbNewLine + vbNewLine
            Correo += "Por favor no responder este correo" + vbNewLine + "Gracias"
            'se envia email ade advertencia
            Dim _Message As New System.Net.Mail.MailMessage()
            Dim _SMTP As New System.Net.Mail.SmtpClient
            'Dim att As New System.Net.Mail.Attachment("\\bimexserver\Desarrollo de Software\Reporte de AUs\AU's Subidos a SEA.xlsx") ', System.Net.Mime.TransferEncoding.Base64

            'CONFIGURACIÓN DEL STMP
            _SMTP.Credentials = New System.Net.NetworkCredential(EnviadoPor, "Row.6078$")
            _SMTP.Host = "smtp.ipower.com"
            _SMTP.Port = 587
            _SMTP.EnableSsl = True
            ' CONFIGURACION DEL MENSAJE
            ' _Message.Bcc.Add(DestinatariosBCC)
            _Message.[To].Add(DestinatariosTO)
            _Message.From = New System.Net.Mail.MailAddress(EnviadoPor, "", System.Text.Encoding.UTF8) 'Quien lo envía
            _Message.Subject = "Reporte de MRP by Forecast"
            _Message.SubjectEncoding = System.Text.Encoding.UTF8 'Codificacion
            _Message.Body = CorreoIssues + vbNewLine + CorreoWarnings
            _Message.BodyEncoding = System.Text.Encoding.UTF8
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            If RutaIssues <> "" Then
                Dim att1 As New System.Net.Mail.Attachment(RutaIssues)
                _Message.Attachments.Add(att1)
            End If
            If RutaWarning <> "" Then
                Dim att2 As New System.Net.Mail.Attachment(RutaWarning)
                _Message.Attachments.Add(att2)
            End If
            _Message.IsBodyHtml = False
            'ENVIO
            _SMTP.Send(_Message)
            'MsgBox("Se ha Enviado el Email", MsgBoxStyle.Information, "EMail Enviado")
        Catch ex As Exception
            MsgBox(ex.ToString.ToString)
        End Try
    End Sub
    '
    Private Function CargaDestinatarios(ByVal Modulo As String, ByVal OpcionEnvio As String)
        Dim Destinatarios As String = ""
        Dim Edo As String = ""
        Dim contador As Long = 0
        Using TE As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT Email FROM tblUserEmails WHERE Module=@Module AND Active=1 AND OptionToSend=@OptionToSend"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@Module", SqlDbType.NVarChar).Value = Modulo
                cmd.Parameters.Add("@OptionToSend", SqlDbType.NVarChar).Value = OpcionEnvio
                cnn.Open()
                dr = cmd.ExecuteReader
                TE.Load(dr)
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error Loading tblMaster")
            End Try
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            If TE.Rows.Count > 0 Then
                For NM As Integer = 0 To TE.Rows.Count - 1
                    Destinatarios += TE.Rows(NM).Item("Email").ToString
                    If NM < TE.Rows.Count - 1 Then Destinatarios += ","
                Next
            End If
        End Using
        Return Destinatarios
    End Function

    Private Function CreaCSVIssues()
        Dim Banderilla As Integer = 0
        Dim ArchivoNombre As String = "Issues " + Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + Now.ToString("HH") + Now.ToString("mm") + ".csv"
        Dim Ruta As String = Path.GetTempPath() & ArchivoNombre
        Try
            Dim fs As FileStream = File.Create(Ruta)
            Dim Cadena As String = "Num,Issue," + vbNewLine
            Dim infoTitulos As Byte() = New UTF8Encoding(True).GetBytes(Cadena)
            fs.Write(infoTitulos, 0, infoTitulos.Length)
            For NM As Integer = 0 To tblIssues.Rows.Count - 1
                Cadena = tblIssues.Rows(NM).Item("ItemRow").ToString + "," + tblIssues.Rows(NM).Item("Issues").ToString + "," + vbNewLine
                'Cadena += tblIssues.Rows(NM).Item("Issues").ToString
                'Cadena += vbNewLine
                Dim info As Byte() = New UTF8Encoding(True).GetBytes(Cadena)
                fs.Write(info, 0, info.Length)
            Next
            fs.Close()
            Banderilla += 1
        Catch ex As Exception
            MessageBox.Show(ex.ToString.ToString + vbNewLine + "Error in Issues.csv", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        If Banderilla = 0 Then
            Ruta = ""
        End If
        Return Ruta
    End Function
    Private Function CreaCSVWarning()
        Dim Banderilla As Integer = 0
        Dim ArchivoNombre As String = "Warning " + Now.ToString("yyyy") + Now.ToString("MM") + Now.ToString("dd") + Now.ToString("HH") + Now.ToString("mm") + ".csv"
        Dim Ruta As String = Path.GetTempPath() & ArchivoNombre
        Try
            Dim fs As FileStream = File.Create(Ruta)
            Dim Cadena As String = "Num,Warning," + vbNewLine
            Dim infoTitulos As Byte() = New UTF8Encoding(True).GetBytes(Cadena)
            fs.Write(infoTitulos, 0, infoTitulos.Length)
            For NM As Integer = 0 To tblWarning.Rows.Count - 1
                Cadena = tblWarning.Rows(NM).Item("ItemRow").ToString + "," + tblWarning.Rows(NM).Item("Warnings").ToString + "," + vbNewLine
                'Cadena += tblIssues.Rows(NM).Item("Issues").ToString
                'Cadena += vbNewLine
                Dim info As Byte() = New UTF8Encoding(True).GetBytes(Cadena)
                fs.Write(info, 0, info.Length)
            Next
            fs.Close()
            Banderilla += 1
        Catch ex As Exception
            MessageBox.Show(ex.ToString.ToString + vbNewLine + "Error in Issues.csv", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        If Banderilla = 0 Then
            Ruta = ""
        End If
        Return Ruta
    End Function

    Private Function BuscaPNConReservado(ByVal PN As String, ByVal SubPN As String)
        Dim Resp As Integer = 0
        Using TN As New Data.DataTable
            Dim Edo As String = ""
            Try 'tblItemsFinantialInventoryControlTempforProductionProcess" & sTempTableName 
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT SubPN, Reserved FROM tblPurchasingTempMRPForecast" & sTempTableName & " WHERE SubPN=@SubPN GROUP BY SubPN, Reserved "
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cmd.Parameters.Add("@SubPN", SqlDbType.NVarChar).Value = SubPN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
                Resp = TN.Rows.Count
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading BuscaPNConReservado function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        Return Resp
    End Function
    '=========================================================== Controles BOM WIP ====================================================================
    Private Sub txbBOMWIP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbBOMWIP.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Asc(e.KeyChar) = 13 Then
            FindPNBOMWIP(txbBOMWIP.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbBOMWIP_TextChanged(sender As Object, e As EventArgs) Handles txbBOMWIP.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnFindBOMWIP_Click(sender As Object, e As EventArgs) Handles btnFindBOMWIP.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        FindPNBOMWIP(txbBOMWIP.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbAUBOMWIP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAUBOMWIP.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbAUBOMWIP.SelectedIndex > -1 Then
            If cmbAUBOMWIP.SelectedValue.ToString <> "System.Data.DataRowView" Then
                cmbRevBOMWIP.Text = "ALL"
                cmbWIPBOMWIP.Text = "ALL"
                CargaComboRevWIP(cmbAUBOMWIP.SelectedValue.ToString)
                CargaComboWIP(cmbAUBOMWIP.SelectedValue.ToString, cmbRevBOMWIP.Text)
                FindBOMWIP(cmbAUBOMWIP.SelectedValue.ToString, cmbRevBOMWIP.Text, cmbWIPBOMWIP.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbRevBOMWIP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRevBOMWIP.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbRevBOMWIP.SelectedIndex > -1 Then
            If cmbRevBOMWIP.SelectedValue.ToString <> "System.Data.DataRowView" Then
                CargaComboWIP(cmbAUBOMWIP.SelectedValue.ToString, cmbRevBOMWIP.SelectedValue.ToString)
                FindBOMWIP(cmbAUBOMWIP.SelectedValue.ToString, cmbRevBOMWIP.Text, cmbWIPBOMWIP.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbWIPBOMWIP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbWIPBOMWIP.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbWIPBOMWIP.SelectedIndex > -1 Then
            FindBOMWIP(cmbAUBOMWIP.SelectedValue.ToString, cmbRevBOMWIP.Text, cmbWIPBOMWIP.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridBOMWIP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMWIP.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridBOMWIP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMWIP.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridBOMWIP_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMWIP.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub FindPNBOMWIP(ByVal PN As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT tblBOMWIP.PN, tblBOMWIP.Qty, tblBOMWIP.Balance, tblBOMWIP.Unit, tblBOMWIP.AU, tblBOMWIP.Rev, tblBOMWIP.WIP, tblBOMWIP.RequieredDate, tblBOMWIP.Description, tblBOMWIP.MaterialGroup, tblBOMWIP.Reference, tblBOMWIP.PickList, tblBOMWIP.Route, tblBOMWIP.CreatedBy, tblBOMWIP.CreatedDate, tblBOMWIP.ModifyBy, tblBOMWIP.ModifyDate FROM tblBOMWIP INNER JOIN tblWIP ON tblBOMWIP.WIP = tblWIP.WIP WHERE ((tblWIP.Status = N'Open') AND (tblBOMWIP.PN = @PN) AND (tblBOMWIP.Balance>0)) ORDER BY tblBOMWIP.AU ASC, tblBOMWIP.Rev DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@PN", SqlDbType.NVarChar).Value = PN
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindPNBOMWIP")
            End Try
            GridBOMWIP.DataSource = TN
            GridBOMWIP.AutoResizeColumns()
            lblRecordsBOMWIP.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub

    Private Sub CargaComboWIP(ByVal AU As Long, ByVal Rev As String)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblWIPBOMWIP.Clear()
        Dim R As DataRow = tblWIPBOMWIP.NewRow
        R.Item("WIP") = "ALL"
        tblWIPBOMWIP.Rows.Add(R)
        Using TN As New Data.DataTable
            Dim Complemento As String = ""
            If Rev <> "ALL" And Rev <> "" Then
                Complemento = " AND Rev=@Rev "
            End If
            Dim Query As String = "SELECT WIP FROM tblWIP WHERE Status='Open' AND AU=@AU " + Complemento + " ORDER BY WIP ASC"
            Try
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                tblWIPBOMWIP.Load(dr)
                cnn.Close()
                'If TN.Rows.Count > 0 Then
                '    For NM As Integer = 0 To TN.Rows.Count - 1
                '        Edo = CStr(TN.Rows(NM).Item("WIP").ToString)
                '        tblWIPBOMWIP.Rows.Add(CStr(TN.Rows(NM).Item("WIP").ToString))
                '    Next
                'End If
                With cmbWIPBOMWIP
                    .DataSource = Nothing
                    .DataSource = tblWIPBOMWIP
                    .DisplayMember = "WIP"
                    .ValueMember = "WIP"
                End With
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close() 'cierra la conexion
                MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboRevWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    '=========================================================== Controles BOM ENG ====================================================================
    Private Sub txbPNBOMENG_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbPNBOMENG.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Asc(e.KeyChar) = 13 Then
            FindPNBOMENG(txbPNBOMENG.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbPNBOMENG_TextChanged(sender As Object, e As EventArgs) Handles txbPNBOMENG.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnFindBOMENG_Click(sender As Object, e As EventArgs) Handles btnFindBOMENG.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        FindPNBOMENG(txbPNBOMENG.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbAUBOMENG_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAUBOMENG.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbAUBOMENG.SelectedIndex > -1 Then
            If cmbAUBOMENG.SelectedValue.ToString <> "System.Data.DataRowView" Then
                'cmbRevBOMENG.Text = "ALL"
                CargaComboRevENG(cmbAUBOMENG.SelectedValue.ToString)
                FindBOMENG(cmbAUBOMENG.SelectedValue.ToString, cmbRevBOMENG.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbRevBOMENG_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRevBOMENG.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbRevBOMENG.SelectedIndex > -1 Then
            If cmbRevBOMENG.SelectedValue.ToString <> "System.Data.DataRowView" Then
                FindBOMENG(cmbAUBOMENG.SelectedValue.ToString, cmbRevBOMENG.SelectedValue.ToString)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridBOMENG_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMENG.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridBOMENG_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMENG.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridBOMENG_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMENG.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    '=========================================================== Controles Search in My Table ====================================================================
    Private Sub cmbPNMyTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPNMyTable.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbPNMyTable.SelectedIndex > -1 Then
            If cmbPNMyTable.SelectedValue.ToString <> "System.Data.DataRowView" Then
                FindPNMyTable(cmbPNMyTable.SelectedValue.ToString)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridMyTable_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridMyTable.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridMyTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridMyTable.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridMyTable_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridMyTable.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    '=========================================================== Controles Sales Order ====================================================================
    Private Sub txbAUSalesOrder_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbAUSalesOrder.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Asc(e.KeyChar) = 13 Then
            If txbAUSalesOrder.Text <> "" Then
                If IsNumeric(txbAUSalesOrder.Text) = True Then CargaComboRevSalesOrder(txbAUSalesOrder.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbAUSalesOrder_LostFocus(sender As Object, e As EventArgs) Handles txbAUSalesOrder.LostFocus
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If txbAUSalesOrder.Text <> "" Then
            If IsNumeric(txbAUSalesOrder.Text) = True Then
                CargaComboRevSalesOrder(txbAUSalesOrder.Text)
                cmbRevSalesOrder.Text = "ALL"
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbAUSalesOrder_TextChanged(sender As Object, e As EventArgs) Handles txbAUSalesOrder.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbRevSalesOrder_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRevSalesOrder.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbRevSalesOrder.SelectedIndex > -1 Then
            If cmbRevSalesOrder.SelectedValue.ToString <> "System.Data.DataRowView" Then
                If IsNumeric(txbAUSalesOrder.Text) = True Then FindSalesOrderByAU(txbAUSalesOrder.Text, cmbRevSalesOrder.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnFindSalesOrder_Click(sender As Object, e As EventArgs) Handles btnFindSalesOrder.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUSalesOrder.Text) = True Then FindSalesOrderByAU(txbAUSalesOrder.Text, cmbRevSalesOrder.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoOpenSalesOrderByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoOpenSalesOrderByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoOpenSalesOrderByAU_Click(sender As Object, e As EventArgs) Handles rdoOpenSalesOrderByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUSalesOrder.Text) = True Then
            FindSalesOrderByAU(txbAUSalesOrder.Text, cmbRevSalesOrder.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoCloseSalesOrderByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCloseSalesOrderByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoCloseSalesOrderByAU_Click(sender As Object, e As EventArgs) Handles rdoCloseSalesOrderByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUSalesOrder.Text) = True Then
            FindSalesOrderByAU(txbAUSalesOrder.Text, cmbRevSalesOrder.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoCancelSalesOrderByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCancelSalesOrderByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoCancelSalesOrderByAU_Click(sender As Object, e As EventArgs) Handles rdoCancelSalesOrderByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUSalesOrder.Text) = True Then
            FindSalesOrderByAU(txbAUSalesOrder.Text, cmbRevSalesOrder.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoAllSalesOrderByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAllSalesOrderByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoAllSalesOrderByAU_Click(sender As Object, e As EventArgs) Handles rdoAllSalesOrderByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUSalesOrder.Text) = True Then
            FindSalesOrderByAU(txbAUSalesOrder.Text, cmbRevSalesOrder.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridAUSalesOrderFind_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridAUSalesOrderFind.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridAUSalesOrderFind_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridAUSalesOrderFind.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridAUSalesOrderFind_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridAUSalesOrderFind.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    '=========================================================== Controles WIP produccion ====================================================================
    Private Sub txbAUWipByAU_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbAUWipByAU.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Asc(e.KeyChar) = 13 Then
            If txbAUWipByAU.Text <> "" Then
                If IsNumeric(txbAUWipByAU.Text) = True Then
                    CargaComboRevWipByAU(txbAUWipByAU.Text)
                    cmbRevWipByAU.Text = "ALL"
                End If
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbAUWipByAU_LostFocus(sender As Object, e As EventArgs) Handles txbAUWipByAU.LostFocus
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If txbAUWipByAU.Text <> "" Then
            If IsNumeric(txbAUWipByAU.Text) = True Then
                CargaComboRevWipByAU(txbAUWipByAU.Text)
                cmbRevWipByAU.Text = "ALL"
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbAUWipByAU_TextChanged(sender As Object, e As EventArgs) Handles txbAUWipByAU.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbRevWipByAU_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRevWipByAU.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbRevWipByAU.SelectedIndex > -1 Then
            If cmbRevWipByAU.SelectedValue.ToString <> "System.Data.DataRowView" Then
                If IsNumeric(txbAUWipByAU.Text) = True Then FindWipByAU(txbAUWipByAU.Text, cmbRevWipByAU.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoOpenWipByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoOpenWipByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoOpenWipByAU_Click(sender As Object, e As EventArgs) Handles rdoOpenWipByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUWipByAU.Text) = True Then FindWipByAU(txbAUWipByAU.Text, cmbRevWipByAU.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoCloseWipByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCloseWipByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoCloseWipByAU_Click(sender As Object, e As EventArgs) Handles rdoCloseWipByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUWipByAU.Text) = True Then FindWipByAU(txbAUWipByAU.Text, cmbRevWipByAU.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoCancelWipByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCancelWipByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoCancelWipByAU_Click(sender As Object, e As EventArgs) Handles rdoCancelWipByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUWipByAU.Text) = True Then FindWipByAU(txbAUWipByAU.Text, cmbRevWipByAU.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub rdoAllWipByAU_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAllWipByAU.CheckedChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub rdoAllWipByAU_Click(sender As Object, e As EventArgs) Handles rdoAllWipByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUWipByAU.Text) = True Then FindWipByAU(txbAUWipByAU.Text, cmbRevWipByAU.Text)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnFindWipByAU_Click(sender As Object, e As EventArgs) Handles btnFindWipByAU.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If IsNumeric(txbAUWipByAU.Text) = True Then
            FindWipByAU(txbAUWipByAU.Text, cmbRevWipByAU.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridWipByAU_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridWipByAU.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridWipByAU_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridWipByAU.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridWipByAU_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridWipByAU.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub FindWipByAU(ByVal AU As String, ByVal Rev As String) 'FindSalesOrderByAU
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try '
                Dim CompRev As String = ""
                If Rev <> "ALL" Then
                    CompRev = " AND Rev=@Rev "
                End If
                Dim CompStatus As String = ""
                'If rdoAllWipByAU .Checked = True Then CompStatus = " AND Status=@"
                If rdoOpenWipByAU.Checked = True Then CompStatus = " AND Status='Open' "
                If rdoCloseWipByAU.Checked = True Then CompStatus = " AND Status='Close' "
                If rdoCancelWipByAU.Checked = True Then CompStatus = " AND Status='Cancel' "
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT WIP, AU, Rev, PN, Qty, DueDateCustomer, DueDateProcess, DueDateAssy, DueDateShipped, CreatedDate, ClosedDate, BalanceProcess, BalanceSubStorage, BalanceAssy, BalancePack, BalanceShipped, Priority, Customer, IT, Notes, KindOfAU, Status, Line  FROM tblWIP WHERE AU=@AU " + CompRev + CompStatus + "  ORDER BY DueDateCustomer DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindWipByAU")
            End Try
            GridWipByAU.DataSource = TN
            If GridWipByAU.RowCount > 0 Then
                With GridWipByAU
                    .Columns("DueDateCustomer").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateProcess").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateAssy").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateShipped").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("CreatedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("ClosedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                End With
            End If
            GridWipByAU.AutoResizeColumns()
            lblRecordsWipByAU.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub
    '=========================================================== Controles Wip Forecast ====================================================================
    Private Sub cmbAUWIPForecast_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAUWIPForecast.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbAUWIPForecast.SelectedIndex > -1 Then
            If cmbAUWIPForecast.SelectedValue.ToString <> "System.Data.DataRowView" Then
                CargaComboRevWipByAUForecastreference(cmbAUWIPForecast.SelectedValue.ToString, lblForecastReference.Text)
                FindAUWIPForecast(cmbAUWIPForecast.SelectedValue.ToString, cmbRevWIPForecast.Text, lblForecastReference.Text)
                'FindWipByAU(txbAUWIPForecast.Text, cmbRevWIPForecast.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbRevWIPForecast_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRevWIPForecast.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbRevWIPForecast.SelectedIndex > -1 Then
            If cmbRevWIPForecast.SelectedValue.ToString <> "System.Data.DataRowView" Then
                If IsNumeric(cmbAUWIPForecast.Text) = True Then
                    FindAUWIPForecast(cmbAUWIPForecast.SelectedValue.ToString, cmbRevWIPForecast.Text, lblForecastReference.Text)
                    'FindWipByAU(cmbAUWIPForecast.Text, cmbRevWIPForecast.Text)
                End If
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnWIPForecast_Click(sender As Object, e As EventArgs) Handles btnWIPForecast.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridWIPForecast_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridWIPForecast.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridWIPForecast_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridWIPForecast.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridWIPForecast_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridWIPForecast.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboAUBOMWIPForecast()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblAUBOMWIPForecastreference.Clear()
        'Dim tblPNs As New DataTable
        Dim Query As String = "SELECT DISTINCT AU FROM tblPurchasingWipFake WHERE Forecastreference=@Forecastreference ORDER BY AU ASC"
        Try
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cmd = New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@Forecastreference", SqlDbType.NVarChar).Value = lblForecastReference.Text
            'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
            cnn.Open()
            dr = cmd.ExecuteReader
            tblAUBOMWIPForecastreference.Load(dr)
            cnn.Close()
            If tblAUBOMWIPForecastreference.Rows.Count > 0 Then
                With cmbAUBOMForecast
                    .DataSource = tblAUBOMWIPForecastreference
                    .DisplayMember = "AU"
                    .ValueMember = "AU"
                    ' .Text = tblItems.Rows(0).Item("ShipTo").ToString
                End With
            End If
            Dim Contador As Long = tblAUBOMWIPForecastreference.Rows.Count
        Catch ex As Exception
            MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboAUWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Edo = cnn.State.ToString
        If Edo = "Open" Then cnn.Close() 'cierra la conexion
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub CargaComboAUWIPForecast()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim Edo As String = ""
        tblAUWIPForecastreference.Clear()
        'Dim tblPNs As New DataTable
        Dim Query As String = "SELECT DISTINCT AU FROM tblPurchasingWipFake WHERE Forecastreference=@Forecastreference ORDER BY AU ASC"
        Try
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cmd = New SqlCommand(Query, cnn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@Forecastreference", SqlDbType.NVarChar).Value = lblForecastReference.Text
            'cmd.Parameters.Add("@Field", SqlDbType.NVarChar).Value = Field
            cnn.Open()
            dr = cmd.ExecuteReader
            tblAUWIPForecastreference.Load(dr)
            cnn.Close()
            If tblAUWIPForecastreference.Rows.Count > 0 Then
                With cmbAUWIPForecast
                    .DataSource = tblAUWIPForecastreference
                    .DisplayMember = "AU"
                    .ValueMember = "AU"
                    ' .Text = tblItems.Rows(0).Item("ShipTo").ToString
                End With
            End If
            Dim Contador As Long = tblAUWIPForecastreference.Rows.Count
        Catch ex As Exception
            Edo = cnn.State.ToString
            If Edo = "Open" Then cnn.Close()
            MessageBox.Show(ex.ToString + vbNewLine + "Error Loading CargaComboAUWIP function", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub FindAUWIPForecast(ByVal AU As String, ByVal Rev As String, ByVal ForecastReference As String)
        Dim Edo As String = ""
        Using TN As New Data.DataTable
            Try '
                Dim CompRev As String = ""
                If Rev <> "ALL" Then
                    CompRev = " AND Rev=@Rev "
                End If
                Dim CompStatus As String = ""
                'If rdoAllWipByAU .Checked = True Then CompStatus = " AND Status=@"
                'If rdoOpenWipByAU.Checked = True Then CompStatus = " AND Status='Open' "
                'If rdoCloseWipByAU.Checked = True Then CompStatus = " AND Status='Close' "
                'If rdoCancelWipByAU.Checked = True Then CompStatus = " AND Status='Cancel' "
                Dim cmd As SqlCommand
                Dim dr As SqlDataReader
                Dim Query As String = "SELECT WIP, AU, Rev, PN, Qty, DueDateCustomer, DueDateProcess, DueDateAssy, DueDateShipped, CreatedDate, Customer, IT, Notes, KindOfAU, Line  FROM tblPurchasingWipFake WHERE ForecastReference=@ForecastReference AND AU=@AU " + CompRev + CompStatus + "  ORDER BY DueDateCustomer, Rev DESC"
                cmd = New SqlCommand(Query, cnn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@AU", SqlDbType.NVarChar).Value = AU
                cmd.Parameters.Add("@Rev", SqlDbType.NVarChar).Value = Rev
                cmd.Parameters.Add("@ForecastReference", SqlDbType.NVarChar).Value = ForecastReference
                cnn.Open()
                dr = cmd.ExecuteReader
                TN.Load(dr)
                cnn.Close()
            Catch ex As Exception
                Edo = cnn.State.ToString
                If Edo = "Open" Then cnn.Close()
                MessageBox.Show(ex.ToString, "Error Loading FindWipByAU")
            End Try
            GridWIPForecast.DataSource = TN
            If GridWIPForecast.RowCount > 0 Then
                With GridWIPForecast
                    .Columns("DueDateCustomer").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateProcess").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateAssy").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("DueDateShipped").DefaultCellStyle.Format = ("dd/MMM/yy")
                    .Columns("CreatedDate").DefaultCellStyle.Format = ("dd/MMM/yy")
                End With
            End If
            GridWIPForecast.AutoResizeColumns()
            lblRecordsGridWIPForecast.Text = "Records " + TN.Rows.Count.ToString
        End Using
    End Sub
    '=========================================================== Controles BOM Forecast ====================================================================
    Private Sub txbPNBOMForecast_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txbPNBOMForecast.KeyPress
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub txbPNBOMForecast_TextChanged(sender As Object, e As EventArgs) Handles txbPNBOMForecast.TextChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub btnFindBOMForecast_Click(sender As Object, e As EventArgs) Handles btnFindBOMForecast.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbAUBOMForecast_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAUBOMForecast.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbAUBOMForecast.SelectedIndex > -1 Then
            If cmbAUBOMForecast.SelectedValue.ToString <> "System.Data.DataRowView" Then
                cmbRevBOMForecast.Text = "ALL"
                cmbWIPBOMForecast.Text = "ALL"
                CargaComboRevWIPForecast(cmbAUBOMForecast.SelectedValue.ToString)
                CargaComboWIPForecast(cmbAUBOMForecast.SelectedValue.ToString, cmbRevBOMWIP.Text)
                FindBOMWIPForecast(cmbAUBOMForecast.SelectedValue.ToString, cmbRevBOMWIP.Text, cmbWIPBOMWIP.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbRevBOMForecast_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRevBOMForecast.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbRevBOMForecast.SelectedIndex > -1 Then
            If cmbRevBOMForecast.SelectedValue.ToString <> "System.Data.DataRowView" Then
                CargaComboWIPForecast(cmbAUBOMForecast.SelectedValue.ToString, cmbRevBOMForecast.SelectedValue.ToString)
                FindBOMWIP(cmbAUBOMForecast.SelectedValue.ToString, cmbRevBOMForecast.Text, cmbWIPBOMForecast.Text)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmbWIPBOMForecast_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbWIPBOMForecast.SelectedIndexChanged
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If cmbWIPBOMForecast.SelectedIndex > -1 Then
            FindBOMWIPForecast(cmbAUBOMForecast.SelectedValue.ToString, cmbRevBOMForecast.Text, cmbWIPBOMForecast.Text)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub GridBOMForecast_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMForecast.CellClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridBOMForecast_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMForecast.CellContentClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub GridBOMForecast_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBOMForecast.CellDoubleClick
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    '=========================================================== Controles  ====================================================================
End Class
