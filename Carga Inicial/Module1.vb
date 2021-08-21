
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography
Imports System.Windows.Forms

Module Module1
    Public Instancia, PCNombre, RFCEstacion, RFCProveedor, PermisoCRE, ArchPFX, Sello As String
    Public HoraCorte As Date
    Public Foliorecepcion, FolioRelacion As Integer
    Public appPath As String = "C:\Users\Public\Admingas"
    Public conn As New SqlClient.SqlConnection 'Crea un objeto Connection
    Public connWeb As New SqlClient.SqlConnection 'Crea un objeto Connection
    Public Tanque, ClaveP, ClaveSub As Integer
    Public StrSql, NombreDoc As String

    Sub conectar()
        Instancia = Trim(LeerIni(Application.StartupPath() & "\convol.ini"))
        Try
            conn.ConnectionString = ("Data Source=" & Instancia & ";Initial Catalog=Admingas;Persist Security Info=True;User ID=sa;Password=admingas")
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
            Exit Sub
        End Try
    End Sub


    Public Function LeerIni(ArchivoINI As String) As String
        Using streamReader As System.IO.StreamReader = System.IO.File.OpenText(ArchivoINI)
            LeerIni = streamReader.ReadLine()
        End Using
    End Function
    Public Sub ComSQL(xstrsql As String)
        Dim comando As SqlCommand
        comando = New SqlCommand(xstrsql, conn)
        comando.ExecuteNonQuery()
    End Sub
    Function CreaSello(ArchXml As Xml.XmlDocument)
        Dim xmlDoc As New System.Xml.XmlDocument
        Dim CadenaOriginal As String
        Dim transformer As System.Xml.Xsl.XslCompiledTransform
        Dim Archivo_XSLT As String
        'Dim document As New System.Xml.XmlDocument
        Dim navigator As System.Xml.XPath.XPathNavigator
        Dim output As New System.IO.StringWriter()
        'document = New System.Xml.XmlDocument()
        transformer = New System.Xml.Xsl.XslCompiledTransform
        Archivo_XSLT = "C:\Users\Public\admingas\ArchivosDigitales\cadenaoriginal_controlesvolumetricos_v1.2.xslt"
        'document.Load(ArchXml)
        navigator = ArchXml.CreateNavigator
        transformer.Load(Archivo_XSLT)
        transformer.Transform(navigator, Nothing, output)
        CadenaOriginal = output.ToString
        output.Close()
        CreaSello = GenerarSello(CadenaOriginal)
        'FrmInicio.TxtCadenaOriginal.Text = CadenaOriginal
        'FrmInicio.TxtCadenaOriginal.Refresh()
        'Dim has As New OC.Core.Crypto.Hash

        ':::Encriptamos el string recibido al metodo SHA256 y los mostramos en el TextBox en minúsculas
        'Dim CadenaEncriptada As String = has.Sha256(CadenaOriginal).ToLower
        'CreaSello = Trim(CadenaEncriptada)
    End Function

    Public Function GenerarSello(CadenaOriginal As String) As String
        ArchPFX = appPath + "\ArchivosDigitales\" + Trim(DREmpresa("RFC")) + ".pfx"
        Dim privateCert As New X509Certificate2(ArchPFX, "admingas", X509KeyStorageFlags.Exportable)
        Dim privateKey As RSACryptoServiceProvider = DirectCast(privateCert.PrivateKey, RSACryptoServiceProvider)
        Dim privateKey1 As New RSACryptoServiceProvider()
        privateKey1.ImportParameters(privateKey.ExportParameters(True))
        Dim stringCadenaOriginal() As Byte = System.Text.Encoding.UTF8.GetBytes(CadenaOriginal)
        Dim signature As Byte() = privateKey1.SignData(stringCadenaOriginal, "SHA256")
        Dim sello256 As String = Convert.ToBase64String(signature)
        'para verificar el sello
        Dim isValid As Boolean = privateKey1.VerifyData(stringCadenaOriginal, "SHA256", signature)
        If isValid Then
            GenerarSello = sello256
        Else
            GenerarSello = "SELLO INVÁLIDO"
        End If

    End Function

    Public Sub CargaFolios()

        DSFolios = New DataSet
        SDAFolios = New SqlClient.SqlDataAdapter("Select * from folios", conn)
        SDAFolios.Fill(DSFolios, "Folios")
        DRFolios = DSFolios.Tables("Folios").Rows(0)
    End Sub

    Public Sub ClaveProds(Clave As Integer)
        Select Case Clave
            Case "32011"
                ClaveP = "07"
                ClaveSub = 1
            Case "32012"
                ClaveP = "07"
                ClaveSub = 2
            Case "34006"
                ClaveP = "03"
                ClaveSub = 3
        End Select
    End Sub
    Public Sub CorregirDespachosNull(Fechanull As String)
        Dim NombreCampo As String
        StrSql = "Select * from Despachos where fechahorag='" + Fechanull + "' and tipoventa<>'D' and importeventa is null"
        SDADespachos = New SqlClient.SqlDataAdapter(StrSql, conn)
        DSDespachos = New DataSet
        SDADespachos.Fill(DSDespachos, "Despachos")
        If DSDespachos.Tables("Despachos").Rows.Count > 0 Then
            For i = 0 To DSDespachos.Tables("Despachos").Rows.Count - 1
                DRDespachos = DSDespachos.Tables("Despachos").Rows(i)
                For Each column In DSDespachos.Tables("despachos").Columns
                    NombreCampo = column.ToString
                    If LCase(NombreCampo) <> "transaccion" And LCase(NombreCampo) <> "Posicion" And LCase(NombreCampo) <> "fechahora" _
                        And LCase(NombreCampo) <> "fechahorag" And LCase(NombreCampo) <> "tipoventa" And LCase(NombreCampo) <> "transi" And LCase(NombreCampo) <> "tpago" Then
                        StrSql = "update despachos set " + NombreCampo + "= (select top(1) " + NombreCampo + " from despachos where posicion=" + _
                            Str(DRDespachos("posicion")) + " and transaccion<" + Str(DRDespachos("transaccion")) + " order by transaccion desc) where transaccion=" + Str(DRDespachos("transaccion"))
                        ComSQL(StrSql)
                        StrSql = "update despachos set solicitud=0,volumen=0,importeventa=0,numticket=0,factura=0,tpago='EFECTIVO' WHERE TRANSACCION=" + Str(DRDespachos("Transaccion"))
                        ComSQL(StrSql)
                    End If
                Next
            Next
        End If
    End Sub
End Module
