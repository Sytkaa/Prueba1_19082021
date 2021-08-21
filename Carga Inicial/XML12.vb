Imports System
Imports System.Xml
Imports System.IO
Imports System.IO.Compression
Public Class Convol12

    Shared Function CREAXML(HoraCorte As Date, Certificado As String, NoCertificado As String) As String
        Try
            '****VERIFICAR ESTACIÓN AUTORIZADA
            conectar()
            CargaEmpresa()
            RFCEstacion = Trim(DREmpresa("RFC"))
            RFCProveedor = "AES130222NU4"
            PermisoCRE = Trim(DREmpresa("PermisoCRE"))
            PermisoCRE = Replace(PermisoCRE, "_", "/")
            If File.Exists(appPath + "\esaut.ini") Then
                My.Computer.FileSystem.DeleteFile(appPath + "\esaut.ini")
            End If
            My.Computer.Network.DownloadFile("ftp://187.188.161.105/esaut.ini", appPath + "\esaut.ini")
            Dim Estaciones As String = LeerIni(appPath + "\esaut.ini")
            If InStr(Estaciones, Trim(DREmpresa("RFC"))) Then

                '*****COMIENZA LA FORMACIÓN DEL XML

                Dim Doc As New XmlDocument, Nodo, EXI, REC, RECCab, RECDet, RECDocs, VTA, VTACabecera, VTADetalle, TQS, DIS As XmlNode
                Nodo = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:ControlesVolumetricos", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                Nodo = Doc.AppendChild(Nodo)

                Dim schemaLocationCV As XmlAttribute = Doc.CreateAttribute("xmlns:controlesvolumetricos")
                Dim schemaLocation As XmlAttribute = Doc.CreateAttribute("xmlns:xs")
                Dim atributo_version As XmlAttribute = Doc.CreateAttribute("version")
                Dim atributo_CRE As XmlAttribute = Doc.CreateAttribute("numeroPermisoCRE")
                Dim atributo_rfc As XmlAttribute = Doc.CreateAttribute("rfc")
                Dim atributo_rfcProveedorSw As XmlAttribute = Doc.CreateAttribute("rfcProveedorSw")
                Dim atributo_sello As XmlAttribute = Doc.CreateAttribute("sello")
                Dim atributo_Nocertificado As XmlAttribute = Doc.CreateAttribute("noCertificado")
                Dim atributo_certificado As XmlAttribute = Doc.CreateAttribute("certificado")
                Dim atributo_fechacorte As XmlAttribute = Doc.CreateAttribute("fechaYHoraCorte")

                schemaLocationCV.Value = "http://www.sat.gob.mx/esquemas/controlesvolumetricos"
                schemaLocation.Value = "http://www.w3.org/2001/XMLSchema"
                atributo_version.Value = "1.2"
                atributo_rfc.Value = RFCEstacion
                atributo_rfcProveedorSw.Value = RFCProveedor
                atributo_CRE.Value = PermisoCRE
                atributo_Nocertificado.Value = NoCertificado
                atributo_certificado.Value = Certificado
                atributo_fechacorte.Value = Format(HoraCorte, "yyyy-MM-ddThh:mm:ss")

                Doc.DocumentElement.SetAttributeNode(schemaLocationCV)
                Doc.DocumentElement.SetAttributeNode(schemaLocation)
                Doc.DocumentElement.SetAttributeNode(atributo_version)
                Doc.DocumentElement.SetAttributeNode(atributo_rfc)
                Doc.DocumentElement.SetAttributeNode(atributo_rfcProveedorSw)
                Doc.DocumentElement.SetAttributeNode(atributo_CRE)
                Doc.DocumentElement.SetAttributeNode(atributo_sello)
                Doc.DocumentElement.SetAttributeNode(atributo_Nocertificado)
                Doc.DocumentElement.SetAttributeNode(atributo_certificado)
                Doc.DocumentElement.SetAttributeNode(atributo_fechacorte)

                'A CONTINUACIÓN LEEMOS LA TABLA DE TANQUES Y EXISTENCIASVOL
                SDATanques = New SqlClient.SqlDataAdapter("Select * from tanques inner join productos on claveproducto=nprodclave order by tanque", conn)
                DSTanques = New DataSet
                SDATanques.Fill(DSTanques, "Tanques")
                For i = 0 To DSTanques.Tables("Tanques").Rows.Count - 1

                    EXI = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:EXI", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    EXI = Doc.DocumentElement.AppendChild(EXI)
                    Dim CveProd As XmlAttribute = Doc.CreateAttribute("claveProducto")
                    Dim CveSubP As XmlAttribute = Doc.CreateAttribute("claveSubProducto")
                    Dim Octanaje As XmlAttribute = Doc.CreateAttribute("composicionOctanajeDeGasolina")
                    Dim Etanol As XmlAttribute = Doc.CreateAttribute("gasolinaConEtanol")
                    Dim CompEtanol As XmlAttribute = Doc.CreateAttribute("composicionDeEtanolEnGasolina")
                    Dim numeroTanque As XmlAttribute = Doc.CreateAttribute("numeroTanque")
                    Dim VolUtil As XmlAttribute = Doc.CreateAttribute("volumenUtil")
                    Dim VolFondaje As XmlAttribute = Doc.CreateAttribute("volumenFondaje")
                    Dim VolAgua As XmlAttribute = Doc.CreateAttribute("volumenAgua")
                    Dim VolDisponible As XmlAttribute = Doc.CreateAttribute("volumenDisponible")
                    Dim VolExtracc As XmlAttribute = Doc.CreateAttribute("volumenExtraccion")
                    Dim VolRec As XmlAttribute = Doc.CreateAttribute("volumenRecepcion")
                    Dim VolTemp As XmlAttribute = Doc.CreateAttribute("temperatura")
                    Dim FechaHoraAct As XmlAttribute = Doc.CreateAttribute("fechaYHoraEstaMedicion")
                    Dim FechaHoraAnt As XmlAttribute = Doc.CreateAttribute("fechaYHoraMedicionAnterior")


                    DRTanques = DSTanques.Tables("Tanques").Rows(i)
                    numeroTanque.Value = Trim(DRTanques("Tanque"))
                    CveProd.Value = Trim(DRTanques("claveCRE"))
                    ClaveProds(Trim(DRTanques("claveproducto")))
                    CveSubP.Value = ClaveSub
                    Octanaje.Value = DRTanques("Octanos")
                    If DRTanques("Etanol") <> 0 Then
                        Etanol.Value = "Sí"
                        CompEtanol.Value = DRTanques("Etanol")
                    Else
                        Etanol.Value = "No"
                    End If
                    StrSql = "Select * from ExistenciasVol where tanque=" + numeroTanque.Value + " and DAY(FechaMedicionActual) =" & _
                                                                  Str(Day(HoraCorte)) + " AND MONTH(FechaMedicionActual) = " + Str(Month(HoraCorte)) + " AND YEAR(FechaMedicionActual) =" + Str(Year(HoraCorte)) + " order by fechamedicionactual desc"
                    SDAExistencias = New SqlClient.SqlDataAdapter(StrSql, conn)
                    DSExistencias = New DataSet
                    SDAExistencias.Fill(DSExistencias, "ExistenciasVol")
                    DRExistencias = DSExistencias.Tables("ExistenciasVol").Rows(0)
                    VolFondaje.Value = Int(DRTanques("VolumenFondaje"))
                    VolAgua.Value = Int(DRExistencias("VolumenAgua"))
                    VolDisponible.Value = Int(DRExistencias("volumenDisponible"))
                    If VolDisponible.Value - VolFondaje.Value < 0 Then
                        VolUtil.Value = 0
                    Else
                        VolUtil.Value = VolDisponible.Value - VolFondaje.Value
                    End If
                    VolExtracc.Value = Int(DRExistencias("Volumenventa"))
                    If VolExtracc.Value < 0 Then VolExtracc.Value = 0
                    VolRec.Value = Int(DRExistencias("volumenrecibido"))
                    VolTemp.Value = DRExistencias("Temperatura")
                    FechaHoraAct.Value = Format(DRExistencias("FechaMedicionActual"), "yyyy-MM-ddTHH:mm:ss")
                    FechaHoraAnt.Value = Format(DRExistencias("FechaMedicionAnte"), "yyyy-MM-ddTHH:mm:ss")

                    Doc.DocumentElement.SetAttributeNode(numeroTanque)
                    Doc.DocumentElement.SetAttributeNode(CveProd)
                    Doc.DocumentElement.SetAttributeNode(CveSubP)
                    If DRTanques("clavecre") = 7 Then
                        Doc.DocumentElement.SetAttributeNode(Octanaje)
                        Doc.DocumentElement.SetAttributeNode(Etanol)
                        If DRTanques("Etanol") <> 0 Then Doc.DocumentElement.SetAttributeNode(CompEtanol)
                    End If
                    Doc.DocumentElement.SetAttributeNode(VolUtil)
                    Doc.DocumentElement.SetAttributeNode(VolFondaje)
                    Doc.DocumentElement.SetAttributeNode(VolAgua)
                    Doc.DocumentElement.SetAttributeNode(VolDisponible)
                    Doc.DocumentElement.SetAttributeNode(VolExtracc)
                    Doc.DocumentElement.SetAttributeNode(VolRec)
                    Doc.DocumentElement.SetAttributeNode(VolTemp)
                    Doc.DocumentElement.SetAttributeNode(FechaHoraAct)

                    EXI.Attributes.Append(numeroTanque)
                    EXI.Attributes.Append(CveProd)
                    EXI.Attributes.Append(CveSubP)
                    'SE VERIFICA SI ES GASOLINA O DIESEL PARA INCLUIR LOS ATRIBUTOS DE OCTANAJE Y ETANOL
                    If CveProd.Value = 7 Then
                        EXI.Attributes.Append(Octanaje)
                        EXI.Attributes.Append(Etanol)
                        If Etanol.Value = "Sí" Then
                            EXI.Attributes.Append(CompEtanol)
                        End If
                    End If
                    EXI.Attributes.Append(VolUtil)
                    EXI.Attributes.Append(VolFondaje)
                    EXI.Attributes.Append(VolAgua)
                    EXI.Attributes.Append(VolDisponible)
                    EXI.Attributes.Append(VolExtracc)
                    EXI.Attributes.Append(VolRec)
                    EXI.Attributes.Append(VolTemp)
                    EXI.Attributes.Append(FechaHoraAct)
                    EXI.Attributes.Append(FechaHoraAnt)
                Next

                'SE CREA NODO REC 
                REC = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:REC", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                Nodo = Doc.DocumentElement.AppendChild(REC)

                'SE CREAN LOS ATRIBUTOS DE REC
                Dim TotRec As XmlAttribute = Doc.CreateAttribute("totalRecepciones")
                Dim TotDocs As XmlAttribute = Doc.CreateAttribute("totalDocumentos")

                'SE CARGAN LAS TABLAS RECEPCIONES, RECEPCIONESCAP y PROVEEDCOMBUST

                StrSql = "select * from recepciones inner join Productos on ClaveProducto=NProdClave  where estadoenvio=0  and day(fecharecepcion)=" + Str(Day(HoraCorte)) & _
                                                              " and month(fecharecepcion)=" + Str(Month(HoraCorte)) + " and year(fecharecepcion)=" + Str(Year(HoraCorte)) + " order by folio"
                SDARecepciones = New SqlClient.SqlDataAdapter(StrSql, conn)
                DSRecepciones = New DataSet
                SDARecepciones.Fill(DSRecepciones, "Recepciones")

                SDARecepcionesCap = New SqlClient.SqlDataAdapter("select * from recepcionescap inner join Productos on ClaveProducto=NProdClave where estadoenvio=0 and fecharecepcion>='" + _
                                            HoraCorte.AddDays(-3) + "' and fecharecepcion<='" + HoraCorte + "' order by folio", conn)
                DSRecepcionesCap = New DataSet
                SDARecepcionesCap.Fill(DSRecepcionesCap, "RecepcionesCap")
                SDAProveed = New SqlClient.SqlDataAdapter("select * from ProveedCombust", conn)
                DsProveed = New DataSet
                SDAProveed.Fill(DsProveed, "ProveedCombust")

                TotRec.Value = DSRecepciones.Tables("Recepciones").Rows.Count
                TotDocs.Value = DSRecepcionesCap.Tables("Recepcionescap").Rows.Count
                Doc.DocumentElement.SetAttributeNode(TotRec)
                Doc.DocumentElement.SetAttributeNode(TotDocs)

                REC.Attributes.Append(TotRec)
                REC.Attributes.Append(TotDocs)

                'SE CREA NODO HIJO DE REC-> RECCabecera
                For i = 0 To DSRecepciones.Tables("Recepciones").Rows.Count - 1
                    RECCab = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:RECCabecera", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    RECCab = REC.AppendChild(RECCab) 'ASÍ SE CREA EL NODO HIJO

                    DRRecepciones = DSRecepciones.Tables("Recepciones").Rows(i)

                    Dim FolUniRec As XmlAttribute = Doc.CreateAttribute("folioUnicoRecepcion")
                    Dim FolUniRel As XmlAttribute = Doc.CreateAttribute("folioUnicoRelacion")
                    Dim CveProd As XmlAttribute = Doc.CreateAttribute("claveProducto")
                    Dim CveSubP As XmlAttribute = Doc.CreateAttribute("claveSubProducto")
                    Dim Octanaje As XmlAttribute = Doc.CreateAttribute("composicionOctanajeDeGasolina")
                    Dim Etanol As XmlAttribute = Doc.CreateAttribute("gasolinaConEtanol")
                    Dim CompEtanol As XmlAttribute = Doc.CreateAttribute("composicionDeEtanolEnGasolina")

                    Doc.DocumentElement.SetAttributeNode(FolUniRec)
                    Doc.DocumentElement.SetAttributeNode(CveProd)
                    Doc.DocumentElement.SetAttributeNode(CveSubP)
                    If DRRecepciones("clavecre") = 7 Then
                        Doc.DocumentElement.SetAttributeNode(Octanaje)
                        Doc.DocumentElement.SetAttributeNode(Etanol)
                        If DRRecepciones("Etanol") <> 0 Then Doc.DocumentElement.SetAttributeNode(CompEtanol)
                    End If
                    Doc.DocumentElement.SetAttributeNode(FolUniRel)
                    Octanaje.Value = DRRecepciones("Octanos")
                    'FolUniRec.Value = Foliorecepcion
                    FolUniRec.Value = DRRecepciones("Folio")
                    If DRRecepciones("Etanol") <> 0 Then
                        Etanol.Value = "Sí"
                        CompEtanol.Value = DRRecepciones("Etanol")
                    Else
                        Etanol.Value = "No"
                    End If
                    'FolUniRel.Value = Foliorecepcion
                    FolUniRel.Value = DRRecepciones("Folio")
                    ClaveProds(Trim(DRRecepciones("claveproducto")))
                    CveProd.Value = Format(ClaveP, "00")
                    CveSubP.Value = ClaveSub
                    RECCab.Attributes.Append(FolUniRec)
                    RECCab.Attributes.Append(CveProd)
                    RECCab.Attributes.Append(CveSubP)
                    If DRRecepciones("ClaveCRE") = 7 Then
                        RECCab.Attributes.Append(Octanaje)
                        RECCab.Attributes.Append(Etanol)
                        If Etanol.Value = "Sí" Then
                            RECCab.Attributes.Append(CompEtanol)
                        End If
                    End If
                    RECCab.Attributes.Append(FolUniRel)
                    'Foliorecepcion += 1
                Next

                'SE CREA NODO HIJO DE REC-> RECDetalle

                For i = 0 To DSRecepciones.Tables("Recepciones").Rows.Count - 1
                    RECDet = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:RECDetalle", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    RECDet = REC.AppendChild(RECDet) 'ASÍ SE CREA EL NODO HIJO

                    DRRecepciones = DSRecepciones.Tables("Recepciones").Rows(i)

                    Dim FolUniRec As XmlAttribute = Doc.CreateAttribute("folioUnicoRecepcion")
                    Dim Tanque As XmlAttribute = Doc.CreateAttribute("numeroDeTanque")
                    Dim VolInic As XmlAttribute = Doc.CreateAttribute("volumenInicialTanque")
                    Dim VolFin As XmlAttribute = Doc.CreateAttribute("volumenFinalTanque")
                    Dim VolRec As XmlAttribute = Doc.CreateAttribute("volumenRecepcion")
                    Dim Temp As XmlAttribute = Doc.CreateAttribute("temperatura")
                    Dim FecRec As XmlAttribute = Doc.CreateAttribute("fechaYHoraRecepcion")
                    Dim FolUniRel As XmlAttribute = Doc.CreateAttribute("folioUnicoRelacion")

                    Doc.DocumentElement.SetAttributeNode(FolUniRec)
                    Doc.DocumentElement.SetAttributeNode(Tanque)
                    Doc.DocumentElement.SetAttributeNode(VolInic)
                    Doc.DocumentElement.SetAttributeNode(VolFin)
                    Doc.DocumentElement.SetAttributeNode(VolRec)
                    Doc.DocumentElement.SetAttributeNode(Temp)
                    Doc.DocumentElement.SetAttributeNode(FecRec)
                    Doc.DocumentElement.SetAttributeNode(FolUniRel)

                    'FolUniRec.Value = Foliorecepcion
                    'Foliorecepcion += 1
                    FolUniRec.Value = DRRecepciones("Folio")
                    Tanque.Value = DRRecepciones("Tanque")
                    VolInic.Value = Int(Math.Round(DRRecepciones("VolumenInicial"), 0))
                    VolFin.Value = Int(Math.Round(DRRecepciones("VolumenFinal"), 0))
                    VolRec.Value = VolFin.Value - VolInic.Value
                    Temp.Value = DRRecepciones("TemperaturaFinal")
                    FecRec.Value = Format(DRRecepciones("FechaRecepcion"), "yyyy-MM-ddTHH:mm:ss")
                    'FolUniRel.Value = FolioRelacion
                    FolUniRel.Value = DRRecepciones("Folio")
                    'FolioRelacion += 1
                    RECDet.Attributes.Append(FolUniRec)
                    RECDet.Attributes.Append(Tanque)
                    RECDet.Attributes.Append(VolInic)
                    RECDet.Attributes.Append(VolFin)
                    RECDet.Attributes.Append(VolRec)
                    RECDet.Attributes.Append(Temp)
                    RECDet.Attributes.Append(FecRec)
                    RECDet.Attributes.Append(FolUniRel)

                Next

                'SE CREA NODO HIJO DE REC-> RECDocumentos

                DRProveed = DsProveed.Tables("ProveedCombust").Rows(0)
                For i = 0 To DSRecepcionesCap.Tables("RecepcionesCap").Rows.Count - 1
                    RECDocs = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:RECDocumentos", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    RECDocs = REC.AppendChild(RECDocs) 'ASÍ SE CREA EL NODO HIJO

                    DRRecepcionesCap = DSRecepcionesCap.Tables("RecepcionesCap").Rows(i)

                    Dim FolUniRec As XmlAttribute = Doc.CreateAttribute("folioUnicoRecepcion")
                    Dim TAR As XmlAttribute = Doc.CreateAttribute("terminalAlmacenamientoYDistribucion")
                    Dim PermAlmDis As XmlAttribute = Doc.CreateAttribute("permisoAlmacenamientoYDistribucion")
                    Dim TipoDoc As XmlAttribute = Doc.CreateAttribute("tipoDocumento")
                    Dim FechDoc As XmlAttribute = Doc.CreateAttribute("fechaDocumento")
                    Dim FolioDoc As XmlAttribute = Doc.CreateAttribute("folioDocumentoRecepcion")
                    Dim VolDoc As XmlAttribute = Doc.CreateAttribute("volumenDocumentado")
                    Dim Precio As XmlAttribute = Doc.CreateAttribute("precioCompra")
                    Dim PermTransp As XmlAttribute = Doc.CreateAttribute("permisoTransporte")
                    Dim CveVehic As XmlAttribute = Doc.CreateAttribute("claveVehiculo")
                    Dim FolUniRel As XmlAttribute = Doc.CreateAttribute("folioUnicoRelacion")
                    Dim TipoProv As XmlAttribute = Doc.CreateAttribute("tipoProveedor")
                    Dim PermImport As XmlAttribute = Doc.CreateAttribute("permisoImportacion")
                    Dim RFCProv As XmlAttribute = Doc.CreateAttribute("rfcProveedor")
                    Dim NomProv As XmlAttribute = Doc.CreateAttribute("nombreProveedor")
                    Dim PermCRE As XmlAttribute = Doc.CreateAttribute("permisoProveedor")

                    Doc.DocumentElement.SetAttributeNode(FolUniRec)
                    If Trim(DRProveed("Tipo")) <> "Nacional" Then Doc.DocumentElement.SetAttributeNode(PermAlmDis)
                    Dim TermAlm As String
                    If IsDBNull(DRRecepcionesCap("TeminalAlmacenamiento")) Then
                        If DREnvVol("enviar") Then
                            MsgBox("Documento de Recepción sin captura de TAR", MsgBoxStyle.Critical)
                            Exit Function
                        Else
                            TermAlm = ""
                        End If
                    Else
                        TermAlm = Trim(DRRecepcionesCap("TeminalAlmacenamiento"))
                        Doc.DocumentElement.SetAttributeNode(TAR)
                    End If
                    Doc.DocumentElement.SetAttributeNode(TipoDoc)
                    Doc.DocumentElement.SetAttributeNode(FechDoc)
                    Doc.DocumentElement.SetAttributeNode(FolioDoc)
                    Doc.DocumentElement.SetAttributeNode(VolDoc)
                    Doc.DocumentElement.SetAttributeNode(Precio)
                    Doc.DocumentElement.SetAttributeNode(PermTransp)
                    Doc.DocumentElement.SetAttributeNode(CveVehic)
                    Doc.DocumentElement.SetAttributeNode(FolUniRel)
                    Doc.DocumentElement.SetAttributeNode(TipoProv)
                    If Trim(DRProveed("Tipo")) <> "Nacional" Then Doc.DocumentElement.SetAttributeNode(PermImport)
                    Doc.DocumentElement.SetAttributeNode(RFCProv)
                    Doc.DocumentElement.SetAttributeNode(NomProv)
                    Doc.DocumentElement.SetAttributeNode(PermCRE)

                    'FolUniRec.Value = Foliorecepcion
                    FolUniRec.Value = DRRecepcionesCap("Folio")
                    'Foliorecepcion += 1
                    If TermAlm <> "" Then
                        TAR.Value = TermAlm
                    End If
                    TipoDoc.Value = DRRecepcionesCap("tipoDocumento")
                    FechDoc.Value = Format(DRRecepcionesCap("fechaDocumento"), "yyyy-MM-ddTHH:mm:ss")
                    FolioDoc.Value = Trim(DRRecepcionesCap("folioDocumento"))
                    VolDoc.Value = Int(Math.Round(DRRecepcionesCap("VolumenPemex"), 0))
                    Precio.Value = DRRecepcionesCap("PrecioCompra")
                    PermTransp.Value = DRRecepcionesCap("TransCRE")
                    CveVehic.Value = Trim(DRRecepcionesCap("ClaveVehiculo"))
                    'FolUniRel.Value = FolioRelacion
                    FolUniRel.Value = DRRecepcionesCap("Folio")
                    'FolioRelacion += 1
                    If Trim(DRProveed("Tipo")) = "Nacional" Then
                        TipoProv.Value = "Nacional"
                        RFCProv.Value = Trim(DRProveed("RFC"))
                        PermCRE.Value = Trim(DRProveed("PermisoCRE"))
                        PermAlmDis.Value = Trim(DRProveed("PermisoAlmDist"))
                    Else
                        TipoProv.Value = "Extranjero"
                        PermImport.Value = Trim(DRProveed("PermisoImport"))
                    End If
                    NomProv.Value = Trim(DRProveed("Nombre"))

                    RECDocs.Attributes.Append(FolUniRec)
                    If TermAlm <> "" Then
                        RECDocs.Attributes.Append(TAR)
                    End If
                    If Trim(DRProveed("Tipo")) = "Nacional" Then RECDocs.Attributes.Append(PermCRE)
                    RECDocs.Attributes.Append(TipoDoc)
                    RECDocs.Attributes.Append(FechDoc)
                    RECDocs.Attributes.Append(FolioDoc)
                    RECDocs.Attributes.Append(VolDoc)
                    RECDocs.Attributes.Append(Precio)
                    RECDocs.Attributes.Append(PermTransp)
                    RECDocs.Attributes.Append(CveVehic)
                    RECDocs.Attributes.Append(FolUniRel)
                    RECDocs.Attributes.Append(TipoProv)
                    If Trim(DRProveed("Tipo")) <> "Nacional" Then RECDocs.Attributes.Append(PermImport)
                    If Trim(DRProveed("Tipo")) = "Nacional" Then RECDocs.Attributes.Append(RFCProv)
                    RECDocs.Attributes.Append(NomProv)
                Next

                'CREAMOS EL NODO VTA
                VTA = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:VTA", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                VTA = Doc.DocumentElement.AppendChild(VTA)

                'CARGAMOS LA TABLA DE DESPACHOS
                Dim Xfecha As String = Format(HoraCorte, "dd/MM/yyyy")
                CorregirDespachosNull(Xfecha)
                StrSql = "Select * from Despachos where fechahorag='" + Xfecha + "' order by fechahora"
                SDADespachos = New SqlClient.SqlDataAdapter(StrSql, conn)
                DSDespachos = New DataSet
                SDADespachos.Fill(DSDespachos, "Despachos")

                Dim TotRegs As XmlAttribute = Doc.CreateAttribute("numTotalRegistrosDetalle")
                Doc.DocumentElement.SetAttributeNode(TotRegs)
                TotRegs.Value = DSDespachos.Tables("Despachos").Rows.Count

                VTA.Attributes.Append(TotRegs)

                'CARGAMOS DESPACHOS para el nodo VTACabecera

                SDADespachos = New SqlClient.SqlDataAdapter("SELECT posicion,manguera,claveprod,ClaveCRE,octanos,Etanol,sum(importeventa)as ventas,sum(volumen) as volumen ,COUNT(posicion) as registros from Despachos inner join productos on claveprod=nprodclave where fechahorag='" + Xfecha + "' group by POSICION,manguera,claveprod,ClaveCRE,octanos,Etanol order by posicion", conn)
                DSDespachos = New DataSet
                SDADespachos.Fill(DSDespachos, "Despachos")

                For i = 0 To DSDespachos.Tables("Despachos").Rows.Count - 1
                    DRDespachos = DSDespachos.Tables("Despachos").Rows(i)

                    'SE CREA EL NODO HIJO VTACabecera
                    VTACabecera = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:VTACabecera", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    VTACabecera = VTA.AppendChild(VTACabecera) 'ASÍ SE CREA EL NODO HIJO

                    'ATRIBUTOS
                    Dim TotRegsPos As XmlAttribute = Doc.CreateAttribute("numeroTotalRegistrosDetalle")
                    Dim Posicion As XmlAttribute = Doc.CreateAttribute("numeroDispensario")
                    Dim Manguera As XmlAttribute = Doc.CreateAttribute("identificadorManguera")
                    Dim SumVol As XmlAttribute = Doc.CreateAttribute("sumatoriaVolumenDespachado")
                    Dim SumVta As XmlAttribute = Doc.CreateAttribute("sumatoriaVentas")
                    Dim CveProd As XmlAttribute = Doc.CreateAttribute("claveProducto")
                    Dim CveSubP As XmlAttribute = Doc.CreateAttribute("claveSubProducto")
                    Dim Octanaje As XmlAttribute = Doc.CreateAttribute("composicionOctanajeDeGasolina")
                    Dim Etanol As XmlAttribute = Doc.CreateAttribute("gasolinaConEtanol")
                    Dim CompEtanol As XmlAttribute = Doc.CreateAttribute("composicionDeEtanolEnGasolina")
                    Doc.DocumentElement.SetAttributeNode(TotRegsPos)
                    Doc.DocumentElement.SetAttributeNode(Posicion)
                    Doc.DocumentElement.SetAttributeNode(Manguera)
                    Doc.DocumentElement.SetAttributeNode(CveProd)
                    Doc.DocumentElement.SetAttributeNode(CveSubP)
                    If DRDespachos("ClaveCRE") = "07" Then
                        Doc.DocumentElement.SetAttributeNode(Octanaje)
                        Doc.DocumentElement.SetAttributeNode(Etanol)
                        If DRDespachos("Etanol") <> 0 Then Doc.DocumentElement.SetAttributeNode(CompEtanol)
                    End If
                    Doc.DocumentElement.SetAttributeNode(SumVol)
                    Doc.DocumentElement.SetAttributeNode(SumVta)

                    'SE ASIGNAN LOS VALORES A LOS ATRIBUTOS
                    TotRegsPos.Value = DRDespachos("registros")
                    Posicion.Value = DRDespachos("posicion")
                    Manguera.Value = DRDespachos("manguera")

                    ClaveProds(Trim(DRDespachos("claveprod")))
                    CveProd.Value = Format(ClaveP, "00")
                    CveSubP.Value = ClaveSub
                    Octanaje.Value = DRDespachos("Octanos")
                    If DRDespachos("Etanol") <> 0 Then
                        Etanol.Value = "Sí"
                        CompEtanol.Value = DRDespachos("Etanol")
                    Else
                        Etanol.Value = "No"
                    End If
                    SumVol.Value = DRDespachos("volumen")
                    SumVta.Value = DRDespachos("ventas")

                    'SE CREAN LOS ATRIBUTOS DE VTACabecera EN DOCUMENTO
                    VTACabecera.Attributes.Append(TotRegsPos)
                    VTACabecera.Attributes.Append(Posicion)
                    VTACabecera.Attributes.Append(Manguera)
                    VTACabecera.Attributes.Append(CveProd)
                    VTACabecera.Attributes.Append(CveSubP)
                    If DRDespachos("ClaveCRE") = "07" Then
                        VTACabecera.Attributes.Append(Octanaje)
                        VTACabecera.Attributes.Append(Etanol)
                        If Etanol.Value = "Sí" Then
                            VTACabecera.Attributes.Append(CompEtanol)
                        End If
                    End If
                    VTACabecera.Attributes.Append(SumVol)
                    VTACabecera.Attributes.Append(SumVta)
                Next

                'CARGAMOS DESPACHOS  para el nodo VTADetalle
                SDADespachos = New SqlClient.SqlDataAdapter("SELECT TIPOREG,TRANSACCION,POSICION,MANGUERA,CLAVEPROD,ClaveCRE,octanos,Etanol,VOLUMEN,PRECIO,IMPORTEVENTA,FECHAHORA FROM Despachos inner join productos on claveprod=nprodclave where fechahorag='" + Xfecha + "' order by transaccion", conn)
                DSDespachos = New DataSet
                SDADespachos.Fill(DSDespachos, "Despachos")

                For i = 0 To DSDespachos.Tables("Despachos").Rows.Count - 1

                    DRDespachos = DSDespachos.Tables("Despachos").Rows(i)

                    'SE CREA EL NODO HIJO VTADetalle
                    VTADetalle = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:VTADetalle", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    VTADetalle = VTA.AppendChild(VTADetalle) 'ASÍ SE CREA EL NODO HIJO

                    'ATRIBUTOS
                    Dim TipoReg As XmlAttribute = Doc.CreateAttribute("tipoDeRegistro")
                    Dim Transac As XmlAttribute = Doc.CreateAttribute("numeroUnicoTransaccionVenta")
                    Dim Posic As XmlAttribute = Doc.CreateAttribute("numeroDispensario")
                    Dim Manguera As XmlAttribute = Doc.CreateAttribute("identificadorManguera")
                    Dim Volumen As XmlAttribute = Doc.CreateAttribute("volumenDespachado")
                    Dim PrecioU As XmlAttribute = Doc.CreateAttribute("precioUnitarioProducto")
                    Dim ImporteT As XmlAttribute = Doc.CreateAttribute("importeTotalTransaccion")
                    Dim FechaVta As XmlAttribute = Doc.CreateAttribute("fechaYHoraTransaccionVenta")
                    Dim CveProd As XmlAttribute = Doc.CreateAttribute("claveProducto")
                    Dim CveSubP As XmlAttribute = Doc.CreateAttribute("claveSubProducto")
                    Dim Octanaje As XmlAttribute = Doc.CreateAttribute("composicionOctanajeDeGasolina")
                    Dim Etanol As XmlAttribute = Doc.CreateAttribute("gasolinaConEtanol")
                    Dim CompEtanol As XmlAttribute = Doc.CreateAttribute("composicionDeEtanolEnGasolina")

                    Doc.DocumentElement.SetAttributeNode(TipoReg)
                    Doc.DocumentElement.SetAttributeNode(Transac)
                    Doc.DocumentElement.SetAttributeNode(Posic)
                    Doc.DocumentElement.SetAttributeNode(Manguera)
                    Doc.DocumentElement.SetAttributeNode(CveProd)
                    Doc.DocumentElement.SetAttributeNode(CveSubP)
                    If DRDespachos("ClaveCRE") = "07" Then
                        Doc.DocumentElement.SetAttributeNode(Octanaje)
                        Doc.DocumentElement.SetAttributeNode(Etanol)
                        If DRDespachos("Etanol") <> 0 Then Doc.DocumentElement.SetAttributeNode(CompEtanol)
                    End If
                    Doc.DocumentElement.SetAttributeNode(Volumen)
                    Doc.DocumentElement.SetAttributeNode(PrecioU)
                    Doc.DocumentElement.SetAttributeNode(ImporteT)
                    Doc.DocumentElement.SetAttributeNode(FechaVta)

                    'SE ASIGNAN LOS VALORES A LOS ATRIBUTOS
                    TipoReg.Value = DRDespachos("TIPOREG")
                    Transac.Value = DRDespachos("TRANSACCION")
                    Posic.Value = DRDespachos("POSICION")
                    Manguera.Value = DRDespachos("MANGUERA")
                    ClaveProds(Trim(DRDespachos("claveprod")))
                    CveProd.Value = Format(ClaveP, "00")
                    CveSubP.Value = ClaveSub
                    Octanaje.Value = DRDespachos("Octanos")
                    If DRDespachos("Etanol") <> 0 Then
                        Etanol.Value = "Sí"
                        CompEtanol.Value = DRDespachos("Etanol")
                    Else
                        Etanol.Value = "No"
                    End If
                    Volumen.Value = DRDespachos("volumen")
                    PrecioU.Value = DRDespachos("PRECIO")
                    ImporteT.Value = DRDespachos("IMPORTEVENTA")
                    FechaVta.Value = Format(DRDespachos("FECHAHORA"), "yyyy-MM-ddTHH:mm:ss")

                    'SE AGREGAN LOS ATRIBUTOS DE VTADetalle EN DOCUMENTO
                    VTADetalle.Attributes.Append(TipoReg)
                    VTADetalle.Attributes.Append(Transac)
                    VTADetalle.Attributes.Append(Posic)
                    VTADetalle.Attributes.Append(Manguera)
                    VTADetalle.Attributes.Append(CveProd)
                    VTADetalle.Attributes.Append(CveSubP)
                    If DRDespachos("ClaveCRE") = "07" Then
                        VTADetalle.Attributes.Append(Octanaje)
                        VTADetalle.Attributes.Append(Etanol)
                        If Etanol.Value = "Sí" Then
                            VTADetalle.Attributes.Append(CompEtanol)
                        End If
                    End If
                    VTADetalle.Attributes.Append(Volumen)
                    VTADetalle.Attributes.Append(PrecioU)
                    VTADetalle.Attributes.Append(ImporteT)
                    VTADetalle.Attributes.Append(FechaVta)
                Next

                'CREAMOS EL NODO TQS
                For i = 0 To DSTanques.Tables("Tanques").Rows.Count - 1
                    DRTanques = DSTanques.Tables("Tanques").Rows(i)
                    TQS = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:TQS", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    TQS = Doc.DocumentElement.AppendChild(TQS)

                    'ATRIBUTOS
                    Dim NoTanq As XmlAttribute = Doc.CreateAttribute("numeroTanque")
                    Dim CapTot As XmlAttribute = Doc.CreateAttribute("capacidadTotalTanque")
                    Dim CapOp As XmlAttribute = Doc.CreateAttribute("capacidadOperativaTanque")
                    Dim CapUt As XmlAttribute = Doc.CreateAttribute("capacidadUtilTanque")
                    Dim CapFond As XmlAttribute = Doc.CreateAttribute("capacidadFondajeTanque")
                    Dim VolMin As XmlAttribute = Doc.CreateAttribute("volumenMinimoOperacion")
                    Dim Estado As XmlAttribute = Doc.CreateAttribute("estadoTanque")
                    Dim CveProd As XmlAttribute = Doc.CreateAttribute("claveProducto")
                    Dim CveSubP As XmlAttribute = Doc.CreateAttribute("claveSubProducto")
                    Dim Octanaje As XmlAttribute = Doc.CreateAttribute("composicionOctanajeDeGasolina")
                    Dim Etanol As XmlAttribute = Doc.CreateAttribute("gasolinaConEtanol")
                    Dim CompEtanol As XmlAttribute = Doc.CreateAttribute("composicionDeEtanolEnGasolina")

                    Doc.DocumentElement.SetAttributeNode(NoTanq)
                    Doc.DocumentElement.SetAttributeNode(CveProd)
                    Doc.DocumentElement.SetAttributeNode(CveSubP)
                    If DRTanques("ClaveCRE") = "07" Then
                        Doc.DocumentElement.SetAttributeNode(Octanaje)
                        Doc.DocumentElement.SetAttributeNode(Etanol)
                        If DRTanques("Etanol") <> 0 Then Doc.DocumentElement.SetAttributeNode(CompEtanol)
                    End If
                    Doc.DocumentElement.SetAttributeNode(CapTot)
                    Doc.DocumentElement.SetAttributeNode(CapOp)
                    Doc.DocumentElement.SetAttributeNode(CapUt)
                    Doc.DocumentElement.SetAttributeNode(CapFond)
                    Doc.DocumentElement.SetAttributeNode(VolMin)
                    Doc.DocumentElement.SetAttributeNode(Estado)

                    'SE ASIGNAN LOS VALORES A LOS ATRIBUTOS
                    NoTanq.Value = Int(DRTanques("Tanque"))
                    ClaveProds(Trim(DRTanques("claveproducto")))
                    CveProd.Value = Format(ClaveP, "00")
                    CveSubP.Value = ClaveSub
                    Octanaje.Value = DRTanques("Octanos")
                    If DRTanques("Etanol") <> 0 Then
                        Etanol.Value = "Sí"
                        CompEtanol.Value = DRTanques("Etanol")
                    Else
                        Etanol.Value = "No"
                    End If
                    CapTot.Value = Int(DRTanques("VolumenTanque"))
                    CapOp.Value = Int(DRTanques("VolumenOperativo"))
                    CapUt.Value = Int(DRTanques("VolumenUtil"))
                    CapFond.Value = Int(DRTanques("VolumenFondaje"))
                    VolMin.Value = Int(DRTanques("VolumenMinimo"))
                    If DRTanques("EstadoTanque") = 0 Then
                        Estado.Value = "O"
                    Else
                        Estado.Value = "F"
                    End If

                    'SE AGREGAN AL DOCUMENTO
                    TQS.Attributes.Append(NoTanq)
                    TQS.Attributes.Append(CveProd)
                    TQS.Attributes.Append(CveSubP)
                    If DRTanques("ClaveCRE") = 7 Then
                        TQS.Attributes.Append(Octanaje)
                        TQS.Attributes.Append(Etanol)
                        If Etanol.Value = "Sí" Then
                            TQS.Attributes.Append(CompEtanol)
                        End If
                    End If
                    TQS.Attributes.Append(CapTot)
                    TQS.Attributes.Append(CapOp)
                    TQS.Attributes.Append(CapUt)
                    TQS.Attributes.Append(CapFond)
                    TQS.Attributes.Append(VolMin)
                    TQS.Attributes.Append(Estado)
                Next

                'CARGAMOS LA TABLA DE DISPENSARIOS
                SDADispensarios = New SqlClient.SqlDataAdapter("Select * from Dispensarios inner join productos on NCodProd=nprodclave order by Npos,Nmang", conn)
                DSDispensarios = New DataSet
                SDADispensarios.Fill(DSDispensarios, "Dispensarios")

                'CREAMOS EL NODO DIS
                For i = 0 To DSDispensarios.Tables("Dispensarios").Rows.Count - 1
                    DRDispensarios = DSDispensarios.Tables("Dispensarios").Rows(i)
                    DIS = Doc.CreateNode(XmlNodeType.Element, "controlesvolumetricos:DIS", "http://www.sat.gob.mx/esquemas/controlesvolumetricos")
                    DIS = Doc.DocumentElement.AppendChild(DIS)

                    'ATRIBUTOS
                    Dim NoDisp As XmlAttribute = Doc.CreateAttribute("numeroDispensario")
                    Dim IdMang As XmlAttribute = Doc.CreateAttribute("identificadorManguera")
                    Dim CveProd As XmlAttribute = Doc.CreateAttribute("claveProducto")
                    Dim CveSubP As XmlAttribute = Doc.CreateAttribute("claveSubProducto")
                    Dim Octanaje As XmlAttribute = Doc.CreateAttribute("composicionOctanajeDeGasolina")
                    Dim Etanol As XmlAttribute = Doc.CreateAttribute("gasolinaConEtanol")
                    Dim CompEtanol As XmlAttribute = Doc.CreateAttribute("composicionDeEtanolEnGasolina")

                    Doc.DocumentElement.SetAttributeNode(NoDisp)
                    Doc.DocumentElement.SetAttributeNode(IdMang)
                    Doc.DocumentElement.SetAttributeNode(CveProd)
                    Doc.DocumentElement.SetAttributeNode(CveSubP)
                    If DRDispensarios("ClaveCRE") = 7 Then
                        Doc.DocumentElement.SetAttributeNode(Octanaje)
                        Doc.DocumentElement.SetAttributeNode(Etanol)
                        If DRDispensarios("Etanol") <> 0 Then Doc.DocumentElement.SetAttributeNode(CompEtanol)
                    End If
                    'SE ASIGNAN LOS VALORES A LOS ATRIBUTOS
                    NoDisp.Value = Int(DRDispensarios("nPos"))
                    IdMang.Value = Int(DRDispensarios("nMang"))
                    ClaveProds(Trim(DRDispensarios("nCodProd")))
                    CveProd.Value = Format(ClaveP, "00")
                    CveSubP.Value = ClaveSub
                    Octanaje.Value = DRDispensarios("Octanos")
                    If DRDispensarios("Etanol") <> 0 Then
                        Etanol.Value = "Sí"
                        CompEtanol.Value = DRDispensarios("Etanol")
                    Else
                        Etanol.Value = "No"
                    End If

                    'SE AGREGAN AL DOCUMENTO
                    DIS.Attributes.Append(NoDisp)
                    DIS.Attributes.Append(IdMang)
                    DIS.Attributes.Append(CveProd)
                    DIS.Attributes.Append(CveSubP)
                    If DRDispensarios("ClaveCRE") = 7 Then
                        DIS.Attributes.Append(Octanaje)
                        DIS.Attributes.Append(Etanol)
                        If Etanol.Value = "Sí" Then
                            DIS.Attributes.Append(CompEtanol)
                        End If
                    End If
                Next

                'SE CREA EL DOCUMENTO XML
                NombreDoc = Replace(PermisoCRE, "/", "_") + Format(Now, "yyyyMMdd.HHmmss") + RFCEstacion + ".XML"
                Dim ArchMov = NombreDoc
                Dim ArchTemp As String = appPath + "\xmltemp\xmltemp.xml"
                Dim ArchZip As String = NombreDoc.Replace("XML", "zip")
                'SI NO EXISTE EL DIRECTORIO SE CREA
                Dim Directorio As String = appPath + "\xmltemp"

                'SI NO EXISTE EL DIRECTORIO SE CREA
                If Not Directory.Exists(Directorio) Then
                    Directory.CreateDirectory(Directorio)
                End If

                'Doc.Save(ArchTemp) 'SE QUITA LA COMILLA INICIAL PARA DETECTAR ERRORES

                'SE GENERA E INSERTA SELLO 
                atributo_sello.Value = Trim(CreaSello(Doc))
                NombreDoc = appPath + "\xmltemp\" + NombreDoc
                Doc.Save(NombreDoc)

                'SE BORRA EL TEMPORAL(sin sello)


                'SE COMPRIME ARCHIVO XML->ZIP Y SE COLOCA EN DIRECTORIO POR ENVIAR
                If File.Exists("C:\controlvolumetricoporenviar\" + ArchZip) Then
                    My.Computer.FileSystem.DeleteFile("C:\controlvolumetricoporenviar\" + ArchZip)
                End If
                ZipFile.CreateFromDirectory(Directorio, "C:\controlvolumetricoporenviar\" + ArchZip)

                'SE MUEVE EL ARCHIVO XML AL DIRECTORIO ControlVolumetrico
                If File.Exists("C:\controlvolumetrico\" + ArchMov) Then
                    My.Computer.FileSystem.DeleteFile("C:\controlvolumetrico\" + ArchMov)
                End If
                My.Computer.FileSystem.MoveFile(NombreDoc, "C:\controlvolumetrico\" + ArchMov)
                CREAXML = "CREADA"
            Else
                CREAXML = "ESTACIÓN DE SERVICIO NO AUTORIZADA"
            End If
        Catch ex As Exception
            CREAXML = ex.Message
        End Try

    End Function
   
End Class
