Imports System.IO
Imports System.Data.SqlClient

Module CargaTablas
    Public DSFolios, DSTanques, DSEmpresa, DSExistencias, DSRecepciones, DSRecepcionesCap, DSDespachos, DSDispensarios, DsProveed, DSEnvVol, DSActivas As DataSet
    Public SDAFolios, SDATanques, SDAEmpresa, SDAExistencias, SDARecepciones, SDARecepcionesCap, SDADespachos, SDADispensarios, SDAProveed, SDAEnvVol, SDAActivas As SqlDataAdapter
    Public DRFolios, DRTanques, DREmpresa, DRExistencias, DRRecepciones, DRRecepcionesCap, DRDespachos, DRDispensarios, DRProveed, DREnvVol, DRActivas As DataRow
    Public Sub CargaEnviosVol()
        SDAEnvVol = New SqlClient.SqlDataAdapter("Select * from EnviosVol order by horaenvio desc", conn)
        DSEnvVol = New DataSet
        SDAEnvVol.Fill(DSEnvVol, "EnviosVol")
        DREnvVol = DSEnvVol.Tables("EnviosVol").Rows(0)
    End Sub
    Sub CargaEmpresa()
        SDAEmpresa = New SqlDataAdapter("Select * from EMPRESASF", conn)
        DSEmpresa = New DataSet
        SDAEmpresa.Fill(DSEmpresa, "EMPRESASF")
        DREmpresa = DSEmpresa.Tables("EMPRESASF").Rows(0)
    End Sub
     
End Module
