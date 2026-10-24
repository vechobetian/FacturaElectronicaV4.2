Imports System.Collections.Generic

Public Class ImedPayload
    Public Property MensajeImedSigedad As MensajeImedSigedadData
End Class

Public Class MensajeImedSigedadData
    Public Property MensajeImedSigedadVersion As String
    Public Property EncabezadoMensaje As EncabezadoMensajeData
    Public Property Comprobante As ComprobanteData
    Public Property DatosVentas As List(Of VentaItemData)
End Class

Public Class EncabezadoMensajeData
    Public Property TipoMsj As Integer
    Public Property CodAccion As Integer
    Public Property IdMsj As Integer
    Public Property InicioTrx As InicioTrxData
    Public Property Software As SoftwareData
    Public Property Prestador As PrestadorData
End Class

Public Class InicioTrxData
    Public Property Fecha As String
    Public Property Hora As String
End Class

Public Class SoftwareData
    Public Property CodigoADESFA As String
    Public Property Nombre As String
    Public Property Version As String
End Class

Public Class PrestadorData
    Public Property GLN As String
    Public Property CodigoImed As Long
    Public Property Cuit As String
    Public Property CUFE As String
End Class

Public Class ComprobanteData
    Public Property Id As Long
    Public Property PtoVenta As Integer
    Public Property TipoComprobanteARCA As Integer
    Public Property Fecha As String
    Public Property Hora As String
    Public Property Tipo As String
    Public Property Fiscal As String
    Public Property CAE As Long
    Public Property ImporteTotal As Decimal
    Public Property TipoDocumento As Integer
    Public Property NroDocumento As Long
    Public Property PersonaRetiro As PersonaRetiroData
End Class

Public Class PersonaRetiroData
    Public Property TipoDocumentoR As Integer
    Public Property NroDocumentoR As Long
    Public Property FechaNacimientoR As String
    Public Property SexoR As String
    Public Property LecturaDNI As String
End Class

Public Class VentaItemData
    Public Property Item As ItemContenidoData
End Class

Public Class ItemContenidoData
    Public Property NroOrden As Integer
    Public Property ItemVenta As ItemVentaData
    Public Property RecetaDigital As RecetaDigitalData
    Public Property Validador As ValidadorData
End Class

Public Class ItemVentaData
    Public Property OrdenItemVenta As Integer
    Public Property NroReferenciaSGF As String
    Public Property NroReferenciaSGFOrigen As String
    Public Property CodBarras As Long
    Public Property CapturaCodBarras As String
    Public Property CodTroquel As Long
    Public Property Alfabeta As Integer
    Public Property Kairos As String
    Public Property Codigo As String
    Public Property IdProducto As String
    Public Property Descripcion As String
    Public Property CantidadVenta As Integer
    Public Property ImporteUnitario As Decimal
End Class

Public Class RecetaDigitalData
    Public Property Fecha As String
    Public Property NroRecetaDigital As Long
    Public Property NroReferenciaConsulta As String
    Public Property TipoPrescriptor As String
    Public Property TipoMatricula As String
    Public Property ProvinciaMatricula As String
    Public Property NroMatricula As Integer
    Public Property IdProductoSolicitado As String
    Public Property AlfabetaSolicitado As String
    Public Property DrogaCodigoSolicitado As String
    Public Property CodDrogaSnomedSolicitado As String
    Public Property DrogaDescripcionSolicitado As String
    Public Property CantidadSolicitado As Integer
    Public Property EstadoSugerido As Integer
    Public Property Financiador As FinanciadorData
End Class

Public Class FinanciadorData
    Public Property CodigoADESFA As String
    Public Property Codigo As String
End Class

Public Class ValidadorData
    Public Property CodigoADESFA As String
    Public Property Nombre As String
    Public Property NroReferencia As String
    Public Property CantidadAprobada As Integer
End Class

