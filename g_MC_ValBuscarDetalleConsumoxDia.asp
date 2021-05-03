<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_MC_ValBuscarDetalleConsumoxDia.asp // 26feb21 - 
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, rsDetalleConsumo, arrDetalleConsumo
	'	
	idConsumo	=	Request.Querystring("id_Consumo")	
	'
	' Buscar los detalles del Consumo
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Consumo.Tiene_Factura,"
	QrySql = QrySql & " PH_Consumo.Id_Canal,"
	QrySql = QrySql & " PH_Consumo.Id_Cadena,"	
	QrySql = QrySql & " PH_Canal.Canal,"
	QrySql = QrySql & " PH_Cadena.Cadena,"
	QrySql = QrySql & " PH_Consumo.Total_Compra,"
	QrySql = QrySql & " PH_Consumo.total_Items,"
	QrySql = QrySql & " PH_Moneda.Id_Moneda,"
	QrySql = QrySql & " PH_Consumo.Validado,"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar,"
	QrySql = QrySql & " PH_Consumo.Resuelto"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN PH_Canal  ON PH_Consumo.Id_Canal  = PH_Canal.Id_Canal"
	QrySql = QrySql & " INNER JOIN PH_Cadena ON PH_Consumo.Id_Cadena = PH_Cadena.Id_Cadena"
	QrySql = QrySql & " INNER JOIN PH_Moneda ON PH_Consumo.Id_Moneda = PH_Moneda.Id_Moneda"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Consumo = " & idConsumo
	'QrySql = QrySql & " AND"
	'QrySql = QrySql & " PH_Consumo.Status_registro = 'G'"
	'
	' Response.Write QrySql '& "<BR><BR>"
	' Response.end
	'
	Set rsDetalleConsumo = Server.CreateObject("ADODB.recordset")
	rsDetalleConsumo.Open QrySql, conexion
	'
	if not rsDetalleConsumo.EOF then
    	arrDetalleConsumo = rsDetalleConsumo.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"		
	'
	sTabla=vbnullstring
	
	if IsArray(arrDetalleConsumo) then
	
		For i = 0 to ubound(arrDetalleConsumo, 2)
			sTabla    =   chr(123)&  chr(34) & "tienefactura"	& chr(34) & ":" & chr(34) & arrDetalleConsumo(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "canal" 	    	& chr(34) & ":" & chr(34) & arrDetalleConsumo(1,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "cadena" 		& chr(34) & ":" & chr(34) & arrDetalleConsumo(2,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "totalproductos"	& chr(34) & ":" & chr(34) & arrDetalleConsumo(6,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "moneda" 		& chr(34) & ":" & chr(34) & arrDetalleConsumo(7,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "validado" 		& chr(34) & ":" & chr(34) & arrDetalleConsumo(8,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "investigar" 	& chr(34) & ":" & chr(34) & arrDetalleConsumo(9,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "resuelto" 		& chr(34) & ":" & chr(34) & arrDetalleConsumo(10,i) & chr(34) & chr(44)
			'sTabla    =    sTabla &  chr(34) & "semana" 		& chr(34) & ":" & chr(34) & arrDetalleConsumo(8,i) & chr(34) & chr(44)
			'
			total = replace(arrDetalleConsumo(5,i),",",".")
			'
			sTabla    =    sTabla &  chr(34) & "totalcompra"    & chr(34) & ":" & chr(34) & total & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
		next
		'
		sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		'
	else
		'Eof()
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	end if
	''
	Response.Write(JsonData)
	conexion.close
	set conexion = nothing	
	
%>