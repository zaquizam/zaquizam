 <%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValBuscarImagenFacturaxDia.asp // 30dic20 - 20ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, rsImagenFactura, arrImagenFactura
	Dim FSOobj,FilePath	
	Set FSOobj = Server.CreateObject("Scripting.FileSystemObject")
	'	
	idConsumo	=	Request.Querystring("id_Consumo")	
	'
	' Buscar los detalles de la Imagen de Factura
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " PH_Consumo_Detalle_Factura.Id_Consumo,"	
	QrySql = QrySql & " PH_Consumo_Detalle_Factura.Nombre_identificador"	
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo_Detalle_Factura"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo_Detalle_Factura.Id_Consumo = " & idConsumo	
	'	
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsImagenFactura = Server.CreateObject("ADODB.recordset")
	rsImagenFactura.Open QrySql, conexion
	'	
	if not rsImagenFactura.EOF then
    	arrImagenFactura = rsImagenFactura.GetRows() ' Convert recordset to 2D Array
	end if
	'
	set rsImagenFactura = nothing
	'
	'Response.ContentType = "application/json"
	'
	sTabla=vbnullstring
	'
	if IsArray(arrImagenFactura) then
		
		For i = 0 to ubound(arrImagenFactura, 2)
		
		
			FilePath=Server.MapPath("images/facturas/" & arrImagenFactura(1,i)) ' located in the same director
			
			if FSOobj.fileExists(FilePath) Then
				
				sTabla    =   chr(123)&  chr(34) & "id" & chr(34) & ":" & chr(34) & arrImagenFactura(0,i) & chr(34) & chr(44)
							'
				'sTabla    =    sTabla &  chr(34) & "imagen"&(i+1) & chr(34) & ":" & chr(34) & arrImagenFactura(0,i) & chr(34) & chr(125) & chr(44)
				'			
				sTabla    =    sTabla &  chr(34) & "imagen" & chr(34) & ":" & chr(34) & arrImagenFactura(1,i) & chr(34) & chr(125) & chr(44)		

			Else
				sTabla    =   chr(123)&  chr(34) & "id" & chr(34) & ":" & chr(34) & "0" & chr(34) & chr(44)
				sTabla    =   sTabla &  chr(34) & "imagen" & chr(34) & ":" & chr(34) & "ndSin_imagen_disponible2.jpg" & chr(34) & chr(125) & chr(44)						
			End if	
			'			    				
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
		next
		'			
		sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		
	else
		'Eof()		
		sTabla    =   chr(123)&  chr(34) & "id" & chr(34) & ":" & chr(34) & "0" & chr(34) & chr(44)
		sTabla    =   sTabla &  chr(34) & "imagen" & chr(34) & ":" & chr(34) & "Sin_imagen_disponible2.jpg" & chr(34) & chr(125) & chr(44)		
		'
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'		
		sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		'
	end if
	'
	Response.Write(JsonData)
	conexion.close
	set conexion = nothing	
	Set FSOobj = Nothing
	''
%>