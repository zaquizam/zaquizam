<!--#include file="conexionRS.asp" -->
<%
'
' RetMen_llenar_cmb5.asp - 15jul21 - 27ene22
'
' Cambio en combo Marca - 
'
Server.ScriptTimeout = 30000
Response.Buffer = True	
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "UTF-8"
'
if conexionRS.errors.count <> 0 Then
  Response.Write ("No hay conexionRS con la BD...!")
  Response.End
end if

Dim opcion, qrySql, idCat, idCliente, idMar
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
'
opcion = Request.Querystring("opcion")
idCat  = Request.Querystring("idCat")
idArea = Request.Querystring("idArea")
idZona = Request.Querystring("idZona")
idCanal = Request.Querystring("idCanal")
idFab   = Request.Querystring("idFab")
idCliente = Request.Querystring("idCli")
'
IF (Cint(opcion) = 5) THEN
	'
	'ReFill combo Marca
	'			
	Dim rsMarca, arrMarca
	'
	' Buscar Datos de todas las Marcas
	'	
	if idCat >= 127 and idCat <= 145 then
		qrySql = vbnullstring	
		qrySql = " SELECT Id_Marca as id, Trim(Marca) + '('+Trim(Fabricante)+')' as nombre FROM RS_DataProcSem WHERE Id_Categoria = " & idCat
		if Len(idArea)<>0 then 
			qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
		end if	
		if Len(idZona)<>0 then 
			qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
		end if
		if Len(idCanal)<>0 then 
			qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
		end if
		if Len(idFab)<>0 then 
			qrySql = qrySql & " AND Id_Fabricante in (" & idFab & ")"
		end if	
		qrySql = qrySql & " GROUP BY Id_Marca, Trim(Marca)+'('+Trim(Fabricante)+')' HAVING Id_Marca <> 0 ORDER BY Trim(Marca)+'('+Trim(Fabricante)+')'"
	else 
		qrySql = vbnullstring	
		qrySql = " SELECT Id_Marca as id, Marca as nombre FROM RS_DataProcSem WHERE Id_Categoria = " & idCat
		if Len(idArea)<>0 then 
			qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
		end if	
		if Len(idZona)<>0 then 
			qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
		end if
		if Len(idCanal)<>0 then 
			qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
		end if
		if Len(idFab)<>0 then 
			qrySql = qrySql & " AND Id_Fabricante in (" & idFab & ")"
		end if	
		qrySql = qrySql & " GROUP BY Id_Marca, Marca HAVING Id_Marca <> 0 ORDER BY Marca "
	end if
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMarca = Server.CreateObject("ADODB.recordset")
	rsMarca.Open qrySql,conexionRS,0,1
	'
	if not rsMarca.EOF then
		arrMarca = rsMarca.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsMarca.Close : Set rsMarca = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrMarca) then

		For i = 0 to ubound(arrMarca, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrMarca(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrMarca(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.close : set conexionRS = nothing
	'		
ELSEIF (Cint(opcion) = 6) THEN
	'
	'ReFill combo Segmento
	'			
	Dim rsSegmento, arrSegmento		
	'
	' Buscar Datos de todas las Segmento
	'
	qrySql = vbnullstring	
	qrySql = " SELECT DISTINCT Id_Segmento as id, Segmento as nombre  FROM  RS_DataProcSem  WHERE Id_Categoria = " & idCat
	if Len(idArea)<>0 then 
		qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
	end if
	if Len(idZona)<>0 then 
		qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
	end if
	if Len(idCanal)<>0 then 
		qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
	end if
	if Len(idFab)<>0 then 
		qrySql = qrySql & " AND Id_Fabricante in (" & idFab & ")"
	end if	
	qrySql = qrySql & " AND Id_Segmento <> 0 ORDER BY  Segmento"	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSegmento = Server.CreateObject("ADODB.recordset")
	rsSegmento.Open qrySql,conexionRS,0,1
	'
	if not rsSegmento.EOF then
		arrSegmento = rsSegmento.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsSegmento.Close : Set rsSegmento = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrSegmento) then

		For i = 0 to ubound(arrSegmento, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrSegmento(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrSegmento(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.close : set conexionRS = nothing
	'	
ELSEIF (Cint(opcion) = 7) THEN
	'
	'ReFill combo Tama??o
	'			
	Dim rsTamano, arrTamano		
	'
	' Buscar Datos de todas las Tamano
	'
	qrySql = vbnullstring	
	qrySql = " SELECT DISTINCT Id_Tamano as id, Tamano as nombre FROM RS_DataProcSem  WHERE Id_Categoria = " & idCat
	if Len(idArea)<>0 then 
		qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
	end if
	if Len(idZona)<>0 then 
		qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
	end if
	if Len(idCanal)<>0 then 
		qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
	end if
	if Len(idFab)<>0 then 
		qrySql = qrySql & " AND Id_Fabricante in (" & idFab & ")"
	end if
	
	qrySql = qrySql & " AND Id_Tamano <> 0 ORDER BY Tamano"	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsTamano = Server.CreateObject("ADODB.recordset")
	rsTamano.Open qrySql,conexionRS,0,1
	'
	if not rsTamano.EOF then
		arrTamano = rsTamano.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsTamano.Close : Set rsTamano = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrTamano) then

		For i = 0 to ubound(arrTamano, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrTamano(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrTamano(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.close : set conexionRS = nothing
	'	
ELSEIF (Cint(opcion) = 8) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'	
	'17nov
	qrySql = vbnullstring	
	qrySql = " SELECT RS_DataProcSem.CodigoBarra as id, TRIM(RS_DataProcSem.Descripcion) as nombre FROM RS_DataProcSem INNER JOIN PH_CB_Fabricante ON RS_DataProcSem.Id_Fabricante = PH_CB_Fabricante.id_Fabricante WHERE RS_DataProcSem.Id_Categoria = " & idCat
	if Len(idArea)<>0 then 
		qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
	end if	
	if Len(idZona)<>0 then 
		qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
	end if		
	if Len(idCanal)<>0 then 
		 qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
	end if	
	if Len(idFab)<>0 then 
		 qrySql = qrySql & " AND PH_CB_Fabricante.id_Fabricante in (" & idFab & ")"
	end if	
	qrySql = qrySql & " AND PH_CB_Fabricante.Ind_MarcaPropia = 0 GROUP BY RS_DataProcSem.CodigoBarra, RS_DataProcSem.Descripcion HAVING ( RS_DataProcSem.CodigoBarra IS NOT NULL AND RS_DataProcSem.CodigoBarra <> '' ) AND ( RS_DataProcSem.Descripcion IS NOT NULL AND RS_DataProcSem.Descripcion <> '' ) ORDER BY nombre"	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsProducto = Server.CreateObject("ADODB.recordset")
	rsProducto.Open qrySql,conexionRS,0,1
	'
	if not rsProducto.EOF then
		arrProducto = rsProducto.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsProducto.Close : Set rsProducto = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrProducto) then

		For i = 0 to ubound(arrProducto, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrProducto(0,i) & chr(34) & chr(44)
			'sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrProducto(1,i) & chr(34) & chr(125) &chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & RemoverSaltodeLinea(arrProducto(1,i)) & " - " & arrProducto(0,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.close : set conexionRS = nothing
	'		
ELSE
	' de lo Contrario
	Response.write "error"
END IF
'

%>