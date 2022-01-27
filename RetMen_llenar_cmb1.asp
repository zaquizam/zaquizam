<!--#include file="conexionRS.asp" -->
<%
'
'RetMen_llenar_cmb1.asp - 12jul21 - 12ene22
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
Server.ScriptTimeout=10000000
'
if conexionRS.errors.count <> 0 Then
  Response.Write ("No hay conexionRS con la BD...!")
  Response.End
end if

Dim opcion, qrySql, idCat, idCliente
'
opcion = Request.Querystring("opcion")
idCat  = Request.Querystring("idCat")
idCliente = Request.Querystring("idCli")
'
IF (Cint(opcion) = 1) THEN
	'
	'Fill combo Categoria
	'				
	Dim rsCategoria, arrCategoria
	'
	' Buscar Datos de todas las Categorias
	' 03nov21
	'
	IF (Cint(idCliente) = 1 ) THEN
		qrySql = vbNullstring
		qrySql = qrySql & " SELECT RS_DataProcSem.Id_Categoria AS id,  RS_DataProcSem.Categoria AS nombre FROM dbo.RS_DataProcSem"
		qrySql = qrySql & " GROUP BY RS_DataProcSem.Id_Categoria, RS_DataProcSem.Categoria ORDER BY RS_DataProcSem.Categoria ASC"
	ELSE
		qrySql = vbNullstring
		qrySql = qrySql & " SELECT RS_DataProcSem.Id_Categoria AS id,  RS_DataProcSem.Categoria AS nombre FROM dbo.RS_DataProcSem"		
		qrySql = qrySql & " INNER JOIN dbo.ss_ClienteCategoria ON  RS_DataProcSem.Id_Categoria = ss_ClienteCategoria.Id_Categoria"
		qrySql = qrySql & " WHERE"
		qrySql = qrySql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		qrySql = qrySql & " and ss_ClienteCategoria.Ind_Mensual = 1 "
		qrySql = qrySql & " GROUP BY RS_DataProcSem.Id_Categoria, RS_DataProcSem.Categoria ORDER BY RS_DataProcSem.Categoria ASC"
	END IF			
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordSet")
	rsCategoria.Open qrySql, conexionRS
	'
	if not rsCategoria.EOF then
		arrCategoria = rsCategoria.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsCategoria.Close : Set rsCategoria = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrCategoria) then

		For i = 0 to ubound(arrCategoria, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCategoria(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCategoria(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'
ELSEIF (Cint(opcion) = 2) THEN
	'
	'Fill combo Area
	'			
	Dim rsArea, arrArea
	'
	' Buscar Datos de todas las Areas
	'
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " Id_Area as id, "
	qrySql = qrySql & " Area as mombre "
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_DataProcSem "
	qrySql = qrySql & " WHERE "
	qrySql = qrySql & " Id_Categoria = " & idCat
	qrySql = qrySql & " GROUP BY "
	qrySql = qrySql & " Id_Area, "
	qrySql = qrySql & " Area "
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " Id_Area "	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsArea = Server.CreateObject("ADODB.recordSet")
	rsArea.Open qrySql, conexionRS
	'
	if not rsArea.EOF then
		arrArea = rsArea.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsArea.Close : Set rsArea = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrArea) then

		For i = 0 to ubound(arrArea, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrArea(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrArea(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'	
ELSEIF (Cint(opcion) = 3) THEN
	'
	'Fill combo Zona
	'			
	Dim rsZona, arrZona
	'
	' Buscar Datos de todas las Zonas
	'
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " Id_Zona, "
	qrySql = qrySql & " Zona "
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_DataProcSem "
	qrySql = qrySql & " WHERE  "
	qrySql = qrySql & " Id_Categoria= " & idCat
	qrySql = qrySql & " GROUP BY "
	qrySql = qrySql & " Id_Zona, "
	qrySql = qrySql & " Zona "
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " Zona "
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsZona = Server.CreateObject("ADODB.recordSet")
	rsZona.Open qrySql, conexionRS
	'
	if not rsZona.EOF then
		arrZona = rsZona.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsZona.Close : Set rsZona = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrZona) then

		For i = 0 to ubound(arrZona, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrZona(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrZona(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'	
	
ELSEIF (Cint(opcion) = 4) THEN
	'
	'Fill combo Canal
	'			
	Dim rsCanal, arrCanal
	'
	' Buscar Datos de todas las Canales
	'
	qrySql = vbNullstring
	
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " Id_Canal as id, "
	qrySql = qrySql & " rtrim(Canal) as nombre"
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_DataProcSem "
	qrySql = qrySql & " WHERE "
	qrySql = qrySql & " Id_Categoria = " & idCat
	qrySql = qrySql & " GROUP BY "
	qrySql = qrySql & " Id_Canal, "
	qrySql = qrySql & " Canal "
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " Canal "
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCanal = Server.CreateObject("ADODB.recordSet")
	rsCanal.Open qrySql, conexionRS
	'
	if not rsCanal.EOF then
		arrCanal = rsCanal.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsCanal.Close : Set rsCanal = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrCanal) then

		For i = 0 to ubound(arrCanal, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCanal(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCanal(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 5) THEN
	'
	' Fill combo Fabricante
	'			
	Dim rsFabricante, arrFabricante	
	'
	' Buscar Datos de todas las Fabricantes
	'
	qrySql = vbNullstring	
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " Id_Fabricante as id, "
	qrySql = qrySql & " Fabricante as nombre "
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_DataProcSem "
	qrySql = qrySql & " WHERE "
	qrySql = qrySql & " Id_Categoria = " & idCat
	qrySql = qrySql & " GROUP BY "
	qrySql = qrySql & " Id_Fabricante, "
	qrySql = qrySql & " Fabricante "
	qrySql = qrySql & " HAVING "
	qrySql = qrySql & " Id_Fabricante <> 0 "
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " Fabricante "
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsFabricante = Server.CreateObject("ADODB.recordSet")
	rsFabricante.Open qrySql, conexionRS
	'
	if not rsFabricante.EOF then
		arrFabricante = rsFabricante.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsFabricante.Close : Set rsFabricante = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrFabricante) then

		For i = 0 to ubound(arrFabricante, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrFabricante(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrFabricante(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 6) THEN
	'
	'Fill combo Marca
	'			
	Dim rsMarca, arrMarca
	'
	' Buscar Datos de todas las Canales
	'
	if idCat >= 127 and idCat <= 145 then
		qrySql = vbNullstring	
		qrySql = qrySql & " SELECT "
		qrySql = qrySql & " Id_Marca as id, "
		qrySql = qrySql & " Trim(Marca)+'('+Trim(Fabricante)+')' as nombre "
		qrySql = qrySql & " FROM "
		qrySql = qrySql & " RS_DataProcSem "
		qrySql = qrySql & " WHERE "
		qrySql = qrySql & " Id_Fabricante <> 0 AND Id_Categoria = " & idCat
		qrySql = qrySql & " GROUP BY "
		qrySql = qrySql & " Id_Marca, "
		qrySql = qrySql & " Trim(Marca)+'('+Trim(Fabricante)+')'"
		qrySql = qrySql & " HAVING "
		qrySql = qrySql & " Id_Marca <> 0 "
		qrySql = qrySql & " ORDER BY "
		qrySql = qrySql & " Trim(Marca) + '('+Trim(Fabricante)+')'"
	else 
		qrySql = vbNullstring	
		qrySql = qrySql & " SELECT "
		qrySql = qrySql & " Id_Marca as id, "
		qrySql = qrySql & " Marca as nombre"
		qrySql = qrySql & " FROM "
		qrySql = qrySql & " RS_DataProcSem "
		qrySql = qrySql & " WHERE "
		qrySql = qrySql & " Id_Categoria = " & idCat
		qrySql = qrySql & " GROUP BY "
		qrySql = qrySql & " Id_Marca, "
		qrySql = qrySql & " Marca "
		qrySql = qrySql & " HAVING "
		qrySql = qrySql & " Id_Marca <> 0 "
		qrySql = qrySql & " ORDER BY "
		qrySql = qrySql & " Marca "
	end if
	'		
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMarca = Server.CreateObject("ADODB.recordSet")
	rsMarca.Open qrySql, conexionRS
	'
	if not rsMarca.EOF then
		arrMarca = rsMarca.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsMarca.Close : Set rsMarca = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrMarca) then

		For i = 0 to ubound(arrMarca, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrMarca(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrMarca(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 7) THEN
	'
	'Fill combo Segmento
	'			
	Dim rsSegmento, arrSegmento
	'
	' Buscar Datos de todas las Canales
	'
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " Id_Segmento as id, "
	qrySql = qrySql & " Segmento as nombre"
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_DataProcSem "
	qrySql = qrySql & " WHERE "
	qrySql = qrySql & " Id_Categoria = " & idCat
	qrySql = qrySql & " GROUP BY "
	qrySql = qrySql & " Id_Segmento, "
	qrySql = qrySql & " Segmento "
	qrySql = qrySql & " HAVING "
	qrySql = qrySql & " Id_Segmento <> 0 "
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " Segmento "
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSegmento = Server.CreateObject("ADODB.recordSet")
	rsSegmento.Open qrySql, conexionRS
	'
	if not rsSegmento.EOF then
		arrSegmento = rsSegmento.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsSegmento.Close : Set rsSegmento = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrSegmento) then

		For i = 0 to ubound(arrSegmento, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrSegmento(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrSegmento(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'		
	
ELSEIF (Cint(opcion) = 8) THEN
	'
	'Fill combo Tamaño
	'			
	Dim rsTamano, arrTamano
	'
	' Buscar Datos de todas las Tamanos
	'
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " Id_Tamano as id, "
	qrySql = qrySql & " CONVERT(DECIMAL(10,0),Tamano) as nombre"
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_DataProcSem "
	qrySql = qrySql & " WHERE "
	qrySql = qrySql & " Id_Categoria =  " & idCat
	qrySql = qrySql & " GROUP BY "
	qrySql = qrySql & " Id_Tamano, "
	qrySql = qrySql & " Tamano "
	qrySql = qrySql & " HAVING "
	qrySql = qrySql & " Id_Tamano <> 0 "
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " CONVERT(DECIMAL(10,0),Tamano) "
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsTamano = Server.CreateObject("ADODB.recordSet")
	rsTamano.Open qrySql, conexionRS
	'
	if not rsTamano.EOF then
		arrTamano = rsTamano.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsTamano.Close : Set rsTamano = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbNullstring

	if IsArray(arrTamano) then

		For i = 0 to ubound(arrTamano, 2)
			'
			'value=Split(arrTamano(1,i),".")			
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrTamano(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrTamano(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'
ELSEIF (Cint(opcion) = 9) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'
	qrySql = vbNullstring	
	qrySql = qrySql & " SELECT"
	qrySql = qrySql & " RS_DataProcSem.CodigoBarra as id,"	
	qrySql = qrySql & " TRIM(RS_DataProcSem.Descripcion) as nombre"
	qrySql = qrySql & " FROM"
	qrySql = qrySql & " RS_DataProcSem INNER JOIN PH_CB_Fabricante ON RS_DataProcSem.Id_Fabricante = PH_CB_Fabricante.id_Fabricante"
	qrySql = qrySql & " WHERE"
	qrySql = qrySql & " RS_DataProcSem.Id_Categoria = " & idCat
	qrySql = qrySql & " AND"
	qrySql = qrySql & " PH_CB_Fabricante.Ind_MarcaPropia = 0"
	qrySql = qrySql & " GROUP BY"
	qrySql = qrySql & " RS_DataProcSem.CodigoBarra,"
	qrySql = qrySql & " RS_DataProcSem.Descripcion"
	qrySql = qrySql & " HAVING"	
	qrySql = qrySql & " ( RS_DataProcSem.CodigoBarra IS NOT NULL AND RS_DataProcSem.CodigoBarra <> '' )"
	qrySql = qrySql & " AND"
	qrySql = qrySql & " ( RS_DataProcSem.Descripcion IS NOT NULL AND RS_DataProcSem.Descripcion <> '' )"	
	qrySql = qrySql & " ORDER BY"
	qrySql = qrySql & " nombre"	
	'	
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsProducto = Server.CreateObject("ADODB.recordSet")
	rsProducto.Open qrySql, conexionRS
	'
	if not rsProducto.EOF then
		arrProducto = rsProducto.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	'Response.end
	rsProducto.Close : Set rsProducto = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrProducto) then

		For i = 0 to ubound(arrProducto, 2)
			'
			sTabla     = chr(123) &  chr(34) & "id" 	& chr(34) & ":" & chr(34) & arrProducto(0,i) & chr(34) & chr(44)
			sTabla     = sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & RemoverSaltodeLinea(arrProducto(1,i)) &  " - "  & RemoverSaltodeLinea(arrProducto(0,i)) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34) & ":" & chr(34)  & "0" 		 & chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'	
ELSEIF (Cint(opcion) = 10) THEN
	'
	'Fill combo Indicadores
	'			
	Dim rsIndicadores, arrIndicadores
	'
	' Buscar Datos de todas las Indicadores
	'	
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT "	
	qrySql = qrySql & " Id_Indicador as id, "
	qrySql = qrySql & " Abreviatura as nombre"
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " RS_Indicadores "
	qrySql = qrySql & " WHERE "	
	if idCliente = 1 then 
		qrySql = qrySql & " Ind_Atenas = 1 " 
	else
		qrySql = qrySql & " Ind_Men = 1 " 
	end if
	'
 	if (idCat > 126 and idCat < 146) or (idCat = 41 or idCat = 18 or idCat = 54) then
		qrySql = qrySql & " AND ( Id_Indicador <> 3 and Id_Indicador <> 15 and Id_Indicador <> 9 ) "
	end if
	'
	qrySql = qrySql & " AND Ind_Activo = 1 " 
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " Id_Indicador "		
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsIndicadores = Server.CreateObject("ADODB.recordSet")
	rsIndicadores.Open qrySql, conexionRS
	'
	if not rsIndicadores.EOF then
		arrIndicadores = rsIndicadores.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsIndicadores.Close : Set rsIndicadores = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrIndicadores) then

		For i = 0 to ubound(arrIndicadores, 2)
			'
			sTabla     =   chr(123)&  chr(34) & "id" 	& chr(34) & ":" & chr(34) & arrIndicadores(0,i) & chr(34) & chr(44)
			sTabla     =   sTabla &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrIndicadores(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		 & chr(34) & ":" & chr(34) & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34) & ":" & chr(34) & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'
ELSEIF (Cint(opcion) = 11) THEN
	'
	'Fill combo Semanas
	'			
	Dim rsSemanas, arrSemanas, iSemanaDes, iSemanaHas, rsSemanario
	'	
	if Cint(idCliente=1) then
		'atenas
		qrySql = vbNullstring
		qrySql = qrySql & " SELECT Min(RS_DataProcSem.Id_Semana) AS desde, Max(RS_DataProcSem.Id_Semana) AS hasta, RS_DataProcSem.Id_Categoria FROM RS_DataProcSem"
		qrySql = qrySql & " GROUP BY RS_DataProcSem.Id_Categoria HAVING (((RS_DataProcSem.Id_Categoria)=" & idCat & "));"
		'		
	else
		qrySql = vbNullstring		
		qrySql = qrySql & " SELECT ss_ClienteCategoria.Id_PeriodoDesde AS desde, ss_ClienteCategoria.Id_PeriodoPub AS hasta FROM ss_ClienteCategoria"
		qrySql = qrySql & " WHERE"
		qrySql = qrySql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		qrySql = qrySql & " AND"
		qrySql = qrySql & " ss_ClienteCategoria.Ind_Mensual = 1"
		qrySql = qrySql & " AND"
		qrySql = qrySql & " ss_ClienteCategoria.Id_Categoria = " & idCat
	end if
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSemanario = Server.CreateObject("ADODB.recordSet")
	rsSemanario.Open qrySql, conexionRS
			
	if not (rsSemanario.EOF and rsSemanario.BOF) then		
		iSemanaDes = rsSemanario("desde").value 'rsSemanario(0)
		iSemanaHas = rsSemanario("hasta").value 'rsSemanario(1)
	else
		iSemanaDes = 0
		iSemanaHas = 0
	end if		
	'
	rsSemanario.Close : Set rsSemanario = Nothing
	'
	' Buscar Datos de todas las Semanas
	'	
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT "
	qrySql = qrySql & " IdSemana as id, "
	qrySql = qrySql & " Semana as nombre "
	qrySql = qrySql & " FROM "
	qrySql = qrySql & " ss_Semana "
	if( iSemanaDes <> 0 and iSemanaHas <> 0 ) then
		qrySql = qrySql & " WHERE "
		qrySql = qrySql & " IdSemana >= " & iSemanaDes
		qrySql = qrySql & " And IdSemana <= " & iSemanaHas
	end if	
	qrySql = qrySql & " ORDER BY "
	qrySql = qrySql & " IdSemana DESC "
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSemanas = Server.CreateObject("ADODB.recordSet")
	rsSemanas.Open qrySql, conexionRS
	'
	if not rsSemanas.EOF then
		arrSemanas = rsSemanas.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsSemanas.Close : Set rsSemanas = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrSemanas) then

		For i = 0 to ubound(arrSemanas, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrSemanas(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrSemanas(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		'
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 12) THEN	
 	'
	' Verificar cliente contrato el servicio				
	'
	Dim rsCliente, arrCliente
	'
	if (CInt(idCliente) = 1) then
	
		Response.Write CInt(idCliente)
	
	else
		qrySql = vbNullstring		
		qrySql = qrySql & " SELECT COUNT(Id_Cliente) as total FROM dbo.ss_ClienteCategoria"
		qrySql = qrySql & " WHERE"
		qrySql = qrySql & " dbo.ss_ClienteCategoria.Id_Cliente = " & idCliente
		qrySql = qrySql & " AND"
		qrySql = qrySql & " dbo.ss_ClienteCategoria.Ind_Mensual = 1"
		'
		Response.Write qrySql & "<BR><BR>"
		'Response.end
		'
		Set rsCliente = Server.CreateObject("ADODB.recordSet")
		rsCliente.Open qrySql, conexionRS
		'
		if not (rsCliente.EOF and rsCliente.BOF) then
			Response.Write rsCliente(0)
		else			
			Response.Write 0			
		end if		
		'
		rsCliente.Close : Set rsCliente = Nothing
		'
		conexionRS.Close : Set conexionRS = Nothing		
	end if	'

ELSEIF (Cint(opcion) = 13) THEN
	'
	'Fill combo Meses
	'			
	Dim rsMeses, arrMeses
	'
	' Buscar Datos de todas las Meses
	'
	Dim iSemDes, iSemHas, rsMensual
	'	
	if Cint(idCliente=1) then
		'Atenas
		qrySql = vbNullstring
		qrySql = qrySql & " SELECT Min(RS_DataProcSem.Id_Semana) AS desde, Max(RS_DataProcSem.Id_Semana) AS hasta, RS_DataProcSem.Id_Categoria FROM RS_DataProcSem"
		qrySql = qrySql & " GROUP BY RS_DataProcSem.Id_Categoria HAVING RS_DataProcSem.Id_Categoria=" & idCat
		'		
	else
		qrySql = vbNullstring		
		qrySql = qrySql & " SELECT ss_ClienteCategoria.Id_PeriodoDesde AS desde, ss_ClienteCategoria.Id_PeriodoPub AS hasta FROM ss_ClienteCategoria"
		qrySql = qrySql & " WHERE"
		qrySql = qrySql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		qrySql = qrySql & " AND"
		qrySql = qrySql & " ss_ClienteCategoria.Ind_Mensual = 1"
		qrySql = qrySql & " AND"
		qrySql = qrySql & " ss_ClienteCategoria.Id_Categoria = " & idCat
	end if
	'
	' Response.Write qrySql & "<BR><BR>"
	' Response.end
	'
	Set rsMensual = Server.CreateObject("ADODB.recordSet")
	rsMensual.Open qrySql, conexionRS
			
	if not (rsMensual.EOF and rsMensual.BOF) then		
		iSemDes = rsMensual("desde").value 'rsSemanario(0)
		iSemHas = rsMensual("hasta").value 'rsSemanario(1)
	else
		iSemDes = 0
		iSemHas = 0
	end if		
	'
	rsMensual.Close : Set rsMensual = Nothing
	'		
	qrySql = vbNullstring
	qrySql = qrySql & " SELECT ss_Periodo.IdPeriodo as id, ss_Periodo.Periodo as nombre"
	qrySql = qrySql & " FROM ss_Periodo INNER JOIN ss_Semana ON ss_Periodo.IdPeriodo = ss_Semana.Id_Periodo"
	qrySql = qrySql & " WHERE ss_Semana.IdSemana >= " & iSemDes & " AND ss_Semana.IdSemana<= " & iSemHas 
	qrySql = qrySql & " GROUP BY ss_Periodo.IdPeriodo, ss_Periodo.Periodo"
	qrySql = qrySql & " ORDER BY ss_Periodo.IdPeriodo DESC;"	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMeses = Server.CreateObject("ADODB.recordSet")
	rsMeses.Open qrySql, conexionRS
	'
	if not rsMeses.EOF then
		arrMeses = rsMeses.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsMeses.Close : Set rsMeses = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrMeses) then

		For i = 0 to ubound(arrMeses, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrMeses(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & RemoverSaltodeLinea(arrMeses(1,i)) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34) & ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla   = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData = chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing	
	'
ELSEIF (Cint(opcion) = 14) THEN
'
	'Find Categoria es Medicina
	'				
	Dim rsMedicina
	'	
	qrySql = vbNullstring
	qrySql = " SELECT PH_CB_Categoria.Ind_Medicina as Medicina FROM dbo.PH_CB_Categoria WHERE PH_CB_Categoria.id_Categoria = " & idCat
	'		
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMedicina = Server.CreateObject("ADODB.recordSet")
	rsMedicina.Open qrySql, conexionRS
			
	if not (rsMedicina.EOF and rsMedicina.BOF) then	
		Response.Write rsMedicina("Medicina").value							
	end if		
	'
	rsMedicina.Close : Set rsMedicina = Nothing	
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing	
	'
ELSE
	' de lo Contrario
	Response.Write "Ups!, Algo Salio Mal..!"
END IF
'
%>