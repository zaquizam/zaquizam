<!--#include file="conexionRS.asp" -->
<%
'
'RetSem_llenar_cmb1.asp - 12jul21 - 
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
if conexionRS.errors.count <> 0 Then
  Response.Write ("No hay conexionRS con la BD...!")
  Response.End
end if

Dim opcion, QrySql, idCat, idCliente
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
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
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT Id_Categoria as id,"
	QrySql = QrySql & " Categoria as nombre"
	QrySql = QrySql & " FROM RS_DataProcSem "
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Categoria, " 
	QrySql = QrySql & " Categoria " 
	
	IF (Cint(idCliente) = 17) THEN
		QrySql = QrySql & " HAVING "
		QrySql = QrySql & " Id_Categoria In (7,8,14,5,6,34,24,21,13,2)"
	END IF
	
	IF (Cint(idCliente) = 30) THEN
		QrySql = QrySql & " HAVING "
		QrySql = QrySql & " Id_Categoria In (33)"
	END IF
	IF (Cint(idCliente) = 7) THEN
		QrySql = QrySql & " HAVING "
		QrySql = QrySql & " Id_Categoria In (21,55,24,23,20,25)"
	END IF
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Categoria "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordset")
	rsCategoria.Open QrySql, conexionRS
	'
	if not rsCategoria.EOF then
		arrCategoria = rsCategoria.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsCategoria.Close : Set rsCategoria = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCategoria) then

		For i = 0 to ubound(arrCategoria, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCategoria(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCategoria(1,i) & chr(34) & chr(125) &chr(44)
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
ELSEIF (Cint(opcion) = 2) THEN
	'
	'Fill combo Area
	'			
	Dim rsArea, arrArea
	'
	' Buscar Datos de todas las Areas
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Area as id, "
	QrySql = QrySql & " Area as mombre "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Area, "
	QrySql = QrySql & " Area "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Id_Area "	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsArea = Server.CreateObject("ADODB.recordset")
	rsArea.Open QrySql, conexionRS
	'
	if not rsArea.EOF then
		arrArea = rsArea.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsArea.Close : Set rsArea = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrArea) then

		For i = 0 to ubound(arrArea, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrArea(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrArea(1,i) & chr(34) & chr(125) &chr(44)
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
ELSEIF (Cint(opcion) = 3) THEN
	'
	'Fill combo Zona
	'			
	Dim rsZona, arrZona
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Zonas
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Zona, "
	QrySql = QrySql & " Zona "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE  "
	QrySql = QrySql & " Id_Categoria= " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Zona, "
	QrySql = QrySql & " Zona "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Zona "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsZona = Server.CreateObject("ADODB.recordset")
	rsZona.Open QrySql, conexionRS
	'
	if not rsZona.EOF then
		arrZona = rsZona.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsZona.Close : Set rsZona = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrZona) then

		For i = 0 to ubound(arrZona, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrZona(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrZona(1,i) & chr(34) & chr(125) &chr(44)
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
	
ELSEIF (Cint(opcion) = 4) THEN
	'
	'Fill combo Canal
	'			
	Dim rsCanal, arrCanal
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Canales
	'
	QrySql = vbnullstring
	
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Canal as id, "
	QrySql = QrySql & " rtrim(Canal) as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Canal, "
	QrySql = QrySql & " Canal "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Canal "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCanal = Server.CreateObject("ADODB.recordset")
	rsCanal.Open QrySql, conexionRS
	'
	if not rsCanal.EOF then
		arrCanal = rsCanal.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsCanal.Close : Set rsCanal = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCanal) then

		For i = 0 to ubound(arrCanal, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCanal(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCanal(1,i) & chr(34) & chr(125) &chr(44)
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
ELSEIF (Cint(opcion) = 5) THEN
	'
	' Fill combo Fabricante
	'			
	Dim rsFabricante, arrFabricante	
	'
	' Buscar Datos de todas las Fabricantes
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Fabricante as id, "
	QrySql = QrySql & " Fabricante as nombre "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Fabricante, "
	QrySql = QrySql & " Fabricante "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Fabricante <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Fabricante "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsFabricante = Server.CreateObject("ADODB.recordset")
	rsFabricante.Open QrySql, conexionRS
	'
	if not rsFabricante.EOF then
		arrFabricante = rsFabricante.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsFabricante.Close : Set rsFabricante = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrFabricante) then

		For i = 0 to ubound(arrFabricante, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrFabricante(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrFabricante(1,i) & chr(34) & chr(125) &chr(44)
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
	'Fill combo Marca
	'			
	Dim rsMarca, arrMarca
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Canales
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Marca as id, "
	QrySql = QrySql & " Marca as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Marca, "
	QrySql = QrySql & " Marca "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Marca <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Marca "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMarca = Server.CreateObject("ADODB.recordset")
	rsMarca.Open QrySql, conexionRS
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
ELSEIF (Cint(opcion) = 7) THEN
	'
	'Fill combo Segmento
	'			
	Dim rsSegmento, arrSegmento
	'
	' Buscar Datos de todas las Canales
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Segmento as id, "
	QrySql = QrySql & " Segmento as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Segmento, "
	QrySql = QrySql & " Segmento "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Segmento <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Segmento "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSegmento = Server.CreateObject("ADODB.recordset")
	rsSegmento.Open QrySql, conexionRS
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
	
ELSEIF (Cint(opcion) = 8) THEN
	'
	'Fill combo Tamaño
	'			
	Dim rsTamano, arrTamano
	'
	' Buscar Datos de todas las Tamanos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Tamano as id, "
	QrySql = QrySql & " CONVERT(DECIMAL(10,2),Tamano) as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria =  " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Tamano, "
	QrySql = QrySql & " Tamano "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Tamano <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " CONVERT(DECIMAL(10,2),Tamano) "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsTamano = Server.CreateObject("ADODB.recordset")
	rsTamano.Open QrySql, conexionRS
	'
	if not rsTamano.EOF then
		arrTamano = rsTamano.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsTamano.Close : Set rsTamano = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
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
ELSEIF (Cint(opcion) = 9) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " CodigoBarra as id, "
	QrySql = QrySql & " Descripcion as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE  "
	QrySql = QrySql & " Id_Categoria= " &  idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " CodigoBarra, "
	QrySql = QrySql & " Descripcion "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " (CodigoBarra Is Not Null "
	QrySql = QrySql & " And CodigoBarra<>'') "
	QrySql = QrySql & " AND (Descripcion Is Not Null "
	QrySql = QrySql & " And Descripcion<>'' ) "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Descripcion "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsProducto = Server.CreateObject("ADODB.recordset")
	rsProducto.Open QrySql, conexionRS
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
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrProducto(0,i) & " " & arrProducto(1,i) & chr(34) & chr(125) &chr(44)
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
ELSEIF (Cint(opcion) = 10) THEN
	'
	'Fill combo Indicadores
	'			
	Dim rsIndicadores, arrIndicadores
	'
	' Buscar Datos de todas las Indicadores
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "	
	QrySql = QrySql & " Id_Indicador as id, "
	QrySql = QrySql & " Abreviatura as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_Indicadores "
	QrySql = QrySql & " WHERE "
	if idCliente = 1 then 
		QrySql = QrySql & " Ind_Atenas = 1 " 
	else
		QrySql = QrySql & " Ind_Sem = 1 " 
	end if
	QrySql = QrySql & " AND Ind_Activo = 1 " 
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Id_Indicador "		
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsIndicadores = Server.CreateObject("ADODB.recordset")
	rsIndicadores.Open QrySql, conexionRS
	'
	if not rsIndicadores.EOF then
		arrIndicadores = rsIndicadores.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsIndicadores.Close : Set rsIndicadores = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbnullstring

	if IsArray(arrIndicadores) then

		For i = 0 to ubound(arrIndicadores, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrIndicadores(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrIndicadores(1,i) & chr(34) & chr(125) &chr(44)
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
ELSEIF (Cint(opcion) = 11) THEN
	'
	'Fill combo Semanas
	'			
	Dim rsSemanas, arrSemanas, iSemanaDes, iSemanaHas
	iSemanaDes = 37
    iSemanaHas = 40
	if idCliente = 1 then 
		iSemanaDes = 24
		iSemanaHas = 40
	end if
	'
	' Buscar Datos de todas las Semanas
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " IdSemana as id, "
	QrySql = QrySql & " Semana as nombre "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " ss_Semana "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " IdSemana >= " & iSemanaDes
	QrySql = QrySql & " And IdSemana <= " & iSemanaHas
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " IdSemana DESC "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSemanas = Server.CreateObject("ADODB.recordset")
	rsSemanas.Open QrySql, conexionRS
	'
	if not rsSemanas.EOF then
		arrSemanas = rsSemanas.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsSemanas.Close : Set rsSemanas = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrSemanas) then

		For i = 0 to ubound(arrSemanas, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrSemanas(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrSemanas(1,i) & chr(34) & chr(125) &chr(44)
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
 ELSEIF (Cint(opcion) = 12) THEN	
	'Verificar cliente contrato el servicio				
	if idCliente = 1 or idCliente = 17 or idCliente = 30 or idCliente  = 7 then
		Response.write true
	else
		Response.write false
	end if	
	'Cerrar conexiones		
	conexionRS.close : set conexionRS = nothing	
ELSE
	' de lo Contrario
	Response.write "error"
END IF
'
%>