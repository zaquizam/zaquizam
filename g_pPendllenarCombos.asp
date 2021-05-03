<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'Response.AddHeader "Content-Type","application/json;charset=utf-8"
	'
	' g_pPendllenarCombos - 11mar21
	'
	Session.lcid		= 1034
	Response.CodePage 	= 65001
	Response.CharSet 	= "utf-8"	
	'
	iOpcion  = Request.QueryString("idOpcion")	
	iBuscar1 = Request.QueryString("idBusqueda1")
	iBuscar2 = Request.QueryString("idBusqueda2")
	'
	IF CInt(iOpcion) = 1 THEN
		' Llenar Combo categoria
		'
		Dim rsCategorias, arrCategorias
		'	
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_categoria, categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_Categoria"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " Categoria"
		'
		Set rsCategorias = Server.CreateObject("ADODB.recordset")
		rsCategorias.Open QrySql, conexion
		'
		if not rsCategorias.EOF then
			arrCategorias = rsCategorias.GetRows()  ' Convert recordset to 2D Array
		end if
		'		
		sTabla=vbnullstring

		if IsArray(arrCategorias) then

			For i = 0 to ubound(arrCategorias, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrCategorias(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrCategorias(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla=vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsCategorias.Close : Set rsCategorias = Nothing
		'		
	ELSEIF CInt(iOpcion) = 2 THEN
		'
		' Llenar Combo Fabricante
		'
		Dim rsFabricante, arrFabricante
		'			
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_fabricante, fabricante"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_Fabricante"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " AND id_categoria=" & CInt(iBuscar1)
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " fabricante"
		'		
		Set rsFabricante = Server.CreateObject("ADODB.recordset")
		rsFabricante.Open QrySql, conexion
		'
		if not rsFabricante.EOF then
			arrFabricante = rsFabricante.GetRows()  ' Convert recordset to 2D Array
		end if
		'
		sTabla=vbnullstring

		if IsArray(arrFabricante) then

			For i = 0 to ubound(arrFabricante, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrFabricante(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrFabricante(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla=vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsFabricante.Close : Set rsFabricante = Nothing
		'
	ELSEIF CInt(iOpcion) = 3 THEN
		'
		' Llenar Combo Marcas
		'
		Dim rsMarca, arrMarca
		'			
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_Marca, Marca"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_Marca"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " AND id_categoria=" & CInt(iBuscar1)
		QrySql = QrySql & " AND id_fabricante=" & CInt(iBuscar2)
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " Marca"
		'		
		Set rsMarca = Server.CreateObject("ADODB.recordset")
		rsMarca.Open QrySql, conexion
		'
		if not rsMarca.EOF then
			arrMarca = rsMarca.GetRows()  ' Convert recordset to 2D Array
		end if
		'
		sTabla=vbnullstring

		if IsArray(arrMarca) then

			For i = 0 to ubound(arrMarca, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrMarca(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrMarca(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla=vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsMarca.Close : Set rsMarca = Nothing
		'
	ELSEIF CInt(iOpcion) = 4 THEN
		'
		' Llenar Combo Segmento
		'
		Dim rsSegmento, arrSegmento
		'			
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_Segmento, Segmento"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_Segmento"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " AND id_categoria=" & CInt(iBuscar1)		
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " Segmento"
		'		
		Set rsSegmento = Server.CreateObject("ADODB.recordset")
		rsSegmento.Open QrySql, conexion
		'
		if not rsSegmento.EOF then
			arrSegmento = rsSegmento.GetRows()  ' Convert recordset to 2D Array
		end if
		'
		sTabla=vbnullstring

		if IsArray(arrSegmento) then

			For i = 0 to ubound(arrSegmento, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrSegmento(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrSegmento(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla = vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsSegmento.Close : Set rsSegmento = Nothing
		'
	ELSEIF CInt(iOpcion) = 5 THEN
		'
		' Llenar Combo Tamaño
		'
		Dim rsTamano, arrTamano
		'			
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_Tamano, Tamano"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_Tamano"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " AND id_categoria=" & CInt(iBuscar1)		
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " Tamano"
		'		
		Set rsTamano = Server.CreateObject("ADODB.recordset")
		rsTamano.Open QrySql, conexion
		'
		if not rsTamano.EOF then
			arrTamano = rsTamano.GetRows()  ' Convert recordset to 2D Array
		end if
		'
		sTabla=vbnullstring

		if IsArray(arrTamano) then

			For i = 0 to ubound(arrTamano, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrTamano(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrTamano(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla = vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsTamano.Close : Set rsTamano = Nothing
		'
	ELSEIF CInt(iOpcion) = 6 THEN
		'
		' Llenar Combo Rango / TamanoRango
		'
		Dim rsRango, arrRango
		'			
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_TamanoRango, TamanoRango"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_TamanoRango"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " AND id_categoria=" & CInt(iBuscar1)		
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " TamanoRango"
		'		
		Set rsRango = Server.CreateObject("ADODB.recordset")
		rsRango.Open QrySql, conexion
		'
		if not rsRango.EOF then
			arrRango = rsRango.GetRows()  ' Convert recordset to 2D Array
		end if
		'
		sTabla=vbnullstring

		if IsArray(arrRango) then

			For i = 0 to ubound(arrRango, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrRango(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrRango(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla = vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsRango.Close :	Set rsRango = Nothing
		'
	ELSEIF CInt(iOpcion) = 7 THEN
		'
		' Llenar Combo Unidad - Medida
		'
		Dim rsUnidadMedida, arrUnidadMedida
		'			
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT id_UnidadMedida, UnidadMedida"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_CB_UnidadMedida"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ind_activo=1"
		QrySql = QrySql & " AND id_categoria=" & CInt(iBuscar1)		
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " UnidadMedida"
		'		
		Set rsUnidadMedida = Server.CreateObject("ADODB.recordset")
		rsUnidadMedida.Open QrySql, conexion
		'
		if not rsUnidadMedida.EOF then
			arrUnidadMedida = rsUnidadMedida.GetRows()  ' Convert recordset to 2D Array
		end if
		'
		sTabla=vbnullstring

		if IsArray(arrUnidadMedida) then

			For i = 0 to ubound(arrUnidadMedida, 2)
				'
				sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrUnidadMedida(0,i)) & chr(44)
				sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrUnidadMedida(1,i)  & chr(34) & chr(125) & chr(44)
				sTablaJson =  sTablaJson & sTabla
				sTabla = vbnullstring
				'
			next

		else
			'Eof()
			sTabla    =   chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & "0" 			      & chr(34) & chr(44)
			sTabla    =   sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & "No hay Registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring

		end if
		''
		sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
		Response.Write(JsonData)
		'
		' Cerrar conexiones
		'
		rsUnidadMedida.Close :	Set rsUnidadMedida = Nothing
		'
	ELSE
		'
	END IF
	'
	conexion.close
	set conexion = nothing
	'
%>