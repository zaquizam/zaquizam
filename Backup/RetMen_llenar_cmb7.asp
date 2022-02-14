<!--#include file="conexionRS.asp" -->
<%
'
' RetMen_llenar_cmb7.asp - 15jul21 - 22nov21
'
' Cambio en combo Segmento - 
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

Dim oPcion, qrySql, idCat, idCliente
'
' oPcion  = Cint(Request.Querystring("oPcion"))
' idQuery = Cint(Request.Querystring("id"))
'
oPcion = Request.Querystring("opcion")
idCat  = Request.Querystring("idCat")
idArea = Request.Querystring("idArea")
idZona = Request.Querystring("idZona")
idCanal = Request.Querystring("idCanal")
idFab   = Request.Querystring("idFab")
idMar   = Request.Querystring("idMar")
idSeg  = Request.Querystring("idSeg")
idCliente = Request.Querystring("idCli")
'
IF (Cint(oPcion) = 7) THEN
	'
	'ReFill combo Tamaño
	'			
	Dim rsTamano, arrTamano		
	'
	' Buscar Datos de todas las Tamano
	'
	qrySql = vbnullstring	
	qrySql = qrySql & " SELECT DISTINCT Id_Tamano as id, Tamano as nombre FROM RS_DataProcSem WHERE"
	qrySql = qrySql & " Id_Categoria = " & idCat
	'
	if Len(idArea) <> 0 then 
		qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
	end if
	if Len(idZona) <> 0 then 
		qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
	end if
	if Len(idCanal) <> 0 then 
		qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
	end if
	if Len(idFab) <> 0 then 
		qrySql = qrySql & " AND Id_Fabricante in (" & idFab & ")"
	end if
	if Len(idMar) <> 0 then 
		qrySql = qrySql & " AND Id_Marca in (" & idMar & ")"
	end if
	if Len(idSeg) <> 0 then 
		qrySql = qrySql & " AND Id_Segmento in (" & idSeg & ")"
	end if	
	qrySql = qrySql & " AND Id_Tamano > 0 ORDER BY Tamano"	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set rsTamano = Server.CreateObject("ADODB.recordset")
	rsTamano.Open qrySql, conexionRS
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
ELSEIF (Cint(oPcion) = 8) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'	
	'17nov
	qrySql = vbnullstring	
	qrySql = qrySql & " SELECT"
	qrySql = qrySql & " RS_DataProcSem.CodigoBarra as id,"	
	qrySql = qrySql & " TRIM(RS_DataProcSem.Descripcion) as nombre"
	qrySql = qrySql & " FROM"
	qrySql = qrySql & " RS_DataProcSem INNER JOIN PH_CB_Fabricante ON RS_DataProcSem.Id_Fabricante = PH_CB_Fabricante.id_Fabricante"
	qrySql = qrySql & " WHERE"
	qrySql = qrySql & " RS_DataProcSem.Id_Categoria = " & idCat
	'
	if Len(idArea) <> 0 then 
		qrySql = qrySql & " AND Id_Area in (" & idArea & ")"
	end if	
	if Len(idZona) <> 0 then 
		qrySql = qrySql & " AND Id_Zona in (" & idZona & ")"
	end if		
	if Len(idCanal) <> 0 then 
		qrySql = qrySql & " AND Id_Canal in (" & idCanal & ")"
	end if	
	if Len(idFab) <> 0 then 
		qrySql = qrySql & " AND PH_CB_Fabricante.id_Fabricante in (" & idFab & ")"
	end if
	if Len(idMar) <> 0 then 
		qrySql = qrySql & " AND Id_Marca in (" & idMar & ")"
	end if
	if Len(idSeg) <> 0 then 
		qrySql = qrySql & " AND Id_Segmento in (" & idSeg & ")"
	end if
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
	Set rsProducto = Server.CreateObject("ADODB.recordset")
	rsProducto.Open qrySql, conexionRS
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