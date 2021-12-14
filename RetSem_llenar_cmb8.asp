<!--#include file="conexionRS.asp" -->
<%
'
' RetSem_llenar_cmb8.asp - 15jul21 - 17nov21
'
' Cambio en combo Tamaño - 
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
if conexionRS.errors.count <> 0 Then
  Response.Write ("No hay conexionRS con la BD...!")
  Response.End
end if

Dim opcion, QrySql, idCat, idCliente, idMar
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
'
opcion  = Request.Querystring("opcion")
idCat   = Request.Querystring("idCat")
idArea  = Request.Querystring("idArea")
idZona  = Request.Querystring("idZona")
idCanal = Request.Querystring("idCanal")
idFab   = Request.Querystring("idFab")
idMar   = Request.Querystring("idMar")
idSeg   = Request.Querystring("idSeg")
idTam   = Request.Querystring("idTam")
idCliente = Request.Querystring("idCli")
'
IF (Cint(opcion) = 8) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'
	'17nov
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " RS_DataProcSem.CodigoBarra as id,"	
	QrySql = QrySql & " TRIM(RS_DataProcSem.Descripcion) as nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " RS_DataProcSem INNER JOIN PH_CB_Fabricante ON RS_DataProcSem.Id_Fabricante = PH_CB_Fabricante.id_Fabricante"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " RS_DataProcSem.Id_Categoria = " & idCat
	if len(idArea)<>0 then 
		QrySql = QrySql & " AND Id_Area in (" & idArea & ")"
	end if	
	if len(idZona)<>0 then 
		QrySql = QrySql & " AND Id_Zona in (" & idZona & ")"
	end if		
	if len(idCanal)<>0 then 
		QrySql = QrySql & " AND Id_Canal in (" & idCanal & ")"
	end if	
	if len(idFab)<>0 then 
		QrySql = QrySql & " AND PH_CB_Fabricante.id_Fabricante in (" & idFab & ")"
	end if
	if len(idMar)<>0 then 
		QrySql = QrySql & " AND Id_Marca in (" & idMar & ")"
	end if
	if Len(idSeg) <> 0 then 
		qRySql = qRySql & " AND Id_Segmento in (" & idSeg & ")"
	end if
	if len(idTam)<>0 then 
		QrySql = QrySql & " AND Id_Tamano in (" & idTam & ")"
	end if
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_CB_Fabricante.Ind_MarcaPropia = 0"
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " RS_DataProcSem.CodigoBarra,"
	QrySql = QrySql & " RS_DataProcSem.Descripcion"
	QrySql = QrySql & " HAVING"	
	QrySql = QrySql & " ( RS_DataProcSem.CodigoBarra IS NOT NULL AND RS_DataProcSem.CodigoBarra <> '' )"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " ( RS_DataProcSem.Descripcion IS NOT NULL AND RS_DataProcSem.Descripcion <> '' )"	
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " nombre"	
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
	Response.write "error RetSem_llenar_cmb8.asp"
	'
END IF
'
%>