<!--#include file="conexionRS.asp" -->
<%
'
' RetSem_llenar_cmb8.asp - 15jul21 - 16jul21
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
' if idArea ="" or len(idArea)=0  then idArea=0
' if idZona ="" or len(idZona)=0  then idZona=0
' if idCanal="" or len(idCanal)=0 then idCanal=0
' if idFab  ="" or len(idFab)=0   then idFab=0
' if idMar  ="" or len(idMar)=0   then idMar=0
' if idSeg  ="" or len(idSeg)=0   then idSeg=0
' if idTam  ="" or len(idTam)=0   then idTam=0
'
IF (Cint(opcion) = 8) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT DISTINCT CodigoBarra as id, Descripcion as nombre FROM RS_DataProcSem WHERE"
	QrySql = QrySql & " Id_Categoria= " &  idCat
	'
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
		QrySql = QrySql & " AND Id_Fabricante in (" & idFab & ")"
	end if
	if len(idMar)<>0 then 
		QrySql = QrySql & " AND Id_Marca in (" & idMar & ")"
	end if
	if len(idSeg)<>0 then 
		QrySql = QrySql & " AND Id_Segmento in (" & idSeg & ")"
	end if
	if len(idTam)<>0 then 
		QrySql = QrySql & " AND Id_Tamano in (" & idTam & ")"
	end if
	
	QrySql = QrySql & " AND CodigoBarra IS NOT NULL AND CodigoBarra <> '' AND Descripcion IS NOT NULL AND Descripcion <> '' ORDER BY Descripcion"
	
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
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrProducto(1,i) & chr(34) & chr(125) &chr(44)
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