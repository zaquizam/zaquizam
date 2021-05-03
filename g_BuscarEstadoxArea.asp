<%@language=vbscript%>
<!--#include file="../Conexion.asp"-->
<%
	Dim rsEstado, idArea
	'
	idArea= Request.Form("idarea")
	'
	' Buscar Los Estados asociados al Area
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_GAreaEstado.Id_AreaEstado AS id,"
	QrySql = QrySql & " ss_Estado.Estado AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_GAreaEstado"
	QrySql = QrySql & " INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_GAreaEstado.Ind_Activo = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_GAreaEstado.Id_Area = " & CInt(idArea)
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " nombre ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsEstado = Server.CreateObject("ADODB.recordset")
	rsEstado.Open QrySql,conexion
	'
	if not rsEstado.EOF then
    	arrEstado = rsEstado.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla=""

    if IsArray(arrEstado) then

        For i = 0 to ubound(arrEstado, 2)
            '
            sTabla     =  chr(123)&  chr(34) & "id" 	& chr(34) & ":" & chr(34) & arrEstado(0,i)  & chr(34) & chr(44)
            sTabla     =  sTabla &  chr(34)  & "nombre" & chr(34) & ":" & chr(34) & arrEstado(1,i)  & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "nombre"         & chr(34)& ":" & chr(34) & "No Aplica" 	& chr(34) & chr(125) & chr(44)
        ''
        sTablaJson = sTablaJson & sTabla
        sTabla=""

    end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsEstado.Close
	Set rsEstado = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>