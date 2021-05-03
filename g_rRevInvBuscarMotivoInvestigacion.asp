<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_rRevInvBuscarMotivoInvestigacion.asp // 14ene21 - 25ene21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'	
	' Buscar la Fecha y dia del Consumo
	'
	idConsumo =	Request.Querystring("id_consumo")
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_InvestigacionItems.InvestigacionItems,"
	QrySql = QrySql & " PH_Consumo_Investigar_Detalle.Observaciones_enviadas"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo_Investigar_Detalle"
	QrySql = QrySql & " INNER JOIN PH_InvestigacionItems ON PH_Consumo_Investigar_Detalle.Id_items_investigacion = PH_InvestigacionItems.Id_InvestigacionItems"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo_Investigar_Detalle.Caso_Cerrado='0'"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo_Investigar_Detalle.Id_Consumo = " & idConsumo	
	'
	Set rsMotivo = Server.CreateObject("ADODB.recordset")
	rsMotivo.Open QrySql, conexion
	''		
	if not rsMotivo.EOF then		
		Resultado = trim(rsMotivo(0)) & "-" & trim(rsMotivo(1))
		Response.Write	Resultado
	 else
		 Response.Write False
	end if
	'
		
	' ''
	' if not rsMotivo.EOF then
    	' arrMotivo = rsMotivo.GetRows()  ' Convert recordset to 2D Array
	' end if
	' '
	' 'Response.ContentType = "application/json"
	' '
	' ' Crear Archivo Array Json
	' '
	' sTabla=""

    ' if IsArray(arrMotivo) then

        ' For i = 0 to UBound(arrMotivo, 2)
            ' ''			
			' sTabla     =  chr(123) &  chr(34) & "motivo"   & chr(34) & ":" & chr(34) & arrMotivo(0,i) & chr(34) & chr(44)
			' sTabla     =  sTabla   &  chr(34) & "observa"  & chr(34) & ":" & chr(34) & arrMotivo(1,i) & chr(34) & chr(125) & chr(44)
            ' sTablaJson =  sTablaJson & sTabla
            ' sTabla=""
            ' '
        ' next

    ' else
        ' 'Eof()
        ' sTabla    =   chr(123)&  chr(34) & "motivo"  & chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        ' sTabla    =   sTabla &   chr(34) & "observa" & chr(34)& ":" & chr(34) & "No Aplica" 	& chr(34) & chr(125) & chr(44)
        ' '
        ' sTablaJson = sTablaJson & sTabla
        ' sTabla=""

    ' end if
	' ''
	' sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	' JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
	' 'JsonData	= 	 sTabla  '& chr(125)
	' Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsMotivo.Close
	Set rsMotivo = Nothing
	'
	conexion.close
	set conexion = nothing
	'
	''
%>