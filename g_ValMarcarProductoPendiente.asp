<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	'
	' g_ValMarcarProductoPendiente.asp // 05ene21 - 19feb21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, QrySql, rsSql, idConsumoPrincipal, rsQuery
	'	
	idConsumo	=	Request.Querystring("idConsumo")		
	'
	' Marcar Producto como Pendiente
	'
	QrySql = vbnullstring
	QrySql = QrySql & " UPDATE PH_Consumo_Detalle_Productos"
    QrySql = QrySql & " SET"
    QrySql = QrySql & " Pendiente='1'"	
    QrySql = QrySql & " WHERE"
    QrySql = QrySql & " PH_Consumo_Detalle_Productos.Id_Consumo_detalle_productos = " & idConsumo
    '		
    Set objExec = conexion.Execute(QrySql)
    Set objExec = Nothing	
	'
	' Buscar el id del Consumo 
	'				
	QrySql = vbnullstring
	QrySql = " SELECT id_consumo FROM PH_Consumo_Detalle_Productos WHERE PH_Consumo_Detalle_Productos.Id_Consumo_detalle_productos = " & idConsumo
	Set rsSql = Server.CreateObject("ADODB.recordset")
	rsSql.Open QrySql, conexion
	if not (rsSql.EOF and rsSql.BOF) then
		idConsumoPrincipal = rsSql(0)
	end if
	rsSql.close
	Set rsSql = Nothing
	'
	' Validar Consumo
	'
	QrySql = vbnullstring
	QrySql = " SELECT COUNT(Id_Consumo) AS total FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND Pendiente='0' AND Id_Consumo = " & idConsumoPrincipal	
	Set rsSql = Server.CreateObject("ADODB.recordset")
	' Response.Write QrySql & "<br><br>"
	'		
	rsSql.Open QrySql, conexion
	'
	if not (rsSql.EOF and rsSql.BOF) then
		Total = rsSql(0)
	end if
	'	
	If CInt(Total) = 0 Then
		'
		' Response.Write "Cero "
		' Response.End
		
		QrySql = vbnullstring
		QrySql = QrySql & " UPDATE PH_Consumo"
		QrySql = QrySql & " SET"
		QrySql = QrySql & " VALIDADO='1'"	
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_Consumo.id_consumo = " & idConsumoPrincipal
		'        
		Set objExec = conexion.Execute(QrySql)
		Set objExec = Nothing
		'
	end if
	'
    If Err.Number = 0 Then
		Response.write Total
    Else
        Response.write (Err.Description)
    End If
    '	
    conexion.Close
    Set conexion = Nothing
	'
%>