<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'Response.AddHeader "Content-Type","application/json;charset=utf-8"
	'
	' g_pValBuscarProductosPendientes - 02mar21
	'
	Session.lcid		= 1034
	Response.CodePage 	= 65001
	Response.CharSet 	= "utf-8"	
	'
	Dim rsProductosPendientes, arrProductosPendientes
	'	
	' Buscar Los Productos Pendientes por info completa del codigo de Barras
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT TOP (100)  PERCENT"
	QrySql = QrySql & " cacevedo_atenas.PH_Consumo_Detalle_Productos.Numero_codigo_barras,"
	QrySql = QrySql & " COUNT ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos ) AS Total"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " cacevedo_atenas.PH_Consumo_Detalle_Productos"
	QrySql = QrySql & " LEFT OUTER JOIN cacevedo_atenas.PH_CB_Producto"
	QrySql = QrySql & " ON cacevedo_atenas.PH_Consumo_Detalle_Productos.Numero_codigo_barras = cacevedo_atenas.PH_CB_Producto.CodigoBarra"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Id_Hogar > 1 )"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Pendiente = 1 )"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Tiene_Codigo_Barras = 1 )"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Status_registro='G')"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " ( cacevedo_atenas.PH_CB_Producto.Id_Producto IS NULL )"
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " cacevedo_atenas.PH_Consumo_Detalle_Productos.Numero_codigo_barras"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " COUNT ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos ) DESC"
	'
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsProductosPendientes = Server.CreateObject("ADODB.recordset")
	rsProductosPendientes.Open QrySql, conexion
	'
	if not rsProductosPendientes.EOF then
    	arrProductosPendientes = rsProductosPendientes.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	'Response.ContentType = "application/json"	 
	'
	' Crear Archivo Array Json
	'
	sTabla=vbnullstring

    if IsArray(arrProductosPendientes) then

        For i = 0 to ubound(arrProductosPendientes, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "Id" 	& chr(34) & ":" & CStr(arrProductosPendientes(0,i)) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Name"  & chr(34) & ":" & chr(34) & arrProductosPendientes(1,i)  & chr(34) & chr(125) & chr(44)
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
	rsProductosPendientes.Close
	Set rsProductosPendientes = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>