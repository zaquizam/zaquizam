<% @LCID = 1034%>
<!DOCTYPE HTML>
<html >
<head>
	<title>McDonald's</title>

    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="sp.css" rel="stylesheet" type="text/css" media="screen" />
    
    <!--link rel="stylesheet" href="http://www.w3schools.com/lib/w3.css"-->
<style>
 .w3-card {border:1px solid #ccc}
.w3-card-2,.w3-example {box-shadow:0 2px 4px 0 rgba(0,0,0,0.16),0 2px 10px 0 rgba(0,0,0,0.12) !important}
.w3-card-4 {box-shadow:0 4px 8px 0 rgba(0,0,0,0.2),0 6px 20px 0 rgba(0,0,0,0.19) !important}
</style>

</head>
<body topmargin="0">
<!--#include file="sp_sub.asp"-->
<!--#include File="in_RepDatJL.asp"-->
<!--#include file="in_DataE.asp"-->


<%
  
'==========================================================================================
' Variables y Constantes
'==========================================================================================
'	Dim SqlReg
'	Dim SqlCla
    dim PDF
    dim gPre
    Dim gDat
    Dim rsp
    Dim sRet
    Dim sCap
    Dim sDis
	
    
	Session.Timeout =60

	sPro="sp_rGeneralN.asp"

    iErr=0
    rr_iwFoto=0
   ' rr_iCombo=1 
	'rr_iDis = 0
	 

    iRep=2
   
Sub rr_ParDat

	'rr_sqlWhe = " WHERE (((pr_O_DataProcesada.IdConsecutivo>= " & rr_MesDes & ") And (pr_O_DataProcesada.idConsecutivo<= " & rr_MesHas & "))  And "
	'rr_sqlWhe = " WHERE (((sp_O_Tiendas.Gerencia) Is Not Null) AND ((Month([Fec_Cierre])+(Year([Fec_Cierre])*12)) Is Not Null) And "
	rr_sqlWhe = "  (sp_O_Tiendas.Gerencia Is Not Null) AND (sp_O_Postulados.Id_Periodo Is Not Null) AND (sp_O_Tiendas.ConsultorOp Is Not Null) "
    ilen=len(rr_sqlwhe)
    
    rr_sqlGru=""
 
  
'-------------------------------------------------------------------------------------------    	
' Configuración de la Data
' 0 = Posición (1=combo, 2=fila,3=columna)
' 1 = Nombre
' 2 = Sql
' 3 = Cambio de Dimension=1  0=nose puedde cambiar
' 4 = Ancho del campo
' 5 = Total fila
' 6 = 0 Es igual  1 = Like 9=no filtro
' 7 = Total combo
' 10 =Campo
'-------------------------------------------------------------------------------------------    	

   
' Llenar Nombre de Dimensiones

    rr_sDim(01,1) = "Gerencia"         	:rr_sDim(01,0)  = 1:rr_sDim(01,3) = 1:rr_sDim(01,4)  = 10:rr_sDim(01,5)=1:rr_sDim(01,7)=1:rr_sDim(01,6)=0
    rr_sDim(02,1) = "Área"         	   	:rr_sDim(02,0)  = 1:rr_sDim(02,3) = 1:rr_sDim(02,4)  = 10:rr_sDim(02,5)=1:rr_sDim(02,7)=1:rr_sDim(02,6)=0
    rr_sDim(03,1) = "Consultor RRHH"   	:rr_sDim(03,0)  = 1:rr_sDim(03,3) = 1:rr_sDim(03,4)  = 10:rr_sDim(03,5)=1:rr_sDim(03,7)=1:rr_sDim(03,6)=0
	rr_sDim(04,1) = "Compañia"   	   	:rr_sDim(04,0)  = 1:rr_sDim(04,3) = 1:rr_sDim(04,4)  = 10:rr_sDim(04,5)=1:rr_sDim(04,7)=1:rr_sDim(04,6)=0
	rr_sDim(05,1) = "Ciudad"   	   		:rr_sDim(05,0)  = 1:rr_sDim(05,3) = 1:rr_sDim(05,4)  = 10:rr_sDim(05,5)=1:rr_sDim(05,7)=1:rr_sDim(05,6)=0
	rr_sDim(06,1) = "Restaurante"   	:rr_sDim(06,0)  = 1:rr_sDim(06,3) = 1:rr_sDim(06,4)  = 10:rr_sDim(06,5)=1:rr_sDim(06,7)=1:rr_sDim(06,6)=0
	rr_sDim(07,1) = "Centro de Costo"  	:rr_sDim(07,0)  = 1:rr_sDim(07,3) = 1:rr_sDim(07,4)  = 10:rr_sDim(07,5)=1:rr_sDim(07,7)=1:rr_sDim(07,6)=0
	rr_sDim(08,1) = "Nacionalidad"    	:rr_sDim(08,0)  = 1:rr_sDim(08,3) = 1:rr_sDim(08,4)  = 18:rr_sDim(08,5)=1:rr_sDim(08,7)=1:rr_sDim(08,6)=0

    rr_sDim(09,1) = "Sexo"           	:rr_sDim(09,0)  = 2:rr_sDim(09,3) = 1:rr_sDim(09,4)  = 18:rr_sDim(09,5)=1:rr_sDim(09,7)=1:rr_sDim(09,6)=0
	rr_sDim(10,1) = "Edad"           	:rr_sDim(10,0)  = 2:rr_sDim(10,3) = 1:rr_sDim(10,4)  = 18:rr_sDim(10,5)=1:rr_sDim(10,7)=1:rr_sDim(10,6)=0
	
	rr_sDim(11,1) = "Período"       	:rr_sDim(11,0)  = 3:rr_sDim(11,3) = 0:rr_sDim(11,4)  = 12:rr_sDim(11,5)=0:rr_sDim(11,7)=1:rr_sDim(11,6)=0    
	rr_sDim(12,1) = "Retención"       	:rr_sDim(12,0)  = 0:rr_sDim(12,3) = 1:rr_sDim(12,4)  = 12:rr_sDim(12,5)=0:rr_sDim(12,7)=1:rr_sDim(12,6)=0    
	rr_sDim(13,1) = "Disponibilidad"    :rr_sDim(13,0)  = 0:rr_sDim(13,3) = 1:rr_sDim(13,4)  = 12:rr_sDim(13,5)=0:rr_sDim(13,7)=1:rr_sDim(13,6)=0    
	rr_sDim(14,1) = "Capacidad"       	:rr_sDim(14,0)  = 0:rr_sDim(14,3) = 1:rr_sDim(14,4)  = 12:rr_sDim(14,5)=0:rr_sDim(14,7)=1:rr_sDim(14,6)=0    
    	
	'rr_sDim(00,10) = " sp_O_Postulados INNER JOIN sp_O_Tiendas ON sp_O_Postulados.Nombre_Local = sp_O_Tiendas.Nombre_Local "
	rr_sDim(00,10) = " (sp_O_Postulados INNER JOIN sp_O_Tiendas ON sp_O_Postulados.Nombre_Local = sp_O_Tiendas.Nombre_Local) INNER JOIN ss_T_Periodo ON sp_O_Postulados.Id_Periodo = ss_T_Periodo.IdPeriodo "
    
    rr_sDim(01,10) = " sp_O_Tiendas.Gerencia "
    rr_sDim(02,10) = " sp_O_Tiendas.Area "
    rr_sDim(03,10) = " sp_O_Tiendas.Consultor_RRHH "
    rr_sDim(04,10) = " sp_O_Tiendas.Compania "
    rr_sDim(05,10) = " sp_O_Tiendas.Ciudad "
    rr_sDim(06,10) = " sp_O_Tiendas.Nombre_Local "
    rr_sDim(07,10) = " sp_O_Tiendas.Centro_Costo "
    rr_sDim(08,10) = " sp_O_Postulados.Nacionalidad "
    rr_sDim(09,10) = " sp_O_Postulados.Sexo "
    rr_sDim(10,10) = " sp_O_Postulados.Edad "
	rr_sDim(11,10) = " sp_O_Postulados.Id_Periodo  "
	rr_sDim(12,10) = " sp_O_Postulados.Num_Retencion "
	rr_sDim(13,10) = " sp_O_Postulados.Num_Disponibilidad "
	rr_sDim(14,10) = " sp_O_Postulados.Num_Capacidad "

'Gerencia
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Gerencia as iGerencia, "
	sql = sql & " sp_O_Tiendas.Gerencia"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Gerencia "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Gerencia Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Gerencia "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(01,2)=sql

'Area
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Area as iArea, "
	sql = sql & " sp_O_Tiendas.Area"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Area "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Area Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Area "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(02,2)=sql

'Consultor_RRHH
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Consultor_RRHH as iConsultor_RRHH, "
	sql = sql & " sp_O_Tiendas.Consultor_RRHH"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Consultor_RRHH "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Consultor_RRHH Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Consultor_RRHH "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(03,2)=sql

'Compania
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Compania as iCompania, "
	sql = sql & " sp_O_Tiendas.Compania"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Compania "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Compania Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Compania "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(04,2)=sql

'Ciudad
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Ciudad as iCiudad, "
	sql = sql & " sp_O_Tiendas.Ciudad"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Ciudad "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Ciudad Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Ciudad "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(05,2)=sql

'Nombre_Local
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Nombre_Local as iNombre_Local, "
	sql = sql & " sp_O_Tiendas.Nombre_Local"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Nombre_Local "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Nombre_Local Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Nombre_Local "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(06,2)=sql

'Centro_Costo
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Tiendas.Centro_Costo as iCentro_Costo, "
	sql = sql & " sp_O_Tiendas.Centro_Costo"
	sql = sql & " FROM "
	sql = sql & " sp_O_Tiendas "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Tiendas.Centro_Costo "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Tiendas.Centro_Costo Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Tiendas.Centro_Costo "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(07,2)=sql

'Nacionalidad
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Nacionalidad as iNacionalidad, "
	sql = sql & " sp_O_Postulados.Nacionalidad"
	sql = sql & " FROM "
	sql = sql & " sp_O_Postulados "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Postulados.Nacionalidad "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Postulados.Nacionalidad Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Nacionalidad "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(08,2)=sql

'Sexo
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Sexo as iSexo, "
	sql = sql & " sp_O_Postulados.Sexo"
	sql = sql & " FROM "
	sql = sql & " sp_O_Postulados "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Postulados.Sexo "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Postulados.Sexo Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Sexo "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(09,2)=sql

'Edad
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Edad as iEdad, "
	sql = sql & " sp_O_Postulados.Edad"
	sql = sql & " FROM "
	sql = sql & " sp_O_Postulados "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Postulados.Edad "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Postulados.Edad Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Edad "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(10,2)=sql

'Periodo
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Id_Periodo, "
	sql = sql & " ss_T_Periodo.Periodo "
	sql = sql & " FROM sp_O_Postulados INNER JOIN ss_T_Periodo ON sp_O_Postulados.Id_Periodo = ss_T_Periodo.IdPeriodo "
	sql = sql & " GROUP BY "
	sql = sql & " sp_O_Postulados.Id_Periodo, "
	sql = sql & " ss_T_Periodo.Periodo "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Id_Periodo "
    rr_sDim(11,2)=sql

'Num_Retencion
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Num_Retencion as iNum_Retencion, "
	sql = sql & " sp_O_Postulados.Num_Retencion"
	sql = sql & " FROM "
	sql = sql & " sp_O_Postulados "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Postulados.Num_Retencion "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Postulados.Num_Retencion Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Num_Retencion "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(12,2)=sql
	
'Num_Disponibilidad
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Num_Disponibilidad as iNum_Disponibilidad, "
	sql = sql & " sp_O_Postulados.Num_Disponibilidad"
	sql = sql & " FROM "
	sql = sql & " sp_O_Postulados "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Postulados.Num_Disponibilidad "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Postulados.Num_Disponibilidad Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Num_Disponibilidad "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(13,2)=sql

'Num_Capacidad
    sql = ""
    sql = sql & " SELECT "
	sql = sql & " sp_O_Postulados.Num_Capacidad as iNum_Capacidad, "
	sql = sql & " sp_O_Postulados.Num_Capacidad"
	sql = sql & " FROM "
	sql = sql & " sp_O_Postulados "
	sql = sql & " GROUP "
	sql = sql & " BY sp_O_Postulados.Num_Capacidad "
	sql = sql & " HAVING "
	sql = sql & " sp_O_Postulados.Num_Capacidad Is Not Null "
	sql = sql & " ORDER BY "
	sql = sql & " sp_O_Postulados.Num_Capacidad "
    'response.write "<br>151 sql:= " & sql
    rr_sDim(14,2)=sql

'-------------------------------------------------------------------------------------------    	      
' Configuración de Variables   
' 0 Normal    100 Porcentaje Normal
' 1 Suma      101 Porcentaje de suma
' 2 Cuenta    102 Porcenta de Cuenta
' 3 Promedio  103 Porcentaje del promedio
' 4 Maximo
' 5 Minimo 
' Posición
'    4 = Calculo de Totales (1=suma, 2Cuenta,3=Promedio,4=Maximo,5=minimo)
'    5 = Calculo de Share (1=si, 0=N0, 2= dividir)
'-------------------------------------------------------------------------------------------    	    
	rr_nVar = 4 ' Numero de columnas de las variables
	rr_mVar = 4 ' Numero de columnas de las variables
	rr_sVar(0,0) = "Variables":
	rr_sVar(0,1) = 6 ' Ancho de la variable
	rr_sVar(0,2) = 6 ' Ancho de la variable
	rr_sVar(0,3) = 6 ' Ancho de la variable
	rr_sVar(0,4) = 6 ' Ancho de la variable
   'Nombre de la variable          	Factor Multiplipli   Numero de Decima   #campo con la data  Tipo de cálculo

	rr_sVar(1,0) = "Prom. Retención"   		 :rr_sVar(1,1) = 1        :rr_sVar(1,2) = 0  :rr_sVar(1,3) = 25  :rr_sVar(1,4) = 1  :rr_sVar(1,5) =2:rr_sVar(1,6) =1:rr_sVar(1,7) =4
	rr_sVar(1,10) = "sum(Num_Retencion)"
	rr_sVar(2,0) = "Prom. Disponibilidad"   		:rr_sVar(2,1) = 1    :rr_sVar(2,2) = 0  :rr_sVar(2,3) = 25  :rr_sVar(2,4) = 1  :rr_sVar(2,5) =2:rr_sVar(2,6) =2:rr_sVar(2,7) =4
	rr_sVar(2,10) = "Sum(Num_Disponibilidad)"
	rr_sVar(3,0) = "Prom. Capacidad"   			:rr_sVar(3,1) = 1    :rr_sVar(3,2) = 0  :rr_sVar(3,3) = 25  :rr_sVar(3,4) = 1  :rr_sVar(3,5) =2:rr_sVar(3,6) =3:rr_sVar(3,7) =4
	rr_sVar(3,10) = "Sum(Num_Capacidad)"
	rr_sVar(4,0) = "Total Postulados"   			:rr_sVar(4,1) = 1    :rr_sVar(4,2) = 0  :rr_sVar(4,3) = 25  :rr_sVar(4,4) = 1  :rr_sVar(4,5) =0
	rr_sVar(4,10) = "Count(Num_Capacidad)"
	ipGra=1 

End Sub
Sub vGraBar 
	exit sub
'    lProceso

    
%>
   <br /><br /><br />
   <!--div style="font-size:10px; text-align:right; margin-top:0px"> Nota: Porcentaje de tiendas visitadas:   </div-->    
   

   <%
end sub

Sub vGrafico (IIReg)
if rr_ipExe=1 then exit sub
	
dim iValor(4)
	'response.write "<br>379 Paso" & iiReg	
  if iiReg<>0 then exit sub
  
  For iiVar=1 to rr_mVar 
      isTot=0
      if rr_gTit(iiReg,0)<>0 then isTot=1
        for iiCol=rr_ncol to rr_nCol
'response.write "<br>219 iireg:=" & iireg & " rr_nCol:=" & rr_ncol & " rr_xVar:=" & rr_mVar & "vALOR:=" 
							if rr_gData(iiReg,iiCol, iiVar) then
								if rr_gData(iiReg,iiCol, iiVar)=0  then 
									response.Write "-"
									'response.Write "No Disponible"
								else
								    ' Share
								    if rr_sVar(iiVar,5)=1 then
								        iV=rr_gData(iiReg,iiCol, iiVar)/rr_gData(0,iiCol, iiVar)*100
								        iV= iV *  rr_sVar(iiVar,1 )
								    elseif rr_sVar(iiVar,5)=2 then
								        ix=rr_gData(iiReg,iiCol, rr_sVar(iiVar,7))
								        'response.write "<br>1218 " & ix & " Var7:=" & rr_sVar(iiVar,7)& " Var6:=" & rr_sVar(iiVar,6)& "<br>"
								        if ix<>0 then
									        iV=rr_gData(iiReg,iiCol, rr_sVar(iiVar,6))/ix
    'response.write "<br>1221 " & " Var6:=" & rr_gData(iiReg,iiCol, rr_sVar(iiVar,6))& " Var7:=" & ix & "<br>"									        
									    else    
									        iv=0
									    end if        
								        iV= iV *  rr_sVar(iiVar,1 )
								    else
								        ' Promedio
								        iv=rr_gData(iiReg,iiCol, iiVar)
    								    if rr_sVar(iiVar,4)= 3 then

    								        if istot<>0 then 
    								     'response.write "<br>1257 =" & rr_gData(iireg,iiCol, 10) & "="     								        
    								            iv=iv/rr_gData(iireg,iiCol, 10)
    								        
    								        end if
									    end if
									    iV= iv *  rr_sVar(iiVar,1 ) 
									   ' response.write "<br>1175 iv:=" & iv
									end if 
									
									ixDec=rr_sVar(iiVar,2)
									
									if iv>99.98 and ixdec>1 then ixdec=1
								    if iv then
								        iPro=iPro +Iv
								        IFre=iFre+1
								      '  response.Write FormatNumber(iV,ixDec, -1, 0, -1)  
								      
								        iValor(iiVar)=IV
									  
								    else
								      '  response.Write "-"
                                    end if									    
								end if	
							else
							    'response.Write " -" 
							end if
        next
    next
    for i=1 to 3
        iValor(i)=iValor(i)+0.4
        iValor(i)=int(iValor(i))
    next    
   
%>
</br>
</br>
</br>
</br>
</br>
</br>
</br>
</br>
</br>
</br>
</br>
</br>
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
   <script type="text/javascript">
      google.charts.load('current', {'packages':['gauge']});
      google.charts.setOnLoadCallback(drawChart);
      function drawChart() {

        var data = google.visualization.arrayToDataTable([
          ['Indicador', 'Value'],
          ['Retención', <%=int(iValor(1)) %>]
 
        ]);

        var options = {
          width: 200, height: 150,
          redFrom: 0, redTo: 11,
          yellowFrom:12, yellowTo: 14,
          greenFrom:15, greenTo: 36,
          minorTicks: 5,
		  max:36
        };

        var chart = new google.visualization.Gauge(document.getElementById('chart_div1'));

        chart.draw(data, options);

    
      }
    </script>   
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
   <script type="text/javascript">
     
      google.charts.setOnLoadCallback(drawChart);
      function drawChart() {

        var data = google.visualization.arrayToDataTable([
          ['Indicador', 'Value'],
          ['Disponibilidad', <%=int(iValor(2)) %>]

        ]);

        var options = {
          width: 200, height: 150,
          redFrom: 0, redTo: 0,
          yellowFrom:0, yellowTo: 5,
          greenFrom:6, greenTo: 12,
          minorTicks: 5,
		  max:12
        };

        var chart = new google.visualization.Gauge(document.getElementById('chart_div2'));

        chart.draw(data, options);

    
      }
    </script>   	
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
   <script type="text/javascript">
     
      google.charts.setOnLoadCallback(drawChart);
      function drawChart() {

        var data = google.visualization.arrayToDataTable([
          ['Indicador', 'Value'],

          ['Capacidad', <%=int(iValor(3)) %>]
        ]);

        var options = {
          width: 200, height: 150,
          redFrom: 0, redTo: 26,
          yellowFrom:27, yellowTo: 48,
          greenFrom:49, greenTo: 54,
          minorTicks: 5,
		  max:54

        };

        var chart = new google.visualization.Gauge(document.getElementById('chart_div3'));

        chart.draw(data, options);

    
      }
    </script>  	
     <div class="w3-card-4" style="background-color:White; width:70%;  padding:20px 0 10px; margin-top:70px; margin-bottom:30px; margin-left:auto; margin-right:auto; height:250px">
     <div style="font-size:30px; font-family:Sans-Serif; text-align:center ">Indicadores del último mes</div>
        <div id="chart_div1" style="width: 200px; height: 100px; margin-top:30px; float:left"></div>
		<div id="chart_div2" style="width: 200px; height: 100px; margin-top:30px; float:left"></div>
		<div id="chart_div3" style="width: 200px; height: 100px; margin-top:30px; float:left"></div>
    </div> 
   
<%   
end sub


'=======================
'=======================
    Apertura
    LeePar
    ipEst=193
    ipCli=10
    lUsuario
    if ed_sBus<>"" then ipCue=ed_sBus
	%>
	<table width="98%" align="center" border="0" bgcolor= "#ffb81c" style="margin:10px 10px 0px 10px;border-radius: 10px;" class="w3-example w3-card-8"> 
		<tr>
		<td> 
		<table width="94%" border="0" height="80" cellspacing="1" cellpadding="0" bgcolor="#ffffff" ID="Table8" style="margin:30px 0px 0px 0px; margin-left:auto; margin-right:auto" >      
	<tr><td>
	<%
	leePar
	rr_LeePar
	if rr_ipExe<>1 then 
	    Encabezado 
	else
        Response.ContentType = "application/vnd.ms-excel"
         Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
    end if
	rr_ipGra=1	
	'rr_ipdis=0

	%>


	</td></tr>  
	<tr><td>
		
	</td></tr>

	<tr><td>
<% if rr_ipExe<>1 then %>		
		   <%ed_MenPri %>
			<table width="100%" align="left" style="margin-left:0px; margin-top:0px">
			<tr>
			<td style="width:14%; height:25px;background-color:#1d9656"></td>
			<td style="width:14%; height:25px;background-color:#61c8cf"></td>
			<td style="width:14%; height:25px;background-color:#0faec1"></td>
			<td style="width:14%; height:25px;background-color:#f5cb0c"></td>
			<td style="width:14%; height:25px;background-color:#3a4a5f"></td>
			<td style="width:14%; height:25px;background-color:#f6891f"></td>
			<td style="width:14%; height:25px;background-color:#ef3157"></td>
			</tr>
			</table>            
<% end if			
	sPro="sp_rGeneralN.asp"
	Calpar
%>

<div style="width:98%; margin-left:auto; margin-right:auto">
<%
    rr_sTit="<div style='font-size:24; font-weight:bold'>Reporte de Indicadores</div>"
   rr_Main

%>
</div>
 

    <%conexion.close%>
	
    </td></tr>
    </table>
    <br />
    </td></tr>
    </table>
   
</body>
</html>
