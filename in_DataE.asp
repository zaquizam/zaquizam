<%

Dim iSession
iSession=Request.Cookies("idsession")
'response.write "<br>" & iSession
if iSession="" then
    iSession=Session.SessionID 
    Response.Cookies("idsession")=Session.SessionID 
end if    
Response.Cookies("idsession").Expires=now()+365

'================================
' Mejoras
' Se incluyó Vercombo 23/06/2015
' Se incluyó Menver
'================================
' Próximas Mejoras
'   Campo de campos
'   Validaciones de campo
'   Guardar en GET ROWS solo registros a presentar
'   Copiar del registro anterior al actual
'   Colocar Frame para mover data a la derecha
'   Logs de todos los cambios
'================================

 %>
 
<script type="text/javascript">	
function mensaje(sMensaje) 
{
alert(sMensaje)
}

function valnum(valor,snomcam)
{
    var valor2 = valor.replace(",", ".");
if (isNaN(valor2)) 
{
alert('El dato debe ser un numero y colocó: ' +document.getElementById(snomcam).value)

document.getElementById(snomcam).focus()
document.getElementById(snomcam).click()
document.getElementById(snomcam).style.background="#800000"
document.getElementById(snomcam).value=""

}
else
{

//En caso contrario (Si era un número) devuelvo el valor
document.getElementById(snomcam).style.background="#ffffff"
}
}
function valfec(valor,snomcam)
{
    var valor2 = valor.replace(",", ".");
if (isNaN(valor2)) 
{
alert('El dato debe ser un numero y colocó: ' +document.getElementById(snomcam).value)

document.getElementById(snomcam).focus()
document.getElementById(snomcam).click()
document.getElementById(snomcam).style.background="#800000"
document.getElementById(snomcam).value=""

}
else
{

//En caso contrario (Si era un número) devuelvo el valor
document.getElementById(snomcam).style.background="#ffffff"
}
}

function valtxt(texto,snomcam) { 
   
    if (texto.charCodeAt(0)==32) 
        {
        alert('El primer valor del campo no pueder ser blanco');
        document.getElementById(snomcam).focus()
        document.getElementById(snomcam).style.background="#fffff0"
         }
        else
        {
        //En caso contrario (Si era un número) devuelvo el valor
        document.getElementById(snomcam).style.background="#ffffff"
        }

}   

function maximaLongitud(texto,maxlong) { 
  var tecla, in_value, out_value; 

  if (texto.value.length > maxlong) { 
  alert('El máximo numero de caracteres es: ' +maxlong)
    in_value = texto.value; 
    out_value = in_value.substring(0,maxlong); 
    texto.value = out_value; 
    return false; 
  } 
  return true; 
}   
 
function valPre(valor,snomcam)
{
alert('El debe ser llenado: ' +document.getElementById(snomcam).value)

if (valor=="") 
{
alert('El debe ser llenado: ' +document.getElementById(snomcam).value)

document.getElementById(snomcam).focus()
document.getElementById(snomcam).click()
document.getElementById(snomcam).style.background="#800000"
document.getElementById(snomcam).value=""

}
else
{

//En caso contrario (Si era un número) devuelvo el valor
document.getElementById(snomcam).style.background="#ffffff"
}
}

function textCounter(field, countfield, maxlimit) 
{
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);
// otherwise, update 'characters left' counter
else 
countfield.value = maxlimit - field.value.length;
}


function  ddate_click( sNomCam, sFec)
{
var xmlhttp;

//alert('El dato debe ser un numero y colocó:'+sFec+'-'+sNomCam)
if (window.XMLHttpRequest)
  {
  // code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else if (window.ActiveXObject)
  {
  // code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
else
  {
  alert("Your browser does not support XMLHTTP!");
  }


xmlhttp.onreadystatechange=function() 
{

if(xmlhttp.readyState==4)
  {
   sx="Div"+sNomCam
   sy=document.getElementById(sx).title
   
   if (sy=="Calendarioxx") 
   {
   document.getElementById(sx).title=""
   document.getElementById(sx).innerHTML=""
   document.getElementById(sx).style.height=1;
   document.getElementById(sx).style.visibility="hidden";
   }
   else
   {
   document.getElementById(sx).title="Calendario"
   document.getElementById(sx).style.height=240;
   document.getElementById(sx).style.visibility="visible";
   document.getElementById(sx).innerHTML=xmlhttp.responseText;
   }
   
  }

}
url="aj_Calendario1.asp?backf="+sNomCam+"&cDate="+sFec;
xmlhttp.open("GET",url, false);
xmlhttp.send(null);
}
function SetDate(aDate,sCampo)
{
sx="Div"+sCampo

document.getElementById(sx).innerHTML="";
document.getElementById(sx).style.visibility="hidden";
document.getElementById(sx).style.height=1;
document.getElementById(sCampo).readOnly = false;
document.getElementById(sCampo).value=aDate;
document.getElementById(sCampo).readOnly = true;

}

   function upload(scam) 

   {
 
    var client;
    if (window.XMLHttpRequest)
      {client=new XMLHttpRequest();}
    else if (window.ActiveXObject)
            {client=new ActiveXObject("Microsoft.XMLHTTP");}
    else
        {alert("Your browser does not support XMLHTTP!");}
    var sImg='/images/loading.gif'

    document.getElementById("ImgFact").src=sImg;
    document.getElementById("ImgFact").width="80";  
    client.onreadystatechange=function() 
    {
    if(client.readyState==4)
    {
    document.getElementById("ImgFact").width="240";
    document.getElementById("ImgFact").src=client.responseText;

    document.getElementById(scam).value=client.responseText;
    //document.getElementById("ImgFact").visibility='visible'
  
    }
    }       
    var file = document.getElementById("fileElem");
    var formData = new FormData();
    formData.append("upload", file.files[0]);
    client.open("post", "aj_UpLoad.asp", true);
    client.setRequestHeader("Content-Type", "multipart/form-data");
    client.send(formData);  /* Send to server */ 
   }
</script>
<%
'==========================================================================================
' Tipo de Campos Permitidos
'==========================================================================================
' Ida_	Es un Id de un campo autonumérico , no requiere un query para hacer
' Id_	Es un Id y requiere un query para hacer Join en ed_sQue(icam,1)  iCam= Es el número del campo
' Idy_	Crea un menu de año
' Idm_	Crea un menu de mes
' Idd_	Crea un menu de día
' Ind_	Crea un menú de indicador Si o NO
' Pass es Cun campo tipo Password (no es visible para el usuario)

' SqlCla		En esta variable se coloca la instrucción sql de la  pagina principal
' sQlreg		En esta variable se almacena la instrucción sql para la pagina del formato de actualización
' sPro			Se almacena el nombre del programa a ejecutar con el click el registro
' iNumCam		El número de campo a presentar en la primera Pantalla
' ed_iRegPag		Número de registros por página
' sNomTab		Nombre de la tabla
' sNomInd		Nombre de la clave principal de la tabla
' ed_cCol		Número de la columna a Ordenar
' ed_cOrd		Sirve para especifica el orden de ordenamiento	' Orden 0=ascendente 1=descendente
' ed_iRan		Se emplea para presentar un ranking de los registros
' ed_iRep		AlAlmacenar uno en esta columna Sirve para convertir el formulario en un reporte
' ed_Bot(1)="disabled"	Desabilita los botones 1=Excel 2=añadir  3=guardar 4=eliminar
' ed_iSwReg     Lee solo un registro
' ed_iDet      Si es <>0 dice que hay un formulario de detalle, el numero indica el numero del campo linkeado




' ed_Edi(99) ' Permite la edición de los campos que contienen un 1 en la pantalla principal
' ed_link ' Es la programa que se usa cuando hay click en la pantalla principal, es su defecto se usa sPro
' ed_Pulsar(9,2) ' Botones que aparecen en cada registro
' ed_sCampo(99,9) '0.- Titulo del campo 1= Default 2= No Presentar 3=Read only 4=1-Obligatorio 5.-Tool Tips 6= Total del campo 7= 1=salto, 8- 1=Copia Valor anterior
' ed_iSum(99) ' Sumatoria del campo
' ed_iGrupo ' Código del grupo Leido en el perfil
' ed_sTarget  Target del sLink
' ed_sBotonC ' Botones 0.- Texto del Boton, 1.- Link , 2.- Target 3.- Tools Tips

'===================== Tipos de Campos =====================
'1- Mas de 50 Caracteres
'2- Combo
'3- Indicador IND
'4- Autonumérico
'5- PAssword
'===========================================================


'==========================================================================================
' Variables y Constantes
'==========================================================================================

    Dim ngCat
	Dim sGra(99) 'Data A Grabar
	Dim ed_iNumCam2
	Dim ed_iNumCam
	
	Dim SqlReg
	Dim SqlCla
	Dim ed_rs1
	
	Dim TipCam(99,4) ' 0= (string, numerico, fecha) 1.-Tipo(indicador, password, etc) 2.-lineas de textarea  3.-Numero de caracteres 4.-Obligatorio 9.-Color
	Dim ed_iPas
	Dim ed_iCla
	Dim ed_iOrd
	Dim ed_iCol
	Dim ed_cOrd
	Dim ed_cCol
	Dim ed_sBus
	Dim ed_iLof
	Dim ed_iRan
	Dim ed_iRegPag ' Numero de registros por pagina
	Dim ed_iSwReg ' Presentar número de registros 0=si 1=no
	Dim ed_iRep
	Dim ed_idet   ' Detalle
	Dim ed_sFil  ' Filtro
	Dim ed_Bot(9)
	dim ed_Edi(99) ' Permite la edición de los campos que contienen un 1 en la pantalla principal
	dim ed_ierr
	dim ed_iErrG ' Swiche de error de grabacion
	dim ed_link ' Es la programa que se usa cuando hay click en la pantalla principal, es su defecto se usa sPro
	Dim ed_Sqlfil ' Sql del menú de opciones
	Dim ed_iFil
	Dim ed_Pulsar(9,2)
	Dim ed_iMaxReg
	Dim ed_sCampo(99,8)
	Dim ed_Formato(99, 4) ' 0=Salto , 1-Columna(pixel) , 2-Ancho Campo (Caracteres), 3-Ancho Texto(Pixel), 4-Fila (textArea)
	dim ed_sTit(99,4) ' 0=Texto 1=Class 2= comienzo del Marco 3= Fin del Marco 4=texto
	dim ed_iSum(99) ' Sumatoria del campo
	dim ed_swSum ' Campo a sumar
	dim ed_sNomTab
	dim ed_sNomInd
    dim ed_sQue(99,1)
    dim ed_iMp
    dim ed_iMs
    dim ed_ipDet ' Detalle
    dim ed_sErr(99) ' Grabar Texto de error
    '* dim mySmartUpload
    Dim ed_sDirUp ' Carpeta para subir los archivos (imagenes, etc)
    dim ed_sJoin(99,4) ' campos del Join 0-Numero del campo de la tabla 1.- Campo Join 2.- Campo a Mostrar 3.- Noombre de tabla 4.-No Hacer Join
    dim ed_iJoin ' Numero de Join
    dim ed_sWhe ' se usa para añadir el where
    dim ed_sMenSec ' Titulo Menu Secundrio
    dim ed_iDes  ' Desarrollo
	Dim sAcc
	Dim ed_iGrupo
    Dim ed_sTarget
    dim ed_sBotonC(9,3)
    dim ed_iAnc1
    Dim ed_sPar(10,1)  'parametros del Combo
    dim ed_iCombo ' Numeros de combos
    dim ed_sCombo(10,6) ' 0 = Titulo , 1 Sql , 2 <>Null es un total, 3 Filtro, 4 width tabla, 5 width del titulo , 6 width del combo


'  Object creation
'  ***************
   '* Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")


'sPro=Request.ServerVariables("URL")
'sPro=mid(sPro,2)
if sPro="" then
    sx=Request.ServerVariables("URL")
    sx=Request.ServerVariables("SERVER_NAME") & sx
    sPro=sx
    spro="http://" & sPro
   ' response.write "pro:=" & sPro
end if   
 ' #E8EEFA Azul claro
 ' #aca899 Gris


	Const ed_Fondo3 ="#dddddd"
	
'#EAF1F3 Gris Clarito
'#F4AC33 NAranja Bonito
'#B9B9B9 Gris Oscuro
'#3162a6 Azul Yahoo
'#95b3de Azul Claro de Yahoo
'#d6deec Azul mas Claro de Yahoo
' #BC131A Rojo
' #C7272D TRojo Notipanel

	Const ed_Espacio="&nbsp;"
	ed_iCla=""
	ed_iCla=Request.QueryString("edcla")
	if ed_iCla ="" then ed_iCla=999999

	ed_iOrd=Request.QueryString("edord")
	ed_iCol=Request.QueryString("edcol")

	ed_iPas=""
	ed_iPas=Request.QueryString("edpas")
	if ed_iPas="" then ed_iPas=1
	ed_iPag=""
	ed_iPag=Request.QueryString("edpag")
	if ed_iPag="" then ed_iPag=1
	
	ed_iFil=""
	ed_iFil=Request.QueryString("ed_fil")
	ed_iMp=""
	ed_iMp=Request.QueryString("ed_mp")
	ed_iMs=""
	ed_iMs=Request.QueryString("ed_ms")
	if isNull(ed_iMs) then ed_iMs=""
	ed_ipDet=Request.QueryString("ed_det")

	urldelavisita=request.servervariables("HTTP_REFERER")
	ix = 0
	ix = instr(1,urldelavisita,"?")-1
	if ix > 0 then
		urldelavisita = mid(urldelavisita,1,ix)
	end if
	if sPro <> urldelavisita then
		ed_sPar(1,0)=""
		ed_sPar(2,0)=""
		ed_sPar(3,0)=""
		ed_sPar(4,0)=""
		ed_sPar(5,0)=""
		ed_sPar(6,0)=""
		ed_sPar(7,0)=""
		ed_sPar(8,0)=""
		ed_sPar(9,0)=""
	else
		ed_sPar(1,0)=Request.QueryString("cc_p1")
		ed_sPar(2,0)=Request.QueryString("cc_p2")
		ed_sPar(3,0)=Request.QueryString("cc_p3")
		ed_sPar(4,0)=Request.QueryString("cc_p4")
		ed_sPar(5,0)=Request.QueryString("cc_p5")
		ed_sPar(6,0)=Request.QueryString("cc_p6")
		ed_sPar(7,0)=Request.QueryString("cc_p7")
		ed_sPar(8,0)=Request.QueryString("cc_p8")
		ed_sPar(9,0)=Request.QueryString("cc_p9")
	end if
'response.write " par:=" & ed_sPar(1,0) & " url"& urldelavisita & " pro:=" & spro
    ed_iDes=Request.QueryString("ed_des")
	
    if ed_ides="" then ed_Ides=0
    
    for ias=1 to 10
       ' ed_sPar(ias,0)= Replace(ed_sPar(ias,0),"@@@", " ")
	next     

	
' Leer Buscar 
'response.write "<br> Paso:=" & ed_ipas	
	if ed_iPas<>3 then ed_sBus = request.Form("bus")
	if ed_sBus="" then
	    ed_sBus=Request.QueryString("edbus")
	end if
	ed_sBus=Replace(ed_sBus,"%20", " ")

%>



<%	
	

	
Sub ed_CalPar (xPas,xCla,xPag,xBus,xCol,xOrd,xFil,xMp,xMs)
	sPar = ""
	CalPar
    
    sB=Replace(xBus," ","%20")
	sPar = sPar & "&edpas="  & xPas 
	sPar = sPar & "&edcla="  & xCla
	sPar = sPar & "&edpag="  & xPag 
	sPar = sPar & "&edbus="  & sB 
	sPar = sPar & "&edcol="  & xCol 
	sPar = sPar & "&edord="  & xOrd 
	sPar = sPar & "&ed_fil=" & xFil
	sPar = sPar & "&ed_mp="  & xMp
	sPar = sPar & "&ed_ms="  & xMs
	
	
	
  
	
	
	for ias=1 to 10
	   ' ed_sPar(ias,0)=Replace(ed_sPar(ias,0)," ", "@@@")
	next   
		
	sPar =  sPar &  "&cc_p1=" & ed_sPar(1,0)
	sPar =  sPar &  "&cc_p2=" & ed_sPar(2,0)
	sPar =  sPar &  "&cc_p3=" & ed_sPar(3,0)
    sPar =  sPar &  "&cc_p4=" & ed_sPar(4,0)
    sPar =  sPar &  "&cc_p5=" & ed_sPar(5,0)
    sPar =  sPar &  "&cc_p6=" & ed_sPar(6,0)
    sPar =  sPar &  "&cc_p7=" & ed_sPar(7,0)
    sPar =  sPar &  "&cc_p8=" & ed_sPar(8,0)
    sPar =  sPar &  "&cc_p9=" & ed_sPar(9,0)
    sPar = sPar & "&ed_des="  & ed_iDes
end Sub



Sub CamTit (sxTit)

	sx=ucase(sxTit)
	
	ix=instr(1,sx,"ID_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,4)
		sxTit=ST
	end if	

	ix=instr(1,sx,"IND_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,5)
		sxTit=ST & "?"
	end if	
	ix=instr(1,sx,"IDA_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,5)
		sxTit=ST 
	end if	
	ix=instr(1,sx,"IDY_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,5)
		sxTit=ST 
	end if	
	ix=instr(1,sx,"IDM_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,5)
		sxTit=ST 
	end if	
	ix=instr(1,sx,"IDD_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,5)
		sxTit=ST 
	end if	
	
	ix=instr(1,sx,"NUM_")
	sT=""
	if ix<>0 then
		sT = Mid(sxTIT,5)
		sxTit="#" & ST
	end if	
	
end sub

Function CheSql (sql,name)
		sy=ucase(sQl)
		ix=instr(sy,"SELECT")
		if ix=0 then %>
			<br><br><center><font face= "Verdana" size="3" color="#800000"><b>
			<%Response.Write "Falta Instrucción Sql para Columna:=" & name %>
			<br><br>
			</b></font>
			<br>
<%			chesql=1
		end if

End function
'==========================================================================================
' Leer un registro de la tabla de Categoría
'==========================================================================================
Sub ed_LeePag1 (SqlInp)

'	Dim ed_rs1
	Dim gTem

	if sqlInp="" then 
	    SqlInp = "Select * FROM " & ed_sNomTab
        SqlInp = SqlInp & " WHERE Fec_Inactivo is  Null "     
    end if    
    
 if ed_ides=1 then   %>
    <div style="width:450px;  margin:10px 10px 10px 10px; border: solid 1px #666666; text-align:justify" >
    <%
    response.write "<br>391 sqlinp:=" & sqlinp%>
    </div><%
 end if   

' Abrir Recordset
	set rso = CreateObject("ADODB.Recordset")
	rso.CursorType = 1 ' 0=El cursor solo avanza 2= Puedes avanzar y retroceder 
	rso.LockType = 1
	rso.MaxRecords =1
'response.write "<br>384 " & sqlinp			
	rso.Open sqlinp,conexion

    

	ed_ierr=0
    if ed_iNumCam>rso.fields.Count-5 then
        ed_iNumCam = rso.fields.Count-5
    end if 
'response.write "<br> ed_iNumCam:=" & ed_iNumCam 
 '   if ed_iNumCam<11 then
  '      ed_iAnc1= ed_iNumCam *12
   ' else    
   '     ed_iAnc1= ed_iNumCam *14
   ' end if    
    'if ed_iAnc1<100 Then ed_iAnc1=100
    if isnull(ed_iNumCam) then ed_iNumCam = rso.fields.Count
    if ed_iNumCam="" then ed_iNumCam = rso.fields.Count-5
    if ed_iPas=4 then ed_iNumCam = rso.fields.Count-5


' Diseñar el Join
    if ed_iJoin<>0 then
    
        dim ed_sCamTem(99)
        for i=0 to rso.fields.Count-1
            ed_sCamTem(i)= ed_sNomTab & "." & rso.Fields(i).name
        next
       ' SqlInp = "Select * FROM " & ed_sNomTab
        sxFrom=" From "& ed_sNomTab
        for i=1 to ed_iJoin
            if ed_sJoin(i,4)<>"1" then            
                ed_sCamtem(ed_sJoin(i,0)) = ed_sJoin(i,3) & "." & ed_sJoin(i,2)
                sx =" INNER JOIN " & ed_sJoin(i,3) & " ON " & ed_sNomTab & "." & rso.Fields(ed_sJoin(i,0)).name & "=" & ed_sJoin(i,3) & "." & ed_sJoin(i,1)
                sxFrom=sxFrom & sx
             end if
        next
      '  response.write "<br> " & sxFrom
' Crear el Select
        sx=" Select "
        for i=0 to rso.fields.Count-1
            sx= sx & ed_sCamTem(i)
            if i<rso.fields.Count-1 then
                sx=sx & ","
            end if    
        next
        SqlInp= sx & sxFrom
        SqlInp = SqlInp & " WHERE (" & ed_sNomTab & ".Fec_Inactivo is  Null) "   
        sqlinp=sqlinp & ed_sWhe
'response.write "<br> 393 " & SqlInp
        
        rso.close    
	    rso.LockType = 1
	    rso.MaxRecords =1
	    rso.Open sqlinp,conexion        
    end if
    
    
' Buscar Palabra
	if ed_sBus<>"" then
	  
		sx=" ("
		ix=0
		ii=0
		if isnumeric(ed_sBus) then	
	'response.write"<br>459" & ed_SBus
			for i=0 to rso.fields.Count -5
			    'response.write "<br>" & rso.fields(i).name & "-" & rso.fields(i).Type
		        select case rso.fields(i).Type
			        case  202 
				        ix= ix+1
				        if ix > 1  then sx= sx & " OR "
				        sx = sx & "([" &  rso.fields(i).name & "] like '%" & Ed_sBus & "%' ) "
				        ii=1
			        case 2, 3
				        ix= ix+1
				        if ix > 1  then sx= sx & " OR "
				        sx = sx & "([" &  rso.fields(i).name & "] = " & Ed_sBus & " ) "
				        ii=1
    			end Select
	    	next 	
		else
		  
		    for i=0 to rso.fields.Count -5
		    		 'response.write"<br>459 i:=" & i & " Campo:=" & rso.fields(i).name &  " " & ed_SBus & " tipo:=" & rso.fields(i).Type
		        select case rso.fields(i).Type
			        case  202 
			        
				        ix= ix+1
				        if ix > 1  then sx= sx & " OR "
				        sx = sx & "([" &  rso.fields(i).name & "] like '%" & Ed_sBus & "%' ) "
				        ii=1

				    case else    
				  
    			end Select
	    	next 	
		end if
		sx =sx & " ) " 
		ed_sFil=sx
'Response.write "<br>330 ===================Filtro:=" & sx		
		if ii=1 then rso.Filter=ed_sFil
	end if	
	



' Crear Order By		
		ix = ed_iCol
		ix=int(ix)
'response.write "<br>687  ix=" & ix & " icol:=" & ed_Icol		
		sx=rso.Fields(ix).name
		sx = " ORDER BY " & sx 
'response.write "<br>400 sx:=" & sx & " icol:=" & ed_Icol
		SqlInp2 = SqlInp & sx
		if ed_iOrd=1 then sqlInp2 = SqlInp2 & " Desc "
		


' Leer registro(s) para vizualizar		
	set ed_rs1 = CreateObject("ADODB.Recordset")
	if ed_sBus<>"" then ed_rs1.Filter=ed_sFil
	ed_rs1.CursorType = 1
	ed_rs1.LockType = 1
	
	if ed_iSwReg = 1 and ed_sBus="" Then 
	    ed_rs1.MaxRecords = ed_iRegPag * ed_iPag*2 
	end if    
'response.write "<br>649 " & sqlinp2
	ed_rs1.Open sqlinp2,conexion
	if Not(ed_rs1.EOF) then 
		if ed_iswRep =1 then
			ngCat=ed_rs1.GetRows(ed_rs1.MaxRecords)
		else
			ngCat=ed_rs1.GetRows
		end if	
		ed_Ilof=1
		
' Totalizar Campo
    for i=0 to ed_iNumCam-1
        if ed_sCampo(i,6)=1 then
            for j=0 to ubound(ngCat,2)
                
                if isnumeric(ngCat(i,j)) then
               ' response.write "<br>" & ngCat(i,j)
                    ed_iSum(i)= ed_iSum(i) + cdbl(ngCat(i,j))
                  '  response.write " -- " & ed_Isum(i)
                end if    
            next
            ed_swSum=i
        end if    
    next
		
		
	else 
		ed_ilof=0
		
	end if	
	

 
 
    
    ed_rs1.close
' Tipo de Campo	
		for i=0 to ed_rs1.fields.Count -1
			TipCam(i,1)=null
			CamTip i, ed_rs1.fields(i).name, ed_rs1.fields(i).DefinedSize, ed_rs1.fields(i).Type
'response.write "<br>515 Nombre:=" & 	ed_rs1.fields(i).name & " campo:=" & i					
		next 	
				
	for i=0 to ed_rs1.fields.Count -1
        ix=1
		for j=1 to 8
		    iz= ix and ed_rs1.fields(i).Attributes
		    if j=6 and iz=0 then
		        ed_sCampo(i,4)=1
		    end if
		    ix=ix+ix
		next 
	next

end sub

'==========================================================================================
' Leer un registro de la tabla de Categoría
'==========================================================================================
Sub ed_LeePag2 (SqlInp)
	Dim gTem




' Leer registro(s) para vizualizar		
	set ed_rs1 = CreateObject("ADODB.Recordset")

	ed_rs1.CursorType = 1
	ed_rs1.LockType = 1

	'if ed_iSwReg = 1 Then ed_rs1.MaxRecords = ed_iRegPag * iPag
	'response.Write "<br>528 sql:= " & sqlinp
if ed_ides=1 then   response.write "<br>585 sqlinp:=" & sqlinp	
	ed_rs1.Open sqlinp,conexion
	if Not(ed_rs1.EOF) then 
		ngCat=ed_rs1.GetRows
		ed_Ilof=1
	else 
		ed_ilof=0
	end if	

	ed_rs1.close
	
	if ed_iNumCam2="" then ed_iNumCam2 = ed_rs1.fields.Count -6

' Parametros de los campos
    for i=0 to ed_rs1.fields.Count -1
        if ed_sCampo(i,3)=1 then ed_sCampo(i,3) ="readonly"
    next


' Tipo de Campo	
	for i=0 to ed_rs1.fields.Count -1
		CamTip i, ed_rs1.fields(i).name, ed_rs1.fields(i).DefinedSize, ed_rs1.fields(i).type
        TipCam(i,3)=ed_rs1.fields(i).DefinedSize
' tipo de campos 1=Numerico 2=texto 3=fecha		
		select case ed_rs1.fields(i).type
		    case 2,3,4 ' Numérico
		      TipCam(i,0)=1  
		      TipCam(i,3)=ed_rs1.fields(i).precision
		    case 135 ' Campo fecha
		      TipCam(i,0)=3
		      TipCam(i,1)=8    
		    case else ' Campo texto
		      TipCam(i,0)=2  
		end select			
		 ix=1
		 for j=1 to 8
    	    iz= ix and ed_rs1.fields(i).Attributes
 		    if ix=32 then Tipcam(i,4)=iz
		    ix=ix+ix
		next		
		
	next 	

	for i=0 to ed_rs1.fields.Count -1
	    if ed_sCampo(i,7)<>"" then ed_swLin=1
        ix=1
		for j=1 to 8
		    iz= ix and ed_rs1.fields(i).Attributes
		    if j=6 and iz=0 then
		        ed_sCampo(i,4)=1
		    end if
		    ix=ix+ix
		next 
	next

' Formato
' Colocar Margen Por Default
    if  ed_Formato(99,1)<>"" then
        for i=0 to ed_rs1.fields.Count -1
            if ed_Formato(i,1)="" then ed_Formato(i,1)=ed_Formato(99,1)
        next
    end if
' Colocar Ancho Por Default
    if  ed_Formato(99,3)<>"" then
        for i=0 to ed_rs1.fields.Count -1
            if ed_Formato(i,3)="" then ed_Formato(i,3)=ed_Formato(99,3)
        next
    end if


    
end sub
Sub VerCampo 

    %>
    <table border="1">
     <tr>
    	<td>Name</td>
		<td>DefinedSize</td>
       
		<td>NumericScale</td>
		<td>Precision</td>		
		<td>Status</td> 
		<td>Type </td>
		<td>Attributes</td>
		<td> 01</td>
		<td> 02</td>
        <td> 04</td>		
		<td> 08</td>
		<td> 16</td>
		<td> Null<br /> 32</td>
		<td> 64</td>
		<td>128</td>
		<td>TipCam</td>
		<td>N/A/F</td>
		<td>Tipo</td>
		<td>Lineas</td>
		<td>Nro.Carac.</td>
		<td>Obligatorio</td>
	</tr>   
    <%
	for i=0 to ed_rs1.fields.Count -1
		%>
		<tr>
		<td><%=i%></td>
		<td><%=ed_rs1.fields(i).name%></td>
		
		<td><%=ed_rs1.fields(i).DefinedSize%></td>
		
		<td><%=ed_rs1.fields(i).NumericScale %></td>
		<td><%=ed_rs1.fields(i).Precision%></td>
		<td><%=ed_rs1.fields(i).Status%></td> 
		<td><%=ed_rs1.fields(i).Type %></td>
		
		<td><%=ed_rs1.fields(i).Attributes%></td>
		
		
		
		<% ix=1
		 for j=1 to 8
		    iz= ix and ed_rs1.fields(i).Attributes
		     response.write "<td bgcolor='#ffffff'>" & iz 
		    if j=6 and iz=0 then
		        response.write   "*"
		        ed_sCampo(i,4)=1
		    end if
		     response.write "</td>"  
		    ix=ix+ix
		next %>
		
		    <td> </td>
		    <td><%=TipCam(i,0)%></td>
            <td><%=TipCam(i,1)%></td>
            <td><%=TipCam(i,2)%></td>
            <td><%=TipCam(i,3)%></td>
        </tr>
        
		<%		
	next%>
	</table>
    <%for i=0 to ed_rs1.fields.Count -1
		%>
		<tr>
		<td><%="sqlcla=sqlcla  " & chr(38) & chr(34) & ed_sNomTab & "." & ed_rs1.fields(i).name & "," & chr(34)%><br /></td>	
	<%next %>	
	<%
	

%>
<table border="1">
<tr><td colspan="2">Variables del Servidor</td></tr>
<%
for each x in Request.ServerVariables%>
  <tr><td><%=x%></td>
  <td><%=Request.ServerVariables(x)%></td></tr>
 <%
next
%>
    </table>
<%				
end sub
Sub CamTip (i, name, DefinedSize, iTipo)

'response.write "<br>Campo:" & i & " Name:=" & Name

' SI posee mas de 50 Caracteres
	if DefinedSize >50 then
		TipCam(i,1)=1

        select case definedSize
            case definedsize>49 and definedsize<100
                TipCam(i,2)=2
            case definedsize>99 and definedsize<300
                TipCam(i,2)=5
            case definedsize>299 and definedsize<500
                TipCam(i,2)=5
            case else
                TipCam(i,2)=5
        end select
		
	end if
	
' Campos Password
	sx= Ucase(name)
	ix = instr(1,sx,"PASS")
	'response.Write ix & "-" & sx & "<br>"
	if ix<> 0 then TipCam(i,1)=5
		
' Buscar  data de los Combos	
	sx= Ucase(name)
	ix = instr(1,sx,"ID_")
					
	if ix<> 0 and i<>0 then 
		'response.Write sx & sqlINp
		TipCam(i,1)=2
		set rsx = CreateObject("ADODB.Recordset")
		rsx.CursorType = 1
		rsx.LockType = 1
		sx = ed_sQue(i,0)
		if sx= "" then
		    if ed_sCampo(i,2)<>"1" then
		    %>
		    <font face= "Verdana" size="3" color="#800000">
		    <%response.Write "<br> Falta el 'Que' en Campo numero:=" & i & "  Nombre:=" &name
		    end if
		    exit sub
	    end if	 
'response.write "<br>744 sx:=" & sx	       
		rsx.Open sx ,conexion
		if Not(rsx.EOF) then
			gTem =rsx.GetRows
			ed_sQue(i,1)=""
			for h=0 to ubound(gTem,2)
				ed_sQue(i,1) = ed_sQue(i,1)  & gTem(0,h) & chr(9) & gTem(1,h) & chr(9)
				'response.write "<br>" & gTem(0,h) & gtem(1,h)
			next
		end if
	end if	
	
' Primer campo Autonumérico			
	sx= Ucase(name)
    ix = mid(sx,1,2)="ID"
	if ix<> 0 and i=0 then 
			TipCam(i,1)=4
	end if
	ix = instr(1,sx,"ID_")
	if ix<> 0 and i=0 then 
			TipCam(i,1)=4
	end if	
	ix = instr(1,sx,"IDA_")
	if ix<> 0 and i=0 then 
			TipCam(i,1)=4
	end if		
	
' Tipo de Campo Indicador (IND)
	sx= Ucase(name)
	ix = instr(1,sx,"IND_")
	if ix<> 0 and i<>0 then 
'response.Write "<br> Campo:=" & i & " name:=" & name & " type:=" & iTipo			
    if iTipo<>11 then%>
       <br>
       <div style="width:600px; border:solid 1px #cccccc; height:60px; text-align:center; vertical-align:middle ">
       <br />
       <font face= 'verdana' size='2' color='#ff0000' >&nbsp;&nbsp;<b> Error: El campo "<%=name%>" no es de tipo Boolean(bit) </b></font>
        </div>       
    <%end if
			ed_sQue(i,1) = "True" & chr(9) & "Si" & chr(9)
			ed_sQue(i,1) = ed_sQue(i,1)  & "False" & chr(9) & "No" & chr(9)
			'ed_sQue(i,1) =  "Verdadero" & chr(9) & "Si" & chr(9)
		'	ed_sQue(i,1) = ed_sQue(i,1)  & "Falso" & chr(9) & "No" & chr(9)
			'ed_sQue(i,1) ="1" & chr(9) & "Si" & chr(9)
			'ed_sQue(i,1) = ed_sQue(i,1)  & "0" & chr(9) & "No" & chr(9)
			TipCam(i,1)=6
	end if	
' campo de colores
	ix = instr(1,sx,"CC_")
	if ix<> 0  then 
		TipCam(i,1)=9
	end if		
' Grabar campo Mes
			sx= Ucase(name)
			ix = instr(1,sx,"IDM_")
			if ix<> 0 and i<>0 then 
					TipCam(i,1)=3
					ed_sQue(i,1) = "1" & chr(9) & "Enero" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "2" & chr(9) & "Febrero" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "3" & chr(9) & "Marzo" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "4" & chr(9) & "Abril" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "5" & chr(9) & "Mayo" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "6" & chr(9) & "Junio" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "7" & chr(9) & "Julio" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "8" & chr(9) & "Agosto" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "9" & chr(9) & "Septiembre" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "10" & chr(9) & "Octubre" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "11" & chr(9) & "Noviembre" & chr(9)
					ed_sQue(i,1) = ed_sQue(i,1)  & "12" & chr(9) & "Diciembre" & chr(9)
			end if	

' Grabar campo Dia
			sx= Ucase(name)
			ix = instr(1,sx,"IDD_")
			if ix<> 0 and i<>0 then 
					TipCam(i,1)=3
					ed_sQue(i,1) = "1" & chr(9) & "1" & chr(9)
					for ia=2 to 31
						ed_sQue(i,1) = ed_sQue(i,1)  & ia & chr(9) & ia & chr(9)
					next	
			end if
						
		 			

' Campo Año
			sx= Ucase(name)
			ix = instr(1,sx,"IDY_")
			if ix<> 0 and i<>0 then 
					TipCam(i,1)=3
					ed_sQue(i,1) = "1900" & chr(9) & "1900" & chr(9)
					for ia=1901 to Year(now())
						ed_sQue(i,1) = ed_sQue(i,1)  & ia & chr(9) & ia &  chr(9)
					next	
			end if

' Campo UpLoad Imagen
			sx= Ucase(name)
			ix = instr(1,sx,"IMG_")
			if ix<> 0  then 
				TipCam(i,1)=7
			end if						
'response.Write "<br>" & i & name & TipCam(i,1)
		
End Sub


Sub VerAscii

response.write "<table>"
response.write "<tr>"
ix=0
for iy=1 to 255
    ix=ix+1
    response.write "<td>" & iy & ".." & chr(iy) & "</td>"
    if ix=16 then 
        response.write "<br></tr><tr>"
        ix=0
    end if    
next 
response.write "</tr>"
response.write "</table>"
					

end sub


'=====================================================================================
' Desplegar data 
' vector es el vector a paginar
' iPag la página a mostrar 
' iRegsPorPag el nº de registros deseado por cada página
'=====================================================================================
Sub ed_VerPag1(iRegsPorPag, ixPag, gData)

%>
<!-- Chequear si hay Data -->

<%	if ed_Ilof = 0 then 
    	ed_CalPar 5,ed_iCla,1,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	    ed_Botones ixPag, iPaginas, gData%>
	    <hr width="100%"  color="#cccccc" />	    
	    <%
	    if ed_sBus="" then
	        sx="No hay registros para Mostrar"
	    else
	        sx="No hay registros que cumpla con su criterio de búsqueda"
	    end if
	    %>
	    
		<br/><font face= "Verdana" size="3" color="#800000"><center><b>
		<%=sx %></b> </center> </font>
		
		<br/>	    
        <hr width="100%"  color="#cccccc" /><br/><br/>	    
<%		exit sub
	end if 
	
	Dim I, J 
	Dim iPaginas
	Dim iPagActual 'Total de páginas y la página que queremos mostrar
	Dim iTotal, iComienzo, iFin 'Total de registros, registro en que empezamos y registro en que terminamos

		

' Calcular total registros
	iTotal = UBound(gData,2)+1
	Ed_iMaxReg =UBound(gData,2)+1
	
' Calculo el numero de páginas
	iPaginas = (iTotal \ iRegsPorPag)
	if iTotal mod iRegsPorPag > 0 then
		iPaginas = iPaginas + 1
	end if
	
'Si no es una página válida, comienzo en la primera
	ixPag=cInt(ixPag)
	if ixPag < 1 then
		ixPag = 1
	end if
'Si es una página mayor al nº de páginas, comienzo en la última
	if ixPag > iPaginas then
		ixPag = iPaginas
	end if


'Calculo el índice donde comienzo:
	iComienzo = (ixPag-1)*iRegsPorPag
	
'Calculo el final y Si no tengo suficientes registros restantes, voy hasta el final
	iFin = iComienzo + (iRegsPorPag-1)
	if iFin > UBound(gData, 2) then
		iFin = UBound(gData, 2)
	end if

    
    ix=0
    sBac="#f1f1f1"
	if ed_iPas=4 then 
	    ix=1 
	    sBac="#ffffff"
    end if
%>

	<table border="0" width="100%" cellpadding="0" cellspacing="0"   align="center" id="Table8" style=" border:1px solid  #dddddd;  border-top-left-radius: 8px;border-top-right-radius:8px;" >
	<%iColSpan=ed_iNumCam+ed_pulsar(0,0) 
	  if ed_iRan=1 then iColSpan=iColSpan+1
	  if ed_iPas<>4 then %>
	
        <tr bgcolor="<%=ed_sBacCol1 %>" >
        <td colspan="<%=iColSpan%>">
    		<%if ed_iPas<>4 then%>
	    		<%  ed_Botones ixPag, iPaginas, gData%>	
		    <% end if %> 
		</td></tr>
 
	<%end if%>	
     </table>	
<!-- TItulo -->
    
   
   
  <div style=" overflow:Auto; width:100%"> 
    <% if ed_iAnc1="" then ed_iAnc1="100%" %> 
    <table border="<%=ix%>" width="<%=ed_iAnc1%>" cellpadding="0" cellspacing="1" bgcolor="<%=sBac %>"   align="center" id="Table4" >
   	<tr >
   	
<%   	
' #e4ecf7
	if ed_iRan=1 then %>
		<td align="center"  class="ed_ti1" background="/images/men.gif">
			Rank
		</td>
<%	end if
   	



   	for i=0 to ed_iNumCam-1
         iO=ed_sCampo(i,2)
		 if ed_iRep<>1 then
			if ed_iPas=4 and ed_sCampo(i,2)="2" then io=""
		 end if	
   		 if iO="1" or iO="2" Then

		 Else 

			sAli ="Center"	
			
%>		
			
			<td  height="30" align="<%=sAli%>"  class="ed_ti1" title="<%=ed_sCampo(i,5)%>" background="/images/men.gif">
				<% if sAli="left" then response.Write "&nbsp;&nbsp;" %>
				<% sx=ed_rs1.fields(i).name
					sxTit=sx
					'if i<>0 then CamTit sxTit
					if ed_sCampo(i,0)<>"" then sxTit=ed_sCampo(i,0)
					
				' Invertir Ordenamiento  
					if ed_iPas= 8 then
							if ed_iOrd = 0 then ed_iOrd = 1 else ed_iOrd = 0
					end if	

					
                    ed_CalPar 8,ed_iCla,ed_iPag,ed_sBus,i,ed_iOrd, ed_ifil,ed_iMp,ed_iMs

				' Invertir Ordenamiento  
					if ed_iPas= 8 then
							if ed_iOrd = 0 then ed_iOrd = 1 else ed_iOrd = 0
					end if	

				
				if ed_iPas<>4 then%>
						
					<a  href="<%=sPar %>" title="<%=ed_sCampo(i,5)%>" ><%=sxTit%>
					<% if i- ed_iCol=0 then %>
						<% if ed_iOrd=0 then %>
							<img src="\images\s0.gif" border="0">
						<% else%>
						    <img src="\images\s1.gif" border="0">
						<% end if%>
					<% end if %>
					</a>
				
				<% else %>
					<%=sxTit%>
				<%end if	%>
				
			</td>
		
<%		End if
	next

	if ed_pulsar(0,0)<>0 then
	    for ip=1 to ed_pulsar(0,0)%>
    	    <td  height="25" align="<%=sAli%>"  class="ed_ti1" >
    	    
	        </td>
	<%	    next
	 end if %>
	</tr>


	
<%
'#FEFEF9
' #9eb6ce
' #f7f7f7
' Mostrar Registros	
	ixL=0

	for i = icomienzo to ifin
        if ixL=0 then	%>
			<tr  height="30" class="ed_l2" ONMOUSEOVER="this.style.backgroundColor='#eeeeee' " ONMOUSEOUT="this.style.backgroundColor='#ffffff'" > 
		<%else	%>
			 <tr height="30" class="ed_l1" ONMOUSEOVER="this.style.backgroundColor='#eeeeee'" ONMOUSEOUT="this.style.backgroundColor='#ffffff'">
		<%end if
		

		ixL = not(ixL)
		
	
			if ed_Link<>"" then
				sP=sPro
				sPro=ed_link
			end if	
		
			ed_CalPar 2,gData(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
			if ed_Link<>"" then
			    ed_CalPar 1,gData(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
				sPro = sP
			end if	


' Escribir Ranking		
			if ed_iRan=1 then
				response.Write "<td>"
				response.Write i+1
				response.Write "</td>"
			end if
' Esctibir el detalle			
        if ed_iDet<>"" then
            sp=sPar
            ed_CalPar 1,gData(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
            sPar=sPar&"&ed_det=" &  gData(ed_iDet,i)
            %>
           
            
        <% ' sPar=sp
        end if
        
		for j=0 to ed_iNumCam-1
		    iO=ed_sCampo(j,2)
			if ed_iRep<>1 then
				if ed_iPas=4 and ed_sCampo(j,2)="2" then io=""
			end if	
		 if iO="1" or iO="2" Then 
		 Else 

			Select Case TipCam(j,1)
				case 2,3
					sAli ="left" 
					%>		
					<td   align="<%=sAli%>" title="<%=ed_sCampo(j,5)%>">
						<% if ed_iPas<>4 then 
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>"> 
								<%end if 
							end if		
							'response.Write gData(j,i) 
							if gData(j,i)<>"" then
								sCam=ed_sQue(j,1)
								iy=0
								do
									ix= instr(1,sCam,chr(9))
									sC=""
									sT=""
									if ix<>0 then
										sC=mid(sCam,1,ix-1)
										sCam=mid(sCam,ix+1)
										ix = instr(1,sCam,chr(9))
										sT=mid(sCam,1,ix-1)
								'	response.Write sC  & "----------" &gData(j,i)
										 if isnumeric(sc) then
										    iz = gData(j,i) - Sc
										    if iz=0 then response.write "&nbsp;&nbsp;" & ST  
										else
										    if gData(j,i) = Sc then response.write "&nbsp;&nbsp;" & ST  
										end if    
										sCam=mid(sCam,ix+1)
									end if
								loop until ix =0					
						'response.write "&nbsp;&nbsp;" & gData(j,i)& "GGGGGG"
							end if
				
				case 6
					sAli ="center" 
					%>		
					<td   align="<%=sAli%>" title="<%=ed_sCampo(j,5)%>">
						<% if ed_iPas<>4 then 
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>"> 
								<%end if 
							end if		
							if  gData(j,i)  then sx="Si" else sx="No"
							if gData(j,i)<>"" then
						       response.write "&nbsp;&nbsp;" & sx 
							end if
                		
				    %> 
				    
				<%
				case else
				    sAli ="left" 
				    if isNumeric(gData(j,i)) then  sAli="right"
				        
					
					if TipCam(j,0)=3 then sAli="center"
					if ed_rs1.fields(j).type="3" or ed_rs1.fields(j).type="2" then sAli ="center"
					if ed_rs1.fields(j).type="135"  then sAli ="center"
					%>
					<td   align="<%=sAli%>"  title="<%=ed_sCampo(j,5)%>" style="margin-right:13px; margin-left:13px">
					<%
						
						if ed_iPas<>4 then
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>">
								<%else %>
										 
								<%end if 		
						end if		
						if ed_iPas<>4 then response.Write " &nbsp;&nbsp;"
						if TipCam(j,1)=5 then
							
							response.Write String (8,"#")
						else%>
						     <% if isnull(gData(j,i)) then
						        else
						            if Len(gData(j,i))<50 then %>
							            <%=gData(j,i)%>
							            <% if ed_iPas<>4 then %>
							                <%="&nbsp;&nbsp;"%> 
							            <%end if %>
							        <%else %>      
							            <div style="width:450px;  margin-bottom:10px; margin-left:8px; border: solid 1px #ffffff; text-align:justify" >
							            <%=gData(j,i)%> 
							            </div>
							        <% end if %>
							    <% end if %>    
						<%end if	
			end Select	
%>			
				<%if ed_iPas<>4 then
				    if ed_iRep<>1 then%> 
					</a>
				<%  end if
				  end if%>	

				</td>
			
			
<%	    	End if
		Next

	    if ed_pulsar(0,0)<>0 then
	        'if ed_iPas<>4 then %>
	            <%for ip=1 to ed_pulsar(0,0)
	                sP=sPro
	                sPro=ed_Pulsar(ip,1)
	                ed_CalPar ed_iPas,gData(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs 
	                sPro=sP %>
	                <td  align="center" height="10" valign="middle" bgcolor="#ffffff"  >    	
                    <div class="ed_boton1"><a href="<%=sPar %>"   title = '<%=ed_Pulsar(ip,2) %>' >
			           <%=ed_Pulsar(ip,0)%></a></div>	
			        </td>    
	            <%next %>
	        
	    <%' end if
	    end if %>		
   		</tr>


 <%	Response.Flush
    Next%>
</table>
<br />
</div>

    <table border="0" width="100%" cellpadding="0" cellspacing="0"   align="center" id="Table1" style=" border-bottom: solid 1px #dddddd; border-left: solid 1px #dddddd; border-right: solid 1px #dddddd;  border-bottom-right-radius: 8px;border-bottom-left-radius: 8px;" >
   


<%if ed_iPas<>4 then%>   

    <%if ed_swSum<>"" then %>
        <tr><td colspan="<%=iColSpan%>" bgcolor="#ffffff" class="ed_l1" valign="middle" height="20" align="center">
            <%sx=ed_rs1.fields(ed_swSum).name
			  sxTit=sx
			  if ed_sCampo(ed_swSum,0)<>"" then sxTit=ed_sCampo(ed_swSum,0)%>
            
            <%="Total " & sxTit& ":=" & ed_iSum(ed_swSum) %>
	    </td></tr>
	 <%end if %>   


<%  if Ed_iMaxReg>ed_iRegPag then %>   
        <tr><td colspan="<%=iColSpan%>" class="ed_bac1" valign="middle" height="30" align="center">
	        <%ed_Paginar    %>
	
	</td></tr>
<% end if %>	
	<tr><td background="/images/men.gif"  style="font-family:verdana; font-size:8pt; font-weight:  bold; color:#666666; background-color:#ffffff; text-align:center" height="25">

       
       Total Registros:= <%=Ed_iMaxReg%> 

	
	</td></tr>
<% end if%>		
</table>
<%
	
		

	
	
'	Ed_VerPag ixPag, iPaginas, Ubound(gData,2)
if ed_ides=1 then   Vercampo	
End Sub

Sub Ed_Botones (ixPag, iPaginas, gData)

%>
    <table width="98%"  border="0" cellpadding="0" cellspacing="0" id="table9"  align="center" >
	    <tr >
		<td width="15%" align="center"  valign="middle" >
		    <%if Ed_iMaxReg>0 Then %>
		        <font face= "Verdana" size="1" color="#000000">
			        <% Response.Write("Página " & ixPag & " de " & iPaginas & "</b>")%><br />
			        
				</TD>
			    </font>
			<%end if %>	
		</td>
        <td width="15%" align="center"  valign="middle" >
	        <%if ed_sBotonC(0,0)<>"" then 
	            sp=sPro
	            sPro=ed_sBotonC(0,1)
	            ed_CalPar ed_iPas,0,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	            sPro=sP%>
	            <div class="ed_boton3"><a href="<%=sPar %>"   title = '<%=ed_sBotonC(0,3)%>' target="<%=ed_sBotonC(0,2) %>" >
	            <%=ed_sBotonC(0,0)%></a></div>	
	        <%end if %>
	     </td>		
	    <%ed_CalPar 7,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs%>
		<form action="<%=sPar%>"  method='post' >
		<td width="40%" align="center" height="50" valign="middle" class="ed_ti2">
			<input type='text' size="32" name ="bus"   value='<%=ed_sBus%>'  />
			<input  type='submit'  value ='Buscar' id="Submit2" name="Accion" />
			<a href="<%=sPar %>"   class="ed_ti2" style="text-decoration:underline" title = 'Ver todos los registros' >
	        Ver Todo</a>
		</td>
		</form>
	
<!--     de VerReg -->	
	    
	   

	    
	    <td width="30%"  valign="middle" align="right" >
	            
                <%if ed_Bot(2)<>"disabled" then
	                ed_CalPar 5,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs%>
		            <div class="ed_boton3" style="float:  right;"><a href="<%=sPar %>"   title = 'Añadir Registro' >
	                Añadir</a></div>
    			<%end if %> 	        
    			
	            <%if ed_Bot(1)<>"disabled" then
	                ed_CalPar 4,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs%>
	                <div class="ed_boton3" style="float:right;  vertical-align:middle"><a href="<%=sPar %>"   title = 'Exportar a Excel' >
	                Excel</a></div>
	           <%end if %> 	
	           <% for i=1 to 9
	                if ed_sBotonC(i,0)<>"" then 
	                    sp=sPro
	                    sPro=ed_sBotonC(i,1)
	                    ed_CalPar ed_iPas,0,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	                    sPro=sP%>
	                    <div class="ed_boton3" style="float: left"><a href="<%=sPar %>"   title = '<%=ed_sBotonC(i,3)%>' target="<%=ed_sBotonC(i,2) %>" >
	                    <%=ed_sBotonC(i,0)%></a></div>	
    	           <% end if 
    	           next%>	           
    	          
		    </td>

		    
		
	
		</tr>
	</table>
<%		
end Sub



'==========================================================================================
' Grabar Data encontrada en el formato
'==========================================================================================
Sub ed_GraDat(gData, SqlInp)
	

    for i=0 to ed_rs1.fields.Count -6
         sx=ed_rs1.Fields(i).name
         sGra(i)=request.Form(sx)
         'response.Write( "<br> Campo:=" & i & " Nombre:=" &   sx & " - Valor:=" & request.Form(sx))
    next   
       
'  Upload
'  ******
'*    mySmartUpload.Upload	
    if ed_sDirUp="" then ed_sDirUp="."
'*    intCount = mySmartUpload.Save(ed_sDirUp)

	
'*    for i=0 to ed_rs1.fields.Count -1
'*       sx=ed_rs1.Fields(i).name
'*       if TipCam(i,1)<>7 then
          '  Response.Write "<br> name:=" & ed_rs1.Fields(i).name & " Dat:=" & mySmartUpload.Form(sx).values
'*           sGra(i)=mySmartUpload.Form(sx).values
'*       else
           ' Response.Write "<br> sx:=" & sx & " File:=" & ed_rs1.Fields(i).name 
          '  Response.Write "<br>  Dat:=" &  mySmartUpload.Files(sx).Filename
'*            if ed_sCampo(i,2)="1" then
'*                sGra(i)=mySmartUpload.Form(sx).values
'*            else
'*                if mySmartUpload.Files(sx).Filename<>""  then
'*                    sGra(i)=mySmartUpload.Files(sx).Filename
'*                else
'*                    sGra(i)=mySmartUpload.Form(sx).values
'*                end if    
'*            end if    
'*        end if    
'*    next
	
' Validar Not Null
   ed_iErrG=0
   for i = 1 to ed_rs1.fields.Count-6	
        if sGra(i)="" and ed_sCampo(i,4)=1 then
            ed_sErr(i)="<br><font face= 'verdana' size='2' color='#ff0000' >&nbsp;&nbsp;<b> Error: Los campos marcados con * son obligatorio </b></font>"
            Response.Write "<br><font face= 'verdana' size='2' color='#ff0000' >&nbsp;&nbsp;<b>Falta Data en el campo:=" & ed_sCampo(i,0) & "</b></font>"
            ed_iErrG=1
        end if
   next 
   if ed_iErrG=1 then 
        Response.Write  "<br><font face= 'verdana' size='2' color='#ff0000' >&nbsp;&nbsp;<b> Error: Los campos marcados con * son obligatorio </b></font>"
        exit sub
    end if    

	
	Dim rsg
	set rsg = CreateObject("ADODB.Recordset")


	rsg.CursorType = 1
	rsg.LockType = 3
'	response.Write  "<br>1202 " &  SqlReg
	rsg.Open sqlReg,conexion

    sAcc=request.Form("Accion")
'response.write "<br> SAcc:=" & sAcc	
	    if sAcc="Añadir" then
			rsg.addnew
		end if

	if sAcc= "Guardar" or sAcc= "Añadir" Then
	
		for i = 1 to rsg.fields.count-6
	'	response.Write ("<br> 944 " & ed_rs1.fields(i).name) & " Grabar..." & sGra(i)
			if sGra(i)<> ""  then 
				'response.Write ("<br>a" & ed_rs1.fields(i).name & "..." & sGra(i))
				select case TipCam(i,1)
					case "6"
						if sGra(i) = "Verdadero" then sGra(i)=True
						if sGra(i) = "True" then sGra(i)=True
						if sGra(i) ="Falso" then sGra(i)= False
						if sGra(i) ="False" then sGra(i)= False
						
					case else

				end select
			'response.Write ("<br>954 i:=" & i & "  name:="  & ed_rs1.fields(i).name & "..." & sGra(i)) & "... len=" & len(sGra(i)) &   "  TipCam:=" & TipCam(i,1)
			    if ed_rs1.fields(i).Precision<>7 then
				    rsg(ed_rs1.fields(i).name)=sGra(i)
				else
				    sx=replace(sGra(i),",",".")
				    rsg(ed_rs1.fields(i).name)=sx
				end if    
			else 
			'response.Write ("<br>b...........	" & ed_rs1.fields(i).name & "..." & sGra(i)) 
				rsg(ed_rs1.fields(i).name)=null
			end if	
		next
		'   response.Write Request.ServerVariables("REMOTE_ADDR") & "..." & len(Request.ServerVariables("REMOTE_ADDR")) 
		rsg("IP")=Request.ServerVariables("REMOTE_ADDR")
		rsg("Fec_Ult_Mod")=DateAdd("n",+30,Now())
		If session("usu")<>"" then 	rsg("Usr")=session("usu") else rsg("Usr") = "Not User"
		rsg("idSession")=iSession
		'Response.Write("<script>alert('Registro Guardado');</script>") 
	elseif sAcc= "Eliminar" Then
		'response.Write "Pasoooooooooo por Auqi" 
		rsg("IP")=Request.ServerVariables("REMOTE_ADDR")
		rsg("Fec_Inactivo")=DateAdd("n",+30,Now())
		If session("usu")<>"" then 	rsg("Usr")=session("usu") else rsg("Usr") = "ed_GraDat"
		rsg("idSession")=iSession
		'Response.Write("<script>alert('Registro Eliminado');</script>") 
	end if	
	
	
	rsg.update
	rsg.close

end sub


'==========================================================================================
' Grabar Data encontrada en el formato
'==========================================================================================
Sub ed_EliReg(gData, SqlInp)

	Dim sGra(99)
	
	Dim rsg
	set rsg = CreateObject("ADODB.Recordset")


	rsg.CursorType = 1
	rsg.LockType = 3
	'response.Write SqlReg
	rsg.Open sqlReg,conexion

	if rsg.eof then
			'rsg.addnew
			'rsg.close
			'rsg.Open sqlReg,conexion
	
	Else
		rsg("IP")=Request.ServerVariables("REMOTE_ADDR")
		rsg("Fec_Inactivo")=DateAdd("n",+30,Now())
		If session("usu")<>"" then 	rsg("Usr")=session("usu") else rsg("Usr") = "ed_EliReg"
		rsg("idSession")=iSession
		rsg.update
	end if	
	rsg.close
end sub



Sub ed_VerPag2 (gData, iFor)
   
	if iFor=1 Then
		ed_CalPar 3,gData(0,0),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	else
	    ed_CalPar 3,"",ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	end if	 
%>	
	<!--table border="0" width="85%"  cellspacing="1"  cellpadding="0" align="center" id="Table2"  bgcolor="#c0c0c0" >
	<tr ><td class="ed_bac2"  -->
	<form action="<%=sPar %>"  method='POST' id="FormPag2" >	</form>
	<table border="0" width="90%"  cellspacing="1"  cellpadding="0" align="center" id="Table2" style="border: solid 1px #cccccc;border-radius:10px; padding:0px 0 10px 0; background-color:#f1f1f1" >
	<tr ><td   >
	
	<table border="0" width="100%" align="center">
	
	
	<tr >
	<td align="center">
	<br />
	
	<%ed_ilin=0
	for i=0 to ed_iNumCam2
' Imprimir texto
        if ed_sTit(ed_ilin,0)<>"" then
            sx=ed_sTit(ed_ilin,1)
            if sx<>"" then sx="class='" & sx & "'"
            %>
            <br />
            <div <%=sx%>><%=ed_sTit(ed_ilin,0) %></div>
            <br />
        <%  ed_sTit(ed_ilin,0)=""
        end if
        ic2=i
        if ed_sTit(ic2,2)="f" then  %>
            <div class="ed_div2">
        <%  ed_sTit(ic2,2)=""
        end if
       if ed_sTit(ic2,4)<>"" then  %>
            <div style="margin-top:-30px;padding-bottom:20px; margin-left:20px; text-align:left"><span style="font-size:12px; font-family:Verdana; color:#666666;  background-color:#F5F8FA; border-top : solid 1px #cccccc; padding:3px 3px 3px 3px; border-radius:5px"><%=trim(ed_sTit(ic2,4))%></span></div>
        <%end if        
 
'	
	    if iFor<>1 then ed_sCampo(0,2)=1
		sxTIT= ed_rs1.Fields(i).name 
        if ed_sCampo(i,0)<>"" then sxTit=ed_sCampo(i,0)
        
		sx=""
		if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
		if ed_sCampo(i,4)=1 then sxTIT=sxTIT & "<font face= 'verdana' color='#ff0000' ><b>*</b></font>"

' Campo que no se va a presentar		
		if ed_sCampo(i,2)="1" or ed_sCampo(i,2)="3" then 
		    if iFor=1 then%>
		        <input type="hidden" name="<%=ed_rs1.Fields(i).name%>" value='<%=gData(i,0)%>' id="Hidden2" form="FormPag2" />			
   		    <%else %>    
    	        <input type="hidden" name="<%=ed_rs1.Fields(i).name%>" value='<%=sx %>' id="Hidden1" form="FormPag2"/>	
	        <%end if %>
		<%else%>

<%
' Mostrar Nombre del campo           
            sb=""
            st=ed_sCampo(i,5)
            if ed_iDes=1 then 
                sb="border: 1px solid #ccdbe4;" 
                st=st &  " Campo=" & i & " s=" & ed_Formato(i,0) & " AncT=" & ed_Formato(i,3) & " M=" & ed_Formato(i,1)
            end if 
            iSal=i
            if isal<0 then iSal=0      
            if ed_Formato(iSal,0)="" then sFl="" else sFl="float:left;"%>
            <div style=" position: relative;<%=sb%>vertical-align:middle; padding: 0px 10px 0px 0px; <%=sFl%> margin-left:<%=ed_Formato(i,1)%>px; text-align:left">
    		<span style="<%=sb%>width:<%=ed_Formato(i,3)%>px;float:left;vertical-align:middle;padding:3px 0px 5px 0px;text-align:left; font-family:Verdana; font-size:12px; color:#000066"  title="<%=st %>" >
    		    <%=sxTIT%>:
	    	</span>
			<span title="<%=ed_sCampo(i,5)%>" style="<%=sb%>text-align:left; <%=sFl%>">
<%				
' Mostrar Campos				
			Select Case TipCam(i,1)%>
			
			<%case 1 ' textarea %>
				<%if iFor=1 then%>
				        <textarea rows="<%=TipCam(i,2)%>" name="<%=ed_rs1.Fields(i).name%>" cols="50" title="<%=ed_sCampo(i,5)%>"   <%=ed_sCampo(i,3)%> onKeyUp="return maximaLongitud(this,<%=TipCam(i,3)%>)" tabindex="<%=i %>"  class="ed_cam" maxlength ="<%=TipCam(i,3)%>" form="FormPag2"><%=gData(i,0)%></textarea>
				    <% else    
		    		     sx=""
						 if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
					     if ed_sCampo(i,8)<>"" then  sx=ed_sCampo(i,0)
						 if ed_iErrG=1 then sx=request.Form(ed_rs1.Fields(i).name)
					    %>
    					<textarea rows="<%=TipCam(i,2)%>" name="<%=ed_rs1.Fields(i).name%>" cols="50" title="<%=ed_sCampo(i,5)%>" onKeyUp="return maximaLongitud(this,<%=TipCam(i,3)%>)" tabindex="<%=i %>"  class="ed_cam" maxlength ="<%=TipCam(i,3)%>" form="FormPag2"><%=sx %></textarea>
					<%end if%> 
					</span></div>
			<%case 2 ' Combo
				if ed_sCampo(i,3)<> "readonly" then%>  
				<select size=""  name="<%=ed_rs1.Fields(i).name%>" id="<%=i%>"  title="<%=ed_sCampo(i,5)%>" tabindex="<%=i %>"  class="ed_cam" form="FormPag2">
				
				<%if ed_sCampo(i,4)<>1 then %>
				<option value="0" ></option>
				<%end if
 
				sCam=ed_sQue(i,1)
				iy=0
				do
					ix= instr(1,sCam,chr(9))
					'response.Write "len:=" & len(scam) & "opcion:=" & iy & "ix:=" & ix
					if ix<>0 then
						sC=mid(sCam,1,ix-1)
						sCam=mid(sCam,ix+1)
						ix = instr(1,sCam,chr(9))
						sT=mid(sCam,1,ix-1)
			
			
						response.Write("<option value=" & chr(34) & sC & chr(34))
						if iFor=1 then
							if isnumeric(Sc) then
						        iz = gData(i,0) - Sc
						        if Iz =0 then response.write " selected "
						    else
						        if Sc=gData(i,0) then response.write " selected "
						    end if  
						else
						    if ed_sCampo(i,8)<>"" then  
						         iz = gData(i,0) - Sc
						         if Iz =0 then response.write " selected "
						    else    
	                            ' Default
								if ed_iErrG=1 then ed_sCampo(i,1)=request.Form(ed_rs1.Fields(i).name)		
						        if ed_sCampo(i,1)<>"" then 
						            if isnumeric(ed_sCampo(i,1)) then
							            iz = Sc - ed_sCampo(i,1)
							            if Iz =0 then response.write " selected "
							        else
							            if Sc=ed_sCampo(i,1) then response.write " selected "
							        end if    
						        end if	
						    end if   
						end if    
						response.Write(">" & sT  &  "</option>")
						
						'response.Write  sC & "---" & ST & "Gdata " & gData(i,0) & "<br>"

						sCam=mid(sCam,ix+1)
					end if
				loop until ix =0
				
				
				else
' Mostrar readonly
				
				%>  
				<select size=""  name="<%=ed_rs1.Fields(i).name%>" id="Select3"  title="<%=ed_sCampo(i,5)%>"   tabindex="<%=i %>"  class="ed_cam" form="FormPag2">
				<!--option value="0" ></option-->
				
<% 
				sCam=ed_sQue(i,1)
				iy=0
				do
					ix= instr(1,sCam,chr(9))
					'response.Write "len:=" & len(scam) & "opcion:=" & iy & "ix:=" & ix
					if ix<>0 then
						sC=mid(sCam,1,ix-1)
						sCam=mid(sCam,ix+1)
						ix = instr(1,sCam,chr(9))
						sT=mid(sCam,1,ix-1)
			
			
						
						iz = gData(i,0) - Sc
						if Iz =0 then 
						    response.Write("<option value=" & chr(34) & sC & chr(34))
						    response.write " selected "
						    response.Write(">" & sT  &  "</option>")
                        end if    						


						sCam=mid(sCam,ix+1)
					end if
				loop until ix =0				
				end if
%>
				</select>
				</span></div>
				<% case 3,6 'Inddicador, Mes , Dia , Año%> 
				
				<select  size="1"  name="<%=ed_rs1.Fields(i).name%>" id="Select2"  title="<%=ed_sCampo(i,5)%>"  <%=ed_sCampo(i,3)%> tabindex="<%=i %>"  class="ed_cam" form="FormPag2">
				<%if ed_sCampo(i,4)<>1 then %>
				<option value="0" ></option>
				<%end if
 
				
              
				sCam=ed_sQue(i,1)
				iy=0
				do
					ix= instr(1,sCam,chr(9))
					'response.Write "len:=" & len(scam) & "opcion:=" & iy & "ix:=" & ix
					if ix<>0 then
						sC=mid(sCam,1,ix-1)
						sCam=mid(sCam,ix+1)
						ix = instr(1,sCam,chr(9))
						sT=mid(sCam,1,ix-1)
			
						response.Write("<option value=" & chr(34) & sC & chr(34))
						if iFor= 1 then
						    if Ucase(gData(i,0)) =UCase(sC) then response.write " SELECTED "
						else
						    if ed_sCampo(i,8)<>"" then  
						       if Ucase(gData(i,0)) =UCase(sC) then response.write " SELECTED "
						    else 
								if ed_iErrG=1 then ed_sCampo(i,1)=request.Form(ed_rs1.Fields(i).name)		
						        if ucase(ed_sCampo(i,1))= Ucase(sC) then  response.write " SELECTED"    
						    end if    
						end if    
						response.Write(">" & sT   & "</option>")
						'response.Write  sC & "---" & ST & "Gdata " & gData(i,0) & "<br>"

						sCam=mid(sCam,ix+1)
					end if
				loop until ix =0
				
				

%>
				</select>
				</span></div>
				<% case 4 ' Tipo ID_ read Only
				    if ifor=1 then
				    %>  
    					<input  type='text' size='<%=TipCam(i,3)%>' name ="<%=ed_rs1.Fields(i).name%>" <%=j %> value='<%=gData(i,0)%>' readonly   title="<%=ed_sCampo(i,5)%>"  tabindex="<%=i %>"/  class="ed_cam" form="FormPag2">
					<%else %>
	    				<input type='HIDDEN' size='<%=ed_rs1.Fields(i).DefinedSize%>' name ="<%=ed_rs1.Fields(i).name%>"    value='' readonly ID="HIDDEN6" name="Text1" title="<%=ed_sCampo(i,5)%>" tabindex="<%=i %>"  class="ed_cam" form="FormPag2"/>
					<%end if %>
					</span></div>
				<% case 5 ' Password
				    if iFor=1 then%>  
					    <input  type='password' size='<%=ed_rs1.Fields(i).DefinedSize%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=gData(i,0)%>'  <%=ed_sCampo(i,3)%> title="<%=ed_sCampo(i,5)%>" tabindex="<%=i %>"   class="ed_cam" form="FormPag2"/>
					<%else 
				        sx=""
		                if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)%> 
					    <input  type='password' size='<%=ed_rs1.Fields(i).DefinedSize%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>'  <%=ed_sCampo(i,3)%> title="<%=ed_sCampo(i,5)%>"  tabindex="<%=i %>"  class="ed_cam" form="FormPag2"/>										
					<%end if %>
					</span></div>
				<% case 66  ' Indicador
				    sx="No"
				    if gData(i,0)=true then sx="Si"
				    %> 
				    <input type="checkbox" name ="<%=ed_rs1.Fields(i).name%>"  value="<%=sx %>"
				     <%if gData(i,0)=true then response.write " checked"%>
				       title="<%=ed_sCampo(i,5)%>" />
				    <!--input type="checkbox" name ="<%=ed_rs1.Fields(i).name%>" <%=j %> value='1'  title="<%=ed_sCampo(i,5)%>" /><%=gData(i,0)%> -->
				   </span></div> 
				<% case 7  'Carga imagen
				
                    if iFor=1 then
                       sx=gData(i,0)  
                    else
                    	sx=""
                    end if%>
                     </span>
                     <br />
                <div style="text-align:center">
                    <br />
    	            <img src="<%=sx %>" align="middle" border="0"  style=" width:200px; height:200px; border-radius:5px; border:solid 1px #cccccc" alt="" id="ImgFact"/>
    	        <br/>
  	    	    <input type="file" id="fileElem" accept="image/*" onchange="upload('<%=ed_rs1.Fields(i).name%>')" form="FormPag2" />
  	    	    <br />
  	    	    
	        	<button id="fileSelect" class="ed_boton2">Cargar</button><br />
    	         <input  type='text' size='<%=TipCam(i,3)%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>' <%=ed_sCampo(i,3)%> 
    	         title="<%=ed_sCampo(i,5) %>" id="<%=ed_rs1.Fields(i).name%>" tabindex="<%=i %>" class="ed_cam" form="FormPag2"/>
	        	</div>
	        	</div>
		       
		
<script>

function click(el) {
  // Simulate click on the element.
  var evt = document.createEvent('Event');
  evt.initEvent('click', true, true);
  el.dispatchEvent(evt);
}

document.querySelector('#fileSelect').addEventListener('click', function(e) {
  var fileInput = document.querySelector('#fileElem');
  //click(fileInput); // Simulate the click with a custom event.
  fileInput.click(); // Or, use the native click() of the file input.
}, false);

</script>	                    
                      
			<%case 8 ' Fecha
                if iFor=1 then
                    sx=gData(i,0)  
                else
                    sx=""
    				if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
    				if ed_sCampo(i,8)<>"" then sx=sGra(i)
					if ed_iErrG=1 then sx=request.Form(ed_rs1.Fields(i).name)
                end if
                if ed_sCampo(i,3)="readonly" then %>	
                    <input  type='text' size='10' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>' <%=ed_sCampo(i,3)%>	/>
                <%else %>
				    <input  type='text' size='10' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>'  
				    onclick="ddate_click('<%=ed_rs1.Fields(i).name%>',this.value);"	
				    title="<%=ed_sCampo(i,5) %>" id="<%=ed_rs1.Fields(i).name%>" tabindex="<%=i%>" class="ed_camf" form="FormPag2"/>
				    </span>
				    </div>
				    
				    <div id="Div<%=ed_rs1.Fields(i).name%>" style="visibility:  hidden;height:1; width:400px" class="ed_bac2" ></div>
				<%end if    
					
            case 9 ' Color
                   if iFor=1 then
                       sx=gData(i,0)
                     
                    else
                    	sx=""
    					if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
    					if ed_sCampo(i,8)<>"" then sx=sGra(i)
					
                    end if   
                    %>	
					    <input  type="color"  name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>' <%=ed_sCampo(i,3)%>  
						    
					     title="<%=ed_sCampo(i,5) %>" id="Text1" tabindex="<%=i %>" class="ed_cam" form="FormPag2" />
					 </span></div> 
<%					                     
			case else
                    if iFor=1 then
                       sx=gData(i,0)  
                    else
                    	sx=""
    					if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
    					if ed_sCampo(i,8)<>"" then sx=sGra(i)
						if ed_iErrG=1 then sx=request.Form(ed_rs1.Fields(i).name)
						
                    end if   
                    %>	
					    <input  type='text' size='<%=TipCam(i,3)%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>' <%=ed_sCampo(i,3)%>  
					    
					
					<% if TipCam(i,0)=1 then %>
					    onkeydown="valnum(this.value, 'campo<%=i%>');" onkeyup="valnum(this.value, 'campo<%=i%>');" 
					    
					<% end if%>
                    <% if TipCam(i,0)=2 then %>
					    maxlength ="<%=TipCam(i,3)%>"
                        onblur="valtxt(this.value, 'campo<%=i%>');" 						
					 <% end if%>   					    
					     title="<%=ed_sCampo(i,5) %>" id="campo<%=i%>" tabindex="<%=i %>" class="ed_cam" form="FormPag2"/>
					</span></div>
			<% end select %>	
				  
				
				    <%if ed_Formato(iSal,0)="" then 
				       ed_iLin=ed_iLin+1%>
				      <div style="height:5px"></div>
				    <%end if   
                       if ed_sTit(ic2,3)="f" then%>
                            </div class='ed_div2'>
                            <%ed_sTit(ic2,3)=""
                        end if				      
 
    	End if
 	next %>    	
		             <!--/div--> 
  	    	    </td>
    		    </tr>  
        	    		 
	        </table>
	    </td></tr>
	
	
<!-- Botones del Formato -->
	<tr ><td height="30" align="center" >
	    
		<table  border="0" cellpadding="0" cellspacing="0"  align="center" height="30"  class="ed_tab2">
			<tr valign="middle " >


                <%if iFor=1 then %>		
					<td width="10"% align="right" height="10" valign="middle" >
					<%if ed_Bot(4)<>"disabled" then%>
							<input type='submit'  value ='Eliminar'  ID="Submit12" NAME="Accion" tabindex="44" <%=ed_Bot(4)%> class="ed_boton2" form="FormPag2">
					<% end if %>
					</td>
					<td width="70%" align="center" height="10" valign="middle""></td>
					<td width="10%" align="center" height="10" valign="middle" >
					<%	if ed_Bot(3)="disabled" then
						else %>
						<input type='submit'  value ='Guardar' ID="Submit6" NAME="Accion" tabindex="1" class="ed_boton2" form="FormPag2">			
						<% end if %>
					</td>	

                <%else%>				
                    <td width="10%" align="center" height="10" valign="middle"">	</td>
					<td width="70%" align="center" height="10" valign="middle"">	</td>
					<td width="10"% align="right" height="10" valign="middle"  >
					
						<input type='submit' VALUE ='Añadir' ID="Submit5" NAME="Accion" tabindex="2" class="ed_boton2" form="FormPag2">			
					</td>
                <%end if%>								
			
            <%ed_CalPar 1,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs%>
			<td width="10%" align="center" height="10" valign="middle"   >
                <div class="ed_boton1"><a href="<%=sPar %>"   title = 'Volver a la pagina principal' >
			        Volver</a></div>	
			</td>
 	        </tr>
	        <%
                sPro=sP
               %>
		</table>
	</td></tr>
	</table>


	
<%
 if ed_ides=1 then     Vercampo
end Sub


Sub Ed_OutExcel
	 
     'exit sub
     Response.ContentType = "application/vnd.ms-excel"
     Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
     'Response.AddHeader "Content-disposition","attachment; filename=tem.pdf"
    ' Response.ContentType = "application/msword"
     'Response.ContentType = "application/pdf"
      'response.ContentType ="application/ms-powerpoint"
      'response.ContentType ="application/zip"
     
End Sub

Sub ed_Paginar 
    if Ed_iMaxReg<ed_iRegPag then exit sub
%>
   <table border="0" cellspacing="0" cellpadding="0"  align="center"  class="ed_bac1" >
       <tr>
        
           
     <% 
     
     
     
' Calcular total registros
	iTotal = Ed_iMaxReg
	
' Calculo el numero de páginas
	iMaxPag = (iTotal \ ed_iRegPag)
	if iTotal mod ed_iRegPag > 0 then
		iMaxPag = iMaxPag + 1
	end if
%>      
      
	<td   align="right" class="ed_pag">
             <%ed_CalPar ed_iPas,ed_iCla,1,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs%>
            <a href= "<%=sPar%>"  onMouseOver="window.status='';return true" onMouseOut="window.status=' ';return false" title="">I</a>
    </td>


	<td   align="right" class="ed_pag">
     <%		ix=ipPag-1
            if ix<1 then ix =1
            ed_CalPar ed_iPas,ed_iCla,ix,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
             %>
             
            <a href= "<%=sPar%>"  onMouseOver="window.status='';return true" onMouseOut="window.status=' ';return false" title="">< Anterior</a>
    </td>
            

            
   
       <%   i1=ed_iPag-5
       		if i1<1 then i1=1
       		i2=i1 +9
            if i2> iMaxPag then i2=iMaxPag
        
            for i=i1 to i2
                 
                ed_CalPar ed_iPas,ed_iCla,i,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
 				ix=i-ed_iPag
				sx="ed_pag"
				if ix=0 then sx ="ed_pag1"
				%>
                  
                    
               <td   class="<%=sx%>" width="5">
                    <a href= "<%=sPar%>"  onMouseOver="window.status='';return true" onMouseOut="window.status=' ';return false" title=""><%=i%></a>
				</td>
                  
            <%next %>
       
 
			<%' ix=iMaxPag-ipPag 
           ' if ix<>0 then
                ix=ed_iPag+1
                if ix>iMaxPag then ix = iMaxPag
                ed_CalPar ed_iPas,ed_iCla,ix,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
                 %>
           		<td  align="right" class="ed_pag">
            		<a href= "<%=sPar%>"  onMouseOver="window.status='';return true" onMouseOut="window.status=' ';return false" title="">Siguiente ></a>
            	</td>	
            
        <%      
           ' end if
         %>          
   
      
       
	<td   align="right" class="ed_pag">
     <%		
            ed_CalPar ed_iPas,ed_iCla,iMaxPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
             %>
            <a href= "<%=sPar%>" onMouseOver="window.status='';return true" onMouseOut="window.status=' ';return false" title="">F</a>
    </td>       
       </tr>

    </table>   
    
          

 	    

<% End Sub 

Sub MenuDet(xSql,sMosTot)

	Dim rs
	Dim gMenFil
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
	rs.Open xSql,conexion
	gMenFil=rs.GetRows
	rs.close
	'cGris1="#800000"
	iAncMen=ianc*98/100
'	iwid=iAncMen/ed_iMAxMen
	%>
    
	

<br>    
<table   border="0" align="left" CELLSPACING="10" CELLPADDING="0" bgcolor="#ffffff"  id="Table5">
	<tr height="28">
<%
    if ed_iFil="" then ed_iFil=0
    ixl=0
    
    if sMosTot="n" and ed_iFil=0 then ed_iFil = gMenFil(0,0)
    for i=0 to ubound(gMenFil,2)
  
        ed_CalPar ed_iPas,ed_iCla,1,"",ed_iCol,ed_iOrd, gMenFil(0,i),ed_iMp,ed_iMs
          
        ix=ed_iFil-gMenFil(0,i)
        if ix=0 then%>
            <td  style="border: 1px solid #ccdbe4; background:#F5F8FA ; width:180px;float:left;vertical-align:middle;padding:5px 0px 5px 0px; text-align:center; vertical-align:middle "  title='<%=gMenFil(1,i)%> ' class="ed_md2" >
            
               <%= "" & gMenFil(1,i) & "" %>
            
		    </td>  
         <%else %> 
            <td style="border: 1px solid #ccdbe4;background:#ffffff ;width:180px;float:left;vertical-align:middle;padding:5px 0px 5px 0px; text-align:center "  title='<%=gMenFil(1,i)%> ' class="ed_md1">

                <a href="<%=sPar& "&prg=" & gMenFil(0,i)%>"  >
                <%="" & gMenFil(1,i)%></a>
            </td>     
         <%end if %>
 
    <%next%>
 
    
    </tr>
 
 
    </table>
 
<%
	

End sub

Sub ed_vCombo


    dim rst    
    
	set rst = server.CreateObject("ADODB.Recordset")
	rst.CursorType = 1
	rst.LockType = 1

%>
    <table width="<%=ed_sCombo(i,4)%>" cellspacing="1"  bgcolor="#ffffff" align="left" >
<%    	
	for i=1 to ed_iCombo
		'if isNull(ed_sCombo(i,3)) or ed_sCombo(i,3) = "" then
		'	'response.write	"<br>2339 NO:=" & ed_sCombo(i,3)    	
		'else
		'	response.write	"<br><br>2338 " & ed_sCombo(i,1)    	
		'	response.write	"<br>2339 " & ed_sCombo(i,3)    	
		'	rst.filter = ed_sCombo(i,3)
		'End if
		'response.write "<br>2352 ed_sCombo(i,1):= "  & ed_sCombo(i,1)
		rst.open ed_sCombo(i,1),conexion

	    dim gX
        gX=rst.getrows
        'response.write "<br>2357 i:= "  & i
		ed_sPar(i,1)=gX(0,0)
        if isnull(ed_sPar(i,0)) or ed_sPar(i,0)="" then 
            if ed_sCombo(i,2)<>"" then 
                ed_sPar(i,0)=ed_sCombo(i,2)
             else   
                ed_sPar(i,0)=gX(0,0)
             end if   
           ' response.write "<br>klasjjjjjjjjjjjjjjjjjjj" & ed_sPar(i,0)
        end if    
        rst.close   

       %>
        
   
	    <tr>
	        <td width="<%=ed_sCombo(i,5)%>" class="ed_dim41">&nbsp;&nbsp;<%=ed_sCombo(i,0) %>:</td>
	        <td width="<%=ed_sCombo(i,6)%>" style="font-size:  10px; background-color:#ffffff; color:#000000;">
	        <%if ed_iPas<> 4 then %>
	    	    <select size="1" name="per" id="Select1"   onchange ="location.href=this.options[this.selectedIndex].value"  style="width:100%; font-size:12px; font-family:Verdana; padding:3px 0 3px 0 ">
            <%  if ed_sCombo(i,2)<>"" then   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=ed_sCombo(i,2)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP

            %>
        			<option value="<%=sPar%>"  <% if ed_sPar(i,0) =ed_sCombo(i,2) then response.Write"selected" %> style="width:100%; font-size:12px; font-family:Verdana; padding:5px 0 5px 0 " >
				    <%="[" & ed_sCombo(i,2) & "]"%>
			        </option>

    		<%  end if
    		    wO=0

    		    for j=0 to ubound(gX,2)
    		   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=gx(0,j)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP
    	    	   
				   if isnumeric(ed_sPar(i,0)) then
						ix=ed_sPar(i,0) -gX(0,j)
				   else
					    ix=1
					    if ed_sPar(i,0) =gX(0,j) then ix=0
				   end if
     	    %>
			    <option value="<%=sPar%>"  
			    <% 
			    if ix=0 then 
			        response.Write"selected" 
			        wO=1
			    end if    
			        %>  >
				    <%=gX(1,j)%>
			    </option>
			    <%next
			     if wO=0 then 
			        if ed_sCombo(i,2)<>"" then 
                        ed_sPar(i,0)=ed_sCombo(i,2)
                    else   
			            ed_sPar(i,0)=ed_sPar(i,1) 
			        end if    
			     end if%>
	        </select>
	       <% else 
	            %>
	                <%=ed_Spar(i,0) %>
	            <%
	          end if %>
	    </td> 
	   </tr>
	             
<%	Next%>
    </table>
<%    
            		       
End sub

Sub MenuFil(ed_iMaxMen, sMosTot)

	Dim rs
	Dim gMenFil
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
'	response.Write "aaaaaaaaaaa" & ed_SqlFil
	'exit sub
	rs.Open ed_sqlfil,conexion
	gMenFil=rs.GetRows
	rs.close
	'cGris1="#800000"
	iAncMen=ianc*98/100
	iwid=iAncmen/ed_iMAxMen
	%>
    
	


<table width="100%"  border="0" align="center" CELLSPACING="1" CELLPADDING="0" bgcolor="#dddddd"  id="Table3">
	<tr height="28">
<%
    if ed_iFil="" then ed_iFil=0
    ixl=0
    
    sMosTot="s"
    if lcase(sMosTot)="s" then   
' Todos

        ed_CalPar ed_iPas,ed_iCla,1,"",ed_iCol,ed_iOrd, "",ed_iMp,ed_iMs
        if ed_iFil=0 then%>
            <td  width="<%=iwid %>" class="ed_ms2" bgcolor="#ffffff" background="/images/men.gif">
                <%= "- Sin Filtro -"%></td>
        <%else %> 
            <td  width="<%=iwid %>"  class="ed_ms1" bgcolor="#ffffff" background="/images/men.gif">  
            <a href="<%=sPar %>" title='Todas las Opciones'>
            <%="- Sin Filtro -"%></a></td>
        <%end if
        ixl=ixl+1    
    end if
    
    if sMosTot="n" and ed_iFil=0 then ed_iFil = gMenFil(0,0)
    for i=0 to ubound(gMenFil,2)
  
        ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, gMenFil(0,i),ed_iMp,ed_iMs
          
        ix=ed_iFil-gMenFil(0,i)
        if ix=0 then%>
            <td  width="<%=iwid %>"  class="ed_ms2" bgcolor="#ffffff" background="/images/men.gif" title='<%=gMenFil(1,i) %>'>
            <%= "" & gMenFil(1,i) & "" %></td>
         <%else %> 
            <td  width="<%=iwid %>" class="ed_ms1" bgcolor="#ffffff" background="/images/men.gif" title='<%=gMenFil(1,i) %>'>  
                <a href="<%=sPar %>"  >
                <%="" & gMenFil(1,i)  %></a></td>
         <%end if %>
        
        <%
        
        next
           
        %>
    </tr>
 
 
    </table>
 
<%
	

End sub

Sub ed_Menu 
	Dim ed_rs
	Dim ed_gMenu
	set ed_rs = CreateObject("ADODB.Recordset")
	ed_rs.CursorType = 1
	ed_rs.LockType = 1


    	
	sql = "Select * FROM SS_S_Menu"
    sql= sql & " WHERE ((id_Perfil=" & iPerUsu & ")AND (Fec_Inactivo is  Null) AND (num_Padre=0)) "
    sql= sql & " ORDER BY Orden "
	
'	response.Write "<br> 1778" & sql
	'exit sub
	ed_rs.Open sql, conexion
	ed_gMenu=ed_rs.GetRows
	ed_rs.close
	if ed_iMp="" then ed_iMp=ed_gMenu(0,0)

	 
	
%>
   <table width="100%" class="ed_tab1" cellspacing="0" height="35" border="0" align="center" >
        <tr ><td colspan="3" height="5"></td>
        <tr >
           <td width="8" height="30" ></td> 
           <%  ' isw=0
                for i=0 to Ubound(ed_gMenu,2)
                '  if ed_gMenu(5,i)="" then im1=ed_gMenu(0,i) 
                  if ed_gMenu(6,i)="" then ed_gMenu(6,i)=ed_gMenu(1,i)
                  ix=ed_gMenu(0,i)-ed_iMp
                  if ix <> 0 then
                    ' isw=1
                   %>
                    <td class="ed_mp1" title = "<%=ed_gMenu(6,i)%>"  >
                          <%
                          spxx=sPro
                          sPro= ed_gMenu(5,i)
                          ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,ed_gMenu(0,i),""
                          if ed_gMenu(9,i)<>"" then sPar=sPar &ed_gMenu(9,i)
                        
                          sPro=spxx%>
                          <a href="<%=sPar%>"   target="<%=ed_gMenu(8,i)%>" title = "<%=ed_gMenu(6,i)%>" ><%=ed_gMenu(1,i) %></a>
                            
                        </td>       
                <%else%>    
                      <td class="ed_mp1" bgcolor="#ffffff" align="center" title = "<%=ed_gMenu(6,i)%>">
                        <%=ed_gMenu(1,i)%>
                      </td>
                <%end if
                  ed_iAnc=ed_iAnc + 16
                next %> 
             
             <td width="<%=100-ed_iAnc%>%"></td>

         </tr>     
     </table>
<%   


	Dim ed_rss
	Dim ed_gMenu2
	set ed_rss = CreateObject("ADODB.Recordset")
	ed_rss.CursorType = 1
	ed_rss.LockType = 1
	
	sql = "Select * FROM SS_S_Menu"
    sql= sql & " WHERE ((id_Perfil=" & iPerUsu & ") AND (Fec_Inactivo is  Null) AND (num_Padre=" & ed_iMp & " )) "
    sql= sql & " ORDER BY Orden "
	
	'response.Write "aaaaaaaaaaa" & sql & " isw:=" & isw
	'exit sub
	ed_rss.Open sql, conexion
	ed_gMenu2=ed_rss.GetRows
	ed_rss.close
	if ed_iMs="" then 
	    ed_iMs=ed_gMenu2(0,0)
	else
	    'if isw=0 then   ed_iMs=ed_gMenu2(0,0)
	end if    
	%>


   <table width="100%" bgcolor="#c0c0c0" cellspacing="0" height="35" border="0" align="center">
          <tr >
           <% for i=0 to Ubound(ed_gMenu2,2)
                  
                  if ed_gMenu2(6,i)="" then ed_gMenu2(6,i)=ed_gMenu2(1,i)
                  iwid=len(ed_gMenu2(1,i))
                  iwid=iwid*10*2
                  ix=ed_iMs
                  ix=ed_gMenu2(0,i)-ix
                  if ix <> 0 then %>
                 
                    <td class="ed_ms1" width="<%=iWid%>" background="/images/men.gif"   title = "<%=ed_gMenu2(6,i)%>">
                          <%
                           spxx=sPro
                          sPro= ed_gMenu2(5,i)
                       '   if sPro="" then sPro=spxx
                          ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,ed_iMp,ed_gMenu2(0,i)
                           if ed_gMenu2(9,i)<>"" then sPar=sPar &ed_gMenu2(9,i)
                          sPro=spxx
                            %>
                            <a href="<%=sPar%>"   target="<%=ed_gMenu2(8,i)%>" title = "<%=ed_gMenu2(6,i)%>"> <%=ed_gMenu2(1,i)%></a>
                            
                        </td>       
                 <%else%>    
                     <%if ed_gMenu2(1,i)<>"#" then %>
                        <td class="ed_ms2" width="<%=iWid%>"  background="/images/men.gif" title = "<%=ed_gMenu2(6,i)%>" > 
                            <%=ed_gMenu2(1,i)%>
                        </td>
                     <%else %>  
                       <td class="ed_ms2" width="<%=iWid%>"  background="/images/men.gif" title = "<%=ed_gMenu2(6,i)%>" > 
                            <%=ed_gMenu2(1,i)%>
                        </td>                      
                     <%end if %>
                  <%end if
                   ed_iAnc=ed_iAnc + 12
                next %> 
         </tr>     
     </table>
<% 
  response.flush 
end Sub

Sub ed_MenVer 
' Abrir Perfil
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
	
    sql = "SELECT  id_Perfil, Perfil, link_acceso, ocultar, idGrupo "
    sql = sql & " FROM  ss_u_Perfil "
	sql = sql & " WHERE (((id_Perfil)=" & iPerUsu & ") AND ((Fec_Inactivo) Is Null)) "
'response.write "<br>2196 sql:=" & sql    	
	rs.Open sql,conexion
	e_gPerUsu=rs.GetRows
	rs.close
	set rs = nothing
	
' Abrir Menu
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
	
    sql = "SELECT  idNivel, Menu, Link , Txt_Tips, Target, Parametros, id_menu, idGrupo,ind_Emergente, Ind_Activo, ind_oculto "
    sql = sql & " FROM  ss_U_Menu "

	sOcultar=e_gPerUsu(3,0)
    sql = sql & " WHERE ((IdGrupo)=" & e_gPerUsu(4,0) & ") AND ((IdEmpresa)=1) AND ((Fec_Inactivo) Is Null) "
	if sOcultar<>"" then sql = sql & " AND (id_Menu NOT IN (" & sOcultar& "))"    
   
    sql = sql & " ORDER BY Orden; "
'response.write "<br>2196 sql:=" & sql    
	rs.Open sql,conexion
	'e_gMenu=rs.GetRows
	'rs.close
	'rs.Open sql,conexion 
 %>

    <table width="100%"  cellspacing="0"   cellpadding="0" border="0" align="center"   bgcolor="#fffff0">
    <tr ><td  bgcolor="#0059B2"  >
        <ul id="nav1">
         <% spxx=sPro
            sPro= rs("Link").value 
            if sPro="" then sPro=spxx
            ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("idnivel").value, rs("Id_Menu").value
            if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
            sPro=spxx   
         %>
         <!--li><a href="<%=sPar%> " title="<%=rs("txt_tips").value%>"> <%=rs("Menu").value%></a-->

        <%iNiv=rs("idNivel").value
        in1=rs("idNivel").value
        Do While NOT rs.EOF
          
         '   if i>ubound(e_gMenu,2) then exit do
       ' For i=1 to  ubound(e_gMenu,2)
            if ed_Ims<>"" then 
                iy=ed_iMs-rs("id_Menu").value
            else
                iy=1
            end if        
            if iy=0 then ed_sMenSec=rs("Menu").value
        %>
            <% ix= iNiv-rs("idNivel").value
                i1=iNiv-rs("idNivel").value
               if ix=0 then %></li>
             <%else%>
                <%if ix<0 then %><ul >
                <%else 
                    if ix=2 then%> 
                        </li></ul></ul></li>  
                    <% else %>    
                        </li></ul></li>  
                    <%end if %>
                <%end if %>
              <%end if
                         spxx=sPro
                          sPro= rs("Link").value
                          if sPro="" then sPro=spxx
                          ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("idNivel").value,rs("id_Menu").value
                           if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
                          sPro=spxx                
             'Sub ed_CalPar (xPas,xCla,xPag,xBus,xCol,xOrd,xFil,xMp,xMs)
                
                 
                if rs("Menu").value<>"#" then
                    if rs("Ind_Activo").value=false then%>
                        <li><div class="MnuIna" ><%=rs("Menu").value%></div>
                    <%else%>
						 <%if rs("Ind_Emergente").value=false then%>
							<li><a href="<%=sPar%>" title="<%=rs("txt_tips").value%>"  >
							<%if rs("idNivel").value = in1 then%>
								<img src="\images\dolar.png" alt="<%=rs("txt_tips").value%>" width="60" /><br />
							<%end if %>
						 <%else%>
							<li><a href="<%=sPar%>" title="<%=rs("txt_tips").value%>"  >
							<%if rs("idNivel").value = in1 then%>
								<img src="\images\dolar.png" alt="<%=rs("txt_tips").value%>" width="60" /><br />
							<%end if %>
                        <%end if %>
                        <%=rs("Menu").value%></a>
                    <%end if %>
                <%else%>  
                    <li><div class="MnuRaya"></div>
                    
                <%end if %>
        <% iNiv = rs("idNivel").value
            rs.MoveNext
              i=i+1
        Loop
        
        rs.close
         %>
        </ul></li></ul>
    </td>
    </tr>
    </table>


<%end Sub 

Sub ed_MenPri
' Abrir Perfil
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1 
    sql = "SELECT  id_PerfilUsuario, PerfilUsuario, link_acceso, ocultar, idGrupo , mostrar, Ind_Cliente, Ind_Operaciones"
    sql = sql & " FROM  ss_u_PerfilUsuario "
	sql = sql & " WHERE (((id_Perfilusuario)=" & iPerUsu & ") AND ((Fec_Inactivo) Is Null)) "
'response.write "<br>2196 sql:=" & sql    	
	rs.Open sql,conexion
	e_gPerUsu=rs.GetRows
	rs.close
	set rs = nothing
	session("perfcli") = e_gPerUsu(6,0)
	session("perfope") = e_gPerUsu(7,0)

	
' Abrir Menu
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1 
	
    sql = "SELECT  idNivel, Menu, Link , Txt_Tips, Target, Parametros, id_menu, idgrupo, Ind_Activo, ind_Oculto "
    sql = sql & " FROM  ss_u_Menu "
	
	sOcultar=e_gPerUsu(3,0)
	sMostrar=e_gPerUsu(5,0)
	ed_iGrupo=e_gPerUsu(4,0)
'response.write "<br>ed_igrupo=" & ed_Igrupo	
    sql = sql & " WHERE ((IdGrupo)=" & e_gPerUsu(4,0) & ") AND ((IdEmpresa)=1) AND ((Fec_Inactivo) Is Null) "
	if sOcultar<>"" then sql = sql & " AND (id_Menu NOT IN (" & sOcultar& "))"    
	if sMostrar<>"" then sql = sql & " AND (id_Menu IN (" & sMostrar& "))"    
	

    sql = sql & " ORDER BY Orden; "
'response.write "<br>2196 sql:=" & sql    
	rs.Open sql,conexion
	'e_gMenu=rs.GetRows
	'rs.close
	'rs.Open sql,conexion
 %>

    <table width="100%"  cellspacing="0"   cellpadding="0" border="0" align="center"   bgcolor="#ffffff">
    <tr ><td  bgcolor="#ffffff"  >
        <ul id="nav">
         <% spxx=sPro
         
            sPro= rs("Link").value 
            if sPro="" then sPro=spxx
            ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("idnivel").value, rs("Id_Menu").value
            if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
            sPro=spxx   
         %>
         <!--li><a href="<%=sPar%> " title="<%=rs("txt_tips").value%>"> <%=rs("Menu").value%></a-->

        <%iNiv=rs("idNivel").value
       
        i=0
        Do While NOT rs.EOF
            if rs("ind_oculto").value=true then
            else
        
            if ed_Ims<>"" then 
                iy=ed_iMs-rs("id_Menu").value
            else
                iy=1
            end if        
            if iy=0 then ed_sMenSec=rs("Menu").value
        %>
            <% ix= iNiv-rs("idNivel").value
               if ix=0 then %></li>
             <%else%>
                <%if ix<0 then %><ul >
                <%else 
                    if ix=2 then%> 
                        </li></ul></ul></li>  
                    <% else %>    
                        </li></ul></li>  
                    <%end if %>
                <%end if %>
              <%end if
                         spxx=sPro
                          sPro= rs("Link").value
                          if sPro="" then sPro=spxx
                          ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("idNivel").value,rs("id_Menu").value
                           if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
                          sPro=spxx                
             'Sub ed_CalPar (xPas,xCla,xPag,xBus,xCol,xOrd,xFil,xMp,xMs)
                
                
                if rs("Menu").value<>"#" then
                    if rs("Ind_Activo").value=false then%>
                        <li><div class="MnuIna" ><%=rs("Menu").value%></div>
                    <%else%>
                        <li><a href="<%=sPar%>" title="<%=rs("txt_tips").value%>"  ><%=rs("Menu").value%></a>
                    <%end if %>
                <%else%>  
                    <li><div class="MnuRaya"></div>
                    
                <%end if %>
                <% iNiv = rs("idNivel").value
            end if    
            rs.MoveNext
              i=i+1
        Loop
        
        rs.close
         %>
        </ul></li></ul>
        
    </td>
    </tr>
    </table>

<%
'response.write "<br>LR"
end Sub 

Sub ed_MenPri2 
' Abrir Menu
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
	
    sql = "SELECT  Nivel, Menu, Link , Txt_Tips, Target, Parametros, id_menu, id_perfil, Ind_Activo "
    sql = sql & " FROM  s_Menu "
    If iPerUsu="" then iPerUsu="0"
    if iPerUsu="0" then
        sql = sql & " WHERE (((IdEmpresa)=1) AND ((Fec_Inactivo) Is Null)) "
	'response.write "<br>iPerUsu:" & iPerUsu
    else
       sql = sql & " WHERE (((Id_Perfil)=" & iPerUsu & ") AND ((IdEmpresa)=1) AND ((Fec_Inactivo) Is Null)) "    
    end if        
    sql = sql & " ORDER BY Orden; "
'response.write "<br>2196 sql:=" & sql    
	rs.Open sql,conexion
	'e_gMenu=rs.GetRows
	'rs.close
	'rs.Open sql,conexion
 %>

    <table width="100%"  cellspacing="0"  border="0" align="center"   height="25" bgcolor="#ffffff">
    <tr ><td  bgcolor="#ffffff" title=""  >
        <ul id="nav">
         <% spxx=sPro
            sPro= rs("Link").value 
            if sPro="" then sPro=spxx
            ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("nivel").value, rs("Id_Menu").value
            if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
            sPro=spxx   
         %>
         <!--li><a href="<%=sPar%> " title="<%=rs("txt_tips").value%>"> <%=rs("Menu").value%></a-->

        <%iNiv=rs("Nivel").value
        i=0
        Do While NOT rs.EOF
          
         '   if i>ubound(e_gMenu,2) then exit do
       ' For i=1 to  ubound(e_gMenu,2)
            if ed_Ims<>"" then 
                iy=ed_iMs-rs("id_Menu").value
            else
                iy=1
            end if        
            if iy=0 then ed_sMenSec=rs("Menu").value
        %>
            <% ix= iNiv-rs("Nivel").value
               if ix=0 then %></li>
             <%else%>
                <%if ix<0 then %><ul >
                <%else 
                    if ix=2 then%> 
                        </li></ul></ul></li>  
                    <% else %>    
                        </li></ul></li>  
                    <%end if %>
                <%end if %>
              <%end if
                         spxx=sPro
                          sPro= rs("Link").value
                          if sPro="" then sPro=spxx
                          ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("Nivel").value,rs("id_Menu").value
                           if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
                          sPro=spxx                
             'Sub ed_CalPar (xPas,xCla,xPag,xBus,xCol,xOrd,xFil,xMp,xMs)
                
                
                if rs("Menu").value<>"#" then
                    if rs("Ind_Activo").value=false then%>
                        <li><div class="MnuIna" ><%=rs("Menu").value%></div>
                    <%else%>
                        <li><a href="<%=sPar%>" title="<%=rs("txt_tips").value%>"  ><%=rs("Menu").value%></a>
                    <%end if %>
                <%else%>  
                    <li><div class="MnuRaya"></div>
                    
                <%end if %>
        <% iNiv = rs("Nivel").value
            rs.MoveNext
              i=i+1
        Loop
        
        rs.close
         %>
        </ul></li></ul>
        
    </td>
    </tr>
    </table>


<%end Sub 

Sub Ed_Main

if ed_Link<>"" or ed_iRep=1 then
    ed_Bot(2)="disabled" ' Añadir
    ed_Bot(3)="disabled" ' Guardar
    ed_Bot(4)="disabled" ' eliminar
end if    

%>
    <%if ed_iPas<>4 then
        ix=85
        if ed_iPas=2 then ix=85 %>
        <table width="<%=ix%>%"  cellspacing="0"  border="0" align="center"   height="25" bgcolor="#ffffff">
        <tr><td  class="ed_ti3"><%=ed_sMenSec %></td></tr>
        </table>
    <%
     end if
	if ed_iCol="" then ed_iCol=ed_cCol
	if ed_iOrd="" then ed_iOrd=ed_cOrd
	SqlReg = "Select * FROM " & ed_sNomTab
	sqlReg= sqlreg &  " WHERE (((" & ed_sNomTab & "." & ed_sNomInd & ")= " & Ed_iCla & ")) "
	
' Crear sQue
    for i=1 to ed_iJoin
        ed_sQue(ed_sJoin(i,0),0) =" SELECT " & ed_sJoin(i,1) & "," & ed_sJoin(i,2) & " FROM  " & ed_sJoin(i,3)&  " WHERE Fec_Inactivo is  Null ORDER BY " & ed_sJoin(i,2)
  'response.write "<br>" & Ed_Sque(ed_sJoin(i,0),0)  & " -" & ed_sJoin(i,0) 
    next	

		
'response.write "<br>2241 reg:=" & ed_iRegPag

    if ed_iRegPag=0 then ed_iRegPag=10

'	if ed_iCol="" then ed_iCol=ed_cCol
	if ed_iOrd="" then ed_iOrd=0
    if ed_iPas=4 then ed_Pulsar(0,0)=0
	Select Case ed_iPas 
		case 1,7,  8 ' Listar
			ed_LeePag1 SqlCla
			if ed_ierr=1 then exit sub
			ed_VerPag1 ed_iRegPag, ed_iPag, ngCat
		case 2 ' Modificar
			ed_LeePag2 SqlReg
			ed_VerPag2 ngCat, 1
			'ed_VerPag2_old ngCat, 1
		case 3 ' Grabar Data
			ed_LeePag2 SqlReg
			ed_GraDat ngCat, SqlReg
			ed_iPas=1
			'*sAcc=mySmartUpload.Form("Accion").values
			sAcc=request.Form("Accion")
'response.write "<br>2674 sAcc:=" & sAcc & "  ierrg=" & 	ed_iErrG & " iFor:=" 
			if sAcc= "Añadir" then
				ed_LeePag2 SqlReg
				ed_VerPag2 ngCat, 2	
			else
				if ed_iErrG=1 then
					ed_LeePag2 SqlReg
					ed_VerPag2 ngCat, 1	
				else
					ed_LeePag1 SqlCla
					ed_VerPag1 ed_iRegPag, ed_iPag, ngCat
				end if	
			end if	
			
		case 4 ' Salida a Excel
			ed_OutExcel
			ed_LeePag1 SqlCla
			'ed_LeePag1 SqlReg

            
	        ix= UBound(ngCat,2)+1
			ed_VerPag1 ix, ed_iPag, ngCat
			
		case 5 ' Agregar Registro
			ed_LeePag2 SqlReg
			ed_VerPag2 ngCat, 2	
		case 6 ' Eliminar  Registro
			ed_LeePag2 SqlReg
			ed_EliReg ngCat, SqlReg
			ed_LeePag1 SqlCla
			ed_VerPag1 ed_iRegPag, ed_iPag, ngCat
	'	case 7 ' Buscar
	'		ed_LeePag1 SqlCla
	'		ed_VerPag1 ed_iRegPag, ed_iPag, ngCat
			
	end Select%>
	<%if ed_iPas<>4 then %>
	<center>
	
	<div style="font-family:Arial; font-size: xx-small ; color:#ffffff">
	ver 13.06.14 <%=iPerUsu%>
	</div>
	</center>
	<%end if%>
	
<%
End Sub
Sub ed_vCombo2


    dim rst    
    
	set rst = server.CreateObject("ADODB.Recordset")
	rst.CursorType = 1
	rst.LockType = 1

%>
    <table width="<%=ed_sCombo(i,4)%>" cellspacing="1"  bgcolor="#ffffff" align="left" >
<%    	
	for i=1 to ed_iCombo
		'if isNull(ed_sCombo(i,3)) or ed_sCombo(i,3) = "" then
		'	'response.write	"<br>2339 NO:=" & ed_sCombo(i,3)    	
		'else
		'	response.write	"<br><br>2338 " & ed_sCombo(i,1)    	
		'	response.write	"<br>2339 " & ed_sCombo(i,3)    	
		'	rst.filter = ed_sCombo(i,3)
		'End if
		'response.write "<br>2344 ed_sCombo(i,1):= " & i & "--" & ed_sCombo(i,1)
		rst.open ed_sCombo(i,1),conexion

	    dim gX
        gX=rst.getrows

			'response.write "<br>2369 spro:=" & sPro
		'	response.write "<br>2372 i:=" & i
		ed_sPar(i,1)=gX(0,0)
        if isnull(ed_sPar(i,0)) or ed_sPar(i,0)="" then 
            if ed_sCombo(i,2)<>"" then 
                ed_sPar(i,0)=ed_sCombo(i,2)
             else   
                ed_sPar(i,0)=gX(0,0)
             end if   
           ' response.write "<br>klasjjjjjjjjjjjjjjjjjjj" & ed_sPar(i,0)
        end if    
        rst.close 
       %>
        
   
	    <tr>
	        <td width="<%=ed_sCombo(i,5)%>" class="ed_dim41">&nbsp;&nbsp;<%=ed_sCombo(i,0) %>:</td>
	        <td width="<%=ed_sCombo(i,6)%>" style="font-size:  10px; background-color:#ffffff; color:#000000;">
	        <%if ed_iPas<> 4 then %>
	    	    <select size="1" name="per" id="Select1"   onchange ="location.href=this.options[this.selectedIndex].value"  style="width:100%; font-size:12px; font-family:Verdana; padding:3px 0 3px 0 ">
            <%  if ed_sCombo(i,2)<>"" then   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=ed_sCombo(i,2)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP

            %>
        			<option value="<%=sPar%>"  <% if ed_sPar(i,0) ="Todas" then response.Write"selected" %> style="width:100%; font-size:12px; font-family:Verdana; padding:13px 0 13px 0 " >
				    <%="[" & ed_sCombo(i,2) & "]"%>
			        </option>

    		<%  end if
    		    wO=0

    		    for j=0 to ubound(gX,2)
    		   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=gx(0,j)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP
				   if isnumeric(ed_sPar(i,0)) then
						ix=ed_sPar(i,0) -gX(0,j)
				   else
					 ix=1
					 if ed_sPar(i,0) =gX(0,j) then ix=0
				   end if
     	    %>
			    <option value="<%=sPar%>"  
			    <% 
			    if ix=0 then 
			        response.Write"selected" 
			        wO=1
			    end if    
			        %>  >
				    <%=gX(1,j) %>
			    </option>
			    <%next
			     if wO=0 then ed_sPar(i,0)=ed_sPar(i,1) %>
	        </select>
	       <% else 
	            %>
	                <%=ed_Spar(i,0) %>
	            <%
	          end if %>
	    </td> 
	   </tr>
	             
<%	Next%>
    </table>
<%    
            		       
End sub



%>
