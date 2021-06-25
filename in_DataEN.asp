

<%

Dim ed_iSession
ed_iSession=Request.Cookies("idsession")
'response.write "<br>" & ed_iSession
if ed_iSession="" then
    ed_iSession=Session.SessionID 
    Response.Cookies("idsession")=Session.SessionID 
end if    
Response.Cookies("idsession").Expires=now()+365
'Elimine el Burcar de la linea 2193
'================================
' Mejoras
' Se agrego el objeto fecha 02/03/2016
' Se corrigió falla de campo de fecha 29/02/2016
' Se incluyó Vercombo 23/06/2015
' Se incluyó Menver 23/06/2015
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
 <link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css" rel="stylesheet" type="text/css"/>
  <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.5/jquery.min.js"></script>
  <script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/jquery-ui.min.js"></script>
  
  
<script type="text/javascript">	

$(document).ready(function() {
        
    $("#ed_Fecha1").datepicker({dateFormat: "dd/mm/yy",changeYear: true,yearRange: "1950:nnnn"});
    $("#ed_Fecha1").datepicker();
    $("#ed_Fecha2").datepicker({dateFormat: "dd/mm/yy",changeYear: true,yearRange: "1950:nnnn"});
    $("#ed_Fecha2").datepicker();    
    $("#ed_Fecha3").datepicker({dateFormat: "dd/mm/yy",changeYear: true,yearRange: "1950:nnnn"});
    $("#ed_Fecha3").datepicker();    
    $("#ed_Fecha4").datepicker({dateFormat: "dd/mm/yy",changeYear: true,yearRange: "1950:nnnn"});
    $("#ed_Fecha4").datepicker();    
    $("#ed_Fecha5").datepicker({dateFormat: "dd/mm/yy",changeYear: true,yearRange: "1950:nnnn"});
    $("#ed_Fecha5").datepicker();    
	
  });

function w3_open() {
 
 if (document.getElementsByClassName("ff_sidenav")[0].style.display=="none") 
        {
        document.getElementsByClassName("ff_sidenav")[0].style.display = "block";
        document.getElementsByClassName("w3-opennav").color="#ffffff";
        document.getElementById("contenido").style.width="75%";
        document.getElementById("contenido").style.marginLeft="10px";
        document.getElementById("contenido").style
         }
        else
        {
        document.getElementsByClassName("ff_sidenav")[0].style.display = "none";
        document.getElementsByClassName("w3-opennav")[0].style.display = "inline-block";
        document.getElementById("contenido").style.width="100%";
        }
  
  
}

function mensaje(sMensaje) 
{
alert(sMensaje)
}

function valnum(valor,snomcam,v2)
{
    var valor2 = valor.replace(",", ".");
if (isNaN(valor2)) 
{
alert('El dato debe ser numérico y colocó: ' +document.getElementById(snomcam).value )

document.getElementById(snomcam).focus()
document.getElementById(snomcam).click()
//document.getElementById(snomcam).style.background="#800000"
document.getElementById(snomcam).value=v2

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
' ed_linkVolver ' Es la programa que se usa cuando presiona volver
' ed_Pulsar(9,2) ' Botones que aparecen en cada registro
' ed_sCampo(99,9) '0.- Titulo del campo 1= Default 2= No Presentar 3=Read only 4=1-Obligatorio 5.-Tool Tips 6= Total del campo 7= 1=salto, 8- 1=Copia Valor anterior
' ed_sTitle2(99) ' Titulo de la segunda pagina
' ed_iSum(99) ' Sumatoria del campo
' ed_iGrupo ' Código del grupo Leido en el perfil
' ed_sTarget  Target del sLink
' ed_sBotonC ' Botones 0.- Texto del Boton, 1.- Link , 2.- Target 3.- Tools Tips 4.- Parametro

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
	Dim sqlExcel
    Dim ngCat
    Dim gData3
	Dim sGra(99) 'Data A Grabar
	Dim ed_iNumCam2
	Dim ed_iNumCam
	Dim ed_iNumCam3
	
	Dim SqlReg
	Dim SqlCla
	dim ed_Sql3
	Dim ed_rs1
	Dim ed_rs3
	
	Dim TipCam(99,5) ' 0= (string, numerico, fecha) 1.-Tipo(indicador, password, etc) 2.-lineas de textarea  3.-Numero de caracteres 4.-Obligatorio 9.-Color
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
	dim ed_linkVolver ' Es la programa que se usa cuando presiona volver
	Dim ed_Sqlfil ' Sql del menú de opciones
	Dim ed_iFil
	Dim ed_Pulsar(9,2)
	Dim ed_iMaxReg
	Dim ed_sCampo(99,8)
	Dim ed_sTitle2(99)
	Dim ed_Formato(99, 4) ' 0=Salto , 1-Columna(pixel) , 2-Ancho Campo (Caracteres), 3-Ancho Texto(Pixel), 4-Fila (textArea)
	Dim ed_Formato3(99, 2) ' 0=Salto , 1-Columna(pixel) , 2-Ancho Campo (Caracteres), 3-Ancho Texto(Pixel), 4-Fila (textArea)
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
    dim ed_sBotonC(9,4)
    dim ed_iAnc1
    Dim ed_sPar(10,1)  'parametros del Combo
    dim ed_iCombo ' Numeros de combos
    dim ed_sCombo(10,6) ' 0 = Titulo , 1 Sql , 2 <>Null es un total, 3 Filtro, 4 width tabla, 5 width del titulo , 6 width del combo
    dim ed_Total(99,1) ' 0=suma 1=cuenta


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
	if ed_iPas="" then ed_iPas=5
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
		ed_sPar(1,0)=Request.QueryString("cc_p1")
		ed_sPar(2,0)=Request.QueryString("cc_p2")
		ed_sPar(3,0)=Request.QueryString("cc_p3")
		ed_sPar(4,0)=Request.QueryString("cc_p4")
		ed_sPar(5,0)=Request.QueryString("cc_p5")
		ed_sPar(6,0)=Request.QueryString("cc_p6")
		ed_sPar(7,0)=Request.QueryString("cc_p7")				
		ed_sPar(8,0)=Request.QueryString("cc_p8")		
		ed_sPar(9,0)=Request.QueryString("cc_p9")
				
		if ed_sPar(1,0)="" then ed_sPar(1,0)=Request.Cookies("cc_p1")	
		if ed_sPar(2,0)="" then ed_sPar(2,0)=Request.Cookies("cc_p2")	
		if ed_sPar(3,0)="" then ed_sPar(3,0)=Request.Cookies("cc_p3")	
		if ed_sPar(4,0)="" then ed_sPar(4,0)=Request.Cookies("cc_p4")	
		if ed_sPar(5,0)="" then ed_sPar(5,0)=Request.Cookies("cc_p5")	
		if ed_sPar(6,0)="" then ed_sPar(6,0)=Request.Cookies("cc_p6")	
		if ed_sPar(7,0)="" then ed_sPar(7,0)=Request.Cookies("cc_p7")	
		if ed_sPar(8,0)="" then ed_sPar(8,0)=Request.Cookies("cc_p8")	
		if ed_sPar(9,0)="" then ed_sPar(9,0)=Request.Cookies("cc_p9")	

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


'==========================================================================================
' Leer un registro de la tabla de Categoría
'==========================================================================================
Sub ed_LeePag1 (SqlInp)

	Dim gTem

	if sqlInp="" then 
	    SqlInp = "Select * FROM " & ed_sNomTab
        if ed_iDes<>2 then SqlInp = SqlInp '& " WHERE Fec_Inactivo is  Null "     
    end if    
    
 if ed_ides=1 then   %>
    <div style="width:450px;  margin:10px 10px 10px 10px; border: solid 1px #666666; text-align:justify" >
    <%
    response.write "<br>391 sqlinp:=" & sqlinp%>
    </div><%
 end if   


' Abrir Recordset
	set rso = CreateObject("ADODB.Recordset")
	rso.CursorType =adOpenKeyset' 1 ' 0=El cursor solo avanza 2= Puedes avanzar y retroceder 
	rso.LockType = 1
	rso.MaxRecords =1

'response.write "<br>384 " & sqlinp			
	rso.Open sqlinp,conexion

    

	ed_ierr=0
    if ed_iNumCam>rso.fields.Count-5 then
        ed_iNumCam = rso.fields.Count-5
    end if 
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
    'response.write "<br>Paso<br>"
    
' Buscar Palabra
	if ed_sBus<>"" then
	  
		sx=" ("
		ix=0
		ii=0
		
		if isnumeric(ed_sBus) and len(ed_sBus) < 10 then	
	'response.write"<br>459" & ed_SBus
			for i=0 to rso.fields.Count -6
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
		  
		    for i=0 to rso.fields.Count -6
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
'response.write "<br>722 sx:=" & sx & " icol:=" & ed_Icol
		SqlInp2 = SqlInp & sx
		if ed_iOrd=1 then sqlInp2 = SqlInp2 & " Desc "
		'response.write "<br>725 SqlInp2:= " & SqlInp2


' Leer registro(s) para vizualizar		
	set ed_rs1 = CreateObject("ADODB.Recordset")
	if ed_sBus<>"" then ed_rs1.Filter=ed_sFil
	'ed_rs1.CursorType = 1
	'ed_rs1.LockType = 1
	'ed_rs1.CursorType = adOpenKeyset
	'ed_rs1.LockType = 2
	
	if ed_iSwReg = 1 and ed_sBus="" Then 
	    ed_rs1.MaxRecords = ed_iRegPag * ed_iPag*2 
	end if    
'response.write "<br>740 " & sqlinp2
	ed_rs1.Open sqlinp2,conexion
'response.write "<br>739 " & sqlinp2	
	if Not(ed_rs1.EOF) then 
		if ed_iswRep =1 then
			ngCat=ed_rs1.GetRows(ed_rs1.MaxRecords)
		else
			ngCat=ed_rs1.GetRows
		end if	
		ed_Ilof=1
		
' Totalizar Campo
    for i=0 to rso.fields.Count -5
        if ed_sCampo(i,6)=1 then
            for j=0 to ubound(ngCat,2)
                
                if isnumeric(ngCat(i,j)) then
               ' response.write "<br>" & ngCat(i,j)
                    ed_iSum(i)= ed_iSum(i) + cdbl(ngCat(i,j))
                    ed_Total(i,0)=ed_Total(i,0)+cdbl(ngCat(i,j))
                    ed_Total(i,1)=ed_Total(i,1)+1
                    
                    
                    'response.write " -- " & ed_Isum(i)
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

Sub ed_LeePag3

	Dim gTem

   
 if ed_ides=1 then   %>
    <div style="width:450px;  margin:10px 10px 10px 10px; border: solid 1px #666666; text-align:justify" >
    <%
    response.write "<br>391 sqlinp:=" & ed_Sql3%>
    </div><%
 end if   

' Abrir Recordset
	set rso = CreateObject("ADODB.Recordset")
	rso.CursorType = 1 ' 0=El cursor solo avanza 2= Puedes avanzar y retroceder 
	rso.LockType = 1
	rso.MaxRecords =1
'response.write "<br>384 " & sqlinp			
	rso.Open ed_Sql3,conexion

    

	ed_ierr=0
    ed_iNumCam3 = rso.fields.Count-5
    ed_iNumCam3 = rso.fields.Count-5
    
    if ed_iPas=4 then ed_iNumCam3 = rso.fields.Count-5


    
    
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
		SqlInp2 = ed_Sql3 & sx
		if ed_iOrd=1 then sqlInp2 = SqlInp2 & " Desc "
		


' Leer registro(s) para vizualizar		
	set ed_rs3 = CreateObject("ADODB.Recordset")
	if ed_sBus<>"" then ed_rs3.Filter=ed_sFil
	ed_rs3.CursorType = 1
	ed_rs3.LockType = 1
	
	if ed_iSwReg = 1 and ed_sBus="" Then 
	    ed_rs3.MaxRecords = ed_iRegPag * ed_iPag*2 
	end if    
'response.write "<br>649 " & sqlinp2
	ed_rs3.Open sqlinp2,conexion
	if Not(ed_rs3.EOF) then 
		if ed_iswRep =1 then
			gData3=ed_rs3.GetRows(ed_rs3.MaxRecords)
		else
			gData3=ed_rs3.GetRows
		end if	
		ed_Ilof=1
		
' Totalizar Campo
    for i=0 to rso.fields.Count -5
        if ed_sCampo(i,6)=1 then
            for j=0 to ubound(ngCat,2)
                
                if isnumeric(ngCat(i,j)) then
               ' response.write "<br>" & ngCat(i,j)
                    ed_iSum(i)= ed_iSum(i) + cdbl(ngCat(i,j))
                    ed_Total(i,0)=ed_Total(i,0)+cdbl(ngCat(i,j))
                    ed_Total(i,1)=ed_Total(i,1)+1
                    
                    
                    'response.write " -- " & ed_Isum(i)
                end if    
            next
            ed_swSum=i
        end if    
    next
		
		
	else 
		ed_ilof=0
		
	end if	
	

 
 
    
    ed_rs3.close
' Tipo de Campo	
		for i=0 to ed_rs3.fields.Count -1
			TipCam(i,1)=null
			CamTip i, ed_rs3.fields(i).name, ed_rs3.fields(i).DefinedSize, ed_rs3.fields(i).Type
'response.write "<br>515 Nombre:=" & 	ed_rs3.fields(i).name & " campo:=" & i					
		next 	
				
	for i=0 to ed_rs3.fields.Count -1
        ix=1
		for j=1 to 8
		    iz= ix and ed_rs3.fields(i).Attributes
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
'response.write "<br>844 Inumcam2:=" & ed_iNumCam2 
' Parametros de los campos
    for i=0 to ed_rs1.fields.Count -1
        if ed_sCampo(i,3)=1 then 
			ed_sCampo(i,3) ="readonly"
			'response.write "<br>PasoLR:="  & i
		end if
    next


' Tipo de Campo	
	for i=0 to ed_rs1.fields.Count -1
		CamTip i, ed_rs1.fields(i).name, ed_rs1.fields(i).DefinedSize, ed_rs1.fields(i).type
        TipCam(i,3)=ed_rs1.fields(i).DefinedSize
' tipo de campos 1=Numerico 2=texto 3=fecha		
'response.write "<br>851 i:= "& i & " " & ed_rs1.fields(i).name & "  " & ed_rs1.fields(i).type & " DAto:=" & ngCat(i,0)
		select case ed_rs1.fields(i).type
		    case 2,3,4 ' Numérico
		      TipCam(i,0)=1  
		      TipCam(i,3)=ed_rs1.fields(i).precision
		    case 6 ' Numérico
		      TipCam(i,0)=1  
		      TipCam(i,3)=14
  
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
    if ed_Formato(99,0)="" then
        ed_Formato(99,0)="w3-col l2  w3-left w3-padding w3-small"
    end if  
    if ed_Formato(99,1)="" then
        ed_Formato(99,1)="w3-text-theme w3-left w3-small"
    end if    
    if ed_Formato(99,2)="" then    
        ed_Formato(99,2)="w3-input w3-border w3-small"
    end if    
      
    for i=0 to ed_rs1.fields.Count -1
        if ed_Formato(i,0)="" then ed_Formato(i,0)=ed_Formato(99,0)
        if ed_Formato(i,1)="" then ed_Formato(i,1)=ed_Formato(99,1)
        if ed_Formato(i,2)="" then ed_Formato(i,2)=ed_Formato(99,2)
    next



    
end sub


Sub VerCampo 

    %>
    <table border="1">
     <tr>
        <td>id</td>
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
	
	
end sub
Sub CamTip (i, name, DefinedSize, iTipo)

'response.write "<br>Campo:" & i & " Name:=" & Name

' SI posee mas de 50 Caracteres
	if DefinedSize >50 then
		TipCam(i,1)=1

        select case definedSize
            case definedsize>49 and definedsize<100
                TipCam(i,2)=1
            case definedsize>99 and definedsize<300
                TipCam(i,2)=1
            case definedsize>299 and definedsize<500
                TipCam(i,2)=1
            case else
                TipCam(i,2)=1
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

' 	Integer	
    if iTipo=3 then TipCam(i,5)=3
' Doble precision
    if iTipo=4 then TipCam(i,5)=4
' Otros    
    TipCam(i,5)=iTipo
	
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
			ed_sQue(i,1) =  "True" & chr(9) & "Si" & chr(9)
			ed_sQue(i,1) = ed_sQue(i,1)  & "False" & chr(9) & "No" & chr(9)
			
			'ed_sQue(i,1) =  "Verdadero" & chr(9) & "Si" & chr(9)
			'ed_sQue(i,1) = ed_sQue(i,1)  & "Falso" & chr(9) & "No" & chr(9)
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
			
' Campos Fecha
	sx= Ucase(name)
	ix = instr(1,sx,"FEC_")
	'response.Write ix & "-" & sx & "<br>"
	if ix<> 0 then TipCam(i,1)=8					
			
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

	<table border="0" width="100%" cellpadding="0" cellspacing="0"   align="center" id="Table6" style=" border:1px solid  #dddddd;  border-top-left-radius: 8px;border-top-right-radius:8px; "  >
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
    <table border="<%=ix%>" width="<%=ed_iAnc1%>" cellpadding="0" cellspacing="1" bgcolor="<%=sBac %>"   align="center" id="Table7" >
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
            Select case ed_rs1.fields(i).type
					    case 135, 200,129,131
					        sAli ="center"
					    case 202,2,3
					        sAli="left"    
					    case else
					        sAli="right"
					end select    
'			sAli ="Center"	
			
%>		
			
			<td  height="30" align="<%=sAli%>"  class="ed_ti1" title="<%=ed_sCampo(i,5)%>" >
				
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
						
					<a  href="<%=sPar %>" title="<%=ed_sCampo(i,5) &  ed_rs1.fields(i).type  %>" ><%=sxTit%>
					<% if i- ed_iCol=0 then %>
						<% if ed_iOrd=0 then %>
							<img src="images\s0.gif" border="0">
						<% else%>
						    <img src="images\s1.gif" border="0">
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
										iz = gData(j,i) - Sc
										if iz=0 then response.write "&nbsp;&nbsp;" & ST  
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
				        
					Select case ed_rs1.fields(j).type
					    case 135,3,2
					        sAli ="center"
					    case 202,200,129
					        sAli="left"    
					    case else
					        sAli="right"
					end select            
					
					
					%>
					<td   align="<%=sAli%>"  title="<%=ed_sCampo(j,5) %>" style="margin-right:13px; margin-left:13px">
					<%
						
						if ed_iPas<>4 then
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>" style="padding: 0px 10px 0 10px">
								<%else %>
										 
								<%end if 		
						end if		
						if ed_iPas<>4 then response.Write " &nbsp;&nbsp;"
						'TipCam(j,1)	= trim(TipCam(j,1))
						if TipCam(j,1)=5 then
							
							response.Write String (8,"#") 
						else%>
						     <% if isnull(gData(j,i)) then
						        else
						            if Len(gData(j,i))<50 then %>
							           		<% Select Case  ed_rs1.fields(j).type %>
							           		    <% case 6,4 %>
							                        <%=formatnumber(gData(j,i),2)%> 
							           		    <% case 131 %>
							                        <%=formatnumber(gData(j,i),2)%> 
							                    <%case else %>
							                        <%=gData(j,i) %>
							                <%end select %>   
							        <%else %>      
							            <span style=" border: solid 1px #ffffff; text-align:justify" >
							                    <%=gData(j,i)%> 
							            </span>
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

    <table border="0" width="100%" cellpadding="0" cellspacing="0"   align="center" id="Table10" style=" border-bottom: solid 1px #dddddd; border-left: solid 1px #dddddd; border-right: solid 1px #dddddd;  border-bottom-right-radius: 8px;border-bottom-left-radius: 8px;" >
   


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




Sub ed_VerPag3

%>
<!-- Chequear si hay Data -->

<%	if ed_Ilof = 0 then 
    	ed_CalPar 5,ed_iCla,1,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	    ed_Botones ed_iPag, iPaginas, gData%>
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
	iTotal = UBound(gData3,2)+1
	Ed_iMaxReg =UBound(gData3,2)+1
	
' Calculo el numero de páginas
	iPaginas = (iTotal \ ed_iRegPag)
	if iTotal mod ed_iRegPag > 0 then
		iPaginas = iPaginas + 1
	end if
	
'Si no es una página válida, comienzo en la primera
	ed_iPag=cInt(ed_iPag)
	if ed_iPag < 1 then
		ed_iPag = 1
	end if
'Si es una página mayor al nº de páginas, comienzo en la última
	if ed_iPag > iPaginas then
		ed_iPag = iPaginas
	end if


'Calculo el índice donde comienzo:
	iComienzo = (ed_iPag-1)*ed_iRegPag
	
'Calculo el final y Si no tengo suficientes registros restantes, voy hasta el final
	iFin = iComienzo + (ed_iRegPag-1)
	if iFin > UBound(gData3, 2) then
		iFin = UBound(gData3, 2)
	end if

    
    ix=0
    sBac="#f1f1f1"
	if ed_iPas=4 then 
	    ix=1 
	    sBac="#ffffff"
    end if
%>
<%if ed_iPas<>4 then%>

<div style=" overflow:Auto; width:100%"> 
	<table border="0" width="100%" cellpadding="0" cellspacing="0"   align="center" id="Table2" style=" border:1px solid  #ffffff;  border-top-left-radius: 8px;border-top-right-radius:8px;" >
	<%iColSpan=ed_iNumCam3+ed_pulsar(0,0) 
	  if ed_iRan=1 then iColSpan=iColSpan+1%>
	  
	
        <tr bgcolor="<%=ed_sBacCol1 %>" >
        <td colspan="<%=iColSpan%>">
    		
	    		<%  ed_Botones ed_iPag, iPaginas, gData%>	
		    
		</td></tr>
 
	
     </table>
  </div>	
<%end if%>	     
<!-- TItulo -->
    
   
   
  <div class="w3-responsive w3-container">
    
    <table class="w3-table w3-border w3-small" id="Table3" >
     <thead>
   	<tr class="w3-theme-light " style="text-decoration:none">
   	
<%   	
' #e4ecf7
	if ed_iRan=1 then %>
		<th >
			Rank
		</th>
<%	end if
   	


'response.write "<br>1398 ed_iNumCam:= " & ed_iNumCam3
   	for i=0 to ed_iNumCam3-1
         iO=ed_sCampo(i,2)
		 if ed_iRep<>1 then
			if ed_iPas=4 and ed_sCampo(i,2)="2" then io=""
		 end if	
   		 if iO="1" or iO="2" Then

		 Else 


			if ic1=0 then ic1=1 else ic1=0
			sC="w3-red"
			if ic1=1 then sc="w3-blue"
			ixt=ed_rs3.fields(i).Type
			'if i=0 then ixt=10
			select case ixt
			    case 3,135,11,2
			        sAli="w3-left-align"
			    case else
			        sAli="w3-right-align"
			end select
%>		
			
			<th  title="<%=ed_sCampo(i,5)%>"  class="<%=sAli%>">
				
				<% sx=ed_rs3.fields(i).name
					sxTit=sx & "(" & ixt&")"
    				'	if ed_sCampo(i,0)<>"" then sxTit=ed_sCampo(i,0)
					
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
				
			</th>
		
<%		End if
	next
	if ed_pulsar(0,0)<>0 then
	    for ip=1 to ed_pulsar(0,0)%>
    	    <th   >
    	    
	        </th>
	<%	    next
	 end if %>
	</tr>
 </thead>

	
<%
'#FEFEF9
' #9eb6ce
' #f7f7f7
' Mostrar Registros	
	ixL=0
    'response.write "<br>1351 iFin:= " & iFin
	for i = icomienzo to ifin
        if ixL=0 then	%>
			<tr  height="30" class="ed_ln w3-hover-theme"   > 
		<%else	%>
			 <tr height="30" class="ed_ln w3-hover-theme" > 
		<%end if
		

		if ixL= 0 then ixL =1 else ixL=0
		
	
			if ed_Link<>"" then
				sP=sPro
				sPro=ed_link
			end if	
		
			ed_CalPar 2,gData3(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
			if ed_Link<>"" then
			    ed_CalPar 1,gData3(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
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
            ed_CalPar 1,gData3(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
            sPar=sPar&"&ed_det=" &  gData3(ed_iDet,i)
            %>
           
            
        <% ' sPar=sp
        end if
        

        
		for j=0 to ed_iNumCam3-1

    		ixt=ed_rs3.fields(j).Type
			select case ixt
			    case 3,135,11,2
			        sAli="w3-left-align"
			    case else
			        sAli="w3-right-align"
			end select
		
		
		    iO=ed_sCampo(j,2)
			if ed_iRep<>1 then
				if ed_iPas=4 and ed_sCampo(j,2)="2" then io=""
			end if	
			
		 if iO="1" or iO="2" Then
		  'response.Write "<br>1390 ed_iNumCam:=" & ed_iNumCam3
		 Else 
            'response.write "<br>1399 TipCam(j,1):= " & TipCam(j,1)
			Select Case TipCam(j,1)
				case 2,3
					
					%>		
					<td   class="<%=sAli%>" title="<%=ed_sCampo(j,5)%>" >
						<% if ed_iPas<>4 then 
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>"> 
								<%end if 
							end if		
							'response.Write gData3(j,i) 
							if gData3(j,i)<>"" then
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
								'	response.Write sC  & "----------" &gData3(j,i)
										 if isnumeric(sc) then
										    iz = gData3(j,i) - Sc
										    if iz=0 then response.write "&nbsp;&nbsp;" & ST  
										else
										    if gData3(j,i) = Sc then response.write "&nbsp;&nbsp;" & ST  
										end if    
										sCam=mid(sCam,ix+1)
									end if
								loop until ix =0					
						'response.write "&nbsp;&nbsp;" & gData3(j,i)& "GGGGGG"
							end if
				
				case 6
					
					%>		
					<td  class="<%=sAli%>" title="<%=ed_sCampo(j,5)%>">
						<% if ed_iPas<>4 then 
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>"> 
								<%end if 
							end if		
							if  gData3(j,i)  then sx="Si" else sx="No"
							if gData3(j,i)<>"" then
						       response.write "&nbsp;&nbsp;" & sx 
							end if
                		
				    %> 
				    
				<%
				case else
				  
				        
					
					
					%>
					<td   class="<%=sAli%>"  title="<%=ed_sCampo(j,5)%>" style=" padding:0 10px 0 10px">
					<%
						
						if ed_iPas<>4 then
								if ed_iRep<>1 then%>
									<a href="<%=sPar%>" title="<%=ed_sCampo(j,5)%>" target="<%=ed_sTarget%>">
								<%else %>
										 
								<%end if 		
						end if		
						
						if TipCam(j,1)=5 then
							
							response.Write String (8,"#")
						else%>
						     <% if isnull(gData3(j,i)) then
						        else
						            if Len(gData3(j,i))<50 then %>
						                <% select case TipCam(j,5)%>
						                    <%case 3 ' Integer%>
						                        <%if isnumeric(gData3(j,i)) then%>
							                        <%=formatnumber(gData3(j,i),0)%>
							                   
                                                <%else %>
                                                    <%=gData3(j,i)%>
                                                <%end if%>
                						    <%case 4, 6  ' Double precision%>
							                   <%=formatnumber(gData3(j,i),2)%>
						                    <%case else %>
							                    <%=gData3(j,i) %>
							            <% end select %>
						        <%else %>      
							            <div style="width:450px;  margin-bottom:10px; margin-left:8px; border: solid 1px #ffffff; text-align:justify" >
							            <%=gData3(j,i)%> 
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
	                ed_CalPar ed_iPas,gData3(0,i),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs 
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

    <table border="0" width="100%" cellpadding="0" cellspacing="0"   align="center" id="Table5" style=" border-bottom: solid 1px #dddddd; border-left: solid 1px #dddddd; border-right: solid 1px #dddddd;  border-bottom-right-radius: 8px;border-bottom-left-radius: 8px;" >
   


<%if ed_iPas<>4 then%>   



<%  if Ed_iMaxReg>ed_iRegPag then %>   
        <tr><td colspan="<%=iColSpan%>" class="ed_bac1" valign="middle" height="30" align="center">
	        <%ed_Paginar    %>
	
	</td></tr>
<% end if %>	
	<tr><td   style="font-family:verdana; font-size:8pt; font-weight:  bold; color:#666666; background-color:#ffffff; text-align:center" height="25">

       
       Total Registros:= <%=Ed_iMaxReg%> 

	
	</td></tr>
<% end if%>		
</table>
<%
	
		

	
	
'	Ed_VerPag ed_iPag, iPaginas, Ubound(gData,2)
if ed_ides=1 then   Vercampo	
End Sub ' Verpag3

Sub Ed_Botones (ixPag, iPaginas, gData)

%>

    <table width="98%"  border="0" cellpadding="0" cellspacing="0" id="table9"  align="center"  >
	    <tr >
		<td width="15%" align="center"  valign="middle" >
		    <%if Ed_iMaxReg>0 Then %>
		        <font face= "Verdana" size="1" color="#000000">
			        <% Response.Write("Pagina " & ixPag & " de " & iPaginas & "</b>")%><br />
			        
				</TD>
			    </font>
			<%end if %>	
		</td>
        <td width="15%" align="center"  valign="middle" >
	        <%if ed_sBotonC(0,0)<>"" then 
	            sp=sPro
	            sPro=ed_sBotonC(0,1)
	            ed_CalPar ed_iPas,0,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	            sPro=sP
	            %>
	            <div class="ed_boton3"><a href="<%=sPar %>"   title = '<%=ed_sBotonC(0,3)%>' target="<%=ed_sBotonC(0,2) %>" >
	            <%=ed_sBotonC(0,0)%></a></div>	
	        <%end if %>
	     </td>		
	    <%ed_CalPar 7,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs%>
		<form action="<%=sPar%>"  method='post' >
		<%
		'response.write "<br>2205 sPar:= " & sPar
		%>
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
	                    sPro=sP
	                    if ed_sBotonC(i,4)<>"" then sPar=sPar&ed_sBotonC(i,4)
	                    %>
	                    <div class="ed_boton3" style="float: left"><a href="<%=sPar %>"   title = '<%=ed_sBotonC(i,3)%>' 
	                    <%if ed_sBotonC(i,2)<>"" then%>
	                            target="<%=ed_sBotonC(i,2)%>" 
	                    <%end if %>
	                    >
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
	
	'LR2019 - 6 del For
	'Response.Write "<br>" & ed_rs1.fields.Count 
    ' for i=0 to ed_rs1.fields.Count -6
		' Response.Write "<br>i:= " & i
        ' if ed_sCampo(i,3)="1" then 
			' ed_sCampo(i,3) =""
			' 'response.write "<br>PasoLR:="  & i
		' end if
    ' next
	'Response.Write("<script>alert('Registro Guardado');</script>") 
	'Response.end
    
	for i=0 to ed_rs1.fields.Count -6
         sx=ed_rs1.Fields(i).name
         sGra(i)=request.Form(sx)
        'response.Write( "<br>1624 Campo:=" & i & " Nombre:=" &   sx & " - Valor:=" & request.Form(sx))
    next   
       
	
' Validar Not Null
   ed_iErrG=0
   for i = 1 to ed_rs1.fields.Count-6	
		'Response.Write "<br>i:=" & i
        if sGra(i)="" and ed_sCampo(i,4)=1 then
            ed_sErr(i)="<br><font face= 'verdana' size='2' color='#ff0000' >&nbsp;&nbsp;<b> Error: Los campos marcados con * son obligatorio </b></font>"
			'Response.Write "<br>i:=" & i & " con error"
            'Response.Write "<br><font face= 'verdana' size='2' color='#ff0000' >&nbsp;&nbsp;<b>Falta Data en el campo:=" & ed_sCampo(i,0) & "</b></font>"
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
		'response.Write ("<br> 944 " & ed_rs1.fields(i).name) & " Grabar..." & sGra(i)
			if sGra(i)<> ""  then 
				'response.Write ("<br>a" & ed_rs1.fields(i).name & "..." & sGra(i))
				select case TipCam(i,1)
					case "6"
						if sGra(i) = "Verdadero" then sGra(i)=True
						if sGra(i) = "True" then sGra(i)=True
						if sGra(i) ="Falso" then sGra(i)= False
						if sGra(i) ="False" then sGra(i)= False
					case "8"
						'response.write "<br>pasooooooooooooooooo1"
						'response.Write ("<br>954 i:=" & i & "  name:="  & ed_rs1.fields(i).name & "..." & sGra(i)) & "... len=" & len(sGra(i)) &   "  TipCam:=" & TipCam(i,1) & " Precision:=" & ed_rs1.fields(i).Precision
					case else

				end select
			'response.Write ("<br>954 i:=" & i & "  name:="  & ed_rs1.fields(i).name & "..." & sGra(i)) & "... len=" & len(sGra(i)) &   "  TipCam:=" & TipCam(i,1) & " Precision:=" & ed_rs1.fields(i).Precision
			    select case ed_rs1.fields(i).Precision
					case 23   
						sd=cDate(sGra(i))			 
						rsg(ed_rs1.fields(i).name)=sD
					case 7,19
						sx=replace(sGra(i),",",".")
						rsg(ed_rs1.fields(i).name)=sx
					case 8
						'response.write "<br>pasooooooooooooooooo2"
					case else
						rsg(ed_rs1.fields(i).name)=sGra(i)		
						'response.write "<br>pasooooooooooooooooo3"
				end select
			else 
			'response.Write ("<br>b...........	" & ed_rs1.fields(i).name & "..." & sGra(i)) 
				rsg(ed_rs1.fields(i).name)=null
			end if	
		next
		'   response.Write Request.ServerVariables("REMOTE_ADDR") & "..." & len(Request.ServerVariables("REMOTE_ADDR")) 
		'LR26Abr2019
		'rsg("IP")=Request.ServerVariables("REMOTE_ADDR")
		'rsg("Fec_Ult_Mod")=Now()
		'If Session("Usuario")<>"" then 	rsg("Usr")=Session("Usuario") else rsg("Usr") = "Not User"
		'if ed_iDes=2 then rsg("Fec_Inactivo")=null
		'rsg("idSession")=ed_iSession
		Response.Write("<script>alert('Registro Guardado');</script>") 
		'Response.end
	elseif sAcc= "Eliminar" Then
		'LR26Abr2019
		'rsg("IP")=Request.ServerVariables("REMOTE_ADDR")
		'rsg("Fec_Inactivo")=Now()
		'If Session("Usuario")<>"" then 	rsg("Usr")=Session("Usuario") else rsg("Usr") = "ed_GraDat"
		'rsg("idSession")=ed_iSession
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
		'LR26Abr2019
		'rsg("IP")=Request.ServerVariables("REMOTE_ADDR")
		'rsg("Fec_Inactivo")=Now()
		'If session("Usuario")<>"" then 	rsg("Usuario")=session("usu") else rsg("Usr") = "ed_EliReg"
		'rsg("idSession")=ed_iSession
		rsg.update
	end if	
	rsg.close
end sub



Sub ed_VerPag2 (gData, iFor)
'response.write "<br>2381 Listo": response.flush
   
	if iFor=1 Then
		ed_CalPar 3,gData(0,0),ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	else
	    ed_CalPar 3,"",ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	end if	 
%>	
	
	<form action="<%=sPar %>"  method='POST' id="FormPag2" class="w3-container" style="margin-top:20px">	</form>
	<div  class=" w3-container  ">
	<div class="w3-row-padding  w3-card-4 w3-theme-l5 w3-container ">
	
	
	
	<%ed_ilin=0
	if ed_iDes=2 then ed_iNumCam2=ed_iNumCam2+4
	for i=0 to ed_iNumCam2

'	
	    if iFor<>1 then ed_sCampo(0,2)=1
		sxTIT= ed_rs1.Fields(i).name 
        if ed_sCampo(i,0)<>"" then sxTit=ed_sCampo(i,0)
        if ed_sTitle2(i) <> "" then  sxTit=ed_sTitle2(i)
		sx=""
		if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
		
		if ed_sCampo(i,4)=1 then sxTIT=sxTIT & "<font face= 'verdana' color='#ff0000' ><b>*</b></font>"

' Campo que no se va a presentar		
		if ed_sCampo(i,2)="1" or ed_sCampo(i,2)="3" then %>
		    <%if iFor=1 then%>
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
            %>
            <div class="<%=ed_Formato(i,0)%>" title="<%=st %>" style="height:70px; ">
            <label class="<%=ed_Formato(i,1)%>">
    		    <%=sxTIT%>
	    	</label>
			
<%				
' Mostrar Campos				
			Select Case TipCam(i,1)%>
			
			<%case 1 ' textarea %>
				<%if iFor=1 then%>
				        <textarea rows="1" name="<%=ed_rs1.Fields(i).name%>" cols="50" title="<%=ed_sCampo(i,5)%>"   <%=ed_sCampo(i,3)%> onKeyUp="return maximaLongitud(this,<%=TipCam(i,3)%>)" tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" maxlength ="<%=TipCam(i,3)%>" form="FormPag2"><%=gData(i,0)%></textarea>
				    <% else    
		    		     sx=""
						 if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)
					     if ed_sCampo(i,8)<>"" then  sx=ed_sCampo(i,0)
						 if ed_iErrG=1 then sx=request.Form(ed_rs1.Fields(i).name)
					    %>
    					<textarea rows="1" name="<%=ed_rs1.Fields(i).name%>" cols="50" title="<%=ed_sCampo(i,5)%>" onKeyUp="return maximaLongitud(this,<%=TipCam(i,3)%>)" tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" maxlength ="<%=TipCam(i,3)%>" form="FormPag2"><%=sx %></textarea>
					<%end if%> 
					
			<%case 2 ' Combo LR2019
				if ed_sCampo(i,3)<> "readonly" then%>  
				<select size=""  name="<%=ed_rs1.Fields(i).name%>" id="<%=i%>"  title="<%=ed_sCampo(i,5)%>" tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" form="FormPag2">
				
				
				
				<option value="0" ></option>
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
						    '   if ed_sCampo(i,8)<>"" then sx=sGra(i)
						         iz =sGra(i) - Sc
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
				
				<select size=""  name="<%=ed_rs1.Fields(i).name%>" id="<%=i%>"  title="<%=ed_sCampo(i,5)%>" disabled  tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" form="FormPag2" >
				 
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
						    '   if ed_sCampo(i,8)<>"" then sx=sGra(i)
						         iz =sGra(i) - Sc
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

				end if
%>
				</select>
				
				<% case 3,6 'Inddicador, Mes , Dia , Año%> 
				
				<select  size="1"  name="<%=ed_rs1.Fields(i).name%>" id="Select2"  title="<%=ed_sCampo(i,5)%>"  <%=ed_sCampo(i,3)%> tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" form="FormPag2">
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
				
				<% case 4 ' Tipo ID_ read Only
				    if ifor=1 then
				    %>  
    					<input  type='text' size='<%=TipCam(i,3)%>' name ="<%=ed_rs1.Fields(i).name%>" <%=j %> value='<%=gData(i,0)%>' readonly   title="<%=ed_sCampo(i,5)%>"  tabindex="<%=i %>"/  class="<%=ed_Formato(i,2)%>" form="FormPag2">
					<%else %>
	    				<input type='HIDDEN' size='<%=ed_rs1.Fields(i).DefinedSize%>' name ="<%=ed_rs1.Fields(i).name%>"    value='' readonly ID="HIDDEN6" name="Text1" title="<%=ed_sCampo(i,5)%>" tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" form="FormPag2"/>
					<%end if %>
					
				<% case 5 ' Password
				    if iFor=1 then%>  
					    <input  type='password' size='<%=ed_rs1.Fields(i).DefinedSize%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=gData(i,0)%>'  <%=ed_sCampo(i,3)%> title="<%=ed_sCampo(i,5)%>" tabindex="<%=i %>"   class="<%=ed_Formato(i,2)%>" form="FormPag2"/>
					<%else 
				        sx=""
		                if ed_sCampo(i,1)<>"" then  sx=ed_sCampo(i,1)%> 
					    <input  type='password' size='<%=ed_rs1.Fields(i).DefinedSize%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>'  <%=ed_sCampo(i,3)%> title="<%=ed_sCampo(i,5)%>"  tabindex="<%=i %>"  class="<%=ed_Formato(i,2)%>" form="FormPag2"/>										
					<%end if %>
					
				<% case 66  ' Indicador
				    sx="No"
				    if gData(i,0)=true then sx="Si"
				    %> 
				    <input type="checkbox" name ="<%=ed_rs1.Fields(i).name%>" tabindex="<%=i %>"  value="<%=sx %>"
				     <%if gData(i,0)=true then response.write " checked"%>
				       title="<%=ed_sCampo(i,5)%>" />
				    <!--input type="checkbox" name ="<%=ed_rs1.Fields(i).name%>" <%=j %> value='1'  title="<%=ed_sCampo(i,5)%>" /><%=gData(i,0)%> -->
				   
				<% case 7  'Carga imagen
				
                    if iFor=1 then
                       sx=gData(i,0)  
                    else
                    	sx=""
                    end if%>
                     
                     <br />
                <div style="text-align:center">
                    <br />
    	            <img src="<%=sx %>" align="middle" border="0"  style=" width:200px; height:200px; border-radius:5px; border:solid 1px #cccccc" alt="" id="ImgFact"/>
    	        <br/>
  	    	    <input type="file" id="fileElem" accept="image/*" onchange="upload('<%=ed_rs1.Fields(i).name%>')" form="FormPag2" tabindex="<%=i %>"  />
  	    	    <br />
  	    	    
	        	<button id="fileSelect" class="ed_boton2">Cargar</button><br />
    	         <input  type='text' size='<%=TipCam(i,3)%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>' <%=ed_sCampo(i,3)%> 
    	         title="<%=ed_sCampo(i,5) %>" id="<%=ed_rs1.Fields(i).name%>" tabindex="<%=i %>" class="<%=ed_Formato(i,2)%>" form="FormPag2"/>
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
                      
			<%case 8 ' Fecha LR2019
			    ixFec=ixfec+1
			    if ed_sCampo(i,8)<>"" then sx=sGra(i)
				if ed_sCampo(i,3)<> "readonly" then  
					if iFor=1 then%> 
						<input  type="text"  class="<%=ed_Formato(1,2)%>" name="<%=ed_rs1.Fields(i).name%>" id="ed_Fecha<%=ixfec%>" value="<%=gData(i,0)%>"  maxlength="10" size="12"  <%=ed_sCampo(i,3)%> form="FormPag2" tabindex="<%=i %>"   />
					<%else %>    
						<input  type="text" class="<%=ed_Formato(1,2)%>" name="<%=ed_rs1.Fields(i).name%>" id="ed_Fecha<%=ixfec%>" value="<%=sx%>"  maxlength="10" size="12"  <%=ed_sCampo(i,3)%>   form="FormPag2" "/>
					<% 
					end if
				else	
					if iFor=1 then%> 
						<input  type="text"  class="<%=ed_Formato(1,2)%>" name="<%=ed_rs1.Fields(i).name%>" id="ed_Fecha<%=ixfec%>" value="<%=gData(i,0)%>"  maxlength="10" size="12"  <%=ed_sCampo(i,3)%> form="FormPag2" tabindex="<%=i %>"   />
					<%else %>    
						<input  type="text" disabled="true" class="<%=ed_Formato(1,2)%>" name="<%=ed_rs1.Fields(i).name%>" id="ed_Fecha<%=ixfec%>" value="<%=sx%>"  maxlength="10" size="12"  <%=ed_sCampo(i,3)%>   form="FormPag2" "/>
					<% 	
					end if
				end if 
				%>    
	        
	        
			<%		
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
					    <input  type='text' class="<%=ed_Formato(i,2)%>" size='<%=TipCam(i,3)%>' name ="<%=ed_rs1.Fields(i).name%>"  value='<%=sx%>' <%=ed_sCampo(i,3)%>  
					    
					
					<% if TipCam(i,0)=1 then %>
					    onkeydown="valnum(this.value, 'campo<%=i%>','<%=sx%>');" onkeyup="valnum(this.value, 'campo<%=i%>','<%=sx%>');" 
					    
					<% end if%>
                    <% if TipCam(i,0)=2 then %>
					    maxlength ="<%=TipCam(i,3)%>"
                        onblur="valtxt(this.value, 'campo<%=i%>');" 						
					 <% end if%>   					    
					     title="<%=ed_sCampo(i,5) %>" id="campo<%=i%>" tabindex="<%=i %>"  form="FormPag2" />
					
			<% end select %>	
	
	
			
			</div>	  
			
				    <%
                      			      
 
    	End if
 	next %>
 	
	<div class="w3-col l12 w3-padding w3-small"></div>
</div>
	<div class="w3-container w3-row w3-margin-top w3-card-4 w3-theme-l5" >
	 
	
	            


                <%if iFor=1 then %>		
					<div class="w3-col l8  w3-left w3-padding " >
					<%if ed_Bot(4)<>"disabled" then%>
							<input type='submit'  value ='Eliminar'  ID="Submit12" NAME="Accion" tabindex="44" <%=ed_Bot(4)%> class="w3-btn w3-theme w3-hover-theme w3-round" form="FormPag2">
					<% end if %>
					</div>
					
                   
					<div class="w3-col l2  w3-left w3-padding ">
					<%	if ed_Bot(3)="disabled" then
						else %>
						<input type='submit'  value ='Guardar' ID="Submit6" NAME="Accion" tabindex="1" class="w3-btn w3-theme w3-hover-theme w3-round" form="FormPag2">			
						<% end if %>
					</div>	

                <%else%>
                <div class="w3-col l8  w3-left w3-padding" ></div>
                
                <div class="w3-col l2  w3-left w3-padding " >
						<input type='submit' VALUE ='Añadir' ID="Submit5" NAME="Accion" tabindex="2" class="w3-btn w3-theme w3-hover-theme w3-round" form="FormPag2" />			
                </div>						
                <%end if%>								
			
            <%
            sp=sPro
            if ed_linkVolver<>"" then 
               sPro=ed_LinkVolver 
            end if   
            ed_CalPar 1,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
            sPro=sp
            %>
          
			<div class="w3-col l2  w3-left w3-padding" >
                
			    <a href="<%=sPar %>" class="w3-btn w3-theme w3-hover-theme w3-round">Volver</a>
			</div>
	        <%
                sPro=sP
               %>
    </div>
</div>
	
<%
 if ed_ides=1 then     Vercampo
end Sub


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

Sub MenuDet(xSql,sMosTot,xTxt)
    
    select case ed_iPas
        case 1,7,8
        case else
            exit sub
    end select        
	Dim rs
	Dim gMenFil
	dim iW
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
	rs.Open xSql,conexion
	if rs.eof then %>
	    <br />
	    <center>
	    <div style="width:30%; border:solid 1px #ff0000; border-radius:5px; height:30px; padding:20px 0 0 0; font-weight:bold; text-align:center ">
	        Advertencia: No hay Datos
	    </div>	
	    </center>
	<%
	    exit sub
	end if
	gMenFil=rs.GetRows
	rs.close
	'cGris1="#800000"
	iAncMen=ianc*98/100
'	iwid=iAncMen/ed_iMAxMen
	%>
    
	
    <div>

<%
    iw=0
    if ed_iFil="" then ed_iFil=0
    ixl=0
    sMosTot=lcase(sMosTot)
    if sMosTot="n" and ed_iFil=0 then ed_iFil = gMenFil(0,0)
    if sMosTot="s" then
        ed_CalPar ed_iPas,ed_iCla,1,"",ed_iCol,ed_iOrd, 0,ed_iMp,ed_iMs%>
       
            <%if ed_iFil=0 then
            iw=1%>
            <div title='<%=gMenFil(1,i)%>' style="border: 1px solid #ccdbe4; background:#F5F8FA ; width:24%;float:left;padding:5px 0px 5px 0px; text-align:center;text-decoration: none; font-size:12px; "   class="w3-container">
               <%=xTxt%>
            </div>
         <%else %> 
             <div title='<%=gMenFil(1,i)%>' style="border: 1px solid #ccdbe4;background:#ffffff ;width:24%;float:left;padding:5px 0px 5px 0px; text-align:center;text-decoration: none;font-size:12px; "   class=" w3-container">
                <a href="<%=sPar%>" style="text-decoration:none" >
                <%=xTxt%></a>
            </div>     
         <%end if           
       
       
    end if         
    
    for i=0 to ubound(gMenFil,2)
  
        ed_CalPar ed_iPas,ed_iCla,1,"",ed_iCol,ed_iOrd, gMenFil(0,i),ed_iMp,ed_iMs
          
        ix=ed_iFil-gMenFil(0,i)
        if ix=0 then
            iw=1%>
            <div title='<%=gMenFil(1,i)%>' style="border: 1px solid #ccdbe4; background:#F5F8FA ; width:24%;float:left;padding:5px 0px 5px 0px; text-align:center;text-decoration: none;font-size:12px; "   class="w3-container">
            
               <%=gMenFil(1,i)%>
            
		    </div>  
         <%else %> 
            <div title='<%=gMenFil(1,i)%>' style="border: 1px solid #ccdbe4;background:#ffffff ;width:24%;float:left;padding:5px 0px 5px 0px; text-align:center;text-decoration: none;font-size:12px;"   class=" w3-container">

                <a href="<%=sPar%>" style="text-decoration:none" >
                <%=gMenFil(1,i)%></a>
            </div>     
         <%end if %>
 
    <%next
     if iw=0 then ed_iFil=0

     %>
 
    </div> 
<%
	

End sub


Sub ed_vCombo


    dim rst    
    
	set rst = server.CreateObject("ADODB.Recordset")
	rst.CursorType = 1
	rst.LockType = 1
	'response.write "<br>2423 ed_sCombo(i,4):= " & ed_sCombo(i,4)

%>
    <table width="100%" cellspacing="1" cellpadding="0"  bgcolor="#ffffff" align="left" class="w3-theme-d2">
<%    	
	for i=1 to ed_iCombo
		
		'response.write "<br>" & ed_iCombo
		
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

        'if ubound(gX,2)<>0 then
       %>
        
   
	    <tr>
	        <td width="<%=ed_sCombo(i,5)%>" align="right" style=" font-family: Calibri; font-size:20px; padding: 0px 10px 0 0 ">
	        <%=ed_sCombo(i,0)  %>:
            </td>
	        <td width="<%=ed_sCombo(i,6)%>" style="font-size:20px; background-color:#ffffff; color:#000000;">
	        <%if ed_iPas<> 4 then %>
	    	    <select size="1" name="per" id="Select1"  onchange ="location.href=this.options[this.selectedIndex].value"  style="width:100%; font-size:20px; font-family:Tahoma; padding:3px 0 3px 0 ">
            <%  if ed_sCombo(i,2)<>"" then   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=ed_sCombo(i,2)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP

            %>
        			<option value="<%=sPar%>"  <% if ed_sPar(i,0) =ed_sCombo(i,2) then response.Write"selected" %> style="width:100%; font-size:20px; font-family:Tahoma; padding:5px 0 5px 0 " >
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
			    <%next%>
			   
	            </select>
	           
	              <%
	              
	              if wO=0 then 
			        if ed_sCombo(i,2)<>"" then 
                        ed_sPar(i,0)=ed_sCombo(i,2)
                    else   
			            ed_sPar(i,0)=ed_sPar(i,1) 
			           '  Response.write "<br />2509 wO:=" & WO & " " & ed_Spar(1,0)
			        end if    
			     end if%>
	       <% else 
	            %>
	                <%=ed_sPar(i,0) %>
	            <%
	          end if %>
	    </td> 
	   </tr>
	    <%
		'end if 
		%>         
<%	Next%>
    </table>
<%    
            		       
End sub

Sub ed_vCombo1


    dim rst    
    
	set rst = server.CreateObject("ADODB.Recordset")
	rst.CursorType = 1
	rst.LockType = 1
	'response.write "<br>2423 ed_sCombo(i,4):= " & ed_sCombo(i,4)

%>

<%    	
	for i=1 to ed_iCombo
		
		'response.write "<br>" & ed_iCombo
		
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

        'if ubound(gX,2)<>0 then
       %>
        
   
	        <%if ed_iPas<> 4 then %>
	    	    <select size="1"  name="per" id="Select1"  onchange ="location.href=this.options[this.selectedIndex].value" >
            <%  if ed_sCombo(i,2)<>"" then   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=ed_sCombo(i,2)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP

            %>
        			<option value="<%=sPar%>"  <% if ed_sPar(i,0) =ed_sCombo(i,2) then response.Write"selected" %> style="width:100%; font-size:20px; font-family:Tahoma; padding:10px 0 10px 0 " >
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
			    <%next%>
			   
	            </select>
	           
	              <%
	              
	              if wO=0 then 
			        if ed_sCombo(i,2)<>"" then 
                        ed_sPar(i,0)=ed_sCombo(i,2)
                    else   
			            ed_sPar(i,0)=ed_sPar(i,1) 
			           '  Response.write "<br />2509 wO:=" & WO & " " & ed_Spar(1,0)
			        end if    
			     end if%>
	       <% else 
	            %>
	                <%=ed_sPar(i,0) %>
	            <%
	          end if %>
	    <%
		'end if 
		%>         
<%	Next%>

<%    
            		       
End sub


Sub ed_MenPri
   
    ' Abrir Perfil
	set rs = CreateObject("ADODB.Recordset")
    rs.CursorType = 1
	rs.LockType = 1

    sql = "SELECT  id_PerfilUsuario, PerfilUsuario, link_acceso, ocultar, idGrupo , mostrar "
    sql = sql & " FROM  ss_U_PerfilUsuario "
	sql = sql & " WHERE (((id_Perfilusuario)=" & iPerUsu & ") AND ((Fec_Inactivo) Is Null)) "
    'response.write "<br>2196 sql:=" & sql    	
	rs.Open sql,conexion
	e_gPerUsu=rs.GetRows
	rs.close
	set rs = nothing

	
' Abrir Menu
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
	
    sql = "SELECT  idNivel, Menu, Link , Txt_Tips, Target, Parametros, id_menu, idgrupo, Ind_Activo, ind_Oculto "
    sql = sql & " FROM  ss_U_Menu "
	
	sOcultar=e_gPerUsu(3,0)
	sMostrar=e_gPerUsu(5,0)
	ed_iGrupo=e_gPerUsu(4,0)
'response.write "<br>ed_igrupo=" & ed_Igrupo	
    sql = sql & " WHERE ((IdGrupo)=" & e_gPerUsu(4,0) & ") AND ((Fec_Inactivo) Is Null) "
	if sOcultar<>"" then sql = sql & " AND (id_Menu NOT IN (" & sOcultar& "))"    
	if sMostrar<>"" then sql = sql & " AND (id_Menu IN (" & sMostrar& "))"    
	

    sql = sql & " ORDER BY Orden; "
'response.write "<br>2196 sql:=" & sql    
	rs.Open sql,conexion
	
 %>

    <table width="100%"  cellspacing="0"   cellpadding="0" border="0" align="center"  bgcolor="#ffffff">
    <tr ><td  style="background-color: #0066aa;"  >
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
     </table>

<%
end Sub 

Sub Ed_OutExcel
	 
     'exit sub
     Response.ContentType = "application/vnd.ms-excel"
     Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
     
End Sub

Sub Ed_Main

if ed_Link<>"" or ed_iRep=1 then
    ed_Bot(2)="disabled" ' Añadir
    ed_Bot(3)="disabled" ' Guardar
    ed_Bot(4)="disabled" ' eliminar
end if    
	if ed_iCol="" then ed_iCol=ed_cCol 'LuisReyes
	'if ed_iCol="" then ed_iCol=ed_iCol
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
		case 1,7,8 ' Listar
			ed_LeePag1 SqlCla
			if ed_ierr=1 then exit sub
			ed_VerPag1 ed_iRegPag, ed_iPag, ngCat
		
		case 2 ' Modificar
			ed_LeePag2 SqlReg
			ed_VerPag2 ngCat, 1
	'		ed_LeePag3
	'		ed_Verpag3

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
                    if ed_linkVolver<>"" then 
                       sPro=ed_LinkVolver 
                       ed_CalPar 1,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
                       response.redirect(sPar)
                    end if   

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
			ed_LeePag1 SqlCla
			ed_VerPag1 ed_iRegPag, ed_iPag, ngCat
			
	end Select%>
	
<%
End Sub


%>
<script>
function w3_open() {
    document.getElementById("mySidebar").style.display = "block";
}
function w3_close() {
    document.getElementById("mySidebar").style.display = "none";
}
function myAccFunc() {
    var x = document.getElementById("demoAcc");
    if (x.className.indexOf("w3-show") == -1) {
        x.className += " w3-show";
        x.previousElementSibling.className += " w3-blue";
    } else { 
        x.className = x.className.replace("w3-show", "");
        x.previousElementSibling.className = 
        x.previousElementSibling.className.replace(" w3-green", "");
    }
}
</script>

<script>
// Accordion
function myFunction(id) {
    var x = document.getElementById(id);
    if (x.className.indexOf("w3-show") == -1) {
        x.className += " w3-show";
       // x.previousElementSibling.className += " w3-theme-l1";
    } else { 
        x.className = x.className.replace("w3-show", "");
        x.previousElementSibling.className = 
        x.previousElementSibling.className.replace(" w3-theme-l1", " w3-theme-l1");
    }
}

// Used to toggle the menu on smaller screens when clicking on the menu button
function openNav() {
    var x = document.getElementById("navDemo");
    if (x.className.indexOf("w3-show") == -1) {
        x.className += " w3-show";
    } else { 
        x.className = x.className.replace(" w3-show", "");
    }
}
</script>
<%
Sub mMenu

' Abrir Perfil
	set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT  id_PerfilUsuario, PerfilUsuario, link_acceso, ocultar, idGrupo , mostrar"
    sql = sql & " FROM  ss_u_PerfilUsuario "
	sql = sql & " WHERE (((id_Perfilusuario)=" & iPerUsu & ") AND ((Fec_Inactivo) Is Null)) "
'response.write "<br>247 sql:=" & sql   
	
	rs.Open sql,conexion
	e_gPerUsu=rs.GetRows
	rs.close
	set rs = nothing
	
' Abrir Menu
	set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT  idNivel, Menu, Link , Txt_Tips, Target, Parametros, id_menu, idgrupo, Ind_Activo, ind_Oculto , icon, Path"
    sql = sql & " FROM  ss_u_Menu "
	
	sOcultar=e_gPerUsu(3,0)
	sMostrar=e_gPerUsu(5,0)
	
    sql = sql & " WHERE ((IdGrupo)=" & e_gPerUsu(4,0) & ")  AND ((Fec_Inactivo) Is Null) "
	if sOcultar<>"" then sql = sql & " AND (Cod_Perfil NOT IN (" & sOcultar& "))"    
	if sMostrar<>"" then sql = sql & " AND (Cod_Perfil IN (" & sMostrar& "))"    
	sql = sql & " AND (ind_activo='true') And (ind_Oculto = 'false' )"

    sql = sql & " ORDER BY Orden,Menu ; "
'response.write "<br>2196 sql:=" & sql    
	rs.Open sql,conexion

 %>

 

 
 <div class="w3-accordion w3-theme-l5">
 
         <% 
        if isnull(rs("Path")) then
           sPro= rs("Link").value 
        else
            sPro= "/" & rs("Path") & "/" & rs("Link").value 
        end if    
         spxx=sPro
            ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("Id_Menu").value, rs("Id_Menu").value
            if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
            sPro=spxx   
      
        iNiv= rs("idNivel").value
      
        i=0
        iDiv=0
        Do While NOT rs.EOF
  
               
              
               if rs("idNivel").value=0 Then ixMen1=rs("id_Menu").value
               if rs("idNivel").value=1 Then ixMen=rs("id_Menu").value

                if ed_ims="" then ed_Ims=0            
                ix=ed_iMs-rs("id_Menu").value
                
                sD="computer"
                sIcon=rs("Icon").value
                if isnull(sIcon) then sIcon=sD
                if rs("idNivel").value<1 Then
                    
                    %>
                    <% if i<>0 then %> 
                        </div> 
                         
                    <%    iDiv=1
                        end if 
                     if iNiv1=1 then%>
                        </div><%
                        iNiv1=0
                        iDiv=1
                     end if           
                                
                    
                     if rs("idNivel").value=0 then 
                        sCla="w3-btn-block w3-theme-l1 w3-left-align " 
                     else 
                        scla="w3-btn-block w3-theme-l4 w3-left-align "
                     end if   
                    'if rs("Menu").value="#" then%>    
                        <button onclick="myFunction('Mnu<%=rs("Id_Menu").value%>')" class="<%=sCla %>">
                        <i class="material-icons w3-large w3-margin-righ"><%=sIcon%></i>
                        <i class=" w3-margin-right "></i><%=rs("Menu").value%>
                        </button>              
         
      

                        <%sCla="w3-accordion-content w3-container"
                     '   response.write "<br>3418 ed_imp:=" & ed_imp & " id_menu:=" & rs("Id_Menu").value
                        if isnumeric(ed_imp) then 
                            ix=ed_imp -rs("Id_Menu").value
                        end if    
                        if ix=0 then  sCla=sCla & " w3-show " %>
                        <div id="Mnu<%=rs("Id_Menu").value%>" class="<%=sCla %>">
                        <%iMen=rs("Id_Menu").value
                        iDiv=0
                   'else%>
                     
                    <!--a href="<%=sPar%>" title="<%=rs("txt_tips").value%>"><%=rs("Menu").value %></a--><%
                  '  end if    
                    
                     %>
                <%else
                
                
                 
                         if rs("Menu").value<>"#" then
                           
                           if rs("Link").value ="#" then
                            'if  rs("idNivel").value=1 then
                                if iNiv1=1 then%>
                                    </div><%
                                    iDiv=1
                                end if
                               %>
                                <%sCla="w3-btn-block w3-theme-l4 w3-left-align w3-small"
                                ix=ed_ims -rs("Id_Menu").value
                                if ix=0 then  sCla=sCla & " w3-show " %>
                                <button onclick="myFunction('Mnu<%=rs("Id_Menu").value%>')" class="<%=sCla %>">
                                   <!--i class="material-icons w3-large w3-margin-righ"><%=sIcon%></i-->
                                   <i class=" w3-margin-right w3-small"></i><%=rs("Menu").value%>
                                </button>
                                <%sCla="w3-accordion-content w3-container"
                                  ix=ed_ims -rs("Id_Menu").value
                                if ix=0 then  sCla=sCla & " w3-show " %>
                                <div id="Mnu<%=rs("Id_Menu").value%>" class="<%=sCla %>">
                                <%iNiv1=1 
                                iDiv=0
                                iMen1=rs("Id_Menu").value%>
                            <%else                                       
                                if isnull(rs("Path")) then
                                    sPro= rs("Link").value 
                                else
                                    sPro= "/" & rs("Path") & "/" & rs("Link").value 
                                end if    
                                spxx=sPro                                
                                sx=spMen
                                spMen=rs("Txt_Tips")
                                ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,ixMen1,ixMen
                                if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
                                sPro=spxx                    
                                spMen=sx %>    
                         
                                <a href="<%=sPar%>" title="<%=rs("txt_tips").value%>" class="w3-small"><%=rs("Menu").value  %></a>
                             <% end if %>
                        <%else%>
                            <div class="w3-theme-l2" style="height:1px; "></div>
                         <% iDiv=0
                            end if    
                  
                   
                
                end if%>
                <% iNiv = rs("idNivel").value
            
            rs.MoveNext
            i=i+1 
        Loop
        
        rs.close
       
        if iDiv=0 then %>
           </div>
        <%
        end if    
         %>
        
   
</div>
<%exit sub %>
 <div class="w3-panel w3-light-grey w3-border w3-round">
    <p>London is the most populous city in the United Kingdom,
    with a metropolitan area of over 9 million inhabitants.</p>
  </div>
   <!-- Interests --> 
      <div class="w3-card-2 w3-round w3-white w3-hide-small">
        <div class="w3-container">
          <p>Interests</p>
          <p>
            <span class="w3-tag w3-small w3-theme-d5">News d5</span>
            <span class="w3-tag w3-small w3-theme-d4">W3Schools d4</span>
            <span class="w3-tag w3-small w3-theme-d3">Labels d3</span>
            <span class="w3-tag w3-small w3-theme-d2">Games d2</span>
            <span class="w3-tag w3-small w3-theme-d1">Friends d1</span>
            <span class="w3-tag w3-small w3-theme">Games</span>
            <span class="w3-tag w3-small w3-theme-l1">Friends l1</span>
            <span class="w3-tag w3-small w3-theme-l2">Food l2</span>
            <span class="w3-tag w3-small w3-theme-l3">Design l3</span>
            <span class="w3-tag w3-small w3-theme-l4">Art l4</span>
            <span class="w3-tag w3-small w3-theme-l5">Photos l5</span>
          </p>
        </div>
      </div>
      
      <br>
<%
end Sub 

Sub DataE%>
<div style="" class="w3-padding  w3-container" ><%ed_Main %></div>
<%
End Sub

%>