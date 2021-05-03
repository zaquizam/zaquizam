<!DOCTYPE HTML>
<html >
<head>
	<title>Subir Archivo</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="meta.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================


    Apertura
%>
<script>
  function fnUpload() 

   {
 var client;
    if (window.XMLHttpRequest)
      {client=new XMLHttpRequest();}
    else if (window.ActiveXObject)
            {client=new ActiveXObject("Microsoft.XMLHTTP");}
    else
        {alert("Your browser does not support XMLHTTP!");}
    var file = document.getElementById("fileElem");
    
    var formData = new FormData();
    formData.append("upload", file.files[0]);
    client.open("post", "rge_pSubirArchivo.asp", true);
    client.setRequestHeader("Content-Type", "multipart/form-data");
    client.send(formData);  /* Send to server */ 
    alert("Archivo cargado");
   
   }
</script>
<!--#include file="xelupload.asp"-->
<style type="text/css">
#fileElem {
  visibility: visible;
  

}

</style>
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    if ed_iPas<>4 then 
        Encabezado
    end if    

    
    iOpc=1

	Dim up, fich
	set up = new xelUpload
	up.Upload()
	response.write "llego1"
	response.end


if up.Ficheros.Count=0 then
    mFormato
else 
    Subir
    sp=sPro
    select Case iOpc
        case 1
            sPro="PH_CB_IncluirProductos.asp"
    end select    
    CalPar
    sPro=sp
    
    if iOpc<>3 then    
    %>
    <div style="width:40%; margin-left:auto; margin-right:auto; text-align:center; margin-top:20px">
    <div class="ed_boton1">
    <a href="<%=sPar%>" >Presione para continuar</a></div>
    </div>
    <%
    end if
end if    

Function DelayEnSegundos(Segundos)
	Dim Inicio
	Dim Fin
	Inicio = Timer()
	Fin = Inicio + Segundos

	Do While Inicio < Fin
	Inicio = Timer()
	Loop
End Function

Sub mFormato
	calpar
	
	
	%>
	<div style="border:solid 1px #ffff00; margin-top:30px ;width:50%; height: auto; margin-left:auto; margin-right:auto; text-align:center; border-radius:5px">
    <div style="border:solid 1px #fffffff; width:80%; height: auto; margin-left:auto; margin-right:auto; text-align:center">
	    <p style=" font-size:14px; font-weight:bold; font-family:Verdana;" >Selecciona el archivo y luego presiona "Cargar"</p>
    </div>	
    
	<%
	'response.write "<br>92 spar:= "&spar
	'response.write "<br>92 iOpc:= "&iOpc
	%>	
	<div style="border:solid 1px #ffffff; width:60%; height:100px; margin-top:10px; margin-left:auto; margin-right:auto; text-align:left">
	<form action="<%=spar&"&opc="&iOpc%>" method="post" enctype="multipart/form-data">
		
			<div style="border:solid 1px #ffffff; margin-top:10px; margin-left:0px">
				<input type="file" name="fichero" size="20"  />
			</div>
			<div style=" border:solid 1px #ffffff;margin-top:20px; margin-left:0px">
				<button id="fileSelect" class="rs_botdiv8">Cargar</button>
			</div>
	</form>
	</div>
	</div>
	<%
end Sub 
Sub Subir

		%>
			<br/>
			<br/>
			<br/>
			<div id='DivProc'  align="center"> 
				<img src="/images/Procesando01.gif" alt="Procesando....."/>
			</div>
		<%

		%>
			<script language="javascript">
				document.getElementById("DivProc").style.visibility="visible";
			</script>
		<%
	'Dim up, fich
	'set up = new xelUpload
	'up.Upload()
	'response.write "<br>93 ipCli:= " & ipCli
	'response.write "<br>94 Opcion:= " & iOpc
	'Response.Flush	
	dim sNombreArc
	For each fich in up.Ficheros.Items
		DelayEnSegundos 2
		response.flush
		'Para guardarlo
		'	Con el nombre de fichero original:
		'	fich.Guardar Server.MapPath("Upload")	
		'	Con otro nombre:
		'	fich.GuardarComo nombrefichero, Server.MapPath("Upload")
		'----------------------------------------------------------------
		Select Case iOpc
			Case 1
				sNombreArc = "ExcelProd" & ".xls"
				'fich.GuardarComo sNombreArc, Server.MapPath("Upload")
				'fich.GuardarComo sNombreArc, "http://datos.rtandna.com/upload"
				fich.GuardarComo sNombreArc, Server.MapPath("Upload")
				
		End Select
        %>
        <div style="border:solid 1px #ffff00;width:40%; margin-left:auto; margin-right:auto; text-align:left; border-radius:5px">
        <%
		'Response.Write("Número de ficheros subidos: " & up.Ficheros.Count & "<br>")
		Response.Write("<ul>")
		Response.Write("<li>Nombre: <b>" & fich.Nombre & "</b></li>")
		Response.Write("<li>Tamaño: <b>" & fich.Tamano & "</b> bytes (" & FormatNumber(fich.Tamano / (1024*1024)) & " Mb)</li>")
	'	Response.Write("<li>Tipo MIME: <b>" & fich.TipoContenido & "</b></li>")
	    Response.Write("<li>Guardado como: <b>" & sNombreArc & "</b></li>")
		Response.Write("</ul>")
	'	response.write "<br>93 ipCli:= " & ipCli
	'	response.write "<br>94 Opcion:= " & iOpc
		DelayEnSegundos 2
		response.flush
		%>
		</div>
		
			<script language="javascript">
				document.getElementById("DivProc").innerHTML="";
				document.getElementById("DivProc").style.height=50;
				document.getElementById("DivProc").style.visibility="hidden";
			</script>
		<%

		%>
			<script language="javascript">
				window.alert("Carga....!\n\n\Realizada con Exito...!");
			</script>
		<%

	Next

	'Limpiamos objeto
	set up = nothing
	end sub

	
%>
		
	<br>
	<div style="width:98%">
	<%
	
	%></div></center>

    <%conexion.close%>
	


</body>
</html>