
<%Session.LCID=8202 
' Antes 1034
%>


<!DOCTYPE HTML>
<html>
<head>
	<title>Inicio</title>
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <meta name="Robots" content="noindex" >
    <meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="refresh" content="60" />
    <link href="lib/de.css" rel="stylesheet" type="text/css" media="screen" />     
    <link  href="lib/w3.css" rel="stylesheet">
	<link rel="icon" href="favicon.ico" type="image/x-icon">
    
</head>
<script  type="text/javascript" LANGUAGE="JavaScript">
	function EjecutarEXE() {
		alert("paso");
		WshShell = new ActiveXObject("WScript.Shell");
		WshShell.Run("C:\\Users\\administrador.SYNTONEX2008\\Desktop\\Luis\\PDF-Correo\\iTextSharpDemo.exe", 1, false);
		//WshShell.Run("iTextSharpDemo.exe");
		alert("paso1");
	}
	
	function abriracrobat(parametros) 
		{ 
			alert("paso2");
			var oShell = new ActiveXObject("Shell.Application"); 

			var aplicacion = "notepad"; 
			
				if (parametros != "") 
				{ 
					var parametros_del_comando = Form1.value; 
				} 
					oShell.ShellExecute(aplicacion, parametros_del_comando, "", "open", "1"); 
		}
	
 </script>
<body topmargin="0" >
<!--#include file="estiloscss.asp"-->
<!--#include file="meta.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="lib/in_KeyP.asp"-->


<%

dim inFal
dim inRec
dim gDatosSol

    Apertura

	'response.write "<br>49LLego"
	'response.end
    'LeePar

	'response.write "<br>12 pr_inicio"
	'response.end
    Encabezado
    
   
		%>
	</table>
            <br><br><br><br><br>
			<br><br><br><br><br>
			<div class="row">
				<!--Contenido General-->
					<img alt="Atenas" src="images/logo/LogoAtenasNew01.jpeg" class="img-responsive center-block" >					
			</div>
			<br><br><br><br><br>
			
			
			<!--#include file="includes/piepagina.asp"-->	            

</body>
</html>
