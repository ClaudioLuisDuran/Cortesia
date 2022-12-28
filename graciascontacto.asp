<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META NAME="Title" CONTENT="Sidra Cortesia - Cortesia Cider">
<META NAME="Author" CONTENT="Claudio Duran">
<META NAME="Subject" CONTENT="Planta elaboradora de Sidras">
<META NAME="Description" CONTENT="Bodega elaboradora de Sidras. Ciders processing plant in Tunuyán, Mendoza, Argentina.">
<META NAME="Keywords" CONTENT="Sidra, Cider, Manzana, Apple, Tunuyán, mendoza, Aconcagua, Vino, Wines, Napatina, Napa, Argentina, Champagne">
<META NAME="Language" CONTENT="Spanish, English">
<META NAME="Revisit" CONTENT="1 day">
<META NAME="Distribution" CONTENT="Global">
<META NAME="Robots" CONTENT="All">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>.:: Sidra Cortesía :: Gracias! ::.</title>

<SCRIPT LANGUAGE="JavaScript">

<!-- Free JavaScript Rollover Buttons from -->
<!-- http://www.creatupropiaweb.com -->

<!-- Begin

image1 = new Image();
image1.src = "images/home2.png";

image2 = new Image();
image2.src = "images/productos2.png";

image3 = new Image();
image3.src = "images/bodega2.png";

image4 = new Image();
image4.src = "images/historia2.png";

image5 = new Image();
image5.src = "images/distribuidores2.png";

image6 = new Image();
image6.src = "images/contacto2.png";


// End -->
</script>


<style type="text/css">
.auto-style1 {
	font-family: Arial;
	font-size: x-small;
	color: #570E0A;
}
.auto-style2 {
	font-family: Arial;
	font-size: medium;
	color: #570E0A;
}
</style>


</head>


<BODY STYLE="background-image:url('images/f2.png'); background-repeat:no-repeat; background-attachment: fixed" bgcolor="#FFFFFF">
</BODY>

<div align="center">
  <center>
  <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1" height="446" cellpadding="0">
    <tr>
      <td width="1052" height="28">
      <div align="center">
        <center>
        <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0">
          <tr>
            <td width="16%">
            
                        <a onmouseover="image1.src='images/home2.png';" onmouseout="image1.src='images/home1.png';" href="index_es.asp">
						<img name="image1" src="images/home1.png" border=0 width="132" height="28"></a>
  
<%
Dim cuerpo
Dim nombre
nombre = request("nombre")
Dim ciudad
ciudad = request("ciudad")
Dim provincia
provincia = request("provincia")
Dim email
email = request("email")
Dim telefono
telefono = request("telefono")
Dim comentarios
comentarios = request("comentarios")
Dim asunto
asunto = request("asunto")
Dim asuntook
asuntook = "[email generado en web Cortesia] " & asunto

cuerpo = " " & VbCrLf
cuerpo = cuerpo & "Nuevo contacto desde la web de Sidra Cortesía." & VbCrLf & VbCrLf
cuerpo = cuerpo & "Asunto:" & asunto & VbCrLf& VbCrLf
cuerpo = cuerpo & "Nombre:" & nombre & VbCrLf
cuerpo = cuerpo & "Ciudad:" & ciudad & VbCrLf
cuerpo = cuerpo & "Provincia:" & provincia & VbCrLf
cuerpo = cuerpo & "Telefono:" & telefono & VbCrLf
cuerpo = cuerpo & "Email:" & email & VbCrLf
cuerpo = cuerpo & "Comentarios:" & comentarios & VbCrLf


Dim objMail
Set objMail = CreateObject("CDONTS.NewMail")
objMail.From = email
'objMail.To = "duranclaudio@ciudad.com.ar"
objMail.To = "marcelo@napatinagroup.com"
objMail.Subject = asuntook
objMail.Body = cuerpo
objMail.importance = cdoHigh
objMail.Send
Set objMail = nothing

Dim objMail2
Set objMail2 = CreateObject("CDONTS.NewMail")
objMail2.From = email
'objMail2.To = "duranclaudio@ciudad.com.ar"
objMail2.To = "juan@napatinagroup.com"
objMail2.Subject = asuntook
objMail2.Body = cuerpo
objMail2.importance = cdoHigh
objMail2.Send
Set objMail2 = nothing


%>          
            
            
            </td>
            <td width="10%">
            
                                    <a onmouseover="image2.src='images/productos2.png';" onmouseout="image2.src='images/productos1.png';" href="productos.asp">
						<img name="image2" src="images/productos1.png" border=0 width="132" height="28"></a>
            
           </td>
            <td width="16%">
            
            <a onmouseover="image3.src='images/bodega2.png';" onmouseout="image3.src='images/bodega1.png';" href="bodega.asp">
						<img name="image3" src="images/bodega1.png" border=0 width="132" height="28"></a>
            
           </td>
            <td width="23%">
            
            <a onmouseover="image4.src='images/historia2.png';" onmouseout="image4.src='images/historia1.png';" href="historia.asp">
						<img name="image4" src="images/historia1.png" border=0 width="132" height="28"></a>            
            
            
            </td>
            <td width="18%">
            
             <a onmouseover="image5.src='images/distribuidores2.png';" onmouseout="image5.src='images/distribuidores1.png';" href="distribuidores.asp">
						<img name="image5" src="images/distribuidores1.png" border=0 width="132" height="28"></a>            
            
</td>
            <td width="17%">
            
                         <a href="contacto.asp" onmouseover="image6.src='images/contacto2.png';"
						onmouseout="image6.src='images/contacto1.png';">
						<img name="image6" src="images/contacto1.png" border=0 width="132" height="28"></a>     
						
          </td>
          </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td width="1052" height="141" valign="top">
      <div align="center">
        <center>
        <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="789" height="185" background="images/marcoamarillo.png" cellpadding="0">
          <tr>
            <td align="center" height="181">
            <img border="0" src="images/logook2.png"></td>
          </tr>
          </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td width="837" height="212" valign="top" background="images/bg3.jpg">
      <div align="center">
        <center>
        <table border="0" cellpadding="7" cellspacing="7" style="border-collapse: collapse" bordercolor="#111111" width="97%" id="AutoNumber1" height="177">
          <tr>
            <td width="100%" height="177" valign="top">
                  <p style="margin-left: 10; margin-right: 10">
                  <span lang="es"><font size="4" color="#570E0A"><b>Contacto</b></font></span></p>
                  <p style="margin-left: 10; margin-right: 10">
                  <img border="0" src="images/fotomanzanas.png" align="left" hspace="15" vspace="15"></p>
                  <p style="margin-left: 10; margin-right: 10" align="center">
                  <b><span lang="es">
                  <font face="Arial" size="2" color="#570E0A">¡</font></span><font face="Arial" size="2" color="#570E0A"><span lang="es">Gracias 
                  por su interés! </span></font></b></p>
                  <p style="margin-left: 10; margin-right: 10" align="center">
                  <font face="Arial" size="2" color="#570E0A"><span lang="es">El 
                  email ha sido enviado con éxito.</span></font></p>
                  <p style="margin-left: 10; margin-right: 10" align="center">
                  <font face="Arial" size="2" color="#570E0A"><span lang="es">Muy pronto nos contactaremos con usted.</span></font></p>
                  <p style="margin-left: 10; margin-right: 10" align="center">
                  &nbsp;</p>
                  </td>
          </tr>
        </table>
        </center>
      </div>
      <p align="center"><span lang="es"><font size="2" color="#570E0A"><b>
      <a href="donde.asp" style="text-decoration: none"><font color="#570E0A">
      Donde estamos</font></a>&nbsp;&nbsp; :::&nbsp;&nbsp;
      <a style="text-decoration: none" href="bodega.asp"><font color="#570E0A">
      La Bodega </font></a>&nbsp;&nbsp; :::&nbsp;&nbsp;&nbsp;
      <a href="productos.asp" style="text-decoration: none">
      <font color="#570E0A">Nuestros productos</font></a>&nbsp;&nbsp; :::&nbsp;&nbsp;
      <a href="news.asp" style="text-decoration: none"><font color="#570E0A">
      Noticias</font></a>&nbsp;&nbsp; :::&nbsp;&nbsp;
      <a href="mendoza.asp" style="text-decoration: none"><font color="#570E0A">
      Conozca Mendoza</font></a>&nbsp;&nbsp; :::&nbsp;&nbsp;
      <a href="contacto.asp" style="text-decoration: none">
      <font color="#570E0A">Contáctenos</font></a></b></font></span><p align="center">
      <span lang="es"><font size="2" color="#570E0A"><b>[ </b>
      <a href="en/index.asp" style="text-decoration: none">
      <font color="#570E0A">English version</font></a><b> ]</b></font></span></td>
    </tr>
    <tr>
      <td width="837" height="109" valign="top" background="images/bg3.jpg">
      <div align="center">
        <center>
        <table border="0" cellspacing="6" style="border-collapse: collapse" width="749" cellpadding="6" height="87">
          <tr>
            <td height="87" bgcolor="#FCF4D8" width="749">
                        <hr color="#DDDDDD" width="333">
            <p align="center"><span lang="es">
            <font face="Arial" size="2" color="#570E0A"><em>Tel. ( +54 2622 423897 
			) Por mal servicio de telefonía fija brindado por Telefónica de 
			Argentina,</em></font></span></p>
						<p align="center"><em><span class="auto-style1">por 
						favor para comunicarse con nosotros llamar al </span>
						<span class="auto-style2"><strong>+54 261 153049334.</strong></span><span class="auto-style1"> 
						Muchas gracias.</span></em></p>
            <p align="center">&nbsp;</p>
            <p align="center"><span lang="es">
            <font face="Arial" size="2" color="#570E0A">Información y ventas
            <a href="contacto.asp"><font color="#570E0A">haga click aquí </font>
            </a> </font></span>
                        <hr color="#DDDDDD" width="333">
                        <p align="center"><span lang="es"><!-- Histats.com  START  -->
<a href="http://www.histats.com/es/" target="_blank" title="contador" ><script  type="text/javascript" language="javascript">
var s_sid = 761501;var st_dominio = 4;
var cimg = 173;var cwi =85;var che =17;
                        </script></a>
<script  type="text/javascript" language="javascript" src="http://s11.histats.com/js9.js"></script>
<noscript><a href="http://www.histats.com/es/" target="_blank">
<img  src="http://s103.histats.com/stats/0.gif?761501&1" alt="contador" border="0"></a></noscript>
<!-- Histats.com  END  --></span></td>
          </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
  </table>
  </center>
</div>





</html>