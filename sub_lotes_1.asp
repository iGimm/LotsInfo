<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="include/MasterHead1.asp"-->
<!--#include file="include/funciones.asp"-->
<!--#include file="include/dg_verificador.asp"-->
<!--#include file="idiomas/publicas_.asp"-->

<%
id_subasta = Vdato(request.QueryString("id_subasta"))
if id_subasta = "S201101061" then 
%>
<script language="javascript">
alert('Esta subasta no tiene lotes disponibles, contacte a Myron Bowling México para mayor información. \nG R A C I A S .');
document.location='sub_index.asp?id_subasta=<%=id_subasta%>';
</script>
<%
response.End()
end if

if len(id_subasta) > 20 then
response.Write("No se encuentra la subasta solicitada.")
response.End()
end if

	if FunEtiquetaIdioma_idioma = "ing_us" then
	sqlCmpSca_nom = "(select top 1 sca_nom_en from subcateg where sca_clave = sub_categoria)"
	else
	sqlCmpSca_nom = "(select top 1 sca_nom from subcateg where sca_clave = sub_categoria)"
	end if


funing.commandtext = "select subinfo.*, tipmon.*, "&sqlCmpSca_nom&" as sca_nom  "&_
					"from subinfo, tipmon "&_
					"where (sub_activarPortal='SI' or sub_activarPortal='HI') "&_
					"and sub_clave = '"&id_subasta&"' "&_
					"and tmo_clave = sub_clave_tmo "
set rs = funing.execute
if rs.eof and rs.bof then
%>
<script language="javascript">
alert('La subasta elejida no existe *<%=id_subasta%>*');
history.back();
</script>
<%
else
sub_nom = rs("sub_nom")
sub_txtwww1 = rs("sub_txtwww1")
sub_nom_en = rs("sub_nom_en")
sub_txtwww1_en = rs("sub_txtwww1_en")
sub_fecsub = rs("sub_fecsub")
sub_fecCierre = rs("sub_fecCierre")
sub_domca = rs("sub_domca")
sub_domco = rs("sub_domco")
sub_domcp = rs("sub_domcp")
sub_domde = rs("sub_domde")
sub_domes = rs("sub_domes")
sub_dompa = rs("sub_dompa")
sub_domte = rs("sub_domte")
sub_contacto = rs("sub_contacto")
usarangos = rs("sub_usarangos")
sub_webMuestraPSalida = rs("sub_webMuestraPSalida")
sub_GaleriasDesde = rs("sub_GaleriasDesde")
sub_muestraCantidad = rs("sub_muestraCantidad")
sub_activarPortal = rs("sub_activarPortal")
	if instr(1, rs("tmo_nom"), "|") > 0 then 
		tmo_nom = rs("tmo_nom")
	else
		tmo_nom = "PESO MX.|PESOS"
	end if
	tipMonArr = split(tmo_nom, "|", -1, 1)

sca_nom = rs("sca_nom")
end if
rs.close
if usarangos = "SI" then
r_piv = 0
sqltxt = "Select * from subrangos where sra_clave_sub='"&id_subasta&"'"
funing.commandtext = sqltxt
set rs = funing.execute()
if (rs.EOF and rs.BOF) then
'response.Write("No HAY RANGOS")
else
 while not rs.eof
 r_piv = r_piv + 1
 rs.movenext
 wend
dim rangos()
redim rangos(r_piv,5)
 rs.movefirst
 r_piv = 0
 while not rs.eof
 r_piv = r_piv + 1
	if FunEtiquetaIdioma_idioma = "ing_us" then
	 rangos(r_piv,0)=rs("sra_nom_en")
	else
	 rangos(r_piv,0)=rs("sra_nom")
	end if
 rangos(r_piv,1)=rs("sra_mon1")
 rangos(r_piv,2)=rs("sra_mon2")
 rangos(r_piv,3)=rs("sra_mdu")
 rangos(r_piv,4)=rs("sra_ocxdu")
 rangos(r_piv,5)=rs("sra_clave")
 rs.movenext
 wend
end if
rs.close
end if

'vERIFICA LA EXISTENCIA DE LAS CARPETAS DE TRABAJO SUBASTA
set fso = server.CreateObject("Scripting.FileSystemObject")
dir_subasta = dir_dv&"\Fotos_Lotes\"&id_subasta

%>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<BODY style="BACKGROUND-IMAGE: none; BACKGROUND-COLOR: #3b3b3b">
<CENTER>
<!--#include file="header1.asp"-->
<!--#include file="header2.asp"-->
<!-- AQUI EMPIEZA CONTENIDO -->
<div class="container">
<div id="main-content" class="main-content">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<!-- PRIMER COLUMNA -->
<tr>
<td align="left" height="200" valign="top" width="90">&nbsp;
</td>
<td align="left" height="200" valign="top" width="20">&nbsp;
</td>
<td aling="center" height="200" valing="top" width="871">
<!-- #########AQUI IVAN LAS SUBASTAS############# -->
<DIV><b><font face="verdana" size="3" color="red"><%if FunEtiquetaIdioma_idioma = "ing_us" then
								response.Write(sub_nom_en)
								else
								response.Write(sub_nom)
								end if%></font></b></DIV>
<hr />
<br />
<!-- pegue todo -->

<div id="PgContenido">
 <div id="PgContenidoSubLista">

<%
	set fso = server.CreateObject("Scripting.FileSystemObject")
	dir_subasta = dir_dv&"\Fotos_Lotes\"&id_subasta
	MuestraFlashGaleria = false
	if fso.FolderExists(dir_subasta) then
		funing.commandtext = "select lot_nom,lot_clave "&_
		"from lotes where lot_clave_sub = '"&id_subasta&"'"
		set rs = funing.execute
		if not (rs.eof and rs.bof) then
			dat_lotes = rs.getrows()
			for lpiv = 0 to Ubound(dat_lotes,2)
				dir_lote = dir_subasta&"\"&dat_lotes(1,lpiv)
				'Response.Write("<br />dir_lote: "&dir_lote&VbCrlf)
				if fso.FolderExists(dir_lote) then
					set dir_comp = fso.getfolder(dir_lote)
					piv = 0
					For each dir_comp In dir_comp.Files 
					piv = piv  + 1
						if instr(1,dir_comp.name,".",1) > 0 then
						arch_arr = split(dir_comp.name,".",-1,1)
						imagen = dir_lote&"\"&dir_comp.name
							if lcase(arch_arr(1)) = "gif" or lcase(arch_arr(1)) = "jpg" and (instr(Ucase(arch_arr(0)),"NO-DISPONIBLE") = 0) then
							MuestraFlashGaleria = true
							end if
						if MuestraFlashGaleria then exit for
						end if
					next		
				if MuestraFlashGaleria then exit for
				end if
			next
		end if
		rs.close
	end if
if MuestraFlashGaleria then
%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="590" height="220" title="Myron Bowling México - Subastas">
  <param name="movie" value="BannerTopSub.swf?var1=<%=id_subasta%>">
  <param name="quality" value="high">
  <param name="wmode" value="opaque">
  <embed src="BannerTopSub.swf?var1=<%=id_subasta%>" quality="high" wmode="opaque" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="590" height="220"></embed>
</object>
<!-- Filter Fields //-->
<%else%>
<!--include file="include/MasterBanner1.asp"-->
<%end if%>


<%if sub_habilitaRadio = "SI" then%>
    <div align="right"><A HREF="javascript://" onClick="window.open('RadioPlayer/index.html', 'RadioPlayer', ',width=500,height=300,resizable=yes,scrollbars=no,menubar=no'); return true;"><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Webcast del Evento")%></A>    
        <img src="img/i/icono_audio.jpg" alt="Escuchar Audio del Evento" width="45" height="45" border="0" align="absmiddle">
    </div>
<%end if%>

<%if sub_requiPreRegis = "SI" and sub_activarPortal="SI" then%>
<div class="SubastaEnlaces" align="right">			 
<%					  if segura(0) then%>
					  <A href="Registro_ASubasta_1.asp?id_subasta=<%=id_subasta%>"><img src="img/i/hospedaje.gif" alt="<%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%>" title="<%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%>" width="32" height="32" border="0" align="absmiddle"><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%></A>
					  <% else%>
					  <A href="verifica_codigob.asp?s=<%=id_subasta%>&MuestraLogIn=no&d=SubInscripcion"><img src="img/i/hospedaje.gif" alt="<%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%>" title="<%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%>" width="32" height="32" border="0" align="absmiddle"><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%></A>
					  <%end if%>
   </div>
<%
  end if
%> 
            

<br>


      <table cellspacing="0" cellpadding="20">
         <tr style="vertical-align: top;">
            <td>

                  <DIV class="SubastaTitPropietario">
		<%
	funing.commandtext = "select distinct(cta_rsocial), lot_clave_cta "&_
	"from lotes, cuentas where lot_clave_sub='"&id_subasta&"' "&_
	"and cta_clave=lot_clave_cta order by cta_rsocial asc"
	set rpro = funing.execute
	sub_cantPro = 0
	NombreProp = ""
	if rpro.eof and rpro.bof then
%>PENDIENTE
        <%	
	else
	reg_arr = rpro.GetRows()
	sub_cantPro = Ubound(reg_arr,2)+1
	if sub_cantPro >= 3 then
		for i = LBound(reg_arr,2) to Ubound(reg_arr,2)
%>
 <!--em><font size="1"><%=reg_arr(0,i)%></font></em><br-->
 <img src="img/logos_pro/logo_<%=reg_arr(1,i)%>.gif" width="97" alt="<%=reg_arr(0,i)%>" border="0"/> 
<%
		if NombreProp <> "" then NombreProp = NombreProp + ", "
		NombreProp = NombreProp + reg_arr(0,i)
		next
	elseif sub_cantPro = 2 then
		for i = LBound(reg_arr,2) to Ubound(reg_arr,2)
%>
 <!--em><font size="1"><%=reg_arr(0,i)%></font></em><br-->
 <img src="img/logos_pro/logo_<%=reg_arr(1,i)%>.gif" width="121" alt="<%=reg_arr(0,i)%>" border="0"/> 
<%
		if NombreProp <> "" then NombreProp = NombreProp + ", "
		NombreProp = NombreProp + reg_arr(0,i)
		next
	elseif sub_cantPro = 1 then
		for i = LBound(reg_arr,2) to Ubound(reg_arr,2)
%>
 <!--em><font size="1"><%=reg_arr(0,i)%></font></em><br-->
 <img src="img/logos_pro/logo_<%=reg_arr(1,i)%>.gif" width="176" alt="<%=reg_arr(0,i)%>" border="0"/> 
<%
		if NombreProp <> "" then NombreProp = NombreProp + ", "
		NombreProp = NombreProp + reg_arr(0,i)
		next
	end if
'	response.Write("("&sub_cantPro&")")
	
	end if
	rpro.close
		%>
</DIV>            
            
            </td>
            <td align="left" >
                <DIV class="SubastaEnlaces">
				<%if FunEtiquetaIdioma_idioma = "ing_us" then
								response.Write(sub_nom_en)
								else
								response.Write(sub_nom)
								end if%>
                <BR>
                  
                <img src="img/spacer.gif" width="30" height="1" /> 
<a href="sub_index.asp?id_subasta=<%=id_subasta%>" class="SubastaEnlaces"> <img src="img/i/puntin.jpg" width="7" height="7" border="0" /> <%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Detalles de la Subasta")%></a>
                  
                  <%if sub_requiPreRegis= "SI" and sub_activarPortal="SI" then
					  if segura(0) then%>
					  <A href="Registro_ASubasta_1.asp?id_subasta=<%=id_subasta%>"  class="SubastaEnlaces"> <img src="img/i/puntin.jpg" width="7" height="7" border="0" /><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%></A>
					  <%
					  else%>
					  <A href="verifica_codigob.asp?s=<%=id_subasta%>&MuestraLogIn=no&d=SubInscripcion"  class="SubastaEnlaces"> <img src="img/i/puntin.jpg" width="7" height="7" border="0" /> <%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Registro al evento")%></A>
					  <%
					  end if
				  end if
				  %>   
              </DIV>
                  <DIV class="SubastaDetalles">
<%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Tipo de Moneda")%>: 
                              <%if FunEtiquetaIdioma_idioma = "ing_us" then
							     if instr(1, tmo_nom, "RES") > 0 then
								  response.Write(tipMonArr(1))
								  else
								  response.Write(tipMonArr(0))
								  end if
								else
								response.Write(tipMonArr(0))
								end if%><br />
<%
		if isdate(sub_fecsub) then
			if sub_fecsub > cdate("01/01/2001") then
			response.Write(FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Inicia")&": "&FormatDateTime(sub_fecsub,2)&" a las "&FormatDateTime(sub_fecsub,4)&" hrs. ")
			end if
		end if
		if isdate(sub_fecCierre) then
			if sub_fecCierre > cdate("01/01/2001") then
			response.Write(FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Cierra")&": "&FormatDateTime(sub_fecCierre,2)&" a las "&FormatDateTime(sub_fecCierre,4)&" hrs. ")
			end if
		end if
		%>
</DIV>                  
            </td>
         </tr>
</table>



<DIV class=page_title style="BACKGROUND-COLOR: gray;">
<table width="850" border="0" cellpadding="0" cellspacing="0">
<tbody align="center">
 <tr>
  <td ><span style="font-size: 12pt; font-weight: bold; color: white"><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Lotes a subastar")%></span></td>
  <td><IMG height="41" width="30" alt="" src="img/spacer.gif"></td>
 </tr>
 </tbody>
</table>      
</DIV>

    <table border="0" cellspacing="5">
<%
'---------------------variavles para funciones de paginación 1
intervalo = 100
complementos="&id_subasta="&id_subasta
'----------------------------comienzan el proceso de paginación
cantRegT = 0
piv = 0
lm1 = cint(Vdato(request.QueryString("lm1")))
lm2 = cint(Vdato(request.QueryString("lm2")))
if lm1 = 0 then
lm1 = 1
end if
if lm2 = 0 then
lm2 = intervalo
end if
'-----------------------fin de complemento


funing.commandtext = "select lotes.* "&_
" from lotes where lot_clave_sub = '"&id_subasta&"' order by lot_numensub asc"
set rs = funing.execute
if rs.eof and rs.bof then
%>
      <tr>
        <td class="txt1" colspan="2">No existen lotes registrados</td>
      </tr>
<%
else
lot_ubi = rs("lot_ubi")
 color = true

	dir_lote = dir_subasta&"\"&rs("lot_clave")


	dir_BieSubasta = dir_dv&"\Fotos_Bienes\"&id_subasta
	
	if not fso.FolderExists(dir_subasta) then
	dir_BieSubastaExiste = false
	else
	dir_BieSubastaExiste = true
	end if
	h = 150
	v = 100
	comp = 400

'---------------------complemento para funciones de paginación con algunos drivers de mysql faya el procedimiento recordcount se hace manualmente....
	reg_arr = rs.GetRows()
 	rs.movefirst()
	cantRegT = Ubound(reg_arr,2)+1 
'-----------------------fin de complemento
 while not rs.eof
'---------------------complemento para funciones de paginación
piv = piv  + 1
if piv >= lm1 and piv <= lm2 then
'-----------------------fin de complemento

'verifica el folder root deL LOTE SOLICITADO, si no existe lo crea

'   fso.CreateFolder(dir_lote)
 '  response.Write("Se creo el espacio de lote<br>")
'end if


%>
      <tr <%if color then response.Write("class='alternate'")%> valign="top">
        <td>
      <%
if sub_GaleriasDesde = "LOTE" then
	'fotos del lote
	dir_lote = dir_dv&"\Fotos_Lotes\"&id_subasta&"\"&rs("lot_clave")

	If (fso.FolderExists(dir_lote)) Then
	'response.Write(dir_lote)
		set dir_comp = fso.getfolder(dir_lote)
		For each dir_comp In dir_comp.Files 
			if instr(1,dir_comp.name,".",1) > 0 then
			arch_arr = split(dir_comp.name,".",-1,1)
			imagen = dir_lote&"\"&dir_comp.name
			if lcase(arch_arr(1)) = "gif" or lcase(arch_arr(1)) = "jpg" then
			'response.Write("<br />"&dir_comp.name)
			'response.Write("<br />"&imagen)
			%><a href="javascript:MM_openBrWindow('sub_lotes_2.asp?id_lote=<%=rs("lot_clave")%>','lote','status=yes,scrollbars=yes,resizable=yes,width=900,height=700')"><img src="r.aspx?command=resize&abspaths=OVA&width=100&src=<%=imagen%>" border="0" /></a>
			<%
			exit for
			end if
			end if
		next		
	end if
else
 'Bienes
	if dir_BieSubastaExiste then
	
	funing.commandtext = "select bie_clave, bie_cantidad, bie_descripcion "&_
	"from bienes where bie_clave_lot = '"&rs("lot_clave")&"' and bie_clave_sub = '"&id_subasta&"' order by bie_m_pbv, bie_clave"
	set rsBie = funing.execute
	if rsBie.eof and rsBie.bof then
		response.Write("Fotos Pendientes - Nr")
		ContinuaBuscandoBienes = false
	else
		dat_bie = rsBie.getrows()
		ContinuaBuscandoBienes = true
	end if
	rsBie.close
	if ContinuaBuscandoBienes then
	for lpiv = 0 to Ubound(dat_bie,2)
		dir_bien = dir_BieSubasta&"\"&dat_bie(0,lpiv)
		'Response.Write("<br />("&dir_dv&"\Fotos_Bienes\"&id_subasta&"\"&dat_bie(0,lpiv)&")"&VbCrlf)
		if fso.FolderExists(dir_bien) then
			set dir_comp = fso.getfolder(dir_bien)
			For each dir_comp In dir_comp.Files 
			'response.Write(dir_comp.name)
				if instr(1,dir_comp.name,".",1) > 0 then
				arch_arr = split(dir_comp.name,".",-1,1)
					if lcase(arch_arr(1)) = "gif" or lcase(arch_arr(1)) = "jpg" and (instr(Ucase(arch_arr(0)),"NO-DISPONIBLE") = 0) then
				'	response.Write(dir_lote)
				'	response.Write(dir_comp.name)
			%><a href="javascript:MM_openBrWindow('sub_lotes_2.asp?id_lote=<%=rs("lot_clave")%>','lote','status=yes,scrollbars=yes,resizable=yes,width=900,height=700')"><img src="r.aspx?command=resize&width=<%=h%>&src=/web/img/GaleriaBienes/<%=id_subasta%>/<%=dat_bie(0,lpiv)%>/<%=dir_comp.name%>"></a><%
					ContinuaBuscandoBienes = false
					exit for
					end if
				end if
			next		
		else
		'response.Write("<br>"&dir_bien&"<br>")
		end if
		if not ContinuaBuscandoBienes then exit for
	next
	end if
	if ContinuaBuscandoBienes then response.Write(FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Fotos Pendientes")&" - Nf")
	else
		response.Write(FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Fotos Pendientes")&" - Lote")
	end if
end if
%>		</td>
		<td align="left">

<a href="#" style="font-style: italic" onClick="MM_openBrWindow('sub_lotes_2.asp?id_lote=<%=rs("lot_clave")%>','lote','status=yes,scrollbars=yes,resizable=yes,width=900,height=700')"><strong>Lote #<%=Funlot_num_f(rs("lot_numensub"))%></strong></a> 
<%
if isnumeric(rs("lot_cant")) and sub_muestraCantidad="SI"  then
		if rs("lot_cant")>1 then response.Write(" "&rs("lot_cant")&" ")
	end if
	if FunEtiquetaIdioma_idioma = "ing_us" then
	 response.Write(rs("lot_nom_en"))
	else
	 response.Write(rs("lot_nom"))
	end if
	%><br />
		<strong><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Ubicado en")%>:</strong> <%=rs("lot_ubi")%> <br><span style="font-style: italic"><strong>
        <!-- AQUI PARA AGREGARLE PALABRA VENDIDO A LA LISTA DE LOTES -->
		<%
		if id_subasta="S201409271" then 
		nlot = rs("lot_numensub")
        if nlot = 11 or nlot = 12 or nlot = 13 or nlot = 14 or nlot = 15 or nlot = 16 or nlot = 17 or nlot = 18  then %>
        <span style="color:#FD0303"><strong>VENDIDO</strong></span>
		<%
		end if
		end if
		%>
        <%
		if id_subasta="S201410151" then 
		nlot = rs("lot_numensub")
        if nlot = 1 then %>
        <span style="color:#FD0303"><strong>VENDIDO</strong></span>
		<%
		end if
		end if
		%>
        <!-- ASTA AQUI PARA AGREGARLE PALABRA VENDIDO A LA LISTA DE LOTES -->
		<%
        if isnumeric(rs("lot_m_pbv")) then 
			if rs("lot_m_pbv") > 1  and sub_webMuestraPSalida = "SI" then 
			response.Write(FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Salida")&": "&formatCurrency(rs("lot_m_pbv"))&" "&tipMonArr(1))
			else
			'response.Write(FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Salida")&": RESERVADO (RESERVED)")
			end if
		end if
		%></strong></span>          
<%if usarangos = "SI" then%>
    <br><span><%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Garantía")%>:  
      <%
	enc_r = false
	sra_Num = 0
    if rs("lot_rango_f")="AUTO" then
		if r_piv <> 0 then
			for y = 1 to r_piv
				if rs("lot_m_pbv") >= rangos(y,1) and rs("lot_m_pbv") <= rangos(y,2) then
				response.Write(rangos(y,0)&" &nbsp;&nbsp; ")
				enc_r = true
				exit for
				end if
			next
			if not enc_r then response.Write("ESTA FUERA LOS RANGOS DE LA SUBASTA")
		else
		response.Write("No hay rangos de precios asignados")
		end if
	else
		sqltxt = "Select sra_clave, sra_nom from subrangos where sra_clave='"&rs("lot_rango_f")&"'"
		dat_rd = info_tablaR3(sqltxt, existeRD)
		if not existeRD then
			response.Write("El RANGO ES FIJO PERO NO EXISTE EL REGISTRO("&rs("lot_rango_f")&")")
			response.End()
		else
			for y = 1 to r_piv
				if rangos(y,5) = dat_rd(0,0) then
				response.Write(rangos(y,0)&" &nbsp;&nbsp; ")
				enc_r = true
				sra_Num = y
				exit for
				end if
			next
			if not enc_r then response.Write("ESTA FUERA LOS RANGOS DE LA SUBASTA")
		end if

	end if	  
	%>	
  </span>
<%
end if
%>

            </td>
      </tr>
<%
 color = not color
'----------------------complemento para la función de paginación
end if
'----------------------fin del complemento para la función de paginación
 rs.movenext
 wend
end if
rs.close
%>
    </table> 


<table>
  <tr>
   <td><%FunReAuto lm1,lm2,cantRegT,intervalo,complementos, FunEtiquetaIdioma_idioma%> </td>
   <td><%FunAvAuto lm1,lm2,cantRegT,intervalo,complementos, FunEtiquetaIdioma_idioma%> </td>
  </tr>
</table>
<p class="txt1">
</p>
<p class="txt1">
<%=FunEtiquetaIdioma(FunEtiquetaIdioma_idioma,"Páginas")%>:<br />
<%
FunListaLimites cantRegT,intervalo,complementos
%>
</p>
<!--  AQUI PARA AGREGARLE FILTRO DE BUSQUEDA --> 
<P>
FILTRO: 
<SELECT name="buscar"> 
<option value="TODO"> TODOS </option>
<%
'dim cont7 = 0
'while Ubound(lot_ubi) >= cont7
%>
<option value="<%=lot_ubi %>"> <%=lot_ubi %> </option>
<% 
'cont7 = cont7 + 1
'wend
 %>
</SELECT>
</P>
<!--  ASTA AQUI PARA AGREGARLE FILTRO DE BUSQUEDA --> 
 </div>
 <div id="PgContenidoMenu">
    <!--include file="include/MasterMenu1.asp"-->
 </div>
</div>

<!-- asta aqui pegue todo -->
</td>
<td align="left" height="200" valign="top" width="10">
<br />
</td>
</tr>
<!-- SEGUNDO COLUMNA -->
<tr>
</tr>
</table>
</div>
</div>
<!-- AQUI TERMINA CONTENIDO -->

<!--#include file="include/MasterPie1.asp"-->

</CENTER>
</BODY>
</HTML>

