// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

  function cargar_datos_str() {

    let myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");

    let requestOptions = {
      method: 'POST',
      headers: myHeaders,
      body: "",
      redirect: 'follow'
    };
    
      //fetch("https://default73c9a419863d4226a83f7a200ad69b.e9.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/15/workflows/8e0dc4da953541c6af3b32bbe54b40e6/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PRdRmuPRX7vTs8X03RmspDc-7zvXsKkQ7jbRHuvId_E", requestOptions)        
      fetch("https://default73c9a419863d4226a83f7a200ad69b.e9.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/01/workflows/5bdb62f710c24557b91fee8a539c2534/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ORS8WbddWepHcGZH9uWJqcQwHvIYQ-4saDQ9V8DDcEk",requestOptions)
      .then(response => response.json())
        .then(result => { dataFirma(result) }) 
        .catch(error => console.log('error', error));
    
  }

  
  let firmasocial="";
  let firmalogo="";
  let firmabanner="";
  let firmanota="";
  let firmaeco="";
  let firmadir="";
  
  function dataFirma(datos) {
    firmasocial="";
    firmalogo="";
    firmabanner="";
    firmaeco="";
    firmanota="";
    firmadir="";

    let Orden = 0.0;
    let Imagen = "";      
    let Enlace = "";
    let Seccion = "";
    let Nota = "";
    let textoAlt="";
   
      for(let i = 0; i < datos.datos.length; i++)
      {
        
        Orden = datos.datos[i].Orden;
        Imagen = datos.datos[i].Imagen;      
        Enlace = datos.datos[i].Enlace;
        Seccion = datos.datos[i].Seccion?.Value || "";
        Nota = datos.datos[i].Nota;
        textoAlt = datos.datos[i].Texto_alternativo;
        
        if(Seccion.toUpperCase() === "LOGO")
        {
          firmalogo +='<a href="'+ Enlace +'"> <img src="'+ Imagen +'" alt="'+ textoAlt +'" width="258" height="87"></a>';
        
        }

        if(Seccion.toUpperCase() === "DIRECCION")
        {
          firmadir +='<strong>'+textoAlt+'</strong><br>'+ Nota +'<br><br>';
        }

        if(Seccion.toUpperCase() === "BANNER")
        {
          firmabanner +='<a href="'+ Enlace +'" ><img src="'+ Imagen +'" style="width:auto; height:auto;" alt="'+ textoAlt +'"> </a>';

        }

       if(Seccion.toUpperCase() === "SOCIAL")
        {
          firmasocial +='<a class="social-icons" href="'+ Enlace +'" target="_blank"><img src="'+ Imagen +'" style="margin:2px;" alt="'+ textoAlt +'" width="24" height="25"></a>';
        }

        
        if(Seccion.toUpperCase() === "ECO")
        {
          firmaeco +='<img src="'+ Imagen +'" width="14" height="14">'
        }

        if(Seccion.toUpperCase() === "NOTA")
        {
          firmanota +='<font color="#7F7F7F" size="1" face="Arial">'+ Nota +'</font><br><br>';
        }



      }
   }



function get_template_A_str(user_info)
{
  let str = ""; 

        str +='<table border="0" cellpadding="1" cellspacing="1"><tbody><tr><td valign="top"><font size="3" color="#17365d" face="Arial">';
        str +='<strong>'+ user_info.name +'</strong></font>';
        //str +='<strong>'+ user_info.name + (is_valid_data(user_info.GrdoAcad) ? ", " + user_info.GrdoAcad : "") +'</strong></font>';
        str +='<br><font size="2" face="Arial">'+ user_info.job +'</font><br>';
        str +='<font size="3" color="#17365d" face="Arial">';
        str += is_valid_data(user_info.pronoun) ? "<strong>" + user_info.pronoun : "";
        str += '</strong></font><br><font size="2" face="Arial">Tel.:';
        str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
        str += user_info.email;
        str += '<br>';
        /*
        str += is_valid_data(user_info.InfoAd1) ? (user_info.InfoAd1.startsWith('http') ? '<a href="' + user_info.InfoAd1 + '">' + user_info.InfoAd1 + '</a><br>' : '<span>' + user_info.InfoAd1 + '</span><br>') : "";
        str += is_valid_data(user_info.InfoAd2) ? (user_info.InfoAd2.startsWith('http') ? '<a href="' + user_info.InfoAd2 + '">' + user_info.InfoAd2 + '</a><br>' : '<span>' + user_info.InfoAd2 + '</span><br>') : "";
        str += is_valid_data(user_info.InfoAd3) ? (user_info.InfoAd3.startsWith('http') ? '<a href="' + user_info.InfoAd3 + '">' + user_info.InfoAd3 + '</a><br>' : '<span>' + user_info.InfoAd3 + '</span><br>') : "";
        
        for (let i = 1; i <= 15; i++)
        {
          let valor = user_info['InfoAd' + i];

          str += is_valid_data(valor)
            ? (valor.startsWith('http')
                ? '<a href="' + valor + '">' + valor + '</a><br>'
                : '<span>' + valor + '</span><br>')
            : "";
        }*/

        str +='</font></td></tr><tr><td><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="240" height="81">';
        str += firmalogo;

        str+= '</td><td width="15"></td>';
        str +='<td style="padding:0 0 0 15px;border-left-style:solid;border-left-width:1pt;border-left-color:#7f7f7f">';
        str +='<p><font size="2" face="Arial">';
        str += firmadir;
        str +='</font></p></td></tr></tbody></table></td></tr><tr><td height="70" align="left" valign="middle">';
        str +='<table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td style="width:auto; height:auto;">';
        
        str += firmabanner;

        str +='</td><td width="15"></td><td class="social" style="display: flex; align-items: center;justify-content: space-around;" width="150" height="70">';

        str += firmasocial;

        str+='</td></tr></tbody></table></td></tr><tr><td>';
        str += firmaeco + '&nbsp;&nbsp;';
        str +='<font color="#7F7F7F" size="1" face="Arial">No me imprimas si no es necesario.</font></td></tr><tr><td>';

        str +='<p style="margin:0">'+ firmanota;
        
        str +='</p></td></tr></tbody></table>';

   //console.log("signature_template_A: " + str)
  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table style='display:none;'>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://www.pucmm.edu.do/PublishingImages/firma-addin/marca-pucmm.jpg' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;
  
  return str;
}
