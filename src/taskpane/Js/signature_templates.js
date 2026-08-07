// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
 import{ firmasocial, firmalogo, firmabanner, firmanota, firmaeco, firmadir } from '../../runtime/Js/autorunshared.js';


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
