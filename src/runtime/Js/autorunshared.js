// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on both Outlook on web and Outlook on Windows.

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */


function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows only. Insert signature into appointment or message.
 * Outlook on Windows can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
    addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Por favor confimar su firma..",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Establecer firma",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */

 function get_template_A_info(user_info) {
  const logoFileName = "marca-pucmm.jpg";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str +='<table border="0" cellpadding="1" cellspacing="1"><tbody><tr><td valign="top"><font size="3" color="#17365d" face="Arial">';
  str +='<strong>'+ user_info.name +'</strong></font>';
  str +='<br><font size="2" face="Arial">'+ user_info.job +'</font><br>';
  str +='<font size="3" color="#17365d" face="Arial">';
  str += is_valid_data(user_info.pronoun) ? "<strong>" + user_info.pronoun : "";
  str += '</strong></font><br><font size="2" face="Arial">Tel.:';
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += user_info.email;
  str += '<br/>';
  str += is_valid_data(user_info.InfoAd1) ? (user_info.InfoAd1.startsWith('http') ? '<a href="' + user_info.InfoAd1 + '">' + user_info.InfoAd1 + '</a><br/>' : '<p>' + user_info.InfoAd1 + '</p>') : "";
  str += is_valid_data(user_info.InfoAd2) ? (user_info.InfoAd2.startsWith('http') ? '<a href="' + user_info.InfoAd2 + '">' + user_info.InfoAd2 + '</a><br/>' : '<p>' + user_info.InfoAd2 + '</p>') : "";
  str += is_valid_data(user_info.InfoAd3) ? (user_info.InfoAd3.startsWith('http') ? '<a href="' + user_info.InfoAd3 + '">' + user_info.InfoAd3 + '</a><br/>' : '<p>' + user_info.InfoAd3 + '</p>') : "";
  str +='</font></td></tr><tr><td><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="240" height="81">';
  str +='<a href="https://pucmm.edu.do/"><img src="https://www.pucmm.edu.do/PublishingImages/firma-addin/marca-pucmm.jpg" width="258" height="87" alt="Pontificia Universidad Católica Madre y Maestra"></a></td><td width="15"></td>';
  str +='<td style="padding:0 0 0 15px;border-left-style:solid;border-left-width:1pt;border-left-color:#7f7f7f">';
  str +='<p><font size="2" face="Arial"><strong>Campus de Santiago:</strong><br>Autopista Duarte km. 1½, Santiago, R.D.';
  str +='<br><br><strong>Campus de Santo Domingo:</strong><br>Av. Abraham Lincoln esq. Av. Simón Bolívar, Santo Domingo, R.D.</font></p></td></tr></tbody></table></td></tr><tr><td height="70" align="left" valign="middle">';
  str +='<table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="600" height="70">';
  str +='<a href="https://pucmm.edu.do/" ><img src="https://pucmm.edu.do/PublishingImages/firma-addin/banner.png" width="700" height="160" alt="Banner PUCMM"> </a>';
  str +='</td><td width="15"></td><td class="social" style="display: flex; align-items: center;justify-content: space-around;" width="150" height="70">';
  str +='<a class="social-icons" href="http://www.facebook.com/pucmm/" target="_blank"><img style="margin:2px;" src="https://www.pucmm.edu.do/PublishingImages/firma-addin/facebook.png" alt="Facebook PUCMM" width="24" height="24"></a>';
  str +='<a class="social-icons" href="http://twitter.com/pucmm/" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-addin/twitter.png" style="margin:2px;" alt="Twitter PUCMM" width="24" height="25"></a>';
  str +='<a class="social-icons" href="http://www.instagram.com/pucmm" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-addin/instagram.png" style="margin:2px;" alt="Instagram PUCMM" width="24" height="25"></a>';
  str +='<a class="social-icons" href="http://www.youtube.com/pucmmtv/" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-addin/youtube.png" style="margin:2px;" alt="Youtube PUCMM" width="24" height="25"></a>'; 
  str +='<a class="social-icons" href="https://www.linkedin.com/edu/school?id=12020" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-addin/linkedin.png" style="margin:2px;" alt="Linkedin PUCMM" width="24" height="25"></a></td></tr></tbody></table></td></tr><tr><td>';
  str +='<img src="https://www.pucmm.edu.do/PublishingImages/firma-addin/Green.gif" width="14" height="14">&nbsp;&nbsp;';
  str +='<font color="#7F7F7F" size="1" face="Arial">No me imprimas si no es necesario.</font></td></tr><tr><td>';
  str +='<p style="margin:0"><font color="#7d7d7d" size="1" face="Arial">NOTA DE CONFIDENCIALIDAD: La información transmitida, incluidos los archivos adjuntos, está dirigida solo a la persona o entidad a la que ha sido remitida y puede contener información confidencial y/o privilegiada. Cualquier difusión u otro uso de la misma, o tomar cualquier acción basada en esta información por personas o entidades distintas al destinatario, está prohibido. Si ha recibido este mensaje por error, por favor contactar al remitente y destruya cualquier copia de esta información.<br>';
  str +='<br>CONFIDENTIALITY NOTE: The information transmitted, including attachments, is intended only for the person or entity to which it is addressed and may contain confidential and/or privileged material. Any review, retransmission, dissemination or other use of, or taking of any action in reliance upon this information by persons or entities other than the intended recipient is prohibited. If you received this in error, please contact the sender and destroy any copies of this information.</font></p></td></tr></tbody></table>';
  
  console.log("autorunshared.js "+ str);
  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

//export {get_template_A_info};

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://www.pucmm.edu.do/PublishingImages/firma-addin/logoFirma.jpg' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
