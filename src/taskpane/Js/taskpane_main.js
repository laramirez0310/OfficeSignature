// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

//import {get_template_A_info} from '../../runtime/Js/autorunshared.js';

let _user_info;

function save_user_settings_to_roaming_settings()
{
  Office.context.roamingSettings.saveAsync(function (asyncResult)
  {
	console.log("save_user_info_str_to_roaming_settings - " + JSON.stringify(asyncResult));
  });
}

function disable_client_signatures_if_necessary()
{
  if ($("#checkbox_sig").prop("checked") === true)
  {
	Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult)
	{
	  //console.log("disable_client_signature_if_necessary - " + JSON.stringify(asyncResult));
	});
  }
}

function save_signature_settings(user_info)
{
  //let user_info_str = localStorage.getItem('user_info');
  let user_info_str = JSON.stringify(user_info);

	if (!user_info)
	{
		user_info_str = localStorage.getItem('user_info');
		if (user_info_str)
		{
		user_info = JSON.parse(user_info_str);
		}
	}

  if (user_info)
  {

	_user_info = user_info;
	Office.context.roamingSettings.set('user_info', user_info_str);
	Office.context.roamingSettings.set('newMail', $("#new_mail option:selected").val());
	Office.context.roamingSettings.set('reply', $("#reply option:selected").val());
	Office.context.roamingSettings.set('forward', $("#forward option:selected").val());

	Office.context.roamingSettings.set('override_olk_signature', $("#checkbox_sig").prop('checked'));

	save_user_settings_to_roaming_settings();

	disable_client_signatures_if_necessary();
	$("#message-successful").show("slow");
  }
  else
  {
	// TBD display an error somewhere?
  }
}

//export {save_signature_settings}; 

function set_body(str)
{
  Office.context.mailbox.item.body.setAsync
  (
	get_cal_offset() + str,

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  //console.log("set_body - " + JSON.stringify(asyncResult));
	}
  );
}

function set_signature(str)
{
  Office.context.mailbox.item.body.setSignatureAsync
  (
	str,

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  //console.log("set_signature - " + JSON.stringify(asyncResult));
	}
  );
}

function insert_signature(str)
{
  if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
  {
	set_body(str);
  }
  else
  {
	set_signature(str);
  }
}


function navigate_to_taskpane2()
{
  window.location.href = 'editsignature.html';
}