<!-- Edited by: Jim Saiya 2017-10-25 -->
<!DOCTYPE html>
<!-- xlsx.js (C) 2013-present  SheetJS sheetjs.com -->
<!-- vim: set ts=2: -->
<html xmlns:mso="urn:schemas-microsoft-com:office:office" xmlns:msdt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ProQuest IMC Marketing List Import Tool</title>
<link href="https://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet" />
<link href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.0.2/css/bootstrap.min.css" rel="stylesheet"/>

<script src="https://code.jquery.com/jquery-1.12.4.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.0.2/js/bootstrap.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-select/1.3.5/bootstrap-select.js"></script>
<script src="js-xlsx/shim.js"></script>
<script src="js-xlsx/jszip.js"></script>
<script src="js-xlsx/xlsx.js"></script>
<style>
body {
	font-family: "Open Sans";
}
h1 {
	font-family: "Open Sans";
	font-size: 31px;
	font-weight: 300;
}
h2 {
	font-family: "Open Sans";
	font-size: 26px;
	font-weight: 300;
}
#drop {
	border: 2px dashed #BBB;
	-moz-border-radius: 5px;
	-webkit-border-radius: 5px;
	border-radius: 5px;
	padding: 25px;
	text-align: center;
	font: 20pt bold,"Vollkorn";
	color: #BBB;
}
.hidden {
	display: none;
}
.header-area {
	background: #333333;
	min-height: 50px;
}
.navbar {
	position: relative;
	/*min-height: 50px;*/
	margin-bottom: 2px;
	border: 1px solid transparent;
}
.navbar-default {
	background-color: transparent;
	border-color: transparent;
}
.proquest-part {
	background: #019FDE;
}
.sectionBlock {
	padding-top: 20px;
	padding-bottom: 10px;
}
.title-part {
	text-align: center;
}
#tsFrame iframe {
	display: block;
	width: 895px;
	height: 435px;
	margin: 0 auto;
}
</style>

<!--[if gte mso 9]><SharePoint:CTFieldRefs runat=server Prefix="mso:" FieldList="FileLeafRef"><xml>
<mso:CustomDocumentProperties>
<mso:display_urn_x003a_schemas-microsoft-com_x003a_office_x003a_office_x0023_SharedWithUsers msdt:dt="string">Stoffregen, Christopher</mso:display_urn_x003a_schemas-microsoft-com_x003a_office_x003a_office_x0023_SharedWithUsers>
<mso:SharedWithUsers msdt:dt="string">935;#Stoffregen, Christopher</mso:SharedWithUsers>
</mso:CustomDocumentProperties>
</xml></SharePoint:CTFieldRefs><![endif]-->
</head>

<body>
	<div class="header-area">
		<div class="row">
			<div class="col-md-12">
			</div>
		</div>
	</div>

	<div class="proquest-part">
		<div class="container">
			<div class="col-md-12">
				<!--<div class="inner-proquest"><img src="http://contentz.mkt5049.com/lp/43888/477000/logo_0.jpg" class="img-responsive" alt="ProQuest" spname="logo_0.jpg" /></div>-->
				<div class="inner-proquest" style="padding: 10px;"><img src="https://myproquest.sharepoint.com/about/Documents/ProQuest%20-%20white.png" width="134" height="50" class="img-responsive" alt="ProQuest" spname="logo_0.jpg" /></div>
			</div>
		</div>
	</div>

	<div class="title-part">
		<div class="container sectionBlock">
			<h1><strong>Marketing List Import/Upload Tool</strong></h1>
		</div>
	</div>

	<div id="etZone" class="row">
		<div class="container sectionBlock" style="text-align: center;">
			Upload Type: &nbsp;
			<select id="selectUploadType" name="selectUploadType" class="btn dropdown-toggle selectpicker btn-default" onchange="etSelect();">
				<option value="select" selected="selected" >--- Select Upload Type ---</option>
				<option value="imcOnly" >Contacts only added to IMC</option>
				<option value="imcSF" >Contacts added to IMC and Salesforce</option>
				<option value="imcSFCmpn" >Contacts added to IMC and Salesforce with Campaign</option>
			</select>
		</div>
	</div>

	<div id="dropZone" class="row">
		<div class="container sectionBlock">
			<div id="drop">Drop an XLSX / XLS file here or select below.</div>
			<br>
			<div id="browseBtnRow">
				<p><input class="filepicked" type="file" name="xlfile" id="xlf" /></p>
			</div>
		</div>
	</div>

	<div id="notifyZone" class="container sectionBlock">
		<div class="col-sm-11 col-md-11 col-lg-11">
			Upload File: <b><span id="dspFilename"></span></b>
		</div>
		<div class="col-sm-1 col-md-1 col-lg-1">
			<button id="btnReload" class="btn btn-success pull-right">Reset</button>
		</div>
		<div class="col-sm-11 col-md-11 col-lg-11">
			Upload Type: <b><span id="dspFiletype"></span></b>
		</div>
		<div class="col-sm-7 col-md-7 col-lg-7 text-left">
			<h2><div id="fnlImportStatus"></div></h2>
			<div>
				<div id="nextSteps" class="text-left"></div>
			</div>
		</div>
	</div>

	<div id="statusZone" class="container sectionBlock">
		<hr>
		<div class="container sectionBlock">
			<div class="col-sm-12 col-md-12 col-lg-12 text-left">
				<h2>Summary:</h2>
			</div>
			<div class="col-sm-5 col-md-5 col-lg-5 text-right">
				Total Spreadsheet Rows for Processing:<br>
				Total Spreadsheet Rows with Required Data:<br>
				Total Spreadsheet Rows with Missing Required Data:<br>
			</div>
			<div class="col-sm-1 col-md-1 col-lg-1">
				<div id="tot_rows"></div>
				<div id="tot_good_rows" class="text-success"></div>
				<div id="tot_bad_rows" class="text-danger"></div>
			</div>
			<div class="col-sm-6 col-md-6 col-lg-6">
				<div id="importSummary" class="text-left"></div>
			</div>
		</div>
		<div class="container sectionBlock">
			<div class="col-sm-4 col-md-4 col-lg-4">
				<a id="myUploadBtn" class="btn btn-primary btn-md" onclick="upload_info();">Start My Upload</a>
			</div>
		</div>
	</div>

	<div id="preZone" class="row">
		<hr>
		<div class="container sectionBlock">
			<pre id="out"></pre>
		</div>
	</div>

	<iframe class="hidden" name="@rcvrForm" height='100%' width='100%' scrolling='no' src='https://www.pages04.net/proquest/GEMS-RCVR/rcvr-noalerts.html' allowfullscreen></iframe>

<!-- ################################################ HTML ABOVE ################################################# -->

<input type="checkbox" class="hidden" name="useworker" checked><br>
<input type="checkbox" class="hidden" name="xferable" checked><br>
<input type="checkbox" class="hidden" name="userabs" checked><br>

<!-- uncomment the next line here and in xlsxworker.js for encoding support -->
<!--<script src="dist/cpexcel.js"></script>-->

<!-- uncomment the next line here and in xlsxworker.js for ODS support -->
<script src="ods.js"></script>

<script>
var arrBlob = [];
var blah = [];
var arrValidate = [];
var totRows = 0;
var fileTypes = ['xlsx', 'xls'];  // acceptable file types

$( document ).ready(function() {
	setProcess();
	$('#selectUploadType').val('select');
	$('#btnReload').click(function() {
		location.reload();
	});
});

// XLSX is the main Object exposed by SheetJS
var X = XLSX;
var XW = {
	/* worker message */
	msg: 'xlsx',
	/* worker scripts */
	rABS: './js-xlsx/xlsxworker2.js',
	norABS: './js-xlsx/xlsxworker1.js',
	noxfer: './js-xlsx/xlsxworker.js'
};

var rABS = typeof FileReader !== 'undefined' && typeof FileReader.prototype !== 'undefined' && typeof FileReader.prototype.readAsBinaryString !== 'undefined';
if (!rABS) {
	// 'rABS' = 'Read As Binary String'
	document.getElementsByName('userabs')[0].disabled = true;
	document.getElementsByName('userabs')[0].checked = false;
}

var use_worker = typeof Worker !== 'undefined';
//use_worker = false;  //////////////////////////////// uncomment if running on a local system
if (!use_worker) {
	document.getElementsByName('useworker')[0].disabled = true;
	document.getElementsByName('useworker')[0].checked = false;
}

var transferable = use_worker;
if (!transferable) {
	document.getElementsByName('xferable')[0].disabled = true;
	document.getElementsByName('xferable')[0].checked = false;
}

// Debugging mode used by SheetJS
var wtf_mode = false;

// User has selected an upload type from the dropdown menu
function etSelect() {
	if (jQuery.inArray('first_name', arrValidate) !== -1) {
		setProcess();
	}
	// set which fields are mandatory
	if ($('#selectUploadType').val() === 'imcOnly') {
		arrValidate = ['email'];
		$('#dropZone').fadeIn('slow');
	} else if ($('#selectUploadType').val() === 'imcSF') {
		arrValidate = ['first_name','last_name','company','email','market','submarket','state','country','lead_source','lead_status','last_touch_lead_source','last_touch_lead_source_date'];
		$('#dropZone').fadeIn('slow');
	} else if ($('#selectUploadType').val() === 'imcSFCmpn') {
		arrValidate = ['first_name','last_name','company','email','market','submarket','state','country','lead_source','lead_status','campaign_id','campaign_member_status','last_touch_lead_source','last_touch_lead_source_date'];
		$('#dropZone').fadeIn('slow');
	} else {
		$('#dropZone').hide().fadeOut('slow');
	}
}

// User clicked the "Start My Upload" button
function upload_info() {
	// Set rcvrFormName to						| to generate these alerts
	// -----------------------------------------|-------------------------------
	var RCVR_NOALERTS = "rcvr-noalerts.html";	// No alerts generated
	var RCVR_ALERTS   = "rcvr-alerts.html";		// Send an email and issue a task
	var RCVR_EMAIL    = "rcvr-email.html";		// Send an email only
//	var z_url_base = "https://www.pages04.net/proquest-sandbox/GEMS-RCVR/";		// SANDBOX CONNECTION
	var z_url_base = "https://www.pages04.net/proquest/GEMS-RCVR/";				// PRODUCTION CONNECTION
	z_url = z_url_base + RCVR_NOALERTS;  // default form file to receive data on the IMC server
	z_target = "@rcvrForm";
	currentProcessed = 0;
	var delay = 0;
	$('#preZone').show();
	// process each row
	$.each(blah.Template, function(i, item) {
		setTimeout(function() {
			arrBlob = blah.Template[i];
			rcvrFormPage = RCVR_NOALERTS;
			leadSourceParam = "";
			// build the #formPostToIframe form element to hold the data to upload (1 spreadsheet row)
			$('body').append('<form action="'+z_url+'" method="post" target="'+z_target+'" id="formPostToIframe"></form>');
			// loop through fields
			for (var key in arrBlob) {
				name = key;
				value = arrBlob[key].toString();
console.log('Row '+(i+1)+': '+name+' = '+value);
				if (name === 'campaign_id' && value !== '') {
					$('#formPostToIframe').append('<input type="text" name="sp_ctc" value="'+value+'" />');
				} else if (name === 'campaign_member_status' && value !== '') {
					$('#formPostToIframe').append('<input type="text" name="sp_cts" value="'+value+'" />');
				} else if (name === 'comments' && value !== '') {
					rcvrFormPage = RCVR_ALERTS;
					$('#formPostToIframe').append('<input type="text" name="'+name+'" value="'+value+'" />');
				} else if (name === 'notes' && value !== '') {
					if (rcvrFormPage !== RCVR_ALERTS)
						rcvrFormPage = RCVR_EMAIL;  // only set this value if 'comments' field is empty
					$('#formPostToIframe').append('<input type="text" name="'+name+'" value="'+value+'" />');
				} else if (name === 'lead_source' && value !== '') {
					leadSourceParam = "?sp_source=" + value;
				} else {
					$('#formPostToIframe').append('<input type="text" name="'+name+'" value="'+value+'" />');
				}
			}
			// build the URL to call when the form is submitted
			z_url_source = z_url_base + rcvrFormPage + leadSourceParam;

//console.log('z_url_base: ' + z_url_base);
console.log('rcvrFormPage: ' + rcvrFormPage);
console.log('leadSourceParam: ' + leadSourceParam);
console.log('z_url_source: ' + z_url_source);

			// trigger form submission and destroy form
			$('#formPostToIframe').attr('action', z_url_source);
			$('#formPostToIframe').submit().remove();

			// running counter
			currentProcessed++;
			$('#out').append('<div>Spreadsheet row number <strong>'+currentProcessed+'</strong> has been processed.</div>');
			if (currentProcessed == totRows) {
				$('#out').append('<div>Finished...!!!</div>');
				$('#myUploadBtn').hide();
			}
		}, delay += 500);
	}); // end $.each(blah.Template)
} // end upload_info()

function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w, l*w + w)));
	o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}

// Convert Array Buffer to String
function ab2str(data) {
	var o = "", l = 0, w = 10240;
	for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint16Array(data.slice(l*w, l*w + w)));
	o += String.fromCharCode.apply(null, new Uint16Array(data.slice(l*w)));
	return o;
}

// Convert String to Array Buffer
function s2ab(s) {
	var b = new ArrayBuffer(s.length * 2), v = new Uint16Array(b);
	for (var i = 0; i != s.length; ++i) v[i] = s.charCodeAt(i);
	return [v, b];
}

function xw_noxfer(data, cb) {
	var worker = new Worker(XW.noxfer);
	worker.onmessage = function(e) {
		switch (e.data.t) {
			case 'ready':
				break;
			case 'e':
				console.error(e.data.d);
				break;
			case XW.msg:
				cb(JSON.parse(e.data.d));
				break;
		}
	};
	var arr = rABS ? data : btoa(fixdata(data));
	worker.postMessage({d:arr, b:rABS});
}

function xw_xfer(data, cb) {
	var worker = new Worker(rABS ? XW.rABS : XW.norABS);
	worker.onmessage = function(e) {
		switch (e.data.t) {
			case 'ready':
				break;
			case 'e':
				console.error(e.data.d);
				break;
			default:
				xx = ab2str(e.data).replace(/\n/g,'\\n').replace(/\r/g,'\\r');
				console.log('done');
				cb(JSON.parse(xx));
				break;
		}
	};
	if (rABS) {
		var val = s2ab(data);
		worker.postMessage(val[1], [val[1]]);
	} else {
		worker.postMessage(data, [data]);
	}
}

function xw(data, cb) {
	transferable = document.getElementsByName('xferable')[0].checked;
	if (transferable) xw_xfer(data, cb);
	else xw_noxfer(data, cb);
}

function to_json(workbook) {
	var result = {};
	var sheetName = workbook.SheetNames[0];
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName]);
		if (roa.length > 0) {
			result[sheetName] = roa;
		}
	});
	return result;
}

function process_wb(wb) {
	$('#statusZone').fadeIn('slow');
	$('#importSummary').empty();

	var output = "";
	output = JSON.stringify(to_json(wb), 2, 2);
	blah = to_json(wb);
	names = Object.getOwnPropertyNames(blah);
	totRows = 0;
	var totGoodRows = 0;
	var totBadRows = 0;

	if (names[0] === 'Template') {
		$.each(blah.Template, function(i, item) {
			var catchit = false;
			totRows++;
			thisTemplateItem = this;
			$.each(arrValidate, function (j, val) {
				if (thisTemplateItem[val] === undefined || thisTemplateItem[val] === '') {
					badXcelRow = totRows;
					$('#importSummary').append('<div>Spreadsheet Row <strong>'+badXcelRow+'</strong> is missing a value for <strong>'+val+'</strong>.</div>');
					catchit = true;
				}
			});
			if (catchit) {
				$('#importSummary').append('<hr>');
				totBadRows++;
			}
		});
		//	##############################   Write totals to File Status Area   ##################################
		$('#tot_rows').html('<b>'+totRows+'</b>');
		$('#tot_good_rows').html('<b>'+(totRows - totBadRows)+'</b>');
		$('#tot_bad_rows').html('<b>'+totBadRows+'</b>');
		if (totBadRows > 0) {
			$('#fnlImportStatus').hide().html('File Import FAILED required fields check!').fadeIn('slow');
			$('#fnlImportStatus').addClass('text-danger');
			$('#nextSteps').hide().html('(Please check the fields below on the import file and retry file upload.)').fadeIn('slow');
			$('#myUploadBtn').prop('disabled', true);
			$('#btnReload').hide().fadeIn('slow');
			$('#dropZone').hide().fadeOut('slow');
			$('#notifyZone').hide().fadeIn('slow');
			$('#etZone').hide().fadeOut('slow');
		} else {
			$('#fnlImportStatus').hide().html('File Import PASSED required fields check!').fadeIn('slow');
			$('#fnlImportStatus').addClass('text-success');
			$('#myUploadBtn').show();
			$('#notifyZone').hide().fadeIn('slow');
			$('#dropZone').hide().fadeOut('slow');
			$('#etZone').hide().fadeOut('slow');
		}
	} else { // first tab is not named 'Template'
		$('#dropZone').hide().fadeOut('slow');
		$('#notifyZone').hide().fadeIn('slow');
		$('#btnReload').hide().fadeIn('slow');
		$('#dspFilename').hide().html('----------').fadeIn('slow');
		$('#dspFiletype').hide().html('----------').fadeIn('slow');
		$('#statusZone').hide().fadeOut('slow');
		$('#fnlImportStatus').hide().html('<h2>Spreadsheet file does not contain TEMPLATE tab.</h2>').fadeIn('slow');
	}
}

// Initialize variables and prepare the display
function setProcess() {
	$('.filepicked').val('');
	$('#importSummary').empty();
	$('#myUploadBtn').hide().fadeOut('slow');
	$('#btnReload').hide().fadeOut('slow');
	$('#notifyZone').hide().fadeOut('slow');
	$('#dropZone').hide().fadeOut('slow');
	$('#statusZone').hide().fadeOut('slow');
	$('#preZone').hide().fadeOut('slow');
	$('#tsFrame').hide();
	arrBlob = [];
	blah = [];
	arrValidate = [];
}

//	##############################   EVENT HANDLERS   ##################################

var drop = document.getElementById('drop');
function handleDrop(e) {
	e.stopPropagation();
	e.preventDefault();
	rABS = document.getElementsByName('userabs')[0].checked;
	use_worker = document.getElementsByName('useworker')[0].checked;
	var files = e.dataTransfer.files;
	var f = files[0];
	{
		var extension = files[0].name.split('.').pop().toLowerCase(),  // file extension from input file
			isSuccess = fileTypes.indexOf(extension) > -1;  // is extension in acceptable types
		if (isSuccess) { // yes
			var reader = new FileReader();
			var name = f.name;
			reader.onload = function(e) {
				if (typeof console !== 'undefined') console.log('onload', new Date(), rABS, use_worker);
				var data = e.target.result;
				if (use_worker) {
					xw(data, process_wb);
				} else {
					var wb;
					if (rABS) {
						wb = X.read(data, {type: 'binary'});
					} else {
						var arr = fixdata(data);
						wb = X.read(btoa(arr), {type: 'base64'});
					}
					process_wb(wb);
				}
			}; // end reader.onload
			$('#notifyZone').hide().fadeOut('slow');
			$('#browseBtnRow').fadeOut('slow');
			$('#btnReload').fadeIn('slow');
			$('#dspFilename').hide().html(name).fadeIn('slow');
			$('#dspFiletype').hide().html($('#selectUploadType option:selected').text()).fadeIn('slow');
			if (rABS) reader.readAsBinaryString(f);
			else reader.readAsArrayBuffer(f);
		} else { // not an acceptable file type
			$('.filepicked').val('');
			$('#dropZone').hide().fadeOut('slow');
			$('#notifyZone').hide().fadeIn('slow');
			$('#btnReload').hide().fadeIn('slow');
			$('#dspFilename').hide().html(f.name).fadeIn('slow');
			$('#dspFiletype').hide().html('----------').fadeIn('slow');
			$('#fnlImportStatus').hide().html('<h2>IMPROPER FILE TYPE</h2>').fadeIn('slow');
			$('#fnlImportStatus').addClass('text-danger');
			$('#nextSteps').hide().html('(Only XLSX or XLS files may be used to upload.)').fadeIn('slow');
		}
	}
}

function handleDragover(e) {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = "copy";
}

if (drop.addEventListener) {
	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
}

var xlf = document.getElementById('xlf');
function handleFile(e) {
	rABS = document.getElementsByName('userabs')[0].checked;
	use_worker = document.getElementsByName('useworker')[0].checked;
	var files = e.target.files;
	var f = files[0];
	{
		if (files && files[0]) {
			var extension = files[0].name.split('.').pop().toLowerCase(),  // file extension from input file
				isSuccess = fileTypes.indexOf(extension) > -1;  // is extension in acceptable types?
			if (isSuccess) { // yes
				var reader = new FileReader();
				var name = f.name;
				reader.onload = function(e) {
					if (typeof console !== 'undefined') console.log('onload', new Date(), rABS, use_worker);
					var data = e.target.result;
					if (use_worker) {
						xw(data, process_wb);
					} else {
						var wb;
						if (rABS) {
							wb = X.read(data, {type: 'binary'});
						} else {
							var arr = fixdata(data);
							wb = X.read(btoa(arr), {type: 'base64'});
						}
						process_wb(wb);
					}
				}; // end reader.onload
				$('#btnReload').fadeIn('slow');
				$('#dspFilename').hide().html(name).fadeIn('slow');
				$('#dspFiletype').hide().html($('#selectUploadType option:selected').text()).fadeIn('slow');
				if (rABS) reader.readAsBinaryString(f);
				else reader.readAsArrayBuffer(f);
			} else { // not an acceptable file type
				$('.filepicked').val('');
				$('#dropZone').hide().fadeOut('slow');
				$('#notifyZone').hide().fadeIn('slow');
				$('#btnReload').hide().fadeIn('slow');
				$('#dspFilename').hide().html(f.name).fadeIn('slow');
				$('#dspFiletype').hide().html('----------').fadeIn('slow');
				$('#fnlImportStatus').html('<h2>IMPROPER FILE TYPE</h2>').fadeIn('slow');
				$('#fnlImportStatus').addClass('text-danger');
				$('#nextSteps').hide().html('(Only XLSX or XLS files may be used to upload.)').fadeIn('slow');
			}
		}
	}
}

if (xlf.addEventListener) {
	xlf.addEventListener('change', handleFile, false);
}
</script>

</body>
</html>
