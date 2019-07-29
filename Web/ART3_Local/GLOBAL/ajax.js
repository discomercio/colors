//<SCRIPT>
var AJAX_REQUEST_IS_NOT_INITIALIZED = 0;
var AJAX_REQUEST_HAS_BEEN_SETUP = 1;
var AJAX_REQUEST_HAS_BEEN_SENT = 2;
var AJAX_REQUEST_IS_IN_PROCESS = 3;
var AJAX_REQUEST_IS_COMPLETE = 4;

function GetXmlHttpObject() {
var xmlHttp=null;

	try
		{
	//  Firefox, Opera 8.0+, Safari
		xmlHttp=new XMLHttpRequest();
		}
	catch (e)
		{
	//  Internet Explorer
		try
			{
		//  Internet Explorer 6.0+
			xmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
			}
		catch (e)
			{
		//  Internet Explorer 5.5+
			xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
			}
		}

	return xmlHttp;
}
