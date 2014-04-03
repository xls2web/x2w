function id(domId) {
    return document.getElementById(domId);
}

function log(message, div) {
	
	if (typeof(div) == "undefined") {div = "JsOutputDiv";}
	var m = new Date();
	var dateString =
	  m.getUTCFullYear() +"/"+
	  ("0" + (m.getUTCMonth()+1)).slice(-2) +"/"+
	  ("0" + m.getUTCDate()).slice(-2) + " " +
	  ("0" + m.getUTCHours()).slice(-2) + ":" +
	  ("0" + m.getUTCMinutes()).slice(-2) + ":" +
	  ("0" + m.getUTCSeconds()).slice(-2);	
    
    var child = document.createTextNode(message);
    var parent = document.getElementById(div);
    parent.appendChild(document.createTextNode("<------------------------>"));
    parent.appendChild(document.createElement("br"));
    parent.appendChild(document.createTextNode(dateString));
    parent.appendChild(document.createElement("br"));
    parent.appendChild(child);
    parent.appendChild(document.createElement("br"));
    parent.scrollTop = parent.scrollHeight;
}

function xls2JSONcall(io){

// I have to call it from iframe, otherwise jQuery stops working correctly

   var iFrame = document.getElementById(io.divId).firstChild;
   var excel = iFrame.contentWindow;
   excel.xls2JSON(io);	   
}



function xlsEmbed(fileToken, div, init, listener){
while( div.hasChildNodes() )
        {
        div.removeChild(div.lastChild);
        }  
        
   	   	
   	var iframe = document.createElement("iframe");
//   	iframe.setAttribute("src", "excel_uploaded.html?filetoken="+encodeURIComponent(fileToken)+"&div="+div.id+"&init="+init+"&listener="+listener);
   	iframe.setAttribute("width", "100%");
   	iframe.setAttribute("height", "100%");
   	iframe.setAttribute("frameborder", "0");
   	iframe.setAttribute("scrolling", "0");
   	iframe.setAttribute("onload", "fakeCall");
   	    
	div.appendChild(iframe);

   	populateXlsContent(iframe, {
   		"filetoken": fileToken,
   		"div": div.id,
   		"init": init,
   		"listener": listener
   	});



}

/**
 * Represents a xls2web plugin.
 *
 * @class xls2web
 * @constructor
 * @param [filetoken] unique token used to capture the embeded excel file, by default is taken from data-token attribute
 **/

$.fn.xls2web = function (options)
{
var filetoken, init, listener;  
if (typeof(options)	== "object") 
{
	filetoken = options.filetoken;
	init = options.init;
	listener = options.listener;
}
else
{
	filetoken = options;
}
var index;

	for (index = 0; index < this.length; ++index) {	
		if (filetoken == null)
		{
		    xlsEmbed(this.get(index).getAttribute('data-token'), this.get(index), init, listener);
	    }
	    else 
	    {
			xlsEmbed(filetoken, this.get(index), init, listener);
	    } 
	}
	return 0;
};

function launchQueryBatch(div, ioIn){

	var ioOut = {};
	ioOut.divId = div;




	if(typeof(ioIn) === "string")
			 	{
			 		ioOut.writeValue="";
			 		ioOut.writeTo="";
			 		ioOut.readFrom = ioIn;
		 		}


	if(typeof(ioIn) === "object")
		{
					ioOut.writeValue=ioIn.writeValue;
			 		ioOut.writeTo=ioIn.writeTo;
			 		if (typeof(ioIn.readFrom)=="undefined")
			 		{
			 			ioOut.readFrom = ioIn.writeTo;
			 		}
			 		else
			 		{
				 		ioOut.readFrom = ioIn.readFrom;	
			 		}			 		

		} 		
		

	ioOut.callback = callbackSelect(div, ioIn);

	if ((div == ioIn.divId) || (typeof(ioIn.divId) == "undefined"))
	{
		xls2JSONcall(ioOut);
	}
	else
	{
		alert('xls2JSONcall fail '+div+ ' '+ioIn.divId ); return 999;
	}
		 


}

function callbackSelect(div, ioIn){
    //console.log(id(div).getAttribute('data-callback'));
    var callback;
	if(typeof(ioIn.callback) == "undefined")
		 	 {
		 	 	if(id(div).getAttribute('data-callback') == null)
			 	 {
			 	 	callback = "xls2JSONcallback";
			 	 }
			 	else
			 	 {
			 	 	callback = id(div).getAttribute('data-callback');
			 	 }
		 	 	
		 	 }
	else
	{
		callback = ioIn.callback;
	}	 	 
    return callback;		 	 
}


$.fn.xls2JSON = function (io)
{
var iojq = io;
if (typeof(io) == "undefined")
{
	alert('No parameters passed to query!');
	return 1;
}	


	for (var index = 0; index < this.length; ++index) {	
		launchQueryBatch(this.get(index).id,iojq);
	}
	return 0;
	
};

//function noCallback(){console.log("no callback")}

function populateXlsContent(iframe, options){
	var filetoken = options.filetoken,
		div = options.div,
		init = options.init,
		listener = options.listener;

	var childDocument = iframe.contentWindow.document;
    var html = childDocument.getElementsByTagName("html")[0];
    var head = childDocument.getElementsByTagName("head")[0];
    var body = childDocument.getElementsByTagName("body")[0];
    
    //var meta = childDocument.createElement("meta");    
    var scriptMs = childDocument.createElement("script");
    var scriptGt = childDocument.createElement("script");
    var scriptJq = childDocument.createElement("script");
    var loadingDiv = childDocument.createElement("div");
    var qrwExcelDiv = childDocument.createElement("div");


    html.setAttribute("lang", "en");
    html.setAttribute("style", "width: 100%;height: 100%");
    //meta.setAttribute("charset", "utf-8");
    //html.appendChild(meta);
    scriptMs.setAttribute("src", "http://r.office.microsoft.com/r/rlidExcelWLJS?v=1&kip=1");
    scriptMs.setAttribute("id", "scriptMs");
    scriptGt.innerHTML = 'function GET(){ var GET = {"filetoken": "'+filetoken+'", "div": "'+div+'", "init": "'+init+'", "listener": "'+listener+'"}; return GET;}';
    scriptJq.setAttribute("src", "http://x2w.xls2web.com/_x2w.js");
    head.appendChild(scriptMs);
    head.appendChild(scriptGt);
    head.appendChild(scriptJq);
    
    body.setAttribute("style", "width: 100%;height: 100%");
    loadingDiv.setAttribute("id", "loadingdiv");
    loadingDiv.setAttribute("style", "font-size: 50%");
    qrwExcelDiv.setAttribute("id", "_qrwExcelDiv");
    qrwExcelDiv.setAttribute("style", "width: 90%; height: 90%");
    body.appendChild(loadingDiv);
    body.appendChild(qrwExcelDiv);

    launchXlsEmbed(div, 0, 1000, 10);

}



function launchXlsEmbed(div, currAttempt, totalAttempts, interval) {	
	var iFrame = document.getElementById(div).firstChild;
	var excel = iFrame.contentWindow;

	if (currAttempt < totalAttempts)
	{
		if ((typeof(excel.loadEwaOnPageLoad) == "function")
			&& (typeof(excel.Ewa) != "undefined"))
		{
			//console.log(typeof(excel.Ewa));
			excel.loadEwaOnPageLoad();			
		}
		else
		{
			//console.log(Date.now());
			setTimeout(function(){launchXlsEmbed(div, currAttempt+1, totalAttempts, interval);},interval);
		}
	}		
}


