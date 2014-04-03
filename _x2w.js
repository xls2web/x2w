  // run the Excel load handler on page load
  if (window.attachEvent) {
    window.attachEvent("onload", loadEwaOnPageLoad);
  } else {
    window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
  }


  



//function loadEwaOnPageLoad1() {alert('sobaka');}
  
  function loadEwaOnPageLoad() {

    document.getElementById("loadingdiv").innerHTML = "Loading...";

    var fileToken = GET().filetoken;

    var props = {
      uiOptions: {
        showGridlines: true,
        showRowColumnHeaders: true,
        showParametersTaskPane: true
        },
        interactivityOptions: {
        allowTypingAndFormulaEntry: true,
        allowParameterModification: true,
        allowSorting: true,
        allowFiltering: true,
        allowPivotTableInteractivity: true
        }   
      };
            
      Ewa.EwaControl.loadEwaAsync(fileToken, "_qrwExcelDiv" , props, function(){
        document.getElementById("loadingdiv").style.display = "none";
        Ewa.EwaControl.add_applicationReady(ewaApplicationReady); 

            });

  }

  
  function ewaApplicationReady() {        	
    	var init = GET().init,
          listener = GET().listener;
      if ((init != null) && (init != ""))
      {
        var callbackFnText = "parent."+init+"();";
        var initCallback = new Function(callbackFnText);
        //console.log(callbackFnText);
        initCallback();        
      }
      else
      {
       return 0; 
      }

      if ((listener != null) && (listener != ""))
      {
        Ewa.EwaControl.getInstances().getItem(0).getActiveWorkbook().add_sheetDataEntered(sheetDataEnteredHandler);
        //console.log("listener added"); 
      }
      
  }

  function sheetDataEnteredHandler(rangeChangeArgs) {
   
      var listener = GET().listener;
      if ((listener != null) && (listener != ""))
      {
        var callbackFnText = "parent."+listener+"();";
        var listenerCallback = new Function(callbackFnText);
        //console.log(callbackFnText);
        listenerCallback();        
      }
      else
      {
       return 0; 
      }

  }
  
  function ewaApplicationRefresh(io) 
  {
    var fileToken = GET().filetoken,
        div = GET().div,
        init = GET().init,
        listener = GET().listener;  
  	//alert(div+" Application is not ready or session timeout, please click ok to refresh then call again!");
    //alert("Location "+window.location)
    //alert("Token = "+fileToken+' Div = '+ div);
    var returnJSON = {
            inputPar: io,
            queryStatus: "fail"
          };
             

    var callbackFnText = "parent."+io.callback+"(response);";
    var usrCallback = new Function("response", callbackFnText);;
    usrCallback(returnJSON);

    parent.xlsEmbed(fileToken, parent.document.getElementById(div), init, listener);

  	

  }
  
  function xls2JSON(io) {
  	     
  //    alert("xls2json called"+ io.divId);
  //    if(io.divId == 'xliFrameContainer'){io.readFrom='Sheet1!A4'}
      var objEwa   = Ewa.EwaControl.getInstances().getItem(0); 
      
      if (objEwa.getActiveWorkbook() != "undefined") {
      	   objEwa.getActiveWorkbook().getRangeA1Async(io.writeTo, setCallback, io);
      	   //console.log("Workbook is here "+io.divId);
      } 
      else
      {
         ewaApplicationRefresh(io);	
      } 

                  	  	   	             
  }
  
  function setCallback(asyncResult)
  {
    //console.log('setCallBack called');
    //console.log(asyncResult);

    var io = asyncResult.getUserContext(); 
//    alert('setCallBack called ' + io.divId);   
    if (asyncResult.getCode() == 0)
    {
		    var range = asyncResult.getReturnValue();
        var arrValue   = new Array();
      	arrValue[0]    = new Array();  
    		arrValue[0][0] = io.writeValue;
    range.setValuesAsync(arrValue,setRangeValues, io);
    //console.log("setCallBack getcode = 0 "+io.divId);
    }
    else
    {
    setRangeValues(asyncResult);
    //console.log("Set skipped "+io.divId);
    }
  }
  
  function setRangeValues(asyncResult)
  {
          var io = asyncResult.getUserContext();
          var ewa = Ewa.EwaControl.getInstances().getItem(0);
          var workbook = ewa.getActiveWorkbook();
          
	if (ewa.getActiveWorkbook() != "undefined") {
            workbook.getRangeA1Async(io.readFrom, getCallBack, io);
            //console.log("setRangeValues was called "+io.divId);
        } 
        else
        {
            ewaApplicationRefresh(io);
        }   
 
  } 
  
  function getCallBack(asyncResult) {      
    var range = asyncResult.getReturnValue();
    var io = asyncResult.getUserContext();
    //console.log("getCallBack getCode = "+ asyncResult.getCode());
    
    if (asyncResult.getCode() != 0)
    {
          ewaApplicationRefresh(io);      
    }
          
    var io = asyncResult.getUserContext();
    range.getValuesAsync(0,getRangeValues,io);
    //console.log("getCallBack was called "+io.divId);
  } 
  
  function getRangeValues(asyncResult)
  {
      // Get the value from asyncResult if the asynchronous operation was successful.
      if (asyncResult.getCode() == 0)
      {                                    
          var returnJSON = {
        		inputPar: asyncResult.getUserContext(),
        		outputArray: asyncResult.getReturnValue(),
            queryStatus: "success"
        	};
         
  	    var io = asyncResult.getUserContext();
        var callbackFnText = "parent."+io.callback+"(response);";
        var usrCallback = new Function("response", callbackFnText);
        usrCallback(returnJSON);
        //console.log(usrCallback);
        //parent.xls2JSONcallback(returnJSON);

      }
      else 
      {
            alert("Operation failed with error message " + asyncResult.getDescription() + ".");
      }    
  }

  window
