x2w.js - xls2web jQuery plug-in
===============================

x2w.js is a jQuery plugin that bridges your custom Microsoft Excel spreadsheet saved on Onedrive (former Skydrive) and your website DOM elements. Inspired by www.excelmashup.com this small library is aimed at making your Excel mash-up development process  as easy as calling your spreadheet by token and further processing of a JSON response.

The following four steps make the basic usage scenario:

1. Add &lt;script src="http://x2w.xls2web.com/x2w.js" &#62; &lt;/script&gt;
2. Load your spreadsheet into container div by adding $("#container").xls2web(options)
3. Call your spreadsheet through ("#container").xls2JSON(options)
4. Create a callback function that processes the response JSON (by default it is xls2JSONCallback(response))

See more details on JSFiddle:
<br> http://jsfiddle.net/xls2web/ugBe2/ 
<br> http://jsfiddle.net/xls2web/RVjLD/
<br> http://jsfiddle.net/xls2web/KMtnc/
<br> http://x2w.co
<br> http://www.xls2web.com
<br> http://www.excelmashup.com
