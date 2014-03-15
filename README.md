x2w
===

Xls2web jQuery plug-in

x2w.js is a jQuery plugin that bridges your custom Excel spreadsheet and your site DOM elements. 
The following four steps make the basic usage scenario:

1. Add <script src="http://x2w.xls2web.com/x2w.js"><script>
2. Load your spreadsheet into container by adding $("#container").xls2web(options)
3. Call your spreadsheet through ("#container").xls2JSON(options)
4. Create a callback function that processes the response JSON (by default it is xls2JSONCallback(response))

See more details on JSFiddle:
http://jsfiddle.net/xls2web/ugBe2/
http://jsfiddle.net/xls2web/RVjLD/
http://jsfiddle.net/xls2web/KMtnc/
