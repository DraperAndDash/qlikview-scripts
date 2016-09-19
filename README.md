
<html>

<head>
  <meta charset="utf-8"/>
</head>

<body>

<h1>Draper & Dash - Qlikview Scripts</h1>

<h2>About</h2>
<p>
	This website explains the functions and scripts that are common across many D&D applications. Including these functions and scripts in your development will help speed up the initial process and also make it easier for maintaining applications throughout their use. All these functions and scripts can be included by using the following line of code:
</p>
<span><script>
$(include=https://raw.githubusercontent.com/GlutesDD/qlikview-scripts/master/DraperDashQlikviewScripts.inc?token=ANDyRRSDETE0ume3mrB9hb1_nK60gbGxks5X6TwtwA%3D%3D);
</script></span>

<div><h2>GenerateCalendar.qvs</h2></div>
<div>CALL GenerateCalendar.qvs(Region, Date, Table)</div>
<div><h3>Type: </h3><span> Function</span></div>
<div><h3>Parameters: </h3>
<div><h4>Region</h4><span> - The first parameter that must be supplied is the geographical region for the calendar. This can be UK, US or OZ for United Kingdom, USA and Australia respectively. The region will define how the financial calendar fields are being calculated.</span></div>
<div><h4>Date</h4><span> - The second parameter is the data field you wish to use to associate the calendar to. This must be exactly the same as it appears in the data.</span></div>
<div><h4>Table</h4><span> - The third and final parameter is the table that your date field is located in. This must match the table name exactly and must be available as a resident.</span></div></div>
<div><h3>Description:</h3><p>This script allows a calendar to be generated from a single function call and is included via the DraperDashQlikviewScripts.inc include file. You must create an inline table called Calendars with a single field Calendar. Type the name of the date fields used to generate each calendar as the field values.</p></div>

<div><h2>GenerateCommonCalendar.qvs</h2></div>
<div>CALL GenerateCommonCalendar.qvs(Region, PrimaryKey, Table)</div>
<div><h3>Type: </h3><span> Function</span></div>
<div><h3>Parameters: </h3>
<div><h4>Region</h4><span> - The first parameter that must be supplied is the geographical region for the calendar. This can be UK, US or OZ for United Kingdom, USA and Australia respectively. The region will define how the financial calendar fields are being calculated.</span></div>
<div><h4>PrimaryKey</h4><span> - The second parameter is the primary key field for your fact table, you must have a primary key in the table to be able to create a link between the common calendar and the fact table. This must be exactly the same as it appears in the data.</span></div>
<div><h4>Table</h4><span> - The third and final parameter is the name of the fact table. This must match the table name exactly and must be available as a resident.</span></div></div>
<div><h3>Description:</h3><p>This script allows a common calendar to be generated which allows for analysing data in the same chart or object with a single data dimension where multiple calendars are needed. This is done from a single function call and is included via the DraperDashQlikviewScripts.inc include file. Please run this function after you have generated all other calendars.</p></div>


</body>

</body>

</html>
