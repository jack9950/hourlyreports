# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# This contains data for the salesemail.py breakdownemail.py scripts
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------


# ------------------------------------------------------------------------------
# This is data for the salesemail.py script
# ------------------------------------------------------------------------------


agentRowStart = """<tr style='height:12.75pt'>"""
agentRowEnd = """</tr>"""

agentIDStart = """<td width=71 valign=bottom style='width:53.0pt;border:solid gray 1.0pt;border-top:none;padding:0in 5.4pt 0in 5.4pt;height:12.75pt'><p class=MsoNormal><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>"""
agentIDEnd = """<o:p></o:p></span></p></td>"""

agentNameStart = """<td width=219 valign=bottom style='width:164.0pt;border-top:none;border-left:none;border-bottom:solid gray 1.0pt;border-right:solid gray 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:12.75pt'><p class=MsoNormal><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>"""
agentNameEnd = """<o:p></o:p></span></p></td>"""

signInTimeStart = """<td width=91 valign=bottom style='width:68.0pt;border-top:none;border-left:none;border-bottom:solid gray 1.0pt;border-right:solid gray 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:12.75pt'><p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>"""
signInTimeEnd = """<o:p></o:p></span></p></td>"""

callsHandledStart = """<td width=91 valign=bottom style='width:68.0pt;border-top:none;border-left:none;border-bottom:solid gray 1.0pt;border-right:solid gray 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:12.75pt'><p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:black'>"""
callsHandledEnd = """<o:p></o:p></span></p></td>"""

AHTStart = """<td width=84 valign=top style='width:63.0pt;border-top:none;border-left:none;border-bottom:solid gray 1.0pt;border-right:solid gray 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:12.75pt'><p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Arial",sans-serif;color:black'>"""
AHTEnd = """<o:p></o:p></span></p></td>"""


grandTotalRowStart = """<tr style='height:16.15pt'>"""
grandTotalRowEnd = """</tr>"""

grandTotalAgentID = """<td width=71 valign=bottom style='width:53.0pt;border-top:none;border-left:solid lightgrey 1.0pt;border-bottom:solid lightgrey 1.0pt;border-right:none;background:maroon;padding:0in 5.4pt 0in 5.4pt;height:16.15pt'><p class=MsoNormal><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>&nbsp;<o:p></o:p></span></b></p></td>"""

grandTotalAgentName = """<td width=219 valign=bottom style='width:164.0pt;border-top:none;border-left:none;border-bottom:solid lightgrey 1.0pt;border-right:solid lightgrey 1.0pt;background:maroon;padding:0in 5.4pt 0in 5.4pt;height:16.15pt'><p class=MsoNormal align=right style='text-align:right'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>NDIHO, JACKSON Total<o:p></o:p></span></b></p></td>"""

grandTotalSignInTime = """<td width=91 valign=bottom style='width:68.0pt;border-top:none;border-left:none;border-bottom:solid lightgrey 1.0pt;border-right:solid lightgrey 1.0pt;background:maroon;padding:0in 5.4pt 0in 5.4pt;height:16.15pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>&nbsp;<o:p></o:p></span></b></p></td>"""

grandTotalCallsHandledStart = """<td width=91 valign=bottom style='width:68.0pt;border-top:none;border-left:none;border-bottom:solid lightgrey 1.0pt;border-right:solid lightgrey 1.0pt;background:maroon;padding:0in 5.4pt 0in 5.4pt;height:16.15pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>"""
grandTotalCallsHandledEnd = """<o:p></o:p></span></b></p></td>"""

grandTotalAHTStart = """<td width=84 valign=bottom style='width:63.0pt;border-top:none;border-left:none;border-bottom:solid lightgrey 1.0pt;border-right:solid lightgrey 1.0pt;background:maroon;padding:0in 5.4pt 0in 5.4pt;height:16.15pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>"""
grandTotalAHTEnd = """<o:p></o:p></span></b></p></td>"""


topOfTable = """
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><style><!--
/* Font Definitions */
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:Verdana;
	panose-1:2 11 6 4 3 5 4 4 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:#0563C1;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:#954F72;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal;
	font-family:"Calibri",sans-serif;
	color:windowtext;}
span.EmailStyle18
	{mso-style-type:personal;
	font-family:"Calibri",sans-serif;
	color:#1F497D;}
span.EmailStyle19
	{mso-style-type:personal-reply;
	font-family:"Calibri",sans-serif;
	color:#1F497D;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-size:10.0pt;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
--></style><!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]--><!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]-->
</head>
<body lang=EN-US link="#0563C1" vlink="#954F72">
	<div class=WordSection1>
		<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0 style='width:416.0pt;margin-left:-.15pt;border-collapse:collapse'>
			<tr style='height:34.5pt'>

				<td width=71 valign=bottom style='width:53.0pt;border:solid lightgrey 1.0pt;background:#4E0000;padding:0in 5.4pt 0in 5.4pt;height:34.5pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>Agent ID<o:p></o:p></span></b></p></td>
				<td width=219 valign=bottom style='width:164.0pt;border:solid lightgrey 1.0pt;border-left:none;background:#4E0000;padding:0in 5.4pt 0in 5.4pt;height:34.5pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>Agent Name<o:p></o:p></span></b></p></td>

				<td width=91 valign=bottom style='width:68.0pt;border:solid lightgrey 1.0pt;border-left:none;background:#4E0000;padding:0in 5.4pt 0in 5.4pt;height:34.5pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>Sign in Time<o:p></o:p></span></b></p></td>

				<td width=91 valign=bottom style='width:68.0pt;border:solid lightgrey 1.0pt;border-left:none;background:#4E0000;padding:0in 5.4pt 0in 5.4pt;height:34.5pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>Calls Handled<o:p></o:p></span></b></p></td>

				<td width=84 valign=bottom style='width:63.0pt;border:solid lightgrey 1.0pt;border-left:none;background:#4E0000;padding:0in 5.4pt 0in 5.4pt;height:34.5pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;font-family:"Verdana",sans-serif;color:white'>AHT<o:p></o:p></span></b></p></td>
			</tr>
"""

# ------------------------------------------------------------------------------
# This is data for the breakdown email
# ------------------------------------------------------------------------------
topOfBreakdownTable = """
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta http-equiv=Content-Type content="text/html; charset=us-ascii">
    <meta name=Generator content="Microsoft Word 15 (filtered medium)">
    <style><!--
      /* Font Definitions */
      @font-face
      	{font-family:"Cambria Math";
      	panose-1:2 4 5 3 5 4 6 3 2 4;}
      @font-face
      	{font-family:Calibri;
      	panose-1:2 15 5 2 2 2 4 3 2 4;}
      /* Style Definitions */
      p.MsoNormal, li.MsoNormal, div.MsoNormal
      	{margin:0in;
      	margin-bottom:.0001pt;
      	font-size:11.0pt;
      	font-family:"Calibri",sans-serif;}
      a:link, span.MsoHyperlink
      	{mso-style-priority:99;
      	color:#0563C1;
      	text-decoration:underline;}
      a:visited, span.MsoHyperlinkFollowed
      	{mso-style-priority:99;
      	color:#954F72;
      	text-decoration:underline;}
      span.EmailStyle17
      	{mso-style-type:personal-compose;
      	font-family:"Calibri",sans-serif;
      	color:windowtext;}
      .MsoChpDefault
      	{mso-style-type:export-only;
      	font-family:"Calibri",sans-serif;}
      @page WordSection1
      	{size:8.5in 11.0in;
      	margin:1.0in 1.0in 1.0in 1.0in;}
      div.WordSection1
      	{page:WordSection1;}
  --></style><!--[if gte mso 9]><xml>
  <o:shapedefaults v:ext="edit" spidmax="1026" />
  </xml><![endif]--><!--[if gte mso 9]><xml>
  <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1" />
  </o:shapelayout></xml><![endif]-->
  </head>
  <body lang=EN-US link="#0563C1" vlink="#954F72">
    <div class=WordSection1>
      <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=1273 style='width:955.0pt;border-collapse:collapse'>
        <tr style='height:33.75pt'>
          <td width=489 colspan=4 valign=bottom style='width:367.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>Electricity Sales<o:p></o:p></span></b></p></td>
          <td width=19 valign=bottom style='width:14.0pt;padding:0in  5.4pt 0in 5.4pt;height:33.75pt'></td>
          <td width=765 colspan=5 valign=bottom style='width:574.0pt;padding:0in 5.4pt 0in  5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>DEPP<o:p></o:p></span></b></p></td>
        </tr>
        <tr style='height:56.25pt'>
		  <td width=167 valign=bottom style='width:125.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
		  <td width=75 valign=bottom style='width:56.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Account Number<o:p></o:p></span></b></p></td>
		  <td width=75 valign=bottom style='width:56.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Order Number<o:p></o:p></span></b></p></td>
		  <td width=173 valign=bottom style='width:130.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Order Status<o:p></o:p></span></b></p></td>
		  <td width=19 valign=bottom style='width:14.0pt;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'></td>
		  <td width=167 valign=bottom style='width:125.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
		  <td width=75 valign=bottom style='width:56.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Account Number<o:p></o:p></span></b></p></td>
		  <td width=75 valign=bottom style='width:56.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Order Number<o:p></o:p></span></b></p></td>
		  <td width=231 valign=bottom style='width:173.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>DEPP Name<o:p></o:p></span></b></p></td>
		  <td width=173 valign=bottom style='width:130.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Order Status<o:p></o:p></span></b></p></td>
	   </tr>
      </table>
      <p class=MsoNormal><o:p>&nbsp;</o:p></p>
      <p class=MsoNormal><o:p>&nbsp;</o:p></p>
      <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=851 style='width:638.0pt;border-collapse:collapse'>
        <tr style='height:33.75pt'>
          <td width=325 nowrap colspan=2 valign=bottom style='width:244.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>FCP Sales<o:p></o:p></span></b></p></td>
          <td width=80 nowrap valign=bottom style='width:60.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'></td>
          <td width=445 nowrap colspan=3 valign=bottom style='width:334.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>FCP Opportunities<o:p></o:p></span></b></p></td>
        </tr>
        <tr style='height:56.25pt'>
          <td width=167 valign=bottom style='width:125.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
          <td width=159 valign=bottom style='width:119.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>First Choice Power Account Number<o:p></o:p></span></b></p></td>
          <td width=80 nowrap valign=bottom style='width:60.0pt;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'></td>
          <td width=167 valign=bottom style='width:125.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
          <td width=128 valign=bottom style='width:96.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Account Number<o:p></o:p></span></b></p></td>
          <td width=151 valign=bottom style='width:113.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Pogo Status<o:p></o:p></span></b></p></td>
        </tr>
      </table>
    </div>
  </body>
</html>
"""
