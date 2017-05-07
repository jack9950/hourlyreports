import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jackson.ndiho@iqor.com'
mail.Subject = 'This is a Test'

emailBody = """
<head></head>
<body lang=EN-US link="#0563C1" vlink="#954F72">
	<div class=WordSection1>
		<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=656 style='width:492.35pt;border-collapse:collapse'>
			<tr style='height:39.0pt'>
				<td width=63 valign=bottom style='width:47.55pt;padding:0in 5.4pt 0in 5.4pt;height:39.0pt'></td>
				<td width=191 valign=bottom style='width:143.0pt;padding:0in 5.4pt 0in 5.4pt;height:39.0pt'></td>
				<td width=161 colspan=2 valign=bottom style='width:120.6pt;border-top:solid windowtext 1.0pt;border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid black 1.0pt;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:39.0pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:14.0pt;color:black'>Calls Handled<o:p></o:p></span></b></p></td>
				<td width=131 colspan=2 valign=bottom style='width:98.1pt;border:none;border-top:solid windowtext 1.0pt;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:39.0pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:14.0pt;color:black'>Sales<o:p></o:p></span></b></p></td>
				<td width=111 colspan=2 valign=bottom style='width:83.1pt;border:solid windowtext 1.0pt;border-right:solid black 1.0pt;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:39.0pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:14.0pt;color:black'>Additional Products<o:p></o:p></span></b></p></td>
			</tr>
			<tr style='height:54.75pt'>
				<td width=63 valign=bottom style='width:47.55pt;border:solid windowtext 1.0pt;border-right:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent ID<o:p></o:p></span></b></p></td>
				<td width=191 valign=bottom style='width:143.0pt;border:solid windowtext 1.0pt;border-left:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
				<td width=80 valign=bottom style='width:60.3pt;border-top:solid windowtext 1.0pt;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Calls Handled<o:p></o:p></span></b></p></td>
				<td width=80 valign=bottom style='width:60.3pt;border:solid windowtext 1.0pt;border-left:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Sales Calls Handled<o:p></o:p></span></b></p></td>
				<td width=73 valign=bottom style='width:55.05pt;border-top:solid windowtext 1.0pt;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Bounce Sales<o:p></o:p></span></b></p></td>
				<td width=57 valign=bottom style='width:43.05pt;border-top:solid windowtext 1.0pt;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Close Rate<o:p></o:p></span></b></p></td>
				<td width=55 valign=bottom style='width:41.55pt;border-top:none;border-left:solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;border-right:none;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>FCP Sales<o:p></o:p></span></b></p></td>
				<td width=55 valign=bottom style='width:41.55pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#9BC2E6;padding:0in 5.4pt 0in 5.4pt;height:54.75pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>DEPP Sales<o:p></o:p></span></b></p></td>
			</tr><tr style='height:15.0pt'>
				<td width=63 nowrap valign=bottom style='width:47.55pt;border:none;border-left:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>2062026<o:p></o:p></span></p></td>
				<td width=191 nowrap valign=bottom style='width:143.0pt;border:none;border-right:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>BECERRA, DOLORES<o:p></o:p></span></p></td>
				<td width=80 nowrap valign=bottom style='width:60.3pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:black'>21<o:p></o:p></span></p></td>
				<td width=80 nowrap valign=bottom style='width:60.3pt;border:none;border-right:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:black'>3<o:p></o:p></span></p></td>
				<td width=73 nowrap valign=bottom style='width:55.05pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:black'>0<o:p></o:p></span></p></td>
				<td width=57 nowrap valign=bottom style='width:43.05pt;background:#FFC7CE;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='color:#9C0006'>0%<o:p></o:p></span></b></p></td>
				<td width=55 nowrap valign=bottom style='width:41.55pt;border:none;border-left:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:black'>0<o:p></o:p></span></p></td>
				<td width=55 nowrap valign=bottom style='width:41.55pt;border:none;border-right:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:black'>0<o:p></o:p></span></p></td>
			</tr>
        </table>
    </div>
</body>
"""

mail.HtmlBody = emailBody
mail.send

print("Done...")
