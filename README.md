<div align="center">

## The Daily Newbie \- Using the DateDiff\(\) Function


</div>

### Description

Newest in the series aimed at teaching newbies (and not so newbies) about the commands that VB has available. This edition gives a basic outline of DateDiff() usage and, as always, some copt-and-paste sample code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-the-daily-newbie-using-the-datediff-function__1-22717/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
<title>Daily Newbie - 04/25/2001</title>
</head>
<body bgcolor="#FFFFFF">
<p> </p>
<p class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoTitle"><font size="7"><strong>The
Daily Newbie</strong></font></p>
<p align="center" class="MsoTitle"><strong>&#8220;To Start Things
Off Right&#8221;</strong></p>
<p align="center" class="MsoTitle"><font size="1">Third
Edition
April 26,
2001
Free</font></p>
<p align="center" class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p class="MsoNormal"><font face="Arial"><strong>About this
feature:</strong></font></p>
<p class="MsoBodyText"><font size="2" face="Arial">Today's Newbie code is the result of a request from a reader of yesterdays (thanks for the suggestion BigCalm).</font></p>
<p class="MsoNormal"><font size="2" face="Arial">Today I am going to discuss the DateDiff() function. Many newbies (and some more experienced coders) spend many hours writing code to do the exact same things that they could do with a single call to DateDiff(). I hope to show you what this function is, how to use it, and how it can make your coding MUCH easier. </font></p>
<p class="MsoNormal"><font size="2" face="Arial">.</font></p>
<p class="MsoNormal"
style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Today&#8217;s Keyword:</strong>
               </font><font
size="4" face="Arial"> DateDiff()</font></p>
<p class="MsoNormal"
style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Name Derived
From:    </strong>     </font>
 <font size="2" face="Arial">"Date Difference "</em></font></p>
 </p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Used for  </strong>
Determining the difference between two dates or times.</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>VB Help Description: </strong>        Returns a Variant (Long) specifying the number of time intervals between two specified dates.</font></p>
<font size="2" face="Arial"><strong>Plain
English:  </strong> Makes adding and subtracting dates easier by allowing you to pass in a start and end date and get difference back. This difference can be in any valid date/time increment (day, week, month, quarter, year, hour, minute, second).</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Syntax:    </strong>               X = DateDiff(Interval, StartDateTime, EndDateTime)</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Usage:    </strong>                intDayCount = DateDiff("d","01/01/1995", "01/01/2001")</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Parameters:    </strong>                <li>Interval - The type of results you want returned. These are:
			yyyy	=Year
			q	=Quarter
			m	=Month
			y	=Day of year
			d	=Day
			w	=Weekday
			ww	=Week
			h	=Hour
			n	=Minute
			s	=Second
<br><br>
<li>StartDateTime - Any valid date, time, or datetime combination. Examples: "01/01/2000" , "01/01/2000 12:25 AM" , "16:30"
<li>EndDateTime - Same criteria as StartDateTime. This is the data the start date will be subtracted from.
</font></p>
<p class="MsoNormal"
style="margin-left:135.35pt;text-indent:-135.35pt"><font size="2"
face="Arial"><strong>Copy & Paste Code:</strong></font></p>
    <p class="MsoNormal"
    style="margin-left:135.35pt;text-indent:-135.35pt"><font
    size="2" face="Arial"></font></p>
       <pre>
<font size="2" face="Arial"><code></code></font></pre>
       <pre
       style="margin-left:1.25in;text-indent:.35pt;tab-stops:45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><font
size="2" face="Arial"><code>
				Dim StartDate As Date<br>
				Dim EndDate As Date<br>
				Dim Interval As String
				<br>
				StartDate = InputBox ("Start Date:")<br>
				EndDate = InputBox ("End Date:")<br>
				Interval = InputBox ("Return In:  s=seconds, m=Minutes h=Hours, d=Days, ww=Weeks, w=WeekDays, yyyy=years"
				MsgBox DateDiff(Interval, StartDate, EndDate)
				</code></font></pre>
 <p class="MsoNormal"
 style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"> </p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Notes: </strong></font></p>
<font size="2" face="Arial">
The DateDiff() function is one of the most useful ones VB has to offer. It literally replaces thousands of lines of code, takes in account leap years, knows how many days and weeks are in a month, and many other things that typically trip up home-brewed date code. Let's face it...those Microsoft guys can write some decent code. They went to a lot of trouble to create these functions in lower level languages so we could just call it and get a result back. Besides being much less likely to error out that your own code, it is also exponentially faster since it exists as true bytecode.<br><br>
		<br>
		<b>A couple of things to watch out for in the DateDiff() Function are:</b><br><br>
		<li><b>Times can mess you up. </b>When you call DateDiff without specifying a time (i.e. "01/10/200" instead of "01/01/2000 9:25:00"), DateDiff assumes a time of midnight (00:00:01). This can have the effect of "skipping" a day if you aren't careful. Check your results a few times and adjust your dates or times to make it right. Once you have it, it will always work the same.<br><br>
		<li><b>Switching dates will return negative values.</b> Not a tragedy, but something you should be aware of.
		<br>
		<br>
		Well I hope today's newsletter has helped save some newbie coders out there from pulling out clumps of hair over date manipulation. If you need more details on using DateDiff() please let me know.
		</font></p>
</body>
</html>

