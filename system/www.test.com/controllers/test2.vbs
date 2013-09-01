'Import "system/libraries/clsThief.vbs"	
Import "system/libraries/RSSContentFeed.vbs"	

Class Test2


	Function index()		
		Dim rss
		Set rss= New RSSContentFeed

		'get content
		rss.ContentURL = "http://rss.mydrivers.com/rss.aspx?Tid=1"
		rss.GetRSS()

		'display content
		response.write "<h3>" & rss.ChannelTitle & "</h3>"

		Dim i 
		For Each i in rss.Results
		  response.write rss.Links(i) & "<br>"
		  response.write rss.ItemHTML(i) & "<br>"
		Next

		'release object
		Set rss= Nothing
	End Function


	Function job()
		'call AnyFunction() once every 6 hrs
		If (ScheduleTask("MyTaskName", "h", 6)) Then 
		  Call AnyFunction()
		end If

	End Function


	FUNCTION ScheduleTask(task_name, period, qty)
	  Dim RunNow
	  Dim last_date
	  Dim diff

	  'boolean result
	  RunNow = False

	  'chcek the value of app setting
	  last_date = Trim(Application("Sched_" & task_name))

	  'is value empty? maybe app just started
	  If (last_date = "") Then
		RunNow = True
	  Else
		'is value old?
		diff = DateDiff(period, last_date, Now())
		If (diff>=qty) Then RunNow = True
	  End if

	  'if scheduled to run now, set the app last run time
	  If (RunNow) Then Application("Sched_" & task_name) = Now()

	  'return result
	  ScheduleTask = RunNow 
	END FUNCTION 

	FUNCTION AnyFunction()
		Echo "Ö´ÐÐ´úÂë"
	END FUNCTION 

End Class