<%
' +---------------------------------------------+
' | RSS Content Feed VBScript Class 1.0         |
' | © 2004 www.tele-pro.co.uk                   |
' | http://www.tele-pro.co.uk/scripts/rss/      |
' +---------------------------------------------+
'   
' Sample VBScript Code for the RSSContentFeed Class 
' Example Code - Setting Properties
' sample_code_properties.asp

%>
<!-- #INCLUDE FILE="rss_content_feed_class.1.asp" -->
<%

'create object
Dim rss
Set rss= New RSSContentFeed

'set content url
rss.ContentURL = "http://freenewsfeed.newsfactor.com/rss" 
 
'set Post data
rss.PostData = "id=123&cat=news" 
 
'set content url
rss.MaxResults = 5 

'set cache
rss.Cache = "\\nas03ent\domains\n\nicksumner.com\user\htdocs\rsscache\" 
 
'cache items for 2 days
rss.CacheDays = 2 
 
'from cache?
if rss.FromCache Then
  'item was returned from cache
End If 

'display properties
Response.Write "<br> rss.MaxResults: " & rss.MaxResults 
Response.Write "<br> rss.ContentURL: " & rss.ContentURL 
Response.Write "<br> rss.PostData: " & rss.PostData 
Response.Write "<br> rss.CacheDays: " & rss.CacheDays 

'release object
Set rss= Nothing

%>