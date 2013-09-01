<%
' +---------------------------------------------+
' | RSS Content Feed VBScript Class 1.0         |
' | © 2004 www.tele-pro.co.uk                   |
' | http://www.tele-pro.co.uk/scripts/rss/      |
' +---------------------------------------------+
'   
' Sample VBScript Code for the RSSContentFeed Class 
' Example Code - Simple
' sample_code1.asp

%>
<!-- #INCLUDE FILE="rss_content_feed_class.1.asp" -->
<%

'create object
Dim rss
Set rss= New RSSContentFeed

'get content
rss.ContentURL = "http://www.sofotex.com/download/xml/24.xml"
rss.GetRSS()

'display content
response.write "<h3>" & rss.ChannelTitle & "</h3>"

Dim i 
For Each i in rss.Results
  response.write rss.Links(i) & "<br>"
Next

'release object
Set rss= Nothing

%>
