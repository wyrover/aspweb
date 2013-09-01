<%
' +---------------------------------------------+
' | RSS Content Feed VBScript Class 1.0         |
' | © 2004 www.tele-pro.co.uk                   |
' | http://www.tele-pro.co.uk/scripts/rss/      |
' +---------------------------------------------+
'   
' Sample VBScript Code for the RSSContentFeed Class 
' Example Code - Amazon RSS Feed 
' sample_code_amazon.asp

%>
<!-- #INCLUDE FILE="rss_content_feed_class.1.asp" -->
<%

'create object
Dim rss
Set rss= New RSSContentFeed

'set the amazon parameters
Dim assocID, DevToken, Kwd, Mode, Title
assocID = "Your-AssociateID"
DevToken = "Your-Developer-Token"
Kwd = "ASP"
Mode = "books-uk"
Title = "ASP Books"

'get amazon rss content
Call rss.GetAmazonRSS(assocID, DevToken, Kwd, Mode, Title)

'display results
Dim i 
For Each i in rss.Results
  response.write rss.ItemHTML(i) & "<br>"
Next

'release object
Set rss= Nothing

%>