<%         
'Sends an email         
Dim mail         
YourName = Request.form("YourName")       
Body = Request.form("Body")       
EmailAddress = Request.form("EmailAddress")        
Set mail = Server.CreateObject("CDO.Message")         
mail.To = "enquiries@webtank.co.uk"         
mail.From = "enquiries@webtank.co.uk"         
mail.Subject = Request.Form("Subject")         
mail.TextBody = YourName + vbcrlf + EmailAddress + vbcrlf + Body         
mail.Send()       
Response.Redirect "http://www.webtank.co.uk/thankyou.asp"       
'Destroy the mail object!         
Set mail = nothing         
%>