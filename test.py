import win32com.client as win32
from tableformat import topOfBreakdownTable

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jackson.ndiho@iqor.com'
mail.Subject = 'This is a Test'

# mail.HtmlBody = emailBody
mail.HtmlBody = topOfBreakdownTable
mail.send

print("Done...")
