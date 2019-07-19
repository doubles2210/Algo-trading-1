
#!/usr/bin/python
import smtplib, socket
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import yfinance as yf

from datetime import date

from datetime import timedelta

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('gspread-date-09-06-2019-8520c2069f94.json', scope)
gc = gspread.authorize(credentials)

wks = gc.open("Share Trading-AutoUpdate-More").sheet1

if credentials.access_token_expired:
     gc_client.login()

entire_sheet = wks.get_all_values()
header_removed_entire_sheet = entire_sheet[1:] # Remove header row
# email-function-open
def send_mail(from_email, to_email, email_sub, email_content):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = email_sub
    email_body = email_content
    msg.attach(MIMEText(email_body, 'plain'))
    complete_email = msg.as_string()
    server.sendmail(from_email, to_email, complete_email)
# email-function-closed

# email-login-open
##server = smtplib.SMTP('mail.4photobooks.in', 587)
##server.login('saket@4photobooks.in','Prem2210')
# email-login-closed
end_date = date.today() - timedelta(1)
count = 2
for row in header_removed_entire_sheet:

    if str(row[0]):
        print(row)
        number_of_days = int(row[1])
        x_days_before_date = date.today() - timedelta(number_of_days)
        ticker_yahoo = yf.Ticker(row[0])
        get_process_history=ticker_yahoo.history(start=x_days_before_date, end=end_date)
        last_price_upd=ticker_yahoo.info['regularMarketPrice']
        #get_process_history=get_history(row[0], start=x_days_before_date, end=end_date)
        hist_price_upd = (get_process_history.Close.max())
#       print(hist_price_upd)
#       print(last_price_upd)
        print("yoo")
        #update x days high - E
        hist_price_upd_col = 'E' + str(count)
        result=wks.update_acell(hist_price_upd_col, hist_price_upd)
        #update todays price - Column H
        last_price_upd_col = 'H' + str(count)
#       print (last_price_upd_col)
        result=wks.update_acell(last_price_upd_col, last_price_upd)

    #Calculate + update days increase - Column I
        days_per_change = round(((last_price_upd - hist_price_upd)/hist_price_upd)*100,2)
        days_per_change_col = 'I' + str(count)
        result=wks.update_acell(days_per_change_col, days_per_change)

    #Calculate + update last investment increase - Column J
        if str(row[5]):
            last_investment_amount = float(row[5])
            last_investment_per_change = round(((last_price_upd - last_investment_amount)/last_investment_amount)*100,2)
            last_investment_change_col = 'J' + str(count)
            result=wks.update_acell(last_investment_change_col, last_investment_per_change)


        #Notify + Update date change
        if str(row[10]) == 'Y':
        # if higher side limit is hit
            if (days_per_change) > float("0.00"):
                if (days_per_change) > float(row[2]):
                    email_header = row[0] +' ('+ ' Upper-days for days ' + row[1] +' )'
                    email_body = row[0] +' ('+ ' upper-days for days ' + row[1] +' ) % Chg is '+ str(days_per_change) + '\n last-high-price '+ row[4] + '\n current-price ' + str(last_price_upd) + '\n our-purchase-price ' + row[5]
##                    send_mail('saket@4photobooks.in', 'ssinho@gmail.com', email_header, email_body)
                    print ("upper-days")
            elif (abs(days_per_change) > float(row[2])):
                email_header = row[0] +' ('+ ' Lower-days for days ' + row[1] +' )'
                email_body = row[0] +' ('+ ' lower-days for days ' + row[1] +' ) % Chg is '+ str(days_per_change) + '\n last-high-price '+ row[4] + '\n current-price ' + str(last_price_upd)  + '\n our-purchase-price ' + row[5]
##                send_mail('saket@4photobooks.in','ssinho@gmail.com', email_header, email_body)
                print ("lower-days")


        #Notify + + Update last investment
        if str(row[10]) == 'Y' and str(row[5]):
        # if higher side limit is hit
            if (last_investment_per_change) > float("0.00"):
                if (last_investment_per_change) > float(row[3]):
                    email_header = row[0] + ' Investment-higher-for-days ' + row[1] +' )'
                    email_body = row[0] +' ('+ ' upper-last-investment for days ' + row[1] +' ) % Chg is '+ str(last_investment_per_change) + '\n last-high-price '+ row[4] + '\n current-price ' + str(last_price_upd) + '\n our-purchase-price ' + row[5]
##                    send_mail('saket@4photobooks.in', 'ssinho@gmail.com', email_header, email_body )
                    print ("upper-last-investment")
            elif (abs(last_investment_per_change) > float(row[3])):
                email_header = row[0] +' ('+ ' Investment-lower-for-days ' + row[1] +' )'
                email_body = row[0] +' ('+ ' lower-last-investment for days ' + row[1] +' ) % Chg is '+ str(last_investment_per_change) + '\n last-high-price '+ row[4] + '\n current-price ' + str(last_price_upd) + '\n our-purchase-price ' + row[5]
##                send_mail('saket@4photobooks.in', 'ssinho@gmail.com', email_header, email_body)
                print ("lower-last-investment")
    count = count +1
##server.quit()
