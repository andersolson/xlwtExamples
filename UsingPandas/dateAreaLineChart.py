from datetime import date
from dateutil.relativedelta import relativedelta

count = 0
dates = []

#Generate 12 months worth of dates and append them to a list
while count < 12:
    count += 1
    iDate = date.today() + relativedelta(months=+count)
    #print(count)
    #print(iDate)
    dates.append(iDate)

print(dates)


    