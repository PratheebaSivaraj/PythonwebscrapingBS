import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
import sqlite3
import datetime

#printing the date and time
now = datetime.datetime.now()
print ("Current date and time : ")
date_time=now.strftime("%Y-%m-%d %H:%M:%S")
print(date_time)

D={}
density={}
keyword_list=[]
stopwords=['a','is','the','to','was','it','an','▼','>>>','▲','#']

#creating the xls
workbook=xlsxwriter.Workbook("d://mysheet.xls")
sheet=1

#creating the database
conn=sqlite3.connect("mysheet.db")
#creating the table
conn.execute("create table project12(words text,density float)")
#Creating a Cursor
c=conn.cursor()


try:
    req = urllib.request.Request( 'https://www.python.org/', data=None, 
             headers={ 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36' } 
                                      ) 
    page = urllib.request.urlopen(req) 

    soup = BeautifulSoup(page,"html.parser")     # page contains html content

# create a new bs4 object from the html data loaded 
    for script in soup(["script", "style"]): 
# remove all javascript and stylesheet code 
                 script.extract() 
# get text 
    text = soup.get_text()
# break into lines and remove leading and trailing space on each 
    lines = (line.strip() for line in text.split())


#remove stopwords
    for i in lines:
        if i not in stopwords:
            keyword_list.append(i)
    #print(keyword_list)


#frequency calculation
    for word in keyword_list:
            if word not in D:
                    D[word]=1
            else:
                    D[word]+=1
    #print(D)
    #print(D.keys())
    #print(D.values())



#density calculation
    for word in D:
        density[word]=(D[word]/len(keyword_list))*100
#inserting data in the worksheet
    worksheet=workbook.add_worksheet()
    row = 1
    col = 1
    worksheet.write('A1',date_time)
    for i,j in zip(density.keys(),density.values()):
        worksheet.write(row, col,i)
        worksheet.write(row, col + 1,j)
        row += 1
#creating the chart
    chart = workbook.add_chart({'type': 'line'})
   
    chart.add_series({'values': '=Sheet%d!$C$1:$C$150'%sheet})
    chart.set_title({'name': 'Results of Web Scraping'})
    chart.set_y_axis({'name': 'Word Density'})
    chart.set_x_axis({'name': 'Sno of Words'})
    worksheet.insert_chart('F5', chart)
#inserting data in database
    for i,j in zip(density.keys(),density.values()):
          conn.execute("INSERT INTO project12 (WORDS,DENSITY) VALUES(?,?)", (i,j))
    conn.commit()
#printing data stored in sqlite3
    data=c.execute("Select * from  project12")
    for i in data:
        print(i)

        
finally:
    conn.execute("drop table project12")
    workbook.close()
    conn.close()
    
