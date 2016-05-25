# coding: utf-8
import urllib2
import urllib
import re
from bs4 import BeautifulSoup
import urlparse
import xlsxwriter
import traceback


def open_with_retries(url):
    attempts = 20
    for attempt in range(attempts):
        try:
            return opener.open(url, timeout = 20)
        except:
            if attempt == attempts - 1:
                print "Time Out"
                #raise


workbook = xlsxwriter.Workbook('Artist_Art_A.xlsx')
worksheet = workbook.add_worksheet()

ro = 0
co = 0

opener = urllib2.build_opener()
opener.addheaders = [('User-agent', 'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6' )]

# First web
response = open_with_retries(unicode("http://amma.artron.net/artronindex_artist.php"))
content = response.read()
pattern = re.compile(u'<li><a href="artronindex_pic.php(.*?) title="(.*?)".*?</a></li>',re.S)
items = re.findall(pattern,content)
for item in items:
    # Second
    try:
        res2 = open_with_retries(str("http://amma.artron.net/artronindex_pic.php?artist="+item[1]))
        soup = BeautifulSoup(res2.read())
        tables = soup.find_all('tbody')
        if len(tables)>0:
            table = tables[0]
            rows = table.findChildren('tr')
#        worksheet.write(ro,co,item[1].decode('utf-8'))
 #   co=co+1
            print item[1]
            for row in rows:
##                cells = row.findChildren('td')
##                for cell in cells:
##                    value = unicode(cell.string)
##                    worksheet.write(ro,co,value)
##                    co = co+1
##                    print value
                links = row.find_all('a',href=True)
                for link in links:
                    url =  link['href']
                    parsed = urlparse.urlparse(url)
                    sort = urlparse.parse_qs(parsed.query)['sort'][0]
                    labe = urlparse.parse_qs(parsed.query)['labe'][0]
                    f = {'sort':sort,'labe':labe}
                    later = urllib.urlencode(f)
       
                    # Third
                    res3 = open_with_retries(str("http://amma.artron.net/artronindex_auctionseason.php?name="+item[1]+"&"+later))
                    soup2 = BeautifulSoup(res3.read())  
                    ttables = soup2.findChildren('tbody')
                    if len(tables)>0:
                        ttable = ttables[0]
                        rrows = ttable.findChildren('tr')
                        for rrow in rrows:                  
                             ccells = rrow.findChildren('td')
                             for ccell in ccells:
                                 vvalue = unicode(ccell.string)
                                 worksheet.write(ro,co,vvalue)
                                 co=co+1
                                 print vvalue
                        ro = ro+1
                        co = 0
                        	# Fourth
                             llinks = rrow.find_all('a',href=True)
                             for llink in llinks:
                                uurl = llink['href']
                                parsed = urlparse.urlparse(uurl)
                                picid = urlparse.parse_qs(parsed.query)['picid'][0]                                      
                                res4 = open_with_retries(str("http://auction.artron.net/paimai-"+picid))
                                soup4 = BeautifulSoup(res4.read())
                                table3 = soup4.find_all('div',{'class':'layL'})[0]
                                row3s = table3.findChildren('tr')
                                for row3 in row3s:
                                    cell3s = row3.findChildren('td')
                                    for cell3 in cell3s:
                                        value3 = cell3.string
                                        if value3 != None:
                                            value4 = unicode(value3)
                                            worksheet.write(ro,co,value4)
                                            co=co+1
                                            print value4
                                ro = ro+1
                                co = 0
    except Exception:
        traceback.print_exc()
workbook.close()
    
