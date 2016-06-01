from threading import Thread
import pandas as pd
import urllib
import json
import os

def process_id(id):
    """process a single ID"""
    # fetch the data
    search = id 
    search = search.replace('&','').replace('#','').replace(' ','%20')
    url = 'https://maps.googleapis.com/maps/api/place/textsearch/json?query='+search+'&key=INSERT API KEY HERE'
    json_obj = urllib.request.urlopen(url).read().decode('UTF-8')
    data = json.loads(json_obj)    
    return data
        
def process_range(id_range, store=None):
    """process a number of ids, storing the results in a dict"""
    if store is None:
        store = {}
    for id in id_range:
        store[id] = process_id(id)
    return store
    

def threaded_process_range(nthreads, id_range):
    """process the id range in a specified number of threads"""
    store = {}
    threads = []
    # create the threads
    for i in range(nthreads):
        ids = id_range[i::nthreads]
        t = Thread(target=process_range, args=(ids,store))
        threads.append(t)

    # start the threads
    [ t.start() for t in threads ]
    # wait for the threads to finish
    [ t.join() for t in threads ]
    return store
    
def main():
    #loading data from csv files
    address_data = pd.read_csv(r"C:\Utkarsh\GIT\Python\GoogleAPI\VendorData.csv")#INSERT DATA FILE PATH
    #names and addresses to be searched
    searches = address_data['VENDOR_ADDRESS']
    file = open(os.path.expanduser(r"'C:\Utkarsh\GIT\Python\GoogleAPI\Results.csv"), "wb")#INSERT RESULT PATH
    file.write(b"Searches,Name,Address,Type,Latitude,Longitude,Isflag" + b"\n")
    noofthreads =10 #no. of threads
    outdata=threaded_process_range(noofthreads, searches)
    for key,value in outdata.items():
       for item in value['results']: 
           formatted_address = item['formatted_address'].replace(',','')
           types = item['types']
           type = '-'.join(types)
           name = item['name'].replace(',','')
           geo = item['geometry']
           loc = geo['location']
           lat = loc['lat']
           lng = loc['lng']
           if 'establishment' in type:
               Isflag='1'
           else:
                Isflag='0'
           
           try:
               CombinedString = key.replace(',','') + "," + name + "," + formatted_address + "," + type + "," + str(lat) + "," + str(lng) + "," + Isflag + '\n'
               file.write(bytes(CombinedString, encoding="ascii", errors='ignore'))
               print(CombinedString)
           except:
               a=1
           CombinedString=""
    file.close()
   
    #Formating the results into a new Excel Sheet
    df = pd.read_csv(r'C:\Utkarsh\GIT\Python\GoogleAPIt\Results.csv')#INSERT RESULT PATH
    df = df.drop_duplicates('Searches')
    #Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(r'C:\Utkarsh\GIT\Python\GoogleAPIt\Results.xlsx', engine='xlsxwriter',)
    df.to_excel(writer, sheet_name='Enwave',index=False)
    workbook = writer.book
    worksheet = writer.sheets['BREP']
    formater = workbook.add_format({'border':1})
    worksheet.set_column(0,len(df.columns)-1,15,formater)
    worksheet.freeze_panes(1,0)
    worksheet.autofilter(0,0,0,len(df.columns)-1)
    writer.save()

    #Open the obtained excel output
    os.system(r'start excel.exe "C:\Utkarsh\GIT\Python\GoogleAPIt\Results.xlsx"')#INSERT RESULT PATH
    
main()
 