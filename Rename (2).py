
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.http.request_options import RequestOptions
import json
from office365.sharepoint.files.system_object_type import FileSystemObjectType
import pandas as pd
import copy
import urllib
import re
import urllib.request
import urllib.parse 
from urllib import parse

# Log in to sharepoint site with the credentials
site_url = "https://nokia.sharepoint.com/sites/lte-a"
ctx = ClientContext(site_url).with_credentials(UserCredential("amine.errafay@nokia.com", "Imation12!"))
web = ctx.web.get().execute_query()
# Print the title of the site
print("Web title: {0}".format(web.properties['Title']))
i=0
#read the exceel file
df = pd.read_excel('worksheet.xlsx')
#iterare in the rows
for i,row in df.iterrows() :
 # extracts the sites
  site=row[0]
 # split the site by delimeter : in order to grab only the site
  site=site.split(":")
 #print(site)
 # take only the site
  site=site[1]
  #Remove the leading space
  site=site.lstrip()
  # remove the Nokia control documents
  #site=site.replace("/sites/lte-a/Nokia Controlled Documents/","Nokia Controlled Documents/")
 # add quotes around the site
  #site='"'+str(site)+'"'
 #print(site)
  with open("sites.txt","a") as f:
   print(site,file=f)
  
  # extracts the updatenames
  with open("updatenames.txt","a") as file:
   updatename=row[1]
 #Remove the leading space
   updatename=updatename.lstrip()
   #updatename='"'+str(updatename)+'"'
   print(updatename,file=file)
   #print(updatename)
  
# push the sites into a list 
f = open("sites.txt", "r")
sites_list= f.readlines()
print(sites_list)
# remove the spaces from the list
converted_list = []

for element in sites_list:
    converted_list.append(element.strip())

#print(converted_list)
#for i in sites_list:
  #print(i)
    #liat=sites_list[i]
   # print(sites_list)
 
  
#sites_list = sites_list.strip()
#print(sites_list)
# push the updatenames into list
f = open("updatenames.txt", "r")
updatesnames_list= f.readlines()
updated_list = []

#for k in updatesnames_list:
 #   updated_list.append(k.strip())
  #  print(updated_list)

#for i in range(len(updatesnames_list)) :
 #print(updatesnames_list[i],sites_list[i] )
 #i+1
#print (updatesnames_list[2])
#print (type((sites_list[2])))





#print(sites_list[7])
folder_url=sites_list[7]

#sites_list[2]=urllib.parse.unquote(sites_list[6])
#print(sites_list[7])


#b=urllib.parse.unquote(folder_url)
#print(b)
print(len(updatesnames_list))
list_test=["/sites/lte-a/Nokia Controlled Documents/LTE-BusProc(OBs)/10 - Project   Program Management","/sites/lte-a/Nokia Controlled Documents/LTE-BusProc(OBs)/04 - Design & Development"]
#for i in range(len(updatesnames_list)) :
 #x=ctx.web.get_folder_by_server_relative_path(converted_list[i]).rename(updated_list[i]).execute_query()
 #i+1
#print(updatesnames_list[1])



#print(list_test[1])
#print(updatesnames_list[1])


 
  #for updatename in updatenames:
   # print(updatename)
  #updatename=row[1]
  #updatename='"'+str(updatename)+'"'
 
#put updatenames in a list

  #print(updatename,file=file)
