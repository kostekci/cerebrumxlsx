import pandas as pd
import string 
from os import listdir
from os.path import isfile, join

paths=['d:\\projects\\german_only_text\\seg\\','d:\\projects\\german_only_text\\eg\\','d:\\projects\\german_only_text\\other\\']

#file names in all dirs
files=[]

for path in paths:
    files += [path+f for f in listdir(path) if isfile(join(path, f))]

#final df for file based wordcounts on all files
dfz = pd.DataFrame({'Word':[], 'Count':[], 'VideoName':[]}, columns = ['Word', 'Count', 'VideoName'])

for file in files:
    text = open(file, encoding="utf8")
    
    #dict to hold word counts on single file on every iteration 
    d = dict() 
      
    #Loop through each line of the file 
    for line in text:
        #Remove the leading spaces and newline character 
        line = line.strip() 
       
        #lowercase to avoid case mismatch 
        line = line.lower() 
      
        #Remove the punctuation marks from the line 
        line = line.translate(line.maketrans("", "", string.punctuation)) 
      
        #Split the line into words 
        words = line.split(" ") 
      
        #Iterate over each word in line 
        for word in words:
            if word=="":
                continue
            #Check if the word is already in dictionary 
            if word in d: 
                #Increment count of word by 1 
                d[word] = d[word] + 1
            else: 
                #Add the word to dictionary with count 1 
                d[word] = 1
            
    dsorted = {k: v for k, v in sorted(d.items(), key=lambda item: item[1])}

    #df for word counts on single file created form sorted dict
    dfy = pd.DataFrame(dsorted.items(),columns=['Word', 'Count'])
    dfy['VideoName']=text.name
    
    dfz = dfz.append(dfy)
    text.close()
#end of for (iterated on all files)

dfz = dfz.sort_values(by=['VideoName','Count'], ascending=False)

#Create a Pandas Excel writer
writer = pd.ExcelWriter('WordStats.xlsx', engine='xlsxwriter')

#Convert the dataframe to an excel object
dfz.to_excel(writer, sheet_name='Stats')

#Close the Pandas Excel writer and output the Excel file.
writer.save()
