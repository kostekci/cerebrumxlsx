import pandas as pd

wordsdf = pd.read_excel('WordStats.xlsx', sheet_name='Stats', usecols=['Word', 'Count', 'VideoName'])

#total word counts on all videos

total_wordsdf = wordsdf.groupby(['Word']).sum().sort_values(by='Count', ascending=False)

writer = pd.ExcelWriter('PostWordStats.xlsx', engine='xlsxwriter')

total_wordsdf.to_excel(writer, sheet_name='WordsToKnowInGeneral')

#number of vidoes word seen as VideoCount

dfq = wordsdf.groupby(['Word'])['VideoName'].size().reset_index(name='VideoCount').sort_values(by='VideoCount', ascending=False)

#video based sum of VideoCount

xdf = wordsdf.merge(dfq, left_on='Word', right_on='Word').groupby(['VideoName']).sum().sort_values(by='VideoCount', ascending=False).drop(columns=['Count'])

xdf.to_excel(writer, sheet_name='WhatOrderToFollow')

#What words to check before watching that video

zdf = xdf.merge(wordsdf, left_on='VideoName', right_on='VideoName')

zdf.to_excel(writer, sheet_name='WordsToCheckBeforeWatching')

writer.save()