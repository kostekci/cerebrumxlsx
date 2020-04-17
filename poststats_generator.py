import pandas as pd

statsdf = pd.read_excel('WordStats.xlsx', sheet_name='Stats', usecols=['Word', 'Count', 'VideoName'])

#total word counts on all videos

total_wordsdf = statsdf.groupby(['Word']).sum().sort_values(by='Count', ascending=False)

writer = pd.ExcelWriter('PostWordStats.xlsx', engine='xlsxwriter')

total_wordsdf.to_excel(writer, sheet_name='WordsToKnowInGeneral')

#number of vidoes word seen as VideoCount

video_baseddf = statsdf.groupby(['Word'])['VideoName'].size().reset_index(name='VideoCount').sort_values(by='VideoCount', ascending=False)

#video based sum of VideoCount

video_prioritydf = statsdf.merge(video_baseddf, left_on='Word', right_on='Word').groupby(['VideoName']).sum()

video_prioritydf['VideoPriority'] = video_prioritydf['VideoCount'] / video_prioritydf['Count']

video_prioritydf = video_prioritydf.sort_values(by='VideoPriority', ascending=False)

video_prioritydf.to_excel(writer, sheet_name='WhatOrderToFollow')

#What words to check before watching that video

word_checkdf = video_prioritydf.merge(statsdf, left_on='VideoName', right_on='VideoName').drop(columns=['VideoCount','Count_x']).merge(total_wordsdf, left_on='Word', right_on='Word').sort_values(by=['VideoPriority','Count_y','Count'], ascending=False)

word_checkdf.to_excel(writer, sheet_name='WordsToCheckBeforeWatching')

writer.save()