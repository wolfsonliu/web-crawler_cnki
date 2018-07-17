import pandas as pd

headers = [
    'SrcDatabase-来源库', 'Title-题名', 'Author-作者', 'Organ-单位', 'Source-文献来源', 'Keyword-关键词', 'Summary-摘要',
    'PubTime-发表时间', 'FirstDuty-第一责任人', 'Fund-基金', 'Year-年', 'Period-期', 'PageCount-页码'
]

#
# with open('CNKI-636597288626875000.txt', 'br') as f:
#     filelines = list()
#     for l in f:
#         filelines.append(l.decode('utf-8').rstrip())
#
# data = list()
# i = -1
#
# while True:
#     i = i + 1
#     if i == len(filelines):
#         break
#     if filelines[i] == '':
#         continue
#     else:
#         if filelines[i].find(headers[0]) >= 0:
#             datdict = dict()
#             initem = True
#             j = 0
#             while True:
#                 if not initem:
#                     # 已经不在 item 内
#                     break
#                 if filelines[i].find(headers[12]) >= 0:
#                     # 到了该item的最后一行
#                     initem = False
#                 if filelines[i].find(headers[j]) >= 0:
#                     # 找到了item的headers
#                     datdict[headers[j]] = filelines[i].replace(headers[j] + ':', '').strip()
#                     j = j + 1
#                 else:
#                     datdict[headers[j-1]] = ';'.join([datdict[headers[j-1]], filelines[i].strip()])
#                 i = i + 1
#             data.append(datdict)
#             i = i - 1
#

data = list()

for x in os.listdir('D:/analysis/cnki_year'):
    yeardata = pd.read_csv(os.path.join('D:/analysis/cnki_year', x), header=None)
    yeardata.columns = ['University', 'Subject', 'start', 'end', 'count']
    data.append(yeardata)

alldata = pd.concat(data, axis=0)

index = pd.MultiIndex.from_arrays([alldata['University'], alldata['Subject'], alldata['start']])

alldata.index = index

unstackdata = alldata['count'].unstack([1, 2]).copy()

unstackdata.to_csv('alldata.csv')