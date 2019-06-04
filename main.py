import pandas as pd

sheet_index1 = [1,2,3,4,6,7,8,14,15,17,18,20,23,25,26,27,28,29,30,31,32,33,34,35,36,
              38,39,40,41,43,44,45,46,47,48,50,51,52,53,54,55,56]
sheet_index2 = [1,2,3,4,6,7,8,14,15,17,18,20,23,25,26,27,28,29,30,31,32,33,34,35,36]
sheet_index3 = [38,39,40,41,43,44,45,46,47,48,50,51,52,53,54,55,56]

filename = 'input.xlsx'
#-----------------------------------------------------------------------------------------------#
df_empty1 = pd.DataFrame(index = ['Partner Countries','Partner Country Code','Reporter Countries',\
                       'Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for i in sheet_index1:
    df_tmp = pd.read_excel(io = filename, sheet_name = str(i))
    if df_tmp is not None:
        df_empty1 = df_empty1.append(df_tmp)
df = df_empty1    
df = df.dropna()
rows = df.shape[0]
print(rows)
Partner_Countries = df['Partner Countries'].tolist()
Partner_Country_Code = df['Partner Country Code'].tolist()
Reporter_Countries = df['Reporter Countries'].tolist()
Reporter_Country_Code = df['Reporter Country Code'].tolist()
Partner_Reporter = set(zip(Partner_Countries, Reporter_Countries))
Partner_Countries_set = set(Partner_Countries)
Reporter_Countries_set = set(Reporter_Countries)
# print(Partner_Reporter)

Partner_Name_Code_set = set(zip(Partner_Countries, Partner_Country_Code))

Partner_Name_Code = {}
for item in Partner_Name_Code_set :
    Partner_Name_Code[item[0]] = item[1]
    
Reporter_Name_Code_set = set(zip(Reporter_Countries, Reporter_Country_Code))
Reporter_Name_Code = {}
for item in Reporter_Name_Code_set:
    Reporter_Name_Code[item[0]] = item[1]
    

# print(len(Partner_Reporter))
# print(len(Partner_Name_Code))
# print(len(Reporter_Name_Code))

total = []
for item in Partner_Reporter:
    sum_VWCblue = 0
    sum_VWCgreen = 0
    sum_VWCtotal = 0
    dict_item_sum = {}
    
    sub_df = df[(df['Partner Countries']==item[0]) & (df['Reporter Countries']==item[1])]
    
    rows_sub = sub_df.shape[0]
    
    sub_VWCblue = sub_df['VWCblue'].tolist()
    sum_VWCblue = sum(sub_VWCblue)
    
    sub_VWCgreen = sub_df['VWCgreen'].tolist()
    sum_VWCgreen = sum(sub_VWCgreen)

    sub_VWCtotal = sub_df['VWCtotal'].tolist()
    sum_VWCtotal = sum(sub_VWCtotal)
    
    dict_item_sum[item] = (sum_VWCblue,sum_VWCgreen,sum_VWCtotal)
    total.append(dict_item_sum)
    
# print(total) 

df_total = pd.DataFrame(index = ['Partner Countries','Partner Country Code','Reporter Countries',\
                       'Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for item in total:
    # print(item)
    keys = list(item.keys())
    keys = keys[0]
    # print(keys)
    value = list(item.values())
    value = value[0]
    # print(value)
    df_tmp = pd.DataFrame({'Partner Countries':[keys[0]], 'Partner Country Code':[Partner_Name_Code[keys[0]]], \
                         'Reporter Countries':[keys[1]], 'Reporter Country Code':[Reporter_Name_Code[keys[1]]], \
                         'VWCblue': [value[0]], 'VWCgreen': [value[1]],  'VWCtotal': [value[2]]})
    df_total = df_total.append(df_tmp)
writer1 = pd.ExcelWriter('first.xlsx')
df_total.to_excel(writer1,index=False,encoding='utf-8',sheet_name='Sheet')
writer1.save()

total_Partner = []
for item in Partner_Countries_set:
    sum_VWCblue = 0
    sum_VWCgreen = 0
    sum_VWCtotal = 0
    dict_item_sum = {}
    sub_df = df[(df['Partner Countries']==item)]
    
    sub_VWCblue = sub_df['VWCblue'].tolist()
    sum_VWCblue = sum(sub_VWCblue)
    
    sub_VWCgreen = sub_df['VWCgreen'].tolist()
    sum_VWCgreen = sum(sub_VWCgreen)

    sub_VWCtotal = sub_df['VWCtotal'].tolist()
    sum_VWCtotal = sum(sub_VWCtotal)
    
    dict_item_sum[item] = (sum_VWCblue,sum_VWCgreen,sum_VWCtotal)
    total_Partner.append(dict_item_sum)
    
df_total_Partner = pd.DataFrame(index = ['Partner Countries','Partner Country Code','VWCblue','VWCgreen','VWCtotal'])
for item in total_Partner:
    # print(item)
    keys = list(item.keys())
    keys = keys[0]
    # print(keys)
    value = list(item.values())
    value = value[0]
    # print(value)
    df_tmp = pd.DataFrame({'Partner Countries':[keys], 'Partner Country Code':[Partner_Name_Code[keys]], \
                         'VWCblue': [value[0]], 'VWCgreen': [value[1]],  'VWCtotal': [value[2]]})
    df_total_Partner = df_total_Partner.append(df_tmp)
    
    writer2 = pd.ExcelWriter('second.xlsx')
df_total_Partner.to_excel(writer2,index=False,encoding='utf-8',sheet_name='Sheet')
writer2.save()

total_Reporter = []
for item in Reporter_Countries_set:
    # print(item)
    sum_VWCblue = 0
    sum_VWCgreen = 0
    sum_VWCtotal = 0
    dict_item_sum = {}
    sub_df = df[df['Reporter Countries']==item]
    
    sub_VWCblue = sub_df['VWCblue'].tolist()
    # print(sub_VWCblue)
    sum_VWCblue = sum(sub_VWCblue)
    # print(sum_VWCblue)
    
    sub_VWCgreen = sub_df['VWCgreen'].tolist()
    sum_VWCgreen = sum(sub_VWCgreen)

    sub_VWCtotal = sub_df['VWCtotal'].tolist()
    sum_VWCtotal = sum(sub_VWCtotal)
    
    dict_item_sum[item] = (sum_VWCblue,sum_VWCgreen,sum_VWCtotal)
    total_Reporter.append(dict_item_sum)

df_total_Reporter = pd.DataFrame(index = ['Reporter Countries','Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for item in total_Reporter:
    # print(item)
    keys = list(item.keys())
    keys = keys[0]
    #print(keys)
    value = list(item.values())
    value = value[0]
    # print(value)
    df_tmp = pd.DataFrame({'Reporter Countries':[keys], 'Reporter Country Code':[Reporter_Name_Code[keys]], \
                         'VWCblue': [value[0]], 'VWCgreen': [value[1]],  'VWCtotal': [value[2]]})
    df_total_Reporter = df_total_Reporter.append(df_tmp)
    
    writer3 = pd.ExcelWriter('third.xlsx')
df_total_Reporter.to_excel(writer3,index=False,encoding='utf-8',sheet_name='Sheet')
writer3.save()
#----------------------------------------------------------------------------------------------#
df_empty2 = pd.DataFrame(index = ['Partner Countries','Partner Country Code','Reporter Countries',\
                       'Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for i in sheet_index2:
    df_tmp = pd.read_excel(io = filename, sheet_name = str(i))
    if df_tmp is not None:
        df_empty2 = df_empty2.append(df_tmp)
df = df_empty2
df = df.dropna()
rows = df.shape[0]
print(rows)
Partner_Countries = df['Partner Countries'].tolist()
Partner_Country_Code = df['Partner Country Code'].tolist()
Reporter_Countries = df['Reporter Countries'].tolist()
Reporter_Country_Code = df['Reporter Country Code'].tolist()
Partner_Reporter = set(zip(Partner_Countries, Reporter_Countries))
Partner_Countries_set = set(Partner_Countries)
Reporter_Countries_set = set(Reporter_Countries)
# print(Partner_Reporter)

Partner_Name_Code_set = set(zip(Partner_Countries, Partner_Country_Code))

Partner_Name_Code = {}
for item in Partner_Name_Code_set :
    Partner_Name_Code[item[0]] = item[1]
    
Reporter_Name_Code_set = set(zip(Reporter_Countries, Reporter_Country_Code))
Reporter_Name_Code = {}
for item in Reporter_Name_Code_set:
    Reporter_Name_Code[item[0]] = item[1]
    

# print(len(Partner_Reporter))
# print(len(Partner_Name_Code))
# print(len(Reporter_Name_Code))

total = []
for item in Partner_Reporter:
    sum_VWCblue = 0
    sum_VWCgreen = 0
    sum_VWCtotal = 0
    dict_item_sum = {}
    
    sub_df = df[(df['Partner Countries']==item[0]) & (df['Reporter Countries']==item[1])]
    
    rows_sub = sub_df.shape[0]
    
    sub_VWCblue = sub_df['VWCblue'].tolist()
    sum_VWCblue = sum(sub_VWCblue)
    
    sub_VWCgreen = sub_df['VWCgreen'].tolist()
    sum_VWCgreen = sum(sub_VWCgreen)

    sub_VWCtotal = sub_df['VWCtotal'].tolist()
    sum_VWCtotal = sum(sub_VWCtotal)
    
    dict_item_sum[item] = (sum_VWCblue,sum_VWCgreen,sum_VWCtotal)
    total.append(dict_item_sum)
    
# print(total) 

df_total = pd.DataFrame(index = ['Partner Countries','Partner Country Code','Reporter Countries',\
                       'Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for item in total:
    # print(item)
    keys = list(item.keys())
    keys = keys[0]
    # print(keys)
    value = list(item.values())
    value = value[0]
    # print(value)
    df_tmp = pd.DataFrame({'Partner Countries':[keys[0]], 'Partner Country Code':[Partner_Name_Code[keys[0]]], \
                         'Reporter Countries':[keys[1]], 'Reporter Country Code':[Reporter_Name_Code[keys[1]]], \
                         'VWCblue': [value[0]], 'VWCgreen': [value[1]],  'VWCtotal': [value[2]]})
    df_total = df_total.append(df_tmp)
writer4 = pd.ExcelWriter('forth.xlsx')
df_total.to_excel(writer4,index=False,encoding='utf-8',sheet_name='Sheet')
writer4.save()
#----------------------------------------------------------------------------------------------#
df_empty3 = pd.DataFrame(index = ['Partner Countries','Partner Country Code','Reporter Countries',\
                       'Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for i in sheet_index3:
    df_tmp = pd.read_excel(io = filename, sheet_name = str(i))
    if df_tmp is not None:
        df_empty3 = df_empty3.append(df_tmp)
df =df_empty3
df = df.dropna()
rows = df.shape[0]
print(rows)
Partner_Countries = df['Partner Countries'].tolist()
Partner_Country_Code = df['Partner Country Code'].tolist()
Reporter_Countries = df['Reporter Countries'].tolist()
Reporter_Country_Code = df['Reporter Country Code'].tolist()
Partner_Reporter = set(zip(Partner_Countries, Reporter_Countries))
Partner_Countries_set = set(Partner_Countries)
Reporter_Countries_set = set(Reporter_Countries)
# print(Partner_Reporter)

Partner_Name_Code_set = set(zip(Partner_Countries, Partner_Country_Code))

Partner_Name_Code = {}
for item in Partner_Name_Code_set :
    Partner_Name_Code[item[0]] = item[1]
    
Reporter_Name_Code_set = set(zip(Reporter_Countries, Reporter_Country_Code))
Reporter_Name_Code = {}
for item in Reporter_Name_Code_set:
    Reporter_Name_Code[item[0]] = item[1]
    

# print(len(Partner_Reporter))
# print(len(Partner_Name_Code))
# print(len(Reporter_Name_Code))

total = []
for item in Partner_Reporter:
    sum_VWCblue = 0
    sum_VWCgreen = 0
    sum_VWCtotal = 0
    dict_item_sum = {}
    
    sub_df = df[(df['Partner Countries']==item[0]) & (df['Reporter Countries']==item[1])]
    
    rows_sub = sub_df.shape[0]
    
    sub_VWCblue = sub_df['VWCblue'].tolist()
    sum_VWCblue = sum(sub_VWCblue)
    
    sub_VWCgreen = sub_df['VWCgreen'].tolist()
    sum_VWCgreen = sum(sub_VWCgreen)

    sub_VWCtotal = sub_df['VWCtotal'].tolist()
    sum_VWCtotal = sum(sub_VWCtotal)
    
    dict_item_sum[item] = (sum_VWCblue,sum_VWCgreen,sum_VWCtotal)
    total.append(dict_item_sum)
    
# print(total) 

df_total = pd.DataFrame(index = ['Partner Countries','Partner Country Code','Reporter Countries',\
                       'Reporter Country Code','VWCblue','VWCgreen','VWCtotal'])
for item in total:
    # print(item)
    keys = list(item.keys())
    keys = keys[0]
    # print(keys)
    value = list(item.values())
    value = value[0]
    # print(value)
    df_tmp = pd.DataFrame({'Partner Countries':[keys[0]], 'Partner Country Code':[Partner_Name_Code[keys[0]]], \
                         'Reporter Countries':[keys[1]], 'Reporter Country Code':[Reporter_Name_Code[keys[1]]], \
                         'VWCblue': [value[0]], 'VWCgreen': [value[1]],  'VWCtotal': [value[2]]})
    df_total = df_total.append(df_tmp)
writer5 = pd.ExcelWriter('fifth.xlsx')
df_total.to_excel(writer5,index=False,encoding='utf-8',sheet_name='Sheet')
writer5.save()