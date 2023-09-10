import openpyxl
import pandas as pd
import itertools


input_data = input("Please enter comma separated codes: ")
input_data = input_data.upper()
input_data = input_data.split(",")
df = pd.read_excel('Input.xlsx')
df = df.map(lambda x: x.replace('_', '') if isinstance(x, str) else x)
df['From'] = df['From'].str.strip()
df['To'] = df['To'].str.strip()
df['From'] = pd.to_datetime(df['From'], format ='%H:%M')
df['From'] = df['From'].dt.time
df['To'] = pd.to_datetime(df['To'], format ='%H:%M')
days = ['Saturday','Sunday','Monday','Tuesday','Wednesday','Thursday','Friday']
df['Day'] = pd.Categorical(df['Day'],categories=days,ordered=True)
df['To'] = df['To'].dt.time
df = df.sort_values(by =['Day','From'])
df1 = df[df['Code'].isin(input_data)]
#1 code, 2 name, 4 type, 5 day,6 from, 7 to
def get_combinations(n, df):
    data_list = df.values.tolist()
    return list(itertools.combinations(data_list, n))

def is_valid(com):
    for i in range(len(com)):
        if(com[i][11] == 'Closed'):
            return False
        for j in range(len(com)):
            if i == j:
                continue
            else:
                if com[i][5] == com[j][5] and ((com[i][6] > com[j][6] and com[i][6] < com[j][7]) or com[i][6] == com[j][6] or com[i][7] == com[j][7]):
                    return False
    types = True

    for i in range(len(com)):
        count = 1
        for j in range(len(com)):
            if(i == j):
                continue
            else:
                if(com[i][1] == com[j][1]):
                    count += 1
                    if(count > 2):
                        return False
                    if(com[i][4] == com[j][4]):
                        print(str(com[i][4]))
                        return False
                    if('MTHN' in com[i][1] and (com[i][3] != com[j][3])):
                        return False

        if (count == 1 and 'GENN' not in com[i][1] and 'CCEN' not in com[i][1]):
            return False
    return True




genns = 0
for i in input_data:
    if 'GENN' in i or 'CCEN' in i:
        genns += 1
n = genns + 2 * (len(input_data)-genns)
combinations = get_combinations(n, df1)
valid = []
for i in combinations:
    if is_valid(i):
        valid.append(i)

#for j in range(len(combinations)):
   # if(is_valid(combinations[j])):
      #  for i in range(len(combinations[j])):
          #  print(  str(combinations[j][i][2])+ '     '  +str(combinations[j][i][5])+'    '+str(combinations[j][i][6]) + '  ' + str(combinations[j][i][7]) + '   ' + str(combinations[j][i][4]))
        #print('\n \n \n')

wb = openpyxl.Workbook()
sheet = wb.active

i = 1

for tup in valid:
    for lst in tup:
        j = 1
        indcies = [1,2,4,5,6,7]
        counter = 0
        for item in lst:
            if(counter in indcies):
                sheet.cell(row=i, column=j, value=item)
                j += 1
            counter += 1
        i += 1
    i += 3


wb.save("output.xlsx")
print(str(len(valid)) + ' schedule(s) generated!')
