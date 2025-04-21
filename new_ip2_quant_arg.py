import pandas as pd
import numpy as np
import math

def inputVar():
    parser = argparse.ArgumentParser(description='TMT proteomics filtering script.  LMO 2024')
  
    
    requiredNamed = parser.add_argument_group('required arguments')
    requiredNamed.add_argument('--input', metavar='<filepath>', type=str, help='full path of input file; ie ~Documents/myfolder/myfile', required=True)
    requiredNamed.add_argument('--output', metavar='<filepath>', type=str, help='full path of output directory; ie ~Documents/myfolder/', required=True)
    
    #requiredNamed.add_argument('-i', '--input', help='Input file name', required=True)
    #parser.parse_args(['-h'])

    args = parser.parse_args()
    print(args.input)
    return args.input, args.output



# copy and paste filename here:
#filename_no_ext = "250113_dmso_da_tmt_pd"
#filename = filename_no_ext + ".csv"
#print(filename)




    

#df = pd.read_excel(filename, skiprows = 1)
df = pd.read_csv(filename)
print(df.iloc[0:2])
#df = df..iloc[1:, :]
#df = pd.read_csv(oldFilePlus)


# sort by norm ratio
#print(df.iloc[0])
print("after")
#dfx = df.iloc[1: , :]
#print(dfx.iloc[0])
df = df.sort_values('NORM_RATIO_2_1', ascending=False)


# delete contaminant or reverse
df_merged = df.drop_duplicates(subset=['GENE'])



#DataFrame.drop_duplicates(subset=None, *, keep='first', inplace=False, ignore_index=False)


new_accession = df_merged.copy()
new_accession["first4access"] = df_merged["ACCESSION"]
new_accession['norm ratio copy'] = (df_merged['NORM_RATIO_2_1'])


# skip this for now
#new_accession['norm protein ratio log2'] = new_accession['norm ratio copy'].apply(log2)

new_accession['first4access2'] = new_accession['first4access'].str.slice(0,4)




df_spec_rev1 = new_accession[new_accession['first4access2'] != ('Reve')]
df_spec_rev = df_spec_rev1[df_spec_rev1['first4access2'] != ('cont')]



#df_spec_rev1 = new_accession[new_accession['first4access2'] != ('REVE')]
#df_spec_rev = df_spec_rev1[df_spec_rev1['first4access2'] != ('cont')]



del df_spec_rev['norm ratio copy']
del df_spec_rev["first4access"]
del df_spec_rev["first4access2"]

#change infinity to an int
#norm_ratio = df_spec_rev['NORM_RATIO_2_1']
norm_ratio = df_spec_rev['NORM_RATIO_2_1'].astype(str)
norm_ratio_inf = norm_ratio.str.replace('inf', '999999999')
norm_ratio_inf_nan = norm_ratio_inf.str.replace('nan', '-1')
df_spec_rev["NORM_RATIO_INT"] = norm_ratio_inf_nan


# remove 0 and infinity ### why is this not working?
df_spec_rev = df_spec_rev[df_spec_rev['NORM_RATIO_2_1'] < 999]
df_spec_rev = df_spec_rev[df_spec_rev['NORM_RATIO_2_1'] != ('0')]


#add a column with nor_

df_spec_rev['NORM_RATIO_INT'] = df_spec_rev['NORM_RATIO_INT'].astype(float)
df_spec_rev['NORM_PVALUE_1'] = df_spec_rev['NORM_PVALUE_1'].astype(float)






df_spec_rev["Norm_Ratio_Log2"] = np.log2(df_spec_rev["NORM_RATIO_INT"])
df_spec_rev["Norm_Pval_log10"] = (np.log10(df_spec_rev["NORM_PVALUE_1"]))
df_spec_rev["Norm_Pval_minuslog10"] = (np.multiply(df_spec_rev["Norm_Pval_log10"], -1))
del df_spec_rev["Norm_Pval_log10"]
#df_spec_rev["Norm_Ratio_Log2"] = np.log2(10)



print(df_spec_rev)

# idea: 10000000000 + the intensity


#print(norm_ratio_inf_nan)
#df_spec_rev['NORM_RATIO_2_1'] = df_spec_rev['NORM_RATIO_2_1'].astype(int)
#df_spec_rev['NORM_RATIO_2_1'] = norm_ratio_int

#export to specific sheet
#df_spec_rev.to_excel(r(user_file_name), sheet_name='raw data', index = False)



#print(df_spec_rev.head(15))


cols = list(df_spec_rev.columns.values)
#print(cols)
#print(cols[24])


#change this    
#newOrder = [0, 24, 6, 25, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
#cols = [cols[i] for i in newOrder]
#df_spec_rev = df_spec_rev[cols]



# remove insignificant p values
statsig1 = df_spec_rev[df_spec_rev['NORM_PVALUE_1'] < 0.05]
statsig = statsig1[statsig1['NORM_RATIO_2_1'] < 99]

#no_DMSO = statsig[statsig['NORM_RATIO_INT'] == '999999999']


#write to excel

import xlsxwriter



user_output_name = 'analyzed/' + filename_no_ext + "_analyzed.xlsx"
#print(user_output_name)

writer = pd.ExcelWriter(user_output_name, engine='xlsxwriter')

#is this necessary?
workbook = writer.book

workbook = xlsxwriter.Workbook(user_output_name, {'constant_memory': True})








#define specific sheets
df.to_excel(writer, sheet_name='Raw data', index = False)
df_spec_rev.to_excel(writer, sheet_name='Curated data', index = False)
statsig.to_excel(writer, sheet_name='Statistically significant', index = False)
#no_DMSO.to_excel(writer, sheet_name='Not in DMSO', index = False)

worksheet = writer.sheets['Curated data']


cell_format = workbook.add_format()
cell_format.set_bold()
#cell_format.set_font_color('green')
#cell_color = workbook.add_format()
#cell_color.set_font_color('red')

worksheet.set_column('A:A', None, cell_format)
worksheet.set_column('B:B', None, cell_format)
worksheet.set_column('C:C', None, cell_format)
worksheet.set_column('D:D', None, cell_format)
#worksheet.set_row('0:', None, cell_color)

#make formatting

#bold = workbook.add_format({'bold': True})
#workbook.write('1:1', 'A:A', bold)

# save the new excel file
writer.save()



