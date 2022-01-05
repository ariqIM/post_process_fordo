path='/home/ariq/Downloads/fordo_30K_annotation_results_9-12-21-partial_supervise_with_pad.xlsx'
path_out='/home/ariq/Downloads/fordo_30K_annotation_results_9-12-21-partial_supervise_with_pad_post_process.xlsx'
import xlrd
import pandas as pd
import fastwer
# EXCEL_FILES_FOLDER = 'excel_files/'
# excel_file_path = EXCEL_FILES_FOLDER+'read_excel.xlsx'
loc = path
# To open Workbook
df1 = pd.read_excel(
     path,
     engine='openpyxl',
)	
print(df1.columns)
cnt=0
avg=0
for index, row in df1.iterrows():
    #print( 'P' in df1.loc[index,'prediction'])
    #print(df1.loc[index,'prediction'])
    #print(df1[index,'prediction'],df1[index,'ground_truth'],df1[index,'cer_accu'])
      pred=str(df1.loc[index,'prediction'])
      gt=str(df1.loc[index,'ground_truth'])
    #   print(pred,gt)
      # pred=pred.replace('=','',len(pred))
      # gt=gt.replace('=','',len(gt))
      pred=pred.replace('P','p',len(pred))
      gt=gt.replace('P','p',len(gt))
      pred=pred.replace('o','০',len(pred))
      gt=gt.replace('o','০',len(gt))
      pred=pred.replace('2','২',len(pred))
      gt=gt.replace('2','২',len(gt))
      pred=pred.replace('K','k',len(pred))
      gt=gt.replace('K','k',len(gt))
      pred=pred.replace('O','০',len(pred))
      gt=gt.replace('O','০',len(gt))
      pred=pred.replace('0','০',len(pred))
      gt=gt.replace('0','০',len(gt))
      pred=pred.replace('8','৪',len(pred))
      gt=gt.replace('8','৪',len(gt))
      pred=pred.replace('S','s',len(pred))
      gt=gt.replace('S','s',len(gt))
      # pred=pred.replace('/','',len(pred))
      # gt=gt.replace('/','',len(gt))
      # pred=pred.replace('.','',len(pred))
      # gt=gt.replace('.','',len(gt))
    #   pred=pred.replace('৬','ড',len(pred))
    #   gt=gt.replace('৬','ড',len(pred))
      pred=pred.replace('C','c',len(pred))
      gt=gt.replace('C','c',len(gt))
      pred=pred.replace('M','m',len(pred))
      gt=gt.replace('M','m',len(gt))
      pred=pred.replace('N','n',len(pred))
      gt=gt.replace('N','n',len(gt))
      pred=pred.replace('W','w',len(pred))
      gt=gt.replace('W','w',len(gt))
      pred=pred.rstrip()
      gt=gt.rstrip()
      df1.loc[index,'prediction']=pred
      df1.loc[index,'ground_truth']=gt
      df1.loc[index,'cer_accu']=max(0,100 - fastwer.score_sent(pred,gt, char_level=True))
      avg+=df1.loc[index,'cer_accu']
      cnt+=1
# for index, row in df1.head(200).iterrows():
#     print( '২' in df1.loc[index,'prediction'])

print(avg)
print(cnt)
print(avg/cnt)
df1.to_excel(path_out)
#wb = xlrd.open_workbook(loc)
# sheetCust = wb.sheet_by_name('sheet1')
# rowsCount = sheetCust.nrows
# colsCount = sheetCust.ncols
# print("rowsCount: ",rowsCount," colsCount: ",colsCount)
# # scrip to print each value from the excel sheet
# for i in range(1,4):
#         col_name = sheetCust.cell_value(i, 3)
#         cell_val = sheetCust.cell_value(i, 4)
#         print(col_name,': ',cell_val)