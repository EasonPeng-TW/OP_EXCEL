import pandas as pd
import requests
import datetime as dt
from datetime import timedelta
import time
import datetime
import matplotlib.pyplot as plt
from matplotlib.font_manager import findfont, FontProperties
import openpyxl
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False

class Op():
    def __init__(self):
        self.dict_web = {'cnp_three':'https://www.taifex.com.tw/cht/3/callsAndPutsDate', #三大法人 put&call 分計
                         'cnp_big10':'https://www.taifex.com.tw/cht/3/largeTraderOptQry', #十大 put&call
                         'fu_three':'https://www.taifex.com.tw/cht/3/futContractsDate', #三大法人 大小台 
                         'fu_big10':'https://www.taifex.com.tw/cht/3/largeTraderFutQry' #十大期貨
                        }
        
    def crawl(self, search_web, search_date= '1'):
        # dt.datetime.now().strftime("%Y/%m/%d")
        data = {'queryType': '1',
                'goDay': '',
                'doQuery': '1',
                'dateaddcnt': '-1',
                'queryDate': search_date}
        headers = {'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.113 Mobile Safari/537.36',
        }
        res = requests.post(self.dict_web[search_web], headers=headers, data=data, timeout = 5).text
        table = pd.read_html(res)
        return table, search_web, search_date 

    def crawl_big10(self, search_web, search_date='1'):
        data = {'datecount': '',
                'contractId2': 'all',
                'queryDate': search_date, 
                'contractId': 'all'}
        headers = {'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Mobile Safari/537.36'}
        res = requests.post(self.dict_web[search_web], headers=headers, data=data, timeout = 5).text
        table = pd.read_html(res)
        return table, search_web, search_date 

    def analysis_op(self, file_mode):
        get_days = int(input('請輸入OP回測天數: '))
        i = 0
        op_data = []
        while True:
            try:
                day_delta = i
                i += 1
                today_to_past = (dt.datetime.now() - timedelta(days= int(day_delta)))
                time.sleep(1)
                true_format = today_to_past.strftime("%Y/%m/%d")

                table, search_web, search_date = self.crawl(search_web='cnp_three', search_date=true_format)
                sfbc_nb = int(table[2].iloc[5][10]) #自營商 BC BP SC SP
                sfbp_nb = int(table[2].iloc[8][10])
                sfsc_nb = int(table[2].iloc[5][12])
                sfsp_nb = int(table[2].iloc[8][12])

                bc_nb = int(table[2].iloc[7][10]) #外資 BC BP SC SP
                bp_nb = int(table[2].iloc[10][10])
                sc_nb = int(table[2].iloc[7][12])
                sp_nb = int(table[2].iloc[10][12])
                # return table, bc_nb, bp_nb, sc_nb, sp_nb, sfbc_nb, sfbp_nb, sfsc_nb, sfsp_nb, search_date
                               
                print(true_format + ' 寫入檔案...')
                op_data.append([search_date, sfbc_nb, sfbp_nb, sfsc_nb, sfsp_nb, bc_nb, bp_nb, sc_nb, sp_nb])

                if i == get_days:
                    with open('op_data.csv', file_mode, encoding = 'utf-8-sig') as f:
                        # f.write('日期, 自營bc, 自營bp, 自營sc, 自營sp, 外資bc, 外資bp, 外資sc, 外資sp\n') #file mode=w 時要打開
                        for p in op_data:
                            f.write(p[0] + ',' + str(p[1]) + ',' + str(p[2]) + ',' +str(p[3]) + ',' + str(p[4]) + ',' + str(p[5]) + ',' + str(p[6])  + ',' + str(p[7]) + ',' + str(p[8]) +'\n')
                    break
            except:
                print(true_format)
                print('可能是假日或沒開盤')
                if i == get_days:
                    with open('op_data.csv', file_mode, encoding = 'utf-8-sig') as f:
                        # f.write('日期, 自營bc, 自營bp, 自營sc, 自營sp, 外資bc, 外資bp, 外資sc, 外資sp\n') #file mode=w 時要打開
                        for p in op_data:
                            f.write(p[0] + ',' + str(p[1]) + ',' + str(p[2]) + ',' +str(p[3]) + ',' + str(p[4]) + ',' + str(p[5]) + ',' + str(p[6])  + ',' + str(p[7]) + ',' + str(p[8]) +'\n')
                    break

    def analysis_op_excel(self):
        get_days = int(input('請輸入OP回測天數: '))
        i = 0
        op_data = []
        while True:
            try:
                day_delta = i
                i += 1
                today_to_past = (dt.datetime.now() - timedelta(days= int(day_delta)))
                time.sleep(1)
                true_format = today_to_past.strftime("%Y/%m/%d")

                table, search_web, search_date = self.crawl(search_web='cnp_three', search_date=true_format)
                print('Connecting to web...')
                sfbc_nb = int(table[2].iloc[5][10]) #自營商 BC BP SC SP
                sfbp_nb = int(table[2].iloc[8][10])
                sfsc_nb = int(table[2].iloc[5][12])
                sfsp_nb = int(table[2].iloc[8][12])

                bc_nb = int(table[2].iloc[7][10]) #外資 BC BP SC SP
                bp_nb = int(table[2].iloc[10][10])
                sc_nb = int(table[2].iloc[7][12])
                sp_nb = int(table[2].iloc[10][12])
                # return table, bc_nb, bp_nb, sc_nb, sp_nb, sfbc_nb, sfbp_nb, sfsc_nb, sfsp_nb, search_date
                               
                print(true_format + ' 寫入檔案...')
                op_data.append([search_date, sfbc_nb, sfbp_nb, sfsc_nb, sfsp_nb, bc_nb, bp_nb, sc_nb, sp_nb])
                if i == get_days:
                    workbook = openpyxl.load_workbook('test.xlsx')
                    sheet = workbook.worksheets[0]
                    # sheet.append(['日期', '自營bc', '自營bp', '自營sc', '自營sp', '外資bc', '外資bp', '外資sc', '外資sp'])
                    for data in op_data:
                        sheet.append(data)                    
                    workbook.save('test.xlsx')
                    break
            except:
                print(true_format)
                print('可能是假日或沒開盤')
                if i == get_days:
                    workbook = openpyxl.load_workbook('test.xlsx')
                    sheet = workbook.worksheets[0]
                    # sheet.append(['日期', '自營bc', '自營bp', '自營sc', '自營sp', '外資bc', '外資bp', '外資sc', '外資sp'])
                    for data in op_data:
                        sheet.append(data)
                    workbook.save('test.xlsx')
                    break


    def anayisis_future(self, file_mode):
        get_days = int(input('請輸入期貨回測天數: '))
        i = 0
        fu_data = []
        while True:
            try:
                day_delta = i
                i += 1
                today_to_past = (dt.datetime.now() - timedelta(days= int(day_delta)))
                time.sleep(1)
                true_format = today_to_past.strftime("%Y/%m/%d") #change date format

                table, search_web, search_date = self.crawl(search_web='fu_three', search_date=true_format)
                sf_fucall = int(table[2].iloc[5][9]) #自營商大台
                sf_fuput = int(table[2].iloc[5][11])
                fucall = int(table[2].iloc[7][9])   #外資大台
                fuput = int(table[2].iloc[7][11])

                sf_smfucall = int(table[2].iloc[14][9]) #自營商小台
                sf_smfuput = int(table[2].iloc[14][11])
                smfucall = int(table[2].iloc[16][9])  #外資小台
                smfuput = int(table[2].iloc[16][11])

                print(true_format + ' 寫入檔案...')
                fu_data.append([search_date, sf_fucall, sf_fuput, fucall, fuput, sf_smfucall, sf_smfuput, smfucall, smfuput])

                if i == get_days:
                    with open('fu_data.csv', file_mode, encoding = 'utf-8-sig') as f:
                        # f.write('日期, 自營大台多, 自營大台空, 外資大台多, 外資大台空, 自營小台多, 自營小台空, 外資小台多, 外資小台空\n') #file mode=w 時要打開
                        for p in fu_data:
                            f.write(p[0] + ',' + str(p[1]) + ',' + str(p[2]) + ',' +str(p[3]) + ',' + str(p[4]) + ',' + str(p[5]) + ',' + str(p[6])  + ',' + str(p[7]) + ',' + str(p[8]) +'\n')
                    break
            except:
                print(true_format)
                print('可能是假日或沒開盤')
                if i == get_days:
                    with open('fu_data.csv', file_mode, encoding = 'utf-8-sig') as f:
                        # f.write('日期, 自營大台多, 自營大台空, 外資大台多, 外資大台空, 自營小台多, 自營小台空, 外資小台多, 外資小台空\n') #file mode=w 時要打開
                        for p in fu_data:
                            f.write(p[0] + ',' + str(p[1]) + ',' + str(p[2]) + ',' +str(p[3]) + ',' + str(p[4]) + ',' + str(p[5]) + ',' + str(p[6])  + ',' + str(p[7]) + ',' + str(p[8]) +'\n')
                    break

    def analysis_big10op(self, file_mode):
        get_days = int(input('請輸入十大OP回測天數: '))
        i = 0
        op_data = []
        while True:
            try:
                day_delta = i
                i += 1
                today_to_past = (dt.datetime.now() - timedelta(days= int(day_delta)))
                time.sleep(1)
                true_format = today_to_past.strftime("%Y/%m/%d")

                table, search_web, search_date = self.crawl_big10(search_web='cnp_big10', search_date=true_format)
                big10_weekbc = int((table[3].iloc[0][4].split(' ', 1)[0]).replace(',','')) # 十大週選op
                big10_weekbp = int((table[3].iloc[0][8].split(' ', 1)[0]).replace(',',''))
                big10_weeksc = int((table[3].iloc[3][4].split(' ', 1)[0]).replace(',',''))
                big10_weeksp = int((table[3].iloc[3][8].split(' ', 1)[0]).replace(',',''))
                # print(table[3])

                big10_monthbc = int((table[3].iloc[1][4].split(' ', 1)[0]).replace(',',''))# 十大月選op
                big10_monthbp = int((table[3].iloc[1][8].split(' ', 1)[0]).replace(',',''))
                big10_monthsc = int((table[3].iloc[4][4].split(' ', 1)[0]).replace(',',''))
                big10_monthsp = int((table[3].iloc[4][8].split(' ', 1)[0]).replace(',',''))
                           
                print(true_format + '寫入檔案...')
                op_data.append([search_date, big10_weekbc, big10_weekbp, big10_weeksc, big10_weeksc, big10_monthbc, big10_monthbp, big10_monthsc, big10_monthsp])

                if i == get_days:
                    with open('big10op_data.csv', file_mode, encoding = 'utf-8-sig') as f:
                        # f.write('日期, 週選bc, 週選bp, 週選sc, 週選sp, 月選bc, 月選bp, 月選sc, 月選sp\n') #file mode=w 時要打開
                        for p in op_data:
                            f.write(p[0] + ',' + str(p[1]) + ',' + str(p[2]) + ',' +str(p[3]) + ',' + str(p[4]) + ',' + str(p[5]) + ',' + str(p[6])  + ',' + str(p[7]) + ',' + str(p[8]) +'\n')
                    break
            except:
                print(true_format)
                print('可能是假日或沒開盤')
                if i == get_days:
                    with open('big10op_data.csv', file_mode, encoding = 'utf-8-sig') as f:
                        f.write('日期, 週選bc, 週選bp, 週選sc, 週選sp, 月選bc, 月選bp, 月選sc, 月選sp\n') #file mode=w 時要打開
                        for p in op_data:
                            f.write(p[0] + ',' + str(p[1]) + ',' + str(p[2]) + ',' +str(p[3]) + ',' + str(p[4]) + ',' + str(p[5]) + ',' + str(p[6])  + ',' + str(p[7]) + ',' + str(p[8]) +'\n')
                    break

    def get_op2daysdata(self):
        try :
            what_day = int(datetime.datetime.now().isoweekday())
            print('今天是星期', what_day)
            if what_day == 1:
                day_delta = 3
            else:
                day_delta = 1
            yesterday = (dt.datetime.now() - timedelta(days= int(day_delta)))
            # yesterday = (dt.datetime.now() - timedelta(days= int(input('星期一的話輸入\'3\', 其他輸入\'1\': ' ))))
            true_format = yesterday.strftime("%Y/%m/%d")
            table, search_web, search_date = self.crawl(search_web='cnp_three', search_date=true_format)
            print('Connecting yesterday opweb..')
            sfbc_nb = int(table[2].iloc[5][10]) #自營商 BC BP SC SP
            sfbp_nb = int(table[2].iloc[8][10])
            sfsc_nb = int(table[2].iloc[5][12])
            sfsp_nb = int(table[2].iloc[8][12])

            bc_nb = int(table[2].iloc[7][10]) #外資 BC BP SC SP
            bp_nb = int(table[2].iloc[10][10])
            sc_nb = int(table[2].iloc[7][12])
            sp_nb = int(table[2].iloc[10][12])
            yesterday_oplist = [sfbc_nb, sfbp_nb, sfsc_nb, sfsp_nb, bc_nb, bp_nb, sc_nb, sp_nb]
            print(yesterday_oplist)
        except:
            print('可能為假日或沒開盤...')

        try:
            today = dt.datetime.now().strftime("%Y/%m/%d")
            table, search_web, search_date = self.crawl(search_web='cnp_three', search_date=today)
            print('Connecting today opweb..')
            sfbc_nb = int(table[2].iloc[5][10]) #自營商 BC BP SC SP
            sfbp_nb = int(table[2].iloc[8][10])
            sfsc_nb = int(table[2].iloc[5][12])
            sfsp_nb = int(table[2].iloc[8][12])

            bc_nb = int(table[2].iloc[7][10]) #外資 BC BP SC SP
            bp_nb = int(table[2].iloc[10][10])
            sc_nb = int(table[2].iloc[7][12])
            sp_nb = int(table[2].iloc[10][12])
            today_oplist = [sfbc_nb, sfbp_nb, sfsc_nb, sfsp_nb, bc_nb, bp_nb, sc_nb, sp_nb]
            print(today_oplist)
        except:
            print('還沒下午三點讀取不到今日資料...')

        return today_oplist, yesterday_oplist

    def analysis_opdata(self):
        try:
            today_oplist, yesterday_oplist = self.get_op2daysdata()
            print('Getting yesterday&today op data...')
            op_dif = [today_oplist[i] - yesterday_oplist[i] for i in range(len(today_oplist))]
            header = ['sfbc_nb', 'sfbp_nb', 'sfsc_nb', 'sfsp_nb', 'bc_nb', 'bp_nb', 'sc_nb', 'sp_nb']
            
            print('今日op資料: ', today_oplist)
            print('昨日op資料: ', yesterday_oplist)
        except:
            print('Fail to analysis')
        return header, op_dif

    def get_fu2daysdata(self):
        try :
            what_day = int(datetime.datetime.now().isoweekday())
            print('今天是星期', what_day)
            if what_day == 1:
                day_delta = 3
            else:
                day_delta = 1
            yesterday = (dt.datetime.now() - timedelta(days= int(day_delta)))
            # yesterday = (dt.datetime.now() - timedelta(days= int(input('星期一的話輸入\'3\', 其他輸入\'1\': ' ))))
            true_format = yesterday.strftime("%Y/%m/%d")
            table, search_web, search_date = self.crawl(search_web='fu_three', search_date=true_format)
            print('Connecting yesterday fuweb..')
            sf_fucall = int(table[2].iloc[5][9]) #自營商大台
            sf_fuput = int(table[2].iloc[5][11])
            fucall = int(table[2].iloc[7][9])   #外資大台
            fuput = int(table[2].iloc[7][11])

            sf_smfucall = int(table[2].iloc[14][9]) #自營商小台
            sf_smfuput = int(table[2].iloc[14][11])
            smfucall = int(table[2].iloc[16][9])  #外資小台
            smfuput = int(table[2].iloc[16][11])
            yesterday_fulist = [sf_fucall, sf_fuput, fucall, fuput, sf_smfucall, sf_smfuput, smfucall, smfuput]
            print(yesterday_fulist)
        except:
            print('可能為假日或沒開盤...')

        try:
            today = dt.datetime.now().strftime("%Y/%m/%d")
            table, search_web, search_date = self.crawl(search_web='fu_three', search_date=today)
            print('Connecting today opweb..')
            sf_fucall = int(table[2].iloc[5][9]) #自營商大台
            sf_fuput = int(table[2].iloc[5][11])
            fucall = int(table[2].iloc[7][9])   #外資大台
            fuput = int(table[2].iloc[7][11])

            sf_smfucall = int(table[2].iloc[14][9]) #自營商小台
            sf_smfuput = int(table[2].iloc[14][11])
            smfucall = int(table[2].iloc[16][9])  #外資小台
            smfuput = int(table[2].iloc[16][11])
            today_fulist = [sf_fucall, sf_fuput, fucall, fuput, sf_smfucall, sf_smfuput, smfucall, smfuput]
            print(today_fulist)
        except:
            print('還沒下午三點讀取不到今日資料...')

        return today_fulist, yesterday_fulist

    def analysis_fudata(self):
        try:
            today_fulist, yesterday_fulist = self.get_fu2daysdata()
            print('Getting yesterday&today fu data...')
            fu_dif = [today_fulist[i] - yesterday_fulist[i] for i in range(len(today_fulist))]
            fu_header = ['sf_fucall', 'sf_fuput', 'fucall', 'fuput', 'sf_smfucall', 'sf_smfuput', 'smfucall', 'smfuput']
            
            print('今日fu資料: ', today_fulist)
            print('昨日fu資料: ', yesterday_fulist)
        except:
            print('Fail to analysis')
        return fu_header, fu_dif

    def autolabel(self, rects):
        """Attach a text label above each bar in *rects*, displaying its height."""
        for rect in rects:
            height = rect.get_height()
            plt.annotate('{}'.format(height),
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 3),  # 3 points vertical offset
                        textcoords="offset points",
                        ha='center', va='bottom')

    def data_visualization(self, x_range, y_range, title):
            rects = plt.bar(x_range, y_range, color=['firebrick','g'])
            plt.title(title)
            self.autolabel(rects) #remark content
            # plt.close('all')
            # path = 'C:\\Users\\RF\\Desktop\\coding\\op\\' + today +'\\'
            # if not os.path.isdir(path):
            #     os.mkdir(path)
            # plt.savefig( path + today + save_name, dpi=200)
            # plt.show()
            
    def data_print(self):
        header, op_dif = self.analysis_opdata()
        self.data_visualization(header[0:2], op_dif[0:2], '自營bcbp')
        self.data_visualization(header[2:4], op_dif[2:4], '自營scsp')
        self.data_visualization(header[4:6], op_dif[4:6], '外資bcbp')
        self.data_visualization(header[6:8], op_dif[6:8], '外資scsp')
        fu_header, fu_dif = self.analysis_fudata()
        self.data_visualization(fu_header[0:2], fu_dif[0:2], '自營大台')
        self.data_visualization(fu_header[2:4], fu_dif[2:4], '外資大台')
        self.data_visualization(fu_header[4:6], fu_dif[4:6], '自營小台')
        self.data_visualization(fu_header[6:8], fu_dif[6:8], '外資小台')


        
op = Op()
if __name__ == '__main__':
   
    while True:
        q = input('執行資料抓取輸入\'1\', \n執行今日昨日資料比對輸入\'2\', \n, 執行excel 測試輸入3, \n離開程式輸入\'q\': ')
        if q == '1':
            print('執行資料抓取...')
            op.analysis_op(file_mode='a')
            op.anayisis_future(file_mode='a')
            op.analysis_big10op(file_mode='a')
            
        elif q == '2':
             print('執行今日昨日資料比對...')
             op.data_print()
        elif q == '3':
            print('Excel test...')
            op.analysis_op_excel()
             
        elif q == 'q':
            break
    # op.get_2daysdata()
    # op.analysis_data()
    
    
    