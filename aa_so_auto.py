import pyautogui as pg
import time
import pyperclip as pc
import pandas as pd
import numpy as np
from datetime import datetime
import webbrowser

start_time = time.time()

# =============================================================================
# Program run time depends on the number of ticker every 100 tickers take around 330 seconds/ 5 mins to run
# =============================================================================

# Open AAstock specific page for SO extraction
webbrowser.open('http://www.aastocks.com/en/stocks/analysis/company-fundamental/basic-information?symbol=00700')  # Go to example.com

# Failsafe
pg.FAILSAFE = True

#pg.position()

# Top 200 market value ticker for HK market (i.e. fundamental ticker is HK, e.g. 9988 HK has BABA US as fundamental ticker so won't be included in this list)
stock_list = [
700,939,2318,1288,941,3988,857,1299,2628,3690,386,1658,3328,883,914,1088,2202,388,16,788,1876,688,1339,66,11,6066,
2388,1928,998,1810,241,3,960,1109,27,1,267,728,2007,2,3333,6862,1113,3692,2020,1211,2328,1918,2269,762,390,12,384,1816,1177,1186,
6186,823,2313,1800,291,1972,1038,6808,2382,2338,20,669,1833,175,6,288,2319,813,270,2899,2238,2688,1997,6098,6823,1093,1193,2016,17,6881,656,2800,
1044,981,3323,3993,322,83,1313,881,101,2638,1055,670,1128,2333,19,151,3969,3799,1579,1929,1099,902,586,1913,3380,2066,817,1169,2331,6110,3618,
1359,177,3319,874,1066,1821,2883,3908,992,708,1801,247,1898,2833,1030,966,135,884,2018,1919,6049,3808,489,880,836,4,486,3918,285,968,358,
696,1138,1171,6198,552,6060,1508,268,868,8,467,3888,293,1114,991,2282,6199,53,1813,6185,1551,2009,2588,772,3383,3698,392,220,2689,1528,1797,2128,
2777,522,144,2799,152,1877,1378,345,659,683,853,3311,1691,916,3320,1619,371]

# Market cap 201-500
add300 = [
2869,3898,1818,2669,667,1958,780,839,1216,6190,257,3331,2778,1548,6169,14,1060,576,3360,3866,3996,1252,1233,1717,2696,
694,6865,1951,1888,3883,3990,69,3998,867,598,123,425,2500,1789,10,1882,6196,604,200,316,551,6158,2828,1916,148,1112,1308,1578,3396,2866,
2772,1530,1347,165,754,2314,636,493,6869,9926,1238,535,2880,2357,9969,2823,3900,1310,973,3633,416,1638,1628,512,570,1995,2005,9922,1999,
43,410,1268,41,9966,2386,1448,81,933,1375,3369,2380,3868,2103,1243,3377,1053,1755,1858,303,1765,1966,363,778,1606,6138,743,2013,3669,
6826,921,1558,3301,2768,56,3308,1622,1908,1963,1199,1686,631,811,3933,1228,1905,2186,777,2068,405,8083,846,1896,6088,816,45,95,9968,179,354,
6122,1478,579,3339,173,107,1610,272,2299,1316,832,665,1083,34,1212,6055,1883,2356,530,1208,1839,494,1458,1387,1111,376,697,71,590,341,1333,
1890,2048,855,1788,119,775,1589,2098,1337,1330,336,588,1910,1585,691,2558,1052,1907,995,737,6100,1996,3613,1098,1727,1031,207,581,1992,1381,1608,
1660,2139,1829,639,658,412,819,1777,2400,506,3813,520,1383,59,242,2019,1282,1769,934,1773,440,2342,62,1076,2858,2606,1636,337,1176,2666,2616,9909,1070,
2006,1666,1089,460,1568,1521,2233,2868,1476,1873,797,1368,2038,189,861,1475,1302,67,3899,1675,1224,1817,142,1983,373,127,256,1836,6855,6068,1811,435,1196,
2362,35,1257,956,1317,369,215,198,546,1860,2001,1501,86,9928,978,826,799,6868,3836,1357,3877,51,302,462,308]

# Check for non-equity ticker
def not_equity(listl):
    for i in listl:
        if i >9999:
            print(i, " is not a common stock code, please delete it from the list")
        else:
            pass

not_eq_map = map(not_equity, [stock_list, add300])
list_not_eq = list(not_eq_map)
for i in range(len(list_not_eq)):
    print(list_not_eq[i])
    
# error_list = []

### functions for return button coordinatein number
    
def findloc_tic(picname): # return position to enter ticker code
    pic = pg.locateOnScreen(str(picname) + '.png')
    pic_mid = pg.center(pic)
    pic_midx, pic_midy = pic_mid
    adj_x = pic_midx/2
    adj_y = pic_midy/2
    tic_x = adj_x
    tic_y = adj_y+56
    return tic_x, tic_y

ticx, ticy = findloc_tic("DQL") # DQL is the picture above the stock code input position

def findloc_SO(picname): # return position for the SO to be copied
    pic = pg.locateOnScreen(str(picname) + '.png')
    pic_mid = pg.center(pic)
    pic_midx, pic_midy = pic_mid
    adj_x = pic_midx/2
    adj_y = pic_midy/2
    SO_x = adj_x +195
    SO_y = adj_y
    return SO_x, SO_y

SOx, SOy = findloc_SO("SO") # SO is the picture at the right of the issued Captial number
    
###

# Auto copy part

so_list =[]
previous_SO = 1    # dummy for checking SO i and SO i-1, should be different
for i,j in enumerate(stock_list):
    pg.moveTo(ticx,ticy) # input stock location
    pg.click()
    pg.doubleClick()
    pg.typewrite(str(j))
    pg.hotkey('enter') 
    time.sleep(1.5) # wait for page load
    pg.moveTo(SOx,SOy) # SO position
    pg.click()
    pg.doubleClick()
    pg.hotkey('command', 'c', interval = 0.15)
    text = pc.paste()  # text will have the content of clipboard
    while text[-1].isdigit() == True and text[0].isdigit() == True and int(text.replace(",","")) != previous_SO: # Check on the copy is the next SO
        break
    else:
        webbrowser.open('http://www.aastocks.com/en/stocks/analysis/company-fundamental/basic-information?symbol=00700')  # Go to example.com
        time.sleep(1.5)
        pg.moveTo(ticx,ticy) # input stock location
        pg.click()
        pg.doubleClick()
        pg.typewrite(str(j))
        pg.hotkey('enter') 
        time.sleep(1.5) # wait for page load
        pg.moveTo(SOx,SOy) # SO position
        pg.click()
        pg.doubleClick()
        pg.hotkey('command', 'c', interval = 0.15)
        text = pc.paste()  # text will have the content of clipboard
        
    text1 = int(text.replace(",",""))

    while text1 > 10000: # Check it is copying the SO and no mistakenly copy other value (such as the ticker code or else)
        break
    else:
        time.sleep(0.1)
        pg.moveTo(SOx,SOy) # SO position
        pg.click()
        pg.doubleClick()
        pg.hotkey('command', 'c', interval = 0.15)
        text = pc.paste()  # text will have the content of clipboard
        text1 = int(text.replace(",",""))
    
    so_list.append(text1)
    previous_SO = text1

# Ticker dataframe
res=[]
for i in stock_list:
    equity = str(i) + " HK Equity"
    res.append(equity)    
stock_list_pd = pd.DataFrame(res)

# so dataframe
so_excel_pd = pd.DataFrame(so_list)

# Bloomberg API dataframe
res1=[]
for i in res:
    API = "=BDP(\"" + str(i) + "\", " "\"EQY_SH_OUT_REAL\")"
    res1.append(API)

API_pd = pd.DataFrame(res1)

# Excel filter function for SO different between BBG and AA stock > 1
res2=[]
for i in stock_list:
    filter_1 = "=if(ABS(C2-D2)>1 ,1 ,0)"
    res2.append(filter_1)
    
filter_pd = pd.DataFrame(res2)

# Pull fundamental ticker with BBG API
res3=[]
for i in stock_list:
    fund = "=RIGHT(BDP(B2, \"EQY_FUND_TICKER\") ,2)"
    res3.append(fund)
    
fund_pd = pd.DataFrame(res3)

# Combine list
frame = [stock_list_pd, so_excel_pd, API_pd, filter_pd, fund_pd]
result_excel = pd.concat(frame, axis = 1, sort = False)
result_excel.columns = ["Ticker","AA SO", "BBG SO", "SO Diff", "Fundamental Ticker" ]

# Export result
today = datetime.today().strftime('%Y-%m-%d')
pd.DataFrame(result_excel).to_excel("AA_SO " + str(today) +".xlsx")

so_check = result_excek.iloc[:,2].tolist()
so_check.sort()
for i in range(len(so_check)-1):
    if so_check[i] == so_check[i+1]:
        print (so_check[i] + " is duplicate")

print(".")
print(".")
print(".")
print("AAstock SO autocopy completed and exported excel file")
print("--- runtime: %s seconds ---" % (time.time() - start_time))

