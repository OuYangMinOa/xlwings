import openpyxl
import xlwings as xw
import pandas as pd
FILE   = 'data.xlsx'
attack_ee = {'87':1629,'88':1641,'89':1652,'90':1664,'91':1676,'92':1688,
             '93':1700,'94':1712,'95':1724,'96':1736,'97':1748,'98':1760,'99':1773}
all_t  = [1629,1641,1652,1664,1676,1688,1689,1700,1712,1724,1736,1748,1760,1773]
#########################
wb     = xw.Book(FILE)
sheet  = wb.sheets[0]
sht    = wb.sheets.active
def fegr():
    write_file = open('get_data.txt','w',encoding="utf-8")
    try:
        for card_num in range(1,1987):
            # 未昇華
            sht.range('A4').value = card_num #編號
            sht.range('B4').value = 99       #滿等    
            sht.range('C4').value = 4        #未昇華       
            big_level     = sht.range('D4').value #最大等級
            card_name     = sht.range('L4').value
            attack        = sht.range('N4').value
            attack_2      = sht.range('AC4').value
            
            try:
                int(attack)
                int(attack_2)
                int(big_level)
            except:
                continue
            
            if (int(attack) < min(all_t) ):
                pass
            elif(int(attack) in all_t):
                print(card_name,"(未昇華)最大攻擊力:",int(attack))
                write_file.write('{} 等級99(未昇華)#{}\n'.format(card_name,int(attack)))
            else:
                for level in range(1,int(big_level)):
                    sht.range('B4').value = level
                    attack = sht.range('N4').value
                    if(int(attack) in all_t):
                        print(card_name,"(未昇華)等級",level,"攻擊力:",int(attack))
                        write_file.write('{} 等級{}(未昇華)#{}\n'.format(card_name,level,int(attack)))
                    elif(int(attack) > 1773):
                        break
            #昇華
            if (attack_2 == attack):
                continue
            
            if (int(attack_2) < min(all_t) ):
                continue
            elif(int(attack_2) in all_t):
                print(card_name,"(有昇華)最大攻擊力:",int(attack_2))
                write_file.write('{} 等級99(有昇華)#{}\n'.format(card_name,int(attack_2)))
            else:
                for level in range(1,int(big_level)):
                    sht.range('B4').value = level
                    if(int(attack_2) in all_t):
                        print(card_name,"(有昇華)等級",level,"攻擊力:",int(attack_2))
                        write_file.write('{} 等級{}(有昇華)#{}\n'.format(card_name,level,int(attack_2)))
                    elif(int(attack_2) > 1773):
                        break
    except Exception as e:
        write_file.close()
        print(e)
    write_file.close()
    now_to = open('進度.txt','w', encoding="utf-8")
    now_to.write(str(card_num))
    now_to.close()
#fegr()
write_file = open('get_data.txt','r',encoding="utf-8")
get = write_file.read().split('\n')
for t_level in attack_ee:
    print('\n三藏等級',t_level,'\n')
    try:
        for j in get:
            go = j.split('#')
            if (int(go[1]) == int(attack_ee[t_level])):
                if ('‧' not in go[0].split("等級")[0]):
                    print('{:　<12}等級{}'.format(go[0].split("等級")[0],go[0].split("等級")[1]))
                else:
                    print('{:　<13}等級{}'.format(go[0].split("等級")[0],go[0].split("等級")[1]))
        print() 
    except:
        pass

print("總共",len(get),"個結果")









            
                
