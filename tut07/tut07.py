import pandas as pd
from platform import python_version
import math
import os
import glob
from datetime import datetime
start_time = datetime.now()
import xlsxwriter
import threading

Graph_ = {0:"+1 " , 1:"-1 "  , 2:"+2 "  , 3:"-2 "  , 4:"+3 "  , 5:"-3 "  , 6:"+4 "  , 7:"-4 "}


def octant_longest_subsequence_count(f):

    number = len(f)
    
    f.loc[0, "U_Avg"] = f["U"].mean()
    f.loc[0, "V_Avg"] = f["V"].mean()
    f.loc[0, "W_Avg"] = f["W"].mean()
    u_mean = f["U"].mean()
    v_mean = f["V"].mean()
    w_mean = f["W"].mean()
    f["u_"] = f["U"]-u_mean
    f["v_"] = f["V"]-v_mean
    f["w_"] = f["W"]-w_mean

    
    f.loc[((f.u_ > 0) & (f.v_ > 0) & (f.w_ > 0)), "Octant"] = "+1"
    f.loc[((f.u_ > 0) & (f.v_ > 0) & (f.w_ < 0)), "Octant"] = "-1"
    f.loc[((f.u_ < 0) & (f.v_ > 0) & (f.w_ > 0)), "Octant"] = "+2"
    f.loc[((f.u_ < 0) & (f.v_ > 0) & (f.w_ < 0)), "Octant"] = "-2"
    f.loc[((f.u_ < 0) & (f.v_ < 0) & (f.w_ > 0)), "Octant"] = "+3"
    f.loc[((f.u_ < 0) & (f.v_ < 0) & (f.w_ < 0)), "Octant"] = "-3"
    f.loc[((f.u_ > 0) & (f.v_ < 0) & (f.w_ > 0)), "Octant"] = "+4"
    f.loc[((f.u_ > 0) & (f.v_ < 0) & (f.w_ < 0)), "Octant"] = "-4"

    Longsum = {"-1": 0, "+1": 0,  "-2": 0, "+2": 0,
               "-3": 0, "+3": 0, "-4": 0, "+4": 0}





    f["      "] = " "
    f["count"] = ""
    f["Longest Subsquence Length"] = ""
    f["Count"] = ""


    f["       "] = " " 
    f["count2"] = ""
    f["Longest Subsquence Length2"] = ""
    f["Count2"] = ""



    k = 0

    while k < number:
        octant1 = f["Octant"][k]
        curr_count = 0
        l = k
        while l < number:
            try:
                octant2 = f["Octant"][l]
                if (octant2 == octant1):
                    curr_count += 1
                else:
                    break
                l += 1
            except:
                print("row not found")
        Longsum[f["Octant"][k]] = max(Longsum[f["Octant"][k]], curr_count)
        k = l

    mapp = {0: "+1", 1: "-1",  2: "+2", 3: "-2",
               4: "+3", 5: "-3", 6: "+4", 7: "-4"}
    for k in range(8):
        try:
            f.loc[k, "count"] = mapp[k]
        except:
            print("row not found")

    for k in range(8):
        try:
            f.loc[k, "Longest Subsquence Length"] = Longsum[mapp[k]]
        except:
            print("row not found")

    k = 0

   
    countLongsum = {"-1": 0, "+1": 0,  "-2": 0,
                  "+2": 0, "-3": 0, "+3": 0, "-4": 0, "+4": 0}
    
    ranLonsum = {"-1": [], "+1": [],  "-2": [],
                    "+2": [], "-3": [], "+3": [], "-4": [], "+4": []}

    while k < number:
        curr_count = 0
        octant1 = ""
        try:
            l = k
            octant1 = f["Octant"][k]
            while l < number:

                octant2 = f["Octant"][l]
                if (octant2 == octant1):
                    curr_count += 1
                else:
                    break
                l += 1
                
            if (curr_count == Longsum[octant1]):
                countLongsum[octant1] += 1
               
                try:
                    lirange = str(f["T"][k]) + "," + str(f["T"][l-1])
                    ranLonsum[octant1].append(lirange)
                except:
                    print("Invalid mod value or invalid data or Excel file is empty")

            k = l
        except:
            print("row not found")
     
 
        
    ind2 = 0
    for ind1 in range(8):
        f.loc[ind1, "Count"] = countLongsum[mapp[ind1]]
        currOctant = mapp[ind1]

        f.loc[ind2, "count2"] = currOctant
        f.loc[ind2, "Longest Subsquence Length2"] = Longsum[currOctant]
        f.loc[ind2, "Count2"] = countLongsum[currOctant]

        ind2 += 1
        f.loc[ind2, "count2"] = "Time"
        f.loc[ind2, "Longest Subsquence Length2"] = "From"
        f.loc[ind2, "Count2"] = "To"

        ind2 += 1
        for TRange in ranLonsum[currOctant]:
            lst = TRange.split(",")
            f.loc[ind2, "Longest Subsquence Length2"] = str(lst[0])
            f.loc[ind2, "Count2"] = str(lst[1])
            ind2 += 1


def octant_range_names(f, mod=5000):
    nu = len(f)
    
    octant_name_id_mapping = {"+1 ": "Internal outward interaction", "-1 ": "External outward interaction", "+2 ": "External Ejection",
                              "-2 ": "Internal Ejection", "+3 ": "External inward interaction", "-3 ": "Internal inward interaction", "+4 ": "Internal sweep", "-4 ": "External sweep"}
    f.loc[0, "U_Avg"] = f["U"].mean()
    f.loc[0, "V_Avg"] = f["V"].mean()
    f.loc[0, "W_Avg"] = f["W"].mean()
    u_mean = f["U"].mean()
    v_mean = f["V"].mean()
    w_mean = f["W"].mean()
    f["u_"] = f["U"]-u_mean
    f["v_"] = f["V"]-v_mean
    f["w_"] = f["W"]-w_mean

    # creating  a octant column 
    f.loc[((f.u_ > 0) & (f.v_ > 0) & (f.w_ > 0)), "Octant"] = "+1"
    f.loc[((f.u_ > 0) & (f.v_ > 0) & (f.w_ < 0)), "Octant"] = "-1"
    f.loc[((f.u_ < 0) & (f.v_ > 0) & (f.w_ > 0)), "Octant"] = "+2"
    f.loc[((f.u_ < 0) & (f.v_ > 0) & (f.w_ < 0)), "Octant"] = "-2"
    f.loc[((f.u_ < 0) & (f.v_ < 0) & (f.w_ > 0)), "Octant"] = "+3"
    f.loc[((f.u_ < 0) & (f.v_ < 0) & (f.w_ < 0)), "Octant"] = "-3"
    f.loc[((f.u_ > 0) & (f.v_ < 0) & (f.w_ > 0)), "Octant"] = "+4"
    f.loc[((f.u_ > 0) & (f.v_ < 0) & (f.w_ < 0)), "Octant"] = "-4"
    f.loc[1, " "] = "userinput"

    # finding the total count 
    h = f['Octant'].value_counts()
   

    f.loc[1, "Octant Id"] = "overall Count"
    f.loc[2, "Octant Id"] = "mod"+" " + str(mod)

    num_blocks = math.ceil(nu/mod)
    l = 0
    m = mod
    st = 0
    end = m-1
    j = 0
    for j in range(num_blocks):
        if (st + mod > nu):
            f.loc[j+3, "Octant Id"] = str(st)+"-" + str(nu-1)
            break
        else:
            f.loc[j+3, "Octant Id"] = str(st)+"-" + str(end)

        j = j+1
        st = st + mod
        end = end + mod

    try:
        f.loc[1, "+1"] = h["+1"]  # creating a column 
    except:
        f.loc[1, "+1"] = 0
    try:
        f.loc[1, "-1"] = h["-1"]   
    except:
        f.loc[1, "-1"] = 0
    try:
        f.loc[1, "+2"] = h["+2"]   
    except:
        f.loc[1, "+2"] = 0
    try:
        f.loc[1, "-2"] = h["-2"]  
    except:
        f.loc[1, "-2"] = 0
    try:
        f.loc[1, "+3"] = h["+3"]  
    except:
        f.loc[1, "+3"] = 0
    try:
        f.loc[1, "-3"] = h["-3"]  
    except:
        f.loc[1, "-3"] = 0
    try:
        f.loc[1, "+4"] = h["+4"]  
    except:
        f.loc[1, "+4"] = 0
    try:
        f.loc[1, "-4"] = h["-4"]
    except:
        f.loc[1, "-4"] = 0
 


    # creating rank columns 
    f.loc[0 , "+1 "] = "Rank1"  #Rank of Octand Id1
    f.loc[0 , "-1 "] = "Rank2"
    f.loc[0 , "+2 "] = "Rank3"
    f.loc[0 , "-2 "] = "Rank4"
    f.loc[0 , "+3 "] = "Rank5"
    f.loc[0 , "-3 "] = "Rank6"
    f.loc[0 , "+4 "] = "Rank7"
    f.loc[0 , "-4 "] = "Rank8"
    f.loc[0, "   "] = "Rank1 Octant Id"
    f.loc[0, "    "] = "Rank1 Octant Name"
    
    rank1_cnt = {"+1 " : 0, "-1 " : 0, "+2 ": 0, "-2 ": 0, "+3 ": 0, "-3 ": 0, "+4 ": 0, "-4 ": 0 }
    r = 0 
    row = 3
    j = 0
    oct1 = oct2 = oct3 = oct4 = oct5 = oct6 = oct7 = oct8 = 0
    for j in range(num_blocks):
        
        for i in range(mod):
            try:
                if f["Octant"][r] == "+1":
                    oct1 = oct1+1
                elif f["Octant"][r] == "-1":
                    oct2 += 1
                elif f["Octant"][r] == "+2":
                    oct3 += 1
                elif f["Octant"][r] == "-2":
                    oct4 += 1
                elif f["Octant"][r] == "+3":
                    oct5 += 1
                elif f["Octant"][r] == "-3":
                    oct6 += 1
                elif f["Octant"][r] == "+4":
                    oct7 += 1
                elif f["Octant"][r] == "-4":
                    oct8 += 1
            except:
                print("Row Not found")

            r = r+1
            if r == nu:
                j = num_blocks+1
                break  
        j = j+1
        try:
            f.loc[row, "+1"] = oct1
            f.loc[row, "+2"] = oct3
            f.loc[row, "-2"] = oct4
            f.loc[row, "+3"] = oct5
            f.loc[row, "-3"] = oct6
            f.loc[row, "+4"] = oct7
            f.loc[row, "-4"] = oct8
            f.loc[row, "-1"] = oct2
        except:
            print("Index out of Bound Error")
            
        ls = [ oct1, oct2 , oct3 , oct4 , oct5 , oct6, oct7 , oct8]
        st = set({})
        dict = {}
        
        for el in ls:
            st.add(el)
        for el in st:
            dict[el] = []
        i = 0
        for el in ls:
            dict[el].append(Graph_[i])
            i+= 1
        for el in st:
            dict[el].sort(reverse = True)
        
        rank1 = "Initiate"
        rank = 8
        for el in st:
            for octant in dict[el]:
                try:
                   f.loc[row, octant] = rank
                except:
                   print("Index out of Bound Error") 
                if(rank == 1):
                    rank1 = octant
                rank -=  1 
        f.loc[row, "   "] = rank1
        f.loc[row, "    "] = octant_name_id_mapping[rank1]
        rank1_cnt[rank1] += 1
        oct1 = oct2 = oct3 = oct4 = oct5 = oct6 = oct7 = oct8 = 0
        
        row = row+1
    row += 4
    f.loc[row, "+1"] = "Octant Id"
    f.loc[row , "-1"] = "Octant Name"
    f.loc[row , "+2"] = "Count of Rank 1 Mod Values"
    
    row += 1
    cnt = 0
    for key in octant_name_id_mapping:
        f.loc[row, "+1"] = Graph_[cnt]
        f.loc[row, "-1"] = octant_name_id_mapping[Graph_[cnt]]
        f.loc[row, "+2"] = rank1_cnt[Graph_[cnt]]
        row += 1
        cnt += 1
def Iintial(f, n, i):
    num = len(f)
    try:
       string1 = f['Octant'][i] + "  "    
       for i in range(8):
            f.iat[n+i, f.columns.get_loc(string1)] = 0
    except:
        print("Index out of Bound Error1")

def upda(f, l, i):
    num = len(f)
    try:
        string1 = f['Octant'][i]
        string2 = f['Octant'][i+1] + "  "

        if string1 == '+1':
            f.iat[l, f.columns.get_loc(string2)] += 1
        elif string1 == '-1':
            f.iat[l+1, f.columns.get_loc(string2)] += 1
        elif string1 == '+2':
            f.iat[l+2, f.columns.get_loc(string2)] += 1
        elif string1 == '-2':
            f.iat[l+3, f.columns.get_loc(string2)] += 1
        elif string1 == '+3':
            f.iat[l+4, f.columns.get_loc(string2)] += 1
        elif string1 == '-3':
            f.iat[l+5, f.columns.get_loc(string2)] += 1
        elif string1 == '+4':
            f.iat[l+6, f.columns.get_loc(string2)] += 1
        elif string1 == '-4':
            f.iat[l+7, f.columns.get_loc(string2)] += 1
    
    except:
        pass
        
        
def fun1(f, m):
    num = len(f)
    f.loc[m+1, "     "] = "From"  
    f.loc[m-1, "+1  "] = "To"
    f.loc[m+1, 'overall id'] = "+"+str(1)
    f.loc[m+2, 'overall id'] = '-1'
    f.loc[m+3, 'overall id'] = "+"+str(2)
    f.loc[m+4, 'overall id'] = '-2'
    f.loc[m+5, 'overall id'] = "+"+str(3)
    f.loc[m+6, 'overall id'] = '-3'
    f.loc[m+7, 'overall id'] = '+4'
    f.loc[m, "overall id"] = "Count"
    f.loc[m, "+1  "] = "+"+str(1)
    f.loc[m, "-1  "] = '-1'
    f.loc[m, "+2  "] = "+"+str(2)
    f.loc[m, "-2  "] = '-2'
    f.loc[m, "+3  "] = "+"+str(3)
    f.loc[m, "-3  "] = '-3'
    f.loc[m, "+4  "] = "+"+str(4)
    f.loc[m, "-4  "] = '-4'
    f.loc[m+8, 'overall id'] = "+"+str(4)
def octant_transition_count(f, mod=5000):
    num = len(f)
    num_blocks = math.ceil(num/mod)
    l = 0
    m = mod
    start = 0
    end = m-1
    j = 0
    for j in range(num_blocks):
        if (start + mod > num):
            break
        else:
            pass

        j = j+1
        start = start + mod
        end = end + mod

    r = 0
    row = 3
    j = 0
    m = n = y = j

    l = 0000
    z = mod-1  # assigning mod value to m
    a = 0
    b = mod-1

    f.loc[0, "     "] = ""
    for x in range(num_blocks):
        if x == 0:
            f.loc[y+2, "overall id"] = "Overall Transition Count"
        else:
            y += 12
            # changing the column overall id 
            try:
                f.loc[y+1, "overall id"] = "Mod transition count "
                f.loc[y+2, "overall id"] = str(a)+"-"+str(b)
            except:
                 print("Row Not Found")
            j = j+1
            l = z+1
            z = z+mod
            a = str(l)

            b = str(z)

    y += 12
    try:
        f.loc[y+1, "overall id"] = "Mod transition count "
        f.loc[y+2, "overall id"] = str(a)+"-"+str(num)
    except:
        print("Row Not Found")

    m += 3

    fun1(f, m)

    n = n+4
    l = n
    
    for i in range(num):
        Iintial(f, n, i)

    for i in range(num):
        upda(f, l, i)

    n += 12
    z = n

    for x in range(0, num, mod):
        for i in range(x, mod+x-1, 1):
            if (i >= num):
                break
            Iintial(f, n, i)

        n += 12

    for x in range(0, num, mod):
        for i in range(x, mod+x-1, 1):
            if (i >= num):
                break
            l = z
            upda(f, l, i)
        z += 12

    m += 12

    for x in range(num_blocks):
        fun1(f, m)
        m += 12
    
      
def octant_analysis(mod=5000):
    try:
      os.mkdir('output')
    except:
      print("Try deleting the Octant directory, and then run the program")
    path = os.getcwd()
    myfiles = glob.glob(os.path.join(path,"input", "*.xlsx"))
    
    for df in myfiles:
         # df is the path name 
        f = pd.read_excel(df)
        fname = os.path.basename(df)
        file_name  = ""
        try:
            i = 0
            sz =len(fname)
            while i < sz-5:
                file_name += fname[i]
                i+= 1
        except:
            print("Program is strictly written for .xlsx extension files ")
        file_name += " cm_vel_octant_analysis_" + str(mod)
        
        trd1 = threading.Thread(target = octant_range_names, args=(f, mod,))
        trd1.start()
        trd1.join()
        trd2 = threading.Thread(target = octant_transition_count , args=(f, mod,))
        trd2.start()
        trd2.join()
        trd3 = threading.Thread(target = octant_longest_subsequence_count, args =(f,))
        trd3.start()
        trd3.join()
        
        
        writer = pd.ExcelWriter(f'output/{file_name}.xlsx', engine ='xlsxwriter')
        
        f.to_excel(writer, sheet_name ='Sheet1')
        workbook = writer.book
        worksheet  = writer.sheets['Sheet1']
        
        border_fmt = workbook.add_format({'bottom':5, 'top':5 , 'left':5 , 'right':5})
        worksheet.conditional_format(xlsxwriter.utility.xl_range(0,0, len(f), len(f.columns)), {'type': 'no_blanks', 'format':border_fmt})
        writer.save()
        

from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod=5000
octant_analysis(mod)

#This shall be the last lines of the code.
end_T = datetime.now()
print('Duration of Program Execution: {}'.format(end_T - start_time))


# https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html   follow this link for more 
  