from platform import python_version
import math
import pandas as pd
from datetime import datetime
start_time = datetime.now()

try:
    f = pd.read_excel("octant_input.xlsx")
except:
    print("File opening error")

numb = len(f)
print(numb)

graph = {0:"+1 " , 1:"-1 "  , 2:"+2 "  , 3:"-2 "  , 4:"+3 "  , 5:"-3 "  , 6:"+4 "  , 7:"-4 "}

def octant_range_names(mod=5000):
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

    
    f.loc[((f.u_ > 0) & (f.v_ > 0) & (f.w_ > 0)), "Octant"] = "+1"
    f.loc[((f.u_ > 0) & (f.v_ > 0) & (f.w_ < 0)), "Octant"] = "-1"
    f.loc[((f.u_ < 0) & (f.v_ > 0) & (f.w_ > 0)), "Octant"] = "+2"
    f.loc[((f.u_ < 0) & (f.v_ > 0) & (f.w_ < 0)), "Octant"] = "-2"
    f.loc[((f.u_ < 0) & (f.v_ < 0) & (f.w_ > 0)), "Octant"] = "+3"
    f.loc[((f.u_ < 0) & (f.v_ < 0) & (f.w_ < 0)), "Octant"] = "-3"
    f.loc[((f.u_ > 0) & (f.v_ < 0) & (f.w_ > 0)), "Octant"] = "+4"
    f.loc[((f.u_ > 0) & (f.v_ < 0) & (f.w_ < 0)), "Octant"] = "-4"
    f.loc[1, ""] = "userinput"

   
    h = f['Octant'].value_counts()


    f.loc[1, "Octant Id"] = "overall Count"
    f.loc[2, "Octant Id"] = "mod"+" " + str(mod)

    number_blocks = math.ceil(numb/mod)
    l = 0
    v = mod
    st_point = 0
    en_point = v-1
    x = 0
    for x in range(number_blocks):
        if (st_point + mod > numb):
            f.loc[x+3, "Octant Id"] = str(st_point)+"-" + str(numb-1)
            break
        else:
            f.loc[x+3, "Octant Id"] = str(st_point)+"-" + str(en_point)

        x = x+1
        st_point = st_point + mod
        en_point = en_point + mod

    try:
        f.loc[1, "+1"] = h["+1"] 
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
   
    f.loc[0 , "+1 "] = "Rank1"  
    f.loc[0 , "-1 "] = "Rank2"
    f.loc[0 , "+2 "] = "Rank3"
    f.loc[0 , "-2 "] = "Rank4"
    f.loc[0 , "+3 "] = "Rank5"
    f.loc[0 , "-3 "] = "Rank6"
    f.loc[0 , "+4 "] = "Rank7"
    f.loc[0 , "-4 "] = "Rank8"
    f.loc[0, "   "] = "Rank1 Octant Id"
    f.loc[0, "    "] = "Rank1 Octant Name"
    
    rank_count = {"+1 " : 0, "-1 " : 0, "+2 ": 0, "-2 ": 0, "+3 ": 0, "-3 ": 0, "+4 ": 0, "-4 ": 0 }
    
    R = 0
    row = 3
    x = 0
    oct1 = oct2 = oct3 = oct4 = oct5 = oct6 = oct7 = oct8 = 0
    for x in range(number_blocks):
        
        for i in range(mod):
            try:
                if f["Octant"][R] == "+1":
                    oct1 = oct1+1
                elif f["Octant"][R] == "-1":
                    oct2 += 1
                elif f["Octant"][R] == "+2":
                    oct3 += 1
                elif f["Octant"][R] == "-2":
                    oct4 += 1
                elif f["Octant"][R] == "+3":
                    oct5 += 1
                elif f["Octant"][R] == "-3":
                    oct6 += 1
                elif f["Octant"][R] == "+4":
                    oct7 += 1
                elif f["Octant"][R] == "-4":
                    oct8 += 1
            except:
                print("Row Not found")

            R = R+1
            if R == numb:
                x = number_blocks+1
                break  # break statement
        x = x+1
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
            
        
        
        # Tut 5 main work 
        #Before moving on to next block 
        ls = [ oct1, oct2 , oct3 , oct4 , oct5 , oct6, oct7 , oct8]
        kt = set({})
        dict = {}
        
        for pl in ls:
            kt.add(pl)
        for pl in kt:
            dict[pl] = []
        # found distinct octant values in st and sorted in increasing order 
        
        i = 0
        for pl in ls:
            dict[pl].append(graph[i])
            i+= 1
        for pl in kt:
            dict[pl].sort(reverse = True)
        
        ran1 = "Initiate"
        # assigning rank to octants 
        rank = 8
        for pl in kt:
            for octant in dict[pl]:
                try:
                   f.loc[row, octant] = rank
                except:
                   print("Index out of Bound Error") 
                if(rank == 1):
                    ran1 = octant
                rank -=  1
                
        # Inserting Rank1 Octant Id and its name 
        f.loc[row, "   "] = ran1
        f.loc[row, "    "] = octant_name_id_mapping[ran1]
        rank_count[ran1] += 1
        oct1 = oct2 = oct3 = oct4 = oct5 = oct6 = oct7 = oct8 = 0
        
        row = row+1
    # inserting overall rank1 count for each octant 
    row += 4
    f.loc[row, "+1"] = "Octant Id"
    f.loc[row , "-1"] = "Octant Name"
    f.loc[row , "+2"] = "Count of Rank 1 Mod Values"
    
    row += 1
    cn = 0
    # graph2 = {0:"+1" , 1:"-1"  , 2:"+2"  , 3:"-2"  , 4:"+3"  , 5:"-3"  , 6:"+4"  , 7:"-4"}
    for key in octant_name_id_mapping:
        f.loc[row, "+1"] = graph[cn]
        f.loc[row, "-1"] = octant_name_id_mapping[graph[cn]]
        f.loc[row, "+2"] = rank_count[graph[cn]]
        row += 1
        cn += 1
        
    f.to_excel('./octant_output_ranking_excel.xlsx')


ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


mod = 5000
octant_range_names(mod)


# This shall be the last lines of the code.
