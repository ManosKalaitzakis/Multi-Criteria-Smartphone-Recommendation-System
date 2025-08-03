from utamethods import UTA, UTASTAR, UTADIS
import pandas as pd
data=pd.read_excel("Book1.xlsx")
bestvalues=[2,10,3]
worstvalues=[30,40,0]
monotonicity=[0,0,1]
intervals=[4,5,5]
res=UTA(data,bestvalues=bestvalues, worstvalues=worstvalues, intervals=intervals, monotonicity=monotonicity)
for a in res.crit_utilities:

    #print(res.crit_utilities)
    print(a)

