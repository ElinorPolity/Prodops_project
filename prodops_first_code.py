import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

data = pd.read_excel("C:\PRODOPS\data_try.xlsx")
data

type_group = data["חובה (מקומי)"].groupby(data["שם חשבון קיזוז"])
x1=type_group.sum()
x2=type_group.size()

data_oo=pd.merge(x1,x2,on="שם חשבון קיזוז")
data_oo=data_oo.reset_index()

data_oo=data_oo.rename(columns={"חובה (מקומי)_x":"costs","חובה (מקומי)_y":"number_of_events"})
data_oo["normalize"]=round(data_oo.costs/data_oo.number_of_events,2)
data_oo["revers_text"]=data_oo["שם חשבון קיזוז"].apply(lambda row : row[::-1])

p1=data_oo.plot(x="revers_text",y="number_of_events",kind="bar")
p1=data_oo.plot(x="revers_text",y="costs",kind="bar")
p2=data_oo.plot(x="revers_text", y=["costs", "normalize"], kind="bar")
