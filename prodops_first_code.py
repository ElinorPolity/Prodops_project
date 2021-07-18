import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

data = pd.read_excel("C:\PRODOPS\data_try.xlsx")
data=data.drop(0).reset_index()
data=data.rename(columns={"חובה (מקומי)":"costs","שם חשבון קיזוז":"compeny_of_service","תאריך ערך":"date"})
data

#let see how much money we pay every year 
data["year"]=data["date"].apply(lambda row: str(str(row)[:4]))
data_per_year_total=pd.DataFrame()
data["sum_of_pay_for_a_year"]=data.groupby("year")["costs"].transform("sum")
plt.plot(data["year"], data["sum_of_pay_for_a_year"])  # Plot the chart
plt.show()

#how much did we spend on each compeny in all the mechine life time/time of service

type_group = data["costs"].groupby(data["compeny_of_service"])
x1=type_group.sum()
x2=type_group.size()

data_oo=pd.merge(x1,x2,on="compeny_of_service")
data_oo=data_oo.reset_index()
data_oo=data_oo.rename(columns={"costs_x":"costs","costs_y":"number_of_events"})
data_oo["normalize"]=round(data_oo.costs/data_oo.number_of_events,2)
data_oo["revers_text"]=data_oo["compeny_of_service"].apply(lambda row : row[::-1])
data_oo

data_oo.plot(x="revers_text",y="number_of_events",kind="bar")
data_oo.plot(x="revers_text",y="costs",kind="bar")
data_oo.plot(x="revers_text", y=["costs", "normalize"], kind="bar")

data_per_year=pd.DataFrame()
data_per_year["compeny_and_year"]=data[["compeny_of_service","year"]].agg('-'.join, axis=1)

# let see how much we take out on each compeny per year
type_group = data["costs"].groupby(data_per_year["compeny_and_year"])
x1=type_group.sum()
x2=type_group.size()

data_new=pd.merge(x1,x2,on="compeny_and_year")
data_new=data_new.reset_index()
data_new=data_new.rename(columns={"costs_x":"costs","costs_y":"number_of_events"})
data_new["normalize"]=round(data_new.costs/data_new.number_of_events,2)
data_new["revers_text"]=data_new["compeny_and_year"].apply(lambda row : row[::-1])
data_new["year"]=data_new["compeny_and_year"].apply(lambda row :row[len(row)-4:])
data_new

#creating plot fo each year
years_arry=data_new["year"].drop_duplicates().sort_values()
years_arry
for year in years_arry :
    ooo=data_new[data_new["year"]==year]
    ooo.plot(x="revers_text", y=["costs", "normalize"], kind="bar",title="year "+year)



##tring another option
data_new.drop("number_of_events")
grouped = data_new.groupby('year')

ncols=2
nrows = int(np.ceil(grouped.ngroups/ncols))

fig, axes = plt.subplots(nrows=nrows, ncols=ncols, figsize=(12,12), sharey=True)

for (key, ax) in zip(grouped.groups.keys(), axes.flatten()):
    grouped.get_group(key).plot(ax=ax)

ax.legend()
plt.show()
