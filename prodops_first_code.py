import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

data = pd.read_excel("C:\PRODOPS\data_try.xlsx")
data=data.drop(0).reset_index()
data=data.rename(columns={"חובה (מקומי)":"costs","שם חשבון קיזוז":"compeny_of_service","תאריך ערך":"date"})
data

#in this part im creating a plot based half a year
data["year"]=data["date"].apply(lambda row: str(str(row)[:4]))
data["part_of_year"]=data["date"].apply(lambda row: "01" if int(str(row)[6:7])<7 else "07")#changing dates to first or seconde part of the year
data["part_of_year"]=data["part_of_year"].apply(lambda row: str(row))
data_per__half_year_total=pd.DataFrame()
data["updated_date"]=data[['year',"part_of_year"]].agg('-'.join, axis=1)#joining together so i could group
data["sum_of_pay_for_a_half_year"]=data.groupby("updated_date")["costs"].transform("sum")#gruping and sum
data_per__half_year_total[["updated_date","sum_of_pay_for_a_half_year"]]=data[["updated_date","sum_of_pay_for_a_half_year"]]
data_per__half_year_total=data_per__half_year_total.drop_duplicates().reset_index()
data_per__half_year_total

x1=mdates.datestr2num(data_per__half_year_total["updated_date"])#turning the str to a date object of mdates
plt.plot_date(x1,data_per__half_year_total["sum_of_pay_for_a_half_year"],fmt="bo", tz=None, xdate=True,linestyle='solid', marker='None')
plt.title("costs per half a year")
plt.show()

#in this part, creating a pie plot per year
data_per_year_total=pd.DataFrame()
data["sum_of_pay_for_a_year"]=data.groupby("year")["costs"].transform("sum")
data_per_year_total[["year","sum_of_pay_for_a_year"]]=data[["year","sum_of_pay_for_a_year"]]
data_per_year_total=data_per_year_total.drop_duplicates().reset_index()
plt.pie(data_per_year_total["sum_of_pay_for_a_year"],labels=data_per_year_total["year"],shadow = True)
plt.legend(title = "year:")
plt.show()

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


