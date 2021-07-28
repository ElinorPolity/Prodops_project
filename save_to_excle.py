import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime
import matplotlib.dates as mdates


def data_analizing(writer,i,data):
    
    data=data.reset_index()
    title_of_machine=str(data.head(1)["הערות"])[5:52]
    data=data.drop(0).reset_index()
    
    #in this part im creating a plot based half a year
    data["year"]=data["date"].apply(lambda row: str(str(row)[:4]))
    data["part_of_year"]=data["date"].apply(lambda row: "06" if int(str(row)[5:7])<7 else "12")#changing dates to first or seconde part of the year
    data["part_of_year"]=data["part_of_year"].apply(lambda row: str(row))
    data=data.sort_values(by=["year","part_of_year"])
    data_per__half_year_total=pd.DataFrame()
    data["updated_date"]=data[['year',"part_of_year"]].agg('-'.join, axis=1)#joining together so i could group
    data["sum_of_pay_for_a_half_year"]=data.groupby("updated_date")["costs"].transform("sum")#gruping and sum
    data_per__half_year_total[["updated_date","sum_of_pay_for_a_half_year"]]=data[["updated_date","sum_of_pay_for_a_half_year"]]
    data_per__half_year_total=data_per__half_year_total.drop_duplicates().reset_index()
    data_per__half_year_total=data_per__half_year_total.drop("index",axis=1)
    data_per__half_year_total["date"]=data_per__half_year_total["updated_date"].apply(lambda row : row[:4]+" - 1" if  row[5:7]=="06" else row[:4]+" - 2")
    
    
    x1=mdates.datestr2num(data_per__half_year_total["updated_date"])#turning the str to a date object of mdates
    fig=plt.plot_date(x1,data_per__half_year_total["sum_of_pay_for_a_half_year"],fmt="bo", tz=None, xdate=True,linestyle='solid', marker='None')
    plt.title("costs per half a year")
    plt.savefig(r'C:\PRODOPS\images\python_pretty_plot'+i+'.png',dpi=300, bbox_inches='tight')
    plt.clf()#tels the plt we are done with tha plot
    
    ##in this part im creating a plot based on quarter
    def quarter(row):
        if int(str(row)[5:7])<4 :
            return "03"
        if int(str(row)[5:7])>3 and int(str(row)[5:7])<7:
            return "06"
        if int(str(row)[5:7])>6 and int(str(row)[5:7])<10:
            return "09"
        if int(str(row)[5:7])>9:
            return "12"
    data["year"]=data["date"].apply(lambda row: str(str(row)[:4]))    
    data["part_of_year"]=data["date"].apply(lambda row:quarter(row) )#changing dates to first or seconde part of the year
    data["part_of_year"]=data["part_of_year"].apply(lambda row: str(row))
    data=data.sort_values(by=["year","part_of_year"])
    data_per__quarter_total=pd.DataFrame()
    data["updated_date"]=data[['year',"part_of_year"]].agg('-'.join, axis=1)#joining together so i could group
    data["sum_of_pay_for_a_quarter"]=data.groupby("updated_date")["costs"].transform("sum")#gruping and sum
    data_per__quarter_total[["updated_date","sum_of_pay_for_a_quarter"]]=data[["updated_date","sum_of_pay_for_a_quarter"]]
    data_per__quarter_total=data_per__quarter_total.drop_duplicates().reset_index()
    data_per__quarter_total=data_per__quarter_total.drop("index",axis=1)
    
    x1=mdates.datestr2num(data_per__quarter_total["updated_date"])#turning the str to a date object of mdates
    plt.plot_date(x1,data_per__quarter_total["sum_of_pay_for_a_quarter"],fmt="bo", tz=None, xdate=True,linestyle='solid', marker='None')
    plt.title("costs per quarter")
    plt.savefig(r'C:\PRODOPS\images\quarter_plot'+i+'.png',dpi=300, bbox_inches='tight')
    plt.clf()
    
    #in this part, creating a pie plot per year
    data_per_year_total=pd.DataFrame()
    data["sum_of_pay_for_a_year"]=data.groupby("year")["costs"].transform("sum")
    data_per_year_total[["year","sum_of_pay_for_a_year"]]=data[["year","sum_of_pay_for_a_year"]]
    data_per_year_total=data_per_year_total.drop_duplicates().reset_index()
    plt.pie(data_per_year_total["sum_of_pay_for_a_year"],labels=data_per_year_total["year"],shadow = True,)
    plt.title("costs per year")
    plt.legend(title = "year:")
    plt.savefig(r'C:\PRODOPS\images\python_pretty_plot_pie'+i+'.png', dpi=300, bbox_inches='tight')
    plt.clf()
    
    #how much did we spend on each compeny in all the mechine life time/time of service

    type_group = data["costs"].groupby(data["compeny_of_service"])
    x1=type_group.sum()
    x2=type_group.size()

    data_oo=pd.merge(x1,x2,on="compeny_of_service")
    data_oo=data_oo.reset_index()
    data_oo=data_oo.rename(columns={"costs_x":"costs","costs_y":"number_of_events"})
    data_oo["normalize"]=round(data_oo.costs/data_oo.number_of_events,2)
    data_oo["revers_text"]=data_oo["compeny_of_service"].apply(lambda row : row[::-1])
    
    p1=data_oo.plot(x="revers_text",y="number_of_events",kind="bar",title='number of events in each compeny')
    fig1 = p1.get_figure()
    fig1.savefig(r'C:\PRODOPS\images\s'+i+'.png', dpi=300, bbox_inches='tight')
    fig1.clf()
    
    p2=data_oo.plot(x="revers_text",y="costs",kind="bar",title='cost for each compeny')
    fig2 = p2.get_figure()
    fig2.savefig(r'C:\PRODOPS\images\cost_per_compeny'+i+'.png', dpi=300, bbox_inches='tight')
    fig2.clf()
    
    p3=data_oo.plot(x="revers_text", y=["costs", "normalize"], kind="bar",title = "cost per events (normalized cost)")
    fig3 = p3.get_figure()
    fig3.savefig(r'C:\PRODOPS\images\n'+i+'.png', dpi=300, bbox_inches='tight')
    fig3.clf()
    
    #anlyzing compenys per year
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


    # Write each dataframe to a different worksheet.
    data.head(0)["הערות"].to_excel(writer, sheet_name='Sheet'+i)
    data_oo=data_oo.drop("revers_text",axis=1)
    
    data_per__excle1=pd.DataFrame()
    data_per__excle2=pd.DataFrame()
    data_per__excle1[["שנה וחציון","עלות כוללת"]] = data_per__half_year_total[["date","sum_of_pay_for_a_half_year"]]
    data_per__excle2[["שנה ורבעון","עלות כוללת"]] = data_per__quarter_total[["updated_date","sum_of_pay_for_a_quarter"]]
    data_per__excle1.to_excel(writer, sheet_name='Sheet'+i,startrow=2 )
    data_per__excle2.to_excel(writer, sheet_name='Sheet'+i, startcol=5,startrow=2)
    data_oo.to_excel(writer, sheet_name='Sheet'+i, startcol=10,startrow=2)
    
    worksheet = writer.sheets['Sheet'+i]
    worksheet.write(0, 10, title_of_machine)
    worksheet.insert_image(46,1,'C:\PRODOPS\images\python_pretty_plot'+i+'.png')
    worksheet.insert_image(46,10,'C:\PRODOPS\images\quarter_plot'+i+'.png')
    worksheet.insert_image(26,0,'C:\PRODOPS\images\python_pretty_plot_pie'+i+'.png')
    worksheet.insert_image(4,18,'C:\PRODOPS\images\\s'+i+'.png')
    worksheet.insert_image(4,27,'C:\PRODOPS\images\cost_per_compeny'+i+'.png')
    worksheet.insert_image(30,22,'C:\PRODOPS\images\\n'+i+'.png')
    
    #in this part we save the plots of the compenys in each year
    years_arry=data_new["year"].drop_duplicates().sort_values()
    years_arry
    index = 1
    for year in years_arry :
        ooo=data_new[data_new["year"]==year]
        p=ooo.plot(x="revers_text", y=["costs", "normalize"], kind="bar",title="year "+year)
        fig = p.get_figure()
        fig.savefig(r'C:\PRODOPS\images\compemy_yearly_sheet'+i+str(index)+'.png', dpi=300, bbox_inches='tight')
        fig.clf()
        worksheet.insert_image(66,10*index-10,'C:\PRODOPS\images\\compemy_yearly_sheet'+i+str(index)+'.png')
        index+=1
    


if __name__ == "__main__":
    writer = pd.ExcelWriter(r'C:\PRODOPS\analized_data.xlsx', engine='xlsxwriter')
    data = pd.read_excel("C:\PRODOPS\Data.xlsx")
    data=data.rename(columns={"חובה (מקומי)":"costs","שם חשבון קיזוז":"compeny_of_service","תאריך ערך":"date"})
    i=1
    start=0
    for index, row in data.iterrows():
        if(row["תאריך אסמכתא"]=="עלות המכירות"and index!=0 ):
            Data_per_mechine=data.loc[ start:index-2 , : ]
            data_analizing(writer,str(i),Data_per_mechine)
            start=index
            i=i+1
        if(index==len(data)-1):
            Data_per_mechine=data.loc[ start:index-2 , : ]
            data_analizing(writer,str(i),Data_per_mechine)
            start=index
            i=i+1
    writer.save()
