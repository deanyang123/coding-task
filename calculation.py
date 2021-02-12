import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

df = pd.read_excel(r"C:\Users\Deany\Documents\Accounting\SPX.xlsx")

df_returns = []
start = df.iloc[0].Close
ret = 0

for i in range(1, len(df)):
    data = df.iloc[i]
    if data["Date"].month == df.iloc[i - 1]["Date"].month:
        ret = ((data.Close / start) - 1) * 100
    else:
        #signifies a new month, append previous result
        ptime = df.iloc[i-1]["Date"]
        timestamp = pd.Timestamp(f"{ptime.year}-{ptime.month}-28")
        df_returns.append([timestamp, ret])
        
        start = data.Close

#edge case for the last month, i need to append it manually
ptime = df.iloc[len(df) - 1]["Date"]
timestamp = pd.Timestamp(f"{ptime.year}-{ptime.month}-28")
df_returns.append([timestamp, ret])


#convert results to a dataframe
df_ret=pd.DataFrame(df_returns, columns= ["Date", "Returns"])


#plotting and saving the graph
plt.plot(df_ret["Returns"])
plt.ylabel("Percentage Return")
plt.xlabel("Month Number")

pathg = r"C:\Users\Deany\Documents\Accounting\returns_graph.png"

plt.savefig(pathg, bbox_inches="tight")
plt.show()


#upload results into an excel sheet
pathxl = r'C:\Users\Deany\Documents\Accounting\returns.xlsx'

workbook = xlsxwriter.Workbook(pathxl)
writer = pd.ExcelWriter(pathxl, engine='xlsxwriter')

worksheet1 = df_ret.to_excel(writer, sheet_name="Sheet1", index=False)
worksheet2 = pd.DataFrame([],[]).to_excel(writer, sheet_name="Sheet2", index=False)

writer.sheets["Sheet2"].insert_image('B5', pathg)
writer.save()