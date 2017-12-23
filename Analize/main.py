import pandas as pd
import statsmodels.tsa.api as tsa
from datetime import timedelta


fn = '..//rates20061226-20170127.xlsx'
print fn

xd = pd.ExcelFile(fn)

adf = xd.parse(xd.sheet_names[-1], header=0, index_col=0)

sdtARMA = (adf.last_valid_index()+timedelta(days=1)).strftime('%Y%m%d')
sdt = adf.last_valid_index()
for i in range(30):
    adf = adf.append(pd.DataFrame([0], index  = [adf.last_valid_index()+timedelta(days=1)], columns = [adf.axes[1][0]]))
edtARMA = adf.last_valid_index().strftime('%Y%m%d')

dfusd = pd.DataFrame()
dfusd['USD'] = adf['USD']

dfeur = pd.DataFrame()
dfeur['EUR'] = adf['EUR']

dfusd['ARMA'] = tsa.ARMA(dfusd[:sdt], (2,1)).fit().predict(sdtARMA, edtARMA)
dfeur['ARMA'] = tsa.ARMA(dfeur[:sdt], (2,1)).fit().predict(sdtARMA, edtARMA)

for w in [5, 10, 15, 20, 25, 30, 120]:
    dfusd['MA' + str(w)] = dfusd['USD'].rolling(window=w).mean()
    dfeur['MA' + str(w)] = dfeur['EUR'].rolling(window=w).mean()


fn = str.format('..//{0}-{1}-{2}.xlsx', 'AnalizeIT', adf.first_valid_index().strftime('%Y%m%d'), sdt.strftime('%Y%m%d'))
writer = pd.ExcelWriter(fn)
dfusd.to_excel(writer,'USD')
dfusd.to_excel(writer,'EUR')
writer.save()
