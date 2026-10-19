# Builds inventory_tracker_real.xlsx from UCI Online Retail (CC BY 4.0), 15 real products.
# ASSUMPTIONS (dataset has no stock or cost data):
#  - avg daily demand = units sold 1 Dec 2010 - 31 Oct 2011 / 334 days
#  - opening stock on 1 Nov 2011 = 60 days of avg demand; reorder level = 14 days of avg demand
#  - unit cost = 60% of avg selling price; current stock = opening - real units sold 1 Nov - 9 Dec 2011
import pandas as pd
df=pd.read_excel('data/Online Retail.xlsx')
d=df[(df.Quantity>0)&(df.Quantity<10000)&(df.UnitPrice>0)&~df.InvoiceNo.astype(str).str.startswith('C')
     &df.StockCode.astype(str).str.match(r'^\d{5}[A-Z]?$')].copy()
d['StockCode']=d.StockCode.astype(str); d['Rev']=d.Quantity*d.UnitPrice
hist=d[d.InvoiceDate<'2011-11-01']; rec=d[d.InvoiceDate>='2011-11-01'].copy()
g=hist.groupby('StockCode').agg(desc=('Description','first'),units=('Quantity','sum'),rev=('Rev','sum'),inv=('InvoiceNo','nunique'))
top=g.sort_values('inv',ascending=False).head(15).copy()
days=(pd.Timestamp('2011-11-01')-hist.InvoiceDate.min().normalize()).days
top['daily']=top.units/days
pr=d[d.StockCode.isin(top.index)].groupby('StockCode').apply(lambda x:x.Rev.sum()/x.Quantity.sum())
cat=lambda s:'Bags' if 'BAG' in s else 'Lighting' if any(k in s for k in['LIGHT','LANTERN','CANDLE']) else 'Kitchen' if any(k in s for k in['CAKE','TIN','JAR','CUP','PLATE','POPCORN','MUG','LUNCH','TEA']) else 'Decor'
rec=rec[rec.StockCode.isin(top.index)]
mk={'United Kingdom':'UK','EIRE':'Ireland','France':'France','Germany':'Germany'}
rec['Market']=rec.Country.map(mk).fillna('Other EU/Intl')
agg=rec.groupby([rec.InvoiceDate.dt.normalize(),'StockCode','Market']).agg(q=('Quantity','sum'),n=('InvoiceNo','nunique')).reset_index()
print('sales rows',len(agg),'days',days)

import json
prods=[[c,r.desc.strip().title().replace('  ',' ').rstrip('.'),cat(r.desc),round(pr[c]*0.6,2),round(pr[c],2),int(round(r.daily*60)),int(round(r.daily*14))] for c,r in top.iterrows()]
sales=[[str(r.InvoiceDate.date()),r.StockCode,int(r.q),r.Market,f"{r.n} invoices"] for r in agg.sort_values('InvoiceDate').itertuples()]
json.dump({'products':prods,'sales':sales},open('data/real_data.json','w'))
