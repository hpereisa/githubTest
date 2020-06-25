import numpy as np
import pandas as pd
import xlrd
import matplotlib.pyplot as plt


#reading EverClear file into dataframe
ec = pd.read_excel('ec.xlsx', dtype={"Transaction Date": 'datetime64',
    "Marketplace ID": 'str',
    "Merchant Customer ID Zerengeti": 'str'
    })

#renaming EverClear df columns
ec.rename(columns={"Transaction Date": 'transaction_date',
    "Marketplace ID": 'market',
    "Merchant Customer ID Zerengeti": 'merchant',
    "Sale Price": 'sale_price',
    "Coupon": 'coupon',
    "Tax": 'tax',
    "Shipping Revenue": 'shipping',
    "Shipping Tax": 'shipping_tax',
    "Credit Card (Txn)": 'credit_card'
    }, inplace=True)

#calculating the net_receivable in df (EC)
# credit card ( does not contain the export principal fee ) 
#ec['net_receivable'] = ec['credit_card'] + ec['coupon']


#calculating the net_receivable in df (EC)
#replacing ['credit_card'] with the following line. This is to make up for export fee not in dryad. (temporary)
ec['net_receivable'] = ec['sale_price'] + ec['tax'] + ec['shipping'] + ec['shipping_tax'] + ec['coupon']


#reading Dryad into df
dryad = pd.read_excel('dryad.xlsx', dtype={"merchant": 'str'}, parse_dates=['transaction_date'])


#merging EC and dryad df into merged df
ship_recon = pd.merge(
        dryad[['transaction_date', 'market', 'merchant', 'net_receivable']],
        ec[['transaction_date', 'merchant', 'net_receivable']],
        on=['transaction_date', 'merchant'],
        how='outer',
        suffixes=['_dryad', '_ec']
        )

# calculating the net variance between the two systems and sorting the values before exporting. 
ship_recon['net_ar_variance'] = ship_recon['net_receivable_dryad'] - ship_recon['net_receivable_ec']
ship_recon['net_ar_var_abs'] = ship_recon['net_ar_variance'].abs()
ship_recon.sort_values(['transaction_date', 'merchant'],  ascending=True, inplace=True)



#setting up excel engine writer
writer = pd.ExcelWriter('ship_recon.xlsx', engine='xlsxwriter')
ship_recon.to_excel(writer, sheet_name='sheet1', index=False, encoding='utf8')
writer.save()

# for plotting

ship_graph = ship_recon.groupby(['transaction_date', 'market'])['net_ar_variance'].sum().unstack().plot()
plt.show()
