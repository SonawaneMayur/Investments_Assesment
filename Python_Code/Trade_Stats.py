from Trade_Environment import Trade_Files
from utility import write_Excel

import pandas as pd

tf = Trade_Files()

order_xl = tf.get_trade_order_dataframe()

confirm_xl = tf.get_trade_confirms_dataframe()

#order_df = pd.DataFrame(order_xl)
#print(confirm_xl.columns)
isin_order_list = order_xl['ISIN']
isin_confirm_list = confirm_xl['ISIN Code']

print(len(isin_order_list))
print(len(isin_order_list.unique()))

print(len(isin_confirm_list))
print(len(isin_confirm_list.unique()))

# print(confirm_xl[confirm_xl.duplicated('ISIN Code')])
# d_isin  = confirm_xl['ISIN Code'].values_counts()
# print(d_isin[d_isin >1])

#print(confirm_xl.groupby('ISIN Code').sum())
out_columns =['ISIN', 'SECURITY_NAME','SIDE','EXECUTED_QUANTITY','ORDER_QUANTITY', 'CURRENCY' ]
completed_order = pd.DataFrame(columns = out_columns)
partial_order = pd.DataFrame(columns= out_columns)
irregularities_order = pd.DataFrame(columns= out_columns)
c_row_id = 0
p_row_id = 0
i_row_id = 0
#ISIN Code	CY	Name of Security	Type	Exchange	 Quantity Executed 	 Quantity Order
complete_order_isin = set()
partial_order_isin = set()

for index,row in confirm_xl.iterrows():
    if row["Quantity Executed"] == row["Quantity Order"] and row['ISIN Code'] not in partial_order_isin:
        completed_order.loc[c_row_id, 'ISIN'] = row['ISIN Code']
        completed_order.loc[c_row_id, 'SECURITY_NAME'] = row['Name of Security']
        completed_order.loc[c_row_id, 'SIDE'] = row['Type']
        completed_order.loc[c_row_id, 'EXECUTED_QUANTITY'] = row['Quantity Executed']
        completed_order.loc[c_row_id, 'ORDER_QUANTITY'] = row['Quantity Order']
        completed_order.loc[c_row_id, 'CURRENCY'] = row['Cur. Comm.']

        complete_order_isin.add(row['ISIN Code'])
        c_row_id += 1

    elif row["Quantity Executed"] < row["Quantity Order"] and row['ISIN Code'] not in complete_order_isin:
        partial_order.loc[p_row_id, 'ISIN'] = row['ISIN Code']
        partial_order.loc[p_row_id, 'SECURITY_NAME'] = row['Name of Security']
        partial_order.loc[p_row_id, 'SIDE'] = row['Type']
        partial_order.loc[p_row_id, 'EXECUTED_QUANTITY'] = row['Quantity Executed']
        partial_order.loc[p_row_id, 'ORDER_QUANTITY'] = row['Quantity Order']
        partial_order.loc[p_row_id, 'CURRENCY'] = row['Cur. Comm.']

        partial_order_isin.add(row['ISIN Code'])
        p_row_id += 1

    else:
        irregularities_order.loc[i_row_id, 'ISIN'] = row['ISIN Code']
        irregularities_order.loc[i_row_id, 'SECURITY_NAME'] = row['Name of Security']
        irregularities_order.loc[i_row_id, 'SIDE'] = row['Type']
        irregularities_order.loc[i_row_id, 'EXECUTED_QUANTITY'] = row['Quantity Executed']
        irregularities_order.loc[i_row_id, 'ORDER_QUANTITY'] = row['Quantity Order']
        irregularities_order.loc[i_row_id, 'CURRENCY'] = row['Cur. Comm.']

        partial_order_isin.add(row['ISIN Code'])
        i_row_id += 1

print('sfsfsf')
print(len(completed_order['ISIN']))
print(len(partial_order['ISIN']))
print(len(irregularities_order['ISIN']))

print(write_Excel(completed_order,partial_order,irregularities_order))


# writer = pd.ExcelWriter('C://Users//Mayur//Documents//Assesment//Qtron_Investments//Trade_Files//Trade_Stats.xlsx', engine='xlsxwriter')
#
# # Write each dataframe to a different worksheet.
# completed_order.to_excel(writer, sheet_name='Complete_Orders')
# partial_order.to_excel(writer, sheet_name='Partial_Orders')
# irregularities_order.to_excel(writer, sheet_name='Irregularities_Orders')
#
# # Close the Pandas Excel writer and output the Excel file.
# writer.save()