import pandas as pd
from Trade_Environment import Trade_Files
tf = Trade_Files()
print(tf.get_trade_output_file())
def write_Excel(completed_order,partial_order,irregularities_order):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(tf.get_trade_output_file(), engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    completed_order.to_excel(writer, sheet_name='Complete_Orders')
    partial_order.to_excel(writer, sheet_name='Partial_Orders')
    irregularities_order.to_excel(writer, sheet_name='Irregularities_Orders')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    return True