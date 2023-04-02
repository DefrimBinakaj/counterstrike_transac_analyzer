import pandas as pd


# ----------------------------------------------------------------------------------------------------------------------------------
# 'Item Name', 'Game Name' , 'Listed On' , 'Acted On', ' Display Price' ,  ' Price in Cents' , 
# ' Type' , ' Market Name' , ' App Id' , ' Context Id' , ' Asset Id' , ' Instance Id' , ' Class Id' , 
# ' Unowned Context Id' , ' Unowned Id' , ' Partner Name', ' Partner Link'

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def highlightPrice(merged_data, worksheet):
    light_blue_fill = PatternFill(start_color="EEFEFF", end_color="EEFEFF", fill_type="solid")
    display_price_col_idx = merged_data.columns.get_loc("Combined_Price") + 1

    for row in range(1, len(merged_data) + 2):
        worksheet.cell(row=row, column=display_price_col_idx).fill = light_blue_fill
    



def highlightType(merged_data, worksheet):

    red_fill = PatternFill(start_color="FFF4EE", end_color="FFF4EE", fill_type="solid")
    green_fill = PatternFill(start_color="EEFFF5", end_color="EEFFF5", fill_type="solid")

    col_idx = merged_data.columns.get_loc(' Type') + 1
    for row in range(2, len(merged_data) + 2):
        cell_value = worksheet.cell(row=row, column=col_idx).value
        if cell_value == 'purchase':
            for col in range(1, len(merged_data.columns) + 1):
                worksheet.cell(row=row, column=col).fill = red_fill
        elif cell_value == 'sale':
            for col in range(1, len(merged_data.columns) + 1):
                worksheet.cell(row=row, column=col).fill = green_fill
    

def highlightGame(merged_data, worksheet):

    orange_fill = PatternFill(start_color="FFF5CC", end_color="FFF5CC", fill_type="solid")
    blue_fill = PatternFill(start_color="CCDAFF", end_color="CCDAFF", fill_type="solid")

    col_idx = merged_data.columns.get_loc('Game Name') + 1
    for row in range(2, len(merged_data) + 2):
        cell_value = worksheet.cell(row=row, column=col_idx).value
        if cell_value == 'TF2':
            worksheet.cell(row=row, column=col_idx).fill = orange_fill
        elif cell_value == 'CSGO':
            worksheet.cell(row=row, column=col_idx).fill = blue_fill


def highlightSheet(output_file_name, merged_data):

    with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
        merged_data.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']

        # type
        highlightType(merged_data, worksheet)

        # price
        # highlightPrice(merged_data, worksheet)

        # game
        highlightGame(merged_data, worksheet)

        writer.save()



def replaceGameNames(data):

    data['Game Name'] = data['Game Name'].replace('Counter-Strike: Global Offensive', 'CSGO')
    data['Game Name'] = data['Game Name'].replace('Team Fortress 2', 'TF2')

    return data


def combineBulk(data):

    # clean up date columns
    data['Acted On'] = pd.to_datetime(data['Acted On']).dt.date
    data['Listed On'] = pd.to_datetime(data['Listed On']).dt.date

    # Create a new column to store the original index
    data['original_index'] = data.index

    # Group by 'Acted On' and 'Price in Cents', keep the first non-null value in each group, and calculate the count
    grouped_data = data.groupby(['Acted On', ' Price in Cents'], as_index=False).first()
    grouped_data['Count'] = data.groupby(['Acted On', ' Price in Cents']).size().values
    grouped_data['Combined_Price'] = grouped_data['Count'] * grouped_data[' Price in Cents'] / 100

    # Sort by the original index
    grouped_data = grouped_data.sort_values('original_index')

    # Drop the original index column
    grouped_data = grouped_data.drop(columns=['original_index'])

    # # clean up Acted On column
    # grouped_data['Acted On'] = grouped_data.to_datetime(grouped_data['Acted On']).dt.date

    # Drop unnecessary columns and reorder
    merged_data = grouped_data[['Item Name', 
                                'Game Name', 
                                'Listed On', 
                                'Acted On', 
                                ' Display Price', 
                                # ' Price in Cents',
                                ' Type', 
                                ' Market Name', 
                                # ' App Id', 
                                # ' Context Id', 
                                ' Asset Id', 
                                # ' Instance Id', 
                                ' Class Id',
                                # ' Unowned Context Id', 
                                ' Unowned Id', 
                                ' Partner Name', 
                                ' Partner Link', 
                                'Count', 
                                'Combined_Price']]
    
    
    return merged_data






def createOutputSheet(merged_data):
    # Save the resulting data to a new Excel file
    output_file_name = 'output_data_csgo_t4.xlsx'
    writer = pd.ExcelWriter(output_file_name, engine='openpyxl')
    merged_data.to_excel(writer, index=False, sheet_name='output1')
    writer.save()







def main():

    file_name = 'steam_market_history.xlsx'
    sheet_name = 'steam_market_history'
    data = pd.read_excel(file_name, sheet_name=sheet_name)

    data = replaceGameNames(data)

    merged_data = combineBulk(data)

    

    # createOutputSheet(merged_data)
    # highlightPrice("output_data_csgo_colour_2.xlsx", merged_data)
    highlightSheet("output_data_csgo_all_2.xlsx", merged_data)




    print("done")


if __name__ == "__main__":
    main()