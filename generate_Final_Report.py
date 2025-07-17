import pandas as pd
from datetime import datetime
from calendar import monthrange

print("Loading Commander Data File...")
master_data = pd.read_excel("/Users/vaishshirsekar/Documents/Automation/Commander_Data.xlsx")

with pd.ExcelWriter("/Users/vaishshirsekar/Documents/Automation/Final_Sales_Report_Generated.xlsx", engine='xlsxwriter') as writer:
    # Update the sheet name to include the current date in the desired format
    today_date = datetime.now().strftime('%m.%d.%y')  # Format as '07.03.25'
    sheet_name_with_date = f"fc_v_actual_Data_{today_date.lstrip('0').replace('.0', '.')}"  # Remove leading zeros

    print(f"Writing commander Data Sheet as '{sheet_name_with_date}'")
    master_data.to_excel(writer, sheet_name=sheet_name_with_date, index=False)

    master_data.columns = master_data.columns.str.strip().str.lower()

    # Update columns_needed to use lowercase column names
    columns_needed = ['forecaster', 'sku', 'sales rep', 'bu', 'starting forecast', 'lag 3 fcst',
                      'forecast', 'shipped', 'delivery notes', 'current allocation', 'open orders']
    data = master_data[columns_needed].copy()
    data['bu'] = data['bu'].str.upper()


    def map_bu_group(value):
        value = str(value).upper()
        if any(sub in value for sub in ["LIFESTYLE-HEADPHONES", "GAMING-HEADPHONES", "SPORT-HEADPHONES"]):
            return "HEADPHONES"
        return value


    data['BU_GROUPED'] = data['bu'].apply(map_bu_group)


    def modify_bu(value):
        value = str(value).upper()
        if "LIFESTYLE" in value:
            return "HEADPHONES - LIFESTYLE"
        elif "GAMING" in value:
            return "HEADPHONES - GAMING"
        elif "SPORTS" in value:
            return "HEADPHONES - SPORTS"
        else:
            return value


    data['BU'] = data['bu'].apply(modify_bu)
    data['Total'] = data['shipped'] + data['delivery notes'] + data['current allocation'] + data['open orders']
    data['Result'] = data['Total'] - data['forecast']
    data['% of Forecast'] = data.apply(lambda row: (row['Total'] / row['forecast']) if row['forecast'] != 0 else 0,
                                       axis=1)
    data['Status'] = data['Result'].apply(
        lambda x: 'Overselling' if x > 0 else ('Underselling' if x < 0 else 'On Target'))

    # Rename columns before calculations
    final_columns = {
        'forecaster': 'FORECASTER', 'sku': 'SKU', 'sales rep': 'Sales Rep', 'bu': 'BU',
        'starting forecast': 'Starting Forecast', 'Lag 3 Fcst': 'Lag 3 Fcst', 'forecast': 'Forecast',
        'shipped': 'Shipped',
        'delivery notes': 'Delivery Notes', 'current allocation': 'Current Allocation', 'open orders': 'Open Orders'
    }
    data.rename(columns=final_columns, inplace=True)

    # Add new columns
    today = datetime.now()
    total_days_in_month = monthrange(today.year, today.month)[1]
    percent_month_passed = round((today.day / total_days_in_month) * 100)  # Round to nearest whole number

    print(f"Today's Day: {today.day}")
    print(f"Total Days in Month: {total_days_in_month}")
    print(f"% of Month Passed (Rounded): {percent_month_passed}")

    data['% of Month Passed'] = percent_month_passed / 100  # Convert back to fraction for Excel formatting
    data['Forecast Consumed (Expected)'] = round(data['Forecast'] * data['% of Month Passed'])  # Rounded values
    data['% of Forecast Consumed'] = data.apply(
        lambda row: (row['Total'] / row['Forecast']) if row['Forecast'] != 0 else 0, axis=1
    )
    data['% Difference'] = data['% of Forecast Consumed'] - data['% of Month Passed']


    def safe_sheet_name(base_name, suffix):
        base_name = str(base_name)
        if base_name.lower() == 'nan':
            base_name = "UnknownBU"
        safe_base = base_name[:31 - len(suffix) - 1]
        return f"{safe_base}_{suffix}"


    print("Writing Full Sales Report Sheet...")
    data.to_excel(writer, sheet_name="Sales Report Generated", index=False)

    workbook = writer.book
    percent_format_no_decimals = workbook.add_format({'num_format': '0%'})  # Format without decimals
    number_format = workbook.add_format({'num_format': '#,##0'})  # Format for rounded numbers
    header_format = workbook.add_format({'bold': True, 'bg_color': '#FFC000', 'border': 1, 'align': 'center'})
    headers_to_format = ['Total', 'Result', 'Forecast Consumed (Expected)', '% of Forecast', 'Status',
                         '% of Month Passed', '% of Forecast Consumed', '% Difference']

    main_sheet = writer.sheets['Sales Report Generated']
    for col_name in headers_to_format:
        col_index = data.columns.get_loc(col_name)
        if col_name in ['Total', 'Result', 'Forecast Consumed (Expected)']:
            main_sheet.set_column(col_index, col_index, 12, number_format)  # Format as rounded numbers
        elif col_name in ['% of Forecast', '% of Month Passed', '% of Forecast Consumed', '% Difference']:
            main_sheet.set_column(col_index, col_index, 12,
                                  percent_format_no_decimals)  # Format as percentages without decimals

    for col_num, col_name in enumerate(data.columns):
        header = "Result (Over FC or Under FC)" if col_name == "Result" else col_name
        if col_name in headers_to_format:
            main_sheet.write(0, col_num, header, header_format)

    # Add conditional formatting for negative values in the 'Result' column
    result_col_index = data.columns.get_loc('Result')
    main_sheet.conditional_format(1, result_col_index, len(data), result_col_index, {
        'type': 'cell',
        'criteria': '<',
        'value': 0,
        'format': workbook.add_format({'font_color': 'red'})
    })

    report_date = datetime.now().strftime("%d-%b-%Y")

    for bu_group in data['BU_GROUPED'].unique():
        print(f"Processing Business Unit: {bu_group}...")
        bu_data = data[data['BU_GROUPED'] == bu_group]
        # Filter Top 10 Overselling and Underselling
        overselling = bu_data[bu_data['Status'] == 'Overselling'].sort_values(by='Result', ascending=False).head(10)
        overselling['Result_Abs'] = overselling['Result'].abs()
        overselling_renamed = overselling.rename(columns={"Result": "Over FC"})

        underselling = bu_data[bu_data['Status'] == 'Underselling'].sort_values(by='Result', ascending=True).head(10)
        underselling['Result_Abs'] = underselling['Result'].abs()
        underselling_renamed = underselling.rename(columns={"Result": "Under FC"})

        # Sheet Setup
        sheet_name = safe_sheet_name(bu_group, "Report")
        sheet = writer.book.add_worksheet(sheet_name)

        sheet.write('A1', "Sales Performance Report")
        sheet.write('A2', f"Business Unit (BU): {bu_group}")
        sheet.write('A3', f"Generated On: {report_date}")
        sheet.write('A4', "Top 10 Overselling and Underselling SKUs")

        # ==============================
        # OVERSSELLING SECTION
        # ==============================
        print(f"Generating Overselling Table for {bu_group}...")
        overselling_renamed.to_excel(writer, sheet_name=sheet_name, startrow=5, index=False)

        for col_num, col_name in enumerate(overselling_renamed.columns):
            if col_name == "Over FC":
                sheet.write(5, col_num, col_name, header_format)  # same yellowish header
            elif col_name in headers_to_format:
                sheet.write(5, col_num, col_name, header_format)

        if not overselling.empty:
            print(f"Creating Overselling Chart for {bu_group}...")
            chart1 = workbook.add_chart({'type': 'column'})
            chart1.add_series({
                'name': 'Overselling',
                'categories': [sheet_name, 6, overselling_renamed.columns.get_loc('SKU'),
                               6 + len(overselling_renamed) - 1, overselling_renamed.columns.get_loc('SKU')],
                'values': [sheet_name, 6, overselling_renamed.columns.get_loc('Over FC'),
                           6 + len(overselling_renamed) - 1, overselling_renamed.columns.get_loc('Over FC')],
            })
            chart1.set_title({'name': 'Top 10 Overselling by Result'})
            chart1.set_y_axis({'name': 'Result'})
            sheet.insert_chart(f'L{6 + len(overselling_renamed) + 2}', chart1)

        # ==============================
        # UNDERSSELLING SECTION
        # ==============================
        underselling_start_row = 6 + len(overselling_renamed) + 15 + 2

        print(f"Generating Underselling Table for {bu_group}...")
        underselling_renamed.to_excel(writer, sheet_name=sheet_name, startrow=underselling_start_row, index=False)

        for col_num, col_name in enumerate(underselling_renamed.columns):
            if col_name == "Under FC":
                sheet.write(underselling_start_row, col_num, col_name, header_format)
            elif col_name in headers_to_format:
                sheet.write(underselling_start_row, col_num, col_name, header_format)

        if not underselling.empty:
            print(f"Creating Underselling Chart for {bu_group}...")
            chart2 = workbook.add_chart({'type': 'column'})
            chart2.add_series({
                'name': 'Underselling',
                'categories': [sheet_name, underselling_start_row + 1, underselling_renamed.columns.get_loc('SKU'),
                               underselling_start_row + len(underselling_renamed),
                               underselling_renamed.columns.get_loc('SKU')],
                'values': [sheet_name, underselling_start_row + 1, underselling_renamed.columns.get_loc('Result_Abs'),
                           underselling_start_row + len(underselling_renamed),
                           underselling_renamed.columns.get_loc('Result_Abs')],
            })
            chart2.set_title({'name': 'Top 10 Underselling by Result'})
            chart2.set_y_axis({'max': underselling['Result_Abs'].max() * 1.2, 'name': 'Absolute Result'})
            chart2.set_x_axis({'name': 'SKU', 'label_position': 'low', 'label_rotation': 45})
            sheet.insert_chart(f'L{underselling_start_row + len(underselling_renamed) + 3}', chart2)
            if "Under FC" in underselling_renamed.columns:
                under_fc_col_index = underselling_renamed.columns.get_loc("Under FC")
                sheet.conditional_format(underselling_start_row + 1, under_fc_col_index,
                                         underselling_start_row + len(underselling_renamed), under_fc_col_index, {
                                             'type': 'cell',
                                             'criteria': '<',
                                             'value': 0,
                                             'format': workbook.add_format({'font_color': 'red'})
                                         })

        # Apply number/percent formatting to renamed sheets
        for col_name in ['Forecast Consumed (Expected)']:
            if col_name in bu_data.columns:
                col_index = bu_data.columns.get_loc(col_name)
                sheet.set_column(col_index, col_index, 12, number_format)

        for col_name in ['% of Forecast', '% of Month Passed', '% of Forecast Consumed', '% Difference']:
            if col_name in bu_data.columns:
                col_index = bu_data.columns.get_loc(col_name)
                sheet.set_column(col_index, col_index, 12, percent_format_no_decimals)

print("Sales Report Successfully Generated!")