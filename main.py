import pandas as pd
import re
import cProfile
import pstats
import io
import matplotlib
import matplotlib.pyplot as plt
import xlsxwriter

matplotlib.use('Agg')


def sanitize_sheet_name(name):
    """Sanitize sheet names by removing invalid characters and trimming to 25 characters.

    Args:
        name (str): The original sheet name.

    Returns:
        str: The sanitized sheet name.
    """
    sanitized = re.sub(r'[\\/*?[\]:]', '', name).strip()
    return sanitized[:25]


def auto_adjust_column_widths(df, worksheet):
    """Adjust the width of the columns in the Excel worksheet.

    Args:
        df (pd.DataFrame): The DataFrame containing the data.
        worksheet (xlsxwriter.Worksheet): The worksheet to adjust.
    """
    for idx, col in enumerate(df.columns):
        max_length = max(df[col].astype(str).map(len).max(), len(col)) + 2  # Adding a little extra space
        worksheet.set_column(idx, idx, max_length)


def create_bar_chart(data, title, x_label, y_label):
    """Create a bar chart using Matplotlib and return it as an in-memory image.

    Args:
        data (pd.Series): The data to plot.
        title (str): The title of the chart.
        x_label (str): The label for the x-axis.
        y_label (str): The label for the y-axis.

    Returns:
        io.BytesIO: The in-memory image of the chart.
    """
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(data.index, data.values)
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    plt.xticks(rotation=30, ha='right')
    plt.tight_layout()

    image_stream = io.BytesIO()
    plt.savefig(image_stream, format='png')
    plt.close(fig)
    image_stream.seek(0)
    return image_stream


def process_faults_file():
    """Process the faults CSV file and create an Excel file with separate sheets for each machine."""
    file_name = 'faults.csv'
    df = pd.read_csv(file_name)

    # Convert 'T_TotalDuration' and 'T_TotalOccur' to numeric, forcing errors to NaN
    df['T_TotalDuration'] = pd.to_numeric(df['T_TotalDuration'], errors='coerce')
    df['T_TotalOccur'] = pd.to_numeric(df['T_TotalOccur'], errors='coerce')

    # Drop rows with NaN values in 'T_TotalDuration' or 'T_TotalOccur'
    df = df.dropna(subset=['T_TotalDuration', 'T_TotalOccur'])

    # Ensure 'D_MachineName', 'D_MsgDesc', and 'D_MsgCode' are strings and fill NaN values with a placeholder
    df['D_MachineName'] = df['D_MachineName'].astype(str).fillna('Unknown')
    df['D_StateDesc'] = df['D_StateDesc'].astype(str).fillna('Unknown')
    df['D_MsgDesc'] = df['D_MsgDesc'].astype(str).fillna('Unknown')
    df['D_MsgCode'] = df['D_MsgCode'].astype(str).fillna('Unknown')

    # Combine 'D_MsgDesc' and 'D_MsgCode' for labeling
    df['Fault_Description'] = df['D_MsgDesc'] + " (" + df['D_MsgCode'] + ")"

    # Aggregate data by 'D_MsgCode' and sum 'T_TotalDuration' and 'T_TotalOccur'
    aggregated_df_total_duration = df.groupby(
        ['D_MachineName', 'D_StateDesc', 'D_MsgCode', 'D_MsgDesc', 'Fault_Description']
    ).agg({'T_TotalDuration': 'sum'}).reset_index()

    aggregated_df_total_occur = df.groupby(
        ['D_MachineName', 'D_StateDesc', 'D_MsgCode', 'D_MsgDesc', 'Fault_Description']
    ).agg({'T_TotalOccur': 'sum'}).reset_index()

    # Calculate the number of unique faults per station
    unique_faults_per_station = df.groupby('D_MachineName').size().reset_index(name='Unique Faults')
    sorted_unique_faults_per_station = unique_faults_per_station.sort_values(by='Unique Faults', ascending=False)

    # Aggregate total duration per station
    total_duration_per_station = aggregated_df_total_duration.groupby('D_MachineName')['T_TotalDuration'].sum().reset_index()
    sorted_total_duration_per_station = total_duration_per_station.sort_values(by='T_TotalDuration', ascending=False)

    # Create an Excel writer object
    excel_file = 'faults_per_machine.xlsx'

    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        # Initialize list for index sheet
        index_data = []

        # Get unique machine names
        unique_machines = aggregated_df_total_duration['D_MachineName'].unique()
        
        # Write data for each machine to a separate sheet
        for machine in unique_machines:
            machine_df_duration = aggregated_df_total_duration[aggregated_df_total_duration['D_MachineName'] == machine].sort_values(
                by='T_TotalDuration', ascending=False
            ).head(10)  # Grab top 10 by duration
            machine_df_occurrences = aggregated_df_total_occur[aggregated_df_total_occur['D_MachineName'] == machine].sort_values(
                by='T_TotalOccur', ascending=False
            ).head(10)  # Grab top 10 by occurrences
            
            # Sanitize sheet name
            sheet_name = sanitize_sheet_name(machine)
            
            # Write duration data to the sheet
            machine_df_duration.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False, header=True)
            
            # Write occurrences data to the sheet below the duration data
            startrow = len(machine_df_duration) + 4
            machine_df_occurrences.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=True)
            
            # Add titles to the tables
            worksheet = writer.sheets[sheet_name]
            worksheet.write(0, 0, 'Total Duration (Top 10)')
            worksheet.write(startrow-1, 0, 'Total Occurrences (Top 10)')
            
            # Auto-adjust column widths
            auto_adjust_column_widths(machine_df_duration, worksheet)
            auto_adjust_column_widths(machine_df_occurrences, worksheet)
            
            # Create bar charts with Matplotlib and insert as images
            duration_image = create_bar_chart(
                machine_df_duration.set_index('Fault_Description')['T_TotalDuration'],
                f'Top 10 Durations for {sheet_name}', 'Fault Description', 'Total Duration'
            )
            occurrences_image = create_bar_chart(
                machine_df_occurrences.set_index('Fault_Description')['T_TotalOccur'],
                f'Top 10 Occurrences for {sheet_name}', 'Fault Description', 'Total Occurrences'
            )
            
            # Insert images into the worksheet
            worksheet.insert_image(20, 0, f'{sheet_name}_duration.png', {'image_data': duration_image, 'x_scale': 1, 'y_scale': 1})
            worksheet.insert_image(20, 10, f'{sheet_name}_occurrences.png', {'image_data': occurrences_image, 'x_scale': 1, 'y_scale': 1})

            # Set sheet zoom level to 50%
            worksheet.set_zoom(50)
            
            # Append to index data
            index_data.append([machine, sheet_name])

        # Create index DataFrame
        index_df = pd.DataFrame(index_data, columns=['Machine Name', 'Sheet Name'])

        # Write index sheet
        index_df.to_excel(writer, sheet_name='Index', index=False)
        worksheet = writer.sheets['Index']
        
        # Create hyperlinks
        for i, sheet in enumerate(index_df['Sheet Name']):
            worksheet.write_url(i + 1, 1, f"internal:'{sheet}'!A1", string=sheet)
            worksheet.write(i + 1, 0, index_df.at[i, 'Machine Name'])
        
        # Auto-adjust column widths for index sheet
        auto_adjust_column_widths(index_df, worksheet)

        # Write sorted unique faults per station sheet
        sorted_unique_faults_per_station.to_excel(writer, sheet_name='Unique Faults', index=False)
        worksheet = writer.sheets['Unique Faults']
        auto_adjust_column_widths(sorted_unique_faults_per_station, worksheet)

        # Write sorted total duration per station sheet
        sorted_total_duration_per_station.to_excel(writer, sheet_name='Total Duration', index=False)
        worksheet = writer.sheets['Total Duration']
        auto_adjust_column_widths(sorted_total_duration_per_station, worksheet)

        # Set zoom level to 50% for summary sheets
        worksheet = writer.sheets['Index']
        worksheet.set_zoom(50)
        worksheet = writer.sheets['Unique Faults']
        worksheet.set_zoom(50)
        worksheet = writer.sheets['Total Duration']
        worksheet.set_zoom(50)

    # Save the workbook and reopen to set the index sheet as the first sheet
    workbook = writer.book
    workbook.filename = excel_file
    workbook.close()

    # Reopen the workbook and set the index sheet as the first sheet
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
        writer.book.active = writer.book.sheetnames.index('Index')

    print(f"Excel file with separate sheets for each machine, an index sheet, a sorted unique faults sheet, and a sorted total duration sheet has been saved as '{excel_file}'.")


if __name__ == "__main__":
    # Profile the function
    pr = cProfile.Profile()
    pr.enable()

    process_faults_file()

    pr.disable()
    s = io.StringIO()
    ps = pstats.Stats(pr, stream=s).sort_stats('cumulative')
    ps.print_stats()

    print(s.getvalue())
