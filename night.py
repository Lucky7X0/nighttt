import pandas as pd
import datetime
import streamlit as st
from io import BytesIO

def load_data(file):
    return pd.read_excel(file, sheet_name=None)  # Load all sheets

def parse_datetime(date_str, time_str):
    return datetime.datetime.strptime(f"{date_str} {time_str}", "%d/%m/%Y %H:%M:%S")

def filter_data_for_day(data, shift_date):
    shift_start_datetime = parse_datetime(shift_date, '17:00:00')
    shift_end_datetime = shift_start_datetime + datetime.timedelta(hours=11)  # Ends at 4 AM the next day

    data['DateTime'] = pd.to_datetime(data['Date'] + ' ' + data['Punch Time'], format="%d/%m/%Y %H:%M:%S")
    
    # Filter data for the night shift (from 5 PM of shift_date to 4 AM the next day)
    filtered_data = data[
        (data['DateTime'] >= shift_start_datetime) & 
        (data['DateTime'] <= shift_end_datetime)
    ]
    
    return filtered_data

def calculate_night_shift(data):
    total_login_logout_time = datetime.timedelta()
    total_break_time = datetime.timedelta()

    first_login = None
    last_logout = None
    in_time = None
    prev_out_time = None

    for _, row in data.iterrows():
        current_datetime = row['DateTime']

        if row['I/O Type'] == 'IN':
            if first_login is None:
                first_login = current_datetime
            in_time = current_datetime
            if prev_out_time and prev_out_time < in_time:
                total_break_time += in_time - prev_out_time
        elif row['I/O Type'] == 'OUT' and in_time:
            last_logout = current_datetime
            prev_out_time = current_datetime

    if first_login and last_logout:
        total_login_logout_time = last_logout - first_login

    def timedelta_to_hours_minutes(td):
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours} hours, {minutes} minutes"

    results = {
        'Total Time from Login to Logout (including breaks)': timedelta_to_hours_minutes(total_login_logout_time),
        'Total Break Time': timedelta_to_hours_minutes(total_break_time),
        'Total Hours Worked (excluding breaks)': timedelta_to_hours_minutes(total_login_logout_time - total_break_time),
        'Break Time (Minutes)': total_break_time.total_seconds() / 60  # For filtering purposes
    }

    return results

def process_all_sheets(file):
    sheets = pd.read_excel(file, sheet_name=None)
    results_dict = {}
    
    for sheet_name, data in sheets.items():
        st.write(f"Processing sheet: {sheet_name}")
        results = []
        
        # Assuming the data has a 'Date' column with daily dates
        working_dates = data['Date'].unique()
        for shift_date in working_dates:
            filtered_data = filter_data_for_day(data, shift_date)
            if not filtered_data.empty:
                result = calculate_night_shift(filtered_data)
                # Check if the results are meaningful (not zeroed out)
                if any(value != '0 hours, 0 minutes' for value in result.values()):
                    # Convert shift_date to datetime to handle next day's checking
                    shift_date_datetime = datetime.datetime.strptime(shift_date, '%d/%m/%Y')
                    next_day = shift_date_datetime + datetime.timedelta(days=1)
                    next_day_str = next_day.strftime('%d/%m/%Y')

                    # Check if next day is a working day (i.e., if there is data for the next day)
                    if next_day_str in working_dates:
                        results.append({
                            'Working Day (Login Date)': shift_date,
                            **result
                        })
        
        if results:
            results_df = pd.DataFrame(results)
            # Filter records where break time exceeded 60 minutes
            results_df = results_df[results_df['Break Time (Minutes)'] > 60]
            results_dict[sheet_name] = results_df
    
    with BytesIO() as buffer:
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for sheet_name, results_df in results_dict.items():
                results_df.to_excel(writer, sheet_name=f'Results_{sheet_name}', index=False)

        buffer.seek(0)
        return buffer.getvalue()

def main():
    st.title("Night Shift Hours Calculator")

    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        if st.button('Process File'):
            try:
                # Process all sheets in the uploaded file and get the new Excel file as bytes
                new_file_bytes = process_all_sheets(uploaded_file)

                # Provide the new Excel file for download
                st.download_button(
                    label="Download Processed File",
                    data=new_file_bytes,
                    file_name="Processed_Shift_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Load and display the new file in Streamlit
                with pd.ExcelFile(BytesIO(new_file_bytes)) as xls:
                    for sheet_name in xls.sheet_names:
                        st.write(f"Sheet: {sheet_name}")
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        st.dataframe(df)
            
            except Exception as e:
                st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
