import streamlit as st
import pandas as pd
import smartsheet

from datetime import datetime





@st.cache_data
def smartsheet_to_dataframe(sheet_id):
    smartsheet_client = smartsheet.Smartsheet(st.secrets['smartsheet']['access_token'])
    sheet             = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns           = [col.title for col in sheet.columns]
    rows              = []
    for row in sheet.rows: rows.append([cell.value for cell in row.cells])
    return pd.DataFrame(rows, columns=columns)





def get_pod_from_amenity_string(row):

    if 'POD' in row['Amenity_Notes']:
        notes = row['Amenity_Notes'].split(' ')
        return notes[-2]
    
    return None





APP_NAME = 'Collaboration Companion'

st.set_page_config(page_title=APP_NAME, page_icon='🤝', layout='wide')

st.image(st.secrets['images']["rr_logo"], width=100)

st.title('🤝 ' + APP_NAME)
st.info('Highlight collaboration opportunities based on property arrivals.')

auth = st.query_params.get('auth')

if auth == st.secrets['auth']['key']:
    current_year = datetime.now().year
    prior_year   = current_year - 1
    next_year    = current_year + 1
    report_url   = f"{st.secrets['escapia_1']}{prior_year}{st.secrets['escapia_2']}{next_year}{st.secrets['escapia_3']}"

    st.link_button('Download the **Housekeeping Report** from **Escapia**', url=report_url, type='secondary', use_container_width=True, help='Housekeeping Arrival Departure Report - Excel 1 line')


c1, c2          = st.columns(2)

start_date      = c1.date_input('Date range start', value=datetime.now())
end_date        = c2.date_input('Date range end', value=datetime.now())

c1, c2, c3      = st.columns(3)

escapia_file    = c1.file_uploader(label='**Housekeeping Arrival Departure Report - Excel 1 line.csv**\n\n Download button above', type='csv')
breezeway_file  = c2.file_uploader(label='**breezeway-task-custom-export.csv**\n\nBreezeway > Tasks > All Tasks > All tasks (collab ops reporting)', type='csv')
pod_file        = c3.file_uploader(label='**Amenity String Report.csv**\n\nEscapia > Units > Reports > Amenity String Report > (Select all) > CSV', type='csv')


if escapia_file and breezeway_file and pod_file:
    gdf = smartsheet_to_dataframe(st.secrets['smartsheet']['sheet_id']['geo'])
    
    edf = pd.read_csv(escapia_file)
    bdf = pd.read_csv(breezeway_file)
    pdf = pd.read_csv(pod_file)

    pdf['Amenity_Notes'] = pdf['Amenity_Notes'].fillna('')
    pdf['Amenity_Notes'] = pdf['Amenity_Notes'].str.upper()
    pdf['Pod']           = pdf.apply(get_pod_from_amenity_string, axis=1)
    pdf                  = pdf[['Unit_Code','Pod']]

    edf['Start_Date'] = pd.to_datetime(edf['Start_Date'])
    edf['PreviousBookingDate'] = pd.to_datetime(edf['PreviousBookingDate'])
    edf['Previous_Departure_Date'] = edf['PreviousBookingDate'] + pd.Timedelta(days=1)
    edf = edf[['Unit_Code','Reservation_Number','ReservationTypeDescription','Start_Date','Previous_Departure_Date']]

    def determine_b2b(row):
        if pd.isna(row['Previous_Departure_Date']):
            return False
        return row['Start_Date'] == row['Previous_Departure_Date']
    
    edf['isB2B'] = edf.apply(determine_b2b, axis=1)

    edf = edf[(edf['Start_Date'] >= pd.to_datetime(start_date)) & (edf['Start_Date'] <= pd.to_datetime(end_date))]
    edf = edf.sort_values(by='Start_Date', ascending=True)
    edf = edf.drop_duplicates(subset=['Reservation_Number'], keep='last')

    bdf = bdf[['Task ID', 'Task title', 'Property marketing ID', 'Due date']]
    bdf['Due date'] = pd.to_datetime(bdf['Due date'])
    bdf = bdf[~bdf['Task title'].str.startswith(tuple(['1.', '2.', '3.', '4.', '5.', '6.']))]

    df = pd.merge(edf, bdf, left_on=['Unit_Code'], right_on=['Property marketing ID'], how='left')
    df = df[(df['Due date'] <= df['Start_Date'])]
    df = df.drop(columns=['Property marketing ID','Reservation_Number'])
    df = df[df['Task ID'].notna()]

    gdf = gdf[['Unit_Code','Address','Order']]

    df = pd.merge(df, gdf, on='Unit_Code', how='left')
    df = df.merge(pdf, on='Unit_Code', how='left')
    df = df.sort_values(by=['Start_Date','Order'], ascending=True)
    
    def print_results(df, pod_filter=None):

        df = df[df['Pod'] == pod_filter] if pod_filter and pod_filter != 'All' else df

        date = ''
        unit = ''

        for _, row in df.iterrows():
            if row['Start_Date'] != date:
                st.title(f'🗓️ {row['Start_Date'].strftime('%A, %B %d, %Y')}')
                date = row['Start_Date']
            
            if row['Unit_Code'] != unit:
                if 'Renter' in row['ReservationTypeDescription']:
                    if row['isB2B']:
                        st.header(f'🏠 {row['Unit_Code']} - {row['Address']} - {row['ReservationTypeDescription']} - {row['Start_Date'].strftime('%m/%d')} - B2B 🔁')
                    else:
                        st.header(f'🏠 {row['Unit_Code']} - {row['Address']} - {row['ReservationTypeDescription']} - {row['Start_Date'].strftime('%m/%d')}')
                elif 'Guest of Owner' in row['ReservationTypeDescription']:
                    if row['isB2B']:
                        st.header(f'🏡 {row['Unit_Code']} - {row['Address']} - {row['ReservationTypeDescription']} - {row['Start_Date'].strftime('%m/%d')} - B2B 🔁')
                    else:
                        st.header(f'🏡 {row['Unit_Code']} - {row['Address']} - {row['ReservationTypeDescription']} - {row['Start_Date'].strftime('%m/%d')}')
                else:
                    if row['isB2B']:
                        st.header(f'🏰 {row['Unit_Code']} - {row['Address']} - {row['ReservationTypeDescription']} - {row['Start_Date'].strftime('%m/%d')} - B2B 🔁')
                    else:
                        st.header(f'🏰 {row['Unit_Code']} - {row['Address']} - {row['ReservationTypeDescription']} - {row['Start_Date'].strftime('%m/%d')}')

                st.write(f'')
                unit = row['Unit_Code']

            st.write(f'> {int(row['Task ID'])} - {row['Task title']} - {row['Due date'].strftime("%m/%d")}')

    st.success('Opportunities are sorted by **arrival date**, then by **geographical order**.', icon='🥇')
    st.info('Only tasks with due dates **on or before** each property\'s **arrival date** are included.', icon='🗓️')

    st.selectbox('Filter by POD', options=['All'] + sorted(df['Pod'].unique()), key='pod_filter')
    with st.expander('Opportunities', expanded=False):
        print_results(df, pod_filter=st.session_state.get('pod_filter'))