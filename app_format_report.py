import pandas as pd
import pytz
import streamlit as st
from datetime import datetime
from io import StringIO

# Constants
TELEMARKETER_CODE = 'INHS'
ORIGINAL_OR_RERUN = 'O'
MEDIA_CLIENT_CODE = 'RAWEDH'
RESPONSE_COUNTER_FIELD = '000001'
PHONE_NUMBER_MAPPING = {
    'Radio 1': '8009096617',
    'Radio 2': '8009097413',
    'Radio 3': '8006196330',
    'Radio 4': '8006197121',
    'Radio 5': '8009171020',
    'Radio 6': '8007819600',
    'Radio 7': '8007838500',
    'Radio 8': '8009171011',
    'Radio 9': '8003072010',
    'Radio 10': '8006832030'
}

# Timezone conversion
PDT = pytz.timezone('America/Los_Angeles')
EST = pytz.timezone('America/New_York')

# Helper function to format the date and time
def format_datetime(pdt_datetime):
    est_datetime = pdt_datetime.astimezone(EST)
    formatted_date = est_datetime.strftime('%Y%m%d')
    formatted_time = est_datetime.strftime('%H%M')
    return formatted_date, formatted_time

# Helper function to determine the response code
def get_response_code(tags):
    if pd.isna(tags):
        return 'VCAL'
    tags_lower = tags.lower()
    if any(keyword in tags_lower for keyword in ['junk', 'missed call', 'test', 'wrong number']):
        return 'CALL'
    return 'VCAL'

# Streamlit app
def main():
    st.title("Excel to TXT File Processor")

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file is not None:
        # Read the Excel file and 'Calls' sheet
        calls_df = pd.read_excel(uploaded_file, sheet_name='Calls')

        # StringIO to store output for downloading
        output_txt = StringIO()

        # Process the file and write output to StringIO
        for _, row in calls_df.iterrows():
            # Telemarketer Code
            telemarketer_code = TELEMARKETER_CODE
            
            # Original or Rerun
            original_or_rerun = ORIGINAL_OR_RERUN
            
            # 12 spaces
            spaces_12 = ' ' * 12
            
            # Media + Client Code
            media_client_code = MEDIA_CLIENT_CODE
            
            # 24 spaces
            spaces_24 = ' ' * 24
            
            # Date and Time of call (Convert PDT to EST)
            start_time_pdt = row['Start Time']
            if pd.isna(start_time_pdt):
                continue  # skip rows with missing 'Start Time'
            date_str, time_str = format_datetime(PDT.localize(start_time_pdt))
            
            # Response Code
            response_code = get_response_code(row.get('Tags', ''))
            
            # Response Counter Field
            response_counter_field = RESPONSE_COUNTER_FIELD
            
            # Phone number based on 'Number Name'
            radio_station = row.get('Number Name', '')
            phone_number = PHONE_NUMBER_MAPPING.get(radio_station, 'Unknown')
            if phone_number == 'Unknown':
                continue  # skip rows with unknown 'Number Name'
            
            # Zip code of caller (ignored, so set to 5 spaces)
            zip_code = ' ' * 5
            
            # Area code of caller (first 3 digits of the phone number)
            area_code = row.get('Phone Number', '')[:3]
            
            # 21 spaces
            spaces_21 = ' ' * 21
            
            # Create the formatted line
            formatted_line = (
                f"{telemarketer_code}"
                f"{original_or_rerun}"
                f"{spaces_12}"
                f"{media_client_code}"
                f"{spaces_24}"
                f"{date_str}"
                f"{time_str}"
                f"{response_code}"
                f"{response_counter_field}"
                f"{phone_number}"
                f"{zip_code}"
                f"{area_code}"
                f"{spaces_21}\n"
            )
            
            # Write the formatted line to StringIO
            output_txt.write(formatted_line)
        
        # Convert StringIO to downloadable content
        output_txt.seek(0)
        txt_data = output_txt.getvalue()

        # Generate the output file name using the current date
        output_file_name = f"WED_{date_str}.txt"

        # Provide download button for the generated txt file
        st.download_button(
            label="Download TXT file",
            data=txt_data,
            file_name=output_file_name,
            mime="text/plain"
        )

if __name__ == "__main__":
    main()
