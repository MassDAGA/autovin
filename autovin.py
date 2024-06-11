import streamlit as st
import pandas as pd
import requests
import os
import openpyxl
import numpy as np
import json
from datetime import datetime
from io import BytesIO

@st.cache_data
def confirm_vin(file_path):
    
    #some excel files have more than 1 sheet, we handle excel files with more than 1 sheet by telling the 
    #code to read the sheet named 'Vehicle & Asset List' as this is the standard naming convention
    #write the information from this sheet into dataframe named 'raw_vin_data'
    wb = openpyxl.load_workbook(file_path)
    res = len(wb.sheetnames)
    if res > 1:
        raw_vin_data = pd.read_excel(file_path, 'Vehicle & Asset List', header = 3) #turns excel into pandas dataframe
    else:
        raw_vin_data = pd.read_excel(file_path, header = 3)
        
    #assign new column names to raw_vin_data dataframe for dataframe to standardize raw_vin_data for query
    for column in raw_vin_data.columns:
        if 'vehicle asset name' in column.lower():
            raw_vin_data.rename(columns={column:'Vehicle Asset Name'}, inplace=True)
        elif 'model year' in column.lower():
            raw_vin_data.rename(columns={column:'Model Year'}, inplace=True)
        elif 'make' in column.lower():
            raw_vin_data.rename(columns={column:'Make'}, inplace=True)
        elif 'model' in column.lower():
            raw_vin_data.rename(columns={column:'Model'}, inplace=True)
        elif 'vin' in column.lower():
            raw_vin_data.rename(columns={column:'VIN'}, inplace=True)
        elif 'fuel type' in column.lower():
            raw_vin_data.rename(columns={column:'Fuel Type'}, inplace=True)
    
    #create base url that will be augmented with VIN for query
    base_url = 'https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVin/'
    
    #create a dataframe for CAN compatability check in standard format with required columns
    vin_data = pd.DataFrame({'VRN':[], 'VIN': [], 'YEAR': [], 
                             'MAKE': [], 'MODEL': [], 'FUEL': [], 'COUNTRY': []})
    
    
    #write relevant info into vin_data dataframe using raw data from original sales document, only includes 
    #info from rows where vin has been entered, excludes NULL/NAN values
    for ind in raw_vin_data.index:
        if pd.isna(raw_vin_data['VIN'][ind]) == False:
            vin_data.loc[ind] = [raw_vin_data['Vehicle Asset Name'][ind],
                             raw_vin_data['VIN'][ind], raw_vin_data['Model Year'][ind],
                             raw_vin_data['Make'][ind], raw_vin_data['Model'][ind],
                             raw_vin_data['Fuel Type'][ind],
                             'US']
    
    #reset the vin dataframe index, index now begins at 0
    vin_data.reset_index(drop = True, inplace = True)
    
    #replace NAN/NULL values indicating empty cell with an empty string
    vin_data.replace(np.nan, '', inplace = True)
    
    #change the values in vin_data dataframe to strings, this is necessary for later string concatenation
    vin_data = vin_data.astype(str)
    
   #create list to store dictionaries, each dictionary will relate to a specific VIN or row of the dataframe
    results = []
    
    #extract VINs from vin_data dataframe into a list of values
    values = vin_data['VIN'].values.tolist()
    
    #create variable to keep track of which index is being used, this keeps track of what row of the dataframe
    #the code is on
    ind = 0
    
    #query the NHTSA VIN database using each VIN from the original sales document to collect info on vehicle 
    #year, make, model, fuel, and vehicle type, as MCF operates in United States all entries for Country = US
    
    #iterate through each VIN in list of VINs
    for value in values:
        #ensure the type of the VIN is string
        value = str(value)
        #remove spaces from VIN, accounts for common data entry error
        value = value.replace(" ", "")
        #create VIN specific link to access details for API query
        url = base_url + value + '?format=json'
        #pulls details from url, bypasses certification verification error created by Michelin firewalls
        response = requests.get(url, verify = False)
        #check to see if vin is accurate, if accurate extract data into dictionary and add to results list
        try:
            #save url information as data variable for query
            data = response.json()
            #create key for decoding desired information from url data
            decoded_values = {item['Variable']: item['Value'] for item in data['Results']}
            #create a dictionary with vehicle information from VIN query, information based on specific VIN
            results.append({

                'VRN': vin_data['VRN'][ind],
                'VIN': value,
                'YEAR': decoded_values.get('Model Year', 'N/A'),
                'MAKE': decoded_values.get('Make', 'N/A'),
                'MODEL': decoded_values.get('Model', 'N/A'),
                'FUEL': decoded_values.get('Fuel Type - Primary', 'N/A'),
                'COUNTRY': 'US',
                'VEHICLE TYPE': decoded_values.get('Vehicle Type', 'N/A'),
                'ERROR CODE': decoded_values.get('Error Text', 'N/A')
            })
            #increase the index by 1, indicates code moves onto next VIN/row
            ind += 1
        #if vin not accurate, use error handling, will only move to this step if url produces empty response 
        #(data variable is empty)
        except json.JSONDecodeError as e:
            results.append({
                'VRN': vin_data['VRN'][ind],
                'VIN': value,
                'YEAR': 'Error',
                'MAKE': 'Error',
                'MODEL': 'Error',
                'FUEL': 'Error',
                'COUNTRY': 'Error',
                'VEHICLE TYPE': 'Error',
                'ERROR CODE': 'Error: No information found for input VIN'
            })
            #increase the index by 1, indicates code moves onto next VIN/row
            ind += 1
        #if code times out, this error handling will make sure the code does not run indefinitely, if
        #encountered the code will stop processing VINs and communicate a time out error to the user
        except requests.exceptions.Timeout as e:
            return "Timed out"
            
    #create dataframe from list of dictionaries, each dictionary is a row within the 'results' dataframe
    results = pd.DataFrame(results)
    
    #create valid_vins dataframe that will be fed into CAN compatability check, exclude trailers, lifts
    #and invalid VINs, exclude all rows where primary fuel type is N/A, Error or None, such results indicate
    #valid vehicles require an energy source
    valid_vins = results[~results.FUEL.isin(['Not Applicable', 'Error', None])]
    
    #exclude vehicles that are older than 30 years from valid_vins dataframe, these require manual check
    #due to manufacturers repeating VINs
    valid_vins = valid_vins[datetime.now().year - valid_vins['YEAR'].astype(int) < 30]
    
    #remove VEHICLE TYPE and ERROR CODE column from valid_vins dataframe to ensure in correct format for 
    #CAN compatability
    valid_vins.drop(['VEHICLE TYPE', 'ERROR CODE'], axis = 1, inplace = True)
    
    #remove duplicate VINs from CAN compatability check document
    valid_vins.drop_duplicates(subset=['VIN'],inplace= True)
    
    #create a list indicating if vehicles need to be checked manually, this will become a column in the
    #processed VINs doccument exported to the employee
    check_list = []
    
    #create a record of checked VINs, this is used later to check if there are duplicate VINs
    vins_checked = []
    
    #create a list of valid VINs appearing in the CAN compatable dataframe, these do not need to be manually
    #checked
    valid_vin_list = valid_vins['VIN'].values.tolist()
    
    #add vehicle type information to vin_data dataframe for employee reference, added to help with manual checks
    vin_data = pd.concat([vin_data, results['VEHICLE TYPE']], axis = 1)
    
    #determine if a manual check of a given vehicle vin is necessary
    #iterate through vin_data dataframe by index
    for ind in vin_data.index:
        #check if the VIN is already in CAN dataframe and not a duplicate, if true no manual check
        if vin_data['VIN'][ind].replace(" ", "") in valid_vin_list and vin_data['VIN'][ind] not in vins_checked:
            check_list.append('NO')
        #if the vehicle is a trailer no manual checl
        elif vin_data['VEHICLE TYPE'][ind] == 'TRAILER':
            check_list.append('NO')
        elif 'trailer' in vin_data['MODEL'][ind].lower() or 'trailer' in vin_data['VRN'][ind].lower():
            check_list.append('NO')
        #if vehicle is a lift, no manual check
        elif 'lift' in vin_data['MODEL'][ind].lower() or 'lift' in vin_data['VRN'][ind].lower():
            check_list.append('NO')
        elif 'example' in vin_data['VIN'][ind].lower():
            check_list.append('NO')
        #if VIN is duplicate manual check is necessary
        elif vin_data['VIN'][ind] in vins_checked:
            check_list.append('YES: Duplicate Vin')
        #otherwise a manual check is necessary
        else:
            check_list.append('YES')
        #add VIN to vins_checked list
        vins_checked.append(vin_data['VIN'][ind])
        
    #update vehicle type to indicate the vehicle is a trailer, lift or type is unkown where necessary
    for ind in vin_data.index:
        if vin_data['VEHICLE TYPE'][ind] == None or vin_data['VEHICLE TYPE'][ind] == 'Error':
                if 'trailer' in vin_data['MODEL'][ind].lower():
                    vin_data['VEHICLE TYPE'][ind] ='TRAILER'
                elif 'lift' in vin_data['MODEL'][ind].lower():
                    vin_data['VEHICLE TYPE'][ind] = 'LIFT'
                elif vin_data['VEHICLE TYPE'][ind] == 'Error':
                    vin_data['VEHICLE TYPE'][ind] = 'UNKNOWN'

    #create vin_data column indicating that somone needs to manually check a vehicle's VIN info using check_list
    vin_data['MANUAL CHECK NEEDED'] = check_list
    
    #add error code information to vin_data dataframe for employee reference
    vin_data = pd.concat([vin_data, results['ERROR CODE']], axis = 1)
    
    #valid_vins should be written to a CSV that is uploaded to SalesForce CAN compatability check, file path
    #should be the same as the input file with _CAN appended
    CAN_file_path = os.path.splitext(file_path)[0] + "_CAN.csv"
    pd.DataFrame(valid_vins).to_csv(CAN_file_path, index = False)
    
    #create file path and name for the processed vins output file, this file path is the same as the input
    #file path with _processed appended
    processed_file_path = os.path.splitext(file_path)[0] + "_processed.xlsx"
    
    #write vin_data dataframe to Excel file table, this will be the inclusive excel file with all VINS,
    #error codes, manual checks and vehicle types for employee reference
    with pd.ExcelWriter(processed_file_path, engine='openpyxl') as writer:
        
        #create an Excel sheet names 'Processed VINs' to hold the dataframe
        vin_data.to_excel(writer, index=False, sheet_name='Processed VINs')

        #access Excel file and worksheet 
        workbook = writer.book
        worksheet = writer.sheets['Processed VINs']

        #iterate through columns, find the max width of the cells in the column
        for idx, column in enumerate(worksheet.columns):
            #skip over 'ERROR CODE' column as it is the last column, formatting is unneccesary here
            if worksheet.cell(row=1, column=idx + 1).value != 'ERROR CODE':
                max_length = 0
                for cell in column:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                    #adjust column width to show all data
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[chr(65 + idx)].width = adjusted_width  #use index to get column letter
            
            #adjust 'ERROR CODE' column to be the width of the title
            if worksheet.cell(row=1, column=idx + 1).value == 'ERROR CODE':
                worksheet.column_dimensions[chr(65 + idx)].width = 12
        
    return processed_file_path, CAN_file_path

st.title("VIN Decoder")

uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xls", "xlsx", "csv"])

if "processed_file_path" not in st.session_state:
    st.session_state["processed_file_path"] = None
    st.session_state["can_file_path"] = None

if uploaded_file is not None:
    with st.spinner('Processing...'):
        input_file_path = uploaded_file.name
        with open(input_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        processed_file_path, can_file_path = confirm_vin(input_file_path)
        st.session_state["processed_file_path"] = processed_file_path
        st.session_state["can_file_path"] = can_file_path
        st.success('File successfully processed!')

if st.session_state["processed_file_path"] and st.session_state["can_file_path"]:
    with open(st.session_state["processed_file_path"], "rb") as f:
        processed_data = f.read()
    with open(st.session_state["can_file_path"], "rb") as f:
        can_data = f.read()
    
    st.download_button(
        label="Download Processed File",
        data=BytesIO(processed_data),
        file_name=os.path.basename(st.session_state["processed_file_path"]),
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    
    st.download_button(
        label="Download CAN File",
        data=BytesIO(can_data),
        file_name=os.path.basename(st.session_state["can_file_path"]),
        mime='text/csv'
    )
