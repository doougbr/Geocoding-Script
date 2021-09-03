import pandas as pd
import numpy as np
import googlemaps
import time
from openpyxl import Workbook
import json
import requests

GOOGLE_API_KEY = 'Insert your API Key here'

# creating a dataframe from Addresses file and setting ID column as index
address_df = pd.read_excel('Insert your file name here including its extension (.xlsx)', engine='openpyxl')
address_df.index.name = 'Index'

# standardize columns headers to lower case
address_df.columns = map(str.lower, address_df.columns)

# replace empty values with np.nan objects
address_df['city'].replace('', np.nan, inplace=True)

# remove rows containing np.nan objects in City column
address_df.dropna(subset=['city'], inplace=True)

# creating new column with concatenated values from dataframe
address_df['fullAddress'] = address_df['city'] + " " + address_df['address'] + " " + address_df['number'].map(str)

# getting values from fullAddress and returning as list
fullAddress_list = address_df['fullAddress'].tolist()

coordinates_list = []

print('requesting data... it may take a while...')


# geocoding an address getting the result from the JSON
def geocode_address():
    api_key = GOOGLE_API_KEY
    gmaps = googlemaps.Client(key=api_key)
    startTime = time.time()
    for address in fullAddress_list:
        try:
            geocode_result = gmaps.geocode(address)
            lat = geocode_result[0]["geometry"]["location"]["lat"]
            lon = geocode_result[0]["geometry"]["location"]["lng"]
            coordinate = str(lat) + "," + str(lon)
            coordinates_list.append(coordinate)
            time.sleep(0.02)
        except:
            coordinates_list.append("Couldn't find coordinates. Please check the if address is correct.")

    endTime = time.time()
    elapsedTime = endTime - startTime
    print('done!')
    print(f'elapsed time:{elapsedTime:.2f} seconds')


geocode_address()

# creating new column named coordinates
newAddress_df = address_df.assign(coordinates=coordinates_list)

# creating a Pandas writer
fileName = 'coordinates_generic.xlsx'
writer = pd.ExcelWriter(fileName, engine='xlsxwriter')

# creating the Excel file with the sheet name
newAddress_df.to_excel(writer, sheet_name='Sheet1')
# creating the Excel file
writer.save()

print(fileName + " saved successfully!")
