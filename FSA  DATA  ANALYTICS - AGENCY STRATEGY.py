#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_Data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Group the data by FSA and calculate attrition
attrition_data = data.groupby('FSA').agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x <= 1).sum())
)
attrition_data['Attrition_Rate'] = (attrition_data['Attrition_Count'] / attrition_data['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data['Attrition_Rate'] = attrition_data['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Select the required columns
attrition_data = attrition_data[['Number_of_Donors', 'Attrition_Rate']]

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Attrition_Analysis_Result.xlsx"
attrition_data.to_excel(output_file_path, index=False)

print(f"Attrition analysis by FSA with percentage has been saved to {output_file_path}")


# In[4]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_Data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Group the data by FSA and calculate attrition
attrition_data = data.groupby('FSA').agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x <= 1).sum())
).reset_index()
attrition_data['Attrition_Rate'] = (attrition_data['Attrition_Count'] / attrition_data['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data['Attrition_Rate'] = attrition_data['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Attrition_Analysis_Result.xlsx"
attrition_data.to_excel(output_file_path, index=False)

print(f"Attrition analysis by FSA with percentage has been saved to {output_file_path}")


# In[5]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_Data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Convert 'Sign-Up Date' to datetime and filter for the year 2021
data['Sign-Up Date'] = pd.to_datetime(data['Sign-Up Date'], format='%d/%m/%Y')
data_2021 = data[data['Sign-Up Date'].dt.year == 2021]

# Group the filtered data by FSA and calculate 12th Month Attrition
attrition_data_2021 = data_2021.groupby('FSA').agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x < 12).sum())
)
attrition_data_2021['Attrition_Rate'] = (attrition_data_2021['Attrition_Count'] / attrition_data_2021['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data_2021['Attrition_Rate'] = attrition_data_2021['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Reset the index to include 'FSA' as a column
attrition_data_2021.reset_index(inplace=True)

# Save the result to a new Excel file
output_file_path_2021 = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\12th_Month_Attrition_Analysis_Result.xlsx"
attrition_data_2021.to_excel(output_file_path_2021, index=False)

print(f"12th Month Attrition analysis for the year 2021 by FSA has been saved to {output_file_path_2021}")


# In[6]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_Data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Group the data by FSA and AGE_RANGE and calculate attrition
attrition_data_age_range = data.groupby(['FSA', 'AGE_RANGE']).agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x <= 1).sum())
)
attrition_data_age_range['Attrition_Rate'] = (attrition_data_age_range['Attrition_Count'] / attrition_data_age_range['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data_age_range['Attrition_Rate'] = attrition_data_age_range['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Reset the index to include 'FSA' and 'AGE_RANGE' as columns
attrition_data_age_range.reset_index(inplace=True)

# Save the result to a new Excel file
output_file_path_age_range = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Attrition_Analysis_Age_Range_Result.xlsx"
attrition_data_age_range.to_excel(output_file_path_age_range, index=False)

print(f"Second month attrition analysis by AGE_RANGE for each FSA has been saved to {output_file_path_age_range}")


# In[7]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_Data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Convert 'Sign-Up Date' to datetime and filter for the year 2021
data['Sign-Up Date'] = pd.to_datetime(data['Sign-Up Date'], format='%d/%m/%Y')
data_2021 = data[data['Sign-Up Date'].dt.year == 2021]

# Group the filtered data by FSA and AGE_RANGE and calculate 12th Month Attrition
attrition_data_2021_age_range = data_2021.groupby(['FSA', 'AGE_RANGE']).agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x < 12).sum())
)
attrition_data_2021_age_range['Attrition_Rate'] = (attrition_data_2021_age_range['Attrition_Count'] / attrition_data_2021_age_range['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data_2021_age_range['Attrition_Rate'] = attrition_data_2021_age_range['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Reset the index to include 'FSA' and 'AGE_RANGE' as columns
attrition_data_2021_age_range.reset_index(inplace=True)

# Save the result to a new Excel file
output_file_path_2021_age_range = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\12th_Month_Attrition_Analysis_Age_Range_Result.xlsx"
attrition_data_2021_age_range.to_excel(output_file_path_2021_age_range, index=False)

print(f"12th Month Attrition analysis for the year 2021 by AGE_RANGE for each FSA has been saved to {output_file_path_2021_age_range}")


# In[1]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\UNICEF_data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Group the data by FSA and calculate attrition
attrition_data = data.groupby('FSA').agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x <= 1).sum())
).reset_index()
attrition_data['Attrition_Rate'] = (attrition_data['Attrition_Count'] / attrition_data['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data['Attrition_Rate'] = attrition_data['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\Attrition_Analysis_Result.xlsx"
attrition_data.to_excel(output_file_path, index=False)

print(f"Attrition analysis by FSA with percentage has been saved to {output_file_path}")


# In[3]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\UNICEF_data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Convert 'Sign-Up Date' to datetime and filter for the year 2021
# Using 'infer_datetime_format=True' to handle inconsistent date formats
data['Sign-Up Date'] = pd.to_datetime(data['Sign-Up Date'], infer_datetime_format=True)
data_2021 = data[data['Sign-Up Date'].dt.year == 2021]

# Group the filtered data by FSA and calculate 12th Month Attrition
attrition_data_2021 = data_2021.groupby('FSA').agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x < 12).sum())
).reset_index()
attrition_data_2021['Attrition_Rate'] = (attrition_data_2021['Attrition_Count'] / attrition_data_2021['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data_2021['Attrition_Rate'] = attrition_data_2021['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Save the result to a new Excel file
output_file_path_2021 = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\12th_Month_Attrition_Analysis_Result.xlsx"
attrition_data_2021.to_excel(output_file_path_2021, index=False)

print(f"12th Month Attrition analysis for the year 2021 by FSA has been saved to {output_file_path_2021}")


# In[4]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\UNICEF_data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Group the data by FSA and AGE_RANGE and calculate attrition
attrition_data_age_range = data.groupby(['FSA', 'AGE_RANGE']).agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x <= 1).sum())
)
attrition_data_age_range['Attrition_Rate'] = (attrition_data_age_range['Attrition_Count'] / attrition_data_age_range['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data_age_range['Attrition_Rate'] = attrition_data_age_range['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Reset the index to include 'FSA' and 'AGE_RANGE' as columns
attrition_data_age_range.reset_index(inplace=True)

# Save the result to a new Excel file
output_file_path_age_range = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\Attrition_Analysis_Age_Range_Result.xlsx"
attrition_data_age_range.to_excel(output_file_path_age_range, index=False)

print(f"Second month attrition analysis by AGE_RANGE for each FSA has been saved to {output_file_path_age_range}")


# In[5]:


import pandas as pd

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\UNICEF_data_for_FSA_Target.xlsx"
data = pd.read_excel(file_path)

# Convert 'Sign-Up Date' to datetime and filter for the year 2021
# Using 'infer_datetime_format=True' to handle inconsistent date formats
data['Sign-Up Date'] = pd.to_datetime(data['Sign-Up Date'], infer_datetime_format=True)
data_2021 = data[data['Sign-Up Date'].dt.year == 2021]

# Group the filtered data by FSA and AGE_RANGE and calculate 12th Month Attrition
attrition_data_2021_age_range = data_2021.groupby(['FSA', 'AGE_RANGE']).agg(
    Number_of_Donors=('FSA', 'count'),
    Attrition_Count=('NUMBEROFPAYMENTS', lambda x: (x < 12).sum())
)
attrition_data_2021_age_range['Attrition_Rate'] = (attrition_data_2021_age_range['Attrition_Count'] / attrition_data_2021_age_range['Number_of_Donors']) * 100

# Format the 'Attrition_Rate' column to show percentages
attrition_data_2021_age_range['Attrition_Rate'] = attrition_data_2021_age_range['Attrition_Rate'].apply(lambda x: f"{x:.2f}%")

# Reset the index to include 'FSA' and 'AGE_RANGE' as columns
attrition_data_2021_age_range.reset_index(inplace=True)

# Save the result to a new Excel file
output_file_path_2021_age_range = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\UNICEF\\12th_Month_Attrition_Analysis_Age_Range_Result.xlsx"
attrition_data_2021_age_range.to_excel(output_file_path_2021_age_range, index=False)

print(f"12th Month Attrition analysis for the year 2021 by AGE_RANGE for each FSA has been saved to {output_file_path_2021_age_range}")


# In[2]:


import pandas as pd
import matplotlib.pyplot as plt

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB 2nd Month FSA Attrition_Data_Result.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a string to use string methods
data['Attrition_Rate'] = data['Attrition_Rate'].astype(str)

# Remove the percentage sign and convert 'Attrition_Rate' to float
data['Attrition_Rate'] = data['Attrition_Rate'].str.rstrip('%').astype(float)

# Filter FSAs with at least 10 donors
filtered_data = data[data['Number_of_Donors'] >= 10]

# Sort the data by 'Attrition_Rate' in ascending order for better visualization
filtered_data = filtered_data.sort_values('Attrition_Rate')

# Create a bar chart
plt.figure(figsize=(12, 8))
plt.bar(filtered_data['FSA'], filtered_data['Attrition_Rate'], color='skyblue')
plt.xlabel('FSA')
plt.ylabel('Attrition Rate (%)')
plt.title('Attrition Rate by FSA for FSAs with at least 10 Donors')
plt.xticks(rotation=90)  # Rotate the FSA labels for better readability
plt.tight_layout()  # Adjust layout to fit all labels

# Save the plot as an image file
output_chart_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\FSA_Attrition_Rate_Chart.png"
plt.savefig(output_chart_path)
plt.show()


# In[4]:


pip install geopandas


# In[6]:


pip install folium


# In[ ]:





# In[10]:


pip install geopy


# In[ ]:





# In[1]:


import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB 2nd Month FSA Attrition_Data_Result.xlsx"
data = pd.read_excel(file_path)

# Initialize the geocoder with a unique user-agent
geolocator = Nominatim(user_agent="your_unique_user_agent")

# Use rate limiter to avoid overloading the geocode API
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)

# Define a function to geocode FSA with exception handling
def geocode_fsa(fsa):
    try:
        location = geocode(fsa + ", Canada")
        if location:
            return pd.Series([location.latitude, location.longitude])
    except Exception as e:
        print(f"Error geocoding {fsa}: {e}")
    return pd.Series([None, None])

# Apply the function to the 'FSA' column
data[['latitude', 'longitude']] = data['FSA'].apply(geocode_fsa)

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data.to_excel(output_file_path, index=False)

print(f"Data with latitude and longitude has been saved to {output_file_path}")


# In[2]:


import pandas as pd
from geopy.geocoders import Nominatim, ArcGIS
from geopy.extra.rate_limiter import RateLimiter

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB 2nd Month FSA Attrition_Data_Result.xlsx"
data = pd.read_excel(file_path)

# Initialize the geocoders with a unique user-agent
geolocator = Nominatim(user_agent="your_unique_user_agent")
arcgis_geolocator = ArcGIS(user_agent="your_unique_user_agent")

# Use rate limiter to avoid overloading the geocode API
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)
arcgis_geocode = RateLimiter(arcgis_geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)

# Define a function to geocode FSA with exception handling and fallback
def geocode_fsa(fsa):
    try:
        location = geocode(fsa + ", Canada")
        if location:
            return pd.Series([location.latitude, location.longitude])
        else:
            # Fallback geocoder
            fallback_location = arcgis_geocode(fsa + ", Canada")
            if fallback_location:
                return pd.Series([fallback_location.latitude, fallback_location.longitude])
    except Exception as e:
        print(f"Error geocoding {fsa}: {e}")
    return pd.Series([None, None])

# Apply the function to the 'FSA' column
data[['latitude', 'longitude']] = data['FSA'].apply(geocode_fsa)

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data.to_excel(output_file_path, index=False)

print(f"Data with latitude and longitude has been saved to {output_file_path}")


# In[5]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float and convert from percentage to a decimal
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float) / 100

# Filter FSAs with at least 10 donors
filtered_data = data[data['Number_of_Donors'] >= 10]

# Invert the attrition rate to use as weight for the heatmap
filtered_data['Weight'] = 1 - filtered_data['Attrition_Rate']

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Weight']].dropna().values.tolist(), radius=25, max_zoom=13).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\FSA_Attrition_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap of FSA attrition rates has been saved to {output_map_path}")


# In[6]:


import pandas as pd
from geopy.geocoders import Nominatim, ArcGIS
from geopy.extra.rate_limiter import RateLimiter
import time

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB 12th_Month_Attrition_Data_Result.xlsx"
data = pd.read_excel(file_path)

# Initialize the geocoders with a unique user-agent
geolocator = Nominatim(user_agent="your_unique_user_agent", timeout=10)  # Increase the timeout to 10 seconds
arcgis_geolocator = ArcGIS(user_agent="your_unique_user_agent", timeout=10)

# Use rate limiter to avoid overloading the geocode API
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)
arcgis_geocode = RateLimiter(arcgis_geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)

# Define a function to geocode FSA with exception handling, fallback, and retries
def geocode_fsa(fsa):
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            location = geocode(fsa + ", Canada")
            if location:
                return pd.Series([location.latitude, location.longitude])
            else:
                # Fallback geocoder
                fallback_location = arcgis_geocode(fsa + ", Canada")
                if fallback_location:
                    return pd.Series([fallback_location.latitude, fallback_location.longitude])
            break  # Break the loop if geocoding is successful
        except Exception as e:
            print(f"Error geocoding {fsa}: {e}")
            time.sleep(2 ** retry_count)  # Exponential backoff
            retry_count += 1
    return pd.Series([None, None])

# Apply the function to the 'FSA' column
data[['latitude', 'longitude']] = data['FSA'].apply(geocode_fsa)

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_12th_Month_Attrition_Data_Result_with_Lat_Lng.xlsx"
data.to_excel(output_file_path, index=False)

print(f"Data with latitude and longitude has been saved to {output_file_path}")


# In[7]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_12th_Month_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float and convert from percentage to a decimal
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float) / 100

# Filter FSAs with at least 10 donors
filtered_data = data[data['Number_of_Donors'] >= 10]

# Invert the attrition rate to use as weight for the heatmap
filtered_data['Weight'] = 1 - filtered_data['Attrition_Rate']

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Weight']].dropna().values.tolist(), radius=25, max_zoom=13).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_12th_Month_Attrition_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap of 12th month FSA attrition rates has been saved to {output_map_path}")


# In[8]:


import pandas as pd
from geopy.geocoders import Nominatim, ArcGIS
from geopy.extra.rate_limiter import RateLimiter
import time

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB 2nd Month Attrition_Analysis_Age_Range_Result.xlsx"
data = pd.read_excel(file_path)

# Initialize the geocoders with a unique user-agent
geolocator = Nominatim(user_agent="your_unique_user_agent", timeout=10)  # Increase the timeout to 10 seconds
arcgis_geolocator = ArcGIS(user_agent="your_unique_user_agent", timeout=10)

# Use rate limiter to avoid overloading the geocode API
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)
arcgis_geocode = RateLimiter(arcgis_geolocator.geocode, min_delay_seconds=1, error_wait_seconds=10)

# Define a function to geocode FSA with exception handling, fallback, and retries
def geocode_fsa(fsa):
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            location = geocode(fsa + ", Canada")
            if location:
                return pd.Series([location.latitude, location.longitude])
            else:
                # Fallback geocoder
                fallback_location = arcgis_geocode(fsa + ", Canada")
                if fallback_location:
                    return pd.Series([fallback_location.latitude, fallback_location.longitude])
            break  # Break the loop if geocoding is successful
        except Exception as e:
            print(f"Error geocoding {fsa}: {e}")
            time.sleep(2 ** retry_count)  # Exponential backoff
            retry_count += 1
    return pd.Series([None, None])

# Apply the function to the 'FSA' column
data[['latitude', 'longitude']] = data['FSA'].apply(geocode_fsa)

# Save the result to a new Excel file
output_file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Result_with_Lat_Lng.xlsx"
data.to_excel(output_file_path, index=False)

print(f"Data with latitude and longitude has been saved to {output_file_path}")


# In[9]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float and convert from percentage to a decimal
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float) / 100

# Filter FSAs with at least 10 donors
filtered_data = data[data['Number_of_Donors'] >= 10]

# Invert the attrition rate to use as weight for the heatmap
filtered_data['Weight'] = 1 - filtered_data['Attrition_Rate']

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
heat_map = HeatMap(data=filtered_data[['latitude', 'longitude', 'Weight']].dropna().values.tolist(), radius=25, max_zoom=13)
m.add_child(heat_map)

# Add popups with AGE_RANGE information
for idx, row in filtered_data.iterrows():
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=f"AGE_RANGE: {row['AGE_RANGE']}",
        icon=folium.Icon(icon='info-sign')
    ).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap with AGE_RANGE popups has been saved to {output_map_path}")


# In[5]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float and convert from percentage to a decimal
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float) / 100

# Filter FSAs with at least 10 donors and Attrition Rate <= 15%
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] <= 0.15)]

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Attrition_Rate']].dropna().values.tolist(), 
        radius=25, 
        max_zoom=13, 
        gradient={1: 'blue'}).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Heatmap_Low_Attrition.html"
m.save(output_map_path)

print(f"Heatmap of 2nd month FSA attrition rates (<= 15%) has been saved to {output_map_path}")


# In[10]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float)

# Filter FSAs with at least 10 donors and Attrition Rate > 0.15
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] > 0.15)]

# Check if there are any FSAs that meet the criteria
if len(filtered_data) == 0:
    print("No FSAs with attrition rate > 15%")
else:
    # Create a base map
    m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

    # Add a heat map layer
    HeatMap(data=filtered_data[['latitude', 'longitude', 'Attrition_Rate']].dropna().values.tolist(), 
            radius=25, 
            max_zoom=13, 
            gradient={1: 'red'}).add_to(m)

    # Save the map to an HTML file
    output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_Better_Attrition_Heatmap.html"
    m.save(output_map_path)

    print(f"Heatmap of 2nd month FSA attrition rates (> 15%) has been saved to {output_map_path}")


# In[12]:


# Filter FSAs with at least 10 donors and Attrition Rate > 15%
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] > 0.15)]

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Attrition_Rate']].dropna().values.tolist(), 
        radius=25, 
        max_zoom=13, 
        gradient={1: 'red'}).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_Poor_Attrition_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap of 2nd month FSA attrition rates (> 15%) has been saved to {output_map_path}")


# In[13]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float)

# Filter FSAs with at least 10 donors and Attrition Rate <= 15%
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] <= 0.15)]

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Attrition_Rate']].dropna().values.tolist(), 
        radius=25, 
        max_zoom=13, 
        gradient={1: 'blue'}).add_to(m)

# Add popups with AGE_RANGE information
for idx, row in filtered_data.iterrows():
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=f"AGE_RANGE: {row['AGE_RANGE']}",
        icon=folium.Icon(icon='info-sign')
    ).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Finala Output Data\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Better_Attrition_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap of 2nd month FSA attrition rates (<= 15%) with AGE_RANGE popups has been saved to {output_map_path}")


# In[ ]:





# In[ ]:


# A different type of Heatmap to make it easy to navigate for the agencies


# In[3]:


import zipfile

zip_file_path = "C:\\Users\\ayode\\anaconda3\\Lib\\site-packages\\plotly\\package_data\\datasets\\lpr_000b21a_e.zip"
extract_to_path = "C:\\Users\\ayode\\anaconda3\\Lib\\site-packages\\plotly\\package_data\\datasets\\"

with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_to_path)


# In[6]:





# In[9]:


print(gdf.columns)


# In[ ]:





# In[1]:


get_ipython().system('pip install chardet')


# In[ ]:





# In[9]:


pip install matplotlib seaborn plotly folium geopandas


# In[ ]:





# In[ ]:





# In[ ]:





# In[11]:


print(data.columns)


# In[14]:


import geopandas as gpd

# Load the Shapefile with FSA boundaries
shapefile_path = "C:\\Users\\ayode\\anaconda3\\Lib\\site-packages\\plotly\\package_data\\datasets\\lfsa000b21a_e\\lfsa000b21a_e.shp"
gdf = gpd.read_file(shapefile_path)

# Print the first few rows of the GeoDataFrame
print(gdf.head())


# In[16]:


import pandas as pd
import folium
import geopandas as gpd
import json  # import the json module

# Load the data
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Final Output Data\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Filter data
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] <= 0.15)]

# Load the Shapefile with FSA boundaries
shapefile_path = "C:\\Users\\ayode\\anaconda3\\Lib\\site-packages\\plotly\\package_data\\datasets\\lfsa000b21a_e\\lfsa000b21a_e.shp"
gdf = gpd.read_file(shapefile_path)

# Convert to GeoJSON
geojson_data = json.loads(gdf.to_json())

# Print the first feature to inspect the structure
print(geojson_data['features'][0])


# In[ ]:





# In[6]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Final Output Data\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float)

# Filter FSAs with at least 10 donors and Attrition Rate <= 15%
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] <= 0.15)]

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Attrition_Rate']].dropna().values.tolist(), 
        radius=25, 
        max_zoom=13, 
        gradient={1: 'blue'}).add_to(m)

# Add circles
for idx, row in filtered_data.iterrows():
    folium.Circle(
        location=[row['latitude'], row['longitude']],
        radius=1000,  # Adjust the size of the circle as needed
        color='blue',
        fill=True,
        fill_color='blue'
    ).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Final Output Data\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Better_Attrition_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap of 2nd month FSA attrition rates (<= 15%) has been saved to {output_map_path}")


# In[ ]:





# In[ ]:


# Use a Minimum of 8 donors - A correction from the management


# In[1]:


import pandas as pd
import folium
from folium.plugins import HeatMap

# Load the data from the Excel file
file_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Final Output Data\\CNIB_2nd_Month_FSA_Attrition_Data_Result_with_Lat_Lng.xlsx"
data = pd.read_excel(file_path)

# Ensure 'Attrition_Rate' is a float
data['Attrition_Rate'] = data['Attrition_Rate'].astype(float)

# Filter FSAs with at least 10 donors and Attrition Rate <= 15%
filtered_data = data[(data['Number_of_Donors'] >= 10) & (data['Attrition_Rate'] <= 0.15)]

# Create a base map
m = folium.Map(location=[56.1304, -106.3468], zoom_start=5)  # Canada's approximate center coordinates

# Add a heat map layer
HeatMap(data=filtered_data[['latitude', 'longitude', 'Attrition_Rate']].dropna().values.tolist(), 
        radius=25, 
        max_zoom=13, 
        gradient={1: 'blue'}).add_to(m)

# Add circles
for idx, row in filtered_data.iterrows():
    folium.Circle(
        location=[row['latitude'], row['longitude']],
        radius=1000,  # Adjust the size of the circle as needed
        color='blue',
        fill=True,
        fill_color='blue'
    ).add_to(m)

# Save the map to an HTML file
output_map_path = "C:\\Users\\ayode\\Desktop\\TNI Projects\\Agency Acquisition Strategy\\FSA Attrition - Income Analysis\\CNIB\\Final Output Data\\CNIB_2nd_Month_Attrition_Analysis_Age_Range_Better_Attrition_Heatmap.html"
m.save(output_map_path)

print(f"Heatmap of 2nd month FSA attrition rates (<= 15%) has been saved to {output_map_path}")


# In[ ]:





# In[ ]:




