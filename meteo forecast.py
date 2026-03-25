import openmeteo_requests

import pandas as pd
import requests_cache
from retry_requests import retry

# Setup the Open-Meteo API client with cache and retry on error
cache_session = requests_cache.CachedSession('.cache', expire_after = 3600)
retry_session = retry(cache_session, retries = 5, backoff_factor = 0.2)
openmeteo = openmeteo_requests.Client(session = retry_session)

# Make sure all required weather variables are listed here
# The order of variables in hourly or daily is important to assign them correctly below
url = "https://api.open-meteo.com/v1/forecast"
params = {
	"latitude": [-0.5603, -2.3328, -2.7334],
	"longitude": [34.0623, 29.0934, 29.7245],
	"hourly": ["temperature_180m", "wind_gusts_10m", "wind_direction_180m", "wind_speed_180m", "rain", "relative_humidity_2m"],
	"timezone": "auto",
	"past_days": 31,
}
responses = openmeteo.weather_api(url, params = params)

# Process 3 locations
for response in responses:
	print(f"\nCoordinates: {response.Latitude()}°N {response.Longitude()}°E")
	print(f"Elevation: {response.Elevation()} m asl")
	print(f"Timezone: {response.Timezone()}{response.TimezoneAbbreviation()}")
	print(f"Timezone difference to GMT+0: {response.UtcOffsetSeconds()}s")
	
	# Process hourly data. The order of variables needs to be the same as requested.
	hourly = response.Hourly()
	hourly_temperature_180m = hourly.Variables(0).ValuesAsNumpy()
	hourly_wind_gusts_10m = hourly.Variables(1).ValuesAsNumpy()
	hourly_wind_direction_180m = hourly.Variables(2).ValuesAsNumpy()
	hourly_wind_speed_180m = hourly.Variables(3).ValuesAsNumpy()
	hourly_rain = hourly.Variables(4).ValuesAsNumpy()
	hourly_relative_humidity_2m = hourly.Variables(5).ValuesAsNumpy()
	
	hourly_data = {"date": pd.date_range(
		start = pd.to_datetime(hourly.Time() + response.UtcOffsetSeconds(), unit = "s", utc = True),
		end =  pd.to_datetime(hourly.TimeEnd() + response.UtcOffsetSeconds(), unit = "s", utc = True),
		freq = pd.Timedelta(seconds = hourly.Interval()),
		inclusive = "left"
	)}
	
	hourly_data["temperature_180m"] = hourly_temperature_180m
	hourly_data["wind_gusts_10m"] = hourly_wind_gusts_10m
	hourly_data["wind_direction_180m"] = hourly_wind_direction_180m
	hourly_data["wind_speed_180m"] = hourly_wind_speed_180m
	hourly_data["rain"] = hourly_rain
	hourly_data["relative_humidity_2m"] = hourly_relative_humidity_2m
	
	hourly_dataframe = pd.DataFrame(data = hourly_data)
	print("\nHourly data\n", hourly_dataframe)
	