import requests
import datetime
import pandas as pd

def get_weather_data(location, api_key):
    """Fetches weather data from the OpenWeatherMap API."""
    url = f"http://api.openweathermap.org/data/2.5/forecast?q={location}&appid={api_key}&units=metric"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error: Unable to fetch data. Status Code: {response.status_code}")
        return None

def is_suitable_for_farming(temp, rain_volume):
    """Determines farming suitability based on temperature and rainfall."""
    # Assuming rain_volume is a float representing rainfall in the last 3 hours
    # The condition is for a temperature between 0 and 30 degrees Celsius with no rainfall
    if temp > 0 and temp < 30 and rain_volume == 0:
        return "Suitable for farming"
    else:
        return "Not suitable for farming"

def display_weather_forecast(location, api_key):
    """
    Fetches and displays a 5-day farming weather forecast,
    and saves the data to an Excel file.
    """
    data = get_weather_data(location, api_key)
    if data is None:
        return  # Exit the function if API request failed

    weather_records = []  # List to hold the weather records

    if 'list' in data:
        print(f"\n5-Day Farming Weather Forecast for {location}:\n")
        print(f"{'Date':<15} {'Temp (°C)':<12} {'Humidity (%)':<15} {'Weather':<20} {'Rain (mm/3h)':<15} {'Farming Advice':<25}")
        print('-' * 115)

        for i in range(0, len(data['list']), 8):  # Taking one reading per day
            day = data['list'][i]
            date = datetime.datetime.fromtimestamp(day['dt']).strftime('%Y-%m-%d')
            temp = day['main']['temp']
            humidity = day['main']['humidity']
            weather = day['weather'][0]['main']
            
            # Check for rain volume; default to 0 if not present
            rain_volume = day.get('rain', {}).get('3h', 0)
            
            farming_advice = is_suitable_for_farming(temp, rain_volume)

            print(f"{date:<15} {temp:<12} {humidity:<15} {weather:<20} {rain_volume:<15} {farming_advice:<25}")

            # Append each day's data to the list
            weather_records.append({
                "Date": date,
                "Temperature (°C)": temp,
                "Humidity (%)": humidity,
                "Weather": weather,
                "Rain (mm/3h)": rain_volume,
                "Farming Advice": farming_advice
            })

        print('-' * 115)

        # Create a DataFrame and save it as an Excel file
        df = pd.DataFrame(weather_records)
        excel_filename = f"{location}_weather_forecast.xlsx"
        df.to_excel(excel_filename, index=False)
        print(f"Weather data saved to {excel_filename}")
    else:
        print("Weather data not found. Please check the location name and API key.")

if __name__ == "__main__":
    api_key = "b7a2c80dd927b3108ab86293101d9dc2"  # Replace with your actual OpenWeatherMap API key
    location = input("Enter a city or village name: ")
    display_weather_forecast(location, api_key)