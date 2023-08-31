import pandas as pd
from selenium import webdriver
import time

# Load the original XLSX file into a DataFrame
input_file = 'links.xlsx'
df = pd.read_excel(input_file)

# Create a new DataFrame to store the extracted URLs
new_df = pd.DataFrame(columns=['Original URL', 'Final URL'])

# Function to get the final URL from a redirected link using Selenium
def get_final_url_selenium(url):
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver = webdriver.Chrome(options=options)

        driver.get(url)
        final_url = driver.current_url

        driver.quit()
        return final_url
    except Exception as e:
        return f"Error: {str(e)}"

# Iterate through each row in the original DataFrame
for index, row in df.iterrows():
    original_url = row['Hyperlink Column']
    final_url = get_final_url_selenium(original_url)
    new_df = new_df.append({'Original URL': original_url, 'Final URL': final_url}, ignore_index=True)
    
    # Save the new DataFrame with the current state of data to a new XLSX file
    new_file = 'path_to_partial_output_file.xlsx'  # Change to your desired partial output file name
    new_df.to_excel(new_file, index=False)

    # Add a waiting time of 2 seconds between requests
    time.sleep(1)

# Save the final DataFrame with extracted final URLs to the complete output XLSX file
final_output_file = 'path_to_your_final_output_file.xlsx'
new_df.to_excel(final_output_file, index=False)
