import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# Set up the WebDriver (replace the path with the location of your WebDriver)
options = webdriver.ChromeOptions()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)


# Navigate to the IRDAI Agent Locator page
driver.get('https://agencyportal.irdai.gov.in/PublicAccess/AgentLocator.aspx')

# Select 'Life' from the Insurance Type dropdown
insurance_type_dropdown = Select(driver.find_element(By.ID, 'ddlInsuranceType'))
insurance_type_dropdown.select_by_value('2')  # Value for 'Life'

# Wait for the insurer dropdown to be populated
wait = WebDriverWait(driver, 10)

# Wait until the insurer dropdown is enabled (not just present)
wait.until(EC.element_to_be_clickable((By.ID, 'ddlInsurer')))

# Now that we are sure the dropdown is present, we can select the insurer
insurer_dropdown = Select(driver.find_element(By.ID, 'ddlInsurer'))

# Wait for the dropdown options to be populated
wait.until(lambda driver: len(insurer_dropdown.options) > 1)  # Ensure there are options other than the default

# Select the insurer by value
try:
    insurer_dropdown.select_by_value('33')  # Value for 'Kotak Mahindra Life Insurance company limited'
    print("Insurer selected successfully.")
except Exception as e:
    print(f"Error selecting insurer by value: {e}. Trying to select by visible text.")
    # Fallback to selecting by visible text if needed
    insurer_dropdown.select_by_visible_text('Kotak Mahindra Life Insurance company limited')

# Get the list of states
state_dropdown = Select(driver.find_element(By.ID, 'ddlState'))
state_options = [option.text for option in state_dropdown.options]

# Create a pandas DataFrame to store the data temporarily
columns = ['New Column', 'Agent Name', 'License No', 'IRDA URN', 'Agent ID', 'Insurance Type', 'Insurer', 'DP ID', 'State', 'District', 'PIN Code', 'Valid From', 'Valid To', 'Absorbed Agent', 'Phone No', 'Mobile No']
data_frame = pd.DataFrame(columns=columns)

# Iterate over each state
for state in state_options[1:]:  # Skip the first option (Select State)
    state_dropdown.select_by_visible_text(state)
    time.sleep(1)  # Wait for districts to load

    # Get the list of districts for the selected state
    district_dropdown = Select(driver.find_element(By.ID, 'ddlDistrict'))

    # Wait until the district dropdown is enabled and has options
    wait.until(EC.element_to_be_clickable((By.ID, 'ddlDistrict')))
    wait.until(lambda driver: len(district_dropdown.options) > 1)  # Ensure there are options other than the default

    # Check if the district dropdown is enabled and has options
    if district_dropdown.options:
        # Iterate over each district
        for district in district_dropdown.options[1:]:  # Skip the first option (Select District)
            district_dropdown.select_by_visible_text(district.text)

            # Click the Locate button
            locate_button = driver.find_element(By.ID, 'btnLocate')
            locate_button.click()
            time.sleep(2)  # Wait for the results to load

            # Extract the data from the table
            try:
                table = driver.find_element(By.ID, 'fgAgentLocator')
                rows = table.find_elements(By.TAG_NAME, 'tr')

                for row in rows:  # Include all rows, including the first one
                    columns = row.find_elements(By.TAG_NAME, 'td')
                    row_data = [column.text for column in columns]

                    # Debugging: Print the length of row_data
                    print(f"Row data length: {len(row_data)}")

                    # Ensure we have the correct number of columns
                    if len(row_data) == 15:  # Adjust this based on your actual data
                        row_data.insert(0, '')  # Add a new column with a placeholder value
                    elif len(row_data) != 16:
                        print(f"Skipping row due to unexpected column count: {len(row_data)}")
                        continue  # Skip this row if the column count is not as expected

                    # Add the row to the DataFrame
                    data_frame.loc[len(data_frame)] = row_data

                    # Display the current row data as a preview
                    print("Current record scraped:", row_data)

                    # Save the current DataFrame to Excel incrementally
                    data_frame.to_excel('output_file.xlsx', index=False)

            except Exception as e:
                print(f"Error extracting data for {state}, {district.text}: {e}")
    else:
        print(f"No districts available for state: {state}")

# Close the WebDriver
driver.quit()

print("Data has been saved to kotak_agents_data.xlsx")