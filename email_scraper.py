from selenium import webdriver
from openpyxl import Workbook

email_link = "https://sis.moe.gov.sg/default.aspx?search=directory"

driver = webdriver.Chrome()
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(email_link)

# Select all checkboxes for search
checkbox_names = ['PriAll', 'SecAll', 'JCAll', 'CAll', 'MixedFAll', 'MixedT1All']

# Click all checkboxes
for checkbox_name in checkbox_names:
    checkbox = driver.find_element_by_name(checkbox_name)
    checkbox.click()

# Click search button
search_button = driver.find_element_by_id("Submit_Link")
search_button.click()

# Open CSV file
with open("school_emails.csv", 'w') as f:
    # Write CSV headers to file
    f.write("School Name, Email\n")

    # Get all matching elements
    school_email_elements = driver.find_elements_by_css_selector("a[class='getSchDetails']")

    # The school_email_elements selector alternates between school name and email.
    for index, school_email_element in enumerate(school_email_elements):

        # Write school name to file with a comma
        # We start at 0, so the first one will be 0, then 2 then 4 etc. This is the school name.
        if index % 2 == 0:
            school_name = school_email_element.text
            # TODO - handle school names with commas (Eg just use excel instead of CSV)
            f.write(school_name + ", ")

            print(school_name)

        # Write email after comma and add new line
        if index % 2 == 1:
            email_address = school_email_element.text
            f.write(email_address + "\n")

            print(email_address)
