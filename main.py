import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import urllib.parse
from colorama import init, Fore, Style


# Initialize colorama
init(autoreset=True)

print(f"{Fore.BLUE}Welcome to BniReferer Scraper.{Style.RESET_ALL}")

print(f"{Fore.YELLOW}---------------------------------------------------------------------------------------------------------------------{Style.RESET_ALL}")
print(f"{Fore.YELLOW}For this scraper to work you must have input.xlsx file in this same directory with all chapter links in first column.{Style.RESET_ALL}")
print(f"{Fore.YELLOW}If the scraper crashes the data will not be saved so try to give scraper small amount of data if you have less time. \nBecause data will be saved after all data is scraped.{Style.RESET_ALL}")
print(f"{Fore.YELLOW}---------------------------------------------------------------------------------------------------------------------{Style.RESET_ALL}")

# Automatically install the appropriate version of ChromeDriver
chromedriver_autoinstaller.install()

def scrape_table_from_url(url):
    try:
        # Initialize Chrome options and set headless mode
        chrome_options = Options()
        chrome_options.add_argument("--headless")

        # Initialize Selenium webdriver with Chrome options
        driver = webdriver.Chrome(options=chrome_options)
        driver.get(url)

        # Wait for the "members" div to be present
        members_div = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "members"))
        )

        # Parse HTML content using BeautifulSoup
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Find the div with id "members"
        members_div = soup.find('div', id='members')
        if not members_div:
            print(f"No div with id 'members' found on {url}")
            return None

        # Extract links and corresponding chapter names
        links_data = []
        for row in members_div.find('table').find_all('tr')[1:]:
            link = row.find('a')['href'] if row.find('a') else None
            chapter_name = extract_chapter_name_from_url(url)
            links_data.append((chapter_name, url, link))

        return links_data
    except Exception as e:
        print(f"Error occurred while scraping data from {url}: {e}")
        return None
    finally:
        driver.quit()

def extract_member_name_from_url(url):
    try:
        # Parse the URL
        parsed_url = urllib.parse.urlparse(url)
        # Get the query parameters
        query_params = urllib.parse.parse_qs(parsed_url.query)
        # Extract the sheet name from the 'name' parameter
        member_name = query_params.get('name', [''])[0]
        # URL decode the sheet name
        decoded_member_name = urllib.parse.unquote(member_name)
        return decoded_member_name
    except Exception as e:
        print(f"Error occurred while extracting sheet name: {e}")
        return None

def scrape_profile_data(profile_url):
    # Initialize Chrome options and set headless mode
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    # Initialize Selenium webdriver with Chrome options
    driver = webdriver.Chrome(options=chrome_options)
    base_url = "https://bnireferral.co.uk/en-GB/"
    full_profile_url = base_url + profile_url  # Add base URL
    data = []
    member_name = extract_member_name_from_url(full_profile_url)
    data.append(member_name)
    print(f"{Fore.GREEN}Scraping data from {profile_url} for Member {member_name}{Style.RESET_ALL}")
    try:
        driver.get(full_profile_url)
        # Wait for the textHolder div to be present
        text_holder_div = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "textHolder"))
        )
        # Parse HTML content using BeautifulSoup
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        # Find the div with class 'col-xs-12 col-sm-12 col-md-12'
        contact_details_div = soup.find_all('div', class_='col-xs-12 col-sm-12 col-md-12')[1]
        if contact_details_div:
            business_name=contact_details_div.find('p').text.strip()
            if business_name is None:
                business_name(' ')
            data.append(business_name)
            # Extract the profession from the h6 tag
            profession = contact_details_div.find('h6').text
            data.append(profession) 

        business_name_p = contact_details_div.find('p')
        website_a = business_name_p.find('a')
        if website_a:
            website = website_a.get('href').strip()
            data.append(website)
        else:
            data.append("No Website")
        # Find the ul with class 'memberContactInfo' 
        ul_contact_info = soup.find('ul', class_='memberContactInfo')
        if ul_contact_info: 
            # Extract phone number if it exists 
            phone_number_tag = ul_contact_info.findAll('a', href=True, attrs={'class': ''})[0] 
            if phone_number_tag: 
                phone_number = phone_number_tag.get('href').split(':')[-1]
                data.append(phone_number) 
            else: 
                data.append('No Phone Number') 
            # Extract email link if it exists 
            email_link_tag = ul_contact_info.findAll('a', href=True, attrs={'class': ''})[1]
            if email_link_tag: 
                email_link = base_url + email_link_tag.get('href')
                data.append(email_link) 
            else: 
                data.append('No Email') 
        textholder_data = []
        # Extract data from textHolder div
        text_holder_div = soup.find('div', class_='textHolder')
        text_holder_h6 = text_holder_div.find('h6')
        if text_holder_h6:
            # Extract text directly if there are no <br> tags
            if not text_holder_h6.find('br'):
                text = text_holder_h6.get_text().strip()
                # Filter out unwanted fields like member name, business name, and profession
                if text != member_name and text != business_name and text != profession:
                    textholder_data.append(text)
            else:
                # Extract all <br> tags and their surrounding text
                br_tags = text_holder_h6.find_all('br')
                previous_text = None
                for br_tag in br_tags:
                    # Get the text before and after each <br> tag
                    previous_sibling = br_tag.previous_sibling.strip() if br_tag.previous_sibling else ''
                    next_sibling = br_tag.next_sibling.strip() if br_tag.next_sibling else ''
                    # Check if the current text is different from the previous one
                    if previous_sibling and previous_sibling != previous_text:
                        if previous_sibling and \
                        member_name not in previous_sibling and \
                        business_name not in previous_sibling and \
                        profession not in previous_sibling:
                            textholder_data.append(previous_sibling.strip())
                            previous_text = previous_sibling
                    if next_sibling and next_sibling != previous_text:
                        if previous_sibling and \
                        member_name not in previous_sibling and \
                        business_name not in previous_sibling and \
                        profession not in previous_sibling:                            
                            textholder_data.append(next_sibling)
                            previous_text = next_sibling

            address = ', '.join(textholder_data).strip()
            data.append(address)
        else:
            print("textHolder div not found.")

        return data
    except Exception as e:
        print(f"Error occurred while scraping profile data from {full_profile_url}: {e}")
        return None
    finally:
        driver.quit()

def scrape_and_save_tables(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
        data_to_save = []

        # Iterate over rows and scrape tables
        for index, row in df.iterrows():
            url = row.iloc[0]
            print(f"{Fore.GREEN}Scraping data from {url}{Style.RESET_ALL}")
            links_data = scrape_table_from_url(url)
            if links_data is not None:
                for chapter_name, chapter_url, profile_url in links_data:
                    profile_data = scrape_profile_data(profile_url)
                    if profile_data is not None:
                        data_to_save.append([chapter_name, chapter_url] + profile_data)
                    else:
                        print(f"{Fore.YELLOW}Failed to scrape profile data from {profile_url}. Skipping...{Style.RESET_ALL}")
            else:
                print(f"{Fore.YELLOW}No links found on {url}. Skipping...{Style.RESET_ALL}")

        # Convert data to DataFrame and save to Excel
        columns = ["Chapter Name", "Chapter URL", "Name", "Company", "Profession/Speciality", "Website", "Phone", "Email", "Address"]
        data_df = pd.DataFrame(data_to_save)
        data_df.to_excel(output_file, index=False)

        print(f"{Fore.GREEN}Scraped data saved to {output_file}{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}Error occurred: {e}{Style.RESET_ALL}")


def extract_chapter_name_from_url(url):
    try:
        # Parse the URL
        parsed_url = urllib.parse.urlparse(url)

        # Get the query parameters
        query_params = urllib.parse.parse_qs(parsed_url.query)

        # Extract the sheet name from the 'name' parameter
        chapter_name = query_params.get('name', [''])[0]

        # URL decode the sheet name
        decoded_chapter_name = urllib.parse.unquote(chapter_name)

        return decoded_chapter_name
    except Exception as e:
        print(f"{Fore.RED}Error occurred while extracting sheet name: {e}{Style.RESET_ALL}")
        return None


# Example usage
input_excel_file = "input.xlsx" 
output_excel_file = "output.xlsx"
scrape_and_save_tables(input_excel_file, output_excel_file)
