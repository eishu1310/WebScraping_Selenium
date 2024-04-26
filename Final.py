from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import os
#pip install python-docx 
import docx
import pandas as pd

def read_word_file(file_path):
    links = []
    doc = docx.Document(file_path)
    for paragraph in doc.paragraphs:
        link = paragraph.text.strip()
        if link.startswith("http"):
            links.append(link)
    return links


word_file_path = r"Python Assigment 2.docx"
links = read_word_file(word_file_path)
#print(links)


def scrape_website(link):
    driver = webdriver.Chrome()  # Change the path to your ChromeDriver
    driver.get(link)

    # Extract relevant information such as text, images, or links
    # Example:
    website_title = driver.title
    #print(website_title)
    #print("----------------")
    website_text = driver.find_element(By.TAG_NAME, 'body').text
    #print(website_text)
    #print("----------------")

    # Close the browser
    driver.quit()
    return website_title, website_text

# Function to scrape data from multiple website links
def scrape_multiple_websites(links):
    data = []
    for link in links:
        try:
            title, text = scrape_website(link)
            data.append({'Website Title': title, 'Website Text': text})
        except Exception as e:
            print(f"Error scraping website {link}: {str(e)}")
    return data


#scraped_data = scrape_website(links)
#print(scraped_data)


scraped_data = scrape_multiple_websites(links)

    # Convert scraped data to DataFrame
df = pd.DataFrame(scraped_data)

    # Store scraped data in an Excel file
excel_file_path = 'scraped_data.xlsx'
df.to_excel(excel_file_path, index=False)
print("Scraped data saved to:", excel_file_path)
