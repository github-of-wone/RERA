#%%
#! Would be best that this be made into FastAPI related program with multi threading
#! Need to use PySimpleGUI to make the data insertion robust and easy for layman
#! IF possible would like to use something else besides selenium since it is VISIBLE

"""
This is a project to download data from MAHARERA website

https://maharerait.mahaonline.gov.in/searchlist/search?MenuID=1069

Example RERA IDs are 
P51800008478 - SENROOF II
P52100002646 - Pune Project
P51800032087 - SENROOF III


1 Create structure as below:
    Excel File with the {RERA NUMBER}.xlsx
    Folders x 6 with the following names:
        0 {RERA NUMBER}
        1 Organization
        2 Project
        3 Co-Promoters
        4 Project Details
        5 Uploaded Documents
    Folders x 20 with the following names in folder '5 Uploaded Documents':
        01 Copy of the legal title report 
        02 Details of encumbrances 
        03 Copy of Layout Approval (in case of layout) 
        04 Building Plan Approval  NA Order for plotted development 
        05 Commencement Certificates  NA Order for plotted development 
        06 Declaration about Commencement Certificate 
        07 Declaration in FORM B 
        08 Architect’s Certificate of Percentage of Completion of Work (Form 1)
        09 Engineer’s Certificate on Cost Incurred on Project (Form 2)
        10 CERSAI details 
        11 Engineers Certificate on Quality Assurance (Form 2A FY 2019-20)
        12 Disclosure of sold/ booked inventory
        13 Annual Audit Report of Statutory CA (Form 5) (FY 2017-18)
        14 Engineers Certificate on Quality Assurance (Form 2A FY 2020-21)
        15 Proforma of the allotment letter and agreement for sale 
        16 Architect’s Certificate on Completion of Project (Form 4)
        17 Status of Formation of Legal Entity (Society/Co Op etc.)
        18 Status of Conveyance
        19 Other
        20 Complaints

2 Create 'Excel File' with the following Sheets in it:
    {RERA NUMBER}
    Organization
    Project
    Co-Promoters
    Project Details
    Uploaded Documents

3 Go to the MAHARERA site and get to the site from where to download the RERA Project Data
<Insert Video here>
    Here one needs to download the table and place the information in the excel file sheet {RERA NUMBER}
    Out of the ten items in the row; 
    1 -4    first 4 are just text that needs to be inserted in the excel sheet, 
    5       the next {View} will open in next tab, 
    6       the 2nd blue button should download the pdf in the folder '0 {RERA NUMBER}'
    7       the {View Certificate} will download the MAHARERA Certificate which needs to be placed in the folder '0 {RERA NUMBER}'
    8       the {View Extension Certificate} also needs to be downloaded in the same Folder
    9       the View on Map is not required
    10      the Directions will give the Latitude and Longitude of the project

4 this step starts on the 5th point in 5 above
    Copy the data from MAHARERA Application until Organization Contact Details - Sheet 'Organization'
    Copy the data of Past Experience Details to Member Information - Sheet 'Organization' on Cell D2 onwards below
        Here there will be photo the promoter which needs to be downloaded in Folder 'Organization' with the file name to be the name of the person.  The link to that file needs to be kept in the 'View Photo' Cell
    Copy the information from Project until 'Co-Promoter' - Sheet 'Co-Promoters'.  Download the documents in '3 Co-Promoters'
    Copy the information from 'Project Details' until 'Litigation Details'
    Copy the information from 'Uploaded Documents' in '5 Uploaded Documents'.  This is where all the PDF in Download part of all the list need to be placed in folder '5 Uploaded Documents'
"""

"""
1 Create structure as below:
    Excel File with the {RERA NUMBER}.xlsx
    Folders x 6 with the following names:
        0 {RERA NUMBER}
        1 Organization
        2 Project
        3 Co-Promoters
        4 Project Details
        5 Uploaded Documents
    Folders x 20 with the following names in folder '5 Uploaded Documents':
        01 Copy of the legal title report
        02 Details of encumbrances
        03 Copy of Layout Approval (in case of layout)
        04 Building Plan Approval  NA Order for plotted development
        05 Commencement Certificates  NA Order for plotted development
        06 Declaration about Commencement Certificate
        07 Declaration in FORM B
        08 Architect’s Certificate of Percentage of Completion of Work (Form 1)
        09 Engineer’s Certificate on Cost Incurred on Project (Form 2)
        10 CERSAI details
        11 Engineers Certificate on Quality Assurance (Form 2A FY 2019-20)
        12 Disclosure of sold/ booked inventory
        13 Annual Audit Report of Statutory CA (Form 5) (FY 2017-18)
        14 Engineers Certificate on Quality Assurance (Form 2A FY 2020-21)
        15 Proforma of the allotment letter and agreement for sale
        16 Architect’s Certificate on Completion of Project (Form 4)
        17 Status of Formation of Legal Entity (Society/Co Op etc.)
        18 Status of Conveyance
        19 Other
        20 Complaints
"""


from pathlib import Path

MahaRERA_Registration_Number = "P51800008478"

dir0_list = [
    f"0 {MahaRERA_Registration_Number}",
    "1 Organization",
    "2 Project",
    "3 Co-Promoters",
    "4 Project Details",
    "5 Uploaded Documents",
]
dir5_list = [
    "01 Copy of the legal title report ",
    "02 Details of encumbrances ",
    "03 Copy of Layout Approval (in case of layout) ",
    "04 Building Plan Approval  NA Order for plotted development ",
    "05 Commencement Certificates  NA Order for plotted development ",
    "06 Declaration about Commencement Certificate ",
    "07 Declaration in FORM B ",
    "08 Architects Certificate of Percentage of Completion of Work (Form 1)",
    "09 Engineers Certificate on Cost Incurred on Project (Form 2)",
    "10 CERSAI details ",
    "11 Engineers Certificate on Quality Assurance (Form 2A FY 2019-20)",
    "12 Disclosure of sold/ booked inventory",
    "13 Annual Audit Report of Statutory CA (Form 5) (FY 2017-18)",
    "14 Engineers Certificate on Quality Assurance (Form 2A FY 2020-21)",
    "15 Proforma of the allotment letter and agreement for sale ",
    "16 Architects Certificate on Completion of Project (Form 4)",
    "17 Status of Formation of Legal Entity (Society/Co Op etc.)",
    "18 Status of Conveyance",
    "19 Other",
    "20 Complaints",
]

for item in dir0_list:
    Path.cwd().joinpath(item).mkdir(parents=True, exist_ok=True)

dir5_location = Path.cwd().joinpath(dir0_list[5])
for item in dir5_list:
    dir5_location.joinpath(item).mkdir(parents=True, exist_ok=True)


# %% [markdown]
""" 2 Create 'Excel File' with the following Sheets in it:
    {RERA NUMBER}
    Organization
    Project
    Co-Promoters
    Project Details
    Uploaded Documents
"""

MahaRERA_Registration_Number = "P51800008478"

import openpyxl

wb = openpyxl.Workbook()

sheet_list = [
    f"{MahaRERA_Registration_Number}",
    "Organization",
    "Project",
    "Co-Promoters",
    "Project Details",
    "Uploaded Documents",
]

for item in sheet_list:
    wb.create_sheet(item)

wb.save(MahaRERA_Registration_Number + ".xlsx")
# wb.remove('Sheet')
#! There is one sheet 'Sheet' that is automatically being inserted, unable to remove


#%%
"""3 Go to the MAHARERA site and get to the site from where to download the RERA Project Data
<Insert Video here>
    Here one needs to download the table and place the information in the excel file sheet {RERA NUMBER}
    Out of the ten items in the row; 
    1 -4    first 4 are just text that needs to be inserted in the excel sheet, 
    5       the next {View} will open in next tab, 
    6       the 2nd blue button should download the pdf in the folder '0 {RERA NUMBER}'
    7       the {View Certificate} will download the MAHARERA Certificate which needs to be placed in the folder '0 {RERA NUMBER}'
    8       the {View Extension Certificate} also needs to be downloaded in the same Folder
    9       the View on Map is not required
    10      the Directions will give the Latitude and Longitude of the project
"""

from lxml import etree
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

""" This 3 options below is to set the download folder location. """
prefs = {"download.default_directory": str(Path.cwd().joinpath(dir0_list[0]))}
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(executable_path=r"C:\pi\chromedriver.exe", chrome_options=options)


def xpath_click(xpath, driver=driver):
    """This function clicks on the xpath passed to it, after waiting until maximum 30 seconds

    Args:
        xpath ([type]): [description]
    """
    # WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, xpath.split('"')[1])))
    print(xpath.split('"')[1])
    driver.find_element_by_xpath(xpath).click()


def xpath_send_keys(xpath, send_keys1, driver=driver):
    """This function sends the keystrokes passed to it on the xpath passed to it, after waiting until maximum 30 seconds

    Args:
        xpath ([type]): [description]
        send_keys1 ([type]): [description]

    Returns:
        [type]: [description]
    """
    # WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, xpath.split('"')[1])))
    print(xpath.split('"')[1])
    driver.find_element_by_xpath(xpath).send_keys(send_keys1)


def xpath_get_value(xpath, driver=driver):
    """This function gets the value on the xpath passed to it, after waiting until maximum 30 seconds

    Args:
        xpath ([type]): [description]
        send_keys1 ([type]): [description]

    Returns:
        [type]: [description]
    """
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, xpath.split('"')[1])))
    print(xpath.split('"')[1])
    return driver.find_element_by_xpath(xpath).get_attribute("value")


# %%

url = "https://maharera.mahaonline.gov.in/"
driver.get(url)

#%%
fifth_item = '//*[@id="Registration"]'
registered_projects = '//*[@id="navbar"]/ul/li[5]/ul/li[1]/a'
xpath_click(fifth_item, driver)
xpath_click(registered_projects, driver)
#! Unable to click on the pop-up which appears here, somehow managing with MANUALLY click on it

# %%
driver.switch_to.window(
    driver.window_handles[1]
)  # source:https://stackoverflow.com/questions/28715942/how-do-i-switch-to-the-active-tab-in-selenium

registered_projects_rb = '//*[@id="Promoter"]'
registered_agents_rb = '//*[@id="Agent"]' #! Unused presently
revoked_projects_rb = '//*[@id="Revocation"]' #! Unused presently
MahaRERA_Registration_Number_txtbox = '//*[@id="CertiNo"]/'

xpath_click(registered_projects_rb, driver)
# xpath_send_keys(MahaRERA_Registration_Number_txtbox,MahaRERA_Registration_Number,driver)
#! Send Keys is also not working, it doesn't allow copy and paste even physically, hence this issue could be coming
# %% #! Do not delete this seperator

search_button_blue = '//*[@id="btnSearch"]'
back_button_red = '//*[@id="btnCancel"]'
reset_button_red = '//*[@id="btnReset"]'

view_details = '//*[@id="gridview"]/div[1]/div/table/tbody/tr/td[5]/b/a'
view_application = '//*[@id="Download"]/i'  # //*[@id="Download"]/i
view_certificate = '//*[@id="btnShow_64410"]/i'  # //*[@id="btnShow_64410"]/i
view_extension_certificate = ""  # TODO yet to place this here
view_on_map = ""  # TODO yet to place this here
directions = ""  # TODO yet to place this here

xpath_click(search_button_blue, driver)

# %%
# head=driver.find_element_by_class_name('table table-striped grid-table')
head = driver.find_element_by_xpath('//*[@id="gridview"]/div[1]/div/table')

#%%
# Source =https://codereview.stackexchange.com/questions/87901/copy-table-using-selenium-and-python
file_header = []
table_data = []
head_line = head.find_element_by_tag_name("tr")
headers = head_line.find_elements_by_tag_name("th")
# for header in headers:
#     header_text = header.text.encode("utf8")
# file_header.append(header_text)
file_header = [header.text.encode("utf8") for header in head_line.find_elements_by_tag_name("th")]

#%%
datas = head.find_elements_by_tag_name("tr")
# table_data = [data.text.encode("utf8") for data in datas.find_elements_by_tag_name('td')]
for data in datas:
    data_text = data.text.encode("utf8")
    table_data.append(data_text)  #! Why getting all the items concatenated and not seperate here?

#TODO to place the table from above in Excel Sheet '{RERA Registration Number}'

#! Need to have the other items download / link (map) here


# file_data.append(",".join(file_header))
# %%
xpath_click(view_application, driver)

import time

time.sleep(5)
xpath_click(
    view_certificate, driver
)  #! This one opens up the same Application inspite of adding time delay

# %%
"""
4 this step starts on the 5th point in 5 above
    Copy the data from MAHARERA Application until Organization Contact Details - Sheet 'Organization'
    Copy the data of Past Experience Details to Member Information - Sheet 'Organization' on Cell D2 onwards below
        Here there will be photo the promoter which needs to be downloaded in Folder 'Organization' with the file name to be the name of the person.  The link to that file needs to be kept in the 'View Photo' Cell
    Copy the information from Project until 'Co-Promoter' - Sheet 'Co-Promoters'.  Download the documents in '3 Co-Promoters'
    Copy the information from 'Project Details' until 'Litigation Details'
    Copy the information from 'Uploaded Documents' in '5 Uploaded Documents'.  This is where all the PDF in Download part of all the list need to be placed in folder '5 Uploaded Documents'
"""

driver.switch_to.window(driver.window_handles[2])
#! Is there an easier way of pasting the data from here to the excel file?
#! Downloading is difficult since it will download in the directory specified above during chrome options,but here will have to download into  different folders namely 1,3,5 and inside 5 in 20 different folders

#Donwloads for Folder 1 Organizations
Photo1='//*[@id="fldindtxt78"]/div[1]/table/tbody/tr[2]/td[3]/a'
Photo2='//*[@id="fldindtxt78"]/div[1]/table/tbody/tr[3]/td[3]/a'


#Donwloads for Folder 3 Co-Promoters
coinvestor_document1 = '//*[@id="DivCoPromoter"]/div[1]/table/tbody/tr[2]/td[5]/a' #There could be 3 or even more, this needs to be fine tuned
coinvestor_Upload_agreement_or_MoU_Copy='//*[@id="btnShow_5027"]'
coinvestor_Declaration_in_Form_B ='//*[@id="btnShow_5029"]'
dir0_list = [

#Donwloads for Folder 5 Uploaded Documents
#! Many need to match the folder with the name specified.
#! Likewise there is Complainant Details too!

#%%
#! Here all loose end need to be closed since the folder is ready as would be expected
