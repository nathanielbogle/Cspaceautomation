import os
import string 
from datetime import datetime
import pandas as pd 
# The most well known data analysis package on python allows us to work with the excel file.
from openpyxl import load_workbook
from openpyxl import Workbook
# Openpyxl allows us to work with the excel file and copy the range of cells that we need.
import numpy as np
# Numpy allows us to work with arrays and slice the data
from flask.logging import create_logger
from flask import Flask
from flask_cors import CORS
from flask import make_response, request, jsonify, current_app
from flask import current_app as app
from flask_sse import sse
# Flask is the web application framework used to create the program, it handles routing the webform request to the back end.
from webdriver_manager.chrome import ChromeDriverManager
# Webdriver manager ensures that the correct version of Selenium is downloaded for the users version of Google Chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
# Selenium handles automating the login and uploading process.

app = Flask(__name__)
CORS(app) 

# Enables the instance of flask using the special variable __name__. The cors line will enable CORS for all routes

#app.logger.setLevel(logging.INFO)
#app.register_blueprint(sse, url_prefix='/stream')


@app.route('/', methods=['POST'])

def script():
    if request.method == 'POST':

            formVars = request.form
            # Takes the variables sent from the http request

            incentiveCommsUrl = 'https://%s' % (formVars['incentiveCommsUrl'])
            incentiveEmail = formVars['incentiveEmail']
            incentivePw = formVars['incentivePw']
            incentivePostTitle = formVars['incentivePostTitle']
            incentivePostDescription = formVars['incentivePostDescription']
            incentiveTotal = (int(formVars['incentiveTotal'])+15)
            incentivesNamesPerColumn = int(formVars['incentiveNamesPerColumn'])
            IncentivePostOrDraft = formVars['IncentivePostOrDraft']

            #This list contains all of the form variables.

            chrome_options = Options()
            chrome_options.add_argument("--headless")
            #Selenium the browser automation package that logs in, uploads the image and simulates writing text is run in the background.
            #The two lines above enable Selenium to automate the process in the background, rather than opening up a brand new chrome window,
            #which would distract the user from their work.


            driver = webdriver.Chrome(ChromeDriverManager().install())
            
            #,chrome_options=chrome_options)
            #This enables Selenium to install the Chrome Driver Manager to the users cache. Chrome driver manager installs the Chrome Driver that
            #matches the version of the firmware version of a users Google Chrome allowing it to work regardless of the Chrome firmware version.
            
            driver.get(incentiveCommsUrl)
            #This tells the Chrome Web driver to go to the website that was submitted in the 'URL' section of the form.

            username = driver.find_element_by_name('Username')
            username.click
            #username.clear()
            username.send_keys(incentiveEmail)
            #This tells the Chrome Webdriver to find the username login field. Click on the field and type out the members email address
            #that they submitted through the form

            password = driver.find_element_by_name('Password')
            #password.clear()
            password.send_keys(incentivePw)
            #This tells the Chrome Webdriver to find the password login field. Click on the field and type out the members password
            #that they submitted through the form

            driver.find_element_by_id("LogOn").click()
            #This clicks the logon button

            incentiveCommsUrl = driver.current_url.strip('fac')
            driver.get(incentiveCommsUrl+'/Impact/Create')
            #These two lines strips the 'fac' part of the the page url post logon. This is because in the next line the url 
            # is then concatenated with /impact/create to take us to the section of the community where we can create a new impact post.
        
            driver.implicitly_wait(2)
            #Simply tells the Chrome Driver to wait as I've had some difficulties with this page and the loading of elements.

            post_title = driver.find_element_by_id('Title')
            #post_title.clear()
            post_title.send_keys(incentivePostTitle)
            #This finds the title part of the 'create a new impact post' page and writes the title submitted in the form.

            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "Description_ifr")))
            
            driver.switch_to_frame('Description_ifr')

            description = driver.find_element_by_id('tinymce')

            #tinyMceLoaded = driver.execute_script("return tinymce")
            tinyMceLoaded = driver.execute_script('return document.getElementById("tinymce").innerHTML')


            if not tinyMceLoaded:
                driver.implicitly_wait(3)

            # Working with iframes and tinymce in Selenium can be difficult. So here the webdriver waits again, checks to make sure that the
            # description part of the impact post page is visible. Then switches to the iframe to access the tinymce text
            # editor contained in the iframe. It then returns the html inside the tinyMce text editor to check that tinyMce is loaded.
             
            description.send_keys(' ')

            driver.execute_script('document.getElementById("tinymce").innerHTML = "%s"' % (incentivePostDescription))

            driver.switch_to.default_content()

            driver.implicitly_wait(5)

            nextBtn = driver.find_element_by_class_name('js-action-button').click()

            # Once its made sure that the tinyMce is working and is available to push data to. It posts the incentive post description
            # outlined by our user. It then switches us out of the iframe and clicks the 'next' button.

            sheet = request.files['file'].filename

            # Takes the file name of the file that was uploaded.

            if sheet:
                docs = os.path.expanduser("~/Documents")
                #app.logger.info(os.path.isdir("%s/IncentiveLoader" % (docs)))

                if not os.path.isdir("%s/IncentiveLoader" % (docs)):
                    os.mkdir("%s/IncentiveLoader" % (docs))

                directory = "%s/IncentiveLoader" % (docs)
                excel_file = request.files['file']

                wb = load_workbook(excel_file, read_only=True)

                # If a user does not already have a folder in documents called 'IncentiveLoader' this will make one, this file will 
                # be where all of the outputs are saved i.e the images, html and excel list of names in case they need to refer back and
                # double check what they uploaded. Excel_file gets the file object that was uploaded and wb loads the workbook in Openpyxl
                # that the user uploaded.

                today = datetime.now()
                today = today.strftime("%d-%b-%Y, %H:%M")
                ssName = "%s/incentives_%s.xlsx" % (directory, today)
                
                # Uses the date time package to automate the files to be called incentives + todays date.

                wsdetails = wb["Details"]
                template = Workbook()
                template_ws = template.active
                template.save(ssName)
                def copyRange(startCol, startRow, endCol, endRow, sheet):
                    rangeSelected = []
                    for i in range(startRow,endRow + 1,1):
                        rowSelected = []
                        for j in range(startCol,endCol+1,1):
                            rowSelected.append(sheet.cell(row = i, column = j).value)
                        rangeSelected.append(rowSelected)

                    return rangeSelected

                data = copyRange(2,15,2,(incentiveTotal-1),wsdetails)
                
                # Creates a function that copies and returns a range of cells based on some preset information i.e always start from cell
                # B15 and then takes information from our form variables incentives total. Only does it from our worksheet named 'details'.

                sortedData = []
                blanks = 0
                
                for name in data:
                    split = name[0].split()
                    name = [i.capitalize() for i in split]
                    lastname = name.pop()[0]
                    name = [' '.join(name) + ' ' + lastname]
                    
                    if '@' not in name[0]:
                        sortedData.append(name)
                    else:
                        blanks += 1
                    
                sortedData.sort()

                while blanks != 0:
                    sortedData.append([' '])
                    blanks -= 1

                # This cleans the data makes sure all the names are capitalized, makes sure emails are never included as any '@' in the name
                # will simply produce a blank name in the output, makes sure that only the last inital of a surname is ever posted
                # including users with middle names i.e the name 'Sandra De Fretias' now becomes 'Sandra De F'.

                counter = 0
                sliceCount = 0
                iteration = 0

                columns = string.ascii_lowercase

                while counter < len(data):
                    counter += incentivesNamesPerColumn
                    slice = np.array(sortedData)
                    k = 1
                                            
                    for name in slice[sliceCount:counter]:
                        cell = columns[iteration] + str(k)
                        template_ws[str(cell)] = name[0]
                        k += 1
                        
                    sliceCount +=incentivesNamesPerColumn
                    iteration +=1 

                template.save(ssName)

                #  Slices the data based on the number inputed into the incentivesNamesperColumn variable and saves the xlsx output as 'ssName'

                html = ssName.replace('.xlsx', '.html')
                img = html.replace('.html', '.png')

                # Creates the file name with the relevant extensions

                df = pd.read_excel(ssName, header=None)

                df.to_html(html, index=False, bold_rows=False,)

                driverTwo = webdriver.Chrome(ChromeDriverManager().install())
                driverTwo.get("file:///%s" % (html))
                driverTwo.save_screenshot(img)

                media_upload = driver.find_element_by_class_name('upload-button')
                media_upload.send_keys(img)
                tile_upload = driver.find_element_by_class_name('file-uploader')
                test = tile_upload.send_keys(img)

                # We use the pandas module to turn the excel output into a html file which we then open using our second web browser.
                # This serves two functions: 1) Visualises the data - this web browser is not headless which means it pops up 
                # allowing our end user to see what they are uploading allowing them to double check for errors.
                # 2) Screenshot - The web browser takes a screen shot of the output of names (the table) and uploads that to the community
                # using the two upload buttons (media upload and tile image).

                if IncentivePostOrDraft == 'post':

                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/main/article/footer/ul/li[2]/button'))).click()

                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div[3]/div/button[1]'))).click()

                    return 'Post has been made live'

                elif IncentivePostOrDraft == 'draft':
                    
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/main/article/fieldset[1]/header/section/section/ul/li[1]/button'))).click()
                    
                    return 'Post has been saved as a draft'

                # We use a conditional to see whether user has specified to save the post as draft, or make the post live


if __name__ == 'main':
    app.run(debug=False)

# Enables the flask app to run and gives us some possible errors through the debugger. 
