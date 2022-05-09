# Intro
An application to assist accountant/business owner automating federal tax deposit for multiple business accounts all at once

- Features:
    * Concurrent payment scheduler, receipt generator and transaction monitoring
    * Auto compute total tax deposit by crawling past complete transactions

# Setup
1. Install [Python](https://www.python.org/downloads/) and [Selenium](https://selenium-python.readthedocs.io/installation.html)

2. Create a new project folder
    - Copy over **genie_pay.py** and **template_data.xlsx**

    - Create **Receipt** and **Payment Summary** folders to store screenshot copies after each submission

    - Download [ChromeDriver](https://chromedriver.chromium.org/downloads), extract and copy file to the project directory

    - Change directory to the project directory and install Python dependencies
        ```
        pip install pipreqs
        pipreqs .
        pip install -r requirements.txt
        ```

3. Open genie_pay.py and edit path variables 
    - CHROME_DRIVER_PATH = "path to chromedriver.exe in the project directory"

    - DATA_PATH = "path to template_data.xlsx"

    - PAYMENT_RECEIPT_PATH = "path to receipt folder"

    - PAYMENT_SUMMARY_PATH = "path to payment history folder"

4. Add company and payment meta data to template_data.xlsx

5. Run program
    ```
    python genie_pay.py
    ```
6. Output
    * **Receipt** folder stores current payments screenshots

    * **Payment Summary** stores past payment screenshots (180-days period) 

    * **Payment_Status.xlsx** displays status of the most recent payment transactions - failed transactions are shown in red 

    * **Paid_Summary.xlsx** computes total tax deposit for a given quarter