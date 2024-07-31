# EmailCrawler

Project developed in Python and Selenium with the goal of automating Gmail login, searching for all unread emails, and converting this data into an Excel spreadsheet to return to the user. In the future, there are plans to integrate with WhatsApp to notify the user about emails and provide access to emails through a chatbot.

## Features
- Automates Gmail login and email retrieval.
- Extracts all unread emails and their details.
- Converts email data into an Excel spreadsheet format.
- Planned integration with WhatsApp for email notifications.
- Future integration with a chatbot for email access and interaction.

## Technologies Used
- Python
- Selenium
- Openpyxl (for spreadsheet generation)
- WhatsApp API (future integration)

## Clone the Repository
```bash
git clone <repository_url>
cd gmail-email-automation
```

# Install Dependencies
```bash
pip install -r requirements.txt
```

## Alternative: Manually Installing Dependencies

If it is not possible to install dependencies from requirements.txt, you can install them manually using the following commands:

``` bash
pip install selenium
pip install webdriver-manager
pip install python-dotenv
pip install validate_email
pip install openpyxl
```

# Installing openpyxl on Windows

### Method 1: Installing with PIP Manager

1- Open the Start Menu

 - Press the [Win] key or click the bottom left section of the desktop.

2- Type cmd in the Menu

 - Type cmd in the search bar.

3- Run cmd as Administrator

  - Right-click on the cmd result and select "Run as administrator".

4 - Install openpyxl Using PIP

 - In the terminal, type the following command and press Enter:

```bash 
py -m pip install openpyxl
```
 - This will install the openpyxl library in your Python distribution.

# Method 2: Installing Using the Official Package

1- Download the Official openpyxl Package

 - Open the download page for openpyxl from the official link.
2- Click on the Source Distribution Link
- Click the source distribution link 
 to start the download.

3- Extract the Files

- Navigate to the directory where the file was downloaded, right-click, and select "Extract here".

4- Open cmd.exe and Navigate to the Extracted Directory

- Open cmd.exe and change the working directory to the path where the files were extracted:

``` bash
cd path/to/extracted_directory
```

5- Install openpyxl Manually

- In the terminal, type the following command and press Enter:
``` bash
py setup.py install
```

- This will install the openpyxl library in your Python distribution.

## Run

``` bash
python main.py
```


# Future Plans

- Implement WhatsApp integration for real-time email notifications.
- Develop a chatbot interface for accessing and interacting with emails.
- Enhance error handling and robustness of the automation scripts.

# Contributing
Contributions are welcome! Feel free to submit issues or pull requests.

# Acknowledgements
- Inspired by the need for efficient email management.
- Built using Python and Selenium.

# Contact
For questions or feedback, please contact [Marco Damasceno](mailto:marcodamasceno0101@outlook.com)