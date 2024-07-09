import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from services.validateEmail import is_valid_gmail
from services.createFolder import move_file_to_directory

load_dotenv()

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-infobars")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def open_browser(url):
    driver.get(url)


def login():
    try:
        open_browser("https://www.google.com/intl/pt-BR/gmail/about/")  # URL do Gmail

        login_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/header/div/div/div/a[2]"))
        )
        login_button.click()

        email_input = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'input'))
        )
        email = input("digit your email:")
        while not is_valid_gmail(email):
            print("invalid email")
            print("digit a email valid please")
            email = input("digit your email:")
        else:
            email_input.send_keys(email)

        button_next_email = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="identifierNext"]/div/button'))
        )
        button_next_email.click()

        password_input = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input'))
        )

        password = input("digit your password:")
        while not password:
            print("the password cannot be empty")
            password = input("digit your password again:")
        else:
            if password_input:
                password_input.send_keys(password)

        button_next_password = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="passwordNext"]/div/button'))
        )
        button_next_password.click()

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="main"]'))
        )

    except Exception as e:
        print(f"Error during login: {str(e)}")
        print(driver.page_source)


def extract_all_emails_text():
    try:

        email_rows = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'tr.zA'))
        )
        print(len(email_rows), 'EMAILS')

        emails_data = []

        for row in email_rows:
            try:

                subject_element = row.find_element(By.CSS_SELECTOR, 'span.bog')
                subject_text = subject_element.text

                snippet_element = row.find_element(By.CSS_SELECTOR, 'span.y2')
                snippet_text = snippet_element.text

                emails_data.append({
                    "Assunto": subject_text,
                    "Resumo": snippet_text
                })

            except Exception as e:
                print(f"Erro ao extrair texto do email: {str(e)}")

        if len(emails_data) > 0:

            wb = Workbook()
            ws = wb.active
            ws.title = "Emails"

            header_font = Font(bold=True, color="FFFFFF")
            fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")

            ws['A1'] = "Assunto"
            ws['A1'].font = header_font
            ws['A1'].fill = fill

            ws['B1'] = "Resumo"
            ws['B1'].font = header_font
            ws['B1'].fill = fill

            row_num = 2
            for email in emails_data:
                ws[f'A{row_num}'] = email["Assunto"]
                ws[f'B{row_num}'] = email["Resumo"]
                row_num += 1

            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column].width = adjusted_width

            wb.save("emails_formatted.xlsx")
            print("Arquivo Excel 'emails_formatted.xlsx' criado com sucesso.")
            time.sleep(10)
            move_file_to_directory("data", "./emails_formatted.xlsx")

    except Exception as e:
        print(f"Erro ao tentar extrair textos dos emails: {str(e)}")
        driver.quit()


def close_browser():
    input("Pressione qualquer tecla para fechar o navegador...")
    driver.quit()


if __name__ == "__main__":
    login()
    extract_all_emails_text()
    close_browser()
