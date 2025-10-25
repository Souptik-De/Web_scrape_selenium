import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


START_ROLL = 12624019001
END_ROLL = 12624019064
SEMESTER = '2'
URL = "http://111.93.160.42:8084/stud25e.aspx"


chrome_options = Options()
chrome_options.add_argument('--headless')  # Comment this line if you want to see the browser
chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)


data = []

def scrape_roll(roll):
    try:
        driver.get(URL)
        wait.until(EC.presence_of_element_located((By.NAME, "roll")))
        
        # Fill roll number and semester
        driver.find_element(By.NAME, "roll").send_keys(str(roll))
        Select(driver.find_element(By.NAME, "sem")).select_by_value(SEMESTER)
        driver.find_element(By.ID, "Button1").click()

        
        wait.until(EC.presence_of_element_located((By.ID, "lblname")))
        time.sleep(1.5)

        name = driver.find_element(By.ID, "lblname").text.replace("Name : ", "")
        rollno = driver.find_element(By.ID, "lblroll").text.replace("Roll No : ", "")
        sgpa_odd = driver.find_element(By.ID, "lblbottom1").text.split(":")[-1].strip()
        sgpa_even = driver.find_element(By.ID, "lblbottom2").text.split(":")[-1].strip()
        ygpa = driver.find_element(By.ID, "lblbottom3").text.split(":")[-1].strip()

        
        subjects = driver.find_elements(By.XPATH, "//table[contains(@style,'border-style: solid;height:500px;')]/tbody/tr")[1:]
        for subj in subjects:
            cols = subj.find_elements(By.TAG_NAME, "td")
            subject_data = {
                "Roll": rollno,
                "Name": name,
                "Subject Code": cols[0].text.strip(),
                "Subject Name": cols[1].text.strip(),
                "Grade": cols[2].text.strip(),
                "Points": cols[3].text.strip(),
                "Credit": cols[4].text.strip(),
                "Credit Points": cols[5].text.strip(),
                "SGPA Odd": sgpa_odd,
                "SGPA Even": sgpa_even,
                "YGPA": ygpa
            }
            data.append(subject_data)

        print(f"✅ Scraped: {roll}")
    except Exception as e:
        print(f"Error with roll {roll}: {e}")
    finally:
        time.sleep(random.uniform(2.5, 5))  # polite delay


for roll in range(START_ROLL, END_ROLL + 1):
    if roll == 12624019031 or roll == 12624019042 or roll == 12624019008 :
        print("Skipping roll as it is not available.")
        continue
    scrape_roll(roll)

# --- SAVE TO EXCEL ---
df = pd.DataFrame(data)

# Pivot subjects into columns with grades
pivot_df = df.pivot_table(
    index=["Roll", "Name", "SGPA Odd", "SGPA Even", "YGPA"],
    columns="Subject Code",
    values="Points",
    aggfunc="first"
).reset_index()


pivot_df.to_excel("heritage_results_by_student.xlsx", index=False)

print("Excel file saved as heritage_results_by_students.xlsx")


driver.quit()
