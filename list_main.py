# Orange-Find-Points List V2.0 Service
# Copyright (c) 2021 Wxbxbwjb. All rights reserved. > Oranger!xuan:)

"""
OFP List V2.0 Service EULA（使用此服务后则自动视为同意）：
1. 用户信息保护与免责声明
   - 合规性：用户必须遵守所有适用的法律法规，包括但不限于隐私保护法、个人信息保护法等。用户不得将用户信息用于任何违法活动，或以任何方式侵犯他人的合法权益。
   - 信息传播与分发：用户不得传播、分发或泄露任何用户信息，包括但不限于成绩单、个人联系信息等。
   - 道德与法律标准：本服务旨在促进学习交流和技术分享，用户必须遵守所有适用的道德标准和法律法规。严禁将用户信息用于任何违法或不道德的行为，如传播、嘲讽、歧视等。用户应理解，任何擅自用于不正当行为所造成的后果将由用户自行承担。
2. 责任限制
   - 在任何情况下，服务提供者不对因用户违反本协议而产生的任何直接、间接、偶然、特殊或后果性损害负责，包括但不限于数据或利润损失。
3. 协议终止
   - 任何个人或实体在违反本协议的任何条款时，其使用许可将自动终止。在此情况下，用户必须立即停止使用本服务，并销毁所有用户信息的副本。
4. 法律适用与争议解决
   - 本协议的解释和效力适用相关法律。因本协议引起的任何争议应首先通过友好协商解决。若协商未能达成一致，任何一方均有权向有管辖权的法院提起诉讼。
5. 用户同意声明
   - 自用户使用本服务起，即表明用户同意了上述声明和本协议的所有条款，并承诺遵守所有适用的法律法规以及道德标准。
"""


import pyautogui
import openpyxl
import pyperclip
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
find_times = int(input('find times：'))
times = 0
ids = ["114514"]
names = ["Oranger"]
a = openpyxl.load_workbook('out.xlsx')
b = a.active
for i in range(find_times):
    edge = Options()
    edge.add_argument('--headless')
    service = Service('C:\\msedgedriver.exe')
    driver = webdriver.Edge(service=service, options=edge)
    driver.get("URL")
    wait = WebDriverWait(driver, 6)
    id2 = driver.find_element(By.NAME, "s_kaohao")
    id2.send_keys(ids[i])
    names2 = driver.find_element(By.NAME, "s_xingming")
    names2.send_keys(names[i])
    button = driver.find_element(By.ID, "yiDunSubmitBtn")
    button.click()
    pyautogui.sleep(3)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "case")))  
    table = driver.find_element(By.CLASS_NAME, "case")
    clipboard_content = ""
    tbody = table.find_element(By.TAG_NAME, "tbody")
    for c in tbody.find_elements(By.TAG_NAME, "tr"):
        cells = c.find_elements(By.TAG_NAME, "td")
        line_content = "\t".join(cell.text.strip() for cell in cells)
        clipboard_content += line_content + "\n"
    pyperclip.copy(clipboard_content)
    cb = clipboard_content.replace("\n", "")
    max_c = b.max_c
    text = cb
    rob = text.split('\n')
    for c in rob:
        cells = c.split('\t')
        b.append(cells)
    driver.quit()
    a.save('out.xlsx')
    if find_times == times:
        pyautogui.alert('OK', 'Oranger')
    else:
        pyautogui.sleep(1)
a.save('out.xlsx')
driver.quit()
a.close()
