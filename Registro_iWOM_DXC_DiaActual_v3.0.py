# ===========================================================
# IWOM - Automatic Workday Registration Script
# File Name Registro_iWOM_DXC_DiaActual_v3.0.py
# Author    César Álvaro Blázquez
# Email     cesar.alvaro-blazquez@dxc.com
# Role      Wintel L4 System Administrator – DXC for Logista
#
# Description
# Automates login to the IWOM portal using corporate credentials and
# registers daily working hours according to the weekday. Generates
# a daily log of actions and keeps only the last 30 entries.
#
# Usage Instructions
# Before running the script, open the file and replace the hardcoded
# user and password with your own DXC credentials
#   your_email@dxc.com
#   your_password
#
# Scheduled Task Notes
# This script is intended to be executed via Windows Task Scheduler
# at 1045 AM (can be modified) only if the user is logged in.
# If the user is not working that day, the script won't run.
# In that case, manually access IWOM the next day and mark the previous
# day as Not Available with a reason.
# The author assumes no responsibility for misuse of this tool or incorrect time tracking.
#
# License
# This script may be used, modified, forwarded or published
# as long as the original author's name (César Álvaro Blázquez)
# is kept in the code comments andor related documentation.
#
# Version History
# v3.1 - 2025-05-16  Added GitHub-ready version with credential placeholders
# ===========================================================

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from datetime import datetime
from colorama import init, Fore
import time
import sys
import os
import ctypes
import contextlib

init(autoreset=True)

# Bring console to front
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 9)
ctypes.windll.user32.SetForegroundWindow(ctypes.windll.kernel32.GetConsoleWindow())

log_dir = rCScriptsIWOMLogs
os.makedirs(log_dir, exist_ok=True)
timestamp = datetime.now().strftime(%Y%m%d_%H%M)
logfile = os.path.join(log_dir, fiwom_log_{timestamp}.txt)

def log(msg, tipo=info)
    now = datetime.now().strftime(%H%M%S)
    line = f[{now}] {msg}
    if tipo == info
        print(Fore.CYAN + line)
    elif tipo == ok
        print(Fore.GREEN + line)
    elif tipo == warn
        print(Fore.YELLOW + line)
    elif tipo == error
        print(Fore.RED + line)
    with open(logfile, a, encoding=utf-8) as f
        f.write(line + n)

def clean_old_logs(directory, keep=30)
    logs = sorted([f for f in os.listdir(directory) if f.startswith(iwom_log_) and f.endswith(.txt)])
    while len(logs)  keep
        old_log = logs.pop(0)
        try
            os.remove(os.path.join(directory, old_log))
        except Exception as e
            log(fFailed to delete old log {old_log} - {e}, warn)

try
    clean_old_logs(log_dir)
    log(Starting IWOM script..., info)
    log(NOTE Technical Edge messages in white text can be ignored., warn)

    options = Options()
    options.add_argument(--inprivate)
    options.add_argument(window-size=1400,900)

    log(Launching Edge browser..., info)
    service = Service(EdgeChromiumDriverManager().install())
    with open(os.devnull, 'w') as fnull, contextlib.redirect_stderr(fnull)
        driver = webdriver.Edge(service=service, options=options)

    driver.set_window_position(900, 100)
    driver.set_window_size(1400, 900)

    log(Opening login page, info)
    driver.get(httpsportal.bpocenter-dxc.com)
    time.sleep(5)

    log(Clicking DXC Login, info)
    driver.find_element(By.CLASS_NAME, boton_dxc).click()
    time.sleep(5)

    log(Entering user email, info)
    driver.find_element(By.ID, i0116).send_keys(your_email@dxc.com)  # Replace with your DXC email

    log(Clicking Next, info)
    driver.find_element(By.ID, idSIButton9).click()
    time.sleep(5)

    log(Entering password, info)
    driver.find_element(By.ID, i0118).send_keys(your_password)  # Replace with your DXC password
    time.sleep(5)

    log(Clicking Sign in, info)
    driver.find_element(By.ID, idSIButton9).click()

    log(Waiting for 2FA (up to 55 seconds)..., info)
    found = False
    for i in range(55)
        try
            btn = driver.find_element(By.XPATH, h4[text()='iWom Corp']ancestorbutton)
            if btn.is_displayed() and btn.is_enabled()
                found = True
                break
        except
            pass
        print(f...waiting {i+1}55 seconds, end=r)
        time.sleep(1)

    if not found
        log(iWom Corp button not found., error)
        log(2FA might have failed or been skipped., error)
        print(Fore.YELLOW + Press ENTER to close the browser...)
        input()
        driver.quit()
        sys.exit()

    log(Opening IWOM registration page, info)
    btn.click()
    driver.get(httpsiwomwa1.bpocenter-dxc.comes-corpappJornadaReg_jornada.aspx)
    time.sleep(5)

    day = datetime.today().weekday()
    if day in [0, 1, 2, 3]  # Monday–Thursday
        h_start, m_start, h_end, m_end, h_effect = 8, 30, 18, 30, 0822
    elif day == 4  # Friday
        h_start, m_start, h_end, m_end, h_effect = 8, 0, 15, 0, 0700
    else
        log(Weekend or non-working day., warn)
        print(Fore.YELLOW + Press ENTER to close the browser...)
        input()
        driver.quit()
        sys.exit()

    log(Filling form fields..., info)
    Select(driver.find_element(By.ID, ctl00_Sustituto_d_hora_inicio1)).select_by_value(h_start)
    Select(driver.find_element(By.ID, ctl00_Sustituto_D_minuto_inicio1)).select_by_value(m_start)
    Select(driver.find_element(By.ID, ctl00_Sustituto_d_hora_final1)).select_by_value(h_end)
    Select(driver.find_element(By.ID, ctl00_Sustituto_d_minuto_final1)).select_by_value(m_end)

    field = driver.find_element(By.ID, ctl00_Sustituto_T_efectivo)
    field.clear()
    field.send_keys(h_effect)

    log(Submitting form..., info)
    submit = driver.find_element(By.ID, ctl00_Sustituto_Btn_Guardar)
    driver.execute_script(arguments[0].click();, submit)

    log(Workday registered successfully., ok)
    print(Fore.YELLOW + Press ENTER to keep the window open or close manually...)
    input()

except Exception as e
    log(ERROR, error)
    log(str(e), error)
    print(Fore.YELLOW + Press ENTER to close the browser...)
    input()

finally
    driver.quit()
    log(Browser closed., info)
