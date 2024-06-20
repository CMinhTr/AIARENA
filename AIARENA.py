import sys,os
if os.getenv('COMPUTERNAME') == 'HCVIP-1':
    sys.path.insert(0,r'F:\Share-MayChu\005-File Chay Tool 2022')
else:
    sys.path.insert(0,r'\\HCVIP-1\Share-MayChu\005-File Chay Tool 2022')

from vinhthoai.Utils_Gmail import *
oExcel=Excel(win32gui.GetWindowText(win32gui.FindWindow('XLMAIN',None)),'title')

CotKetQua ="E"

if __name__=="__main__":
    for For1 in range(1,999):

        if oExcel._ExcelReadCell(f'{CotKetQua}{For1}')!=None:continue
        Reset_Modem()
        oExcel._ExcelBookSave()

        DD_ProfileChrome = oExcel._ExcelReadCell(f"A{For1}")
        if DD_ProfileChrome==None:sys.exit("Doc xong het DD_ProfileChrome")
        print(f"!!! STT {For1}:DD_ProfileChrome :  {DD_ProfileChrome}")
                        
        NguoiDung= oExcel._ExcelReadCell(f"B{For1}")
        if NguoiDung==None:sys.exit("Doc xong het Nguoi Dung")
        print(f"!!! STT {For1}:Nguoi Dung :  {NguoiDung}")

        Email = oExcel._ExcelReadCell(f"F{For1}")
        if Email == None:sys.exit("Doc xong het Email")
        print(f"!!! STT {For1}:Email :  {Email}")

        Password = oExcel._ExcelReadCell(f"G{For1}")
        if Password == None:sys.exit("Doc xong het Password")
        print(f"!!! STT {For1}:Password :  {Password}")

        FullName = oExcel._ExcelReadCell(f"H{For1}")
        if FullName == None:sys.exit("Doc xong het FullName")
        print(f"!!! STT {For1}:FullName :  {FullName}")

        UserName = oExcel._ExcelReadCell(f"Z{For1}")
        if UserName == None:sys.exit("Doc xong het UserName")
        print(f"!!! STT {For1}:UserName :  {UserName}")

        LinkRef = oExcel._ExcelReadCell(f"X{For1}")
        if LinkRef == None:sys.exit("Doc xong het LinkRef")
        print(f"!!! STT {For1}:LinkRef :  {LinkRef}")

        Chrome_Driver = Google_Chrome()
        Chrome_Driver.ChromeDriver_Kill(BROWER_DEFAUT)

        driver = Chrome_Driver.ChromeDriver_Setup(user_data_dir=DD_ProfileChrome, undetect_chrome=True)
        driver.maximize_window()
        gmail_client = Gmail_Client(email=Email,password=Password,driver=driver)
        try:
            gmail_client.CheckLogin_Gmail()
        except Exception as e:
            driver.quit()
            oExcel._ExcelWriteCell(str(e),f'{CotKetQua}{For1}')
            continue

        driver.get(LinkRef)

        try:
            element = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@placeholder="Email Address"]')))
            element.send_keys(Email)
            driver.find_element(By.XPATH,'//*[@type="submit"]').click()
            Verification = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@class="max-w-sm p1 text-center"]'))).text
            if Verification == 'Verification Sent':
                print('Đã gữi code đăng ký')
        except:
            oExcel._ExcelWriteCell('Error Refernal code added',f'{CotKetQua}{For1}')
            driver.quit()
            continue

        try:
            id_message = gmail_client.ReadIDMail_Gmail('Welcome To AI Arena Complete your login by clicking the link below Secure Login',60)
            driver.get(f'https://mail.google.com/mail/u/0/?fs=1&source=atom#all/{id_message}')
        except Exception as e:
            oExcel._ExcelWriteCell(str(e),f'{CotKetQua}{For1}')
            driver.quit()
            continue

        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            try:
                element = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[text()="Secure Login"]')))
                get_href = element.get_attribute('href')
                oExcel._ExcelWriteCell(get_href, f'AG{For1}')
                print(get_href)
                break
            except:
                time.sleep(1)
        else:
            oExcel._ExcelWriteCell('Error Get Href', f'AG{For1}')
            driver.quit()

        driver.get(get_href)
        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            try:
                element = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="AI Arena\'s Privacy Policy"]')))
                driver.execute_script("arguments[0].click();",element)
                break
            except:
                time.sleep(1)
        else:
            oExcel._ExcelWriteCell('I agree to AI Arena\'s Privacy Policy', f'{CotKetQua}{For1}')
            driver.quit()
            
        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            try:
                element = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="Harbor\'s Privacy Policy"]')))
                driver.execute_script("arguments[0].click();",element)
                break
            except:
                time.sleep(1)
        else:
            oExcel._ExcelWriteCell('I agree to Harbor\'s Privacy Policy', f'{CotKetQua}{For1}')
            driver.quit()   
            
        try:
            driver.find_element(By.XPATH,'//*[@type="submit"]').click()
        except:
            oExcel._ExcelWriteCell('Error submit',f'{CotKetQua}{For1}')
            driver.quit()
            continue

        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            try:
                element = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@placeholder="Username"]')))
                element.send_keys(UserName)
                break
            except:
                time.sleep(1)
        else:
            oExcel._ExcelWriteCell('Error UserName', f'{CotKetQua}{For1}')
            driver.quit()
        
        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            try:
                driver.find_element(By.XPATH,'//*[@type="submit"]').click()
                break
            except:
                time.sleep(1)
        else:
            oExcel._ExcelWriteCell('Error submit',f'{CotKetQua}{For1}')
            driver.quit()

        
       

        
        time.sleep(3)
        driver.get('https://hub.aiarena.io/modules/refer')
        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            try:
                element = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@class="relative !w-full py-2 px-4 btn-accent"]')))
                refernal = element.get_attribute('value')
                print(refernal)
                oExcel._ExcelWriteCell(refernal,f'{CotKetQua}{For1}')
                driver.quit()
                break
            except:
                time.sleep(1)
        else:
            oExcel._ExcelWriteCell('Error Refernal', f'{CotKetQua}{For1}')
            driver.quit()

    





