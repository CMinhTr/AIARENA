import sys,os
if os.getenv('COMPUTERNAME') == 'HCVIP-1':
    sys.path.insert(0,r'F:\Share-MayChu\005-File Chay Tool 2022')
else:
    sys.path.insert(0,r'\\HCVIP-1\Share-MayChu\005-File Chay Tool 2022')

from vinhthoai.Utils_Gmail import *
oExcel=Excel(win32gui.GetWindowText(win32gui.FindWindow('XLMAIN',None)),'title')
CotKetQua ="D"

def loginWallet():
    Timeint = time.time(); TimeOut= 90;Error = 0
    while time.time() < Timeint + TimeOut:
            
        current = driver.current_window_handle
        driver.switch_to.window(current)

        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if handle != current:
                driver.close()

        driver.get('chrome-extension://nkbihfbeogaeaoehlefnkodbefgpgknn/home.html#unlock')
        
        try:
            print('Nhập mật khẩu')
            element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/form/div/div/input")))
            element.send_keys('Hungcuong@6789')
            element.send_keys(Keys.ENTER)
            time.sleep(3)
            driver.find_element(By.XPATH,'//*[@class="app-header__metafox-logo--horizontal"]')
            print("Đăng nhập thành công")
            break
        except: 
            print('Error Nhập mật khẩu')
    else:
        oExcel._ExcelReadCell(f'{CotKetQua}{For1}','Import Không Thành Công')
        print('Import Không Thành Công')
        
for For1 in range(1,999):

        if oExcel._ExcelReadCell(f'{CotKetQua}{For1}')!=None:continue
        oExcel._ExcelBookSave()
        Reset_Modem()
        DD_ProfileChrome = oExcel._ExcelReadCell(f"A{For1}")
        if DD_ProfileChrome==None:sys.exit("Doc xong het DD_ProfileChrome")
        print(f"!!! STT {For1}:DD_ProfileChrome :  {DD_ProfileChrome}")
        Chrome_Driver = Google_Chrome()
        Chrome_Driver.ChromeDriver_Kill(BROWER_DEFAUT)

        PasswordTwitch = oExcel._ExcelReadCell(f"AA{For1}")
        if PasswordTwitch == None:sys.exit("Doc xong het PasswordTwitch")
        print(f"!!! STT {For1}:PasswordTwitch :  {PasswordTwitch}")

        UserName = oExcel._ExcelReadCell(f"Z{For1}")
        if UserName == None:sys.exit("Doc xong het UserName")
        print(f"!!! STT {For1}:UserName :  {UserName}")

        driver = Chrome_Driver.ChromeDriver_Setup(user_data_dir=DD_ProfileChrome, undetect_chrome=True)
        driver.maximize_window()
        
        TimeInt = time.time(); TimeOut = 120
        while time.time() < TimeInt+TimeOut:
            driver.get('https://hub.aiarena.io/modules/quests/216')

            try:
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class=" btn-primary text-black uppercase font-black text-xl leading-5 rounded-full text-center py-5 px-10"]')))
                element.click()
            except:
                time.sleep(0.5)
                
            try:
                WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[text()="claimed"]')))
                break
            except:
                time.sleep(0.5)

            TimeInt = time.time(); TimeOut = 60
            while time.time() < TimeInt+TimeOut:
                try:
                    element = driver.find_element(By.XPATH,'//*[@class="mt-6 w-full flex gap-2 items-center justify-center py-3.5 rounded-2xl text-lg leading-[18px] bg-[#5865F2]"]')
                    authorize_discord = element.get_attribute('href')
                    driver.get(authorize_discord)
                    Authorize = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'(//*[@class="contents_dd4f85"])[2]')))
                    Authorize.click()
                    time.sleep(3)
                    WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH,'(//*[text()="Social Quests"])[1]')))
                    print('Authorize Discord OK')
                    break
                except:
                    time.sleep(0.5)
            else: 
                print('Authorize Discord Không Thành Công')
                
            try:
                driver.get('https://discord.gg/aiarenaplaytest')
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class="contents_dd4f85"]')))
                element.click()
                time.sleep(3)
            except:
                time.sleep(0.5)
        
        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            driver.get('https://hub.aiarena.io/modules/quests/219')

            try:
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class="fa-solid fa-circle-check"]')))
                element.click()
            except:
                time.sleep(0.5)
            try:
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class=" btn-primary text-black uppercase font-black text-xl leading-5 rounded-full text-center py-5 px-10"]')))
                element.click()
            except:
                time.sleep(0.5)
                
            try:
                WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[text()="claimed"]')))
                break
            except:
                time.sleep(0.5)
            
            try:
                element = driver.find_element(By.XPATH,'//*[@class="mt-6 w-full flex gap-2 items-center justify-center py-3.5 rounded-2xl text-lg leading-[18px] bg-[#4EA2E3]"]')
                authorize_twitter = element.get_attribute('href')
                driver.get(authorize_twitter)
                Authorize = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class="submit button selected"]')))
                Authorize.click()
                time.sleep(3)
                print('Authorize Twitter OK')
            except:
                time.sleep(0.5)
        
            try:
                driver.get('https://twitter.com/aiarena_')
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@aria-label="Following @aiarena_"]')))
                print('Fowllow aiarena_ OK')
            except: 
                try:
                    element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@aria-label="Follow @aiarena_"]')))
                    element.click()
                    print('Fowllow aiarena_ OK')
                except: time.sleep(0.5)

        TimeInt = time.time(); TimeOut = 60
        while time.time() < TimeInt+TimeOut:
            driver.get('https://hub.aiarena.io/modules/quests/220')

            try:
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class="fa-solid fa-circle-check"]')))
                element.click()
            except:
                time.sleep(0.5)
            try:
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@class=" btn-primary text-black uppercase font-black text-xl leading-5 rounded-full text-center py-5 px-10"]')))
                element.click()
            except:
                time.sleep(0.5)
                
            try:
                WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[text()="claimed"]')))
                break
            except:
                time.sleep(0.5)
            
            try:
                driver.get('https://twitter.com/arenaxlabs')
                element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@aria-label="Following @arenaxlabs"]')))
                print('Fowllow aiarena_ OK')
            except: 
                try:
                    element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@aria-label="Follow @arenaxlabs"]')))
                    element.click()
                    print('Fowllow aiarena_ OK')
                except: time.sleep(0.5)

        try:
            element = driver.find_element(By.XPATH,'(//*[@class="h4"])[2]').text
            oExcel._ExcelWriteCell(element,f'{CotKetQua}{For1}')
        except: 
            time.sleep(0.5)
            oExcel._ExcelWriteCell('Error Point',f'{CotKetQua}{For1}')
        driver.get('https://www.twitch.tv/signup?lang=en')

        try:
            element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="signup-username"]')))
            element.send_keys(UserName)
        except:
            time.sleep(0.5)
        try:
            element = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="password-input"]')))
            element.send_keys(PasswordTwitch)
        except:
            time.sleep(0.5)
            
        droplist = Select(driver.find_element(By.XPATH,'//*[@aria-label="Select your birthday month"]'))
        droplist.select_by_value('05')
        select = droplist.first_selected_option
        print('Select first: ' + select)
        input('---')
