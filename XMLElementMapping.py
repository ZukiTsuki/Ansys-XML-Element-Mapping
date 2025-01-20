import xml.etree.ElementTree as ET
from openpyxl import load_workbook 
from lxml import etree
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException

def unitconversions(units,contents):
    match units:
        case units if "g/cc" in units:
                contents=contents*1000
        case units if "lb/ft³" in units:
                contents=contents*16.0185  
        case units if "lb/gal" in units:
                contents=contents*119.826
        case units if "lb/in³" in units:
                contents=contents*27679.904710191
        case units if "N/m³" in units:
                contents=contents*0.101971621
        case units if "MPa" in units:
                contents=contents*(10**6)
        case units if  "GPa" in units:
                contents=contents*(10**9)
        case units if  "J/g-°C" in units:
                contents=contents*1000                
        case _:
                contents=contents
    return contents

materialsearch=input("Enter search keyword of material: ")
material=input("Enter name of material: ")
#materialsearch="Aluminum 2117-T4"
#material="Aluminum 2117-T4"
url='https://www.matweb.com/Search/MaterialGroupSearch.aspx?GroupID=202'
service = Service()
options = webdriver.ChromeOptions()
options.add_argument('--load-extension={}'.format(r'C:\Users\kazuk\AppData\Local\Google\Chrome\User Data\Default\Extensions\cjpalhdlnbpafiamejdnhcphjbkeiagm\1.59.0_1'))
options.add_argument("--headless")
driver = webdriver.Chrome(service=service, options=options)
driver.get(url)
driver.find_element("name","ctl00$txtQuickText").send_keys(material)
driver.execute_script('btnQuickTextSearch_Click()')
#WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,"/html/body/form[2]/div[4]/table[1]/tbody/tr/td[3]/div[2]/table[4]/tbody/tr[2]/td/strong/span")))
#/html/body/form[2]/div[4]/table/tbody/tr/td[2]/div/table[3]/tbody/tr[188]/td[3]/a
i=2

while(True):
    try:
        MatW=driver.find_element(By.XPATH,"/html/body/form[2]/div[4]/table/tbody/tr/td[2]/div/table[3]/tbody/tr["+str(i)+"]/td[3]/a").text
        if MatW==material:
              WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/form[2]/div[4]/table/tbody/tr/td[2]/div/table[3]/tbody/tr["+str(i)+"]/td[3]/a"))).click()
              break
        else:
           i=i+1
           continue    
    except NoSuchElementException:
            try:
               if driver.find_element(By.XPATH,"html/body/form[2]/div[4]/table/tbody/tr/td[2]/div/strong").text=="No Materials were found using the selected search criteria....":
                   print("Material not found")
                   exit()
            except NoSuchElementException:
                  driver.find_element(By.LINK_TEXT, '[Next Page]').click()
                  i=2
                  continue
      
while(True):
  try:
      WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[6]/td/div/div[1]/span[2]/a/small")))
      break
  except NoSuchElementException:
      pass

tree = ET.parse(r'E:\Courses\SCRD\Vibration Analysis\Aluminum 6151 T6.xml')
root = tree.getroot()
tree.find('Notes').text= (#driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[3]/td").text + (
     #"\n "+"    "+driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[3]/td/p").text +
     #"\n "+"   "+driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[3]/td/p/br").text + "\n "
     "\n " + "   Table name: MaterialUniverse " + 
     "\n " + "   Database edition: Metals plus" + 
     "\n " + "   Date exported: "+ datetime.today().strftime('%Y-%m-%d %H:%M') + " \n  "
  )
#root[0][0][0][0][0].text= driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[3]/td").text + (
     #"\n "+"    "+driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[3]/td/p").text + "\n "
     #"\n "+"   "+driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[1]/tbody/tr[3]/td/p/br").text + "\n "
#)
material=driver.find_element(By.XPATH,"/html/body/form[2]/div[4]/div/table[1]/tbody/tr[1]/th").text
root[1][0][0][0][0].text=material
root[1][0][0][0][1].text= ("Materials data from MatWeb"+"\n "+
      "Record name: "+ material + "\n "+
      "Table name: MaterialUniverse"+"\n "+"\n "+
      material+"\n "+"         "
    )


rows = len(driver.find_elements(By.XPATH,"/html/body/form[2]/div[4]/div/table[2]/tbody/tr"))+1


for i in range(3,rows):
    try:
      content=str(driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[2]/tbody/tr["+str(i)+"]/td[1]").text)
      unit=(driver.find_element(By.XPATH, "/html/body/form[2]/div[4]/div/table[2]/tbody/tr["+str(i)+"]/td[2]").text).split(" ")
      try:
        units=unit[1]
      except IndexError:
        pass
      try:
        contents=float(unit[0])
      except ValueError:
        contents=unit[0]
    except NoSuchElementException:
         continue
    match content:
       case content if "Density" in content:
                root[1][0][0][0][3][1][0].text = str(unitconversions(units,contents))

       case content if "Tensile Strength, Ultimate" in content:
                root[1][0][0][0][7][1][0].text = str(int(unitconversions(units,contents)))

       case content if "Tensile Strength, Yield" in content:
                root[1][0][0][0][6][1][0].text = str(int(unitconversions(units,contents))) #General
                root[1][0][0][0][5][2][0].text = str(int(unitconversions(units,contents))) #Isotropic Hardening

       case content if "Modulus of Elasticity" in content:
                root[1][0][0][0][4][2][0].text = str(int(unitconversions(units,contents)))

       case content if "Poissons Ratio" in content:
                root[1][0][0][0][4][3][0].text = str(unit[0])

       case content if "Shear Modulus" in content:
                root[1][0][0][0][5][2][1].text = str(int(unitconversions(units,contents)))

       case content if "Specific Heat Capacity" in content:
                root[1][0][0][0][11][3][0].text = str(unitconversions(units,contents))

       case content if "Thermal Conductivity" in content:
                root[1][0][0][0][10][3][0].text = str(unitconversions(units,contents))

       case content if "CTE, linear" in content:
                root[1][0][0][0][9][3][0].text = str(unit[0])
    continue  
fpath='E:\\Ansys Material Library\\'
file=str(material)
extension='.xml'
filepath=fpath+file+extension
tree.write(filepath, encoding="utf-8", xml_declaration=True)

tree=etree.parse(filepath)
root = tree.getroot()
for element in root.xpath('//*[not(node())][not(count(./@*))>0]'):
    element.getparent().remove(element)
tree.write(filepath, encoding="utf-8", xml_declaration=True)
driver.close() 

