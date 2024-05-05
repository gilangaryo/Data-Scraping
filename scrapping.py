from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time

link = "https://www.tokopedia.com/search?st=&q=macbook&srp_component_id=02.01.00.00&srp_page_id=&srp_page_title=&navsource="
opsi = webdriver.ChromeOptions()
opsi.add_argument("--start-maximized")
servis = Service('chromedriver.exe')
driver = webdriver.Chrome(service=servis, options=opsi)
driver.get(link)
rentang = 1000

for i in range(1, 10):
    akhir = rentang * i 
    perintah = "window.scrollTo(0," + str(akhir) + ")"
    driver.execute_script(perintah)
    print("loading ke-" + str(i))
    time.sleep(1)  
time.sleep(7)

content = driver.page_source

list_nama,list_gambar,list_harga,list_link,list_terjual=[],[],[],[],[]

data = BeautifulSoup(content,'html.parser')
i = 1
for area in data.find_all('div', class_="css-llwpbs"):
    nama_element = area.find('div', class_="prd_link-product-name css-3um8ox")
    nama = nama_element.get_text().strip() if nama_element else None

    harga_element = area.find('div', class_="prd_link-product-price css-h66vau")
    harga = harga_element.get_text().strip() if harga_element else None
    
    gambar = area.find('img', class_="css-1q90pod")
    gambar_url = gambar['src'] if gambar and 'src' in gambar.attrs else None
    
    a_tag = area.find('a')  
    product_link = a_tag['href'] if a_tag and 'href' in a_tag.attrs else None
    
    terjual_element = area.find('span', class_="prd_label-integrity css-1sgek4h")
    terjual = terjual_element.get_text() if terjual_element else None
    
    if nama:
        print(i)
        print("Nama Produk:", nama)
        print("Harga Produk:", harga)
        print("Link Gambar:", gambar_url)
        print("Link Produk:", product_link)
        print("Terjual:", terjual)
        i += 1
        print("-------")

    list_nama.append(nama)
    list_gambar.append(gambar_url)
    list_harga.append(harga)
    list_link.append(product_link)
    list_terjual.append(terjual)

gambar_hyperlinks = [f'=HYPERLINK("{url}", "Foto Produk {index+1}")' for index, url in enumerate(list_gambar)]

df = pd.DataFrame({
    'Nama': list_nama,
    'Gambar': gambar_hyperlinks ,
    'Harga': list_harga,
    'Link': list_link,
    'Terjual': list_terjual
})

writer = pd.ExcelWriter('data.xlsx', engine='openpyxl')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer._save()  


