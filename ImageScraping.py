from bs4 import BeautifulSoup
import urllib.request as urllib
from requests import get
import os

# use this image scraper from the location that 
#you want to save scraped images to

# typeFurniture = {'concrete-furniture-singapore' : ['concrete-dining-table','concrete-coffee-table','concrete-side-table','concrete-seating-arrangements','concrete-home-decor'], 
#                  'custom-marble-furniture-singapore' : ['marble-dining-table-singapore','coffee-table-marble']
# }

image_base_url = 'https://www.martlewood.com/product-category/'
def make_soup(url):
    response = get(url)
    page_html = BeautifulSoup(response.text, 'html.parser')
    return page_html

def get_images(url):
    product_url = image_base_url + 'concrete-furniture-singapore/' + 'concrete-dining-table/'
    print(product_url)
    soup = make_soup(product_url)
    #this makes a list of bs4 element tags
    images = [img for img in soup.findAll('img')]
    directory = 'images/tables'
    print (str(len(images)) + " images found.")
    print ('Downloading images to current working directory.')
    for image in images :
        print(image)
    #compile our unicode list of image links
    image_links = [image.get('src') for image in images]
    for image in image_links:
        if not os.path.isdir(directory) :
            os.makedirs(directory)
        filename=image.split('/')[-1]
        urllib.urlretrieve(image, directory + filename)
    return image_links


images = get_images(image_base_url)
