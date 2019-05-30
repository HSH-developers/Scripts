import tkinter as tk
from tkinter import filedialog
import pandas as pd
from PIL import ImageGrab
import win32com.client as win32
import os
import requests

root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    dir = os.getcwd()
    # path_destination = "{0}\\{1}\\".format(dir, "Destination")
    
    import_file_path = filedialog.askopenfilename()

    wb = get_workbook_excel_file(import_file_path)
    upload_data(import_file_path, wb)
    
def get_rows_data(filepath) :
    df = pd.read_excel(filepath)
    data = []
    for index, row in df.iterrows() :

        num = row.No
        name = row.Name
        brand = row.Brand
        price = row.Price
        category = row.Category
        ftype = row.Type
        dimension = row['Dimension (mm)']
        description = row.Description
        row_data = [num, name, brand , price, category ,ftype, dimension, description]
        data.append(row_data)
    
    return data
    
def get_workbook_excel_file(filename):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filename)
    return wb


def DiscartImage(shape):
    if (shape.Height >= 40 and shape.Height <= 70 and shape.Width >= 130 and shape.Width <= 160):
        return True
    if (shape.Height >= 45 and shape.Height <= 50 and shape.Width >= 45 and shape.Width <= 50):
        return True
    return False

def extract_file_images_and_models(workbook):
    filename = workbook.Name

    id_ordem = filename.replace('.xlsm', '')

    for sheet in workbook.Worksheets:
        
        listShape = enumerate(sheet.Shapes)
        images = [shape for i, shape in enumerate(sheet.Shapes) if shape.Name.startswith('Picture')]
        models = [shape for i, shape in enumerate(sheet.Shapes) if shape.Name.startswith('3D Model')]

        return images, models
        

def upload_data(filepath, workbook) :

    rows_data = get_rows_data(filepath)
    images, models = extract_file_images_and_models(workbook)
    
    model_dict = {}
    script_dir = os.path.dirname(__file__) #<-- absolute dir the script is in
    for filepath in os.listdir('./models') :
        furniture_name = os.path.splitext(filepath)[0]
        
        rel_path = "models/" + filepath
        abs_file_path = os.path.join(script_dir, rel_path)
        with open(abs_file_path, "rb") as data :
            model_dict[furniture_name] = data

    print(model_dict)

    
    for i in range(len(rows_data)) :
        num, name, brand , price, category ,ftype, dimension, description = rows_data[i]
        imageShape = images[i]
        modelShape = models[i]
        
        try:
            imageShape.Copy()
            image = ImageGrab.grabclipboard()

         
            model = model_dict[name]
            print(model)
            # post_data = {
            #     'furnitureName' : name,
            #     'furnitureBrand' : brand,
            #     'furnitureType' : ftype,
            #     'furnitureCategory' : category,
            #     'furnitureDimension' : dimension,
            #     'furniturePrice' : price,
            #     'image' : image,
            #     'model' : model
            # }
            
            # print(post_data)

        except:
            continue


browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)

root.mainloop()