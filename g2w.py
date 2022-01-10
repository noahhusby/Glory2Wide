import os
import numpy as np
import cv2
import shutil
import zipfile
from alive_progress import alive_bar
from pptx import Presentation
from pptx.util import Pt
from pptx.shapes.picture import Picture
import json

def crop_image(im):
    img = cv2.imread(im)
    img = img[40:-70,0:-1]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = 255*(gray < 128).astype(np.uint8)
    gray = cv2.morphologyEx(gray, cv2.MORPH_OPEN, np.ones((2, 2), dtype=np.uint8))
    coords = cv2.findNonZero(gray)
    x, y, w, h = cv2.boundingRect(coords)
    w = w + x + 100
    rect = img[y:y+h, 0:w]
    cv2.imwrite(im, rect)
    
def extract_all_melodies():
    print("\n1. Extracting all melodies located in \"hymns\" folder.")
    melodies = {}
    for subdir, dirs, files in os.walk(r'hymns'):
        for filename in files:
            filepath = subdir + os.sep + filename
            if filepath.endswith(".ppt") and "Melody" in filepath:
                melodies[filename] = filepath
                    
    with alive_bar(len(melodies)) as bar:
        for key, value in melodies.items():
            with zipfile.ZipFile(value,"r") as zip_ref:
                zip_ref.extractall("temp/" + key.replace('.ppt', ''))
                bar()
    print("\n")
    
def generate_title_slide_information():
    print("2. Generating title slide information.")
    melodies = {}
    for subdir, dirs, files in os.walk(r'hymns'):
        for filename in files:
            filepath = subdir + os.sep + filename
            if filepath.endswith(".ppt") and "Melody" in filepath:
                melodies[filename.replace('.ppt', '')] = filepath
   
    with alive_bar(len(melodies)) as bar:
        for key, value in melodies.items():
            d = "temp/" + key
            prs = Presentation(value)
            shapes = prs.slides[0].shapes
            if len(shapes) < 2:
                shapes = prs.slides[1].shapes
            hymn_name = shapes[0].text
            hymn_number = shapes[1].text
            hymn_credits = shapes[2].text
            data = {}
            data['name'] = hymn_name
            data['number'] = hymn_number
            data['credits'] = hymn_credits
            with open(d + "/title.json", 'w') as outfile:
                json.dump(data, outfile)
            bar()
    print("\n")
    
def crop_images():
    print("3. Cropping melody images.")
    images = []
    for subdir, dirs, files in os.walk(r'temp'):
        for filename in files:
            filepath = subdir + os.sep + filename
            if filepath.endswith(".png"):
                images.append(filepath)
                    
    with alive_bar(len(images)) as bar:
        for file in images:
            crop_image(file)
            bar()
    print("\n")

def create_presentations():
    print("4. Generating final hymns.")
    temp = []
    for path in os.listdir("temp"):
        temp.append(path)
        
    with alive_bar(len(temp)) as bar:
        for path in temp:
            d = "temp/" + path
            prs = Presentation("template.pptx")
            shapes = prs.slides[0].shapes
            with open(d + "/title.json", 'r') as f:
                title_assets =  json.loads(f.read())
            name = shapes[0].text_frame.paragraphs[0]
            name.text = title_assets['name']
            name.font.size = Pt(60)
            name.font.bold = True
            number = shapes[1].text_frame.paragraphs[0]
            number.text = title_assets['number']
            number.font.size = Pt(30)
            number.font.italic = True
            credits = shapes[2].text_frame.paragraphs[0]
            credits.text = title_assets['credits']
            credits.font.size = Pt(14)
            credits.font.italic = True
            images = {}
            for subdir, dirs, files in os.walk(d):
                for filename in files:
                    filepath = subdir + os.sep + filename
                    if filepath.endswith(".png"):
                        images[int(filename.replace('image','').replace('.png', ''))] = filepath
            images = dict(sorted(images.items()))
            for key, value in dict(sorted(images.items())).items():
                    slide = prs.slides.add_slide(prs.slide_layouts[0])
                    picture = slide.shapes.add_picture(value, 0, 0, prs.slide_width)
                    calc_top_value = round((prs.slide_height - picture.height) / 2)
                    picture.top = calc_top_value            
            prs.save("out/" + path + "_wide.pptx")
            bar()
    print("\n")
    
def clean_up():
    print("5. Cleaning up.")
    temp = []
    for path in os.listdir("temp"):
        temp.append(path)
        
    with alive_bar(len(temp)) as bar:
        for path in temp:
            shutil.rmtree("temp/" + path)
            bar()
    shutil.rmtree("temp")
    print("\n")
    
if __name__ == "__main__":
    print("\nGlory2Wide by Noah Husby")
    if os.path.isdir("temp"):
        shutil.rmtree("temp")
    if os.path.isdir("out") is False:
        os.mkdir("out")
    if os.path.isdir("hymns") is False:
        os.mkdir("hymns")
        print("Please extract the Glory to God hymn pack into the \"hymns\" directory")
        quit()
    extract_all_melodies()
    generate_title_slide_information()
    crop_images()
    create_presentations()
    clean_up()
    print("Successfully converted hymns to wide format.\nThe exported hymns are in the \"out\" folder.\n")

    