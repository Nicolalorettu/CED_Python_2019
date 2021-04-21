from PIL import Image
import pytesseract

img = Image.open(r'C:\Users\Nicola Pc\Desktop\unnamed.png')

print(pytesseract.image_to_string(img, config='-psm 6'))
