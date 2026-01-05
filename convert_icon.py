from PIL import Image
img = Image.open("icon.png")
img.save("icon.ico", format='ICO')
print("Converted icon.png to icon.ico")
