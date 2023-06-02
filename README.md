# Add_text_in_image

from PIL import Image, ImageFont, ImageDraw
import numpy as np
import textwrap
from openpyxl import load_workbook
from moviepy.editor import *
import os


# Iterate over the rows
def read_excel_row(row_num, image):
    # input text
    Input_Text = []
    for col in range(2, 3):
        cell_value = worksheet.cell(row=row_num, column=col).value
        Input_Text.append(cell_value)
    # x-coordinate
    X = []
    for col in range(3, 4):
        cell_value = worksheet.cell(row=row_num, column=col).value
        X.append(cell_value)
    # y-coordinate
    Y = []
    for col in range(4, 5):
        cell_value = worksheet.cell(row=row_num, column=col).value
        Y.append(cell_value)
    # font size
    Font_Size = []
    for col in range(5, 6):
        cell_value = worksheet.cell(row=row_num, column=col).value
        Font_Size.append(cell_value)
    # font style
    Font_Style = []
    for col in range(6, 7):
        cell_value = worksheet.cell(row=row_num, column=col).value
        Font_Style.append(cell_value)
    # final file name
    OutPut_file_name = []
    for col in range(7, 8):
        cell_value = worksheet.cell(row=row_num, column=col).value
        OutPut_file_name.append(cell_value)
    # folder name in which file will be saved
    Folder = []
    for col in range(8, 9):
        cell_value = worksheet.cell(row=row_num, column=col).value
        Folder.append(cell_value)

    # Set the duration of the image clip (in seconds)
    image_clip = ImageClip(np.array(image)).set_duration(5)

    # Define the text
    text = Input_Text[0]

    # Define the square area coordinates
    x, y = X[0], Y[0]  # Top-left corner coordinates

    # Define the square area width and height
    width = 300
    height = 300

    # Set the font properties
    font_size = Font_Size[0]
    font_color = (255, 0, 0)  # White color
    font_path = Font_Style[0]  # Replace with the path to your font file
    # print(font_path)
    if font_path == "AF":
        font_path = "Aleo-Regular.ttf"
        font = ImageFont.truetype(font_path, font_size)
    elif font_path == "EN":
        font_path = "BerkshireSwash-Regular.ttf"
        font = ImageFont.truetype(font_path, font_size)
    else:
        font_path = "AlexBrush-Regular.ttf"
        font = ImageFont.truetype(font_path, font_size)

    # Create a blank image with the same dimensions as the original image
    text_image = Image.new("RGBA", image.size, (0, 0, 0, 0))

    # Create a draw object
    draw = ImageDraw.Draw(text_image)

    # Wrap the text into multiple lines to fit within the square area
    lines = textwrap.wrap(text, width=15)  # Adjust the width parameter as needed

    # Calculate the height of each line
    line_height = font.getsize("hg")[1]  # Assuming 'hg' is the highest and lowest characters

    # Calculate the total height of the text
    total_height = len(lines) * line_height

    # Calculate the starting Y-coordinate for the text to be vertically centered
    start_y = y + (height - total_height) // 2

    # Draw the text line by line within the square area
    for line in lines:
        line_width, line_height = font.getsize(line)
        text_position = ((x + width - line_width) // 2, start_y)
        draw.text(text_position, line, font=font, fill=font_color)
        start_y += line_height

    # Convert the modified image back to a MoviePy clip
    modified_image_clip = ImageClip(np.array(text_image)).set_duration(5)

    # Composite the original image clip and the modified image clip
    composite_clip = CompositeVideoClip([image_clip, modified_image_clip])

    # Set the duration of the composite clip
    composite_clip = composite_clip.set_duration(5)

    # Generate the final composite image
    composite_image = composite_clip.to_ImageClip()

    # Create a folder with the same name as the image (if it doesn't exist)
    folder_name = Folder[0]
    os.makedirs(folder_name, exist_ok=True)

    # Save the composite image as a PNG file in the folder
    output_path = os.path.join(folder_name, OutPut_file_name[0])
    composite_image.save_frame(output_path, t=0, withmask=False)


# reading the excel file
excel_file_sheets = "data.xlsx"
workbook = load_workbook("data.xlsx")

# Select the active worksheet
worksheet = workbook.active

# Find the number of rows with data in the excel sheet
num_rows = 0
for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row):
    if any(cell.value for cell in row):
        num_rows += 1

# Iterate over the rows
for row_number in range(2, num_rows + 1):
    # Load the image
    Input_File_Name = []
    for col in range(1, 2):
        cell_value = worksheet.cell(row=row_number, column=col).value
        Input_File_Name.append(cell_value)

    # Load the image using PIL
    try:
        image = Image.open(Input_File_Name[0])
        # Process the row
        read_excel_row(row_number, image)
    except:
        print("image file is not available for the row number: - ", row_number, Input_File_Name[0])
        with open('error_log.txt', 'a') as f:
            data = str(Input_File_Name[0]) + "--->image file is not available for the row_number--->" + str(row_number)
            f.write(data + "\n")
