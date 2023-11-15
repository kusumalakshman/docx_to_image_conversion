import docx
from PIL import Image, ImageDraw, ImageFont
import textwrap
import ipywidgets as widgets
from IPython.display import display, clear_output
from IPython.display import FileLink  # Add this import
import io
import os

# Define a function to process the uploaded DOCX file and create an image
def process_uploaded_docx(b):
    with output:
        clear_output()  # Clear the previous output

        # Check if any file is uploaded
        if len(file_upload.value) == 0:
            print("Please upload a DOCX file.")
        else:
            uploaded_file = list(file_upload.value.values())[0]['content']

            # Convert the uploaded content to a bytes object
            docx_content = uploaded_file

            # Open the DOCX content directly
            doc = docx.Document(io.BytesIO(docx_content))

            # Create an image with the extracted text
            img_width, img_height = 600, 400
            background_color = (255, 255, 255)
            text_color = (0, 0, 0)
            font_size = 12

            image = Image.new('RGB', (img_width, img_height), background_color)
            draw = ImageDraw.Draw(image)

            y = 10

            for paragraph in doc.paragraphs:
                text = paragraph.text

                text_lines = textwrap.wrap(text, width=90)

                for line in text_lines:
                    font = ImageFont.load_default()  # Use a default font
                    draw.text((10, y), line, fill=text_color, font=font)
                    y += font.getsize(line)[1]

            # Display the image
            display(image)

            # Save the image to a file
            image_path = 'output_image.png'
            image.save(image_path)

            # Provide a download link for the image
            display(FileLink(image_path, result_html_prefix="Click here to download the image: "))

# Create widgets
file_upload = widgets.FileUpload(accept=".docx", description="Upload DOCX file")
run_button = widgets.Button(description="Convert to Image")
output = widgets.Output()

# Register the function to be called when the run button is clicked
run_button.on_click(process_uploaded_docx)

# Display widgets
display(file_upload, run_button, output)
