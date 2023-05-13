import os
import re
import requests
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from dotenv import load_dotenv


def print_shareholder_info():
    """
    Prints shareholder information from the given API keys and company no. from environment variables
    """
    load_dotenv()

    # Insert API key
    api_key = os.environ.get("COMPANY_HOUSE_API_KEY")

    # Insert company number to retrieve its confirmation statement
    company_number = os.environ.get("COMPANY_NO")

    # Make a filing history request to get all confirmation statements
    url = f"https://api.company-information.service.gov.uk/company/{company_number}/filing-history?category=confirmation-statement&items_per_page=100"

    response = requests.get(url, auth=(api_key, ""))

    # Parse the response for the link to the document metadata
    response_json = response.json()
    items = response_json["items"]

    # Check each confirmation statement to only select the latest confirmation statement with updates
    for item in items:
        description = item["description"]
        if "confirmation-statement-with-updates" in description.lower():
            metadata_link = item["links"]["document_metadata"]
            metadata_response = requests.get(
                metadata_link,
                auth=(api_key, ""),
                headers={"Accept": "application/json"},
            )
            metadata_response_json = metadata_response.json()
            document_link = metadata_response_json["links"]["document"]

            # Get the content of the document
            document_response = requests.get(document_link, auth=(api_key, ""))
            document_content = document_response.content

            # Write the content to a file named "confirmation_statement.pdf"
            with open("confirmation_statement.pdf", "wb") as f:
                f.write(document_content)
            break

    pages = convert_from_path(
        "confirmation_statement.pdf",
        250,
        poppler_path=r"C:\Program Files\poppler-23.01.0\Library\bin",
    )

    images = []
    for page in pages:
        images.append(page)

    # Combine all the images into one
    combined_image = Image.new(
        "RGB", (images[0].width, sum([i.height for i in images]))
    )
    y_offset = 0
    for image in images:
        combined_image.paste(image, (0, y_offset))
        y_offset += image.height

    # Save the combined image (note for later: see if i can save the image as a temp file)
    combined_image.save("confirmation_statement.png", "PNG")

    # Set the path to Tesseract
    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )

    # Load the image
    combined_image = Image.open("confirmation_statement.png")

    # Use pytesseract to extract text from the image and format it into one paragraph, while limiting instances of three spaces to two
    text = (
        pytesseract.image_to_string(combined_image)
        .replace("\n\n", "\n")
        .replace("\n", "  ")
        .replace("confirmation  statement", "confirmation statement")
        .replace("this  confirmation", "this confirmation")
        .replace("of  this confirmation", "of this confirmation")
        .replace("date  of this", "date of this")
        .replace(r"\b([A-Z][a-z]+)  ([A-Z][a-z]+)\b", r"\1 \2")
    )

    print(text)

    # Find all instances of "(?<!0 )(\d+) ORDINARY shares held as at the date of this confirmation statement" and the name that follows
    shares = re.findall(
        r"(?<!0 )(\d+) ORDINARY shares held as at the date of this confirmation statement  Name: (\S+(?: \S+)*)",
        text,
    )

    # Calculate the total number of shares
    total_shares = sum([int(share[0]) for share in shares])

    # Create a list of the shareholders and their percentage of shares
    shareholders = []
    for share in shares:
        percentage = round(int(share[0]) / total_shares * 100, 2)
        shareholders.append(f"- {share[1]} - {percentage}%")

    # Combine the shareholders into a string and print it
    shareholders_string = "\n".join(shareholders)
    print(f"The company has the following shareholders:\n{shareholders_string}")
