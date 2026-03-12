import os
import pandas as pd
import requests


def download_pdf(catalog_number, url, save_folder):
    """
    Download a PDF from a given URL and save it using the catalog number.
    """
    response = requests.get(url)

    if response.status_code == 200:
        file_path = os.path.join(save_folder, f"{catalog_number}.pdf")

        with open(file_path, "wb") as f:
            f.write(response.content)

        print(f"Downloaded: {catalog_number}")

    else:
        print(f"Failed to download: {catalog_number}")


def download_pdfs_from_excel(excel_file, save_folder):
    """
    Read an Excel file containing model numbers and spec sheet URLs,
    then download all available PDFs.
    """
    df = pd.read_excel(excel_file, engine="openpyxl")

    # Create save directory if it does not exist
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    for index, row in df.iterrows():
        catalog_number = row["ModelNumber"]
        pdf_link = row["Quick Specs"]

        if pd.notna(pdf_link):
            download_pdf(catalog_number, pdf_link, save_folder)
        else:
            print(f"No PDF available for catalog number: {catalog_number}")


def main():
    excel_file = "Fromm PDB Pricing 7.2024 w dimensions.xlsx"
    save_folder = "./SpecSheetLib"

    download_pdfs_from_excel(excel_file, save_folder)


if __name__ == "__main__":
    main()