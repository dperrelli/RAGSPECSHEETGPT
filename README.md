This project is a Python automation pipeline that downloads manufacturer specification sheets and extracts structured appliance dimensions (height, width, depth) using AI-powered document analysis.

This project was built to automate the extraction of product dimensions from thousands of manufacturer specification PDFs, eliminating the need for manual data entry.

==OVERVIEW==

Many product catalogs store technical specifications in unstructured PDF spec sheets, making automated data ingestion difficult. To solve this, we undergo the following process:
  - Downloading specification sheets from manufacturer URLs
  - Uploading each PDF to an AI assistant for analysis
  - Extracting structured dimension data (height, width, depth)
  - Normalizing results using regex parsing
  - Exporting the structured dataset to Excel

==ASSUMPTIONS==

The current implementation makes several assumptions:
  - The input Excel file contains the columns ModelNumber and Quick Specs, where Quick Specs contains a valid URL to the specification sheet PDF.
  - Each specification sheet contains clearly labeled dimension information using terms such as Height, Width, and Depth.
  - Dimensions appear in a format that can be interpreted by the AI assistant and parsed using regex patterns.
  - Each PDF corresponds to a single product model.
  - The OpenAI API is available and the user has a valid API key configured as an environment variable.
  - The PDF files are accessible and downloadable via HTTP requests.

==FEATURES==

  - Automated download of specification sheet PDFs
  - AI-powered document analysis using the OpenAI API
  - Extraction of structured dimensions from unstructured documents
  - Regex-based parsing for consistent output formatting
  - Batch processing of large PDF libraries
  - Excel dataset generation for downstream product catalog systems

==EXECUTION==

1. Run the download_spec_shets.py script to download all available PDFs into a local library.
2. Run the extract_dimentions.py script to upload the PDFs to an AI assistant which analyzes the specification sheet and returns values in excel format.



