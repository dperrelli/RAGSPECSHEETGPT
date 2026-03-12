import os
import re
import pandas as pd
from openai import OpenAI
from openai.types.beta.threads.message_create_params import (
    Attachment,
    AttachmentToolFileSearch,
)


def extract_dimensions(input_str):
    """
    Extract height, width, and depth values from the assistant response using regex.
    """

    pattern_height = r"Height:\s*([^,【]+)"
    pattern_width = r"Width:\s*([^,【]+)"
    pattern_depth = r"Depth:\s*([^,【]+)"

    match_height = re.search(pattern_height, input_str)
    match_width = re.search(pattern_width, input_str)
    match_depth = re.search(pattern_depth, input_str)

    height = match_height.group(1).strip() if match_height else "N/A"
    width = match_width.group(1).strip() if match_width else "N/A"
    depth = match_depth.group(1).strip() if match_depth else "N/A"

    return height, width, depth


def get_assistant(client):
    """
    Retrieve an existing assistant or create one if it does not exist.
    """

    assistants = client.beta.assistants.list()

    for assistant in assistants:
        if assistant.name == "Specsheet Assistant":
            return assistant

    return client.beta.assistants.create(
        model="gpt-4o",
        description="Assistant for extracting appliance dimensions from spec sheets",
        instructions="Extract the height, width, and depth from the provided spec sheets.",
        tools=[{"type": "file_search"}],
        name="Specsheet Assistant",
    )


def process_pdf(client, assistant, pdf_path):
    """
    Upload a PDF, run the assistant, and return extracted dimensions.
    """

    catalog_number = os.path.splitext(os.path.basename(pdf_path))[0]

    with open(pdf_path, "rb") as pdf_file:
        file = client.files.create(
            file=pdf_file,
            purpose="assistants",
        )

    thread = client.beta.threads.create()

    prompt = (
        "Please extract the overall height, width, and depth of the appliance "
        "in this PDF spec sheet. If you cannot find any of these dimensions, "
        "respond with 'N/A'. Format the output as: Height: X, Width: Y, Depth: Z."
    )

    client.beta.threads.messages.create(
        thread_id=thread.id,
        role="user",
        content=prompt,
        attachments=[
            Attachment(
                file_id=file.id,
                tools=[AttachmentToolFileSearch(type="file_search")],
            )
        ],
    )

    run = client.beta.threads.runs.create_and_poll(
        thread_id=thread.id,
        assistant_id=assistant.id,
        timeout=600,
    )

    if run.status != "completed":
        print(f"Run failed for {catalog_number}")
        return catalog_number, "N/A", "N/A", "N/A"

    messages_cursor = client.beta.threads.messages.list(thread_id=thread.id)
    messages = [message for message in messages_cursor]

    message = messages[0]
    res_text = message.content[0].text.value

    print(f"Response from assistant for {catalog_number}: {res_text}")

    try:
        height, width, depth = extract_dimensions(res_text)
    except Exception as e:
        print(f"Error extracting dimensions from {catalog_number}: {e}")
        height, width, depth = "N/A", "N/A", "N/A"

    client.files.delete(file.id)

    return catalog_number, height, width, depth


def main():
    client = OpenAI()

    folder_path = "SpecSheetLib"

    df = pd.DataFrame(columns=["Catalog Number", "Height", "Width", "Depth"])

    assistant = get_assistant(client)

    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):

            pdf_path = os.path.join(folder_path, filename)

            catalog_number, height, width, depth = process_pdf(
                client, assistant, pdf_path
            )

            print(
                f"Extracted dimensions for {catalog_number}: "
                f"Height: {height}, Width: {width}, Depth: {depth}"
            )

            df = pd.concat(
                [
                    df,
                    pd.DataFrame(
                        {
                            "Catalog Number": [catalog_number],
                            "Height": [height],
                            "Width": [width],
                            "Depth": [depth],
                        }
                    ),
                ],
                ignore_index=True,
            )

    output_file = "extracted_dimensions.xlsx"

    df.to_excel(output_file, index=False)

    print(f"Process completed. Data saved to {output_file}")


if __name__ == "__main__":
    main()