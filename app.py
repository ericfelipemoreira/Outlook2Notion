import streamlit as st
import requests, os

from notion.client import NotionClient
from datetime import datetime
from dotenv import load_dotenv

def init_env():
    load_dotenv(".env")

    st.set_page_config(page_title=" Outlook2Notion", page_icon=">>")
    st.title("Outlook2Notion")

    os.environ["OUTLOOK_API_URL"] = os.getenv("OUTLOOK_API_URL")
    os.environ["NOTION_TOKEN_V2"] = os.getenv("NOTION_TOKEN_V2")
    os.environ["NOTION_DATABASE_ID"] = os.getenv("NOTION_DATABASE_ID")
    os.environ["OUTLOOK_ACCESS_TOKEN"] = os.getenv("OUTLOOK_ACCESS_TOKEN")
    os.environ["ONEDRIVE_FOLDER_URL"] = os.getenv("ONEDRIVE_FOLDER_URL")
    

# Function to authenticate and get Notion client
def authenticate_notion(token_v2):
    return NotionClient(token_v2=token_v2)

# Function to create a new Notion page
def create_notion_page(client, database_id, properties):
    database = client.get_collection_view(database_id)
    new_page = database.collection.add_row()
    for key, value in properties.items():
        setattr(new_page, key, value)
    return new_page

# Function to save attachment to OneDrive
def save_to_onedrive(attachment_url, onedrive_folder_url):
    # Implement your OneDrive saving logic here
    # Use appropriate libraries or APIs for OneDrive interactions
    pass


def main():
    # Outlook API URL
    outlook_api_url = os.environ["OUTLOOK_API_URL"]

    # Notion API Token_v2 (replace with your actual token)
    notion_token_v2 = os.environ["NOTION_TOKEN_V2"]

    # Notion Database ID (replace with your actual database ID)
    notion_database_id = os.environ["NOTION_DATABASE_ID"]

    # Fetching Outlook emails
    outlook_response = requests.get(outlook_api_url, headers={"Authorization": "Bearer " + os.environ["OUTLOOK_ACCESS_TOKEN"]})

    if outlook_response.status_code == 200:
        emails = outlook_response.json().get("value", [])
        
        # Authenticating Notion
        notion_client = authenticate_notion(notion_token_v2)
        
        for email in emails:
            # Extracting relevant information from the email
            email_subject = email.get("subject", "")
            email_date = datetime.strptime(email.get("receivedDateTime", ""), "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%d %H:%M:%S")
            # attachment_url = email.get("attachments", [])[0].get("contentBytes", "")  # Assuming only one attachment for simplicity
            attachment_url = ""
            
            # Saving attachment to OneDrive
            # onedrive_folder_url = os.environ["ONEDRIVE_FOLDER_URL"]
            # save_to_onedrive(attachment_url, onedrive_folder_url)
            
            # Creating a new Notion page with email information
            notion_properties = {"Name": email_subject, "Date": email_date, "Attachment_Link": attachment_url}
            create_notion_page(notion_client, notion_database_id, notion_properties)

        print("Emails processed successfully.")
    else:
        print(f"Failed to fetch emails from Outlook. Status code: {outlook_response.status_code}")


if __name__ == "__main__":
    init_env()
    main()
