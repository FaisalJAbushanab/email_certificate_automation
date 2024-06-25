# Python Script Automation for Email Sending with Attachment

## Overview

This Python script automates the process of sending emails with attachments using the `google api` library. It's designed to streamline the task of sending personalized emails to recipients with optional attachments.

## Features

- **Email Sending:** Automates the sending of emails through google api.
- **Attachment Support:** Allows attaching files (e.g., documents, images) to emails.
- **Personalization:** Supports customizable email content for each recipient.
- **Authentication:** Allows authenticate to google accounts using google api

## Requirements

- Python 3.10
- Dependencies listed in `requirements.txt` 

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your_username/your_repo.git
   cd your_repo
   ```
2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
# Usage

    In order for the script to function appropriatly, you have to enable google email api in your GCP account and give permission for the service account you use to perform send operation. Then download the key in json format and attach it as `cred.json`
    
    Include the attachments in a subdirectory, default directory is `./certificates`
    
    Include the attendance sheet in the same directory as the script, default is `attendance.xlsx`

    Run the script:

    bash

    python send_email.py
