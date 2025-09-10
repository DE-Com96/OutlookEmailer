# Outlook Emailer

## Description

This Python script  automates the creation of draft emails in Outlook. It reads recipient information and email content from a configuration file, populates an email template, and adds specified attachments, saving each email as a draft in Outlook.

---

## Features

* **Bulk Email Drafting**: Creates multiple email drafts from a list of recipients in an Excel file.
* **Dynamic Email Body**: Uses a text file as a template for the email body, with placeholders for personalized information.
* **Attachment Support**: Attaches one or more files to each email.
* **Easy Configuration**: All email parameters (recipients, CC, subject, attachments) are configured in a single Excel file.

---

## Requirements

* Python 3
* pandas library (`pip install pandas`)
* pywin32 library (`pip install pywin32`)
* Microsoft Outlook installed and configured on your machine.

---

## How to Use

1.  **Clone the repository**:
    ```bash
    git clone <repository-url>
    cd <repository-name>
    ```
2.  **Install the required packages**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Configure your emails**:
    * Edit the **`Config.xlsx`** file within the `config` folder to include the recipient details. Each row represents a separate email. 
        * You may modify/add/remove all white-text columns depending on how many variable fields you require for the email.
        * Do not remove or rename the red-text columns, this information is always required (TO, CC, Subject).
        * The attachments columns can be left blank if not required. To add attachments, enter the file name of the attachment seperated by ` ; ` (semicolon with spaces on either side)
    * Modify the **`EmailBody.txt`** file within the `config` folder to create your desired email template. Use placeholders denoted by square brackets `[ColumnName]` that correspond to the column names in `Config.xlsx`.
    * Place all necessary attachments in the **`attachments`** folder. The names of the files should correspond to the names in the "ATTACHMENTS" column of the `Config.xlsx` file.
4.  **Run the script**:
    ```bash
    python main.py
    ```
5.  **Check your Outlook**:
    The script will create the emails and save them in your Outlook drafts folder.

---

## Configuration Details

### `Config.xlsx`

This file is the main configuration for the emails you want to send. Each row corresponds to one email draft. The columns are:

* **Title, FirstName, LastName, Email, RefNum**: These are placeholder fields that will be used to populate the email body. You can add or remove columns as needed, but make sure to update `EmailBody.txt` accordingly.
* **TO**: The primary recipient's email address.
* **CC**: The CC recipients' email addresses.
* **SUBJECT**: The subject line of the email.
* **ATTACHMENTS**: The names of the files to be attached, separated by ` ; `. These files must be present in the `attachments` folder.

### `EmailBody.txt`

This file serves as the template for the email body. You can customize the text and use placeholders that correspond to the column names in `Config.xlsx`. For example, `[FirstName]` in the text file will be replaced by the value from the "FirstName" column in the Excel file for each email.

### `attachments` folder

This folder should contain all the files that you wish to send as attachments. The filenames in this folder must match the filenames listed in the `ATTACHMENTS` column of the `Config.xlsx` file.