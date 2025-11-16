# Universal Document Search

A desktop application built with Python and Tkinter that performs a powerful semantic search across your local files and Google Drive. It ranks documents by relevance (TF-IDF and Cosine Similarity) and displays matching sentences with their corresponding line and page numbers.

## Features

-   **Multi-Source Search**: Simultaneously search through local folders and Google Drive.
-   **Recursive Scanning**: Thoroughly scans all nested sub-folders in both local paths and Google Drive.
-   **Broad File Support**: Ingests and processes `.pdf`, `.docx`, `.xlsx`, and `.txt` files.
-   **Google Workspace Integration**: Correctly handles and exports native Google Docs and Google Sheets.
-   **Intelligent Ranking**: Uses TF-IDF and Cosine Similarity to rank documents by relevance to your query.
-   **Contextual Results**: Displays the exact sentences that match your query, complete with line numbers, page numbers (for PDFs), and highlighted terms.
-   **User-Friendly GUI**: A simple and intuitive graphical interface built with Tkinter.

## Setup Instructions

### 1. Prerequisites

-   Python 3.7+

### 2. Clone the Repository

```
git clone <your-repository-url>
cd <repository-name>
```

### 3. Install Dependencies

Install all the required Python libraries using the `requirements.txt` file.

```
pip install -r requirements.txt
```

### 4. Set Up Google Drive API Credentials

To enable search in Google Drive, you need to authorize the application.

**A. Enable the Google Drive API**

1.  Go to the [Google Cloud Console](https://console.cloud.google.com/apis/library).
2.  Select an existing project or create a new one.
3.  In the API Library, search for "**Google Drive API**" and click **Enable**.

**B. Create OAuth 2.0 Credentials**

1.  Navigate to the [Credentials page](https://console.cloud.google.com/apis/credentials) in the Google Cloud Console.
2.  Click **+ CREATE CREDENTIALS** and select **OAuth client ID**.
3.  Set the **Application type** to **Desktop app**.
4.  Give it a name (e.g., "Document Search App") and click **Create**.
5.  A pop-up will appear. Click **DOWNLOAD JSON**.
6.  Rename the downloaded file to `credentials.json` and place it in the root directory of this project.

> **CRITICAL SECURITY NOTE:** The `credentials.json` file is your application's private key. The `.gitignore` file is already configured to prevent this file from being uploaded to GitHub. **Never share it or commit it to version control.**

## How to Run the Application

Execute the main script from your terminal:

```
python search_app.py
```

## How to Use the GUI

1.  **Choose Search Source**:
    -   **Local Storage**: To search files on your computer.
    -   **Google Drive**: To search files in a specific Google Drive folder.
    -   **Both**: To search across both local and Google Drive sources.

2.  **Provide Paths**:
    -   If using **Local Storage**, click **"Add Folder..."** to select one or more folders to scan. All sub-folders will be included automatically.
    -   If using **Google Drive**, paste the **Folder ID** into the corresponding field. The ID is the last part of the folder's URL (e.g., for `.../folders/1aBcDeFg...`, the ID is `1aBcDeFg...`).

3.  **First-Time Google Drive Authentication**:
    -   The first time you run a search involving Google Drive, a browser window will open.
    -   Log in to your Google account and grant the application permission to view your Drive files.
    -   A `token.json` file will be created in the project folder to store your authorization. The `.gitignore` file also protects this file from being uploaded.

4.  **Enter Your Query & Search**:
    -   Type your search terms into the query box and click **Search**.
    -   The status and results will appear in the text area below.
```