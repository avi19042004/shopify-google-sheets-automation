# 🚀 Zero-Cost ERP: Shopify to Google Sheets Automation

A lightweight, serverless Order Management System (OMS) built entirely on **Google Apps Script**. This tool acts as a bridge between Shopify, Google Sheets, Telegram, and WhatsApp, allowing small businesses to manage orders without expensive ERP software.

![Google Apps Script](https://img.shields.io/badge/Built%20With-Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-drive&logoColor=white)
![Shopify](https://img.shields.io/badge/Shopify-API-95BF47?style=for-the-badge&logo=shopify&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

## ✨ Features

* **🔄 Real-Time Sync:** Fetches new orders from Shopify Admin API and syncs them to Google Sheets automatically.
* **📱 Instant Alerts:** Sends a Telegram notification immediately when a new order is received.
* **💬 WhatsApp Integration:** Automatically generates "Click-to-Chat" WhatsApp links pre-filled with customer details and tracking info.
* **📊 Built-in Dashboard:** Aggregates sales data and product counts directly in the sheet.
* **🔐 Licensing System:** Includes a token-based security check to restrict script execution to authorized users.

## 📥 Quick Start (The Easy Way)

To ensure the script works correctly, you need the exact column structure. I have created a template you can copy.

### 1. Get the Spreadsheet Template
Click the link below to open the template. Then go to **File > Make a copy** to save it to your own Google Drive.

👉 **[Click Here to Open Google Sheet Template](https://docs.google.com/spreadsheets/d/1miO7IFjL9UW_pBGI6TSMoUljPHZLeh5ueiSNndzHstE/edit?usp=sharing)**

*(Note: The template is View Only. You must make a copy to edit it.)*

### 2. Configure Your Keys
Once you have your own copy, go to the **`Dashbord`** sheet (check the bottom tabs) and enter your API keys in the specific cells:

| Cell | Content | Description |
| :--- | :--- | :--- |
| **A2** | `Shopify API Key` | From Shopify Admin > Apps > App development |
| **A5** | `Shopify Password` | The Admin API Access Token |
| **A8** | `Shopify API URL` | e.g., `https://your-store.myshopify.com/admin/api/2023-04/orders.json` |
| **A11**| `Telegram Bot Token`| From @BotFather |
| **A14**| `Telegram Chat ID` | The ID of the user/group receiving alerts |
| **A17**| `License Token` | A unique string for the security check (e.g., "USER-123") |

## 🛠️ Manual Installation (If starting from scratch)

If you prefer to set up the sheet manually instead of using the template:

1.  Create a new Google Sheet.
2.  Rename the first tab to `Order`.
3.  Create a second tab named `Dashbord`.
4.  Open **Extensions > Apps Script**.
5.  Paste the code from `code.js` in this repository.

## ⚙️ How to Deploy

1.  Open your Google Sheet.
2.  Go to **Extensions** > **Apps Script**.
3.  Replace the default code with the code from this repository.
4.  **Critical Step:** Update the spreadsheet IDs in the code:
    * Look for `YOUR_MAIN_SPREADSHEET_ID` and replace it with your sheet's ID (found in the URL between `/d/` and `/edit`).
5.  Save and Run `onOpen` to verify permissions.

## 🚀 Usage

### Manual Run
1.  Refresh your Google Sheet.
2.  A new menu item **"Custom Menu"** will appear in the toolbar.
3.  Click **Custom Menu > Run** to fetch orders immediately.

### Automatic Run (Triggers)
To have this run automatically (e.g., every hour):
1.  Open the Apps Script editor.
2.  Click on the **Clock Icon (Triggers)** on the left sidebar.
3.  Click **+ Add Trigger**.
4.  Select `checkTokenAndExecute` -> `Head` -> `Time-driven` -> `Every hour`.

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
