# CI/CD Deployment Guide for Google Apps Script

This project is configured for automated deployment using **Clasp** and **GitHub Actions**.

## 🛠 Prerequisites

1.  **Node.js**: Install it from [nodejs.org](https://nodejs.org/).
2.  **Clasp**: Install globally using `npm install -g @google/clasp`.
3.  **Enable Apps Script API**: Go to [script.google.com/home/settings](https://script.google.com/home/settings) and turn on the Google Apps Script API.

## 🚀 Setup Steps

### 1. Link to your Script
Open `.clasp.json` and replace `YOUR_SCRIPT_ID_HERE` with your actual Script ID.
*You can find the Script ID in the Google Apps Script editor under **Project Settings** (Gear icon).*

### 2. Login Locally
Run the following command in your terminal:
```bash
clasp login
```
This will open a browser to authorize Clasp.

### 3. Setup GitHub Secrets
To enable auto-deploy from GitHub, you need to provide your authentication credentials to GitHub:

1.  After running `clasp login`, find the file `~/.clasprc.json` on your computer.
2.  Go to your GitHub Repository > **Settings** > **Secrets and variables** > **Actions**.
3.  Create a **New repository secret**:
    - **Name**: `CLASPRC_JSON`
    - **Value**: Paste the entire content of your `~/.clasprc.json` file.

### 4. Push to Deploy
Now, whenever you push code to the `main` branch of your GitHub repository, the action will automatically:
- Check out the code.
- Install Clasp.
- Use your credentials to push code to Google Apps Script.

## 📝 Commands
- `clasp push`: Manually push local changes to Google Apps Script.
- `clasp pull`: Pull changes from Google Apps Script editor to your local machine.
