# Pakistan Equity Investment Daily News Agent

An autonomous agent that compiles daily morning Pakistan equity research into a formatted `.docx` Word file and automatically uploads it to your Google Drive folder (`Daily News`) every day at **7:30 AM PKT**, even when your computer is off.

---

## 📋 Features & Specifications
- **Runs Headlessly in Cloud:** Runs on GitHub Actions on a cron schedule (`02:30 UTC = 07:30 AM PKT`).
- **Targeted Sources:** Scrapes and curates news from:
  - *Business Recorder* (`brecorder.com`)
  - *Dawn Business* (`dawn.com/business`)
  - *The News International - Business* (`thenews.com.pk/business`)
  - *Pakistan Today - Profit* (`profit.pakistantoday.com.pk`)
  - *Tribune Business* (`tribune.com.pk/business`)
- **Strict Filtering:** Last 24 hours only. Excludes non-material international news.
- **Assigned Sectors:** Economy, Pharmaceutical, Automobile, Sugar, Technology, Food, and Modaraba.
- **Report Organization & Typography:**
  - Times New Roman, size 12.
  - Heading: *The Daily News Coverage* + Date.
  - Page-separated sections: **Main News**, **Economic News**, **Sector-Specific News** (omitted if empty), **Other Notables**, and a final page with **Sources & References**.
- **Automated Delivery:** Creates and places the `.docx` into Google Drive folder `Daily News`.

---

## 🚀 Setup Guide (Takes 5 minutes)

### Step 1: Create a Google Cloud Service Account
1. Open the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (e.g. `Pakistan-News-Agent`).
3. Enable the **Google Drive API** under **APIs & Services > Library**.
4. Go to **APIs & Services > Credentials** > Click **Create Credentials** > **Service Account**.
5. Once created, click on the Service Account > Go to **Keys** tab > **Add Key** > **Create new key (JSON)**.
6. A JSON key file will download. Note down the service account email (e.g., `agent@project-id.iam.gserviceaccount.com`).
7. Open Google Drive in your browser, create/find a folder named `Daily News`, and **Share** it with the service account email as **Editor**.

### Step 2: Push this Repository to GitHub
1. Create a new repository on [GitHub](https://github.com/) (Public or Private).
2. Push this project folder to your repository:
   ```bash
   git init
   git add .
   git commit -m "feat: daily pakistan news agent"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```

### Step 3: Add GitHub Secrets
In your GitHub repository, go to **Settings > Secrets and variables > Actions > New repository secret**:
1. `GEMINI_API_KEY`: Your Gemini API key from [Google AI Studio](https://aistudio.google.com/).
2. `GDRIVE_SERVICE_ACCOUNT_JSON`: The entire raw contents of the JSON file downloaded in Step 1.

### Step 4: Verify & Run
- Navigate to the **Actions** tab on your GitHub repository.
- Select **Daily Pakistan News Coverage Agent** and click **Run workflow**.
- Check your Google Drive folder `Daily News` for your generated `.docx` file!
