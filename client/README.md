# Reports of Collections and Deposits (RCD) System - LGU Concepcion

This is a modern web application for the LGU Concepcion, Romblon.

## Tech Stack
- React (Vite)
- TypeScript
- Tailwind CSS (v4)
- React Router DOM (Navigation)
- Lucide React (Icons)
- Google Apps Script (Serverless Backend)

## Setup

1.  Clone the repository.
2.  Install dependencies:
    ```bash
    cd client
    npm install
    ```
3.  Start the development server:
    ```bash
    npm run dev
    ```

## Backend Deployment (Google Apps Script)

This project uses Google Apps Script as a serverless backend to communicate with Google Sheets.

1.  **Prepare the Google Sheet**:
    - Create a new Google Sheet.
    - Go to **Extensions > Apps Script**.
    - Copy the content of `scripts/GAS_CODE.js` from this repository and paste it into the script editor (`Code.gs`).
    - Save the project.

2.  **Deploy as Web App**:
    - Click **Deploy** > **New deployment**.
    - Select **Web app**.
    - Description: `RCD Backend v1`
    - Execute as: **Me** (your email).
    - Who has access: **Anyone** (this is required for the client to access it without OAuth complexity).
    - Click **Deploy**.
    - **Copy the Web App URL** (starts with `https://script.google.com/macros/s/...`).

3.  **Configure Environment**:
    - Create a `.env` file in the `client` folder (copy `.env.example` if it exists, or just create it).
    - Add your Web App URL:
    ```env
    VITE_GOOGLE_SCRIPT_URL=https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec
    ```

## Frontend Deployment (GitHub Pages)

To deploy the frontend to GitHub Pages:

1.  Ensure your changes are committed.
2.  Run the deploy script:
    ```bash
    npm run deploy
    ```
    This will build the project and push the `dist` folder to the `gh-pages` branch.

## Default Login (Mock)
If the backend is not connected or fails, the app may fall back to local storage or mock data.
- Email: `admin@lgu.gov.ph`
- Password: `admin`
