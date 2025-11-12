# Outlook Add-in - Logistics Claims

This Outlook add-in allows users to create logistics claims directly from Outlook emails. It's adapted from the Zendesk module to work with Outlook's Office.js API.

## Features

- **AI-Powered Email Analysis**: Automatically extracts claim information from email content using Gemini AI
- **Smart Form Pre-filling**: Pre-fills claim forms based on email content
- **Invoice Item Extraction**: Automatically extracts items from uploaded invoices
- **Multiple Carrier Support**: Currently supports UPS (extensible to other carriers)
- **File Upload**: Support for proof of purchase and supporting evidence documents
- **Shipment Items Management**: Add multiple items with quantities, amounts, and currencies

## Prerequisites

- Node.js (v14 or higher)
- npm or yarn
- Outlook (desktop, web, or mobile)
- Office 365 account (for testing)

## Setup Instructions

### 1. Install Dependencies

```bash
npm install
```

### 2. Update Manifest Configuration

Before running the add-in, you need to update the `manifest.xml` file:

1. **Generate a unique GUID** for the add-in ID:
   - You can use an online GUID generator or run: `node -e "console.log(require('crypto').randomUUID())"`
   - Replace `00000000-0000-0000-0000-000000000000` in the `<Id>` tag with your GUID

2. **Update the URLs** in `manifest.xml`:
   - Replace `https://localhost:3000` with your actual server URL (for development, localhost:3000 is fine)
   - Update icon URLs if you have custom icons

3. **Add your logo**:
   - Place your logo file in an `assets` folder
   - Update the icon URLs in the manifest to point to your logo

### 3. Start the Development Server

```bash
npm start
```

This will:
- Install SSL certificates for HTTPS (required for Office add-ins)
- Start a local development server
- Provide instructions for sideloading the add-in

### 4. Sideload the Add-in in Outlook

#### For Outlook on the Web:

1. Go to [Outlook on the Web](https://outlook.office.com)
2. Click the gear icon (Settings) in the top right
3. Click "View all Outlook settings"
4. Go to "Mail" > "Manage add-ins"
5. Click "+ Add a custom add-in" > "Add from file"
6. Upload the `manifest.xml` file

#### For Outlook Desktop (Windows/Mac):

1. Open Outlook
2. Go to "Get Add-ins" from the ribbon
3. Click "My Add-ins" > "Add a Custom Add-in" > "Add from File"
4. Select the `manifest.xml` file

#### For Outlook Mobile:

Add-ins are typically managed through the web interface and will sync to mobile.

### 5. Test the Add-in

1. Open an email in Outlook
2. Look for the "Obsydian AI" button in the ribbon
3. Click "Create Claim" to open the task pane
4. The add-in will analyze the email and pre-fill the form

## Configuration

### API Endpoints

The add-in uses the following API endpoints (configured in `taskpane.js`):

- **Claims API**: `https://api-obsydian.up.railway.app/api/claims/create-claim`
- **Invoice Extraction API**: `https://api-obsydian.up.railway.app/api/invoices/extract-items`
- **Gemini API**: For AI-powered email analysis

### API Authentication

Update the `API_AUTH_TOKEN` and `API_ORGANIZATION_ID` constants in `taskpane.js` with your actual credentials.

### Gemini API Key

Update the `GEMINI_API_KEY` constant in `taskpane.js` with your Gemini API key.

## Project Structure

```
outlook-module/
├── manifest.xml          # Outlook add-in manifest
├── taskpane.html         # Main UI HTML
├── taskpane.js           # Main JavaScript logic
├── taskpane.css          # Styles
├── commands.html         # Commands page (required by manifest)
├── package.json          # Node.js dependencies
└── README.md            # This file
```

## Development

### Making Changes

1. Edit the relevant files (`taskpane.html`, `taskpane.js`, `taskpane.css`)
2. The development server will automatically reload
3. Refresh the Outlook add-in to see changes

### Validating the Manifest

```bash
npm run validate
```

### Stopping the Server

```bash
npm run stop
```

## Differences from Zendesk Module

The Outlook version has been adapted from the Zendesk module with the following key changes:

1. **API Integration**: Uses Office.js instead of ZAF (Zendesk App Framework)
2. **Email Reading**: Reads email body/subject instead of ticket comments
3. **User Context**: Gets user email from Outlook instead of Zendesk
4. **Manifest Format**: Uses XML manifest instead of JSON
5. **Source Identifier**: Sets `source: 'outlook'` in API payloads instead of `source: 'zendesk'`

## Troubleshooting

### Add-in Not Loading

- Ensure the development server is running
- Check that SSL certificates are installed (`npm start` installs them automatically)
- Verify the manifest.xml is valid (`npm run validate`)
- Check browser console for errors

### API Errors

- Verify API endpoints are correct
- Check API authentication tokens
- Ensure CORS is configured on the API server

### Office.js Not Available

- Ensure you're running the add-in in a supported Outlook client
- Check that Office.js is loaded correctly in the HTML

## Support

For issues or questions, please contact the Obsydian AI team or open an issue in the repository.

## License

MIT

