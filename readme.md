# Polygon.io Excel Add-in

This Excel add-in integrates with Polygon.io API to fetch financial data directly into your spreadsheets.

[Click to Watch](https://streamable.com/z5hydz)

## Features

* `=POLYGON.getTickerDetails("AAPL")` - Get ticker details (name, market)
* `=POLYGON.getTickerNews("AAPL", 5)` - Get latest news articles

## Prerequisites

* Node.js (v14+) and npm
* Microsoft Excel (Desktop or Online)
* [Polygon.io API key](https://polygon.io/)

## Setup & Installation (Windows with PowerShell)

1. **Clone & Install**
   ```powershell
   git clone https://github.com/yourusername/polygon-ai.git
   cd polygon-ai
   npm install
   ```

2. **Get Development Certificates**
   ```powershell
   npx office-addin-dev-certs install
   ```
   > Note: If you receive a security prompt, select 'Yes' to install the certificate

3. **Build the Project**
   ```powershell
   npm run build
   ```

4. **Start Development Server**
   ```powershell
   npm run dev-server
   ```
   > The server will start at https://localhost:3000

## Sideload the Add-in

### For Excel Desktop (Upload Method)
1. Open Excel and a blank workbook
2. Go to **Insert** tab > **Add-ins** > **My Add-ins** > **Manage My Add-ins** > **Upload My Add-in**
3. Browse to the project folder and select `manifest.xml`

### For Excel Desktop (Network Share Method - Windows)
1. Create a network share using PowerShell (or use an existing share):
   ```powershell
   # Create a folder for sharing
   New-Item -Path "C:\PolygonShare" -ItemType Directory -Force
   
   # Share the folder (requires admin privileges)
   New-SmbShare -Name "PolygonShare" -Path "C:\PolygonShare" -FullAccess Everyone
   ```

2. Create the add-in folder and copy the manifest:
   ```powershell
   # Create add-in folder
   New-Item -Path "C:\PolygonShare\Polygon" -ItemType Directory -Force
   
   # Copy manifest to the shared folder
   Copy-Item -Path ".\manifest.xml" -Destination "C:\PolygonShare\Polygon\"
   ```

3. In Excel, go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**
4. Add the network path to your shared folder (e.g., `\\localhost\PolygonShare\Polygon`)
5. Check the "Show in Menu" option and click **OK**
6. Restart Excel
7. Go to **Insert** tab > **Add-ins** > **My Add-ins**
8. Select the "Shared Folder" tab, and you should see your add-in

### For Excel Online
1. Open Excel Online and create a new workbook
2. Go to **Insert** tab > **Office Add-ins**
3. Select **Upload My Add-in** at the bottom of the dialog
4. Browse to and select the `manifest.xml` file
5. Click **Upload**

## Using the Add-in

1. After sideloading, the add-in appears in the Home tab
2. Click the add-in button to open the task pane
3. Set your Polygon.io API key in the task pane
4. Use the custom functions in your spreadsheet:
   * `=POLYGON.getTickerDetails("AAPL")`
   * `=POLYGON.getTickerNews("AAPL", 5)`

## Development Commands

* `npm run build` - Build for production
* `npm run build:dev` - Build for development
* `npm run dev-server` - Start dev server
* `npm run start` - Build and start debugging
* `npm run validate` - Validate the manifest file
* `npm run lint` - Check for linting issues

## Project Structure

* `/src/functions/` - Custom Excel functions implementation
* `/src/taskpane/` - Task pane UI code
* `manifest.xml` - Add-in configuration file

## Deployment

1. Update `webpack.config.js` to replace development URL with your production URL
2. Build the project: `npm run build`
3. Host the files from the `/dist` folder on your web server
4. Update the manifest.xml to point to your production URL
5. Distribute the manifest.xml file to users

## Troubleshooting (Windows)

* If the add-in doesn't load, check browser console for errors (press F12 in Excel Online)
* Ensure your Polygon.io API key has been set
* Verify the dev server is running on https://localhost:3000
* Check Windows Firewall settings if connecting from another device
* PowerShell commands may require elevated permissions:
  ```powershell
  # Run PowerShell as Administrator
  Start-Process powershell -Verb RunAs
  ```
* If you see certificate errors:
  ```powershell
  # Reinstall development certificates
  npx office-addin-dev-certs install --machine
  ```
* Verify the manifest is correctly formatted using the validator:
  ```powershell
  npm run validate
  ```

## Resources

* [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/)
* [Polygon.io API Documentation](https://polygon.io/docs/)