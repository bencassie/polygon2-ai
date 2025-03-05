# Polygon.io Excel Add-in

This Excel add-in integrates with Polygon.io API to fetch financial data directly into your spreadsheets.

## Features

* `=POLYGON.getTickerDetails("AAPL")` - Get ticker details (name, market)
* `=POLYGON.getTickerNews("AAPL", 5)` - Get latest news articles

## Prerequisites

* Node.js (v14+) and npm
* Microsoft Excel (Desktop or Online)
* [Polygon.io API key](https://polygon.io/)

## Setup & Installation

1. **Clone & Install**
   ```bash
   git clone https://github.com/yourusername/polygon-ai.git
   cd polygon-ai
   npm install
   ```

2. **Get Development Certificates**
   ```bash
   npx office-addin-dev-certs install
   ```

3. **Build the Project**
   ```bash
   npm run build
   ```

4. **Start Development Server**
   ```bash
   npm run dev-server
   ```

## Sideload the Add-in

### For Excel Desktop
1. Open Excel and a blank workbook
2. Go to **Insert** tab > **Add-ins** > **My Add-ins** > **Manage My Add-ins** > **Upload My Add-in**
3. Browse to the project folder and select `manifest.xml`

### For Excel Online
1. Open Excel Online and create a new workbook
2. Go to **Insert** tab > **Add-ins** > **Manage My Add-ins** > **Upload My Add-in**
3. Browse to and select the `manifest.xml` file

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

## Troubleshooting

* If the add-in doesn't load, check browser console for errors
* Ensure your Polygon.io API key has been set
* Verify the dev server is running on https://localhost:3000

## Resources

* [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/)
* [Polygon.io API Documentation](https://polygon.io/docs/)