# Polygon.io Excel Add-in

Welcome to the Polygon.io Excel Add-in project! This add-in allows you to create custom Excel functions using JavaScript to fetch real-time financial data from Polygon.io directly into your spreadsheets.

## Project Overview

This project provides an Excel add-in that integrates with the Polygon.io API to retrieve stock ticker details and news articles. It uses Office.js and modern JavaScript development tools to extend Excel's capabilities.

## Features

* Fetch ticker details (e.g., company name, market) using `=POLYGON.getTickerDetails()`
* Retrieve news articles for a specific ticker with `=POLYGON.getTickerNews()`
* User-friendly task pane to set your Polygon.io API key and insert functions
* Built with modern JavaScript and TypeScript for reliability and maintainability

## Prerequisites

* Node.js (version 14 or higher recommended)
* npm (included with Node.js)
* Microsoft Excel (Desktop or Online version supporting Office Add-ins)
* A Polygon.io API key (sign up at https://polygon.io/)

## Installation

1. Clone the repository: `git clone https://github.com/OfficeDev/Excel-Custom-Functions-JS.git`
2. Navigate to the project directory: `cd polygon-ai`
3. Install dependencies: `npm install`
4. Build the project: `npm run build`

## Development Setup

1. Start the development server: `npm run dev-server`
2. Sideload the add-in into Excel:
   * Open Excel and create a new workbook
   * Follow the instructions at https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing
   * Use the manifest.xml file in the project root
3. Open the task pane via the Home tab and set your Polygon.io API key

## Usage

* Use `=POLYGON.getTickerDetails("AAPL")` to get ticker details for Apple Inc.
* Use `=POLYGON.getTickerNews("AAPL", 5)` to fetch the latest 5 news articles for AAPL
* Access the task pane from the Home tab to manage your API key or insert sample formulas

## Building for Production

1. Build the production version: `npm run build`
2. Update the manifest.xml to point to your production URL (replace https://localhost:3000/ with your deployed URL)
3. Deploy the built files (dist folder) to your web server

## Scripts

* `npm run build`: Build the project for production
* `npm run build:dev`: Build for development
* `npm run dev-server`: Start the development server
* `npm run start`: Start debugging in Excel Desktop
* `npm run stop`: Stop debugging
* `npm run lint`: Check for linting issues
* `npm run lint:fix`: Fix auto-fixable linting issues

## Dependencies

* core-js: Polyfills for older browsers
* regenerator-runtime: Runtime for async/await support
* See package.json for a full list of dependencies and devDependencies

## Contributing

This project follows the Microsoft Open Source Code of Conduct. For details, see CODE_OF_CONDUCT.md. To contribute:

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## Support

For issues or questions, please file a GitHub Issue at https://github.com/OfficeDev/Excel-Custom-Functions-JS/issues. Support is limited to the GitHub Issues platform.

## License

This project is licensed under the MIT License. See the repository for full license details.

## Additional Resources

* Office Add-ins Documentation: https://learn.microsoft.com/office/dev/add-ins/
* Polygon.io API Documentation: https://polygon.io/docs/
* Project Repository: https://github.com/OfficeDev/Excel-Custom-Functions-JS