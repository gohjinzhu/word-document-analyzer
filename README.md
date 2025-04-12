# Word Document Analyzer

A Microsoft Word Add-in that analyzes document formatting and runs formatting test cases.

## Overview

Word Document Analyzer is a task pane add-in for Microsoft Word that helps you analyze document formatting and verify that text meets specific formatting requirements. The add-in can detect and report on text formatting properties such as bold, underline, font size, and more.

## Features

- **Document Analysis**: Analyze the formatting of text within your Word document
- **Formatting Tests**: Run test cases to verify that document text meets specific formatting requirements
- **User-Friendly Interface**: Clean, responsive UI with support for both light and dark modes

## Prerequisites

- Microsoft Word (desktop or online)
- Node.js and npm for local development

## Installation

### Local Development Setup

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/word-document-analyzer.git
   cd word-document-analyzer
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Generate SSL certificates for local development (required for Office Add-ins):
   ```
   openssl req -newkey rsa:2048 -nodes -keyout key.pem -x509 -days 365 -out cert.pem
   ```

4. Start the local server:
   ```
   npm start
   ```

5. Sideload the add-in in Word:
   - In Word, go to Insert > Add-ins > My Add-ins
   - Choose "Upload My Add-in" and select the manifest.xml file from this project
   - The add-in should now appear in the task pane

### Production Deployment

For production deployment, host the add-in files on a web server and update the URLs in the manifest.xml file accordingly.

## Usage

1. Open a Word document
2. Launch the Word Document Analyzer add-in from the ribbon
3. Use the "Analyze Document" button to check formatting in your document
4. Use the "Run Tests" button to verify that text meets specific formatting requirements

## How It Works

The add-in uses the Word JavaScript API to:

1. Examine the content of the document
2. Check formatting properties (bold, underline, font size, etc.)
3. Report results through the task pane UI

## Development

### Project Structure

- `app.js` - Main application logic
- `index.html` - Add-in user interface
- `manifest.xml` - Add-in manifest file
- `assets/` - Icons and other static assets

### Key Functions

- `analyzeDocument()` - Analyzes the formatting of text in the document
- `runTests()` - Runs predefined test cases to verify formatting requirements

## License

[MIT License](LICENSE)