# Automated Church Presentation Slides  

A Google Apps Script that **automates the creation of weekly church service slides**, saving hours of manual work by dynamically pulling data from Google Sheets and web sources.  

## Motivation  
Church staff manually compiled slides every week—a repetitive and error-prone process. This tool was built to:  
- **Eliminate manual copy-pasting** of scriptures, hymns, and duty assignments.  
- **Ensure accuracy** by programmatically fetching the latest data.  
- **Save time** for volunteers.  

## Features  
- **Google Sheets Integration**: Pulls duty assignments, scripture references, and hymn numbers from a spreadsheet.  
- **Web Scraping**: Fetches verse text and hymn lyrics from online sources.  
- **Dynamic Slide Generation**: Inserts data into preformatted Google Slides.  

## How It Works  
1. Script reads a Google Sheet with service details (e.g., Sabbath duties, scripture references).  
2. Fetches verse text from an online Bible API/scraper and hymn lyrics from a hymnal website.  
3. Populates a Google Slides template with the retrieved data.  

## Technologies Used  
- **Google Apps Script** (Slides + Sheets API)  
- **Web Scraping** 
- **Automation**  

## Setup  
1. Clone this repository or copy the script into your Google Slides project.  
2. Link your Google Sheet (with duty assignments) and Slides template.  
3. Configure API/website endpoints for scripture/hymn data.  

