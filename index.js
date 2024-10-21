const axios = require('axios');
const cheerio = require('cheerio');
const xlsx = require('xlsx');
const fs = require('fs');

// The URL to scrape data from
const URL = 'https://www.bankbazaar.com/gold-rate-uttar-pradesh.html';

// Function to format date and time as a string in local timezone
function formatDateTime(date) {
    return date.toLocaleString(); // Converts to a string in local time
}

// Function to scrape data
async function scrapeGoldPrice() {
    try {
        console.log("Scraping gold price...");

        // Fetch the webpage
        const response = await axios.get(URL);
        const html = response.data;

        // Load HTML into Cheerio
        const $ = cheerio.load(html);

        // Extract the gold price
        const goldPrice = $('span.white-space-nowrap').first().text().trim();

        // Get the current date and time in local timezone
        const now = new Date();
        const dateTime = formatDateTime(now); // Use local time

        console.log(`Gold price on ${dateTime}: ${goldPrice}`);

        // Prepare the data to write into Excel
        const data = [[dateTime, goldPrice]];

        // Check if the file already exists
        const fileName = 'gold_prices.xlsx';
        let workbook;
        if (fs.existsSync(fileName)) {
            // If file exists, read the existing data
            workbook = xlsx.readFile(fileName);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            existingData.push([dateTime, goldPrice]); // Append new data
            const newSheet = xlsx.utils.aoa_to_sheet(existingData);
            workbook.Sheets[workbook.SheetNames[0]] = newSheet;
        } else {
            // If file doesn't exist, create a new workbook and sheet
            workbook = xlsx.utils.book_new();
            const worksheet = xlsx.utils.aoa_to_sheet([['Date and Time', 'Gold Price'], ...data]);
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Gold Prices');
        }

        // Write the data to the Excel file
        xlsx.writeFile(workbook, fileName);
        console.log('Data saved to Excel file successfully.');
    } catch (error) {
        console.error('Error scraping the gold price:', error);
    }
}

// Function to start scraping at regular intervals
function startScraping() {
    console.log('Starting gold price scraper...');

    // Run the scraper immediately
    scrapeGoldPrice();

    // Run the scraper every minute (60000 milliseconds)
    const intervalId = setInterval(scrapeGoldPrice, 60000 *5);

    // Stop the scraper after 1 hour (3600000 milliseconds)
    setTimeout(() => {
        clearInterval(intervalId);
        console.log('Scraping terminated after 1 hour.');
    }, 3600000);
}

// Start the scraping
startScraping();
