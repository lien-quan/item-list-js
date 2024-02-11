const { JSDOM } = require("jsdom");
const ExcelJS = require("exceljs");
const axios = require("axios");

// Function to extract price from text
function extractPriceFromText(text) {
  const pricePattern = /Giá:\s*(\d+)/;
  const match = text.match(pricePattern);
  if (match) {
    text = text.replace(match[0], "").trim(); // Remove the price text from the detail text
    return { price: match[1], textWithoutPrice: text };
  } else {
    return { price: "N/A", textWithoutPrice: text };
  }
}

// Function to format HTML content into a more readable text format
function formatText(htmlString) {
  // Convert HTML to text, preserving line breaks for <br> and paragraphs for <p>
  let text = htmlString
    .replace(/<br\s*[\/]?>/gi, "\n") // Convert <br> to line breaks
    .replace(/<p>/gi, "\n") // Start a new line for <p>
    .replace(/<\/p>/gi, "") // Do not need to replace closing </p> with anything
    .replace(/&nbsp;/gi, " "); // Replace &nbsp; with a space

  // Trim and remove multiple consecutive newlines
  text = text.replace(/\n\s*\n/g, "\n").trim();

  return text;
}

// Function to parse HTML and extract data
function parseHTML(htmlString) {
  const dom = new JSDOM(htmlString);
  const document = dom.window.document;
  const items = document.querySelectorAll(".group-items");
  return Array.from(items).map((item) => {
    const imgElement = item.querySelector(".items img");
    const imgSrc = imgElement ? imgElement.src : "N/A";
    const nameElement = item.querySelector(".name");
    const name = nameElement ? nameElement.textContent : "N/A";

    // Extract the item's id number to match with the modal's id
    const itemId = item.id.split("-")[1]; // Split 'item-0' to get '0'
    const modal = document.querySelector(`div[id="${itemId}"].modal.fade`);

    const txtElement = modal ? modal.querySelector(".txt") : null;
    const txt = txtElement ? formatText(txtElement.innerHTML) : "N/A";

    const { price, textWithoutPrice } = extractPriceFromText(txt);

    return { Image: imgSrc, Name: name, Text: textWithoutPrice, Price: price };
  });
}

// Function to fetch HTML content from a web page
async function fetchHTMLContent(url) {
  try {
    const response = await axios.get(url);
    return response.data; // return HTML content
  } catch (error) {
    console.error("Error fetching HTML content:", error);
    throw error;
  }
}

// Function to download image and return as a buffer
async function downloadImage(url) {
  const response = await axios({
    url,
    responseType: "arraybuffer",
  });
  return response.data;
}

async function writeToExcelWithImages(data, fileName) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Items");

  worksheet.columns = [
    { header: "Image", key: "image", width: 10 }, // Image column first
    { header: "Name", key: "name", width: 30 },
    { header: "Price", key: "price", width: 15 },
    { header: "Text", key: "text", width: 30 },
  ];

  // Convert 1.32 cm to row height in points
  const cmToInch = 1 / 2.54; // Convert cm to inches
  const inchToPoints = 72; // Convert inches to points
  const rowHeightInCm = 1.32;

  const rowHeightInPoints = rowHeightInCm * cmToInch * inchToPoints;

  for (let i = 0; i < data.length; i++) {
    console.info(`Processing item ${i + 1}/${data.length}: ${data[i].Name}`);

    const rowData = {
      name: data[i].Name,
      price: data[i].Price, // Add price data here
      text: data[i].Text,
    };
    const row = worksheet.addRow(rowData);

    // Set the row height to 1.32 cm in points
    row.height = rowHeightInPoints;

    if (data[i].Image !== "N/A") {
      const imageBuffer = await downloadImage(data[i].Image);
      const imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: "jpeg",
      });

      // Place the image in a column after 'Text', adjust cell reference as needed
      worksheet.addImage(imageId, {
        tl: { col: 0, row: i + 1 },
        ext: { width: 50, height: 50 },
      });
    }
  }

  await workbook.xlsx.writeFile(fileName);
}

// Main function to process the HTML file and create an Excel file
async function processHTMLContent(url) {
  try {
    const htmlContent = await fetchHTMLContent(url);
    const data = parseHTML(htmlContent);
    await writeToExcelWithImages(data, "items.xlsx");
    console.log("Excel file created successfully with images!");
  } catch (error) {
    console.error("Error processing the HTML file:", error);
  }
}

// URL of the web page to fetch HTML content
const url = "https://lienquan.garena.vn/trang-bi";
processHTMLContent(url);
