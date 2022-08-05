const puppeteer = require("puppeteer");
const stringSimilarity = require("string-similarity");

(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: 0,
  });
  const page = await browser.newPage();
  const xlsx = require("xlsx");
  const file = xlsx.readFile("Input.xlsx");
  let data = [];
  let sheet = file.Sheets["Sheet1"];
  let sheetData = xlsx.utils.sheet_to_json(sheet);

  let ISBN = sheetData.map((item) => String(item.ISBN));

  let nameTitle = sheetData.map((item) => item.Title);

  const sortData = [];
  for (let i = 0; i < ISBN.length; i++) {
    await page.goto("https://www.snapdeal.com/", {
      waitUntil: "load",
      timeout: 60000,
    });
    await page.type("#inputValEnter", ISBN[i]);

    await Promise.all([
      page.click(".searchformButton.col-xs-4.rippleGrey"),
      page.waitForNavigation(),
    ]);
    const cData = [];
    const compareData = [];
    const listData = [];

    const divHandle = await page.$$(".product-desc-rating");
    for (const div of divHandle) {
      try {
        const link = await div.$eval("a", (a) => {
          return a.getAttribute("href");
        });
        const title = await div.$eval("a .product-title", (a) => {
          return a.innerText;
        });
        const price = await div.$eval("a .product-price", (a) => {
          return a.innerText;
        });
        let thisData = {
          name: title,
          price: price,
          id: link,
        };
        cData.push(thisData);
      } catch (error) {}
    } // end of for loop

    // Remove unnecessary data from the array

    for (let i = 0; i < cData.length; i++) {
      split = cData[i].name.split("-");
      cData[i].name = split[0];
      split2 = cData[i].name.split("(");
      cData[i].name = split2[0];
      split3 = cData[i].name.split(":");
      cData[i].name = split3[0];
      let tString = cData[i].name.toLowerCase();
      cData[i].name = tString.replace(/[^a-zA-Z0-9 *]/g, "").trim();
      let pData = cData[i].price.toString();
      cData[i].price = pData.replace(/[^0-9]/g, "");
      cData[i].price = parseInt(cData[i].price);

      //sorting data in ascending order to find the minimum price...
      const eitherSort = (cData = []) => {
        const sorter = (a, b) => {
          return +a.price - +b.price;
        };
        cData.sort(sorter);
      };
      eitherSort(cData);
      compareData.push(cData);
    }
    for (let i = 0; i < nameTitle.length; i++) {
      nameTitle[i] = nameTitle[i].toLowerCase();
      for (let j = 0; j < cData.length; j++) {
        let cString = cData[j].name;
        let sString = nameTitle[i];
        let similarity = stringSimilarity.compareTwoStrings(cString, sString);
        if (similarity > 0.9) {
          listData.push(cData[j]);
        }
      }
    }
    for (let i = 0; i < listData.length; i++) {
      if (listData[i].price <= listData[0].price) {
        sortData.push(listData[i]);
      }
    }
  } // end of for loop

  // new page for the sorted data...
  for (let i = 0; i < sortData.length; i++) {
    const page2 = await browser.newPage();
    await page2.goto(sortData[i].id, { waitUntil: "load", timeout: 60000 });
    page2.bringToFront();
    // fetching data from product page 
    let Url = page2.url();
      let Name =await page2.$eval(".col-xs-22 h1", (a) => {return a.innerText});
      Name = Name.split(new RegExp('[-+()/:?]', 'g'));
      Name = Name[0];
      Name = Name.toString().replace(/[^a-zA-Z0-9 *]/g, "");
      let Price = await page2.$eval(".pdp-e-i-PAY-r.disp-table-cell.lfloat .pdp-final-price span", (a) => {return a.innerText});
      Price = parseInt(Price.trim());
      let isbn = await page2.$eval(".spec-body.p-keyfeatures > ul > li:nth-child(1) > span.h-content", (a) => {return a.innerText});
      isbn = isbn.split(":");
      isbn = isbn[1].replace(/[^0-9]/g, "");
      let publisher = await page2.$eval(".spec-body.p-keyfeatures > ul > li:nth-child(3) > span.h-content", (a) => {return a.innerText});
      publisher = publisher.split(":");
      publisher = publisher[1];
      let author = await page2.$eval(".spec-body.p-keyfeatures > ul > li:nth-child(5) > span.h-content", (a) => {return a.innerText});
      author = author.split(":");
      author = author[1];
      let newdata = {
        No: i,
        Title: Name,
        ISBN: isbn,
        Site: "https://www.snapdeal.com/",
        Found: "yes",
        Url: Url,
        Price: Price,
        Author: author,
        Publisher: publisher,
        InStock: "yes",
      }
      data.push(newdata);
    page2.close();
  }// end of for loop


  // Writing data to an excel file...
  let newWb = xlsx.utils.book_new();
  let newWs = xlsx.utils.json_to_sheet(data);
  xlsx.utils.book_append_sheet(newWb, newWs, "Sheet1");
  xlsx.writeFile(newWb, "Output.xlsx");
  // end of writing data to an excel file...

  browser.close();
})(); // end of async function

