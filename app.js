const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');
const express = require("express");
var http = require("http");
worksheet.columns = [
  { header: 'Title', key: 'title', width: 30 },
  { header: 'Price', key: 'price', width: 20 },
  { header: 'Monthly rate', key: 'monthly_rate', width: 10 },
  { header: 'Attributes', key: 'attributes', width: 20 },
  { header: 'Description', key: 'description', width: 20 },
  { header: 'Status', key: 'status', width: 15 },
  { header: 'Location', key: 'location', width: 15 },
];

const app = express();
var server = http.createServer(app);

app.get("/export", async (req, res) => {
  // What is the file name
  const FILE_NAME = 'porsche.xlsx';
  // How many pages you want to loop
  // const PAGE = 5;
  const url = 'https://finder.porsche.com/de/de-DE/search?page=';
  const DATA = [];
  let page = 1;
  let hasData = true;

  while (hasData) {
    try {
      const { data } = await axios.get(url + page, {
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36',
        },
      });
      const $ = cheerio.load(data);
      const totalSectios = $('section').length;
      if (totalSectios === 0) {
        hasData = false;
        break;
      }
      $('section').each((index, element) => {
        const title = $(element).find('h2').prop('innerText');
        const price = $(element)
          .find('.eqyrYc ._1j9sent0._1j9sent6._1j9sentg')
          .prop('innerText');
        const monthly_rate = $(element)
          .find('.jrgyyB')
          .find('fnssr-p-text')
          .find('.fwaSxz')
          .prop('innerText');
        const attribute_span = $(element).find('.kLOBlW').find('span');
        let attributes = [];
        attribute_span.each((index, span) => {
          attributes.push($(span).prop('innerText'));
        });
        const description = $(element).find('.iuoUSR').prop('innerText');
        const status = $(element).find('.elNpsT').prop('innerText');
        const location = $(element).find('.knQEyZ').prop('innerText');
        const obj = {
          title,
          price,
          monthly_rate,
          attributes: attributes.join(', '),
          description,
          status,
          location,
        };
        DATA.push(obj);
      });
      page++;
      console.log('Total page:', page);
      res.send('Total page:', page);
    } catch (error) {
      console.log('Something went wrong');
      console.error(error?.response?.data?.message);
      hasData = false;
      break;
    }
  }

  DATA.forEach((row) => {
    worksheet.addRow(row);
  });
  // Save the workbook to a file
  workbook.xlsx
    .writeFile(FILE_NAME)
    .then(() => {
      console.log('Excel file created successfully.');
    })
    .catch((error) => {
      console.error('Error creating Excel file:', error);
    });
});

const port = process.env.PORT || 3000;

server.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
