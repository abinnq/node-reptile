const fs = require('fs');
const request = require('request');
const cheerio = require('cheerio');
const querystring = require('querystring');
const XLSX = require('xlsx');
const nodeXlsx = require('node-xlsx');

const FILE_NAME = 'SheetJS.xlsx';
const KEYWORD= '鸟类群落';
const pageSize = 10;
let pageIndex = 1;


function getFirst() {
  const url = 'http://wap.cnki.net/touch/web/article/Search?kw=%E4%BA%BA%E5%B7%A5%E6%99%BA%E8%83%BD&field=101';
  const options = {
    method: 'GET',
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36',
      accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
      cacheControl: 'no-cache',
    }
  }
  request(url, options, function(err, res, body) {
    if(err || res.statusCode !== 200) {
      console.error(err);
    }
    readHtml(body);
  });
}

function getListByPage(pageIndex) {
  const countInfo = `(${(pageIndex - 1)*pageSize} - ${pageIndex*pageSize})`;
  // console.log(`列表获取中${countInfo}...`)
  const url = 'http://wap.cnki.net/touch/web/Article/Search';
  const form = {
    keyword: KEYWORD,
    fieldtype: 101,
    sorttype: 0,
    articletype: -1,
    yeartype: 0,
    screentype: 0,
    searchtype: 0,
    pageindex: pageIndex,
    pagesize: pageSize,
  };
  const formData = querystring.stringify(form);
  const contentLength = formData.length;

  const options = {
    uri: url,
    method: 'POST',
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36',
      accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
      cacheControl: 'no-cache',
      'Content-Length': contentLength,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body: formData,
  }
  request(options, (err, res, body) => {
    if(err || res.statusCode !== 200) {
      console.error(`列表${countInfo}获取失败!`, err);
    }
    readHtml(body).then(list => {
      console.log(`列表获取成功${countInfo}`)
      xlsxWrite(list).then(() => {
        if(list.length < pageSize) {
          console.log('数据全部写入成功!')
        } else {
          pageIndex++;
          getListByPage(pageIndex) 
        }
      }) 
    });
  })
}

async function readHtml(body) {
  var $ = cheerio.load(body);
  var items = [];
  await $('.c-book__person-outer .c-company__body-item').each(function (idx, element) {
    var $element = $(element).find('a');
    items.push([
      $element.first().children().eq(0).text().trim() || '',
      $element.first().children().eq(1).text().trim() || '',
      $element.last().text().trim() || '',
      $element[0].attribs.href || '',
    ]);
  });
  return items;
}

async function xlsxWrite(list = []) {
  await isFileExisted(FILE_NAME);
  var workbook = XLSX.readFile(FILE_NAME, { type: 'binary'});
  var first_sheet_name = workbook.SheetNames[0];
  var worksheet = workbook.Sheets[first_sheet_name];
  let result = XLSX.utils.sheet_to_json(worksheet, {header:1});

  const ws = XLSX.utils.aoa_to_sheet([...result, ...list]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet");
  await XLSX.writeFile(wb, FILE_NAME);
}

function isFileExisted(path_way) {
  return new Promise((resolve, reject) => {
    fs.access(path_way, (err) => {
      if (err) {
        const initTitles = ['title', 'user', 'downloadInfo', 'href'];
        const buffer = nodeXlsx.build([{name: "Sheet", data: [initTitles]}]);
        fs.appendFileSync(path_way, buffer);
        resolve(true);
      } else {
        resolve(true);
      }
    })
  })
};

async function main() {
  console.log(`${KEYWORD} 关键词查询中...`)
  getListByPage(pageIndex)
}
main();



