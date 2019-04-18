const request = require('request');
const cheerio = require('cheerio');
const Excel = require('exceljs');

let index = 0;
let title = [];
let director = [];
let actor = [];
let year = [];
let area = [];
let type = [];
let star = [];

// 获取cookie
function getCookie(callback) {
  request({
    url: 'https://movie.douban.com/top250'
  }, (err, res, html) => {
    var cookie = res.headers['set-cookie'][0].split(';')[0];
    callback(err, cookie);
  });
}

// 爬取信息
function getData(cookie) {
  request({
    url: 'https://movie.douban.com/top250?start=' + index + '&filter=',
    method: 'GET',
    headers: {
      cookie: cookie,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'
    }
  }, (err, res, html) => {
    $ = cheerio.load(html);
    // 解析电影名称
    $('.info .hd').each((index, element) => {
      title.push($(element).children().first().first().text().split('\n')[1].trim());
    })
    // 解析导演
    $('.info .bd').each((index, element) => {
      director.push($(element).children().first().text().split('\n')[1].split('导演: ')[1].split('主演: ')[0].trim());
    })
    // 解析主演
    $('.info .bd').each((index, element) => {
      actor.push($(element).children().first().text().split('\n')[1].split('主演: ')[1] === undefined ? '无' : $(element).children().first().text().split('\n')[1].split('主演: ')[1]);
    })
    // 解析年份
    $('.info .bd').each((index, element) => {
      year.push($(element).children().first().text().split('\n')[2].split('/')[0].trim());
    })
    // 解析国家地区
    $('.info .bd').each((index, element) => {
      area.push($(element).children().first().text().split('\n')[2].split('/')[1].trim());
    })
    // 解析电影类别
    $('.info .bd').each((index, element) => {
      type.push($(element).children().first().text().split('\n')[2].split('/')[2].trim());
    })
    // 解析评分人数
    $('.info .bd .star').each((index, element) => {
      star.push($(element).children().last().text());
    })

    // 将所有数据分别对应写入对象
    let arr = [title, director, actor, year, area, type, star]
    let allData = []
    for (var i = 0; i < 250; i++) {
      allData.push({
        id: i + 1,
        title: arr[0][i],
        director: arr[1][i],
        actor: arr[2][i],
        year: arr[3][i],
        area: arr[4][i],
        type: arr[5][i],
        star: arr[6][i]
      })
    }

    // 创建一个Workbook对象
    const workbook = new Excel.Workbook();
    // 创建一个worksheet并命名
    const worksheet = workbook.addWorksheet('My Sheet');
    // 设置列表属性
    worksheet.columns = [
      { header: '编号', key: 'id', width: 8 },
      { header: '电影名称', key: 'title', width: 12 },
      { header: '导演', key: 'director', width: 25 },
      { header: '主演', key: 'actor', width: 25 },
      { header: '年份', key: 'year', width: 8 },
      { header: '国家', key: 'area', width: 12 },
      { header: '分类', key: 'type', width: 25 },
      { header: '评分人数', key: 'star', width: 25 }
    ];
    // 行、列均从1开始计数，header作为第一行数据
    worksheet.addRows(allData);
    // 写入表格
    workbook.xlsx.writeFile('movie.xlsx')

    // 爬取1-10页所有数据
    index += 25;
    if (index <= 225) {
      getData();
    } else {
      console.log('over');
    }
  })
}

// 执行
getCookie((err, cookie) => {
  console.log(cookie);
  getData()
})