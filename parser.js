// Require library
var excel = require("excel4node");
var fetch = require("node-fetch");

var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet("Sheet 1");

var style = workbook.createStyle({
  font: {
    color: "#000000",
    size: 14,
  },
});

var headStyle = workbook.createStyle({
  font: {
    color: "#000000",
    size: 14,
    bold: true,
  },
});

(async () => {
  try {
    worksheet.column(1).setWidth(50);
    worksheet.column(2).setWidth(20);
    worksheet.column(3).setWidth(20);

    worksheet.cell(1, 1).string("Банки").style(headStyle);
    worksheet.cell(1, 2).string("Покупка").style(headStyle);
    worksheet.cell(1, 3).string("Продажа").style(headStyle);

    worksheet.cell(2, 1).string("QQB").style(style);
    worksheet.cell(3, 1).string("Ipak Yuli Bank").style(style);
    worksheet.cell(4, 1).string("Sanoat Qurilish Bank ПСБ").style(style);
    worksheet.cell(5, 1).string("Ziraat Bank").style(style);

    await (async () => {
      try {
        const response = await fetch(
          "https://manage.qishloqqurilishbank.uz/api/currency-rates/last"
        );
        const result = await response.json();

        worksheet
          .cell(2, 2)
          .string(result.data.currency_rate.currencies[0].buy_rate)
          .style(style);
        worksheet
          .cell(2, 3)
          .string(result.data.currency_rate.currencies[0].sell_rate)
          .style(style);
      } catch (e) {
        console.log(e);
      }
    })();

    await (async () => {
        try {
          const response = await fetch("https://ipakyulibank.uz:8888/webapi/physical/exchange-rates", {
              method: 'POST',
              headers: {
                'X-AppKey': 'blablakey',
                'X-AppLang': 'uz',
                'X-AppRef': '/physical/valyuta-ayirboshlash/kurslar',
              },
          });
          const result = await response.json();
  
          worksheet.cell(3, 2).string(result.data.USD.rates['5'].Course.toString().slice(0, -2)).style(style);
          worksheet.cell(3, 3).string(result.data.USD.rates['4'].Course.toString().slice(0, -2)).style(style);
        } catch (e) {
          console.log(e);
        }
      })();

    await (async () => {
      try {
        const response = await fetch("https://sqb.uz/api/exchanges/");
        const result = await response.json();

        worksheet.cell(4, 2).string(result[0].psb_buy).style(style);
        worksheet.cell(4, 3).string(result[0].psb_sell).style(style);
      } catch (e) {
        console.log(e);
      }
    })();

    await (async () => {
      try {
        const response = await fetch(
          "https://www.ziraatbank.uz/tr/GetCurrency"
        );
        const result = await response.json();

        worksheet.cell(5, 2).string(result[0].value.toString()).style(style);
        worksheet.cell(5, 3).string(result[0].difference.toString()).style(style);
      } catch (e) {
        console.log(e);
      }
    })();

    workbook.write("Excel.xlsx");
  } catch (e) {
    console.log(e);
  }
})();
