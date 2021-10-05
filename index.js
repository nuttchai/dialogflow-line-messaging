const functions = require("firebase-functions");
const { WebhookClient, Payload } = require("dialogflow-fulfillment");
const { google } = require("googleapis");

process.env.DEBUG = "dialogflow:debug"; // enables lib debugging statements

exports.dialogflowFirebaseFulfillment = functions.https.onRequest(
  async (request, response) => {
    const agent = new WebhookClient({ request, response });
    const auth = await google.auth.getClient({
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const sheetsAPI = google.sheets({ version: "v4", auth });
    const spreadsheetId = "1G3ldv6UJPnGRGET4KUaqoSVyRahXnvPQMehVmPxihMA";

    console.log(
      "Dialogflow Request -headers: " + JSON.stringify(request.headers)
    );
    console.log("Dialogflow Request body: " + JSON.stringify(request.body));

    function welcome(agent) {
      agent.add(`Welcome to my agent!`);
    }

    function fallback(agent) {
      agent.add(`I didn't understand`);
      agent.add(`I'm sorry, can you try again?`);
    }

    // "ฟังชั่นสำหรับการเพิ่มข้อมูลสต๊อก"
    async function inputStock(agent) {
      const productName = agent.parameters.productName;
      const productQuantity = agent.parameters.productQuantity;
      const productPrice = agent.parameters.productPrice;

      let values = [
        // สร้างข้อมูล row ใน sheet
        [
          getCurrentDate(),
          productName.toString(),
          productQuantity.toString(),
          productPrice.toString(),
          productQuantity.toString(),
        ],
      ];

      try {
        // สั่งเพิ่มข้อมูลในชีทด้วยข้อมูลที่สร้าง
        await appendStock(values);

        // ถ้าสำเร็จให้ตอบว่าเพิ่มสินค้าเรียบร้อย
        agent.add(`เพิ่มสินค้าเรียบร้อย`);
      } catch (err) {
        // ถ้าไม่สำเร็จให้ตอบว่าไม่สำเร็จเพราะอะไร
        agent.add(`${err}`);
      }
    }

    // "ฟังชั่นสำหรับการเพิ่มบันทึกการขาย และตัดสต๊อกตามจำนวนที่ขาย"
    async function inputSale(agent) {
      const productName = agent.parameters.productName;
      const soldPrice = agent.parameters.price;

      let existingStock = await readAllStockRows();
      let soldQuantity = agent.parameters.productQuantity;

      // บันทึกข้อมูลการขาย
      await appendSale([
        [getCurrentDate(), productName, soldQuantity, soldPrice],
      ]);

      // ตัดสต๊อก
      try {
        while (soldQuantity != 0) {
          const index = existingStock.findIndex(
            (row) => row[1].toString().replace(/\s/g, '') == productName && row[4] != 0
          );

          if (index == -1) {
            throw new Error("ไม่มีสินค้าในระบบให้ตัดสต๊อก");
          }

          const existingQuantity = parseInt(existingStock[index][4]);

          if (existingQuantity - soldQuantity >= 0) {
            const newQuantity = existingQuantity - soldQuantity;
            existingStock[index][4] = newQuantity;
            soldQuantity = 0;
            break;
          } else {
            soldQuantity -= existingQuantity;
            existingStock[index][4] = 0;
          }
        }

        // ลบข้อมูลสต๊อกเก่า
        await deleteAllStockRows();
        // เพิ่มข้อมูลสต๊อกใหม่ไปแทน
        await appendStock(existingStock);

        // ถ้าสำเร็จให้ตอบว่าบันทึกรายการสำเร็จ
        agent.add("บันทึกรายการสำเร็จ");
      } catch (err) {
        // ถ้าไม่สำเร็จให้ตอบว่าไม่สำเร็จเพราะอะไร
        agent.add(`${err}`);
      }
    }

    // "ฟังชั่นสำหรับการสรุปสต๊อก"
    async function outputStock(agent) {
      const result = await getStockSummaryText(); // ของที่ return = { "resultText" : result, "payload" : payload }
      const stockSummaryText = result.resultText;
      const payload = result.payload;

      if (stockSummaryText.length > 10 || payload != null) {
        agent.add(payload);
        // agent.add(stockSummaryText);
      } else {
        agent.add("ไม่มีสินค้าในสต๊อก");
      }
    }

    // "ฟังชั่นสำหรับการสรุปต้นทุน
    async function outputCost(agent) {
        const result = await getCostSummaryText();
        const costSummaryText = result.resultText;
        const payload = result.payload;
  
        if (costSummaryText.length > 10 || payload != null) {
          agent.add(payload);
          // agent.add(costSummaryText);
        } else {
          agent.add("ไม่มีสินค้าในสต๊อก");
        }
      }

    // "ฟังชั่นสำหรับการคำนวนราคาขายจากราคาทุนและกำไรที่ต้องการ"
    function computeProfit(agent) {
      const price = agent.parameters.price;
      const profit = agent.parameters.profit;
      const sellingPrice = price * (profit / 100 + 1);

      agent.add(`สินค้าราคา ${price} กำไร ${profit}% ต้องขาย ${sellingPrice}`);
    }

    async function deleteAllStockRows() {
      const request = {
        spreadsheetId: spreadsheetId,
        range: "stock!A2:G",
        auth: auth,
      };

      (await sheetsAPI.spreadsheets.values.clear(request)).data;
    }

    async function appendSale(values) {
      const request = {
        spreadsheetId: spreadsheetId,
        range: "sales!A1:D1",
        valueInputOption: "RAW",
        resource: {
          values: values,
        },
        auth: auth,
      };

      await sheetsAPI.spreadsheets.values.append(request);
    }

    async function appendStock(values) {
      const request = {
        spreadsheetId: spreadsheetId,
        range: "stock!A1:E1",
        valueInputOption: "RAW",
        resource: {
          values: values,
        },
        auth: auth,
      };

      await sheetsAPI.spreadsheets.values.append(request);
    }

    function getCurrentDate() {
      const date = new Date();
      const month = (date.getMonth() + 1).toString().padStart(2, "0");
      const day = date.getDate().toString().padStart(2, "0");
      const year = date.getFullYear().toString().slice(2);
      const formattedDate = day + "/" + month + "/" + year;

      return formattedDate;
    }

    async function readAllStockRows() {
      const request = {
        spreadsheetId: spreadsheetId,
        range: "stock!A2:E",
        majorDimension: "ROWS",
        auth: auth,
      };

      try {
        const response = await sheetsAPI.spreadsheets.values.get(request);

        return response.data.values;
      } catch (err) {
        throw new Error(`Read Stock Error because ${err}`);
      }
    }

    async function getStockSummaryText() {
      let templateStock = {};

      const existingStock = (await readAllStockRows()).filter(
        (row) => parseInt(row[4]) != 0
      );
      
      const stockSummaryData = existingStock.reduce((result, row) => {
        const date = row[0];
        const name = row[1];
        const price = row[3];
        const quantity = row[4];

        if (name in result) {
          result[name]["textFifo"] += `\n  ${date} ${quantity}@${price}`;
          templateStock[name].fifo.push({ 
            date: date, 
            value: `${quantity}@${price}`
          });

        } else {
          result[name] = {};

          result[name]["textFifo"] = ` FIFO\n  ${date} ${quantity}@${price}`;
          templateStock[name] = {
            name: name.replace(/\s/g, ''),
            fifo: [{
              date: `${date}`,
              value: `${quantity}@${price}`
            }],
            wa: {
              date: "",
              value: ""
            },
            itemLeft: ""
          }
        }

        return result;
      }, {});
      
      Object.keys(stockSummaryData).forEach((name) => {
        const rows = existingStock.filter((row) => row[1] == name);
        const stockCount = rows.reduce((sum, row) => sum + parseInt(row[4]), 0);
        const costSum = rows.reduce(
          (sum, row) => sum + parseInt(row[3] * row[4]),
          0
        );
        const avgPrice = parseFloat(costSum / stockCount).toFixed(2);
        const date = rows[rows.length - 1][0];

        stockSummaryData[name][
          "textWA"
        ] = ` WA\n  ${date} ${stockCount}@${avgPrice}`;
        stockSummaryData[name]["stockCount"] = stockCount;

        templateStock[name].wa.date = `${date}`;
        templateStock[name].wa.value = `${stockCount}@${avgPrice}`;
        templateStock[name].itemLeft = `${stockCount}`;
      });

      const result = Object.keys(stockSummaryData)
        .sort()
        .reduce((summaryString, name) => {
          return summaryString.concat(
            `\n สินค้า ${name}\n ${stockSummaryData[name].textFifo}\n\n ${stockSummaryData[name].textWA}\n\n  คงเหลือ ${stockSummaryData[name].stockCount} ชิ้น\n`
          );
        }, "คลังสินค้า\n");
      
      const payload = await createPayloadStockSummary(templateStock);

      return { "resultText" : result, "payload" : payload };
    }

    async function getCostSummaryText() {
      const templateCost = {
        "products" : [],
        "summary" : {
            "fifoSum" : "0",
            "waSum" : "0"
        }
      };

      const existingStock = (await readAllStockRows()).filter(
        (row) => row[4] != 0
      );

      let fifoSum = 0;
      let waSum = 0;

      const stockFIFO = existingStock.reduce((result, row) => {
        const name = row[1];

        result[name] = {};

        return result;
      }, {});

      Object.keys(stockFIFO).forEach((name) => {
        const rows = existingStock.filter((row) => row[1] == name);
        const stockCount = parseInt(
          rows.reduce((sum, row) => sum + parseInt(row[4]), 0)
        );
        const costSum = parseInt(
          rows.reduce((sum, row) => sum + parseInt(row[3] * row[4]), 0)
        );
        const avgPrice = parseFloat(costSum);
        
        templateCost.products.push(
            {
                "name" : name.toString().replace(/\s/g, ''),
                "fifo" : costSum.toString(),
                "wa" : avgPrice.toFixed(2).toString()
            }
        )

        stockFIFO[name]["textFIFO"] = ` FIFO ${costSum}`;
        stockFIFO[name]["textWA"] = ` WA ${avgPrice.toFixed(2)}`;
        fifoSum += costSum;
        waSum += avgPrice;
      });

      const result =
        Object.keys(stockFIFO)
          .sort()
          .reduce((summaryString, name) => {
            return summaryString.concat(
              `\n สินค้า ${name}\n ${stockFIFO[name].textFIFO}\n ${stockFIFO[name].textWA}\n`
            );
          }, "ต้นทุน\n") + `\nรวม FIFO ${fifoSum}\nรวม WA ${waSum.toFixed(2)}`;
      
      templateCost.summary.fifoSum = fifoSum.toString();
      templateCost.summary.waSum = waSum.toFixed(2).toString();

      const payload = await createPayloadCostSummary(templateCost);
          
      // เราต้องการจะ return 2 อย่าง
      return { "resultText" : result, "payload" : payload };
    }

    async function createPayloadStockSummary(templateStock) {      
      if (!Object.keys(templateStock).length) {
        return null;
      }

      let bodyContent = [];
      
      // loop list สินค้่า
      for (const property in templateStock)  {
        const productInfo = templateStock[property];

        bodyContent.push(
          {
            "type": "text",
            "text": `สินค้า ${productInfo.name}`,
            "weight": "bold",
            "align": "start",
            "margin": "md",
            "decoration": "underline",
            "contents": []
          },
          createItem("FIFO"),
        );
        
        // loop list ของข้างใน fifo
        productInfo.fifo.forEach(eachFifo => {
          bodyContent.push(
            createItem(eachFifo.date, eachFifo.value)
          )
        });

        bodyContent.push(
          createItem("WA"),
          createItem(productInfo.wa.date, productInfo.wa.value),
          createItem("คงเหลือ", `${productInfo.itemLeft} ชิ้น`)
        );
      }

      const payloadJson = 
        {
          "type": "flex",
          "altText": "Flex Message",
          "contents": {
            "type": "bubble",
            "direction": "ltr",
            "header": {
              "type": "box",
              "layout": "vertical",
              "contents": [
                {
                    "type": "text",
                    "text": "คลังสินค้า",
                    "weight": "bold",
                    "size": "xl",
                    "align": "center",
                    "margin": "md",
                    "contents": []
                },
                {
                  "type": "separator",
                  "margin": "xxl"
                },
                {
                  "type": "box",
                  "layout": "vertical",
                  "margin": "xxl",
                  "spacing": "sm",
                  "contents": bodyContent
                }
              ]
            },
            "styles": {
              "footer": {
                "separator": true
              }
            }
          }
        }

        let payload = new Payload('LINE', payloadJson, {sendAsMessage: true});
  
        return payload; 
    }

    async function createPayloadCostSummary(templateCost) {
        if (!templateCost.products.length) { 
            return null;
        }
        
        let bodyContent = [];
        
        templateCost.products.forEach(product => {
          bodyContent.push(
            createItem(`สินค้า ${product.name}`),
            createItem("FIFO", product.fifo),
            createItem("WA", product.wa)
          );
        });

        bodyContent.push(
          {
            "type": "separator",
            "margin": "xxl"
          },
          {
            "type": "box",
            "layout": "horizontal",
            "margin": "xxl",
            "contents": [createItem("รวม FIFO", templateCost.summary.fifoSum)]
          },
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [createItem("รวม WA", templateCost.summary.waSum)]
          }
        )
  
        const payloadJson = 
        {
          "type": "flex",
          "altText": "Flex Message",
          "contents": {
            "type": "bubble",
            "body": {
              "type": "box",
              "layout": "vertical",
              "contents": [
                {
                  "type": "text",
                  "text": "รายงานต้นทุน",
                  "weight": "bold",
                  "size": "xxl",
                  "margin": "md"
                },
                {
                  "type": "separator",
                  "margin": "xxl"
                },
                {
                  "type": "box",
                  "layout": "vertical",
                  "margin": "xxl",
                  "spacing": "sm",
                  "contents": bodyContent
                }
              ]
            },
            "styles": {
              "footer": {
                "separator": true
              }
            }
          }
        };
        
        let payload = new Payload('LINE', payloadJson, {sendAsMessage: true});
  
        return payload;
    }
    
    function createItem(title, value="") {
        let result = {
            "type": "box",
            "layout": "horizontal",
            "margin": "sm",
            "contents": [
                {
                    "type": "text",
                    "text": title,
                    "size": "sm",
                    "color": "#555555"
                }
            ]
        };

        if (value) {
            result.contents.push({
                "type": "text",
                "text": value,
                "size": "sm",
                "color": "#111111",
                "align": "end"
            });
        }

        return result;
    }

    // Run the proper function handler based on the matched Dialogflow intent name
    let intentMap = new Map();

    intentMap.set("Default Welcome Intent", welcome);
    intentMap.set("Default Fallback Intent", fallback);
    intentMap.set("input_stock", inputStock);
    intentMap.set("input_sale", inputSale);
    intentMap.set("output_stock", outputStock);
    intentMap.set("output_cost", outputCost);
    intentMap.set("compute_profit", computeProfit);

    agent.handleRequest(intentMap);
  }
);
