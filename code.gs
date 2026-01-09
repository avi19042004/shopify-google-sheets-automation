function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Run', 'checkTokenAndExecute')
      .addItem('Graph', 'countOrdersInRange')
      .addToUi();
}

function checkTokenAndExecute() {
  // Access the second sheet
  var sheet2 = SpreadsheetApp.openById('1zKw-w8MxZIAY947IvicuikpMbFKjmzft7lTItP-PaE8').getSheets()[0];
  
  // Get data from B column
  var range = sheet2.getRange("B:B");
  var values = range.getValues().flat().filter(Boolean);
  
  // Check if P4 value exists in the array
  var token = SpreadsheetApp.openById('1miO7IFjL9UW_pBGI6TSMoUljPHZLeh5ueiSNndzHstE').getSheetByName('Dashbord').getRange('A17').getValue();
  if (values.includes(token)) {
    // Token exists, execute the existing script
    fetchShopifyOrders();
  } else {
    // Token doesn't exist, show a message
    SpreadsheetApp.getUi().alert('Check your Token.');
  }
}

function fetchShopifyOrders() {
  var spreadsheet = SpreadsheetApp.openById('1miO7IFjL9UW_pBGI6TSMoUljPHZLeh5ueiSNndzHstE');
  var sheet = spreadsheet.getSheetByName('Order');
  var sheet1 = spreadsheet.getSheetByName('Dashbord');

  var shopifyApiKey = sheet1.getRange('A2').getValue();
  var shopifyPassword = sheet1.getRange('A5').getValue();
  var shopifyApiUrl = sheet1.getRange('A8').getValue();

  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(shopifyApiKey + ':' + shopifyPassword)
  };

  try {
    var response = UrlFetchApp.fetch(shopifyApiUrl, {
      headers: headers
    });

    var responseCode = response.getResponseCode();
    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      var orders = data.orders;

      var lastRow = sheet.getLastRow(); 

      var headers = ['Order Date', 'Order ID', 'Customer Name', 'Customer Mobile', 'Tax', 'Address', 'Total Price', 'Order Status', 'Tracking Number', 'IP Address', 'Product Names', 'Whatsapp', 'Tracking URL']; // Added 'Tracking ID' to headers

      if (lastRow === 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }

      var processedOrderIds = [];
      var totalPriceSum = 0;

      var existingOrderIds = sheet.getRange("B2:B" + lastRow).getValues().flat() || [];

      existingOrderIds.forEach(function(orderId) {
        processedOrderIds.push(orderId.toString());
      });

      var orderData = [];
      var orderCounts = {};

      for (var i = 0; i < orders.length; i++) {
        var order = orders[i];
        var customer = order.customer;

        var orderId = order.order_number.toString(); // Fetch order number instead of order ID
        var orderDate = new Date(order.created_at); // Get the order creation date

        if (processedOrderIds.indexOf(orderId) === -1) {
          var fulfillmentStatus = 'Yet to Ship';
          if (order.fulfillment_status === 'fulfilled') {
            fulfillmentStatus = 'In Transit';
          } else if (order.fulfillment_status === 'delivered') {
            fulfillmentStatus = 'Delivered';
          } else if (order.fulfillment_status === 'in_transit') {
            fulfillmentStatus = 'In Transit';
          }

          var courierNumber = 'No tracking number found';
          var trackingID = ''; // Added tracking ID variable

          if (order.fulfillments && order.fulfillments.length > 0) {
            if (order.fulfillments[0].tracking_number) {
              courierNumber = order.fulfillments[0].tracking_number;
              trackingID = courierNumber; // Set tracking ID
            }
          }

          var lineItems = order.line_items;
          var productNames = lineItems.map(function(item) {
            return item.name;
          });

          // Update orderCounts with the count of each product name
          productNames.forEach(function(productName) {
            if (orderCounts[productName]) {
              orderCounts[productName]++;
            } else {
              orderCounts[productName] = 1;
            }
          });

          var orderTotalPrice = parseFloat(order.total_price);
          totalPriceSum += orderTotalPrice;

          var trackingURL = '';

          if (order.fulfillments && order.fulfillments.length > 0) {
            if (order.fulfillments[0].tracking_url) {
              trackingURL = order.fulfillments[0].tracking_url;
            }
          }

          var rowData = [
            orderDate,
            orderId,
            customer.first_name + ' ' + customer.last_name,
            customer.phone,
            order.total_tax,
            customer.default_address.address1,
            orderTotalPrice,
            fulfillmentStatus,
            courierNumber,
            (order.client_details && order.client_details.browser_ip) ? order.client_details.browser_ip : 'No IP address available',
            productNames.join('\n'), // Combine product names into a single string
            trackingID,
            trackingURL
          ];
          orderData.push(rowData);

          processedOrderIds.push(orderId);
        }
      }

      if (orderData.length > 0) {
        sheet.getRange(lastRow + 1, 1, orderData.length, headers.length).setValues(orderData);
        sendTelegramNotifications(orderData); 
        sheet1.getRange('Z10').clearContent();
        var totalPriceSum = orderData.reduce(function(acc, order) {
          return acc + order[5];
        }, 0);

        sheet1.getRange('D11').setValue(parseFloat(totalPriceSum));

        // Update Dashboard sheet with product counts
        var existingProductCounts = sheet1.getRange("D33:E").getValues(); // Assuming product counts start from D24
        var productCounts = {};

        if (existingProductCounts) {
          existingProductCounts.forEach(function(row) {
            var productName = row[0];
            var count = row[1];
            productCounts[productName] = count;
          });
        }

        // Update Dashboard sheet with product counts
        Object.keys(orderCounts).forEach(function(productName) {
          if (productCounts[productName]) {
            productCounts[productName] += orderCounts[productName];
          } else {
            productCounts[productName] = orderCounts[productName];
          }
        });

        var dashboardData = Object.keys(productCounts).map(function(productName) {
          return [productName, productCounts[productName]];
        });

        var dashboardRange = sheet1.getRange(33, 4, dashboardData.length, 2); // Assuming product counts start from D24
        dashboardRange.setValues(dashboardData);

        console.log('Total Price Sum: ' + totalPriceSum);
        console.log('Shopify orders fetched and added to the Google Sheet.');
      }

    } else {
      console.log('Failed to fetch Shopify orders. Response code: ' + responseCode);
    }
  } catch (e) {
    console.log('Error: ' + e.toString());
  }
  createWhatsAppHyperlink();
}

// ... (rest of your code)


function sendTelegramNotifications(newOrders) {
  var spreadsheet = SpreadsheetApp.openById('1miO7IFjL9UW_pBGI6TSMoUljPHZLeh5ueiSNndzHstE');
  var sheet = spreadsheet.getSheetByName('Order');
  var sheet1 = spreadsheet.getSheetByName('Dashbord');
  var botToken = sheet1.getRange('A11').getValue();
  var chatId = sheet1.getRange('A14').getValue();

  var telegramApiUrl = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  
  for (var i = 0; i < newOrders.length; i++) {
    var order = newOrders[i];
    var message = 'New Order Received:\n';
    message += 'Order ID: ' + order[1] + '\n';
    message += 'Customer Name: ' + order[2] + '\n';
    message += 'Phone Number: ' + order[3] + '\n';
    message += 'Address: ' + order[5] + '\n';
    message += 'Total Price: ' + order[6] + '\n';
    message += 'Product Names:\n' + order[10] + '\n';
    message += 'Order Status URL:\n' + order[12] + '\n'; 
    
    var payload = {
      'chat_id': chatId,
      'text': message,
    };

    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
    };

    UrlFetchApp.fetch(telegramApiUrl, options);
  }
}


function createWhatsAppHyperlink() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.openById('1miO7IFjL9UW_pBGI6TSMoUljPHZLeh5ueiSNndzHstE');
  var sheet1 = spreadsheet.getSheetByName('Order');
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 12); // Updated to include the new column

  var data = dataRange.getValues();
  var whatsappLinks = [];

  for (var i = 0; i < data.length; i++) {
    var phoneNumber = data[i][3];
    var customerName = data[i][2];
    var productNames = data[i][10];
    var totalPrice = data[i][6];
    var orderStatusUrl = data[i][11];
    var orderDate = data[i][0];
    var trackingNumber = data[i][8];

    var message = 'Hi ' + customerName + ',\n\n';
    message += 'Thank you for your order. We have received the order of following items:\n';
    message += productNames + 'on ' + orderDate + '\n';
    message += 'Total Price: ₹ ' + totalPrice + '\n';
     message += 'Your Tracking Id is : ' + trackingNumber + '\n';
    message += 'You can track your order status [here](' + orderStatusUrl + ').';

    var whatsappLink = "https://api.whatsapp.com/send?phone=" + phoneNumber + "&text=" + encodeURIComponent(message);
    var displayText = "Click to send";
    var hyperLinkFormula = '=HYPERLINK("' + whatsappLink + '", "' + displayText + '")';
    whatsappLinks.push([hyperLinkFormula]);
  }

  var columnL = sheet1.getRange(2, 12, whatsappLinks.length, 1);
  columnL.setFormulas(whatsappLinks);
}

function countOrdersInRange() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashbord'); // Change the source sheet name if needed
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashbord');
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order');
  
  var startDate = new Date(sourceSheet.getRange('A20').getValue());
  var endDate = new Date(sourceSheet.getRange('A23').getValue());

  var lastRow = ss.getLastRow();
  var lastColumn = ss.getLastColumn();
  var values = ss.getRange(1, 1, lastRow, lastColumn).getValues();
  
  // Initialize an object to store order counts by date
  var orderCounts = {};

  // Iterate through the data
  for (var i = 0; i < values.length; i++) {
    var currentDate = new Date(values[i][0]);
    var productName = values[i][10]; // Get product name from column K
    
    // Check if currentDate is a valid date and it falls within the date range
    if (!isNaN(currentDate.getTime()) && currentDate >= startDate && currentDate <= endDate) {
      var dateString = currentDate.toDateString();
      if (!orderCounts[dateString]) {
        orderCounts[dateString] = {
          count: 0,
          products: [] // Initialize an array to store products for each date
        };
      }
      orderCounts[dateString].count++;
      orderCounts[dateString].products.push({productName: productName}); // Add product info
    }
  }

  // Write the data to the Chart sheet
  var resultData = [];

  for (var date in orderCounts) {
    resultData.push([date, orderCounts[date].count]);
  }

  if (resultData.length > 0) {
    var targetRange = targetSheet.getRange(26, 1, targetSheet.getLastRow(), 2); // Selects range A25:B999
    targetRange.clearContent(); // Clears the content in the selected range\

    var targetRange = targetSheet.getRange(26, 1, resultData.length, resultData[0].length);
    targetRange.setValues(resultData);
  }
}

function checkShopifyForUpdates() {
  var updatedOrders = fetchUpdatedOrders();
  
  if (updatedOrders.length > 0) {
    updatedOrders.forEach(function(order) {
      updateTrackingInfo(order.orderId, order.trackingNumber, order.trackingURL, order.orderStatus);
    });
  }
}
