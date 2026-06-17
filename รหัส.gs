// ==========================================
// ⚙️ ส่วนตั้งค่า (CONFIG) 
// ==========================================
const ADMIN_EMAIL = "chumphonburicoopltd@gmail.com"; // ✉️ ใส่อีเมลที่ต้องการรับแจ้งเตือนที่นี่
const FOLDER_ID = "1nwG2OGUZDdJO6ze1MlYXjbd2kC067Y5s";     // 📁 ID โฟลเดอร์เก็บรูปสลิปใน Google Drive
// [NEW FEATURE] Start - ตั้งค่า PromptPay ของสหกรณ์ฯ
// ❌ ห้ามใส่เลขบัญชี (เช่น 015262938259) เพราะระบบจะสแกนไม่ติด
// ✅ ให้ใส่ "เบอร์โทรศัพท์ 10 หลัก" หรือ "เลขนิติบุคคล 13 หลัก" ที่ใช้ผูกกับ ธ.ก.ส.
const PROMPTPAY_ID = "0994000332076"; 
// [NEW FEATURE] End
const SHEET_ORDERS = "Orders";
const SHEET_PRODUCTS = "Products";

// ==========================================
// 1. ส่วนเชื่อมต่อหน้าเว็บ (API ROUTING) - 🌟 อัปเดตใหม่สำหรับ GitHub Pages
// ==========================================
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let result;

    // แยกการทำงานตาม Action ที่หน้าเว็บ (GitHub) ส่งมา
    if (action === 'getProductData') {
      result = getProductData();
    } else if (action === 'submitOrder') {
      result = submitOrder(request.data);
    } else if (action === 'getOrders') {
      result = getOrders();
    } else if (action === 'getProductsAdmin') {
      result = getProductsAdmin();
    } else if (action === 'updateOrderStatus') {
      result = updateOrderStatus(request.orderId, request.newStatus);
    } else if (action === 'saveProduct') {
      result = saveProduct(request.product);
    } else if (action === 'deleteProduct') {
      result = deleteProduct(request.id);
    } else if (action === 'getPromptPayQR') {
      result = getPromptPayQR(request.amount);
    } else {
      result = { error: 'ไม่พบคำสั่ง Action' };
    }

    // ส่งผลลัพธ์กลับไปที่หน้าเว็บ
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// เปิด doGet ไว้เพื่อให้ลิงก์เว็บแอปไม่ขึ้น Error เวลาคนเผลอกดเข้ามาตรงๆ
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'API is running successfully' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. ฟังก์ชันหลัก (API) สำหรับหน้าบ้าน (React)
// ==========================================

// ดึงข้อมูลสินค้า (แสดงเฉพาะที่ "พร้อมจำหน่าย" หรือ "สินค้าหมด")
function getProductData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ws) return []; 
  
  // ใช้ getDisplayValues เพื่อความชัวร์เรื่อง Format แต่ต้องระวังเรื่องตัวเลข
  const data = ws.getDataRange().getDisplayValues();
  const products = [];
  
  // เริ่มแถวที่ 1 (ข้าม Header)
  for (let i = 1; i < data.length; i++) {
    // โครงสร้าง: [ID, Name, Category, Price, Unit, Image, Status, Stock]
    // กรองสินค้าที่ "เลิกจำหน่าย" ออกไป
    if (data[i][0] && data[i][6] !== 'เลิกจำหน่าย') { 
      products.push({
        id: data[i][0],
        name: data[i][1],
        category: data[i][2],
        price: Number(data[i][3].replace(/,/g, '')), // ลบลูกน้ำออกก่อนแปลง
        unit: data[i][4],
        image: data[i][5],
        status: data[i][6],
        stock: Number(data[i][7].replace(/,/g, '')) || 0 
      });
    }
  }
  return products;
}

// 🛒 ฟังก์ชันรับออเดอร์ (ทำงานเมื่อลูกค้ากดสั่งซื้อ)
function submitOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(SHEET_ORDERS);
  
  // ถ้ายังไม่มี Sheet Orders ให้สร้างใหม่
  if (!ws) {
    ws = ss.insertSheet(SHEET_ORDERS);
    ws.appendRow(['Order ID', 'Date', 'Member ID', 'Name', 'Phone', 'Address', 'Items (JSON)', 'Total', 'Slip URL', 'Status']);
  }

  // 1. จัดการอัปโหลดรูปสลิป (ถ้ามี)
  let slipUrl = "";
  if (data.slipBase64) {
    try {
      if (FOLDER_ID && FOLDER_ID !== "YOUR_FOLDER_ID_HERE") {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        // แปลง Base64 กลับเป็นไฟล์รูปภาพ
        const contentType = data.slipBase64.substring(5, data.slipBase64.indexOf(';'));
        const base64Data = data.slipBase64.substr(data.slipBase64.indexOf('base64,')+7);
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, "Slip_" + new Date().getTime());
        
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // เปิดแชร์เพื่อให้ดูรูปได้
        slipUrl = file.getUrl();
      }
    } catch (e) {
      slipUrl = "Error Upload: " + e.toString();
    }
  }

  // 2. สร้าง Order ID และบันทึกลง Sheet
  const orderId = 'ORD-' + Math.floor(Math.random() * 1000000);
  const timestamp = new Date();
  
  ws.appendRow([
    orderId,
    timestamp,
    data.memberId || "-",
    data.memberName,
    "'" + data.phoneNumber, // ใส่ ' นำหน้าเพื่อป้องกัน Excel ตัดเลข 0
    data.address,
    JSON.stringify(data.cart),
    data.total,
    slipUrl,
    'รอยืนยัน'
  ]);

  // 3. ตัดสต๊อกสินค้า (Optional)
  try {
    updateStock(data.cart);
  } catch(e) { console.log("Stock Update Error: " + e.toString()); }
  
  // 4. 📧 ส่งอีเมลแจ้งเตือนแอดมิน (เรียกฟังก์ชันด้านล่าง)
  try {
    sendOrderEmail({
      id: orderId, 
      name: data.memberName, 
      phone: data.phoneNumber,
      address: data.address, 
      total: data.total, 
      items: data.cart, 
      slip: slipUrl
    });
  } catch (e) { 
    console.error("Email Error: " + e.toString()); 
  }
  
  return { status: 'success', orderId: orderId };
}

// ฟังก์ชันช่วยตัดสต๊อก
function updateStock(cartItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_PRODUCTS);
  const data = ws.getDataRange().getValues(); // ใช้ getValues ปกติสำหรับการเปรียบเทียบ ID
  
  cartItems.forEach(item => {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == item.id) {
        const currentStock = Number(data[i][7]) || 0;
        const newStock = Math.max(0, currentStock - item.qty);
        ws.getRange(i + 1, 8).setValue(newStock); // อัปเดตคอลัมน์ที่ 8 (Stock)
        break;
      }
    }
  });
}

// ==========================================
// 3. ฟังก์ชันสำหรับฝั่ง ADMIN (BACKEND API)
// ==========================================

function getOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_ORDERS);
  if (!ws) return [];
  // [NEW FIX] ใช้ getDisplayValues เพื่อให้ข้อมูลออกมาเป็น Text ทั้งหมด (แก้ปัญหาวันที่/ตัวเลขเพี้ยน)
  const data = ws.getDataRange().getDisplayValues();
  return data.length > 1 ? data.slice(1).reverse() : [];
}

function getProductsAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ws) { // สร้าง Sheet Products ถ้ายังไม่มี
    ws = ss.insertSheet(SHEET_PRODUCTS);
    ws.appendRow(['ID', 'Name', 'Category', 'Price', 'Unit', 'Image', 'Status', 'Stock']);
    return [];
  }
  
  // ตรวจสอบว่ามีคอลัมน์ Stock หรือยัง
  if (ws.getLastColumn() < 8) ws.getRange(1, 8).setValue("Stock");

  // [NEW FIX] ใช้ getDisplayValues และจัดการตัวเลขที่มีลูกน้ำ
  const data = ws.getDataRange().getDisplayValues();
  const products = [];
  
  for (let i = 1; i < data.length; i++) {
    products.push({
      id: data[i][0],
      name: data[i][1],
      category: data[i][2],
      price: Number(data[i][3].replace(/,/g, '')), // ลบลูกน้ำ
      unit: data[i][4],
      image: data[i][5],
      status: data[i][6],
      stock: Number(data[i][7].replace(/,/g, '')) || 0 // ลบลูกน้ำ
    });
  }
  return products;
}

function updateOrderStatus(orderId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_ORDERS);
  const data = ws.getDataRange().getValues(); // ใช้ getValues เพื่อเทียบ ID ได้แม่นยำ
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == orderId) {
      ws.getRange(i + 1, 10).setValue(newStatus); 
      return 'success';
    }
  }
  return 'not found';
}

function saveProduct(product) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ws) {
    ws = ss.insertSheet(SHEET_PRODUCTS);
    ws.appendRow(['ID', 'Name', 'Category', 'Price', 'Unit', 'Image', 'Status', 'Stock']);
  }

  const data = ws.getDataRange().getValues();
  let rowIndex = -1;

  if (product.id) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == product.id) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const rowData = [
    product.id || ('PROD-' + new Date().getTime()),
    product.name,
    product.category,
    Number(product.price),
    product.unit,
    product.image,
    product.status,
    Number(product.stock) || 0
  ];

  if (rowIndex !== -1) {
    ws.getRange(rowIndex, 1, 1, 8).setValues([rowData]);
  } else {
    ws.appendRow(rowData);
  }
  return 'success';
}

function deleteProduct(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_PRODUCTS);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      ws.deleteRow(i + 1);
      return 'success';
    }
  }
  return 'not found';
}

// ==========================================
// 4. ระบบส่งอีเมล (EMAIL SYSTEM)
// ==========================================
function sendOrderEmail(orderData) {
  try {
    const subject = `📢 ออเดอร์ใหม่! ${orderData.id} โดยคุณ ${orderData.name}`;
    
    let itemsHtml = "";
    if (orderData.items && orderData.items.length > 0) {
       itemsHtml = orderData.items.map(item => 
         `<tr>
            <td style="padding:8px; border-bottom:1px solid #eee;">${item.name}</td>
            <td style="padding:8px; text-align:right; border-bottom:1px solid #eee;">x${item.qty}</td>
          </tr>`
       ).join("");
    }

    const slipHtml = orderData.slip 
      ? `<div style="text-align:center; margin-top:20px; padding:10px; background:#f9fafb; border-radius:8px;">
           <p style="margin:0 0 10px 0; font-size:12px; color:#666;">หลักฐานการโอนเงิน:</p>
           <a href="${orderData.slip}" style="background:#2563eb; color:white; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold;">ดูสลิปการโอน</a>
         </div>` 
      : `<p style="text-align:center; color:#999; margin-top:20px;">- ไม่พบสลิปแนบ -</p>`;

    const htmlBody = `
      <div style="font-family:'Sarabun', sans-serif; max-width:600px; border:1px solid #e5e7eb; border-radius:12px; overflow:hidden; margin:0 auto; background-color:#ffffff;">
        <div style="background:#16a34a; padding:20px; text-align:center;">
          <h2 style="margin:0; color:white; font-size:20px;">🌱 รายการสั่งซื้อสินค้าใหม่</h2>
        </div>
        
        <div style="padding:25px;">
          <table style="width:100%; font-size:14px; color:#374151;">
            <tr><td style="color:#6b7280; width:100px; padding-bottom:5px;">รหัสคำสั่งซื้อ:</td><td style="font-weight:bold; padding-bottom:5px;">${orderData.id}</td></tr>
            <tr><td style="color:#6b7280; padding-bottom:5px;">ลูกค้า:</td><td style="padding-bottom:5px;">${orderData.name}</td></tr>
            <tr><td style="color:#6b7280; padding-bottom:5px;">เบอร์โทร:</td><td style="padding-bottom:5px;"><a href="tel:${orderData.phone}" style="color:#2563eb;">${orderData.phone}</a></td></tr>
            <tr><td style="color:#6b7280; vertical-align:top;">ที่อยู่จัดส่ง:</td><td>${orderData.address}</td></tr>
          </table>

          <div style="margin-top:20px; border:1px solid #e5e7eb; border-radius:8px; overflow:hidden;">
            <table style="width:100%; border-collapse:collapse; font-size:14px;">
              <thead style="background:#f3f4f6;">
                <tr><th style="padding:10px; text-align:left;">รายการสินค้า</th><th style="padding:10px; text-align:right;">จำนวน</th></tr>
              </thead>
              <tbody>${itemsHtml}</tbody>
            </table>
            <div style="padding:15px; text-align:right; border-top:1px solid #e5e7eb; background:#f9fafb;">
              ยอดรวมสุทธิ: <span style="font-size:18px; font-weight:bold; color:#16a34a;">฿${Number(orderData.total).toLocaleString()}</span>
            </div>
          </div>

          ${slipHtml}
          
          <div style="margin-top:30px; text-align:center; font-size:12px; color:#9ca3af;">
            ระบบจัดการสหกรณ์การเกษตรชุมพลบุรี
          </div>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      htmlBody: htmlBody
    });
    
  } catch (e) { console.error("Email Error: " + e.toString()); }
}

// [NEW FEATURE] Start - ฟังก์ชันสร้างลิงก์รูปภาพ PromptPay QR Code
function getPromptPayQR(amount) {
  if (!amount || amount <= 0) return "";
  // ใช้บริการ API มาตรฐานสำหรับสร้าง PromptPay QR (คืนค่าเป็น URL ของรูปภาพ)
  return "https://promptpay.io/" + PROMPTPAY_ID + "/" + amount + ".png";
}
// [NEW FEATURE] End