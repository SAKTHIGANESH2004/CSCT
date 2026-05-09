import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";
import pg from "pg";
import Groq from "groq-sdk";
import multer from "multer";
import * as XLSX from "xlsx";
import path from "path";
import { fileURLToPath } from "url";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const upload = multer({ dest: "uploads/" });
const groq = new Groq({ apiKey: process.env.GROQ_KEY });
const pool = new pg.Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.NODE_ENV === "production" ? { rejectUnauthorized: false } : false,
});

// ═══════════════════════════════════════
// DATABASE
// ═══════════════════════════════════════
async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS known_customers (
      id SERIAL PRIMARY KEY,
      phone TEXT UNIQUE NOT NULL,
      name TEXT NOT NULL,
      category TEXT NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )`);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS price_list (
      id SERIAL PRIMARY KEY,
      product TEXT NOT NULL,
      brand TEXT,
      category TEXT,
      dealer_price NUMERIC DEFAULT 0,
      rp_lt_price NUMERIC DEFAULT 0,
      end_user_price NUMERIC DEFAULT 0,
      updated_at TIMESTAMP DEFAULT NOW()
    )`);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS new_enquiries (
      id SERIAL PRIMARY KEY,
      phone TEXT NOT NULL,
      name TEXT,
      location TEXT,
      referred_by TEXT,
      requirement TEXT,
      created_at TIMESTAMP DEFAULT NOW()
    )`);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS chat_sessions (
      id SERIAL PRIMARY KEY,
      phone TEXT UNIQUE NOT NULL,
      step TEXT DEFAULT 'start',
      temp_name TEXT,
      temp_location TEXT,
      temp_referral TEXT,
      updated_at TIMESTAMP DEFAULT NOW()
    )`);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS messages (
      id SERIAL PRIMARY KEY,
      phone TEXT NOT NULL,
      role TEXT NOT NULL,
      content TEXT NOT NULL,
      created_at TIMESTAMP DEFAULT NOW()
    )`);

  // Sample customers - update phone numbers via admin panel
  await pool.query(`
    INSERT INTO known_customers (phone, name, category) VALUES
    ('917000000001', 'Raj', 'Dealer'),
    ('917000000002', 'Mohan', 'RP/LT'),
    ('917000000003', 'Rudhran', 'End User')
    ON CONFLICT (phone) DO NOTHING`);

  // Sample prices - update via Excel upload
  await pool.query(`
    INSERT INTO price_list (product,brand,category,dealer_price,rp_lt_price,end_user_price) VALUES
    ('WiFi Dome Camera 2MP','CP Plus','Camera',2720,2560,2400),
    ('Bullet Camera 4MP','Hikvision','Camera',3825,3600,3375),
    ('PTZ Camera 2MP','Dahua','Camera',7225,6800,6375),
    ('4CH DVR 1080P','CP Plus','DVR/NVR',4080,3840,3600),
    ('8CH NVR 4K','Hikvision','DVR/NVR',8075,7600,7125),
    ('Solar WiFi Camera','CP Plus','Camera',4420,4160,3900),
    ('Indoor Dome 2MP','Hikvision','Camera',2380,2240,2100),
    ('16CH DVR 5MP','CP Plus','DVR/NVR',6970,6560,6150),
    ('2TB Surveillance HDD','Seagate','Storage',3570,3360,3150),
    ('Video Door Phone','CP Plus','Accessories',3230,3040,2850)
    ON CONFLICT DO NOTHING`);

  console.log("✅ Database ready!");
}

// ═══════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════
async function getCustomer(phone) {
  const r = await pool.query("SELECT * FROM known_customers WHERE phone=$1", [phone]);
  return r.rows[0] || null;
}

async function getSession(phone) {
  const r = await pool.query("SELECT * FROM chat_sessions WHERE phone=$1", [phone]);
  return r.rows[0] || null;
}

async function setSession(phone, step, data = {}) {
  const ex = await getSession(phone);
  if (ex) {
    await pool.query(
      `UPDATE chat_sessions SET step=$1,
       temp_name=COALESCE($2,temp_name),
       temp_location=COALESCE($3,temp_location),
       temp_referral=COALESCE($4,temp_referral),
       updated_at=NOW() WHERE phone=$5`,
      [step, data.name || null, data.location || null, data.referral || null, phone]
    );
  } else {
    await pool.query(
      `INSERT INTO chat_sessions(phone,step,temp_name,temp_location,temp_referral)
       VALUES($1,$2,$3,$4,$5)`,
      [phone, step, data.name || null, data.location || null, data.referral || null]
    );
  }
}

async function getPrices() {
  const r = await pool.query("SELECT * FROM price_list ORDER BY category,brand");
  return r.rows;
}

function formatPrices(products, category) {
  const field =
    category === "Dealer" ? "dealer_price" :
    category === "RP/LT" ? "rp_lt_price" : "end_user_price";

  const grouped = {};
  products.forEach(p => {
    const cat = p.category || "General";
    if (!grouped[cat]) grouped[cat] = [];
    grouped[cat].push(p);
  });

  let msg = `📋 *Classic Traders — Price List*\n`;
  msg += `🏷 Type: *${category}*\n`;
  msg += `📅 ${new Date().toLocaleDateString("en-IN")}\n\n`;

  Object.keys(grouped).forEach(cat => {
    msg += `*📦 ${cat}:*\n`;
    grouped[cat].forEach(p => {
      const price = parseFloat(p[field] || 0).toLocaleString("en-IN");
      msg += `• ${p.brand ? p.brand + " " : ""}${p.product}: ₹${price}\n`;
    });
    msg += "\n";
  });

  msg += `_⚠️ Prices valid this week only._\n`;
  msg += `_📞 Bulk orders: 7603868752_`;
  return msg;
}

async function sendWA(to, text) {
  try {
    const r = await fetch(
      `https://graph.facebook.com/v18.0/${process.env.PHONE_ID}/messages`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${process.env.WA_TOKEN}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          messaging_product: "whatsapp",
          to,
          type: "text",
          text: { body: text },
        }),
      }
    );
    const d = await r.json();
    if (d.error) console.error("❌ WA error:", d.error.message);
    else console.log(`✅ Sent to ${to}`);
  } catch (err) {
    console.error("❌ Send error:", err.message);
  }
}

async function saveMsg(phone, role, content) {
  await pool.query(
    "INSERT INTO messages(phone,role,content)VALUES($1,$2,$3)",
    [phone, role, content]
  );
}

async function getHistory(phone) {
  const r = await pool.query(
    `SELECT role,content FROM(
      SELECT role,content,created_at FROM messages
      WHERE phone=$1 ORDER BY created_at DESC LIMIT 10
    )sub ORDER BY created_at ASC`,
    [phone]
  );
  return r.rows;
}

async function alertClient(session, phone) {
  const msg =
    `🔔 *NEW ENQUIRY — Classic Traders*\n\n` +
    `👤 Name: ${session.temp_name || "Unknown"}\n` +
    `📍 Location: ${session.temp_location || "Unknown"}\n` +
    `👥 Referred by: ${session.temp_referral || "Direct"}\n` +
    `📞 Phone: +${phone}\n` +
    `🕐 ${new Date().toLocaleString("en-IN")}\n\n` +
    `✅ End User price list shared automatically.\n` +
    `Please follow up! 📞`;
  await sendWA(process.env.CLIENT_PHONE, msg);
}

// ═══════════════════════════════════════
// KNOWN CUSTOMER
// ═══════════════════════════════════════
async function handleKnown(phone, text, customer) {
  const products = await getPrices();
  const priceText = formatPrices(products, customer.category);
  const history = await getHistory(phone);
  await saveMsg(phone, "user", text);

  const res = await groq.chat.completions.create({
    model: "llama-3.3-70b-versatile",
    max_tokens: 1024,
    messages: [
      {
        role: "system",
        content: `You are WhatsApp sales assistant for Classic Traders — CCTV & Security Systems dealer in Tamil Nadu.

CUSTOMER NAME: ${customer.name}
CUSTOMER TYPE: ${customer.category}

PRICE LIST FOR ${customer.category.toUpperCase()} ONLY:
${priceText}

STRICT RULES:
- Always greet customer by name: "${customer.name}"
- Show ONLY ${customer.category} prices — NEVER reveal other price levels exist
- Never mention Dealer/RP/LT/End User price differences
- For product not in list: say "Please call 7603868752"
- Shop timing: 9AM - 7PM, Monday to Saturday
- For quotation requests: show full price list neatly
- Reply in Tamil if customer writes in Tamil
- Reply in English if customer writes in English
- Be friendly and professional
- Keep replies short and clear for WhatsApp`
      },
      ...history,
      { role: "user", content: text }
    ]
  });

  const reply = res.choices[0].message.content;
  await saveMsg(phone, "assistant", reply);

  // Alert client on price enquiry
  const priceKeys = ["price", "rate", "list", "quotation", "quote", "விலை", "கொள்முதல்", "pricelist"];
  if (priceKeys.some(k => text.toLowerCase().includes(k))) {
    await sendWA(
      process.env.CLIENT_PHONE,
      `📊 *PRICE ENQUIRY ALERT*\n\n` +
      `👤 ${customer.name} (${customer.category})\n` +
      `📞 +${phone}\n` +
      `💬 "${text}"\n` +
      `🕐 ${new Date().toLocaleString("en-IN")}`
    );
  }

  return reply;
}

// ═══════════════════════════════════════
// UNKNOWN CUSTOMER
// ═══════════════════════════════════════
async function handleUnknown(phone, text) {
  let session = await getSession(phone);

  // Step 1 — Ask name
  if (!session || session.step === "start") {
    await setSession(phone, "ask_name");
    return (
      `🙏 Welcome to *Classic Traders!*\n\n` +
      `We specialize in:\n` +
      `📷 CCTV Cameras\n` +
      `📹 DVR / NVR Systems\n` +
      `🔒 Security Solutions\n\n` +
      `*Please share your name to continue.*`
    );
  }

  // Step 2 — Ask location
  if (session.step === "ask_name") {
    await setSession(phone, "ask_location", { name: text });
    return `Nice to meet you *${text}*! 😊\n\n📍 What is your *city / location*?`;
  }

  // Step 3 — Ask referral
  if (session.step === "ask_location") {
    await setSession(phone, "ask_referral", { location: text });
    return (
      `Got it! 📍\n\n` +
      `👥 *Who referred you* to Classic Traders?\n` +
      `_(Type name or "Direct" if no referral)_`
    );
  }

  // Step 4 — Ask requirement
  if (session.step === "ask_referral") {
    await setSession(phone, "ask_requirement", { referral: text });
    return (
      `Thank you! 🙏\n\n` +
      `🛒 What *product or service* are you looking for?\n` +
      `_(Example: 4 CCTV cameras, 8CH DVR, NVR, Installation etc.)_`
    );
  }

  // Step 5 — Save + Alert + Share price list
  if (session.step === "ask_requirement") {
    const s = await getSession(phone);

    // Save enquiry to database
    await pool.query(
      `INSERT INTO new_enquiries(phone,name,location,referred_by,requirement)
       VALUES($1,$2,$3,$4,$5)`,
      [phone, s.temp_name, s.temp_location, s.temp_referral, text]
    );

    // Alert client immediately
    await alertClient(s, phone);

    // Update session
    await setSession(phone, "done");

    // Get end user price list
    const products = await getPrices();
    const priceText = formatPrices(products, "End User");

    return (
      `Thank you *${s.temp_name}*! 🙏\n\n` +
      `Our team will contact you shortly.\n\n` +
      `Meanwhile here is our current price list:\n\n` +
      `${priceText}`
    );
  }

  // After done — answer questions with AI
  const products = await getPrices();
  const priceText = formatPrices(products, "End User");

  const res = await groq.chat.completions.create({
    model: "llama-3.3-70b-versatile",
    max_tokens: 512,
    messages: [
      {
        role: "system",
        content: `You are WhatsApp assistant for Classic Traders — CCTV & Security dealer.

END USER PRICE LIST:
${priceText}

RULES:
- Answer product questions helpfully
- For more info: "Call 7603868752"
- Timing: 9AM-7PM, Mon-Sat
- Reply Tamil if customer writes Tamil
- Keep replies short for WhatsApp`
      },
      { role: "user", content: text }
    ]
  });

  return res.choices[0].message.content;
}

// ═══════════════════════════════════════
// WEBHOOK
// ═══════════════════════════════════════
app.get("/webhook", (req, res) => {
  if (req.query["hub.verify_token"] === process.env.VERIFY_TOKEN) {
    console.log("✅ Webhook verified!");
    res.send(req.query["hub.challenge"]);
  } else {
    res.sendStatus(403);
  }
});

app.post("/webhook", async (req, res) => {
  res.sendStatus(200);
  try {
    const msg = req.body?.entry?.[0]?.changes?.[0]?.value?.messages?.[0];
    if (!msg || msg.type !== "text") return;

    const text = msg.text.body.trim();
    const from = msg.from;
    console.log(`📩 From ${from}: ${text}`);

    const customer = await getCustomer(from);
    const reply = customer
      ? await handleKnown(from, text, customer)
      : await handleUnknown(from, text);

    await sendWA(from, reply);
  } catch (err) {
    console.error("❌ Webhook error:", err.message);
  }
});

// ═══════════════════════════════════════
// ADMIN PANEL
// ═══════════════════════════════════════
const adminHTML = `<!DOCTYPE html>
<html>
<head>
<title>Classic Traders Admin</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#f0f2f5;padding:16px}
h1{color:#1a1a2e;margin-bottom:16px;font-size:20px}
h2{color:#333;font-size:15px;margin-bottom:10px}
.card{background:white;padding:16px;border-radius:12px;margin-bottom:14px;box-shadow:0 2px 8px rgba(0,0,0,0.08)}
label{font-size:12px;font-weight:bold;color:#555;display:block;margin-bottom:3px}
input,select{width:100%;padding:9px;margin-bottom:10px;border:1px solid #ddd;border-radius:8px;font-size:13px}
.btn{display:block;width:100%;background:#25D366;color:white;border:none;padding:11px;border-radius:8px;cursor:pointer;font-size:14px;margin-bottom:8px;text-align:center;text-decoration:none}
.btn:hover{background:#128C7E}
.btn-blue{background:#0084ff}.btn-blue:hover{background:#0066cc}
.btn-red{background:#dc3545}.btn-red:hover{background:#c82333}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.excel-box{background:#f8f9fa;border:1px solid #ddd;border-radius:8px;padding:10px;margin-bottom:10px;font-size:11px}
.excel-box table{width:100%;border-collapse:collapse}
.excel-box th,.excel-box td{border:1px solid #ccc;padding:4px 6px}
.excel-box th{background:#e9ecef}
@media(max-width:600px){.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<h1>🏪 Classic Traders — Admin Panel</h1>
<div class="grid">
<div>
<div class="card">
<h2>➕ Add / Update Customer</h2>
<form action="/admin/customer" method="POST">
<label>Phone (with 91, no +)</label>
<input name="phone" placeholder="917XXXXXXXXX" required>
<label>Customer Name</label>
<input name="name" placeholder="Raj Kumar" required>
<label>Category</label>
<select name="category">
<option>Dealer</option>
<option>RP/LT</option>
<option>End User</option>
</select>
<button class="btn" type="submit">💾 Save Customer</button>
</form>
</div>
<div class="card">
<h2>💰 Add Single Product</h2>
<form action="/admin/price" method="POST">
<label>Product Name</label>
<input name="product" placeholder="WiFi Dome Camera 2MP" required>
<label>Brand</label>
<input name="brand" placeholder="CP Plus">
<label>Category</label>
<input name="category" placeholder="Camera / DVR/NVR / Accessories">
<label>Dealer Price ₹</label>
<input name="dealer_price" type="number" placeholder="2720" required>
<label>RP/LT Price ₹</label>
<input name="rp_lt_price" type="number" placeholder="2560" required>
<label>End User Price ₹</label>
<input name="end_user_price" type="number" placeholder="2400" required>
<button class="btn" type="submit">💾 Save Product</button>
</form>
</div>
</div>
<div>
<div class="card">
<h2>📤 Upload Excel Price List</h2>
<div class="excel-box">
<b>Excel columns required:</b><br><br>
<table>
<tr><th>Product</th><th>Brand</th><th>Category</th><th>Dealer</th><th>RP/LT</th><th>EndUser</th></tr>
<tr><td>Dome 2MP</td><td>CP Plus</td><td>Camera</td><td>2720</td><td>2560</td><td>2400</td></tr>
</table>
</div>
<form action="/admin/upload" method="POST" enctype="multipart/form-data">
<label>Select Excel File (.xlsx)</label>
<input type="file" name="pricefile" accept=".xlsx,.xls" required>
<button class="btn btn-blue" type="submit">📊 Upload & Update All Prices</button>
</form>
</div>
<div class="card">
<h2>📋 View Data</h2>
<a class="btn" href="/admin/enquiries">🆕 New Enquiries</a>
<a class="btn" href="/admin/customers">👥 All Customers</a>
<a class="btn" href="/admin/prices">💰 Price List</a>
</div>
</div>
</div>
</body>
</html>`;

app.get("/admin", (req, res) => res.send(adminHTML));

app.post("/admin/customer", async (req, res) => {
  const { phone, name, category } = req.body;
  await pool.query(
    `INSERT INTO known_customers(phone,name,category)VALUES($1,$2,$3)
     ON CONFLICT(phone)DO UPDATE SET name=$2,category=$3`,
    [phone, name, category]
  );
  res.send(`<h2 style="font-family:Arial;padding:20px">✅ Customer "${name}" saved!</h2><a href="/admin" style="color:green">← Back to Admin</a>`);
});

app.post("/admin/price", async (req, res) => {
  const { product, brand, category, dealer_price, rp_lt_price, end_user_price } = req.body;
  await pool.query(
    `INSERT INTO price_list(product,brand,category,dealer_price,rp_lt_price,end_user_price,updated_at)
     VALUES($1,$2,$3,$4,$5,$6,NOW())`,
    [product, brand, category, dealer_price, rp_lt_price, end_user_price]
  );
  res.send(`<h2 style="font-family:Arial;padding:20px">✅ Product "${product}" saved!</h2><a href="/admin" style="color:green">← Back</a>`);
});

app.post("/admin/upload", upload.single("pricefile"), async (req, res) => {
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws);

    await pool.query("DELETE FROM price_list");

    let count = 0;
    for (const row of data) {
      const product = row["Product"] || row["product"] || row["PRODUCT"] || "";
      const brand = row["Brand"] || row["brand"] || row["BRAND"] || "";
      const category = row["Category"] || row["category"] || row["CATEGORY"] || "General";
      const dealer = parseFloat(row["Dealer"] || row["dealer"] || row["DEALER"] || 0);
      const rplt = parseFloat(row["RP/LT"] || row["RPLT"] || row["rp_lt"] || row["RP"] || 0);
      const enduser = parseFloat(row["EndUser"] || row["End User"] || row["end_user"] || row["ENDUSER"] || 0);

      if (product.trim()) {
        await pool.query(
          `INSERT INTO price_list(product,brand,category,dealer_price,rp_lt_price,end_user_price,updated_at)
           VALUES($1,$2,$3,$4,$5,$6,NOW())`,
          [product, brand, category, dealer, rplt, enduser]
        );
        count++;
      }
    }
    res.send(`<h2 style="font-family:Arial;padding:20px">✅ ${count} products uploaded! Prices updated instantly!</h2><a href="/admin" style="color:green">← Back</a>`);
  } catch (err) {
    res.send(`<h2 style="font-family:Arial;padding:20px;color:red">❌ Error: ${err.message}</h2><a href="/admin">← Back</a>`);
  }
});

app.get("/admin/enquiries", async (req, res) => {
  const r = await pool.query("SELECT * FROM new_enquiries ORDER BY created_at DESC LIMIT 100");
  let html = `<html><body style="font-family:Arial;padding:20px"><h1>🆕 New Enquiries</h1><a href="/admin">← Back</a><br><br>
  <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%;font-size:13px">
  <tr style="background:#f0f0f0"><th>Name</th><th>Phone</th><th>Location</th><th>Referred By</th><th>Requirement</th><th>Time</th></tr>`;
  r.rows.forEach(row => {
    html += `<tr>
      <td>${row.name || ""}</td><td>+${row.phone}</td>
      <td>${row.location || ""}</td><td>${row.referred_by || ""}</td>
      <td>${row.requirement || ""}</td>
      <td>${new Date(row.created_at).toLocaleString("en-IN")}</td>
    </tr>`;
  });
  html += `</table></body></html>`;
  res.send(html);
});

app.get("/admin/customers", async (req, res) => {
  const r = await pool.query("SELECT * FROM known_customers ORDER BY category,name");
  let html = `<html><body style="font-family:Arial;padding:20px"><h1>👥 Customers</h1><a href="/admin">← Back</a><br><br>
  <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%;font-size:13px">
  <tr style="background:#f0f0f0"><th>Name</th><th>Phone</th><th>Category</th></tr>`;
  r.rows.forEach(row => {
    html += `<tr><td>${row.name}</td><td>+${row.phone}</td><td>${row.category}</td></tr>`;
  });
  html += `</table></body></html>`;
  res.send(html);
});

app.get("/admin/prices", async (req, res) => {
  const r = await pool.query("SELECT * FROM price_list ORDER BY category,brand");
  let html = `<html><body style="font-family:Arial;padding:20px"><h1>💰 Price List</h1><a href="/admin">← Back</a><br><br>
  <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%;font-size:13px">
  <tr style="background:#f0f0f0"><th>Product</th><th>Brand</th><th>Category</th><th>Dealer ₹</th><th>RP/LT ₹</th><th>End User ₹</th><th>Updated</th></tr>`;
  r.rows.forEach(row => {
    html += `<tr>
      <td>${row.product}</td><td>${row.brand || ""}</td><td>${row.category || ""}</td>
      <td>₹${parseFloat(row.dealer_price || 0).toLocaleString("en-IN")}</td>
      <td>₹${parseFloat(row.rp_lt_price || 0).toLocaleString("en-IN")}</td>
      <td>₹${parseFloat(row.end_user_price || 0).toLocaleString("en-IN")}</td>
      <td>${new Date(row.updated_at).toLocaleDateString("en-IN")}</td>
    </tr>`;
  });
  html += `</table></body></html>`;
  res.send(html);
});

app.get("/", (req, res) => res.send("✅ Classic Traders Bot is running!"));

// Error handlers
process.on("uncaughtException", err => console.error("❌ Uncaught:", err.message));
process.on("unhandledRejection", err => console.error("❌ Rejection:", err.message));

// Start server
initDB().then(() => {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`🚀 Classic Traders Bot running on port ${PORT}!`);
    console.log(`🌐 Admin: http://localhost:${PORT}/admin`);
    console.log(`📡 Waiting for messages...`);
  });
}).catch(err => console.error("❌ DB Error:", err.message));
