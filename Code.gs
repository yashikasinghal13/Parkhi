/**
 * Parkhi back end: Google Sheet + Drive.
 * Every order, customer and review lands in this spreadsheet in real time;
 * review photos are saved to a Drive folder called "Parkhi review photos".
 *
 * Setup (once):
 *  1. Open your Google Sheet > Extensions > Apps Script. Delete what is there and paste this whole file.
 *  2. Save, then Deploy > New deployment > type "Web app".
 *     Execute as: Me.  Who has access: Anyone.  Click Deploy and allow access when Google asks.
 *  3. Copy the Web app URL (ends in /exec) into the website: SHOP.backendUrl in index.html.
 *  After any change to this file: Deploy > Manage deployments > Edit > Version: New version > Deploy
 *  (the URL stays the same).
 *
 * Hiding a review: in the "Reviews" tab, set the "Show" cell to No. It disappears from the site.
 * Forgotten password: in the "Customers" tab, type a temporary password in that customer's "Reset to" cell
 * and send it to them. It works once; they can then sign in and it becomes their password.
 */

// Prices the site charges, in rupees. Keep in step with the website.
const PRODUCTS = {
  aqua:    { name: 'Gulnaar',    price: 999 },
  indigo:  { name: 'Neel Kamal', price: 999 },
  ivory:   { name: 'Haveli',     price: 999 },
  mustard: { name: 'Sarson',     price: 999 },
};
const SIZES = ['XS', 'S', 'M', 'L', 'XL', 'XXL'];
const FITS = ['A', 'B', 'C', 'D'];
const UPI = { id: 'yashika.singhal@ybl', name: 'Parkhi' };
const PHOTO_FOLDER = 'Parkhi review photos';
const SESSION_DAYS = 60;

const SHEETS = {
  Orders:    ['Date', 'Order ID', 'Name', 'Mobile', 'Address', 'City', 'State', 'Pincode', 'Items', 'Total (₹)', 'Payment', 'Status', 'Customer ID', 'Lines'],
  Customers: ['Customer ID', 'Name', 'Mobile', 'Email', 'Address', 'City', 'State', 'Pincode', 'Joined', 'Password hash', 'Salt', 'Reset to', 'Wishlist'],
  Sessions:  ['Token', 'Customer ID', 'Expires'],
  Reviews:   ['Date', 'Review ID', 'Kurta', 'Product ID', 'Name', 'Rating', 'Size', 'Review', 'Photo 1', 'Photo 2', 'Photo 3', 'Show'],
};

/* ---------- entry points ---------- */
function doGet() { return json({ ok: true, shop: 'Parkhi' }); }

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Bad request.' }); }
  const reads = { listReviews: 1, me: 1, myOrders: 1 };
  const lock = LockService.getScriptLock();
  if (!reads[d.action] && !lock.tryLock(20000)) return json({ ok: false, error: 'Parkhi is busy. Please try again in a moment.' });
  try {
    const fn = ACTIONS[d.action];
    return json(fn ? fn(d) : { ok: false, error: 'Unknown action.' });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: 'Something went wrong on our side. Please try again.' });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

const ACTIONS = {
  /* ----- reviews ----- */
  listReviews() {
    const rows = rowsOf('Reviews');
    const reviews = rows.filter(r => String(r.Show).toLowerCase() !== 'no').map(r => ({
      id: r['Review ID'], product: r['Product ID'], name: r.Name, rating: Number(r.Rating) || 0, size: r.Size, text: r.Review,
      photos: [r['Photo 1'], r['Photo 2'], r['Photo 3']].filter(Boolean),
      createdAt: r.Date instanceof Date ? r.Date.toISOString() : String(r.Date),
    }));
    return { ok: true, reviews };
  },

  addReview(d) {
    const p = PRODUCTS[d.product];
    // rating, name, size, text and photos are all optional, but a review needs a rating, some text or a photo
    const rating = d.rating ? Math.round(Number(d.rating)) : 0;
    const name = clean(d.name, 40), text = clean(d.text, 800);
    const size = String(d.size || '');
    const photos = Array.isArray(d.photos) ? d.photos.slice(0, 3) : [];
    if (!p) return { ok: false, error: 'We could not find that kurta.' };
    if (rating && !(rating >= 1 && rating <= 5)) return { ok: false, error: 'Ratings go from 1 to 5 stars.' };
    if (!rating && !text && !photos.length) return { ok: false, error: 'Add a rating, a few words or a photo to post.' };
    if (size && !validSize(size)) return { ok: false, error: 'Please choose the size from the list.' };
    const id = 'R' + Date.now().toString(36).toUpperCase();
    const urls = [];
    try {
      photos.forEach((dataUrl, i) => urls.push(savePhoto(dataUrl, `${id}-${i + 1}`)));
    } catch (err) {
      return { ok: false, error: 'Your photo did not upload. Try a smaller photo, or post without it.' };
    }
    const now = new Date();
    sheet('Reviews').appendRow([now, id, p.name, d.product, cell(name), rating || '', size, cell(text), urls[0] || '', urls[1] || '', urls[2] || '', 'Yes']);
    return { ok: true, review: { id, product: d.product, name, rating, size, text, photos: urls, createdAt: now.toISOString() } };
  },

  /* ----- orders ----- */
  placeOrder(d) {
    const c = d.customer || {};
    const err = !clean(c.name, 80) ? 'Please enter your full name.' : !mobile(c.phone) ? 'Please enter a 10-digit Indian mobile number.' : addressError(c);
    if (err) return { ok: false, error: err };
    const items = Array.isArray(d.items) ? d.items : [];
    if (!items.length) return { ok: false, error: 'Your bag is empty.' };
    for (const l of items) {
      if (!PRODUCTS[l.id] || !validSize(String(l.size)) || !(Number(l.qty) >= 1 && Number(l.qty) <= 20)) return { ok: false, error: 'Something in your bag is not right. Please refresh and try again.' };
    }
    const orderId = /^PK-\d{6}-[A-Z0-9]{3,8}$/.test(String(d.orderId)) ? String(d.orderId) : 'PK-' + Date.now().toString(36).toUpperCase();
    const total = items.reduce((s, l) => s + PRODUCTS[l.id].price * Number(l.qty), 0);
    const u = d.token ? userFor(d.token) : null;
    const itemText = items.map(l => `${PRODUCTS[l.id].name} (${l.size}) x${Number(l.qty)}`).join(', ');
    const orders = sheet('Orders');
    const existing = findRow(orders, 'Order ID', orderId);
    const row = [new Date(), orderId, cell(clean(c.name, 80)), "'" + mobile(c.phone), cell(clean(c.addr, 300)), cell(clean(c.city, 60)), cell(clean(c.state, 60)), "'" + String(c.pin),
      itemText, total, d.paid ? 'Customer says paid' : 'Pay later', 'New', u ? u['Customer ID'] : '',
      JSON.stringify(items.map(l => ({ id: l.id, size: String(l.size), qty: Number(l.qty), price: PRODUCTS[l.id].price })))];
    if (existing) orders.getRange(existing, 1, 1, row.length).setValues([row]); else orders.appendRow(row);
    if (u && d.saveAddress) setAddress(u._row, c);
    return { ok: true, orderId, total, upiId: UPI.id,
      upiLink: `upi://pay?pa=${encodeURIComponent(UPI.id)}&pn=${encodeURIComponent(UPI.name)}&am=${total}.00&cu=INR&tn=${encodeURIComponent(orderId)}` };
  },

  myOrders(d) {
    const u = userFor(d.token);
    if (!u) return signin();
    // orders placed while signed in, plus guest orders from the same mobile number
    const orders = rowsOf('Orders').filter(o => o['Customer ID'] === u['Customer ID'] || String(o.Mobile) === String(u.Mobile)).reverse().map(o => ({
      orderId: o['Order ID'], total: Number(o['Total (₹)']) || 0, items: o.Items, payment: o.Payment, status: o.Status, lines: linesOf(o),
      date: o.Date instanceof Date ? Utilities.formatDate(o.Date, 'Asia/Kolkata', 'd MMM yyyy') : String(o.Date),
    }));
    return { ok: true, orders };
  },

  /* ----- accounts ----- */
  register(d) {
    const m = mobile(d.mobile), name = clean(d.name, 80);
    if (!name) return { ok: false, error: 'Please enter your name.' };
    if (!m) return { ok: false, error: 'Please enter a 10-digit Indian mobile number.' };
    if (String(d.password || '').length < 6) return { ok: false, error: 'Choose a password of at least 6 characters.' };
    const cs = sheet('Customers');
    if (findRow(cs, 'Mobile', m)) return { ok: false, error: 'An account with this mobile number already exists. Please sign in.' };
    const id = 'C' + Date.now().toString(36).toUpperCase(), salt = Utilities.getUuid();
    cs.appendRow([id, cell(name), "'" + m, cell(clean(d.email, 120)), '', '', '', '', new Date(), hash(d.password, salt), salt, '', '']);
    const u = rowsOf('Customers').find(r => r['Customer ID'] === id);
    return { ok: true, token: newSession(id), user: pub(u) };
  },

  login(d) {
    const m = mobile(d.mobile);
    const u = rowsOf('Customers').find(r => String(r.Mobile) === m);
    if (u && u['Reset to'] && String(d.password) === String(u['Reset to']) && String(d.password).length >= 6) {
      const salt = Utilities.getUuid(), cs = sheet('Customers');   // one-time reset set by Parkhi
      cs.getRange(u._row, 10, 1, 3).setValues([[hash(d.password, salt), salt, '']]);
      return { ok: true, token: newSession(u['Customer ID']), user: pub(u) };
    }
    if (!u || hash(d.password, u.Salt) !== u['Password hash']) return { ok: false, error: 'Mobile number or password is not right.' };
    return { ok: true, token: newSession(u['Customer ID']), user: pub(u) };
  },

  logout(d) {
    const s = sheet('Sessions'), r = findRow(s, 'Token', String(d.token || ''));
    if (r) s.deleteRow(r);
    return { ok: true };
  },

  me(d) { const u = userFor(d.token); return u ? { ok: true, user: pub(u) } : signin(); },

  updateProfile(d) {
    const u = userFor(d.token);
    if (!u) return signin();
    const name = clean(d.name, 80), m = mobile(d.mobile), email = clean(d.email, 120), np = String(d.newPassword || '');
    if (!name) return { ok: false, error: 'Please enter your name.' };
    if (!m) return { ok: false, error: 'Please enter a 10-digit Indian mobile number.' };
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Please check your email address.' };
    if (np && np.length < 6) return { ok: false, error: 'Choose a new password of at least 6 characters.' };
    const changingKey = m !== String(u.Mobile) || !!np;
    if (changingKey && hash(d.password, u.Salt) !== u['Password hash']) return { ok: false, error: 'Your current password is not right.' };
    const cs = sheet('Customers');
    if (m !== String(u.Mobile) && findRow(cs, 'Mobile', m)) return { ok: false, error: 'Another account already uses this mobile number.' };
    cs.getRange(u._row, 2, 1, 3).setValues([[cell(name), "'" + m, cell(email)]]);
    if (np) { const salt = Utilities.getUuid(); cs.getRange(u._row, 10, 1, 2).setValues([[hash(np, salt), salt]]); }
    return { ok: true, user: pub(userFor(d.token)) };
  },

  saveWishlist(d) {
    const u = userFor(d.token);
    if (!u) return signin();
    const list = (Array.isArray(d.wishlist) ? d.wishlist : []).filter(id => PRODUCTS[id]).slice(0, 50);
    sheet('Customers').getRange(u._row, 13).setValue(list.join(','));
    return { ok: true };
  },

  saveAddress(d) {
    const u = userFor(d.token);
    if (!u) return signin();
    const e = addressError(d.address || {});
    if (e) return { ok: false, error: e };
    setAddress(u._row, d.address);
    return { ok: true, user: pub(userFor(d.token)) };
  },
};

/* ---------- helpers ---------- */
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0009\u000b-\u001f]/g, ' ').trim().slice(0, max); }
// what goes into a cell: text that starts like a formula is kept as plain text
function cell(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function mobile(v) { let x = String(v || '').replace(/\D/g, ''); if (x.length > 10 && /^(91|0)/.test(x)) x = x.slice(-10); return /^[6-9]\d{9}$/.test(x) ? x : ''; }
function validSize(s) { const [b, f] = s.split('-'); return SIZES.includes(b) && (f === undefined || FITS.includes(f)); }
function addressError(a) {
  return !String(a.addr || '').trim() ? 'Please enter your delivery address.' : !String(a.city || '').trim() ? 'Please enter your city.'
    : !String(a.state || '').trim() ? 'Please enter your state.' : !/^[1-9]\d{5}$/.test(String(a.pin || '')) ? 'Please enter a 6-digit pincode.' : '';
}
function signin() { return { ok: false, code: 'signin', error: 'Please sign in again.' }; }
function hash(pw, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(pw), Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}

function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold');
  }
  // add any new columns to a sheet made by an older version of this script
  const want = SHEETS[name], have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  if (have.length < want.length || want.some((h, i) => have[i] !== h)) {
    want.forEach((h, i) => { if (have[i] !== h && !have[i]) sh.getRange(1, i + 1).setValue(h).setFontWeight('bold'); });
  }
  return sh;
}
function rowsOf(name) {
  const sh = sheet(name), v = sh.getDataRange().getValues(), head = v.shift() || [];
  return v.map((r, i) => { const o = { _row: i + 2 }; head.forEach((h, j) => o[h] = r[j]); return o; });
}
function findRow(sh, col, value) {
  const v = sh.getDataRange().getValues(), j = v[0].indexOf(col);
  for (let i = 1; i < v.length; i++) if (String(v[i][j]) === String(value)) return i + 1;
  return 0;
}

function newSession(id) {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  sheet('Sessions').appendRow([t, id, new Date(Date.now() + SESSION_DAYS * 864e5)]);
  return t;
}
function userFor(token) {
  if (!token) return null;
  const s = rowsOf('Sessions').find(r => r.Token === String(token));
  if (!s || new Date(s.Expires) < new Date()) return null;
  return rowsOf('Customers').find(r => r['Customer ID'] === s['Customer ID']) || null;
}
function pub(u) {
  return { id: u['Customer ID'], name: u.Name, mobile: String(u.Mobile), email: u.Email,
    wishlist: String(u.Wishlist || '').split(',').filter(id => PRODUCTS[id]),
    address: { addr: u.Address || '', city: u.City || '', state: u.State || '', pin: String(u.Pincode || '') } };
}
// an order's items as [{id, size, qty, price}]; older orders are read from the Items text
function linesOf(o) {
  try { const l = JSON.parse(o.Lines); if (Array.isArray(l)) return l; } catch (err) {}
  const byName = {}; Object.keys(PRODUCTS).forEach(id => byName[PRODUCTS[id].name] = id);
  return String(o.Items || '').split(', ').map(t => /^(.+) \((.+)\) x(\d+)$/.exec(t)).filter(m => m && byName[m[1]])
    .map(m => ({ id: byName[m[1]], size: m[2], qty: Number(m[3]), price: PRODUCTS[byName[m[1]]].price }));
}
function setAddress(row, a) {
  sheet('Customers').getRange(row, 5, 1, 4).setValues([[cell(clean(a.addr, 300)), cell(clean(a.city, 60)), cell(clean(a.state, 60)), "'" + String(a.pin)]]);
}

/* ---------- review photos ---------- */
function photoFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const f = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
function savePhoto(dataUrl, name) {
  const m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/.exec(String(dataUrl));
  if (!m) throw new Error('not an image');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 3 * 1024 * 1024) throw new Error('too big');
  const file = photoFolder().createFile(Utilities.newBlob(bytes, 'image/' + m[1], name + '.' + (m[1] === 'jpeg' ? 'jpg' : m[1])));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200';
}
