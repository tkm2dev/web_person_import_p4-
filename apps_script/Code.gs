/**
 * ระบบรวมข้อมูลบุคคล ภ.4 — ไฟล์เดียวจบ
 * วางไฟล์นี้ทั้งไฟล์ลงใน Apps Script (script.google.com > โปรเจกต์ใหม่ > แทนที่ Code.gs ทั้งหมด)
 *
 * ใช้กับไฟล์จังหวัด 12 ไฟล์ที่อัปโหลดเข้า Drive และแปลงเป็น Google Sheet ไว้แล้ว
 *
 * ลำดับการรัน (เลือกชื่อฟังก์ชันจาก dropdown ด้านบน แล้วกด "เรียกใช้")
 *
 *   0. setProvinceFolderId('<folder id>')  ครั้งเดียว เก็บ id โฟลเดอร์ลง Script Properties
 *   1. auditSharing()         ดูสิทธิ์ปัจจุบันของทุกไฟล์ — อ่านอย่างเดียว ไม่แก้อะไร
 *   2. verifyProvinceFiles()  ตรวจว่าครบ 12 ไฟล์ มี 8 แท็บ หัวตารางตรง template
 *   3. bindExisting()         สร้างไฟล์ ภ.4 + ติดตั้ง trigger + รวมข้อมูลรอบแรก
 *   4. hardenProvinceFiles()  dropdown / ไฮไลต์ช่องบังคับ / ล็อกหัวตาราง + ตั้งสิทธิ์ตาม SHARE_MODE
 *   5. shareProvinceLinks()   เฉพาะโหมด LINK_EDIT — ล็อกโฟลเดอร์ แล้วพิมพ์ลิงก์ 12 จังหวัดให้คัดลอกไปส่ง
 *
 * ข้อ 3 และ 4 ถ้าหมดเวลา 6 นาทีของ Apps Script ให้กด "เรียกใช้" ซ้ำ — จำความคืบหน้าไว้
 *
 * ไม่มี Drive id ใดๆ อยู่ในซอร์ส เพราะไฟล์นี้ขึ้น repo สาธารณะ — ทุก id อยู่ใน Script Properties
 *
 * เมนู "ภ.4" จะโผล่ในไฟล์ส่วนกลางหลังทำข้อ 3 เสร็จ
 * ตอนจะส่งส่วนกลางใช้ exportForCentral() — ลบแท็บช่วยงานออก เหลือโครงตรง template
 *
 * snapshotMaster() ทำงานอัตโนมัติทุกวันตี 1 สำรองไฟล์ ภ.4 ไว้ CONFIG.SNAPSHOT_KEEP ชุด
 * installOverview() สร้างแท็บ ภาพรวม (KPI + กราฟ) — bindExisting() เรียกให้อยู่แล้ว
 *                   ไฟล์ ภ.4 ที่สร้างไปก่อนหน้านี้ ให้รันฟังก์ชันนี้เองครั้งเดียว
 *
 * สำคัญ: โครงสร้างคอลัมน์ของทุกแท็บต้องตรงกับ person_import_template_p4.xlsx 100%
 *        สคริปต์อ่านหัวตารางจากไฟล์จังหวัดโดยตรง ไม่ hardcode และไม่เพิ่มคอลัมน์ใดๆ
 */

// ════════════════════════════════════════════════════════════════════
// 1. ตั้งค่า — แก้เฉพาะส่วนนี้
// ════════════════════════════════════════════════════════════════════

const CONFIG = {
  /**
   * โฟลเดอร์ที่เก็บไฟล์จังหวัด — ปล่อยว่างไว้ อย่าใส่ค่าจริงลงซอร์ส
   * รัน setProvinceFolderId('<folder id>') ครั้งเดียว ค่าจะถูกเก็บใน Script Properties
   * (ใส่ตรงนี้ก็ได้ถ้าซอร์สไม่เคยขึ้น repo)
   */
  PROVINCE_FOLDER_ID: '',

  ROOT_FOLDER_NAME: 'ข้อมูลบุคคล ภ.4',
  MASTER_NAME: 'ภ.4 — ศูนย์รวมข้อมูลบุคคล (ส่วนกลาง)',

  /** ความถี่ที่ ภ.4 ดึงข้อมูลจากจังหวัด (นาที) */
  SYNC_INTERVAL_MINUTES: 5,

  /**
   * วิธีให้สิทธิ์ไฟล์จังหวัด
   *
   *   'LINK_EDIT'  ใครถือลิงก์ไฟล์จังหวัดนั้น แก้ไขได้ทันที ไม่ต้องล็อกอิน ไม่ต้องเก็บอีเมล
   *                ต้องรัน shareProvinceLinks() เพื่อสร้างลิงก์และล็อกโฟลเดอร์
   *   'EMAIL_ONLY' ไฟล์เป็น PRIVATE แชร์เฉพาะอีเมลใน editors/viewers ข้างล่าง
   *
   * ไฟล์ ภ.4 ส่วนกลางเป็น PRIVATE เสมอไม่ว่าตั้งค่านี้เป็นอะไร เพราะรวมข้อมูลครบ 12 จังหวัด
   *
   * ข้อแลกเปลี่ยนของ LINK_EDIT ที่ต้องรับได้
   *   - ประวัติการแก้ไขขึ้นว่า "ผู้ใช้ที่ไม่ระบุชื่อ" สืบไม่ได้ว่าใครแก้หรือลบ
   *   - ใครถือลิงก์ลบแถวได้ และ syncAuto() จะเขียนทับไฟล์ ภ.4 ตามภายใน 5 นาที
   *     ตัวกันคือ snapshotMaster() ที่สำรองไฟล์ ภ.4 ทุกวัน (ดู SNAPSHOT_KEEP)
   *   - ใครถือลิงก์ ดาวน์โหลด/คัดลอกทั้งไฟล์ได้
   */
  SHARE_MODE: 'LINK_EDIT',

  /** เก็บไฟล์สำรองรายวันของไฟล์ ภ.4 ไว้กี่ชุด */
  SNAPSHOT_KEEP: 14,

  /**
   * editors/viewers — ใช้เฉพาะเมื่อ SHARE_MODE = 'EMAIL_ONLY'
   * (โหมด LINK_EDIT ก็ยังเพิ่มอีเมลที่ระบุไว้ได้ ถ้าอยากให้บางคนมีตัวตนในประวัติการแก้ไข)
   *
   * fileId ปล่อยว่างไว้โดยตั้งใจ — ห้าม commit ลง repo
   * ในโหมด LINK_EDIT ใครถือ spreadsheet id ก็เปิดแก้ข้อมูลได้ทันที มันคือรหัสผ่านดีๆ นี่เอง
   * สคริปต์หาไฟล์จาก PROVINCE_FOLDER_ID + ชื่อไฟล์แทน (โฟลเดอร์เป็น PRIVATE ID เปล่าๆ ใช้ไม่ได้)
   * ถ้าชื่อไฟล์ไม่ตรงชื่อจังหวัด ให้รัน setProvinceIds({...}) ครั้งเดียว ค่าจะถูกเก็บใน Script Properties
   */
  PROVINCES: [
    { name: 'กาฬสินธุ์',    editors: [], viewers: [], fileId: '' },
    { name: 'ขอนแก่น',      editors: [], viewers: [], fileId: '' },
    { name: 'นครพนม',       editors: [], viewers: [], fileId: '' },
    { name: 'บึงกาฬ',       editors: [], viewers: [], fileId: '' },
    { name: 'มหาสารคาม',    editors: [], viewers: [], fileId: '' },
    { name: 'มุกดาหาร',     editors: [], viewers: [], fileId: '' },
    { name: 'ร้อยเอ็ด',     editors: [], viewers: [], fileId: '' },
    { name: 'เลย',          editors: [], viewers: [], fileId: '' },
    { name: 'สกลนคร',       editors: [], viewers: [], fileId: '' },
    { name: 'หนองคาย',      editors: [], viewers: [], fileId: '' },
    { name: 'หนองบัวลำภู',  editors: [], viewers: [], fileId: '' },
    { name: 'อุดรธานี',     editors: [], viewers: [], fileId: '' },
  ],

  /** สิทธิ์บนไฟล์ ภ.4 ส่วนกลาง (เห็นข้อมูลทั้ง 12 จังหวัด) */
  MASTER_EDITORS: [],
  MASTER_VIEWERS: [],

  /** true = ยอมให้ resetAll() ย้ายไฟล์ ภ.4 ลงถังขยะ (ไม่แตะไฟล์จังหวัด) */
  ALLOW_RESET: false,
};

// ════════════════════════════════════════════════════════════════════
// 2. นิยามแท็บ — ต้องตรงกับ SHEETS ใน scripts/build_p4.py
// ════════════════════════════════════════════════════════════════════

const TABS = [
  { tab: 'กลุ่ม5สี11กลุ่ม',        type: 'กลุ่ม 5 สี 11 กลุ่ม' },
  { tab: 'ถวายฎีกา',               type: 'บุคคลถวายฎีกา' },
  { tab: 'ม112_มั่นคง',            type: 'บุคคล ม.112/คดีความมั่นคง' },
  { tab: 'จิตเวชรักษา',            type: 'บุคคลจิตเวชมีประวัติการรักษา' },
  { tab: 'เร่ร้อน',                type: 'บุคคลเร่ร้อน' },
  { tab: 'ร้องทุกข์_ดำรงธรรม',     type: 'ข้อมูลยื่นเรื่องราวร้องทุกข์ผ่านศูนย์ดำรงธรรม' },
  { tab: 'ร้องทุกข์_หน่วยงานอื่น', type: 'ข้อมูลยื่นเรื่องราวร้องทุกข์ หน่วยงานอื่นๆ' },
  { tab: 'เฝ้าระวัง_ทะเลาะวิวาท',  type: 'กลุ่มบุคคลเฝ้าระวัง (ทะเลาะวิวาท)' },
];

const COL_TYPE = 'ประเภทข้อมูล*';
const COL_PROVINCE = 'จังหวัด*';

/**
 * สองคอลัมน์นี้ถูกละเว้นตอนตัดสินว่าแถวว่างหรือไม่
 * เพราะแถวสำรองท้ายตารางถูกเติมค่าไว้ให้เจ้าหน้าที่แล้ว
 */
const IGNORE_WHEN_BLANK_CHECK = [COL_TYPE, COL_PROVINCE];

const README_TAB = 'README';
const LISTS_TAB = 'Lists';
const REGISTRY_TAB = '_ทะเบียนไฟล์';
const OVERVIEW_TAB = 'ภาพรวม';
const CHARTDATA_TAB = '_ข้อมูลกราฟ';
const HISTORY_TAB = '_ประวัติรายวัน';

/** ของเก่า ถูกยุบเข้า ภาพรวม แล้ว — คงชื่อไว้เพื่อลบทิ้งจากไฟล์ที่สร้างก่อนหน้านี้ */
const LEGACY_DASHBOARD_TAB = 'Dashboard';

/**
 * แท็บที่ ภ.4 ใช้เอง — exportForCentral() ลบทิ้งก่อนส่งต่อ
 * ต้องลบ ภาพรวม ก่อน _ข้อมูลกราฟ เพราะกราฟอ้างช่วงข้อมูลในนั้น
 */
const HELPER_TABS = [OVERVIEW_TAB, CHARTDATA_TAB, HISTORY_TAB, REGISTRY_TAB, LEGACY_DASHBOARD_TAB];

/**
 * คอลัมน์ที่ใช้วัด "ความครบถ้วน" ในแท็บ ภาพรวม
 * ชื่อ มีทั้ง 'ชื่อ*' และ 'ชื่อ' (แท็บ จิตเวชรักษา ไม่มี *) จึงต้องลองหลายชื่อ
 */
const COMPLETENESS_COLS = [
  { label: 'อำเภอ/เขต', names: ['อำเภอ/เขต*'] },
  { label: 'หน่วยงานเจ้าของข้อมูล', names: ['หน่วยงานเจ้าของข้อมูล*'] },
  { label: 'วันที่บันทึก', names: ['วันที่บันทึก*'] },
  { label: 'ชื่อ', names: ['ชื่อ*', 'ชื่อ'] },
  { label: 'นามสกุล', names: ['นามสกุล'] },
  { label: 'เลขบัตรประชาชน', names: ['เลขบัตรประชาชน'] },
];

/** ข้อความที่ build_p4.py ต่อท้าย README ของไฟล์จังหวัด — ตัดออกจากไฟล์ ภ.4 */
const README_PROVINCE_MARKER = '— ฉบับแยกจังหวัด (ภ.4) —';

const DATE_COLS = ['วันที่บันทึก*', 'วันเกิด', 'วันที่ติดตามถัดไป'];

const TEXT_COLS = ['เลขบัตรประชาชน', 'เลขเอกสารอื่น', 'โทรศัพท์',
                   'เลขที่อ้างอิงต้นทาง', 'เลขคำร้อง/เลขฎีกา'];

/** หัวคอลัมน์ในแท็บข้อมูล -> หัวรายการในแท็บ Lists (ค่าทั้งหมดอ่านจาก Lists ไม่คิดเอง) */
const DROPDOWN_SOURCE = {
  'คำนำหน้า': 'คำนำหน้า',
  'เพศ': 'เพศ',
  'ช่องทางรับเรื่อง': 'ช่องทางรับเรื่อง',
};

const HEADER_BG = '#1f3864';
const HEADER_BG_REQUIRED = '#7b2d26';
const MISSING_REQUIRED_BG = '#fce8e6';

/** ช่องแสดงผลการดึงข้อมูล ถัดจากปุ่มบนแท็บ ภาพรวม */
const SYNC_STATUS_CELL = 'E2';

/**
 * ปุ่ม "ดึงข้อมูลล่าสุด" — รูป PNG ที่ผูกกับ syncFromButton() ผ่าน OverGridImage.assignScript()
 * เป็นปุ่มกดจริง ไม่ใช่ checkbox และไม่ต้องพึ่ง onEdit trigger
 *
 * สร้างจาก scripts/make_button.py (Pillow + ฟอนต์ Leelawadee UI) ขนาด 520x88 ย่อแสดงที่ 260x44
 */
const SYNC_BUTTON_PNG_B64 = [
  'iVBORw0KGgoAAAANSUhEUgAAAggAAABYCAMAAACnM/+PAAAAP1BMVEXN09wjO2ZecI+stcWYpLh3h6FDWH27wtA9U3kfOGT9',
  '/v4AAAAwR29QZIbn6u6Kl64AAAAAAAAAAAAAAAAAAAABW77SAAAAFXRSTlP//////////////wD////////////nnEVzAAAG',
  'QElEQVR42u2c6XbrKgxGGTxShvd/28sMtmUnzW1We+JPP85KqcBU2mBJkMO+mqyQm0nnfAYKwEIPAixycxIYMAAKFQTYAiQw',
  'cABZAQKkAwF2gAQQYAWIJwEgQAACpAMBNoAEAQgQgAABCBCAAAEIEIAAAQgQgAABCBCAAAEIEIAAuT0IWio2XSkYptjAwyen',
  'gP7ngiCVUpZfgKKkZmoMHwcl4fSPA8FIq/xSN4r5PcGc6wklPAHxT2XKvWEi7i2jAoTnhPvVbbgWTPnFbryvH+wI1n+a1JUe',
  'QPgbopL0LVfqgxoSEFLNelD6QtXDIpl/J3gc3vIHA4Qfp2DDwpaKw37PcljA/Z7AzPXwi2STcX7ACSD8ZRDUXmrrRYTo2tb/',
  'cMNnaVj2Fg4Awts4iACoKxC4Db+1ls3SCf9yeAIEJgVfAcLfBUFdyikJnGu9TEaMbpiZ5f+Asf2raQYIL3Kg1Afl6u/JWT4O',
  'hMt4gSojDMy/H+x+x/dppU8Xyw/jHHSKinfFkn8hYm8mx9Z5tF3Q6VXL58mlB7lNZnJs5WlGu3dQ1ROqxrfns78vCIS7nwJB',
  'sKZhzSaKVNXiS1ViywaE0bbOZZWO8UexB2Ga24OG6jSiVbQx8+N2eg2t89nfFgTS20+QYIK1x0VzLYKpx656ZIVmyZ/au8ZN',
  'XBuvYnUHgoy9fWcjW2erhJa5ONFACA+yg9F8GVlb0lRrGEoanWpcVu/0WF8HP5397UG4jhuO3aKL124pL40QGVpcds1UtwnX',
  'gcDaO8CUhcpD7dHUcK6A4Ad3pWDhRxnqIw+tHiK9Vyx6oYXV3uezvy0IZ55+FCi6atQgISrY7QguujYsPJFLDawHQXaRQYnl',
  '045wAMGKTc5aXi3HVjd125JSvNMLx58muF8/mP3NQTg2PxGAd6tI96ePMUaIIIg0uk4fg+9LN9cFaDyr5BhhD4KYKP7o1mOC',
  'QOtdzB4gfLvfzu66zxoSCH7XT299znJExqhdeC7vhhhBzsesYROZyPWZ1oHKFKvexexvCsKrHPDUS5dMkW0PEVwCYQqvgxAz',
  'FFfTIMjmNEOCsAws54lCtVMrurWBMJ7qPZg9QPheR10yxeq1YQcCz3mAqY6qIOjBqlk/B4Jred7UXE63HkGg9MjZA4RXQIiL',
  'aFF28nGhKaasaX4GITvBs1C8UkHoiw0PQPCD2HHhenKeq+pyuvUIAqlHzh4gvAJCNLQIsZcLFs8gqJwMFBBWv+591tC23gKC',
  'j89MKTY8AEHXmsCqYy1yPW89gkDrkbMHCK+AELIA/09cU9WUrJDg1OY4sKZtFYRYLsjFhgcgjGpz3p1dSbceQaD1yNkDhFdA',
  'SHWBVKediim9WdXMSRD4HoS4I4xbEFLdd06HB8U7m/h/6NY00XoEgdYjZw8QXjpfjM4Nb/+4hLMpY/1OPwfCIUbQrfwvl+ad',
  'eEW61B3qW55uPYJwokfOHnWEF6PFJdVqdWfKUEPwJDwEIZaadlmD7zoEV4XDA9+9B6EOtwHh2EqDQOjRswcIL3SUwXhLMeEm',
  '12P6FIQSLejedTmWDE0pRgjHiEMDYdamDDA1EKhWCgRS72z2dwdBvQSCGOsRbmfKMXzV4QQE7xK7cX63TnkI4HLtL78lcowg',
  'grreupJupWIEUu909rcF4ez08RkQVAoI9qYU/W7cg6BFO/QLQYIT4Rx6KB73maYuBCUSMgjUmh6e3BGG0x3hZPZ3B+FwH+Ep',
  'EOo9kY0pDQFCDQNLD3G4mZJLgKngFK7I5zqC7YarrqRbiToCrXcx+7uCcHpD6TEI3XcZtqacLA2CnbuDID3KeFdEjuUSgQvZ',
  'p+iS/nxDibW4v7mSbj1WFmm9q9nfFYTzO4sr5J4gfO71ZYDwf0j4pdmE70nmqyw/PzKbAMI3UfitqXA1LuV600/betCMAYTv',
  'kbD+KghzrEfYN4BgAcJ3UPjNiYS8cl7f8Z/sGKYUQHi5tvQLYtNJ0Bu+bTADhH9JPAh66L+eBhBuCkIo/7p3XCwGCBDUESAA',
  'AQIQIAABAhAgAAECECAAAQIQIAABAhAgAAECECAfD8IXbABZ1y+AAAEIkB4EkABZvwIIIAEcAARIBwJIAAcJBJBwew4yCEDh',
  '5hg0EEDCvTloIICFG1Pg5T+wz0Q38uCNngAAAABJRU5ErkJggg==',
].join('');

// ════════════════════════════════════════════════════════════════════
// 3. ตัวช่วย
// ════════════════════════════════════════════════════════════════════

const P = PropertiesService.getScriptProperties();
const TZ = 'Asia/Bangkok';
const CHUNK = 5000;                     // แถวต่อการเขียน 1 ครั้ง
const TIME_BUDGET_MS = 4.5 * 60 * 1000; // Apps Script ตัดที่ 6 นาที เผื่อไว้ 1.5 นาที

const K_MASTER = 'MASTER_ID';
const K_PROVS = 'PROV_IDS';
const K_STAMPS = 'LAST_SYNC_STAMPS';
const K_SYNC_AT = 'LAST_SYNC_AT';
const K_HARDENED = 'HARDENED';
/** { ชื่อแท็บ: { ชื่อจังหวัด: จำนวนแถว } } — บอกว่าแต่ละจังหวัดกินแถวไหนในชีตส่วนกลาง */
const K_COUNTS = 'TAB_COUNTS';
/** เฉพาะไฟล์ที่สคริปต์สร้างเอง — resetAll() ลบแค่ไฟล์ในรายการนี้ */
const K_CREATED = 'CREATED_IDS';
/** { ชื่อจังหวัด: spreadsheetId } เก็บนอกซอร์สโค้ด เพราะ id = สิทธิ์เข้าถึงในโหมด LINK_EDIT */
const K_FILE_IDS = 'PROVINCE_FILE_IDS';
/** id โฟลเดอร์ไฟล์จังหวัด — เก็บนอกซอร์สเช่นกัน ซอร์สอาจขึ้น repo สาธารณะ */
const K_FOLDER_ID = 'PROVINCE_FOLDER_ID';

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function propGet_(key, fallback) {
  const v = P.getProperty(key);
  return v === null ? fallback : v;
}

function propJson_(key, fallback) {
  const v = P.getProperty(key);
  if (!v) return fallback;
  try { return JSON.parse(v); } catch (e) { return fallback; }
}

function propSetJson_(key, value) {
  P.setProperty(key, JSON.stringify(value));
}

function markCreated_(id) {
  const list = propJson_(K_CREATED, []);
  if (list.indexOf(id) === -1) {
    list.push(id);
    propSetJson_(K_CREATED, list);
  }
}

function folder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function rootFolder_() {
  return folder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER_NAME);
}

function isBlank_(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

/**
 * แถวมีข้อมูลจริงหรือไม่ — ไม่นับคอลัมน์ ประเภทข้อมูล และ จังหวัด
 * (กฎเดียวกับที่ build_p4.py ใช้ตัดแถวเปล่า 1,519 แถวออกจากไฟล์ต้นฉบับ)
 */
function rowHasData_(row, headers) {
  for (let i = 0; i < headers.length; i++) {
    if (IGNORE_WHEN_BLANK_CHECK.indexOf(headers[i]) !== -1) continue;
    if (!isBlank_(row[i])) return true;
  }
  return false;
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function trimSize_(sh, rows, cols) {
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
}

function writeRows_(sh, startRow, values) {
  if (!values.length) return;
  const cols = values[0].length;
  ensureSize_(sh, startRow + values.length - 1, cols);
  for (let i = 0; i < values.length; i += CHUNK) {
    const block = values.slice(i, i + CHUNK);
    sh.getRange(startRow + i, 1, block.length, cols).setValues(block);
  }
}

function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

function sheetByName_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบแท็บ "' + name + '" ในไฟล์ ' + ss.getName());
  return sh;
}

function clearData_(sh) {
  const r = sh.getMaxRows(), c = sh.getMaxColumns();
  if (r > 1) sh.getRange(2, 1, r - 1, c).clearContent();
}

function styleHeader_(sh, headers, protect) {
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true);
  headers.forEach(function (h, i) {
    sh.getRange(1, i + 1).setBackground(String(h).slice(-1) === '*' ? HEADER_BG_REQUIRED : HEADER_BG);
  });
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 42);

  if (protect) {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
      if (p.getDescription() === 'หัวตาราง') p.remove();
    });
    const prot = sh.getRange(1, 1, 1, sh.getMaxColumns()).protect().setDescription('หัวตาราง');
    prot.removeEditors(prot.getEditors());
    if (prot.canDomainEdit()) prot.setDomainEdit(false);
  }
}

/**
 * อ่านทั้งแท็บ ตัดแถวเปล่าออก และปรับค่าให้ปลอดภัย
 *   Date -> ข้อความ yyyy-MM-dd  (กัน locale สลับ วัน/เดือน)
 *   ตัวเลขในคอลัมน์เลขบัตร/เลขเอกสาร -> ข้อความ (กัน 1.47020121503E+12)
 */
function readTab_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) return { headers: [], rows: [] };
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (!lastCol) return { headers: [], rows: [] };
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (lastRow < 2) return { headers: headers, rows: [] };

  const raw = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const isText = headers.map(function (h) { return TEXT_COLS.indexOf(h) !== -1; });
  const rows = [];
  for (let r = 0; r < raw.length; r++) {
    const row = raw[r];
    if (!rowHasData_(row, headers)) continue;
    for (let i = 0; i < row.length; i++) {
      const v = row[i];
      if (v instanceof Date) row[i] = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
      else if (isText[i] && typeof v === 'number') row[i] = String(v);
    }
    rows.push(row);
  }
  return { headers: headers, rows: rows };
}

/** อ่านเฉพาะแถวหัวตารางของทั้ง 8 แท็บ (ไม่แตะข้อมูล จึงเร็ว) */
function readHeaders_(ss) {
  const out = {};
  TABS.forEach(function (t) {
    const sh = ss.getSheetByName(t.tab);
    if (!sh) throw new Error('ไฟล์ "' + ss.getName() + '" ไม่มีแท็บ "' + t.tab + '"');
    const lastCol = sh.getLastColumn();
    if (!lastCol) throw new Error('แท็บ "' + t.tab + '" ในไฟล์ "' + ss.getName() + '" ว่างเปล่า');
    out[t.tab] = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  });
  return out;
}

/** อ่านแท็บ Lists -> { หัวรายการ: [ค่า, ...] } */
function readLists_(ss) {
  const sh = sheetByName_(ss, LISTS_TAB);
  const grid = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const out = {};
  for (let c = 0; c < grid[0].length; c++) {
    const head = String(grid[0][c] || '').trim();
    if (!head) continue;
    const vals = [];
    for (let r = 1; r < grid.length; r++) {
      if (!isBlank_(grid[r][c])) vals.push(String(grid[r][c]).trim());
    }
    if (vals.length) out[head] = vals;
  }
  return out;
}

/** แชร์เฉพาะอีเมลที่ระบุ ปิดลิงก์สาธารณะ — ใช้กับไฟล์ ภ.4 เสมอ */
function applySharing_(fileId, editors, viewers) {
  const file = DriveApp.getFileById(fileId);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  addPeople_(file, editors, viewers);
}

/** สิทธิ์ไฟล์จังหวัด ตาม CONFIG.SHARE_MODE */
function shareProvinceFile_(fileId, cfg) {
  const file = DriveApp.getFileById(fileId);
  if (CONFIG.SHARE_MODE === 'LINK_EDIT') {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } else {
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  }
  addPeople_(file, cfg.editors, cfg.viewers);
}

function addPeople_(file, editors, viewers) {
  (editors || []).forEach(function (e) {
    if (e) try { file.addEditor(e); } catch (err) { Logger.log('addEditor ' + e + ': ' + err); }
  });
  (viewers || []).forEach(function (e) {
    if (e) try { file.addViewer(e); } catch (err) { Logger.log('addViewer ' + e + ': ' + err); }
  });
}

/** ล็อกทั้งชีต ให้เหลือเจ้าของแก้ได้คนเดียว (ใช้กับ README และ Lists) */
function protectSheet_(sh, desc) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (p.getDescription() === desc) p.remove();
  });
  const prot = sh.protect().setDescription(desc);
  prot.removeEditors(prot.getEditors());
  if (prot.canDomainEdit()) prot.setDomainEdit(false);
}

function provinceCfg_(name) {
  const c = CONFIG.PROVINCES.filter(function (p) { return p.name === name; })[0];
  if (!c) throw new Error('ไม่รู้จักจังหวัด ' + name);
  return c;
}

function masterUrl_() {
  const id = propGet_(K_MASTER, '');
  return id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '(ยังไม่ได้สร้าง)';
}

// ════════════════════════════════════════════════════════════════════
// 4. ค้นหาและตรวจไฟล์จังหวัด
// ════════════════════════════════════════════════════════════════════

/** id โฟลเดอร์ไฟล์จังหวัด — จาก CONFIG ก่อน ถ้าว่างค่อยดู Script Properties */
function folderId_() {
  return CONFIG.PROVINCE_FOLDER_ID || propGet_(K_FOLDER_ID, '');
}

/**
 * เก็บ id โฟลเดอร์ลง Script Properties รันครั้งเดียวหลังวางสคริปต์
 *
 *   setProvinceFolderId('14wXv...')
 */
function setProvinceFolderId(id) {
  const v = String(id || '').trim();
  if (!v) throw new Error('ต้องระบุ id โฟลเดอร์');
  DriveApp.getFolderById(v);   // โยน error ทันทีถ้า id ผิดหรือไม่มีสิทธิ์
  P.setProperty(K_FOLDER_ID, v);
  Logger.log('เก็บ id โฟลเดอร์แล้ว');
}

/**
 * เก็บ spreadsheet id ของจังหวัดลง Script Properties (ไม่ผ่านซอร์สโค้ด)
 * ใช้เมื่อชื่อไฟล์ใน Drive ไม่ตรงชื่อจังหวัด จนค้นจากโฟลเดอร์ไม่เจอ
 *
 *   setProvinceIds({ 'ขอนแก่น': '<spreadsheet id>', 'เลย': '<spreadsheet id>' })
 */
function setProvinceIds(map) {
  const cur = propJson_(K_FILE_IDS, {});
  Object.keys(map || {}).forEach(function (n) {
    provinceCfg_(n);            // โยน error ถ้าชื่อจังหวัดผิด
    cur[n] = map[n];
  });
  propSetJson_(K_FILE_IDS, cur);
  Logger.log('เก็บ id ของ ' + Object.keys(cur).length + ' จังหวัดแล้ว');
}

/** ลบ id ที่เก็บไว้ทั้งหมด */
function clearProvinceIds() {
  P.deleteProperty(K_FILE_IDS);
  Logger.log('ล้าง id ที่เก็บไว้แล้ว');
}

/**
 * คืน { ชื่อจังหวัด: spreadsheetId }
 * ลำดับความสำคัญ: CONFIG.fileId -> Script Properties -> ค้นจาก PROVINCE_FOLDER_ID ด้วยชื่อไฟล์
 */
function resolveProvinceFiles_() {
  const out = {};
  const need = [];
  const stored = propJson_(K_FILE_IDS, {});
  CONFIG.PROVINCES.forEach(function (p) {
    const id = p.fileId || stored[p.name];
    if (id) out[p.name] = id;
    else need.push(p.name);
  });

  if (need.length) {
    const fid = folderId_();
    if (!fid) {
      throw new Error('ยังไม่ได้ตั้งโฟลเดอร์ — รัน setProvinceFolderId(\'<folder id>\') ก่อน ' +
                      '(จังหวัดที่ยังหาไม่เจอ: ' + need.join(', ') + ')');
    }
    const folder = DriveApp.getFolderById(fid);
    const sheets = [];
    const it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) {
      const f = it.next();
      sheets.push({ id: f.getId(), name: f.getName() });
    }
    if (!sheets.length) {
      const xl = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
      if (xl.hasNext()) {
        throw new Error('โฟลเดอร์นี้มีแต่ .xlsx ที่ยังไม่ได้แปลง — เปิดแต่ละไฟล์แล้วสั่ง ' +
                        '"ไฟล์ > บันทึกเป็น Google ชีต" ก่อน');
      }
      throw new Error('ไม่พบ Google Sheet ในโฟลเดอร์ ' + folder.getName());
    }

    need.forEach(function (name) {
      let hit = sheets.filter(function (s) { return s.name === name; });
      if (!hit.length) hit = sheets.filter(function (s) { return s.name.indexOf(name) !== -1; });
      if (!hit.length) {
        throw new Error('ไม่พบไฟล์ของจังหวัด ' + name + ' — เปลี่ยนชื่อไฟล์ให้มีคำว่า "' + name +
                        '" หรือระบุ fileId ใน CONFIG');
      }
      if (hit.length > 1) {
        throw new Error('จังหวัด ' + name + ' ตรงกับหลายไฟล์: ' +
                        hit.map(function (h) { return h.name; }).join(', ') + ' — ระบุ fileId ให้ชัดเจน');
      }
      out[name] = hit[0].id;
    });
  }

  const seen = {};
  Object.keys(out).forEach(function (name) {
    if (seen[out[name]]) throw new Error('ไฟล์เดียวกันถูกจับคู่กับทั้ง ' + seen[out[name]] + ' และ ' + name);
    seen[out[name]] = name;
  });
  return out;
}

/** ไฟล์จังหวัดไฟล์แรก ใช้เป็นแม่แบบของหัวตาราง / README / Lists */
function templateSpreadsheet_(ids) {
  const first = CONFIG.PROVINCES.filter(function (p) { return ids[p.name]; })[0];
  if (!first) throw new Error('ไม่มีไฟล์จังหวัดสักไฟล์');
  return SpreadsheetApp.openById(ids[first.name]);
}

/**
 * เปิดไฟล์จังหวัดทีละไฟล์ เทียบหัวตารางกับแม่แบบ และนับแถว
 * นับจาก getLastRow() ไม่อ่านข้อมูลจริง — เร็วพอที่จะไม่ชนลิมิต 6 นาที
 * คืน { problems: [], counts: { จังหวัด: จำนวนแถวในชีต } }
 */
function checkProvinceFiles_(ids, ref) {
  const problems = [];
  const counts = {};

  Object.keys(ids).forEach(function (name) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(ids[name]);
    } catch (e) {
      problems.push(name + ': เปิดไฟล์ไม่ได้ (' + e.message + ')');
      return;
    }
    let rows = 0;
    TABS.forEach(function (t) {
      const sh = ss.getSheetByName(t.tab);
      if (!sh) { problems.push(name + ': ไม่มีแท็บ "' + t.tab + '"'); return; }

      const got = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const want = ref[t.tab];
      if (got.length !== want.length) {
        problems.push(name + '/' + t.tab + ': คอลัมน์ ' + got.length + ' ควรเป็น ' + want.length);
        return;
      }
      for (let i = 0; i < want.length; i++) {
        if (got[i] !== want[i]) {
          problems.push(name + '/' + t.tab + ' คอลัมน์ที่ ' + (i + 1) +
                        ': พบ "' + got[i] + '" ควรเป็น "' + want[i] + '"');
          return;
        }
      }
      rows += Math.max(sh.getLastRow() - 1, 0);
    });
    counts[name] = rows;
  });
  return { problems: problems, counts: counts };
}

/** ตรวจอย่างเดียว ไม่แก้ไขอะไร */
function verifyProvinceFiles() {
  const ids = resolveProvinceFiles_();
  Logger.log('พบไฟล์ ' + Object.keys(ids).length + ' / ' + CONFIG.PROVINCES.length + ' จังหวัด');

  const ref = readHeaders_(templateSpreadsheet_(ids));
  const res = checkProvinceFiles_(ids, ref);

  Object.keys(res.counts).forEach(function (n) {
    Logger.log('  ' + n + ' — ' + res.counts[n] + ' แถวในชีต (รวมแถวสำรองท้ายตาราง)');
  });

  if (!res.problems.length) {
    Logger.log('ผ่าน — ทุกไฟล์มีครบ 8 แท็บ และหัวตารางตรง template');
  } else {
    res.problems.forEach(function (p) { Logger.log('ผิดพลาด: ' + p); });
    throw new Error('พบปัญหา ' + res.problems.length + ' รายการ — แก้ก่อนรัน bindExisting()');
  }
}

// ════════════════════════════════════════════════════════════════════
// 5. สิทธิ์การเข้าถึง
// ════════════════════════════════════════════════════════════════════

/**
 * รายงานสิทธิ์ปัจจุบัน อ่านอย่างเดียว
 * เตือนเฉพาะเมื่อสิทธิ์ไม่ตรงกับที่ตั้งใจ ไม่ใช่ทุกครั้งที่ไฟล์เปิดสาธารณะ
 *   ไฟล์จังหวัด  ควรตรงกับ CONFIG.SHARE_MODE
 *   ไฟล์ ภ.4     ต้องเป็น PRIVATE เสมอ
 */
function auditSharing() {
  const ids = resolveProvinceFiles_();
  const masterId = propGet_(K_MASTER, '');
  const wantProvince = CONFIG.SHARE_MODE === 'LINK_EDIT' ? 'ANYONE_WITH_LINK' : 'PRIVATE';
  let bad = 0;

  const report = function (name, id, want) {
    try {
      const f = DriveApp.getFileById(id);
      const access = String(f.getSharingAccess());
      const editors = f.getEditors().map(function (u) { return u.getEmail(); });
      const viewers = f.getViewers().map(function (u) { return u.getEmail(); });
      const ok = access === want;
      if (!ok) bad++;
      Logger.log((ok ? '' : '*** ผิดจากที่ตั้งใจ (ควรเป็น ' + want + ') *** ') + name +
                 ' | access=' + access + ' permission=' + f.getSharingPermission() +
                 ' | editors=[' + editors.join(', ') + '] viewers=[' + viewers.join(', ') + ']');
    } catch (e) {
      bad++;
      Logger.log(name + ': อ่านสิทธิ์ไม่ได้ (' + e.message + ')');
    }
  };

  CONFIG.PROVINCES.forEach(function (p) {
    if (ids[p.name]) report(p.name, ids[p.name], wantProvince);
  });
  if (masterId) report('[ภ.4 ส่วนกลาง]', masterId, 'PRIVATE');

  if (folderId_()) {
    try {
      const folder = DriveApp.getFolderById(folderId_());
      const access = String(folder.getSharingAccess());
      const ok = access === 'PRIVATE';
      if (!ok) bad++;
      Logger.log((ok ? '' : '*** โฟลเดอร์เปิดสาธารณะ — เห็นครบ 12 จังหวัด รัน lockFolder() *** ') +
                 '[โฟลเดอร์] | access=' + access);
    } catch (e) {
      Logger.log('[โฟลเดอร์]: อ่านสิทธิ์ไม่ได้ (' + e.message + ')');
    }
  }

  Logger.log(bad ? '--- พบ ' + bad + ' รายการที่ต้องแก้ ---'
                 : '--- สิทธิ์ถูกต้องทั้งหมด (โหมด ' + CONFIG.SHARE_MODE + ') ---');
}

/**
 * ปิดสิทธิ์ของ "โฟลเดอร์" ที่เก็บไฟล์จังหวัด
 *
 * สำคัญที่สุดในโหมด LINK_EDIT — ถ้าโฟลเดอร์ยังเปิดสาธารณะ
 * คนที่ถือลิงก์โฟลเดอร์จะเห็นไฟล์ครบทั้ง 12 จังหวัด การแยกไฟล์เป็นโมฆะทันที
 */
function lockFolder() {
  const fid = folderId_();
  if (!fid) throw new Error('ยังไม่ได้ตั้งโฟลเดอร์ — รัน setProvinceFolderId(\'<folder id>\') ก่อน');
  const f = DriveApp.getFolderById(fid);
  const before = String(f.getSharingAccess());
  f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  Logger.log('โฟลเดอร์ "' + f.getName() + '": ' + before + ' -> PRIVATE');
}

/**
 * โหมด LINK_EDIT — ล็อกโฟลเดอร์ แล้วเปิดลิงก์แก้ไขรายไฟล์
 * พิมพ์ลิงก์ 12 จังหวัดออกมาให้คัดลอกไปส่งได้เลย
 *
 * ล็อกโฟลเดอร์ก่อนเสมอ เพราะการปิดโฟลเดอร์อาจดึงสิทธิ์ที่ไฟล์สืบทอดมาออกไปด้วย
 * จึงต้องตั้งสิทธิ์รายไฟล์ทีหลัง
 */
function shareProvinceLinks() {
  if (CONFIG.SHARE_MODE !== 'LINK_EDIT') {
    throw new Error('CONFIG.SHARE_MODE ไม่ใช่ LINK_EDIT — ถ้าจะแชร์รายอีเมลให้ใช้ hardenProvinceFiles()');
  }
  lockFolder();

  const ids = resolveProvinceFiles_();
  const lines = [];
  CONFIG.PROVINCES.forEach(function (p) {
    const id = ids[p.name];
    if (!id) return;
    shareProvinceFile_(id, p);
    lines.push(p.name + '\nhttps://docs.google.com/spreadsheets/d/' + id + '/edit');
  });

  // ไฟล์ ภ.4 มีข้อมูลครบ 12 จังหวัด ห้ามหลุดเป็นลิงก์สาธารณะไม่ว่ากรณีใด
  const m = propGet_(K_MASTER, '');
  if (m) {
    const mf = DriveApp.getFileById(m);
    if (String(mf.getSharingAccess()) !== 'PRIVATE') {
      mf.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      Logger.log('เตือน: ไฟล์ ภ.4 เคยถูกเปิดสาธารณะ — ตั้งกลับเป็น PRIVATE แล้ว');
    }
  }

  Logger.log('\n===== ลิงก์สำหรับส่งให้แต่ละจังหวัด (แก้ไขได้ ไม่ต้องล็อกอิน) =====\n\n' +
             lines.join('\n\n'));
  Logger.log('\nรัน auditSharing() เพื่อยืนยัน');
}

/** ปิดการแชร์ผ่านลิงก์ของทุกไฟล์ — ใช้เมื่อจะย้ายไปโหมด EMAIL_ONLY หรือเลิกใช้ระบบ */
function lockDownProvinceFiles() {
  const ids = resolveProvinceFiles_();
  const m = propGet_(K_MASTER, '');
  if (m) ids['[ภ.4 ส่วนกลาง]'] = m;

  Object.keys(ids).forEach(function (name) {
    try {
      const f = DriveApp.getFileById(ids[name]);
      const before = String(f.getSharingAccess());
      if (before === 'PRIVATE') { Logger.log(name + ': PRIVATE อยู่แล้ว'); return; }
      f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      Logger.log(name + ': ' + before + ' -> PRIVATE');
    } catch (e) {
      Logger.log(name + ': ปิดสิทธิ์ไม่ได้ (' + e.message + ')');
    }
  });
  if (folderId_()) lockFolder();
  Logger.log('เสร็จ — รัน auditSharing() เพื่อยืนยัน');
}

// ════════════════════════════════════════════════════════════════════
// 6. สร้างไฟล์ ภ.4 และผูกไฟล์จังหวัด
// ════════════════════════════════════════════════════════════════════

/** ผูกไฟล์ + สร้างไฟล์ ภ.4 + ติดตั้ง trigger + รวมข้อมูลรอบแรก */
function bindExisting() {
  const ids = resolveProvinceFiles_();
  const tmpl = templateSpreadsheet_(ids);
  const ref = readHeaders_(tmpl);

  const res = checkProvinceFiles_(ids, ref);
  if (res.problems.length) {
    res.problems.forEach(function (p) { Logger.log('ผิดพลาด: ' + p); });
    throw new Error('หัวตารางไม่ตรง template ' + res.problems.length + ' รายการ — ยกเลิกการผูกไฟล์');
  }

  propSetJson_(K_PROVS, ids);
  Logger.log('ผูกไฟล์ครบ ' + Object.keys(ids).length + ' จังหวัด');

  if (!propGet_(K_MASTER, '')) {
    const masterId = buildMasterFile_(ref, tmpl, rootFolder_());
    P.setProperty(K_MASTER, masterId);
    markCreated_(masterId);   // ไฟล์จังหวัดเป็นของคุณ resetAll() จะไม่แตะ
    Logger.log('สร้างไฟล์ ภ.4 แล้ว (' + masterId + ')');
  }

  installTriggers();
  syncNow();
  Logger.log('เสร็จ. ไฟล์ ภ.4: ' + masterUrl_());
  Logger.log('ขั้นถัดไป: hardenProvinceFiles()');
}

function buildMasterFile_(headersByTab, tmpl, root) {
  const ss = SpreadsheetApp.create(CONFIG.MASTER_NAME);
  const placeholder = ss.getSheets()[0];

  copySheet_(tmpl, README_TAB, ss, 0);
  stripProvinceReadme_(sheetByName_(ss, README_TAB));

  TABS.forEach(function (t, i) {
    buildMasterTab_(ss, t.tab, headersByTab[t.tab], i + 1);
  });

  copySheet_(tmpl, LISTS_TAB, ss, TABS.length + 1);
  buildRegistryTab_(ss);
  buildOverview_(ss, headersByTab);   // แทรกแท็บ ภาพรวม ไว้ถัดจาก README

  ss.deleteSheet(placeholder);
  ss.setActiveSheet(sheetByName_(ss, OVERVIEW_TAB));

  DriveApp.getFileById(ss.getId()).moveTo(root);
  applySharing_(ss.getId(), CONFIG.MASTER_EDITORS, CONFIG.MASTER_VIEWERS);
  return ss.getId();
}

/** ชีตข้อมูลในไฟล์ ภ.4: หัวตาราง + รูปแบบคอลัมน์เท่านั้น (Sync เขียนทับทุกครั้ง) */
function buildMasterTab_(ss, tabName, headers, index) {
  const sh = ss.insertSheet(tabName, index);
  ensureSize_(sh, 2, headers.length);
  styleHeader_(sh, headers, false);
  trimSize_(sh, 2, headers.length);
  return sh;
}

function copySheet_(src, name, dest, index) {
  const sh = sheetByName_(src, name).copyTo(dest);
  sh.setName(name);
  dest.setActiveSheet(sh);
  dest.moveActiveSheet(index + 1);
  return sh;
}

/**
 * ตัดข้อความ "ฉบับแยกจังหวัด" ที่ build_p4.py ต่อท้าย README ของไฟล์จังหวัดออก
 * เพราะไฟล์ ภ.4 ต้องส่งต่อส่วนกลาง README จึงต้องเหลือเฉพาะข้อความจาก template เดิม
 */
function stripProvinceReadme_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const col = sh.getRange(1, 1, last, 1).getValues();
  for (let r = 0; r < col.length; r++) {
    if (String(col[r][0]).trim() !== README_PROVINCE_MARKER) continue;
    const markerRow = r + 1;
    const from = Math.max(markerRow - 1, 2);   // ลบแถวว่างที่คั่นอยู่ข้างบนด้วย
    sh.deleteRows(from, last - from + 1);
    return;
  }
}

function buildRegistryTab_(ss) {
  const sh = ss.insertSheet(REGISTRY_TAB);
  const head = ['จังหวัด', 'Spreadsheet ID', 'ลิงก์', 'จำนวนแถวรวม', 'แก้ไขไฟล์ล่าสุด', 'ซิงค์ล่าสุด'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(HEADER_BG);
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 380);
  trimSize_(sh, CONFIG.PROVINCES.length + 1, head.length);
  return sh;
}

// ════════════════════════════════════════════════════════════════════
// 7. จัดรูปแบบไฟล์จังหวัดที่ import มา
// ════════════════════════════════════════════════════════════════════

/**
 * dropdown / ไฮไลต์ช่องบังคับ / ล็อกหัวตาราง + ตั้งสิทธิ์ตาม CONFIG.SHARE_MODE
 * การ import .xlsx พา data validation มาด้วย แต่ไม่พา conditional formatting และ protected range
 * ทำทีละจังหวัดและจำความคืบหน้า — หมดเวลาแล้วรันซ้ำได้
 */
function hardenProvinceFiles() {
  const t0 = Date.now();
  const ids = propJson_(K_PROVS, {});
  if (!Object.keys(ids).length) throw new Error('ยังไม่ได้ผูกไฟล์ — รัน bindExisting() ก่อน');

  const done = propJson_(K_HARDENED, []);
  for (let i = 0; i < CONFIG.PROVINCES.length; i++) {
    const p = CONFIG.PROVINCES[i];
    if (!ids[p.name] || done.indexOf(p.name) !== -1) continue;
    if (Date.now() - t0 > TIME_BUDGET_MS) {
      Logger.log('ใกล้หมดเวลา — รัน hardenProvinceFiles() อีกครั้ง (เหลือ ' +
                 (CONFIG.PROVINCES.length - done.length) + ' จังหวัด)');
      return;
    }
    hardenOne_(ids[p.name], p.name);
    done.push(p.name);
    propSetJson_(K_HARDENED, done);
    Logger.log('จัดรูปแบบ ' + p.name + ' แล้ว');
  }
  Logger.log('จัดรูปแบบครบ ' + done.length + ' จังหวัด');
}

function hardenOne_(id, province) {
  const ss = SpreadsheetApp.openById(id);
  const lists = readLists_(ss);
  const cfg = provinceCfg_(province);

  TABS.forEach(function (t) {
    const sh = sheetByName_(ss, t.tab);
    const nCol = sh.getLastColumn();
    const nRow = sh.getMaxRows();
    if (nRow < 2 || !nCol) return;
    const headers = sh.getRange(1, 1, 1, nCol).getValues()[0].map(String);

    applyColumnFormats_(sh, headers, nRow);
    styleHeader_(sh, headers, false);
    applyValidations_(sh, headers, nRow, province, t.type, lists);
    applyRequiredHighlight_(sh, headers, nRow);
    styleHeader_(sh, headers, true);   // ล็อกหัวตารางเป็นขั้นสุดท้าย
  });

  // แหล่งของ dropdown และคำอธิบาย ห้ามให้ใครแก้
  protectSheet_(sheetByName_(ss, LISTS_TAB), 'Lists');
  protectSheet_(sheetByName_(ss, README_TAB), 'README');

  shareProvinceFile_(id, cfg);
}

/** เลขบัตร/เลขเอกสาร = ข้อความ, คอลัมน์วันที่ = yyyy-mm-dd */
function applyColumnFormats_(sh, headers, nRow) {
  if (nRow < 2) return;
  headers.forEach(function (h, i) {
    const rng = sh.getRange(2, i + 1, nRow - 1, 1);
    if (TEXT_COLS.indexOf(h) !== -1) rng.setNumberFormat('@');
    else if (DATE_COLS.indexOf(h) !== -1) rng.setNumberFormat('yyyy-mm-dd');
  });
}

/** province = ล็อก dropdown จังหวัดเป็นค่าเดียว, type = ล็อกประเภทข้อมูลตามแท็บ */
function applyValidations_(sh, headers, nRow, province, type, lists) {
  headers.forEach(function (h, i) {
    let list = null;
    if (h === COL_TYPE) list = [type];
    else if (h === COL_PROVINCE) list = [province];
    else if (DROPDOWN_SOURCE[h] && lists[DROPDOWN_SOURCE[h]]) list = lists[DROPDOWN_SOURCE[h]];

    const rng = sh.getRange(2, i + 1, nRow - 1, 1);
    if (list) {
      rng.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(list, true)
        .setAllowInvalid(false)
        .setHelpText('เลือกจากรายการที่กำหนด')
        .build());
    } else if (DATE_COLS.indexOf(h) !== -1) {
      rng.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)   // เตือนอย่างเดียว ไม่บล็อกการวางข้อมูล
        .setHelpText('ใช้ ค.ศ. รูปแบบ yyyy-mm-dd เช่น 2026-06-08')
        .build());
    }
  });
}

/**
 * ช่องบังคับ (*) ที่ว่าง ขึ้นพื้นแดง เฉพาะแถวที่มีข้อมูลอื่นอยู่แล้ว
 *
 * เริ่มนับ COUNTA จากคอลัมน์ถัดจาก ประเภทข้อมูล/จังหวัด เพราะแถวสำรอง 500 แถว
 * ถูกเติมสองคอลัมน์นั้นไว้ล่วงหน้า ถ้านับรวมด้วย แถวสำรองจะขึ้นแดงทั้งหมด
 */
function applyRequiredHighlight_(sh, headers, nRow) {
  const ti = headers.indexOf(COL_TYPE);
  const pi = headers.indexOf(COL_PROVINCE);
  const startCol = (ti === 0 && pi === 1) ? 3 : 1;
  const startL = colLetter_(startCol);
  const lastL = colLetter_(headers.length);

  const rules = [];
  headers.forEach(function (h, i) {
    if (String(h).slice(-1) !== '*') return;
    const c = colLetter_(i + 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISBLANK(' + c + '2), COUNTA($' + startL + '2:$' + lastL + '2)>0)')
      .setBackground(MISSING_REQUIRED_BG)
      .setRanges([sh.getRange(2, i + 1, nRow - 1, 1)])
      .build());
  });
  sh.setConditionalFormatRules(rules);
}

// ════════════════════════════════════════════════════════════════════
// 7.5 แท็บภาพรวม + กราฟ
// ════════════════════════════════════════════════════════════════════

/** ใส่ชื่อชีตในสูตรอย่างปลอดภัย */
function q_(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}

/** ตัวอักษรคอลัมน์ของหัวคอลัมน์ชื่อใดชื่อหนึ่งในแท็บนั้น (null ถ้าไม่มี) */
function colOf_(headersByTab, tab, names) {
  const h = headersByTab[tab];
  for (let i = 0; i < names.length; i++) {
    const j = h.indexOf(names[i]);
    if (j !== -1) return colLetter_(j + 1);
  }
  return null;
}

/**
 * ตำแหน่งแถวบนแท็บ ภาพรวม — คำนวณจากจำนวนจังหวัด ไม่ hardcode
 * ทั้งสูตรและช่วงข้อมูลของกราฟอ้างค่าจากที่นี่ที่เดียว
 */
function overviewLayout_() {
  const nProv = CONFIG.PROVINCES.length;
  const nTabs = TABS.length;
  const nCols = COMPLETENESS_COLS.length;
  const countHead = 8;
  const compHead = countHead + nProv + 4;   // เว้น 1 แถวรวม + 2 บรรทัดคั่น + หัวข้อ
  return {
    nProv: nProv, nTabs: nTabs, nCols: nCols,
    kpiHead: 4, kpi: 5,
    countTitle: countHead - 1, countHead: countHead,
    countFirst: countHead + 1, countLast: countHead + nProv, countTotal: countHead + nProv + 1,
    countTotalCol: colLetter_(nTabs + 2),                  // คอลัมน์ "รวม"
    compTitle: compHead - 1, compHead: compHead,
    compFirst: compHead + 1, compLast: compHead + nProv, compTotal: compHead + nProv + 1,
    compAvgCol: 'B', compLastCol: colLetter_(2 + nCols),
    chartRow: compHead + nProv + 4,
  };
}

/**
 * สร้าง/สร้างใหม่แท็บ ภาพรวม, _ข้อมูลกราฟ, _ประวัติรายวัน ในไฟล์ ภ.4
 * รันซ้ำได้ — _ประวัติรายวัน จะถูกเก็บไว้ ไม่ถูกล้าง
 *
 * ใช้กับไฟล์ ภ.4 ที่สร้างไปแล้วก็ได้ (bindExisting ไม่ต้องรันใหม่)
 * แท็บ Dashboard เดิมถูกยุบเข้ามาที่นี่แล้ว จึงลบทิ้ง
 */
function installOverview() {
  const masterId = propGet_(K_MASTER, '');
  if (!masterId) throw new Error('ยังไม่ได้รัน bindExisting()');

  // ต้องกัน syncAuto (ทุก 5 นาที) ไม่ให้เขียนไฟล์เดียวกันพร้อมกัน
  // ไม่งั้น insertSheet จะล้มด้วย "Service Spreadsheets failed while accessing document"
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(90000)) throw new Error('มีการซิงค์ทำงานอยู่ — รอสักครู่แล้วลองใหม่');

  try {
    const master = SpreadsheetApp.openById(masterId);
    buildOverview_(master, readHeaders_(master));
    appendHistory_(master, propJson_(K_COUNTS, null));
    removeLegacyEditTrigger_();
    Logger.log('สร้างแท็บ ภาพรวม เรียบร้อย: ' + masterUrl_());
  } finally {
    lock.releaseLock();
  }
}

/** ปุ่มเป็นรูปที่ assignScript แล้ว ไม่ต้องใช้ onEdit trigger อีก — ลบของเก่าทิ้ง */
function removeLegacyEditTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditMaster') {
      ScriptApp.deleteTrigger(t);
      Logger.log('ลบ trigger onEditMaster ที่ไม่ใช้แล้ว');
    }
  });
}

/**
 * ลำดับสำคัญ: ต้องมีแท็บ ภาพรวม อยู่จริงก่อนเขียนสูตรใน _ข้อมูลกราฟ ที่อ้างถึงมัน
 * ไม่งั้น Google Sheets จะเปลี่ยนการอ้างอิงเป็น #REF! อย่างถาวร
 */
function buildOverview_(master, headersByTab) {
  [OVERVIEW_TAB, CHARTDATA_TAB, LEGACY_DASHBOARD_TAB].forEach(function (n) {
    const sh = master.getSheetByName(n);
    if (sh) master.deleteSheet(sh);
  });
  SpreadsheetApp.flush();   // ให้ฝั่งเซิร์ฟเวอร์ตามทันก่อนแทรกชีตใหม่

  const sh = master.insertSheet(OVERVIEW_TAB, 1);   // ถัดจาก README
  fillOverviewTables_(sh, headersByTab);
  const cd = buildChartDataTab_(master);
  ensureHistoryTab_(master);
  insertOverviewCharts_(master, sh, cd);
  return sh;
}

/** KPI + ตารางจำนวนข้อมูล + ตารางความครบถ้วน ทั้งหมดอยู่บนแท็บเดียว */
function fillOverviewTables_(sh, headersByTab) {
  const L = overviewLayout_();
  const tabNames = TABS.map(function (t) { return t.tab; });
  const provinces = CONFIG.PROVINCES.map(function (p) { return p.name; });
  const TC = L.countTotalCol;

  sh.getRange('A1').setValue('ภาพรวมข้อมูลบุคคล ภ.4').setFontSize(18).setFontWeight('bold');
  sh.getRange('A2').setValue('อัปเดตล่าสุด').setFontWeight('bold');
  sh.getRange('B2').setValue(propGet_(K_SYNC_AT, '(ยังไม่เคยซิงค์)'));

  insertSyncButton_(sh);
  sh.getRange(SYNC_STATUS_CELL).setValue('กดปุ่มเพื่อดึงข้อมูลจากทั้ง 12 จังหวัด');

  // ---- ตารางจำนวนข้อมูล: จังหวัด × ประเภท
  const provCol = {};
  TABS.forEach(function (t) {
    provCol[t.tab] = colOf_(headersByTab, t.tab, [COL_PROVINCE]);
  });

  sh.getRange(L.countTitle, 1).setValue('จำนวนข้อมูล (จังหวัด × ประเภท)').setFontWeight('bold');
  sh.getRange(L.countHead, 1, 1, L.nTabs + 2).setValues([['จังหวัด'].concat(tabNames, ['รวม'])])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(HEADER_BG)
    .setWrap(true).setHorizontalAlignment('center');

  const countRows = provinces.map(function (prov, i) {
    const r = L.countFirst + i;
    const cells = [prov];
    tabNames.forEach(function (tab) {
      cells.push('=COUNTIF(' + q_(tab) + '!$' + provCol[tab] + ':$' + provCol[tab] + ',$A' + r + ')');
    });
    cells.push('=SUM(B' + r + ':' + colLetter_(L.nTabs + 1) + r + ')');
    return cells;
  });
  sh.getRange(L.countFirst, 1, L.nProv, L.nTabs + 2).setValues(countRows);

  const countTotals = ['รวม'];
  for (let c = 2; c <= L.nTabs + 2; c++) {
    const cl = colLetter_(c);
    countTotals.push('=SUM(' + cl + L.countFirst + ':' + cl + L.countLast + ')');
  }
  sh.getRange(L.countTotal, 1, 1, L.nTabs + 2).setValues([countTotals]).setFontWeight('bold');
  sh.getRange(L.countFirst, 2, L.nProv + 1, L.nTabs + 1).setNumberFormat('#,##0');

  // ---- ตารางความครบถ้วน: จังหวัด × คอลัมน์บังคับ
  sh.getRange(L.compTitle, 1).setValue('ความครบถ้วนของช่องบังคับ (จังหวัด × คอลัมน์)').setFontWeight('bold');
  sh.getRange(L.compHead, 1, 1, L.nCols + 2)
    .setValues([['จังหวัด', 'เฉลี่ย'].concat(COMPLETENESS_COLS.map(function (c) { return c.label; }))])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(HEADER_BG)
    .setWrap(true).setHorizontalAlignment('center');

  const compRows = provinces.map(function (prov, i) {
    const rk = L.compFirst + i;
    const rc = L.countFirst + i;          // แถวเดียวกันในตารางจำนวน = ตัวหาร
    const cells = [prov, '=IFERROR(AVERAGE(C' + rk + ':' + L.compLastCol + rk + '),0)'];
    COMPLETENESS_COLS.forEach(function (spec) {
      const terms = [];
      TABS.forEach(function (t) {
        const pc = provCol[t.tab];
        const vc = colOf_(headersByTab, t.tab, spec.names);
        if (!pc || !vc) return;
        terms.push('COUNTIFS(' + q_(t.tab) + '!$' + pc + ':$' + pc + ',$A' + rk +
                   ',' + q_(t.tab) + '!$' + vc + ':$' + vc + ',"<>")');
      });
      cells.push(terms.length ? '=IFERROR((' + terms.join('+') + ')/$' + TC + '$' + rc + ',0)' : 0);
    });
    return cells;
  });
  sh.getRange(L.compFirst, 1, L.nProv, L.nCols + 2).setValues(compRows);

  // แถวรวมถ่วงน้ำหนักด้วยจำนวนแถวของแต่ละจังหวัด (ไม่ใช่เฉลี่ยธรรมดา)
  const w = '$' + TC + '$' + L.countFirst + ':$' + TC + '$' + L.countLast;
  const compTotals = ['รวม (ถ่วงน้ำหนัก)'];
  for (let c = 2; c <= L.nCols + 2; c++) {
    const cl = colLetter_(c);
    compTotals.push('=IFERROR(SUMPRODUCT(' + cl + L.compFirst + ':' + cl + L.compLast + ',' + w +
                    ')/SUM(' + w + '),0)');
  }
  sh.getRange(L.compTotal, 1, 1, L.nCols + 2).setValues([compTotals]).setFontWeight('bold');
  sh.getRange(L.compFirst, 2, L.nProv + 1, L.nCols + 1).setNumberFormat('0.0%');

  // ---- KPI (อ้างตารางบนแท็บเดียวกัน)
  sh.getRange(L.kpiHead, 1, 1, 4)
    .setValues([['รวมทุกประเภท', 'จังหวัดที่มีข้อมูล', 'ความครบถ้วนช่องบังคับ', 'ช่องบังคับที่ยังว่าง']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(HEADER_BG)
    .setHorizontalAlignment('center').setWrap(true).setVerticalAlignment('middle');
  sh.setRowHeight(L.kpiHead, 34);
  sh.getRange(L.kpi, 1, 1, 4).setFormulas([[
    '=$' + TC + '$' + L.countTotal,
    '=COUNTIF(' + w + ',">0")',
    '=' + L.compAvgCol + L.compTotal,
    '=ROUND((1-C' + L.kpi + ')*$' + TC + '$' + L.countTotal + '*' + L.nCols + ',0)',
  ]]).setFontSize(16).setHorizontalAlignment('center');
  sh.getRange(L.kpi, 1).setNumberFormat('#,##0');
  sh.getRange(L.kpi, 3).setNumberFormat('0.0%');
  sh.getRange(L.kpi, 4).setNumberFormat('#,##0');

  sh.setColumnWidth(1, 200);
  return sh;
}

/**
 * วางปุ่มกดจริงบนแท็บ ภาพรวม
 *
 * OverGridImage.assignScript() ผูกรูปเข้ากับฟังก์ชัน กดแล้วรันทันที
 * ไม่ต้องใช้ checkbox และไม่ต้องมี onEdit trigger คอยดัก
 */
function insertSyncButton_(sh) {
  sh.getImages().forEach(function (img) { img.remove(); });
  sh.setRowHeight(2, 50);

  const blob = Utilities.newBlob(Utilities.base64Decode(SYNC_BUTTON_PNG_B64), 'image/png', 'sync.png');
  sh.insertImage(blob, 3, 2, 4, 3)   // คอลัมน์ C แถว 2
    .setWidth(260)
    .setHeight(44)
    .assignScript('syncFromButton');
}

/**
 * ฟังก์ชันที่ผูกกับปุ่ม — ต้องเป็นฟังก์ชันระดับบนสุด ไม่รับพารามิเตอร์
 * ใช้ sync แบบ incremental (~10 วิ) ไม่ใช่ force rebuild (~106 วิ) กดบ่อยก็ไม่กินโควตา
 */
function syncFromButton() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERVIEW_TAB);
  if (!sh) return;
  const status = sh.getRange(SYNC_STATUS_CELL);

  status.setValue('กำลังดึงข้อมูล...');
  SpreadsheetApp.flush();
  try {
    status.setValue(runSync_(false) || 'เสร็จ');
  } catch (e) {
    status.setValue('ผิดพลาด: ' + e.message);
    throw e;
  }
}

/**
 * ชีตซ่อน เก็บเฉพาะบล็อก 2 คอลัมน์ที่กราฟวงกลม/กราฟรายคอลัมน์ต้องใช้
 * (ตัวเลขจริงอยู่บนแท็บ ภาพรวม แล้ว ที่นี่แค่จัดรูปให้กราฟกินได้)
 */
function buildChartDataTab_(master) {
  const sh = master.insertSheet(CHARTDATA_TAB);
  const L = overviewLayout_();
  const OV = q_(OVERVIEW_TAB);

  const pie = [['ประเภทข้อมูล', 'จำนวนแถว']];
  TABS.forEach(function (t, i) {
    pie.push([t.tab, '=' + OV + '!' + colLetter_(i + 2) + L.countTotal]);
  });
  sh.getRange(1, 1, pie.length, 2).setValues(pie);

  const byCol = [['คอลัมน์บังคับ', '% ครบถ้วน']];
  COMPLETENESS_COLS.forEach(function (spec, j) {
    byCol.push([spec.label, '=' + OV + '!' + colLetter_(3 + j) + L.compTotal]);
  });
  sh.getRange(1, 4, byCol.length, 2).setValues(byCol);
  sh.getRange(2, 5, COMPLETENESS_COLS.length, 1).setNumberFormat('0.0%');

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(4, 180);
  sh.hideSheet();
  return sh;
}

function insertOverviewCharts_(master, sh, cd) {
  const L = overviewLayout_();
  const hist = sheetByName_(master, HISTORY_TAB);

  const charts = [
    { type: 'column', stacked: true, ranges: [sh.getRange(L.countHead, 1, L.nProv + 1, L.nTabs + 1)],
      title: 'จำนวนข้อมูลรายจังหวัด แยกตามประเภท', row: L.chartRow, col: 1, w: 900, h: 400 },
    { type: 'pie', ranges: [cd.getRange(1, 1, L.nTabs + 1, 2)],
      title: 'สัดส่วนตามประเภทข้อมูล', row: L.chartRow, col: 11, w: 520, h: 400 },
    { type: 'bar', ranges: [sh.getRange(L.compHead, 1, L.nProv + 1, 2)],
      title: 'ความครบถ้วนของช่องบังคับ รายจังหวัด', row: L.chartRow + 22, col: 1, w: 700, h: 420 },
    { type: 'column', ranges: [cd.getRange(1, 4, L.nCols + 1, 2)],
      title: 'ความครบถ้วนของช่องบังคับ รายคอลัมน์', row: L.chartRow + 22, col: 11, w: 720, h: 420 },
    { type: 'line', ranges: [hist.getRange('A:B')],
      title: 'จำนวนข้อมูลรวม รายวัน', row: L.chartRow + 45, col: 1, w: 900, h: 350 },
  ];

  charts.forEach(function (c) {
    let b = sh.newChart().setNumHeaders(1)
      .setPosition(c.row, c.col, 0, 0)
      .setOption('title', c.title)
      .setOption('width', c.w)
      .setOption('height', c.h)
      .setOption('legend', { position: 'right' });
    c.ranges.forEach(function (r) { b = b.addRange(r); });

    if (c.type === 'pie') b = b.asPieChart();
    else if (c.type === 'bar') b = b.asBarChart();
    else if (c.type === 'line') b = b.asLineChart();
    else { b = b.asColumnChart(); if (c.stacked) b = b.setStacked(); }

    sh.insertChart(b.build());
  });
}

function ensureHistoryTab_(master) {
  let sh = master.getSheetByName(HISTORY_TAB);
  if (sh) return sh;
  sh = master.insertSheet(HISTORY_TAB);
  const head = ['วันที่', 'รวมทั้งหมด'].concat(CONFIG.PROVINCES.map(function (p) { return p.name; }));
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.hideSheet();
  return sh;
}

/** บันทึกยอดของวันนี้ — เขียนทับถ้ามีแถวของวันนี้อยู่แล้ว */
function appendHistory_(master, counts) {
  if (!counts) return;
  const sh = ensureHistoryTab_(master);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  const perProv = CONFIG.PROVINCES.map(function (p) {
    let n = 0;
    TABS.forEach(function (t) { n += (counts[t.tab] && counts[t.tab][p.name]) || 0; });
    return n;
  });
  const total = perProv.reduce(function (a, b) { return a + b; }, 0);
  const row = [today, total].concat(perProv);

  const last = sh.getLastRow();
  if (last >= 2 && sh.getRange(last, 1).getDisplayValue() === today) {
    sh.getRange(last, 1, 1, row.length).setValues([row]);
  } else {
    ensureSize_(sh, last + 1, row.length);
    sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
  }
}

// ════════════════════════════════════════════════════════════════════
// 8. รวมข้อมูลเข้าไฟล์ ภ.4
// ════════════════════════════════════════════════════════════════════

function onOpenMaster(e) {
  SpreadsheetApp.getUi()
    .createMenu('ภ.4')
    .addItem('รวมข้อมูลเดี๋ยวนี้', 'syncNow')
    .addItem('สถานะการซิงค์', 'showSyncStatus')
    .addSeparator()
    .addItem('สร้าง/รีเฟรชแท็บภาพรวม', 'installOverview')
    .addItem('สร้างไฟล์ส่งส่วนกลาง (.xlsx)', 'exportForCentral')
    .addToUi();
}

function syncNow() { return runSync_(true); }
function syncAuto() { return runSync_(false); }

/** คืนข้อความสรุปผล เพื่อให้ปุ่มบนแท็บ ภาพรวม เอาไปแสดงได้ */
function runSync_(force) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(force ? 60000 : 1000)) {
    const busy = 'มีการซิงค์อื่นทำงานอยู่ — ข้ามรอบนี้';
    Logger.log(busy);
    return busy;
  }
  try {
    return doSync_(force);
  } finally {
    lock.releaseLock();
  }
}

/**
 * แถวในชีตส่วนกลางเรียงตามลำดับจังหวัดใน CONFIG.PROVINCES เสมอ
 * K_COUNTS จึงบอกได้ว่าแต่ละจังหวัดกินแถวไหนถึงแถวไหน -> แก้เฉพาะบล็อกที่เปลี่ยนได้
 *
 * สำคัญเพราะบัญชี @gmail.com มีโควตา trigger รวม 90 นาที/วัน
 * ถ้า rebuild ทั้ง 17,553 แถวทุกครั้งที่มีคนแก้ 1 ช่อง (~106 วินาที) โควตาหมดตั้งแต่กลางวัน
 */
function doSync_(force) {
  const t0 = Date.now();
  const masterId = propGet_(K_MASTER, '');
  const ids = propJson_(K_PROVS, {});
  if (!masterId || !Object.keys(ids).length) throw new Error('ยังไม่ได้รัน bindExisting()');

  // 1. มีจังหวัดไหนแก้ไฟล์บ้าง (12 Drive call ถูกกว่าเปิดอ่านทั้ง 12 ไฟล์มาก)
  const prevStamps = propJson_(K_STAMPS, {});
  const stamps = {};
  const changedNames = [];
  CONFIG.PROVINCES.forEach(function (p) {
    if (!ids[p.name]) return;
    const t = DriveApp.getFileById(ids[p.name]).getLastUpdated().getTime();
    stamps[p.name] = t;
    if (prevStamps[p.name] !== t) changedNames.push(p.name);
  });
  if (!force && !changedNames.length) {
    const none = 'ไม่มีจังหวัดใดแก้ไข — ข้อมูลเป็นปัจจุบันแล้ว (' + stampNowShort_() + ')';
    Logger.log(none);
    return none;
  }

  const master = SpreadsheetApp.openById(masterId);
  const stampNow = now_();
  const masterHeaders = {};
  TABS.forEach(function (t) {
    const msh = sheetByName_(master, t.tab);
    masterHeaders[t.tab] = msh.getRange(1, 1, 1, msh.getLastColumn()).getValues()[0].map(String);
  });

  let counts = propJson_(K_COUNTS, null);
  let mode;

  if (force || !countsUsable_(counts)) {
    counts = fullRebuild_(master, ids, masterHeaders, t0);
    mode = 'rebuild เต็ม';
  } else {
    try {
      incrementalSync_(master, ids, masterHeaders, changedNames, counts, t0);
      mode = 'แก้เฉพาะ ' + changedNames.join(', ');
    } catch (e) {
      Logger.log('แก้เฉพาะบล็อกล้มเหลว (' + e.message + ') — rebuild ทั้งหมดแทน');
      counts = fullRebuild_(master, ids, masterHeaders, t0);
      mode = 'rebuild เต็ม (กู้จากข้อผิดพลาด)';
    }
  }
  SpreadsheetApp.flush();

  // ไฟล์ ภ.4 ที่สร้างก่อนมีแท็บภาพรวม จะยังไม่มีชีตนี้ — ไม่ให้ล้มทั้ง sync
  const ov = master.getSheetByName(OVERVIEW_TAB);
  if (ov) ov.getRange('B2').setValue(stampNow);

  writeRegistry_(master, ids, counts, stampNow);
  appendHistory_(master, counts);

  propSetJson_(K_COUNTS, counts);
  propSetJson_(K_STAMPS, stamps);
  P.setProperty(K_SYNC_AT, stampNow);

  const msg = 'รวมข้อมูลสำเร็จ ' + totalRows_(counts).toLocaleString() + ' แถว [' + mode + '] ใน ' +
              Math.round((Date.now() - t0) / 1000) + ' วินาที';
  Logger.log(msg);
  return msg;
}

function stampNowShort_() {
  return Utilities.formatDate(new Date(), TZ, 'HH:mm');
}

/** counts ใช้ได้ก็ต่อเมื่อมีครบทุกแท็บและทุกจังหวัดที่ผูกไว้ */
function countsUsable_(counts) {
  if (!counts) return false;
  for (let i = 0; i < TABS.length; i++) {
    const c = counts[TABS[i].tab];
    if (!c) return false;
    for (let j = 0; j < CONFIG.PROVINCES.length; j++) {
      if (typeof c[CONFIG.PROVINCES[j].name] !== 'number') return false;
    }
  }
  return true;
}

function totalRows_(counts) {
  let n = 0;
  TABS.forEach(function (t) {
    CONFIG.PROVINCES.forEach(function (p) { n += counts[t.tab][p.name] || 0; });
  });
  return n;
}

/** เปิดสเปรดชีตจังหวัดครั้งเดียวต่อการซิงค์ 1 รอบ */
function provinceSS_(ids, name, cache) {
  if (!cache[name]) cache[name] = SpreadsheetApp.openById(ids[name]);
  return cache[name];
}

/**
 * อ่านแท็บหนึ่งของจังหวัดหนึ่ง แล้วแปลงให้ตรงคอลัมน์ของชีตส่วนกลาง
 * จับคู่ด้วยชื่อหัวคอลัมน์ ไม่ใช่ตำแหน่ง — จังหวัดเผลอสลับคอลัมน์ก็ยังรวมถูก
 */
function readProvinceTab_(ss, t, dstHeaders, provinceName) {
  const res = readTab_(ss, t.tab);
  if (!res.headers.length) return null;

  const map = dstHeaders.map(function (h) { return res.headers.indexOf(h); });
  const iType = dstHeaders.indexOf(COL_TYPE);
  const iProv = dstHeaders.indexOf(COL_PROVINCE);

  return res.rows.map(function (r) {
    const out = new Array(dstHeaders.length);
    for (let k = 0; k < map.length; k++) out[k] = map[k] === -1 ? '' : r[map[k]];
    // แหล่งความจริงคือ "มาจากไฟล์ไหน แท็บไหน" ไม่ใช่สิ่งที่เจ้าหน้าที่พิมพ์
    if (iType !== -1) out[iType] = t.type;
    if (iProv !== -1) out[iProv] = provinceName;
    return out;
  });
}

/** อ่านทั้ง 12 จังหวัด เขียนทับทุกชีต คืน counts ชุดใหม่ */
function fullRebuild_(master, ids, masterHeaders, t0) {
  const cache = {};
  const buckets = {};
  const counts = {};
  const skipped = [];
  TABS.forEach(function (t) { buckets[t.tab] = []; counts[t.tab] = {}; });

  CONFIG.PROVINCES.forEach(function (p) {
    if (!ids[p.name]) return;
    let ss;
    try {
      ss = provinceSS_(ids, p.name, cache);
    } catch (e) {
      skipped.push(p.name + ' (เปิดไฟล์ไม่ได้: ' + e.message + ')');
      TABS.forEach(function (t) { counts[t.tab][p.name] = 0; });
      return;
    }
    TABS.forEach(function (t) {
      const rows = readProvinceTab_(ss, t, masterHeaders[t.tab], p.name);
      if (rows === null) {
        skipped.push(p.name + '/' + t.tab + ' (ไม่พบแท็บ)');
        counts[t.tab][p.name] = 0;
        return;
      }
      buckets[t.tab] = buckets[t.tab].concat(rows);
      counts[t.tab][p.name] = rows.length;
    });
    if (Date.now() - t0 > TIME_BUDGET_MS) {
      throw new Error('อ่านข้อมูลไม่ทันใน 4.5 นาที (ค้างที่ ' + p.name + ')');
    }
  });

  TABS.forEach(function (t) {
    const msh = sheetByName_(master, t.tab);
    const headers = masterHeaders[t.tab];
    const rows = buckets[t.tab];
    clearData_(msh);
    ensureSize_(msh, Math.max(rows.length + 1, 2), headers.length);
    if (rows.length) {
      applyColumnFormats_(msh, headers, rows.length + 1);
      writeRows_(msh, 2, rows);
    }
  });

  if (skipped.length) Logger.log('ข้าม: ' + skipped.join(' | '));
  return counts;
}

/** อ่านเฉพาะจังหวัดที่เปลี่ยน แล้วเปลี่ยนเฉพาะบล็อกแถวของจังหวัดนั้นในแต่ละชีต */
function incrementalSync_(master, ids, masterHeaders, changedNames, counts, t0) {
  const cache = {};
  const order = CONFIG.PROVINCES.map(function (p) { return p.name; });
  const changed = order.filter(function (n) { return changedNames.indexOf(n) !== -1; });

  TABS.forEach(function (t) {
    const msh = sheetByName_(master, t.tab);
    const headers = masterHeaders[t.tab];

    changed.forEach(function (name) {
      if (!ids[name]) return;
      const rows = readProvinceTab_(provinceSS_(ids, name, cache), t, headers, name);
      if (rows === null) throw new Error('ไม่พบแท็บ ' + t.tab + ' ในไฟล์ ' + name);
      spliceBlock_(msh, headers, counts[t.tab], order, name, rows);
    });

    if (Date.now() - t0 > TIME_BUDGET_MS) {
      throw new Error('ซิงค์ไม่ทันใน 4.5 นาที (ค้างที่แท็บ ' + t.tab + ')');
    }
  });
}

/**
 * แทนที่บล็อกแถวของจังหวัดหนึ่งในชีตส่วนกลาง แล้วอัปเดต tabCounts
 * offset นับจาก tabCounts ปัจจุบันเสมอ จึงถูกต้องแม้หลายจังหวัดเปลี่ยนพร้อมกัน
 */
function spliceBlock_(msh, headers, tabCounts, order, province, rows) {
  const nCol = headers.length;
  const newN = rows.length;
  const oldN = tabCounts[province] || 0;

  let offset = 1;   // แถวหัวตาราง
  for (let i = 0; i < order.length; i++) {
    if (order[i] === province) break;
    offset += tabCounts[order[i]] || 0;
  }

  // ถ้าจำนวนแถวที่จำไว้เพี้ยนจากชีตจริง ให้ล้มไปทาง rebuild เต็มดีกว่าเขียนผิดที่
  if (offset + oldN > msh.getMaxRows()) {
    throw new Error('จำนวนแถวที่จำไว้ไม่ตรงกับชีตจริง (' + province + ')');
  }

  if (newN === oldN) {
    if (newN) msh.getRange(offset + 1, 1, newN, nCol).setValues(rows);
    return;
  }

  if (oldN) msh.deleteRows(offset + 1, oldN);
  if (newN) {
    const maxRows = msh.getMaxRows();
    if (offset + 1 > maxRows) msh.insertRowsAfter(maxRows, offset + newN - maxRows);
    else msh.insertRowsBefore(offset + 1, newN);
    msh.getRange(offset + 1, 1, newN, nCol).setValues(rows);
    applyFormatsToBlock_(msh, headers, offset + 1, newN);
  }
  tabCounts[province] = newN;
}

/** แถวที่แทรกใหม่ไม่รับรูปแบบคอลัมน์เสมอไป — ตั้งซ้ำเฉพาะบล็อก */
function applyFormatsToBlock_(sh, headers, startRow, n) {
  if (n < 1) return;
  headers.forEach(function (h, i) {
    if (TEXT_COLS.indexOf(h) !== -1) sh.getRange(startRow, i + 1, n, 1).setNumberFormat('@');
    else if (DATE_COLS.indexOf(h) !== -1) sh.getRange(startRow, i + 1, n, 1).setNumberFormat('yyyy-mm-dd');
  });
}

function writeRegistry_(master, ids, counts, stampNow) {
  const sh = sheetByName_(master, REGISTRY_TAB);
  const rows = CONFIG.PROVINCES.map(function (p) {
    const id = ids[p.name] || '';
    if (!id) return [p.name, '', '', 0, '', 'ยังไม่ได้ผูกไฟล์'];

    let total = 0;
    TABS.forEach(function (t) { total += (counts[t.tab] && counts[t.tab][p.name]) || 0; });

    let updated;
    try {
      updated = Utilities.formatDate(DriveApp.getFileById(id).getLastUpdated(), TZ, 'yyyy-MM-dd HH:mm:ss');
    } catch (e) {
      updated = '(อ่านไม่ได้)';
    }
    return [p.name, id, 'https://docs.google.com/spreadsheets/d/' + id + '/edit',
            total, updated, stampNow];
  });
  ensureSize_(sh, rows.length + 1, 6);
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
}

function showSyncStatus() {
  const ids = propJson_(K_PROVS, {});
  SpreadsheetApp.getUi().alert(
    'สถานะการซิงค์ ภ.4',
    'ซิงค์ล่าสุด: ' + propGet_(K_SYNC_AT, '(ยังไม่เคยซิงค์)') +
    '\nไฟล์จังหวัดที่ผูกไว้: ' + Object.keys(ids).length + ' / ' + CONFIG.PROVINCES.length +
    '\nรอบอัตโนมัติ: ทุก ' + CONFIG.SYNC_INTERVAL_MINUTES + ' นาที',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════════
// 9. ส่งต่อส่วนกลาง และงานดูแลระบบ
// ════════════════════════════════════════════════════════════════════

/**
 * สำเนาไฟล์ ภ.4 ที่ลบแท็บช่วยงานออกแล้ว เหลือ README + 8 แท็บข้อมูล + Lists
 * ตรงโครง template ต้นฉบับ พร้อมลิงก์ดาวน์โหลด .xlsx
 */
function exportForCentral() {
  const masterId = propGet_(K_MASTER, '');
  if (!masterId) throw new Error('ยังไม่ได้รัน bindExisting()');

  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const copy = DriveApp.getFileById(masterId).makeCopy('ภ.4_ส่งส่วนกลาง_' + stamp, rootFolder_());
  const ss = SpreadsheetApp.openById(copy.getId());

  HELPER_TABS.forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(sheetByName_(ss, README_TAB));
  SpreadsheetApp.flush();

  const xlsx = 'https://docs.google.com/spreadsheets/d/' + copy.getId() + '/export?format=xlsx';
  Logger.log('ไฟล์สำหรับส่งส่วนกลาง: ' + copy.getUrl());
  Logger.log('ดาวน์โหลด .xlsx: ' + xlsx);
  return xlsx;
}

/**
 * สำเนาไฟล์ ภ.4 เก็บไว้เป็นชุดสำรอง เก็บล่าสุด CONFIG.SNAPSHOT_KEEP ชุด
 *
 * จำเป็นในโหมด LINK_EDIT เพราะ doSync_() เขียนทับชีตส่วนกลางทั้งชีต
 * ถ้ามีคนลบแถวในไฟล์จังหวัด ข้อมูลจะหายจากไฟล์ ภ.4 ภายใน SYNC_INTERVAL_MINUTES นาที
 * ชุดสำรองคือทางเดียวที่จะกู้กลับมาได้
 */
function snapshotMaster() {
  const m = propGet_(K_MASTER, '');
  if (!m) { Logger.log('ยังไม่มีไฟล์ ภ.4 — ข้ามการสำรอง'); return; }

  const snapFolder = folder_(rootFolder_(), 'สำรองข้อมูลรายวัน');
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmm');
  const copy = DriveApp.getFileById(m).makeCopy('ภ.4_สำรอง_' + stamp, snapFolder);
  copy.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  Logger.log('สำรองข้อมูลแล้ว: ' + copy.getName());

  // ชุดสำรองไม่เข้ารายการ CREATED_IDS โดยตั้งใจ — resetAll() จะไม่ลบทิ้ง
  const files = [];
  const it = snapFolder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ f: f, t: f.getDateCreated().getTime() });
  }
  files.sort(function (a, b) { return b.t - a.t; });
  files.slice(CONFIG.SNAPSHOT_KEEP).forEach(function (x) {
    x.f.setTrashed(true);
    Logger.log('ลบชุดสำรองเก่า: ' + x.f.getName());
  });
}

function installTriggers() {
  const masterId = propGet_(K_MASTER, '');
  if (!masterId) throw new Error('ยังไม่มีไฟล์ ภ.4 — รัน bindExisting() ก่อน');

  // onEditMaster อยู่ในรายการเพื่อ "ลบ" ของเก่าเท่านั้น ไม่สร้างใหม่ — ปุ่มใช้ assignScript แทนแล้ว
  const MINE = ['syncAuto', 'onOpenMaster', 'onEditMaster', 'snapshotMaster'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (MINE.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAuto').timeBased().everyMinutes(CONFIG.SYNC_INTERVAL_MINUTES).create();
  ScriptApp.newTrigger('onOpenMaster').forSpreadsheet(masterId).onOpen().create();
  ScriptApp.newTrigger('snapshotMaster').timeBased().atHour(1).everyDays(1).create();
}

/** แชร์ไฟล์จังหวัดเพิ่มภายหลัง: shareProvince('ขอนแก่น', ['a@b.go.th'], []) */
function shareProvince(name, editors, viewers) {
  const ids = propJson_(K_PROVS, {});
  if (!ids[name]) throw new Error('ยังไม่ได้ผูกไฟล์ของจังหวัด ' + name);
  applySharing_(ids[name], editors || [], viewers || []);
  Logger.log('แชร์ ' + name + ' ให้ ' + (editors || []).join(', '));
}

function listFiles() {
  Logger.log('ภ.4 ส่วนกลาง: ' + masterUrl_());
  const ids = propJson_(K_PROVS, {});
  Object.keys(ids).forEach(function (n) {
    Logger.log(n + ': https://docs.google.com/spreadsheets/d/' + ids[n] + '/edit');
  });
}

/** ล้างความคืบหน้าของ hardenProvinceFiles() เพื่อสั่งจัดรูปแบบใหม่ทั้งหมด */
function resetHardenProgress() {
  P.deleteProperty(K_HARDENED);
  Logger.log('ล้างแล้ว — รัน hardenProvinceFiles() ใหม่ได้');
}

/**
 * ย้ายเฉพาะไฟล์ที่สคริปต์สร้างเอง (ไฟล์ ภ.4) ลงถังขยะ แล้วล้างสถานะ
 * ไฟล์จังหวัด 12 ไฟล์ที่คุณอัปโหลดเองจะไม่ถูกแตะ
 */
function resetAll() {
  if (!CONFIG.ALLOW_RESET) throw new Error('ตั้ง CONFIG.ALLOW_RESET = true ก่อน หากต้องการลบจริง');

  propJson_(K_CREATED, []).forEach(function (id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      Logger.log('ลงถังขยะ: ' + id);
    } catch (e) {
      Logger.log('ลบไม่ได้ ' + id + ': ' + e);
    }
  });
  Logger.log('ไม่แตะไฟล์จังหวัดทั้ง 12 ไฟล์');

  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  P.deleteAllProperties();
  Logger.log('ล้างสถานะเรียบร้อย');
}
