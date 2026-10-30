/**
 * Lớp dữ liệu phòng nghỉ.
 * File này chỉ đọc và chuẩn hóa Excel, không thao tác giao diện.
 */
(function createResortData(global) {
  'use strict';

  const EXCEL_PATH = './data/rooms.xlsx';
  const SHEET_NAME = 'Rooms';
  const REQUIRED_HEADERS = [
    'id', 'name', 'type', 'description', 'pricePerNight', 'capacity',
    'areaM2', 'bedType', 'amenities', 'image', 'status'
  ];
  const VALID_STATUSES = new Set(['Available', 'Maintenance']);
  let roomCache = [];

  async function loadRooms() {
    if (!global.XLSX) {
      throw new Error('Không tải được thư viện đọc Excel. Hãy kiểm tra kết nối Internet rồi tải lại trang.');
    }

    let response;
    try {
      // no-store giúp thay đổi trong Excel xuất hiện ngay sau khi tải lại trang.
      response = await fetch(EXCEL_PATH, { cache: 'no-store' });
    } catch (error) {
      throw new Error('Không thể tải file Excel. Hãy chạy website bằng Live Server, không mở bằng file://.');
    }
    if (!response.ok) {
      throw new Error(`Không đọc được ${EXCEL_PATH} (HTTP ${response.status}).`);
    }

    let workbook;
    try {
      const bytes = await response.arrayBuffer();
      workbook = global.XLSX.read(bytes, { type: 'array' });
    } catch (error) {
      throw new Error('File rooms.xlsx không hợp lệ hoặc đang bị hỏng.');
    }

    const worksheet = workbook.Sheets[SHEET_NAME];
    if (!worksheet) throw new Error(`Không tìm thấy worksheet "${SHEET_NAME}" trong rooms.xlsx.`);

    const matrix = global.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: true });
    if (!matrix.length) throw new Error('Worksheet Rooms đang trống.');
    validateHeaders(matrix[0]);

    const rawRows = global.XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: true });
    if (!rawRows.length) throw new Error('Worksheet Rooms chưa có dữ liệu phòng.');

    const ids = new Set();
    roomCache = rawRows.map((row, index) => {
      const room = normalizeRoom(row, index + 2);
      if (ids.has(room.id)) throw new Error(`Mã phòng bị trùng: ${room.id} (dòng ${index + 2}).`);
      ids.add(room.id);
      return Object.freeze(room);
    });
    return [...roomCache];
  }

  function validateHeaders(headerRow) {
    const headers = headerRow.map((value) => String(value).trim());
    const missing = REQUIRED_HEADERS.filter((header) => !headers.includes(header));
    if (missing.length) throw new Error(`Excel thiếu cột bắt buộc: ${missing.join(', ')}.`);
  }

  function normalizeRoom(row, rowNumber) {
    const text = (key, max = 1000) => String(row[key] ?? '').trim().slice(0, max);
    const number = (key) => Number(row[key]);
    const id = text('id', 40);
    const name = text('name', 120);
    const type = text('type', 60);
    const pricePerNight = number('pricePerNight');
    const capacity = number('capacity');
    const areaM2 = number('areaM2');
    const status = text('status', 30);

    if (!id || !name || !type) throw new Error(`Dòng ${rowNumber}: id, name và type không được để trống.`);
    if (!Number.isFinite(pricePerNight) || pricePerNight <= 0) throw new Error(`Dòng ${rowNumber}: pricePerNight phải là số lớn hơn 0.`);
    if (!Number.isInteger(capacity) || capacity < 1) throw new Error(`Dòng ${rowNumber}: capacity phải là số nguyên dương.`);
    if (!Number.isFinite(areaM2) || areaM2 <= 0) throw new Error(`Dòng ${rowNumber}: areaM2 phải là số lớn hơn 0.`);
    if (!VALID_STATUSES.has(status)) throw new Error(`Dòng ${rowNumber}: status chỉ được là Available hoặc Maintenance.`);

    return {
      id,
      name,
      type,
      description: text('description', 1500),
      pricePerNight,
      capacity,
      areaM2,
      bedType: text('bedType', 100),
      amenities: text('amenities', 1000).split(';').map((item) => item.trim()).filter(Boolean),
      image: normalizeImagePath(text('image', 300)),
      status
    };
  }

  function normalizeImagePath(value) {
    if (!value) return './images/hero-resort.png';
    if (/^(https?:|data:)/i.test(value)) return value;
    return value.startsWith('./') ? value : `./${value.replace(/^\/+/, '')}`;
  }

  function getRoomById(id) {
    return roomCache.find((room) => room.id === String(id)) || null;
  }

  global.ResortData = Object.freeze({ loadRooms, getRoomById, EXCEL_PATH });
})(window);
