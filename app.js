/**
 * Điều khiển giao diện website tĩnh.
 * Không gọi API và không ghi dữ liệu cá nhân vào Excel hay localStorage.
 */
'use strict';

const CONFIG = {
  resortEmail: 'hello@annhienresort.vn',
  fallbackImage: './images/hero-resort.png',
  currency: new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND', maximumFractionDigits: 0 })
};

const state = { rooms: [], visibleRooms: [], selectedRoom: null, search: {}, lastFocused: null };
const elements = {};

document.addEventListener('DOMContentLoaded', initializeWebsite);

async function initializeWebsite() {
  cacheElements();
  setDefaultDates();
  bindEvents();
  observeReveals();
  restoreNonSensitiveDraft();
  elements.currentYear.textContent = new Date().getFullYear();
  await loadExcelRooms();
}

function cacheElements() {
  [
    'siteHeader', 'menuToggle', 'mainNav', 'searchForm', 'searchCheckIn', 'searchCheckOut', 'searchGuests',
    'searchMessage', 'filterType', 'filterPrice', 'filterGuests', 'filterStatus', 'sortRooms', 'resetFilters',
    'roomsStatus', 'roomsGrid', 'emptyState', 'emptyReset', 'roomModal', 'roomModalContent', 'bookingModal',
    'bookingForm', 'bookingRoomId', 'bookingCheckIn', 'bookingCheckOut', 'bookingGuests', 'bookingMessage',
    'bookingSummary', 'toast', 'currentYear'
  ].forEach((id) => { elements[id] = document.getElementById(id); });
}

async function loadExcelRooms() {
  showRoomsLoading();
  try {
    state.rooms = await window.ResortData.loadRooms();
    populateTypes();
    applyFilters();
  } catch (error) {
    showRoomsError(error.message);
    console.error('Lỗi dữ liệu phòng:', error);
  }
}

function showRoomsLoading() {
  elements.roomsStatus.hidden = false;
  elements.roomsStatus.innerHTML = '<span class="spinner dark"></span> Đang đọc dữ liệu từ rooms.xlsx...';
  elements.roomsGrid.innerHTML = '';
  elements.emptyState.hidden = true;
}

function showRoomsError(message) {
  elements.roomsStatus.hidden = false;
  elements.roomsStatus.innerHTML = `<p><strong>Chưa thể tải phòng nghỉ.</strong><br>${escapeHtml(message)}</p><button class="button" type="button" id="retryExcel">Thử lại</button>`;
  document.getElementById('retryExcel').addEventListener('click', loadExcelRooms);
}

function bindEvents() {
  window.addEventListener('scroll', () => elements.siteHeader.classList.toggle('scrolled', window.scrollY > 30), { passive: true });
  elements.menuToggle.addEventListener('click', () => {
    const isOpen = elements.mainNav.classList.toggle('open');
    elements.menuToggle.setAttribute('aria-expanded', String(isOpen));
  });
  elements.mainNav.addEventListener('click', () => {
    elements.mainNav.classList.remove('open');
    elements.menuToggle.setAttribute('aria-expanded', 'false');
  });
  elements.searchForm.addEventListener('submit', handleSearch);
  elements.searchCheckIn.addEventListener('change', syncSearchDates);
  [elements.filterType, elements.filterPrice, elements.filterGuests, elements.filterStatus, elements.sortRooms]
    .forEach((node) => node.addEventListener('change', applyFilters));
  elements.resetFilters.addEventListener('click', resetFilters);
  elements.emptyReset.addEventListener('click', resetFilters);
  elements.roomsGrid.addEventListener('click', handleRoomAction);
  document.querySelectorAll('[data-close-modal]').forEach((node) => node.addEventListener('click', () => closeModal(elements.roomModal)));
  document.querySelectorAll('[data-close-booking]').forEach((node) => node.addEventListener('click', () => closeModal(elements.bookingModal)));
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      closeModal(elements.roomModal);
      closeModal(elements.bookingModal);
    }
  });
  elements.bookingForm.addEventListener('submit', prepareBookingRequest);
  [elements.bookingCheckIn, elements.bookingCheckOut, elements.bookingGuests]
    .forEach((node) => node.addEventListener('input', updateBookingSummary));
}

function setDefaultDates() {
  const today = new Date();
  const tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
  const nextDay = new Date(today); nextDay.setDate(today.getDate() + 2);
  const min = toLocalDate(today);
  elements.searchCheckIn.min = min;
  elements.searchCheckOut.min = toLocalDate(tomorrow);
  elements.searchCheckIn.value = toLocalDate(tomorrow);
  elements.searchCheckOut.value = toLocalDate(nextDay);
  elements.bookingCheckIn.min = min;
  elements.bookingCheckOut.min = toLocalDate(tomorrow);
}

function handleSearch(event) {
  event.preventDefault();
  const checkIn = elements.searchCheckIn.value;
  const checkOut = elements.searchCheckOut.value;
  if (!checkIn || !checkOut || checkOut <= checkIn) {
    elements.searchMessage.textContent = 'Ngày trả phòng phải sau ngày nhận phòng.';
    return;
  }
  elements.searchMessage.textContent = '';
  state.search = { checkIn, checkOut, guests: Number(elements.searchGuests.value) };
  elements.filterGuests.value = String(state.search.guests) === '1' ? '' : String(state.search.guests);
  applyFilters();
  document.getElementById('phong-nghi').scrollIntoView({ behavior: 'smooth' });
  showToast('Đã lọc phòng theo sức chứa. Resort sẽ xác nhận tình trạng thực tế sau khi nhận yêu cầu.');
}

function syncSearchDates() {
  if (!elements.searchCheckIn.value) return;
  const nextDay = addDays(elements.searchCheckIn.value, 1);
  elements.searchCheckOut.min = nextDay;
  if (!elements.searchCheckOut.value || elements.searchCheckOut.value <= elements.searchCheckIn.value) {
    elements.searchCheckOut.value = nextDay;
  }
}

function populateTypes() {
  const types = [...new Set(state.rooms.map((room) => room.type))].sort((a, b) => a.localeCompare(b, 'vi'));
  elements.filterType.innerHTML = '<option value="">Tất cả</option>' +
    types.map((type) => `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`).join('');
}

function applyFilters() {
  let rooms = [...state.rooms];
  const type = elements.filterType.value;
  const maxPrice = Number(elements.filterPrice.value);
  const guests = Number(elements.filterGuests.value);
  const status = elements.filterStatus.value;
  if (type) rooms = rooms.filter((room) => room.type === type);
  if (maxPrice) rooms = rooms.filter((room) => room.pricePerNight <= maxPrice);
  if (guests) rooms = rooms.filter((room) => room.capacity >= guests);
  if (status) rooms = rooms.filter((room) => room.status === status);
  if (elements.sortRooms.value === 'asc') rooms.sort((a, b) => a.pricePerNight - b.pricePerNight);
  if (elements.sortRooms.value === 'desc') rooms.sort((a, b) => b.pricePerNight - a.pricePerNight);
  state.visibleRooms = rooms;
  renderRooms();
}

function renderRooms() {
  elements.roomsStatus.hidden = true;
  elements.emptyState.hidden = state.visibleRooms.length > 0;
  elements.roomsGrid.innerHTML = state.visibleRooms.map((room) => `
    <article class="room-card">
      <div class="room-image">
        <img src="${escapeHtml(room.image)}" alt="${escapeHtml(room.name)}" loading="lazy" onerror="this.onerror=null;this.src='${CONFIG.fallbackImage}'">
        <span class="room-status ${room.status === 'Available' ? '' : 'maintenance'}">${room.status === 'Available' ? 'Đang nhận yêu cầu' : 'Đang bảo trì'}</span>
      </div>
      <div class="room-body">
        <div class="room-meta"><span>${escapeHtml(room.type)}</span><span>${room.areaM2} m² · Tối đa ${room.capacity} khách</span></div>
        <h3>${escapeHtml(room.name)}</h3>
        <p>${escapeHtml(room.description)}</p>
        <div class="room-amenities">${room.amenities.slice(0, 3).map((item) => `<span>${escapeHtml(item)}</span>`).join('')}</div>
        <div class="room-footer">
          <div class="room-price">${formatPrice(room.pricePerNight)}<small>mỗi đêm</small></div>
          <div class="room-actions">
            <button type="button" data-action="detail" data-room-id="${escapeHtml(room.id)}">Xem chi tiết</button>
            <button class="book-button" type="button" data-action="book" data-room-id="${escapeHtml(room.id)}" ${room.status !== 'Available' ? 'disabled' : ''}>Gửi yêu cầu</button>
          </div>
        </div>
      </div>
    </article>`).join('');
}

function resetFilters() {
  elements.filterType.value = '';
  elements.filterPrice.value = '';
  elements.filterGuests.value = '';
  elements.filterStatus.value = '';
  elements.sortRooms.value = 'default';
  applyFilters();
}

function handleRoomAction(event) {
  const button = event.target.closest('[data-room-id]');
  if (!button) return;
  const room = window.ResortData.getRoomById(button.dataset.roomId);
  if (!room) return;
  if (button.dataset.action === 'detail') openRoomDetail(room);
  else openBooking(room);
}

function openRoomDetail(room) {
  state.selectedRoom = room;
  elements.roomModalContent.innerHTML = `
    <article class="room-detail">
      <div class="room-detail-image"><img src="${escapeHtml(room.image)}" alt="${escapeHtml(room.name)}" onerror="this.onerror=null;this.src='${CONFIG.fallbackImage}'"></div>
      <div class="room-detail-content">
        <p class="eyebrow">${escapeHtml(room.type)}</p><h2 id="roomModalTitle">${escapeHtml(room.name)}</h2><p>${escapeHtml(room.description)}</p>
        <div class="detail-facts"><div><span>Diện tích</span><strong>${room.areaM2} m²</strong></div><div><span>Sức chứa</span><strong>${room.capacity} khách</strong></div><div><span>Giường</span><strong>${escapeHtml(room.bedType)}</strong></div></div>
        <p class="eyebrow">Tiện ích trong phòng</p><ul class="detail-amenities">${room.amenities.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
        <div class="request-note">Đây là website tĩnh. Tình trạng phòng sẽ được resort kiểm tra và xác nhận sau khi nhận yêu cầu.</div>
        <div class="detail-price"><div class="room-price">${formatPrice(room.pricePerNight)}<small>mỗi đêm</small></div><button class="button" type="button" id="modalBookButton" ${room.status !== 'Available' ? 'disabled' : ''}>${room.status === 'Available' ? 'Gửi yêu cầu' : 'Đang bảo trì'}</button></div>
      </div>
    </article>`;
  const bookButton = document.getElementById('modalBookButton');
  if (!bookButton.disabled) bookButton.addEventListener('click', () => { closeModal(elements.roomModal); openBooking(room); });
  openModal(elements.roomModal);
}

function openBooking(room) {
  if (room.status !== 'Available') {
    showToast('Phòng này đang bảo trì và chưa nhận yêu cầu.');
    return;
  }
  state.selectedRoom = room;
  removePreparedRequest();
  elements.bookingForm.reset();
  clearFieldErrors();
  elements.bookingRoomId.value = room.id;
  elements.bookingCheckIn.value = state.search.checkIn || elements.searchCheckIn.value;
  elements.bookingCheckOut.value = state.search.checkOut || elements.searchCheckOut.value;
  elements.bookingGuests.value = state.search.guests || Math.min(2, room.capacity);
  elements.bookingGuests.max = room.capacity;
  elements.bookingMessage.textContent = '';
  updateBookingSummary();
  openModal(elements.bookingModal);
}

function updateBookingSummary() {
  const room = state.selectedRoom;
  if (!room) return;
  const checkIn = elements.bookingCheckIn.value;
  const checkOut = elements.bookingCheckOut.value;
  const nights = checkIn && checkOut && checkOut > checkIn ? dateDifference(checkIn, checkOut) : 0;
  elements.bookingSummary.innerHTML = `
    <img src="${escapeHtml(room.image)}" alt="" onerror="this.onerror=null;this.src='${CONFIG.fallbackImage}'">
    <p class="eyebrow light">Chi phí dự kiến</p><h3>${escapeHtml(room.name)}</h3>
    <div class="summary-line"><span>Nhận phòng</span><strong>${checkIn ? formatDate(checkIn) : '—'}</strong></div>
    <div class="summary-line"><span>Trả phòng</span><strong>${checkOut ? formatDate(checkOut) : '—'}</strong></div>
    <div class="summary-line"><span>Số đêm</span><strong>${nights || '—'}</strong></div>
    <div class="summary-line"><span>Đơn giá</span><strong>${formatPrice(room.pricePerNight)}</strong></div>
    <div class="summary-total">${nights ? formatPrice(nights * room.pricePerNight) : 'Chọn ngày hợp lệ'}</div>
    <div class="request-note">Giá và tình trạng phòng chưa được xác nhận theo thời gian thực.</div>`;
  saveNonSensitiveDraft();
}

function prepareBookingRequest(event) {
  event.preventDefault();
  clearFieldErrors();
  if (!validateBookingForm()) return;

  const values = Object.fromEntries(new FormData(elements.bookingForm).entries());
  const room = state.selectedRoom;
  const nights = dateDifference(values.checkIn, values.checkOut);
  const totalPrice = nights * room.pricePerNight;
  const requestCode = createRequestCode();
  const emailBody = buildEmailBody({ ...values, nights, totalPrice, requestCode, room });
  renderPreparedRequest({ ...values, nights, totalPrice, requestCode, room, emailBody });
}

function validateBookingForm() {
  const values = Object.fromEntries(new FormData(elements.bookingForm).entries());
  const errors = {};
  const today = toLocalDate(new Date());
  if (values.guestName.trim().length < 2) errors.guestName = 'Vui lòng nhập họ tên.';
  if (!/^[0-9+().\s-]{8,24}$/.test(values.phone)) errors.phone = 'Số điện thoại chưa hợp lệ.';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(values.email)) errors.email = 'Email chưa hợp lệ.';
  if (!values.checkIn || values.checkIn < today) errors.checkIn = 'Ngày nhận không được ở quá khứ.';
  if (!values.checkOut || values.checkOut <= values.checkIn) errors.checkOut = 'Ngày trả phải sau ngày nhận.';
  if (!values.guests || Number(values.guests) < 1 || Number(values.guests) > state.selectedRoom.capacity) errors.guests = `Phòng nhận tối đa ${state.selectedRoom.capacity} khách.`;
  if (state.selectedRoom.status !== 'Available') errors.roomId = 'Phòng đang bảo trì.';

  Object.entries(errors).forEach(([name, message]) => {
    const input = elements.bookingForm.elements[name];
    if (!input) return;
    const field = input.closest('.field');
    if (field) { field.classList.add('invalid'); field.querySelector('.field-error').textContent = message; }
  });
  const firstInvalid = elements.bookingForm.querySelector('.invalid input, .invalid textarea');
  if (firstInvalid) firstInvalid.focus();
  else if (errors.roomId) elements.bookingMessage.textContent = errors.roomId;
  return Object.keys(errors).length === 0;
}

function buildEmailBody(request) {
  return [
    'YÊU CẦU KIỂM TRA VÀ ĐẶT PHÒNG',
    `Mã yêu cầu: ${request.requestCode}`,
    '',
    `Khách hàng: ${request.guestName}`,
    `Điện thoại: ${request.phone}`,
    `Email: ${request.email}`,
    `Phòng quan tâm: ${request.room.name} (${request.room.id})`,
    `Nhận phòng: ${request.checkIn}`,
    `Trả phòng: ${request.checkOut}`,
    `Số đêm: ${request.nights}`,
    `Số khách: ${request.guests}`,
    `Đơn giá tham khảo: ${formatPrice(request.room.pricePerNight)}`,
    `Tổng dự kiến: ${formatPrice(request.totalPrice)}`,
    `Ghi chú: ${request.specialRequest || 'Không có'}`,
    '',
    'Vui lòng kiểm tra tình trạng phòng và liên hệ lại để xác nhận.',
    'Lưu ý: Đây là yêu cầu được chuẩn bị từ website tĩnh, chưa phải booking đã xác nhận.'
  ].join('\n');
}

function renderPreparedRequest(request) {
  const panel = elements.bookingModal.querySelector('.booking-panel');
  panel.querySelector('.booking-form-wrap').hidden = true;
  elements.bookingSummary.hidden = true;
  const success = document.createElement('div');
  success.className = 'booking-success prepared-request';
  success.innerHTML = `
    <div class="success-mark">✓</div>
    <p class="eyebrow">Yêu cầu đã được chuẩn bị</p>
    <h2>Kiểm tra lại trước khi gửi email.</h2>
    <p>Mã tham chiếu của bạn:</p><div class="booking-code">${escapeHtml(request.requestCode)}</div>
    <p><strong>${escapeHtml(request.room.name)}</strong><br>${formatDate(request.checkIn)} — ${formatDate(request.checkOut)} · ${request.nights} đêm<br>Chi phí dự kiến: ${formatPrice(request.totalPrice)}</p>
    <div class="request-note">Website chưa gửi hoặc xác nhận booking. Nút email chỉ mở ứng dụng email với nội dung đã điền; bạn vẫn cần bấm Gửi trong ứng dụng email.</div>
    <div class="request-actions">
      <a class="button" id="emailRequest" href="mailto:${encodeURIComponent(CONFIG.resortEmail)}?subject=${encodeURIComponent(`Yêu cầu đặt phòng ${request.requestCode}`)}&body=${encodeURIComponent(request.emailBody)}">Mở ứng dụng email</a>
      <button class="button button-secondary" type="button" id="copyRequest">Sao chép nội dung</button>
      <button class="button button-secondary" type="button" id="editRequest">Chỉnh sửa</button>
    </div>`;
  panel.appendChild(success);
  document.getElementById('copyRequest').addEventListener('click', () => copyRequestText(request.emailBody));
  document.getElementById('editRequest').addEventListener('click', removePreparedRequest);
  clearDraft();
}

async function copyRequestText(text) {
  try {
    await navigator.clipboard.writeText(text);
    showToast('Đã sao chép nội dung yêu cầu.');
  } catch (error) {
    showToast('Trình duyệt không cho phép sao chép tự động. Hãy dùng nút email.');
  }
}

function removePreparedRequest() {
  const success = elements.bookingModal.querySelector('.prepared-request');
  if (success) success.remove();
  const formWrap = elements.bookingModal.querySelector('.booking-form-wrap');
  if (formWrap) formWrap.hidden = false;
  elements.bookingSummary.hidden = false;
}

function createRequestCode() {
  if (crypto.randomUUID) return `AN-${crypto.randomUUID().split('-')[0].toUpperCase()}`;
  return `AN-${Date.now().toString(36).toUpperCase()}-${Math.random().toString(36).slice(2, 6).toUpperCase()}`;
}

function saveNonSensitiveDraft() {
  if (!state.selectedRoom) return;
  const draft = {
    roomId: state.selectedRoom.id,
    checkIn: elements.bookingCheckIn.value,
    checkOut: elements.bookingCheckOut.value,
    guests: elements.bookingGuests.value
  };
  try { localStorage.setItem('anNhienRequestDraft', JSON.stringify(draft)); } catch (error) { /* Trình duyệt có thể chặn storage. */ }
}

function restoreNonSensitiveDraft() {
  try {
    const draft = JSON.parse(localStorage.getItem('anNhienRequestDraft') || 'null');
    if (!draft) return;
    if (draft.checkIn) elements.searchCheckIn.value = draft.checkIn;
    if (draft.checkOut) elements.searchCheckOut.value = draft.checkOut;
    if (draft.guests) elements.searchGuests.value = draft.guests;
  } catch (error) { clearDraft(); }
}

function clearDraft() {
  try { localStorage.removeItem('anNhienRequestDraft'); } catch (error) { /* Không ảnh hưởng trải nghiệm chính. */ }
}

function openModal(modal) {
  state.lastFocused = document.activeElement;
  modal.hidden = false;
  document.body.classList.add('modal-open');
  requestAnimationFrame(() => modal.querySelector('.modal-panel').focus());
}

function closeModal(modal) {
  if (!modal || modal.hidden) return;
  modal.hidden = true;
  if (modal === elements.bookingModal) removePreparedRequest();
  if (elements.roomModal.hidden && elements.bookingModal.hidden) document.body.classList.remove('modal-open');
  if (state.lastFocused && document.contains(state.lastFocused)) state.lastFocused.focus();
}

function clearFieldErrors() {
  elements.bookingForm.querySelectorAll('.field').forEach((field) => {
    field.classList.remove('invalid');
    field.querySelector('.field-error').textContent = '';
  });
}

function observeReveals() {
  if (!('IntersectionObserver' in window)) {
    document.querySelectorAll('.reveal').forEach((node) => node.classList.add('visible'));
    return;
  }
  const observer = new IntersectionObserver((entries) => entries.forEach((entry) => {
    if (entry.isIntersecting) { entry.target.classList.add('visible'); observer.unobserve(entry.target); }
  }), { threshold: 0.12 });
  document.querySelectorAll('.reveal').forEach((node) => observer.observe(node));
}

function showToast(message) {
  elements.toast.textContent = message;
  elements.toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => elements.toast.classList.remove('show'), 3500);
}

function formatPrice(value) { return CONFIG.currency.format(value); }
function formatDate(value) { return new Intl.DateTimeFormat('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(new Date(`${value}T00:00:00`)); }
function toLocalDate(date) { const offset = date.getTimezoneOffset(); return new Date(date.getTime() - offset * 60000).toISOString().slice(0, 10); }
function addDays(value, days) { const date = new Date(`${value}T00:00:00`); date.setDate(date.getDate() + days); return toLocalDate(date); }
function dateDifference(from, to) { return Math.round((new Date(`${to}T00:00:00`) - new Date(`${from}T00:00:00`)) / 86400000); }
function escapeHtml(value) { return String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[char]); }
