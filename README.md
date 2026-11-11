# meeting-room-calenda
/**
 * รหัส.gs (Google Apps Script API)
 * นำโค้ดนี้ไปแทนที่ใน Apps Script แล้วทำการ Deploy ใหม่เป็น Web App:
 * - Execute as: Me (บัญชีของคุณ)
 * - Who has access: Anyone (ทุกคน)
 */

const CALENDAR_IDS = {
  room1: "4209e3af7cc58769f2c43a438933189dfa640457dddd06e6d0882d103c5d8755@group.calendar.google.com",
  room2: "07c496984c50fec3caeeec5336865ac4db2f0d6d08163fb6b79fd6ede45ee3de@group.calendar.google.com",
  room3: "a24308c5d9b036774f1880fe4f426a9700b018c9ae29f0467fa85c07bca31fb2@group.calendar.google.com"
};

// จัดการคำขอแบบ POST จาก GitHub Pages
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    let result = {};

    if (action === 'getScheduleData') {
      result = getScheduleData(postData.date);
    } else if (action === 'getMonthOverview') {
      result = getMonthOverview(postData.year, postData.month);
    } else if (action === 'saveBooking') {
      result = saveBookingData(postData.data);
    } else {
      result = { status: "error", message: "ไม่พบ Action ที่ระบุ" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// เผื่อเปิดทดสอบผ่าน Browser ตรงๆ
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: "ready", message: "Meeting Room API is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 1. ดึงสถานะการจองตามวันที่
function getScheduleData(dateStr) {
  try {
    const parts = dateStr.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);

    const morningStart = new Date(year, month, day, 8, 30, 0).getTime();
    const morningEnd   = new Date(year, month, day, 12, 0, 0).getTime();
    const afternoonStart = new Date(year, month, day, 13, 0, 0).getTime();
    const afternoonEnd   = new Date(year, month, day, 16, 30, 0).getTime();
    
    const dayStart = new Date(year, month, day, 0, 0, 0);
    const dayEnd   = new Date(year, month, day, 23, 59, 59);

    let scheduleResult = {};

    for (let roomId in CALENDAR_IDS) {
      const cal = CalendarApp.getCalendarById(CALENDAR_IDS[roomId]);
      if (!cal) throw new Error("ไม่พบปฏิทิน: " + roomId);

      const events = cal.getEvents(dayStart, dayEnd);

      let morningStatus = { status: "available" };
      let afternoonStatus = { status: "available" };
      let fulldayStatus = { status: "available" };

      events.forEach(event => {
        let evStart = event.getStartTime().getTime();
        let evEnd   = event.getEndTime().getTime();
        let details = {
          title: event.getTitle(),
          description: event.getDescription() || "-"
        };

        // คาบเกี่ยวช่วงเช้า
        if (evStart < morningEnd && evEnd > morningStart) {
          morningStatus = { status: "booked", ...details };
        }
        // คาบเกี่ยวช่วงบ่าย
        if (evStart < afternoonEnd && evEnd > afternoonStart) {
          afternoonStatus = { status: "booked", ...details };
        }
      });

      // หากมีการจองเช้า หรือ บ่าย -> ตลอดวันจะไม่ว่าง
      if (morningStatus.status === "booked" || afternoonStatus.status === "booked") {
        let reason = [];
        if (morningStatus.status === "booked") reason.push(`ช่วงเช้า: ${morningStatus.title}`);
        if (afternoonStatus.status === "booked") reason.push(`ช่วงบ่าย: ${afternoonStatus.title}`);
        
        fulldayStatus = {
          status: "booked",
          title: "ติดจองบางช่วงเวลา",
          description: reason.join("\n")
        };
      }

      if (morningStatus.status === "booked" && afternoonStatus.status === "booked") {
        fulldayStatus = {
          status: "booked",
          title: morningStatus.title === afternoonStatus.title ? morningStatus.title : "มีการจองเต็มวัน",
          description: `ช่วงเช้า: ${morningStatus.title}\nช่วงบ่าย: ${afternoonStatus.title}`
        };
      }

      scheduleResult[roomId] = {
        morning: morningStatus,
        afternoon: afternoonStatus,
        fullday: fulldayStatus
      };
    }

    return { status: "success", data: scheduleResult };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// 2. ดึงภาพรวมคิวจองทั้งเดือน
function getMonthOverview(year, month) {
  try {
    const startDate = new Date(year, month, 1, 0, 0, 0);
    const endDate = new Date(year, month + 1, 0, 23, 59, 59);

    let monthData = {};

    for (let roomId in CALENDAR_IDS) {
      const cal = CalendarApp.getCalendarById(CALENDAR_IDS[roomId]);
      if (!cal) continue;

      const events = cal.getEvents(startDate, endDate);
      events.forEach(ev => {
        const evStart = ev.getStartTime();
        const y = evStart.getFullYear();
        const m = String(evStart.getMonth() + 1).padStart(2, '0');
        const d = String(evStart.getDate()).padStart(2, '0');
        const dStr = `${y}-${m}-${d}`;

        if (!monthData[dStr]) monthData[dStr] = {};
        monthData[dStr][roomId] = true;
      });
    }

    return { status: "success", data: monthData };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// 3. บันทึกข้อมูลการจองห้อง
function saveBookingData(data) {
  try {
    const cal = CalendarApp.getCalendarById(CALENDAR_IDS[data.room]);
    if (!cal) throw new Error("ไม่พบปฏิทินของห้องที่เลือก");

    const parts = data.date.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);

    let startTime, endTime;
    let slotLabel = "";

    if (data.slot === 'morning') {
      startTime = new Date(year, month, day, 8, 30, 0);
      endTime   = new Date(year, month, day, 12, 0, 0);
      slotLabel = "[ครึ่งวันเช้า]";
    } else if (data.slot === 'afternoon') {
      startTime = new Date(year, month, day, 13, 0, 0);
      endTime   = new Date(year, month, day, 16, 30, 0);
      slotLabel = "[ครึ่งวันบ่าย]";
    } else if (data.slot === 'fullday') {
      startTime = new Date(year, month, day, 8, 30, 0);
      endTime   = new Date(year, month, day, 16, 30, 0);
      slotLabel = "[ตลอดวัน]";
    }

    // ตรวจสอบห้องซ้ำ
    if (cal.getEvents(startTime, endTime).length > 0) {
      return { status: "error", message: "ช่วงเวลานี้ถูกจองไปแล้ว กรุณาเลือกช่วงเวลาใหม่" };
    }

    const desc = `ผู้จอง: ${data.booker}\nเบอร์ภายใน: ${data.phone}\nLINE UID: ${data.userId || '-'}\nรายละเอียดเพิ่มเติม: ${data.detail || '-'}`;
    const fullTitle = `${data.title} (${data.booker}) ${slotLabel}`;
    cal.createEvent(fullTitle, startTime, endTime, { description: desc });

    return { status: "success", message: "จองห้องประชุมสำเร็จเรียบร้อย!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}
