// -------------------- ユーザーによる設定が必要な項目 --------------------
const CALENDAR_IDS = [
  'ここに公開済みのカレンダーのIDを入れてね',
]; // 監視対象とするGoogleカレンダーのIDを配列で指定します。

const DISCORD_WEBHOOK_URL = 'ここにWebHookのURLを入れてね';

const HOLIDAY_CALENDAR_IDS = [
  'ja.japanese#holiday@group.v.calendar.google.com' // 日本の祝日カレンダーのID
  // 必要に応じて他の休日カレンダーIDを追加できます
];

class WorkSchedule {
  constructor(dayOfWeek, startTime, endTime, ignoreOnHoliday = false, startDate = null, endDate = null) {
    this.dayOfWeek = dayOfWeek; // 曜日 (0:日, 1:月, ..., 6:土)
    this.startTime = startTime;   // 開始時間 (時)
    this.endTime = endTime;     // 終了時間 (時)
    this.ignoreOnHoliday = ignoreOnHoliday; // 祝日の場合でも適用するかどうか (true: 適用する, false: 適用しない)
    this.startDate = startDate ? new Date(startDate) : null; // 開始日 (Date型)
    this.endDate = endDate ? new Date(endDate) : null;       // 終了日 (Date型)
  }

  isApplicable(date, isHoliday) {
    if (this.startDate && date < this.startDate) return false;
    if (this.endDate && date > this.endDate) return false;
    if (date.getDay() !== this.dayOfWeek) return false;
    return !isHoliday || this.ignoreOnHoliday;
  }
}

// 出勤等予定設定（予定がある時間帯）
const WORK_SCHEDULES = [
  new WorkSchedule(1, 9, 19, false), // 例: 毎週月曜日 9時〜19時 (祝日は休み)
  new WorkSchedule(2, 9, 19, false), // 例: 毎週火曜日 9時〜19時 (祝日は休み)
  new WorkSchedule(3, 9, 19, false), // 例: 毎週水曜日 9時〜19時 (祝日は休み)
  new WorkSchedule(4, 9, 19, false), // 例: 毎週木曜日 9時〜19時 (祝日は休み)
  new WorkSchedule(5, 9, 19, false), // 例: 毎週金曜日 9時〜19時 (祝日は休み)
  // 必要に応じて勤務スケジュールを追加・変更・削除できます
];

// -------------------- 通常は変更不要な項目 --------------------
const MESSAGE_ID_CACHE_KEY = 'discordWebhookMessageId'; // メッセージIDをキャッシュに保存するキー
const NUM_OF_MONTHS = 3; // 処理対象の月数（当月、翌月、翌々月）
const WEEK_DAYS = ['(日)', '(月)', '(火)', '(水)', '(木)', '(金)', '(土)'];
const DAYTIME_START_HOUR = 12;
const DAYTIME_END_HOUR = 18;
const NIGHTTIME_START_HOUR = 20;
const NIGHTTIME_END_HOUR = 25; // 翌日1時
const MIN_GAP_HOURS = 3;
const PROCESSING_INTERVAL_MINUTES = 30; // イベント判定の処理単位（分）


class DiscordNotifier {
  constructor(webhookUrl) {
    this.webhookUrl = webhookUrl;
  }

  postMessage(message = '', embeds = []) {
    const urlWithWait = this.webhookUrl + '?wait=true'; // ?wait=true を追加
    const payload = JSON.stringify({
      content: message,
      embeds: embeds,
      flags: 4096 // @silent
    });

    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: payload
    };

    try {
      const response = UrlFetchApp.fetch(urlWithWait, options);
      const responseJson = JSON.parse(response.getContentText());
      const messageId = responseJson.id;
      Logger.log(`Discordにメッセージを送信しました。メッセージID: ${messageId}`);
      return messageId;
    } catch (error) {
      Logger.log(`Discordへの送信に失敗しました: ${error}`);
      return null;
    }
  }

  editMessage(messageId, embeds) {
    const cache = CacheService.getScriptCache();

    // messageId が null の場合は編集処理をスキップして新規投稿
    if (messageId === null) {
      Logger.log('編集対象のメッセージIDがnullのため、新規投稿します。');
      const newMessageId = this.postMessage('', embeds);
      if (newMessageId) {
        cache.put(MESSAGE_ID_CACHE_KEY, newMessageId, 60 * 60 * 6); // IDを保存
        Logger.log(`新しいメッセージを送信しました。メッセージID: ${newMessageId}`);
        return newMessageId;
      } else {
        Logger.log('新しいメッセージの送信に失敗しました。');
        return null;
      }
    }

    const editUrl = `${this.webhookUrl}/messages/${messageId}`;
    const options = {
      'method': 'patch',
      'contentType': 'application/json',
      'payload': JSON.stringify({ embeds: embeds })
    };
    try {
      const editResponse = UrlFetchApp.fetch(editUrl, options);
      const responseCode = editResponse.getResponseCode();
      if (responseCode === 200) {
        const responseJson = JSON.parse(editResponse.getContentText());
        const newMessageId = responseJson.id;
        cache.put(MESSAGE_ID_CACHE_KEY, newMessageId, 60 * 60 * 6); // IDを更新
        Logger.log(`メッセージ (ID: ${messageId}) を編集しました。新しいメッセージID: ${newMessageId}`);
        return newMessageId;
      } else if (responseCode === 404) {
        Logger.log(`編集対象のメッセージ (ID: ${messageId}) は存在しませんでした。新規投稿します。`);
        const newMessageId = this.postMessage('', embeds);
        if (newMessageId) {
          cache.put(MESSAGE_ID_CACHE_KEY, newMessageId, 60 * 60 * 6); // IDを保存
          Logger.log(`新しいメッセージを送信しました。メッセージID: ${newMessageId}`);
          return newMessageId;
        } else {
          Logger.log('新しいメッセージの送信に失敗しました。');
          return null;
        }
      } else {
        Logger.log(`メッセージ (ID: ${messageId}) の編集に失敗しました。ステータスコード: ${responseCode}`);
        return null;
      }
    } catch (error) {
      Logger.log(`メッセージ (ID: ${messageId}) の編集中にエラーが発生しました: ${error}`);
      const newMessageId = this.postMessage('', embeds);
      if (newMessageId) {
        cache.put(MESSAGE_ID_CACHE_KEY, newMessageId, 60 * 60 * 6); // IDを保存
        Logger.log(`新しいメッセージを送信しました。メッセージID: ${newMessageId}`);
        return newMessageId;
      } else {
        Logger.log('新しいメッセージの送信に失敗しました。');
        return null;
      }
    }
  }
}

class CalendarHelper {
  constructor(calendarIds, holidayCalendarIds) {
    this.calendarIds = calendarIds;
    this.holidayCalendarIds = holidayCalendarIds;
    this.calendars = {};
    this.holidayCalendars = {};
  }

  getCalendar(calendarId) {
    if (!this.calendars[calendarId]) {
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) {
        this.calendars[calendarId] = calendar;
      } else {
        Logger.log(`カレンダーID '${calendarId}' は見つかりませんでした。`);
        return null;
      }
    }
    return this.calendars[calendarId];
  }

  getHolidayCalendar(calendarId) {
    if (!this.holidayCalendars[calendarId]) {
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) {
        this.holidayCalendars[calendarId] = calendar;
      } else {
        Logger.log(`祝日カレンダーID '${calendarId}' は見つかりませんでした。`);
        return null;
      }
    }
    return this.holidayCalendars[calendarId];
  }

  getEventsForRange(startTime, endTime) {
    let publicBusyEvents = [];
    this.calendarIds.forEach(calendarId => {
      const calendar = this.getCalendar(calendarId);
      if (calendar) {
        const events = calendar.getEvents(startTime, endTime);
        events.forEach(event => {
          publicBusyEvents.push(event);
        });
      }
    });
    return publicBusyEvents; // 公開設定が「予定あり」のイベントのみを返す
  }

  getHolidayDates(startDate, endDate) {
    let allHolidays = [];
    this.holidayCalendarIds.forEach(holidayCalendarId => {
      const calendar = this.getHolidayCalendar(holidayCalendarId);
      if (calendar) {
        const events = calendar.getEvents(startDate, endDate);
        allHolidays = allHolidays.concat(events.map(event => event.getStartTime()));
      }
    });
    return allHolidays;
  }
}


class EventStatusAnalyzer {
  constructor(daytimeStartHour, daytimeEndHour, nighttimeStartHour, nighttimeEndHour, minGapHours, weekDays, workSchedules, processingIntervalMinutes = 60) {
    this.daytimeStartHour = daytimeStartHour;
    this.daytimeEndHour = daytimeEndHour;
    this.nighttimeStartHour = nighttimeStartHour;
    this.nighttimeEndHour = nighttimeEndHour;
    this.minGapHours = minGapHours;
    this.weekDays = weekDays;
    this.workSchedules = workSchedules; // WorkSchedule クラスの配列
    this.processingIntervalMinutes = processingIntervalMinutes;
  }

  getDayStatuses(allEvents, date, holidays) {
    const year = date.getFullYear();
    const month = date.getMonth();
    const dayOfMonth = date.getDate();
    const dayOfWeek = date.getDay();
    const isHoliday = holidays.some(holiday =>
      holiday.getFullYear() === year &&
      holiday.getMonth() === month &&
      holiday.getDate() === dayOfMonth
    );
    const targetDateMidnight = new Date(year, month, dayOfMonth).getTime();

    // その日の終日イベントを取得
    const allDayEventsOnDate = allEvents.filter(event => {
      if (!event.isAllDayEvent()) {
        return false;
      }
      const eventStartDate = new Date(event.getStartTime().getFullYear(), event.getStartTime().getMonth(), event.getStartTime().getDate()).getTime();
      const eventEndDate = new Date(event.getEndTime().getFullYear(), event.getEndTime().getMonth(), event.getEndTime().getDate()).getTime();
      return eventEndDate > targetDateMidnight && eventStartDate <= targetDateMidnight;
    });
    let daytimeStatus = '◯'; // 昼間の初期ステータスは空き
    let nighttimeStatus = '◯'; // 夜間の初期ステータスは空き
    let hasAllDayEvent = allDayEventsOnDate.length > 0;
    let isDaytimeBusy = false;
    let isNighttimeBusy = false;

    this.workSchedules.forEach(schedule => {
      if (schedule.isApplicable(date, isHoliday)) {
        const scheduleStartTime = schedule.startTime;
        const scheduleEndTime = schedule.endTime;
        if ((scheduleStartTime < this.daytimeEndHour && scheduleEndTime > this.daytimeStartHour)) {
          isDaytimeBusy = true;
        }
        if ((scheduleStartTime < this.nighttimeEndHour && scheduleEndTime > this.nighttimeStartHour) ||
            (scheduleStartTime < 24 && scheduleEndTime > 24 && scheduleStartTime < this.nighttimeEndHour % 24) ||
            (scheduleStartTime < 24 && scheduleEndTime > 24 && scheduleEndTime > this.nighttimeStartHour)) {
          isNighttimeBusy = true;
        }
      }
    });

    if (hasAllDayEvent) {
      daytimeStatus = '✕';
      nighttimeStatus = '✕';
    } else {
      daytimeStatus = this.checkTimeRangeStatus(allEvents, date, this.daytimeStartHour, this.daytimeEndHour, isDaytimeBusy);
      nighttimeStatus = this.checkTimeRangeStatus(allEvents, date, this.nighttimeStartHour, this.nighttimeEndHour, isNighttimeBusy);
    }

    return [daytimeStatus, nighttimeStatus]; // 昼と夜のステータスを配列で返す
  }

  getEventsInTimeRange(allEvents, date, startHour, endHour) {
    const year = date.getFullYear();
    const month = date.getMonth();
    const dayOfMonth = date.getDate();
    const startTime = new Date(year, month, dayOfMonth, startHour, 0, 0, 0).getTime();
    let endTime;
    if (endHour > 23) {
      const nextDay = new Date(year, month, dayOfMonth + 1, endHour % 24, 0, 0, 0);
      endTime = nextDay.getTime();
    } else {
      endTime = new Date(year, month, dayOfMonth, endHour, 0, 0, 0).getTime();
    }

    return allEvents.filter(event => {
      const eventStartTime = event.getStartTime().getTime();
      const eventEndTime = event.getEndTime().getTime();
      return !event.isAllDayEvent() && eventStartTime < endTime && eventEndTime > startTime;
    });
  }

  checkTimeRangeStatus(allEvents, date, startHour, endHour, isWorkScheduleActive) {
    const events = this.getEventsInTimeRange(allEvents, date, startHour, endHour);
    const hasEventOrSchedule = isWorkScheduleActive || events.length > 0;

    if (!hasEventOrSchedule) {
      return '◯';
    }

    const allBusyIntervals = [];

    // カレンダーイベントの期間を busy interval として追加
    events.forEach(event => {
      allBusyIntervals.push({ start: event.getStartTime().getTime(), end: event.getEndTime().getTime() });
    });

    // 出勤時間を busy interval として追加
    if (isWorkScheduleActive) {
      const targetDate = new Date(date);
      const scheduleStartTime = new Date(targetDate);
      scheduleStartTime.setHours(this.workSchedules.find(schedule => schedule.isApplicable(targetDate, this.isHoliday(targetDate, []))).startTime, 0, 0, 0);

      const scheduleEndTime = new Date(targetDate);
      scheduleEndTime.setHours(this.workSchedules.find(schedule => schedule.isApplicable(targetDate, this.isHoliday(targetDate, []))).endTime, 0, 0, 0);

      // 翌日にまたがる勤務時間も考慮
      if (scheduleEndTime.getHours() < scheduleStartTime.getHours()) {
        scheduleEndTime.setDate(scheduleEndTime.getDate() + 1);
      }

      const periodStart = new Date(targetDate);
      periodStart.setHours(startHour, 0, 0, 0);
      const periodEnd = new Date(targetDate);
      periodEnd.setHours(endHour > 23 ? endHour % 24 : endHour, 0, 0, 0);
      if (endHour > 23) {
        periodEnd.setDate(periodEnd.getDate() + 1);
      }

      // 判定対象の時間帯に出勤時間帯が一部でも含まれていれば busy interval とする
      if (scheduleStartTime.getTime() < periodEnd.getTime() && scheduleEndTime.getTime() > periodStart.getTime()) {
        allBusyIntervals.push({ start: scheduleStartTime.getTime(), end: scheduleEndTime.getTime() });
      }
    }

    if (allBusyIntervals.length === 0) {
      return '◯';
    }

    allBusyIntervals.sort((a, b) => a.start - b.start);

    const periodStart = new Date(date);
    periodStart.setHours(startHour, 0, 0, 0);
    const periodEnd = new Date(date);
    periodEnd.setHours(endHour > 23 ? endHour % 24 : endHour, 0, 0, 0);
    if (endHour > 23) {
      periodEnd.setDate(periodEnd.getDate() + 1);
    }
    let currentTime = periodStart.getTime();
    let hasLongGap = false;

    for (const interval of allBusyIntervals) {
      const gap = (interval.start - currentTime) / (1000 * 60 * 60);
      if (gap >= this.minGapHours) {
        hasLongGap = true;
        break;
      }
      currentTime = Math.max(currentTime, interval.end);
    }

    const lastIntervalEnd = allBusyIntervals.length > 0 ? allBusyIntervals[allBusyIntervals.length - 1].end : periodStart.getTime();
    const lastGap = (periodEnd.getTime() - lastIntervalEnd) / (1000 * 60 * 60);
    if (lastGap >= this.minGapHours) {
      hasLongGap = true;
    }

    return hasLongGap ? '△' : '✕';
  }

  isHoliday(date, holidays) {
    const year = date.getFullYear();
    const month = date.getMonth();
    const dayOfMonth = date.getDate();
    return holidays.some(holiday =>
      holiday.getFullYear() === year &&
      holiday.getMonth() === month &&
      holiday.getDate() === dayOfMonth
    );
  }
}

/**
 * Googleカレンダーのイベント情報を取得し、指定されたDiscordのWebhook URLへ送信します。
 */
function sendCalendarEventsToDiscord() {
  const discordNotifier = new DiscordNotifier(DISCORD_WEBHOOK_URL);
  const calendarHelper = new CalendarHelper(CALENDAR_IDS, HOLIDAY_CALENDAR_IDS);
  const eventStatusAnalyzer = new EventStatusAnalyzer(
    DAYTIME_START_HOUR,
    DAYTIME_END_HOUR,
    NIGHTTIME_START_HOUR,
    NIGHTTIME_END_HOUR,
    MIN_GAP_HOURS,
    WEEK_DAYS,
    WORK_SCHEDULES,
    PROCESSING_INTERVAL_MINUTES // 処理単位をコンストラクタに渡す
  );

  // 現在の日付情報を取得
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  const timestamp = now.toISOString();

  // 処理対象の月の開始日と終了日を計算
  const firstDayOfCurrentMonth = new Date(currentYear, currentMonth, 1);
  const lastDayOfTargetMonth = new Date(currentYear, currentMonth + NUM_OF_MONTHS, 0);

  // 祝日と休日の日付をまとめて取得
  const allHolidaysInRange = calendarHelper.getHolidayDates(firstDayOfCurrentMonth, lastDayOfTargetMonth);

  const monthInfos = [];

  for (let i = 0; i < NUM_OF_MONTHS; i++) {
    const targetMonth = (currentMonth + i) % 12;
    let targetYear = currentYear;
    if (currentMonth + i >= 12) {
      targetYear++;
    }

    const startDate = new Date(targetYear, targetMonth, 1, 0, 0, 0, 0);
    const endDate = new Date(targetYear, targetMonth + 1, 0, 23, 59, 59, 999);
    const events = calendarHelper.getEventsForRange(startDate, endDate);
    const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
    let monthText = '```\n';

    monthText += '   　 　昼 夜\n';

    for (let day = 1; day <= daysInMonth; day++) {
      const date = new Date(targetYear, targetMonth, day);
      const dayOfWeek = WEEK_DAYS[date.getDay()];
      // 当該日の祝日・休日のみをフィルタリング
      const holidaysOnDate = allHolidaysInRange.filter(holiday =>
        holiday.getFullYear() === targetYear &&
        holiday.getMonth() === targetMonth && holiday.getDate() === day
      );
      const [daytimeStatus, nighttimeStatus] = eventStatusAnalyzer.getDayStatuses(events, date, holidaysOnDate);
      const dayStr = Utilities.formatString('%2d', day);
      monthText += `${dayStr}${dayOfWeek}：${daytimeStatus} ${nighttimeStatus}\n`;
    }
    monthText += '```';

    monthInfos.push({
      name: `${targetMonth + 1}月`,
      value: daysInMonth > 0 ? monthText : '情報なし',
      inline: true
    });
  }

  let embeds = [];
  embeds.push({
    description: '-# ◯：空き、△：要相談、✕：予定あり',
    fields: monthInfos,
    timestamp: timestamp
  });

  // CacheService の初期化
  const cache = CacheService.getScriptCache();
  const prevMessageId = cache.get(MESSAGE_ID_CACHE_KEY);

  // Discordのメッセージを編集または新規投稿
  const messageId = discordNotifier.editMessage(prevMessageId, embeds);
  if (messageId == null) {
    cache.put(MESSAGE_ID_CACHE_KEY, prevMessageId, 60 * 60 * 6); // IDを保存
    return;
  }
  else {
    cache.put(MESSAGE_ID_CACHE_KEY, messageId, 60 * 60 * 6); // IDを保存
  }
}
