const APP = {
  sheets: {
    students: 'Students',
    sessions: 'AuthSessions',
    races: 'RaceSessions',
    progress: 'RaceProgress',
    words: 'Words',
    results: 'Results',
    passages: 'Passages',
    achievements: 'Achievements',
    settings: 'Settings'
  },
  accuracyThreshold: 90,
  sessionHours: 8,
  raceDurations: [30, 60, 120, 300],
  kdfIterations: 60000,
  sessionPollMs: 3000
};

const HEADERS = {
  Students: ['studentId', 'username', 'displayName', 'section', 'status', 'salt', 'passwordHash', 'role', 'createdAt'],
  AuthSessions: ['tokenHash', 'studentId', 'expiresAt', 'revoked'],
  RaceSessions: ['raceId', 'roomCode', 'mode', 'duration', 'difficulty', 'passageId', 'status', 'hostId', 'participantIds', 'createdAt', 'startedAt', 'expiresAt'],
  RaceProgress: ['raceId', 'studentId', 'wpm', 'accuracy', 'progress', 'updatedAt'],
  Words: ['word', 'category', 'active'],
  Results: ['resultId', 'raceId', 'studentId', 'mode', 'duration', 'rawWpm', 'validWpm', 'accuracy', 'correctWords', 'invalidWords', 'errors', 'correctChars', 'typedChars', 'elapsedSeconds', 'eligible', 'completedAt', 'practice'],
  Passages: ['passageId', 'title', 'text', 'difficulty', 'type', 'active'],
  Achievements: ['studentId', 'achievementId', 'awardedAt'],
  Settings: ['key', 'value']
};

const STARTER_PASSAGES = [
  ['easy-1', 'A quiet morning', 'The sun rose above the hills. Birds sang near the old stone bridge, and a cool breeze moved through the trees.', 'easy', 'copy', true],
  ['easy-2', 'A good habit', 'Reading for a few minutes each day can teach us new ideas. A small habit can lead to lasting progress.', 'easy', 'copy', true],
  ['medium-1', 'The observatory', 'At dusk, the observatory opened its doors to the public. Visitors gathered beneath the dome to watch the stars appear.', 'medium', 'copy', true],
  ['medium-2', 'A careful plan', 'Before beginning the journey, the team checked its maps, packed enough water, and agreed to stay together on the trail.', 'medium', 'copy', true],
  ['hard-1', 'A changing coastline', 'Along the rugged coastline, patient researchers recorded subtle changes in the cliffs, where wind and waves had shaped the rock for centuries.', 'hard', 'copy', true],
  ['hard-2', 'A useful discovery', 'Although the first experiment produced unexpected results, careful observation helped the students revise their method and reach a reliable conclusion.', 'hard', 'copy', true]
];

const STARTER_WORDS = (
  'a about above across act after again against age ago air all also always among an and animal another answer any are around as ask at away back be because become before begin best better between big bird black body book both boy bring build but by call came can car carry change children city close cold come common complete country course create day did different do does down during each earth early easy end enough even every example eye face family far fast father feel few find first follow food for form found friend from full game gave get give go good great group grow had hand hard has have he head hear help her here high him his home house how however I idea if important in into is it its just keep kind know land large last later learn leave left length let life light like line list little live long look made make man many map mark may me mean measure might mile mind minute money more morning most mother move much music must name near never new next night no not number object of off often old on once one only open or other our out over own page paper part people person place plan plant play point possible power problem put question quick quite ran read real right river road room round run said same saw say school sea second see seem sentence set several she short show side simple since small so some something sound spell start state still story study such take talk tell than that the their them then there these they thing think this those thought three through time to together too took toward tree turn two under until up us use usual very voice wait walk want warm was watch water way we well went were what when where which while white who why will wind with without word work world would write year you young your zero islet keyboard practice race correct accurate accuracy speed compete classroom challenge computer type typed typing progress score result student teacher daily weekly achievement record personal improve improve improvement precision race winner room code section test happy family forest mountain river country language people planet science history careful curious brave honest patient thoughtful brilliant quiet river sky cloud star light energy nature garden flower golden silver red blue green black brown person children children’s teacher student school laptop keyboard sentence paragraph punctuation apostrophe hyphen today tomorrow yesterday monday tuesday wednesday thursday friday saturday sunday bangladesh dhaka london paris tokyo canada india nepal japan germany africa america australia pacific atlantic earth mars venus europe asia samira rahim salma imran arif'
).split(/\s+/);

function doGet() {
  return jsonOutput_({ ok: true, service: 'WPM RACER API' });
}

function doPost(e) {
  try {
    const body = e && e.postData ? String(e.postData.contents || '') : '';
    if (!body || body.length > 65536) throw new Error('Request body is missing or too large.');
    const envelope = JSON.parse(body);
    if (!envelope || typeof envelope !== 'object' || Array.isArray(envelope)) {
      throw new Error('Invalid proxy request.');
    }
    const expectedSecret = PropertiesService.getScriptProperties().getProperty('WPM_PROXY_SECRET');
    if (!expectedSecret || !constantTimeEqual_(String(envelope.proxySecret || ''), expectedSecret)) {
      throw new Error('Proxy authentication failed.');
    }
    return jsonOutput_(apiRequest(envelope.apiRequest));
  } catch (error) {
    return jsonOutput_({ ok: false, error: error && error.message ? error.message : 'Invalid API request.' });
  }
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok_(data) {
  return { ok: true, data: data };
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
      sheet.setFrozenRows(1);
    }
  });
  const passages = ss.getSheetByName(APP.sheets.passages);
  if (passages.getLastRow() < 2) {
    passages.getRange(2, 1, STARTER_PASSAGES.length, HEADERS.Passages.length).setValues(STARTER_PASSAGES);
  }
  const words = ss.getSheetByName(APP.sheets.words);
  if (words.getLastRow() < 2) {
    const seedWords = Array.from(new Set(STARTER_WORDS.map(normalizeWord_))).filter(Boolean);
    words.getRange(2, 1, seedWords.length, 3).setValues(seedWords.map(function(word) {
      return [word, 'common', true];
    }));
  }
  const settings = ss.getSheetByName(APP.sheets.settings);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 3, 2).setValues([
      ['minimumAccuracy', String(APP.accuracyThreshold)],
      ['timezone', 'Asia/Dhaka'],
      ['pollMs', String(APP.sessionPollMs)]
    ]);
  }
  return 'WPM RACER sheets are ready. Add temporary admin credentials in Script Properties and run provisionAdminAccount.';
}

function createAdminAccount(username, password, displayName) {
  if (!username || !password || password.length < 12) {
    throw new Error('Provide a username and a password of at least 12 characters.');
  }
  return withLock_(function() {
    const sheet = sheet_(APP.sheets.students);
    const rows = records_(sheet);
    if (rows.some(function(row) { return String(row.username).toLowerCase() === String(username).toLowerCase(); })) {
      throw new Error('That username already exists.');
    }
    const salt = Utilities.getUuid() + Utilities.getUuid();
    const hash = passwordHash_(password, salt);
    sheet.appendRow([
      'ADMIN-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      String(username).trim(),
      String(displayName || 'Teacher').trim(),
      '',
      'active',
      salt,
      hash,
      'admin',
      new Date().toISOString()
    ]);
    return 'Admin account created. Keep the password private and do not put it in a sheet.';
  });
}

function provisionAdminAccount() {
  const properties = PropertiesService.getScriptProperties();
  const username = properties.getProperty('WPM_ADMIN_USERNAME');
  const password = properties.getProperty('WPM_ADMIN_PASSWORD');
  const displayName = properties.getProperty('WPM_ADMIN_NAME') || 'Teacher';
  if (!username || !password) throw new Error('Set WPM_ADMIN_USERNAME and WPM_ADMIN_PASSWORD in Script Properties first.');
  const result = createAdminAccount(username, password, displayName);
  properties.deleteProperty('WPM_ADMIN_PASSWORD');
  return result + ' The bootstrap password property has been removed.';
}

function createStudentAccount(studentId, username, password, displayName, section) {
  if (!studentId || !username || !password || password.length < 12 || !displayName) {
    throw new Error('Student ID, username, display name, and a password of at least 12 characters are required.');
  }
  return withLock_(function() {
    const sheet = sheet_(APP.sheets.students);
    const rows = records_(sheet);
    if (rows.some(function(row) {
      return String(row.username).toLowerCase() === String(username).toLowerCase() ||
        String(row.studentId).toLowerCase() === String(studentId).toLowerCase();
    })) {
      throw new Error('Student ID and username must be unique.');
    }
    const salt = Utilities.getUuid() + Utilities.getUuid();
    sheet.appendRow([
      String(studentId).trim(),
      String(username).trim(),
      String(displayName).trim(),
      String(section || '').trim(),
      'active',
      salt,
      passwordHash_(password, salt),
      'student',
      new Date().toISOString()
    ]);
    return 'Student account created.';
  });
}

function apiRequest(request) {
  try {
    if (!request || typeof request.action !== 'string') throw new Error('Invalid request.');
    const payload = request.payload || {};
    if (request.action === 'login') return ok_(login_(payload.username, payload.password));
    if (request.action === 'logout') return ok_(logout_(payload.token));
    const user = requireSession_(payload.token);
    switch (request.action) {
      case 'bootstrap': return ok_(bootstrap_(user));
      case 'leaderboard': return ok_(leaderboard_(payload.scope || 'classroom', user));
      case 'createRace': return ok_(createRace_(payload, user));
      case 'joinRace': return ok_(joinRace_(payload, user));
      case 'beginRace': return ok_(beginRace_(payload, user));
      case 'raceState': return ok_(raceState_(payload, user));
      case 'liveProgress': return ok_(liveProgress_(payload, user));
      case 'submitRace': return ok_(submitRace_(payload, user));
      case 'adminOverview': requireAdmin_(user); return ok_(adminOverview_());
      case 'adminExport': requireAdmin_(user); return ok_(adminExport_());
      case 'adminCreateStudent':
        requireAdmin_(user);
        return ok_(createStudentAccount(payload.studentId, payload.username, payload.password, payload.displayName, payload.section));
      case 'adminAddPassage':
        requireAdmin_(user);
        return ok_(adminAddPassage_(payload));
      case 'adminSetAccuracy':
        requireAdmin_(user);
        return ok_(adminSetAccuracy_(payload.value));
      default: throw new Error('Unknown action.');
    }
  } catch (error) {
    return { ok: false, error: error && error.message ? error.message : 'Request failed.' };
  }
}

function login_(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  if (!username || !password || username.length > 100 || password.length > 256) throw new Error('Enter your username and password.');
  const cache = CacheService.getScriptCache();
  const key = 'login:' + digest_(username.toLowerCase());
  const count = Number(cache.get(key) || 0);
  if (count >= 8) throw new Error('Too many login attempts. Please wait a few minutes and try again.');
  const rows = records_(sheet_(APP.sheets.students));
  const student = rows.find(function(row) {
    return String(row.username).toLowerCase() === username.toLowerCase() ||
      String(row.studentId).toLowerCase() === username.toLowerCase();
  });
  const valid = student && String(student.status).toLowerCase() === 'active' &&
    constantTimeEqual_(passwordHash_(password, String(student.salt)), String(student.passwordHash));
  if (!valid) {
    cache.put(key, String(count + 1), 300);
    throw new Error('Username or password is incorrect.');
  }
  cache.remove(key);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expiresAt = Date.now() + APP.sessionHours * 60 * 60 * 1000;
  sheet_(APP.sheets.sessions).appendRow([digest_(token), student.studentId, expiresAt, false]);
  return { token: token, user: publicStudent_(student) };
}

function logout_(token) {
  const session = findSession_(token);
  if (session) {
    const sheet = sheet_(APP.sheets.sessions);
    sheet.getRange(session.row, 4).setValue(true);
  }
  return true;
}

function requireSession_(token) {
  const session = findSession_(token);
  if (!session) throw new Error('Your session expired. Please sign in again.');
  const student = records_(sheet_(APP.sheets.students)).find(function(row) {
    return String(row.studentId) === String(session.studentId);
  });
  if (!student || String(student.status).toLowerCase() !== 'active') throw new Error('This account is unavailable.');
  return student;
}

function findSession_(token) {
  if (!token || String(token).length < 40 || String(token).length > 100) return null;
  const key = digest_(String(token));
  const rows = records_(sheet_(APP.sheets.sessions));
  const record = rows.find(function(row) {
    return String(row.tokenHash) === key && row.revoked !== true && String(row.revoked).toLowerCase() !== 'true' && Number(row.expiresAt) > Date.now();
  });
  return record ? { row: record._row, studentId: record.studentId } : null;
}

function bootstrap_(user) {
  const results = records_(sheet_(APP.sheets.results)).filter(function(row) {
    return String(row.studentId) === String(user.studentId) && String(row.practice).toLowerCase() !== 'true';
  });
  const sorted = results.slice().sort(function(a, b) { return new Date(b.completedAt) - new Date(a.completedAt); });
  const eligible = results.filter(function(row) { return truthy_(row.eligible); });
  const best = eligible.reduce(function(current, row) { return !current || Number(row.validWpm) > Number(current.validWpm) ? row : current; }, null);
  const avg = function(key) {
    return results.length ? round_(results.reduce(function(total, row) { return total + Number(row[key] || 0); }, 0) / results.length, 1) : 0;
  };
  const passages = records_(sheet_(APP.sheets.passages)).filter(function(row) { return truthy_(row.active); });
  const daily = dailyPassage_(passages);
  const active = records_(sheet_(APP.sheets.races)).find(function(row) {
    return (row.status === 'waiting' || row.status === 'racing') &&
      Number(row.expiresAt) > Date.now() &&
      parseJson_(row.participantIds, []).includes(String(user.studentId));
  });
  const activePassage = active && active.passageId
    ? passages.find(function(row) { return String(row.passageId) === String(active.passageId); })
    : null;
  const recent = sorted.slice(0, 8).map(publicResult_);
  const achievements = records_(sheet_(APP.sheets.achievements)).filter(function(row) {
    return String(row.studentId) === String(user.studentId);
  }).map(function(row) { return String(row.achievementId); });
  const classResults = records_(sheet_(APP.sheets.results)).filter(function(row) {
    return truthy_(row.eligible) && String(row.practice).toLowerCase() !== 'true';
  });
  const rank = classroomRank_(user, classResults);
  return {
    user: publicStudent_(user),
    stats: {
      races: results.length,
      validWords: results.reduce(function(total, row) { return total + Number(row.correctWords || 0); }, 0),
      bestWpm: best ? Number(best.validWpm) : 0,
      averageWpm: avg('validWpm'),
      bestAccuracy: results.length ? Math.max.apply(null, results.map(function(row) { return Number(row.accuracy || 0); })) : 0,
      averageAccuracy: avg('accuracy'),
      dailyStreak: dailyStreak_(user.studentId, results),
      rank: rank
    },
    recent: recent,
    chart: sorted.slice(0, 12).reverse().map(function(row) {
      return { date: String(row.completedAt).slice(0, 10), wpm: Number(row.validWpm), accuracy: Number(row.accuracy) };
    }),
    achievements: achievements,
    passages: passages.map(publicPassage_),
    daily: daily,
    activeRace: active ? racePayload_(active, user, activePassage) : null,
    settings: { accuracyThreshold: getAccuracyThreshold_(), pollMs: APP.sessionPollMs },
    leaderboard: leaderboard_('classroom', user).slice(0, 5),
    dailyLeaderboard: leaderboard_('daily', user).slice(0, 5)
  };
}

function dailyPassage_(passages) {
  if (!passages.length) return null;
  const date = Utilities.formatDate(new Date(), getSetting_('timezone') || 'Asia/Dhaka', 'yyyy-MM-dd');
  const timezone = getSetting_('timezone') || 'Asia/Dhaka';
  return withLock_(function() {
    const sheet = sheet_(APP.sheets.settings);
    const settingKey = 'dailyChallenge:' + date;
    const settings = records_(sheet);
    const saved = settings.find(function(row) { return String(row.key) === settingKey; });
    let selected = saved ? passages.find(function(row) { return String(row.passageId) === String(saved.value); }) : null;
    if (!selected) {
      const index = parseInt(digest_(date).slice(0, 8), 16) % passages.length;
      selected = passages[index];
      if (saved) sheet.getRange(saved._row, 2).setValue(String(selected.passageId));
      else sheet.appendRow([settingKey, String(selected.passageId)]);
    }
    return { date: date, timezone: timezone, duration: 60, passage: publicPassage_(selected) };
  });
}

function createRace_(payload, user) {
  const cache = CacheService.getScriptCache();
  const createKey = 'race-create:' + digest_(user.studentId);
  const createCount = Number(cache.get(createKey) || 0);
  if (createCount >= 12) throw new Error('Too many race-room requests. Please wait a few minutes.');
  cache.put(createKey, String(createCount + 1), 300);
  const mode = String(payload.mode || 'copy');
  let duration = Number(payload.duration);
  const validModes = ['copy', 'free', 'accuracy', 'practice', 'daily'];
  if (validModes.indexOf(mode) < 0 || APP.raceDurations.indexOf(duration) < 0) throw new Error('Choose a supported mode and duration.');
  if (mode === 'daily') duration = 60;
  const isPractice = mode === 'practice' || payload.practice === true;
  let passages = records_(sheet_(APP.sheets.passages)).filter(function(row) { return truthy_(row.active); });
  if (!passages.length) throw new Error('No active passages are configured.');
  if (payload.difficulty && payload.difficulty !== 'any' && mode !== 'free' && mode !== 'practice') {
    const filtered = passages.filter(function(row) { return String(row.difficulty) === String(payload.difficulty); });
    if (filtered.length) passages = filtered;
  }
  let passage = null;
  if (mode === 'daily') {
    passage = dailyPassage_(passages).passage;
    payload = Object.assign({}, payload, { passageId: passage.passageId });
  }
  if (mode !== 'free' && mode !== 'practice') {
    passage = payload.passageId
      ? passages.find(function(row) { return String(row.passageId) === String(payload.passageId); })
      : passages[Math.floor(Math.random() * passages.length)];
    if (!passage) throw new Error('That passage is no longer available.');
  }
  const raceId = Utilities.getUuid();
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const now = Date.now();
  const room = {
    raceId: raceId,
    roomCode: payload.privateRoom ? code : '',
    mode: mode,
    duration: duration,
    difficulty: String(payload.difficulty || 'any'),
    passageId: passage ? passage.passageId : '',
    status: 'waiting',
    hostId: user.studentId,
    participantIds: [String(user.studentId)],
    createdAt: new Date(now).toISOString(),
    startedAt: '',
    expiresAt: now + 15 * 60 * 1000,
    practice: isPractice
  };
  let selectedPassage = passage;
  let existingRoom = null;
  withLock_(function() {
    if (payload.privateRoom) {
      const roomCodes = records_(sheet_(APP.sheets.races)).map(function(item) { return String(item.roomCode || ''); });
      while (roomCodes.indexOf(room.roomCode) >= 0) {
        room.roomCode = String(Math.floor(100000 + Math.random() * 900000));
      }
    }
    const activeRooms = records_(sheet_(APP.sheets.races)).filter(function(item) {
      return (item.status === 'waiting' || item.status === 'racing') &&
        Number(item.expiresAt) > now &&
        parseJson_(item.participantIds, []).includes(String(user.studentId));
    });
    if (activeRooms.length) {
      existingRoom = activeRooms[0];
      return;
    }
    if (!payload.privateRoom && !isPractice) {
      const waiting = records_(sheet_(APP.sheets.races)).find(function(row) {
        return row.status === 'waiting' && !row.roomCode && row.mode === mode &&
          Number(row.duration) === duration && row.difficulty === room.difficulty &&
          parseJson_(row.participantIds, []).length < 12 &&
          Number(row.expiresAt) > now && !parseJson_(row.participantIds, []).includes(String(user.studentId));
      });
      if (waiting) {
        const participants = parseJson_(waiting.participantIds, []);
        participants.push(String(user.studentId));
        writeField_(APP.sheets.races, waiting._row, 'participantIds', JSON.stringify(participants));
        room.raceId = waiting.raceId;
        room.roomCode = '';
        room.status = waiting.status;
        room.startedAt = waiting.startedAt;
        room.participantIds = participants;
        selectedPassage = waiting.passageId
          ? passages.find(function(item) { return String(item.passageId) === String(waiting.passageId); })
          : null;
      } else {
        appendRace_(room);
      }
    } else {
      appendRace_(room);
    }
  });
  if (existingRoom) {
    const currentPassage = existingRoom.passageId
      ? records_(sheet_(APP.sheets.passages)).find(function(item) { return String(item.passageId) === String(existingRoom.passageId); })
      : null;
    return racePayload_(existingRoom, user, currentPassage);
  }
  return racePayload_(room, user, selectedPassage);
}

function appendRace_(race) {
  sheet_(APP.sheets.races).appendRow([
    race.raceId, race.roomCode, race.mode, race.duration, race.difficulty, race.passageId,
    race.status, race.hostId, JSON.stringify(race.participantIds), race.createdAt,
    race.startedAt, race.expiresAt
  ]);
}

function joinRace_(payload, user) {
  const cache = CacheService.getScriptCache();
  const attemptKey = 'room-join:' + digest_(user.studentId);
  const attempts = Number(cache.get(attemptKey) || 0);
  if (attempts >= 12) throw new Error('Too many room-code attempts. Please wait a few minutes.');
  cache.put(attemptKey, String(attempts + 1), 300);
  const code = String(payload.code || '').trim();
  if (!/^\d{6}$/.test(code)) throw new Error('Enter a valid six-digit room code.');
  return withLock_(function() {
    const room = records_(sheet_(APP.sheets.races)).find(function(row) { return String(row.roomCode) === code; });
    if (!room || room.status !== 'waiting' || Number(room.expiresAt) <= Date.now()) throw new Error('Room not found or no longer accepting players.');
    const participants = parseJson_(room.participantIds, []);
    if (!participants.includes(String(user.studentId))) {
      const alreadyActive = records_(sheet_(APP.sheets.races)).some(function(item) {
        return (item.status === 'waiting' || item.status === 'racing') &&
          Number(item.expiresAt) > Date.now() &&
          parseJson_(item.participantIds, []).includes(String(user.studentId));
      });
      if (alreadyActive) throw new Error('You are already in an active race.');
      if (participants.length >= 12) throw new Error('This room is full.');
      participants.push(String(user.studentId));
      writeField_(APP.sheets.races, room._row, 'participantIds', JSON.stringify(participants));
    }
    cache.remove(attemptKey);
    const passage = records_(sheet_(APP.sheets.passages)).find(function(row) { return String(row.passageId) === String(room.passageId); });
    return racePayload_(Object.assign({}, room, { participantIds: participants }), user, passage);
  });
}

function beginRace_(payload, user) {
  return withLock_(function() {
    const room = raceById_(payload.raceId);
    if (!room || !parseJson_(room.participantIds, []).includes(String(user.studentId))) throw new Error('Race session not found.');
    if (room.status === 'waiting') {
      if (Number(room.expiresAt) <= Date.now()) {
        writeField_(APP.sheets.races, room._row, 'status', 'expired');
        throw new Error('This room expired before the race started. Create a new race.');
      }
      writeField_(APP.sheets.races, room._row, 'status', 'racing');
      writeField_(APP.sheets.races, room._row, 'startedAt', new Date().toISOString());
      writeField_(APP.sheets.races, room._row, 'expiresAt', Date.now() + (Number(room.duration) + 60) * 1000);
      room.status = 'racing';
      room.startedAt = new Date().toISOString();
    }
    if (room.status !== 'racing') throw new Error('This race is no longer active.');
    return { startedAt: room.startedAt, duration: Number(room.duration), status: room.status };
  });
}

function raceState_(payload, user) {
  const room = raceById_(payload.raceId);
  if (!room || !parseJson_(room.participantIds, []).includes(String(user.studentId))) throw new Error('Race session not found.');
  if (room.status === 'waiting' && Number(room.expiresAt) <= Date.now()) {
    writeField_(APP.sheets.races, room._row, 'status', 'expired');
    room.status = 'expired';
  }
  const allResults = records_(sheet_(APP.sheets.results)).filter(function(row) { return String(row.raceId) === String(room.raceId); });
  const participantIds = parseJson_(room.participantIds, []);
  const elapsed = room.startedAt ? (Date.now() - new Date(room.startedAt).getTime()) / 1000 : 0;
  if (room.status === 'racing' &&
      (allResults.length >= participantIds.length || elapsed > Number(room.duration) + 15)) {
    writeField_(APP.sheets.races, room._row, 'status', 'completed');
    room.status = 'completed';
  }
  const progress = records_(sheet_(APP.sheets.progress)).filter(function(row) { return String(row.raceId) === String(room.raceId); });
  const studentRows = records_(sheet_(APP.sheets.students));
  const competitors = participantIds.map(function(id) {
    const student = studentRows.find(function(row) { return String(row.studentId) === String(id); });
    const result = allResults.find(function(row) { return String(row.studentId) === String(id); });
    const live = progress.find(function(row) { return String(row.studentId) === String(id); });
    return {
      name: student ? String(student.displayName) : 'Racer',
      submitted: Boolean(result),
      wpm: result ? Number(result.validWpm) : (live ? Number(live.wpm) : 0),
      accuracy: result ? Number(result.accuracy) : (live ? Number(live.accuracy) : 0),
      progress: result ? 100 : (live ? Number(live.progress) : 0)
    };
  });
  const passage = room.passageId ? records_(sheet_(APP.sheets.passages)).find(function(row) { return String(row.passageId) === String(room.passageId); }) : null;
  const completed = room.status === 'completed';
  const standings = completed ? allResults.filter(function(row) {
    return truthy_(row.eligible) && String(row.practice).toLowerCase() !== 'true';
  }).sort(compareResults_) : [];
  const isWinner = room.mode !== 'practice' && completed && standings.length > 0 &&
    String(standings[0].studentId) === String(user.studentId);
  const namedResults = completed ? standings.map(function(row, index) {
    const student = studentRows.find(function(item) { return String(item.studentId) === String(row.studentId); });
    return Object.assign(publicResult_(row), { rank: index + 1, name: student ? String(student.displayName) : 'Racer' });
  }) : [];
  return {
    status: room.status,
    startedAt: room.startedAt,
    duration: Number(room.duration),
    mode: room.mode,
    competitors: competitors,
    passage: passage ? publicPassage_(passage) : null,
    results: namedResults,
    isWinner: isWinner,
    winnerVideoUrl: isWinner ? (PropertiesService.getScriptProperties().getProperty('WPM_WINNER_VIDEO_URL') || '') : '',
    leaderboard: completed ? leaderboard_('race:' + room.raceId, user) : []
  };
}

function liveProgress_(payload, user) {
  const room = raceById_(payload.raceId);
  if (!room || room.status !== 'racing' || !parseJson_(room.participantIds, []).includes(String(user.studentId))) {
    throw new Error('This race is not active.');
  }
  const elapsed = Math.max(1, (Date.now() - new Date(room.startedAt).getTime()) / 1000);
  if (elapsed > Number(room.duration) + 10) return false;
  const typed = String(payload.text || '').slice(0, 12000);
  const wpm = round_(typed.length / 5 / (elapsed / 60), 1);
  const passage = room.passageId ? records_(sheet_(APP.sheets.passages)).find(function(row) {
    return String(row.passageId) === String(room.passageId);
  }) : null;
  const progress = passage && Number(room.duration) > 0
    ? Math.min(99, round_(typed.length / Math.max(1, String(passage.text).length) * 100, 1))
    : Math.min(99, round_(wpm / 120 * 100, 1));
  const sheet = sheet_(APP.sheets.progress);
  const old = records_(sheet).find(function(row) {
    return String(row.raceId) === String(room.raceId) && String(row.studentId) === String(user.studentId);
  });
  if (old && Date.now() - new Date(old.updatedAt).getTime() < 2000) return false;
  const accuracy = Math.max(0, Math.min(100, Number(payload.accuracy) || 0));
  const values = [room.raceId, user.studentId, wpm, accuracy, progress, new Date().toISOString()];
  if (old) sheet.getRange(old._row, 1, 1, values.length).setValues([values]);
  else sheet.appendRow(values);
  return true;
}

function submitRace_(payload, user) {
  return withLock_(function() {
    const room = raceById_(payload.raceId);
    if (!room || !parseJson_(room.participantIds, []).includes(String(user.studentId))) throw new Error('Race session not found.');
    const already = records_(sheet_(APP.sheets.results)).some(function(row) {
      return String(row.raceId) === String(room.raceId) && String(row.studentId) === String(user.studentId);
    });
    if (already) throw new Error('A result has already been submitted for this race.');
    if (room.mode === 'daily') {
      const timezone = getSetting_('timezone') || 'Asia/Dhaka';
      const today = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
      const completedToday = records_(sheet_(APP.sheets.results)).some(function(row) {
        return String(row.studentId) === String(user.studentId) && row.mode === 'daily' &&
          Utilities.formatDate(new Date(row.completedAt), timezone, 'yyyy-MM-dd') === today;
      });
      if (completedToday) throw new Error('You have already completed today’s challenge.');
    }
    if ((room.status !== 'racing' && room.status !== 'completed') || !room.startedAt) throw new Error('This race has not started.');
    if (Date.now() > new Date(room.startedAt).getTime() + (Number(room.duration) + 15) * 1000) {
      throw new Error('The race submission window has closed.');
    }
    const elapsed = Math.max(1, Math.min(Number(room.duration), (Date.now() - new Date(room.startedAt).getTime()) / 1000));
    if (elapsed < 3) throw new Error('A race must run for at least 3 seconds before submission.');
    if (typeof payload.text !== 'string' || payload.text.length > 12000) {
      throw new Error('Race text is missing or exceeds the 12,000-character limit.');
    }
    const text = payload.text;
    const passage = room.passageId ? records_(sheet_(APP.sheets.passages)).find(function(row) { return String(row.passageId) === String(room.passageId); }) : null;
    const stats = calculateScore_(room.mode, text, passage ? String(passage.text) : '', elapsed);
    if (stats.rawWpm > 300) throw new Error('This result exceeds the 300 WPM consistency limit and was not saved.');
    const isPractice = room.mode === 'practice';
    const eligible = !isPractice && stats.accuracy >= getAccuracyThreshold_() && stats.validWpm > 0;
    const result = [
      Utilities.getUuid(), room.raceId, user.studentId, room.mode, Number(room.duration),
      stats.rawWpm, stats.validWpm, stats.accuracy, stats.correctWords, stats.invalidWords,
      stats.errors, stats.correctChars, stats.typedChars, round_(elapsed, 1), eligible, new Date().toISOString(), isPractice
    ];
    sheet_(APP.sheets.results).appendRow(result);
    const participants = parseJson_(room.participantIds, []);
    const submittedCount = records_(sheet_(APP.sheets.results)).filter(function(row) {
      return String(row.raceId) === String(room.raceId);
    }).length;
    if (submittedCount >= participants.length) {
      writeField_(APP.sheets.races, room._row, 'status', 'completed');
    }
    const submittedResults = records_(sheet_(APP.sheets.results)).filter(function(row) {
      return String(row.raceId) === String(room.raceId);
    });
    const raceResults = submittedResults.filter(function(row) {
      return String(row.raceId) === String(room.raceId) && truthy_(row.eligible);
    }).sort(compareResults_);
    const complete = submittedResults.length >= parseJson_(room.participantIds, []).length ||
      Date.now() >= new Date(room.startedAt).getTime() + (Number(room.duration) + 15) * 1000;
    const isWinner = !isPractice && complete && eligible && raceResults.length > 0 &&
      String(raceResults[0].studentId) === String(user.studentId);
    const response = {
      result: publicResult_(rowFromValues_(HEADERS.Results, result)),
      personalBest: updateAchievements_(user, stats, eligible, isPractice, Number(room.duration), room.raceId),
      leaderboard: leaderboard_('race:' + room.raceId, user),
      raceComplete: complete,
      isWinner: isWinner,
      rank: isWinner ? 1 : (complete && eligible ? raceResults.findIndex(function(row) {
        return String(row.studentId) === String(user.studentId);
      }) + 1 : null),
      winnerVideoUrl: isWinner ? (PropertiesService.getScriptProperties().getProperty('WPM_WINNER_VIDEO_URL') || '') : ''
    };
    return response;
  });
}

function calculateScore_(mode, typed, reference, elapsed) {
  const typedChars = typed.length;
  const minutes = elapsed / 60;
  const rawWpm = round_(typedChars / 5 / minutes, 1);
  if (mode === 'free' || mode === 'practice') {
    const tokens = typed.match(/[A-Za-z]+(?:['’-][A-Za-z]+)*/g) || [];
    const dictionary = dictionaryWords_();
    const validWords = tokens.filter(function(word) {
      return isDictionaryWord_(word, dictionary);
    });
    const accuracy = tokens.length ? round_(validWords.length / tokens.length * 100, 1) : 0;
    return {
      rawWpm: rawWpm,
      validWpm: round_(validWords.length / minutes, 1),
      accuracy: accuracy,
      correctWords: validWords.length,
      invalidWords: tokens.length - validWords.length,
      errors: tokens.length - validWords.length,
      correctChars: validWords.reduce(function(total, word) { return total + word.length; }, 0),
      typedChars: typedChars
    };
  }
  const compared = Math.min(typed.length, reference.length);
  let correctChars = 0;
  let errors = Math.max(0, typed.length - reference.length);
  for (let i = 0; i < compared; i++) {
    if (typed.charAt(i) === reference.charAt(i)) correctChars++;
    else errors++;
  }
  const accuracy = typed.length ? round_(correctChars / typed.length * 100, 1) : 0;
  const correctWords = completeWords_(typed, reference);
  return {
    rawWpm: rawWpm,
    validWpm: round_(correctWords / minutes, 1),
    accuracy: accuracy,
    correctWords: correctWords,
    invalidWords: Math.max(0, (typed.match(/\S+/g) || []).length - correctWords),
    errors: errors,
    correctChars: correctChars,
    typedChars: typedChars
  };
}

function completeWords_(typed, reference) {
  const ref = reference.match(/\S+/g) || [];
  let offset = 0;
  let count = 0;
  ref.forEach(function(word) {
    const start = reference.indexOf(word, offset);
    const end = start + word.length;
    const startBoundary = start === 0 || typed.charAt(start - 1) === reference.charAt(start - 1);
    const endBoundary = end === reference.length || typed.length === end ||
      typed.charAt(end) === reference.charAt(end);
    if (end <= typed.length && startBoundary && endBoundary && typed.slice(start, end) === word) count++;
    offset = end;
  });
  return count;
}

function leaderboard_(scope, user) {
  const today = Utilities.formatDate(new Date(), getSetting_('timezone') || 'Asia/Dhaka', 'yyyy-MM-dd');
  const results = records_(sheet_(APP.sheets.results)).filter(function(row) {
    if (String(row.practice).toLowerCase() === 'true' || !truthy_(row.eligible)) return false;
    if (scope.indexOf('race:') === 0) return String(row.raceId) === scope.slice(5);
    if (scope === 'weekly') return new Date(row.completedAt).getTime() >= Date.now() - 7 * 24 * 60 * 60 * 1000;
    if (scope === 'daily') {
      return row.mode === 'daily' &&
        Utilities.formatDate(new Date(row.completedAt), getSetting_('timezone') || 'Asia/Dhaka', 'yyyy-MM-dd') === today;
    }
    return true;
  });
  const students = records_(sheet_(APP.sheets.students));
  const bestByStudent = {};
  results.forEach(function(row) {
    const student = students.find(function(item) { return String(item.studentId) === String(row.studentId); });
    if (scope === 'classroom' && user.section && student && String(student.section) !== String(user.section)) return;
    const previous = bestByStudent[row.studentId];
    if (!previous || compareResults_(row, previous) < 0) bestByStudent[row.studentId] = row;
  });
  return Object.keys(bestByStudent).map(function(id) {
    const row = bestByStudent[id];
    const student = students.find(function(item) { return String(item.studentId) === String(id); });
    return {
      name: student ? String(student.displayName) : 'Racer',
      wpm: Number(row.validWpm),
      accuracy: Number(row.accuracy),
      races: results.filter(function(item) { return String(item.studentId) === String(id); }).length,
      completedAt: row.completedAt
    };
  }).sort(function(a, b) {
    return b.wpm - a.wpm || b.accuracy - a.accuracy || a.completedAt.localeCompare(b.completedAt);
  }).map(function(row, index) { row.rank = index + 1; return row; });
}

function dailyStreak_(studentId, results) {
  const timezone = getSetting_('timezone') || 'Asia/Dhaka';
  const days = new Set(results.filter(function(row) { return row.mode === 'daily'; }).map(function(row) {
    return Utilities.formatDate(new Date(row.completedAt), timezone, 'yyyy-MM-dd');
  }));
  const today = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  const yesterday = new Date(new Date(today + 'T00:00:00Z').getTime() - 86400000).toISOString().slice(0, 10);
  let date = days.has(today) ? today : (days.has(yesterday) ? yesterday : '');
  let streak = 0;
  while (date && days.has(date)) {
    streak++;
    date = new Date(new Date(date + 'T00:00:00Z').getTime() - 86400000).toISOString().slice(0, 10);
  }
  return streak;
}

function adminOverview_() {
  return {
    students: records_(sheet_(APP.sheets.students)).filter(function(row) { return row.role === 'student'; }).map(publicStudent_),
    recentResults: records_(sheet_(APP.sheets.results)).slice(-30).reverse().map(publicResult_),
    activeRooms: records_(sheet_(APP.sheets.races)).filter(function(row) {
      return (row.status === 'waiting' || row.status === 'racing') && Number(row.expiresAt) > Date.now();
    }),
    accuracyThreshold: getAccuracyThreshold_(),
    passages: records_(sheet_(APP.sheets.passages)).map(publicPassage_)
  };
}

function adminExport_() {
  const headers = HEADERS.Results;
  const rows = records_(sheet_(APP.sheets.results));
  const csv = [headers.join(',')].concat(rows.map(function(row) {
    return headers.map(function(key) {
      const value = String(row[key] === undefined ? '' : row[key]).replace(/"/g, '""');
      return '"' + value + '"';
    }).join(',');
  })).join('\r\n');
  return { filename: 'wpm-racer-results.csv', content: csv };
}

function adminAddPassage_(payload) {
  const title = String(payload.title || '').trim();
  const text = String(payload.text || '').trim();
  const difficulty = String(payload.difficulty || 'medium');
  if (!title || text.length < 20 || text.length > 2000) throw new Error('Enter a title and passage of 20 to 2,000 characters.');
  if (['easy', 'medium', 'hard'].indexOf(difficulty) < 0) throw new Error('Choose a valid difficulty.');
  const id = Utilities.getUuid();
  sheet_(APP.sheets.passages).appendRow([id, title.slice(0, 100), text, difficulty, 'copy', true]);
  return { passageId: id };
}

function adminSetAccuracy_(value) {
  value = Number(value);
  if (!Number.isFinite(value) || value < 50 || value > 100) throw new Error('Accuracy threshold must be between 50 and 100.');
  const sheet = sheet_(APP.sheets.settings);
  const rows = records_(sheet);
  const row = rows.find(function(item) { return item.key === 'minimumAccuracy'; });
  if (row) sheet.getRange(row._row, 2).setValue(String(value));
  else sheet.appendRow(['minimumAccuracy', String(value)]);
  return value;
}

function updateAchievements_(user, stats, eligible, practice, duration, raceId) {
  if (practice) return false;
  const results = records_(sheet_(APP.sheets.results)).filter(function(row) {
    return String(row.studentId) === String(user.studentId) && String(row.practice).toLowerCase() !== 'true';
  });
  const prior = results.filter(function(row) { return String(row.raceId) !== String(raceId); });
  const best = prior.reduce(function(max, row) { return Math.max(max, Number(row.validWpm || 0)); }, 0);
  const awards = [];
  if (results.length === 1) awards.push('first-lap');
  if (stats.validWpm >= 60) awards.push('speed-demon');
  if (stats.accuracy === 100) awards.push('precision-master');
  if (stats.errors === 0) awards.push('perfect-run');
  if (Number(duration) === 300) awards.push('marathon');
  if (results.length >= 10) awards.push('consistency');
  if (eligible && stats.validWpm > best) awards.push('comeback-racer');
  if (raceWinCount_(user.studentId) >= 5) awards.push('champion');
  const sheet = sheet_(APP.sheets.achievements);
  const existing = records_(sheet).filter(function(row) { return String(row.studentId) === String(user.studentId); });
  awards.forEach(function(id) {
    if (!existing.some(function(row) { return row.achievementId === id; })) {
      sheet.appendRow([user.studentId, id, new Date().toISOString()]);
    }
  });
  return eligible && stats.validWpm > best;
}

function classroomRank_(user, allRows) {
  const eligibleStudents = {};
  const students = records_(sheet_(APP.sheets.students));
  allRows.forEach(function(row) {
    const student = students.find(function(item) { return String(item.studentId) === String(row.studentId); });
    if (!student || (user.section && String(student.section) !== String(user.section))) return;
    eligibleStudents[row.studentId] = Math.max(eligibleStudents[row.studentId] || 0, Number(row.validWpm));
  });
  const myBest = eligibleStudents[user.studentId] || 0;
  if (!myBest) return null;
  return Object.keys(eligibleStudents).filter(function(id) { return eligibleStudents[id] > myBest; }).length + 1;
}

function raceWinCount_(studentId) {
  const races = {};
  const completed = new Set(records_(sheet_(APP.sheets.races)).filter(function(row) {
    return row.status === 'completed';
  }).map(function(row) { return String(row.raceId); }));
  records_(sheet_(APP.sheets.results)).forEach(function(row) {
    if (!completed.has(String(row.raceId)) || !truthy_(row.eligible) || String(row.practice).toLowerCase() === 'true') return;
    const current = races[row.raceId];
    if (!current || compareResults_(row, current) < 0) races[row.raceId] = row;
  });
  return Object.keys(races).filter(function(id) { return String(races[id].studentId) === String(studentId); }).length;
}

function racePayload_(room, user, passage) {
  return {
    raceId: room.raceId,
    roomCode: room.roomCode || '',
    mode: room.mode,
    duration: Number(room.duration),
    status: room.status,
    passage: passage ? publicPassage_(passage) : null,
    participantCount: parseJson_(room.participantIds, []).length
  };
}

function raceById_(id) {
  return records_(sheet_(APP.sheets.races)).find(function(row) { return String(row.raceId) === String(id); });
}

function publicStudent_(row) {
  return {
    studentId: String(row.studentId),
    username: String(row.username),
    displayName: String(row.displayName),
    section: String(row.section || ''),
    role: String(row.role || 'student')
  };
}

function publicPassage_(row) {
  return { passageId: String(row.passageId), title: String(row.title), text: String(row.text), difficulty: String(row.difficulty) };
}

function publicResult_(row) {
  return {
    resultId: String(row.resultId || ''),
    raceId: String(row.raceId || ''),
    mode: String(row.mode || ''),
    duration: Number(row.duration || 0),
    wpm: Number(row.validWpm || 0),
    rawWpm: Number(row.rawWpm || 0),
    accuracy: Number(row.accuracy || 0),
    correctWords: Number(row.correctWords || 0),
    invalidWords: Number(row.invalidWords || 0),
    errors: Number(row.errors || 0),
    completedAt: String(row.completedAt || ''),
    eligible: truthy_(row.eligible),
    practice: truthy_(row.practice)
  };
}

function rowFromValues_(headers, values) {
  const row = {};
  headers.forEach(function(key, index) { row[key] = values[index]; });
  return row;
}

function requireAdmin_(user) {
  if (String(user.role) !== 'admin') throw new Error('Administrator access required.');
}

function getAccuracyThreshold_() {
  const row = records_(sheet_(APP.sheets.settings)).find(function(item) { return item.key === 'minimumAccuracy'; });
  return row ? Number(row.value) : APP.accuracyThreshold;
}

function getSetting_(key) {
  const row = records_(sheet_(APP.sheets.settings)).find(function(item) { return String(item.key) === String(key); });
  return row ? String(row.value) : '';
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Database is not initialized. Run setupDatabase from the Apps Script editor.');
  return sheet;
}

function dictionaryWords_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('wordbank');
  if (cached) return new Set(JSON.parse(cached));
  const words = new Set(STARTER_WORDS.map(normalizeWord_));
  records_(sheet_(APP.sheets.words)).forEach(function(row) {
    if (truthy_(row.active)) words.add(normalizeWord_(row.word));
  });
  const list = Array.from(words);
  cache.put('wordbank', JSON.stringify(list), 600);
  return words;
}

function records_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(function(values, index) {
    const row = { _row: index + 2 };
    headers.forEach(function(header, column) { row[header] = values[column]; });
    return row;
  }).filter(function(row) { return Object.keys(row).some(function(key) { return key !== '_row' && row[key] !== ''; }); });
}

function writeField_(sheetName, rowIndex, field, value) {
  const headers = HEADERS[sheetName];
  const column = headers.indexOf(field);
  if (column < 0) throw new Error('Invalid database field.');
  sheet_(sheetName).getRange(rowIndex, column + 1).setValue(value);
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return callback(); } finally { lock.releaseLock(); }
}

function passwordHash_(password, salt) {
  const passwordBytes = Utilities.newBlob(String(password)).getBytes();
  const saltBytes = Utilities.newBlob(String(salt)).getBytes();
  let u = Utilities.computeHmacSha256Signature(saltBytes.concat([0, 0, 0, 1]), passwordBytes);
  const output = u.slice();
  for (let i = 1; i < APP.kdfIterations; i++) {
    u = Utilities.computeHmacSha256Signature(u, passwordBytes);
    for (let j = 0; j < output.length; j++) output[j] = output[j] ^ u[j];
  }
  return u8hex_(output);
}

function digest_(value) {
  return u8hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value)));
}

function u8hex_(bytes) {
  return bytes.map(function(value) { return ('0' + ((value + 256) % 256).toString(16)).slice(-2); }).join('');
}

function constantTimeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let mismatch = 0;
  for (let i = 0; i < a.length; i++) mismatch |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return mismatch === 0;
}

function normalizeWord_(word) {
  return String(word).toLowerCase().replace(/[’]/g, "'").replace(/^[^a-z]+|[^a-z]+$/g, '');
}

function isDictionaryWord_(word, dictionary) {
  const normalized = normalizeWord_(word);
  if (dictionary.has(normalized)) return true;
  if (normalized.indexOf('-') >= 0) {
    return normalized.split('-').every(function(part) { return part && dictionary.has(part); });
  }
  return false;
}

function parseJson_(value, fallback) {
  try { return JSON.parse(value || ''); } catch (error) { return fallback; }
}

function truthy_(value) {
  return value === true || String(value).toLowerCase() === 'true' || String(value) === '1';
}

function round_(value, places) {
  const factor = Math.pow(10, places);
  return Math.round(Number(value) * factor) / factor;
}

function compareResults_(a, b) {
  return Number(b.validWpm) - Number(a.validWpm) ||
    Number(b.accuracy) - Number(a.accuracy) ||
    Number(a.errors) - Number(b.errors) ||
    new Date(a.completedAt) - new Date(b.completedAt);
}

