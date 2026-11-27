const SHEET_ID = '1XgL6X-wMNnhXyoOVBPcQhFVQBScbxulk9wNlRlZHqX0';
const ROOM_CACHE_TTL = 21600; // 6 hours
const CONFIG_CACHE_TTL = 30;  // 30 seconds

const ALLOWED_API_ACTIONS = {
  getGameConfig: true,
  createRoom: true,
  joinRoom: true,
  startGame: true,
  getGameState: true,
  updateGameState: true,
  useActiveSkill: true
};

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj !== undefined ? obj : { status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiCall(e.parameter.action, e.parameter.args ? JSON.parse(e.parameter.args) : []);
  }
  return jsonResponse({ ok: true, app: 'War of Wannakhadi API', status: 'running' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return handleApiCall(payload.action, payload.args || []);
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) });
  }
}

function handleApiCall(action, args) {
  try {
    if (!ALLOWED_API_ACTIONS[action]) {
      return jsonResponse({ error: 'ไม่พบฟังก์ชันที่อนุญาต: ' + action });
    }
    const fn = globalThis[action];
    if (typeof fn !== 'function') {
      return jsonResponse({ error: 'ไม่พบฟังก์ชัน: ' + action });
    }
    const result = fn.apply(null, Array.isArray(args) ? args : []);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) });
  }
}

function getSS() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  } catch (e) {
    return null;
  }
}

function getOrCreateRoomsSheet(ss) {
  let sheet = ss.getSheetByName('Rooms');
  if (!sheet) {
    sheet = ss.insertSheet('Rooms');
    sheet.appendRow(['RoomCode', 'Status', 'Data', 'LastUpdated']);
  }
  return sheet;
}

function roomCacheKey(roomCode) {
  return 'wow_room_' + String(roomCode).trim();
}

function writeRoomCache(roomCode, status, gameData) {
  try {
    CacheService.getScriptCache().put(
      roomCacheKey(roomCode),
      JSON.stringify({ status: status, data: gameData }),
      ROOM_CACHE_TTL
    );
  } catch (e) {}
}

function pushGameLog(gameData, msgHtml) {
  if (!Array.isArray(gameData.logs)) gameData.logs = [];
  gameData.logs.unshift('<div class="log-entry">' + msgHtml + '</div>');
  if (gameData.logs.length > 3) gameData.logs.length = 3;
}

function findRoomRow(sheet, roomCode) {
  const code = String(roomCode).trim();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === code) {
      return { rowIndex: i + 1, status: data[i][1], rawData: data[i][2] };
    }
  }
  return null;
}

function saveRoomState(sheet, rowIndex, roomCode, status, gameData) {
  gameData.v = (Number(gameData.v) || 0) + 1;
  gameData.status = status;
  const serialized = JSON.stringify(gameData);
  writeRoomCache(roomCode, status, gameData);
  sheet.getRange(rowIndex, 2, 1, 3).setValues([[status, serialized, new Date()]]);
  return gameData;
}

function getGameConfig(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    try {
      const cached = cache.get('wow_game_config');
      if (cached) return JSON.parse(cached);
    } catch (e) {}
  }

  let config = { board: {}, items: {}, characters: [], events: {}, questions: [], error: '' };
  const ss = getSS();
  if (!ss) {
    config.error = '🚨 ไม่สามารถเชื่อมต่อกับ Google Sheets ได้';
    return config;
  }

  try {
    const sheetB = ss.getSheetByName('BoardConfig');
    if (sheetB) {
      const dataB = sheetB.getDataRange().getValues();
      for (let i = 1; i < dataB.length; i++) {
        const cellNum = parseInt(dataB[i][0], 10);
        if (cellNum > 0) {
          config.board[cellNum] = {
            msg: dataB[i][1] ? dataB[i][1].toString() : '',
            power: parseInt(dataB[i][2], 10) || 0,
            move: parseInt(dataB[i][3], 10) || 0,
            skip: parseInt(dataB[i][4], 10) || 0,
            item: dataB[i][5] ? dataB[i][5].toString().trim() : '',
            warpTo: parseInt(dataB[i][6], 10) || 0,
            damage: parseInt(dataB[i][7], 10) || 0,
            isShop: Boolean(dataB[i][8]),
            isEvent: Boolean(dataB[i][9]),
            bossPower: parseInt(dataB[i][10], 10) || 0,
            bossName: dataB[i][11] ? dataB[i][11].toString() : 'มอนสเตอร์ปริศนา',
            bossImg: dataB[i][12] ? dataB[i][12].toString() : 'https://cdn-icons-png.flaticon.com/512/3204/3204642.png',
            isChest: Boolean(dataB[i][14])
          };
        }
      }
    }

    const sheetI = ss.getSheetByName('Items');
    if (sheetI) {
      const dataI = sheetI.getDataRange().getValues();
      for (let i = 1; i < dataI.length; i++) {
        const id = dataI[i][0] ? dataI[i][0].toString().trim() : '';
        if (id) {
          config.items[id] = {
            name: dataI[i][1] ? dataI[i][1].toString().trim() : 'ไอเท็ม',
            type: dataI[i][2] ? dataI[i][2].toString().trim() : 'power',
            val: parseInt(dataI[i][3], 10) || 0,
            icon: dataI[i][4] ? dataI[i][4].toString().trim() : '📦',
            price: parseInt(dataI[i][5], 10) || 200
          };
        }
      }
    }

    const sheetC = ss.getSheetByName('Characters') || ss.getSheetByName('Charaters') || ss.getSheetByName('ตัวละคร');
    if (sheetC) {
      const dataC = sheetC.getDataRange().getValues();
      for (let i = 1; i < dataC.length; i++) {
        if (dataC[i][0] && dataC[i][1]) {
          const rawActive = dataC[i][5] ? dataC[i][5].toString() : '';
          const splitActive = rawActive.split('//');

          const skillName = (splitActive[0] && splitActive[0].trim()) || (dataC[i][6] ? dataC[i][6].toString().trim() : '');
          let effectType = (splitActive[1] && splitActive[1].trim()) || (dataC[i][7] ? dataC[i][7].toString().trim() : '');
          if (!effectType && skillName === 'ควบม้าสีหมอก') {
            effectType = 'move_boost';
          }

          config.characters.push({
            name: dataC[i][0].toString().trim(),
            img: dataC[i][1].toString().trim(),
            skillType: dataC[i][2] ? dataC[i][2].toString().trim() : '',
            skillVal: parseInt(dataC[i][3], 10) || 0,
            basePower: dataC[i][4] !== '' && dataC[i][4] !== undefined ? parseInt(dataC[i][4], 10) : 200,
            skill_name: skillName,
            effect_type: effectType,
            activeSkillVal: parseInt(splitActive[2], 10) || parseInt(dataC[i][8], 10) || 0,
            activeCost: parseInt(splitActive[3], 10) || parseInt(dataC[i][9], 10) || 0,
            activeCooldown: parseInt(splitActive[4], 10) || parseInt(dataC[i][10], 10) || 0
          });
        }
      }
    } else {
      config.error = "🚨 หาแผ่นงานชื่อ 'Characters' ไม่เจอครับ";
    }

    const sheetE = ss.getSheetByName('Events');
    if (sheetE) {
      const dataE = sheetE.getDataRange().getValues();
      for (let i = 1; i < dataE.length; i++) {
        const id = dataE[i][0] ? dataE[i][0].toString().trim() : '';
        if (id) {
          config.events[id] = {
            name: dataE[i][1] ? dataE[i][1].toString().trim() : 'เหตุการณ์',
            type: dataE[i][2] ? dataE[i][2].toString().trim() : 'power_all',
            val: parseInt(dataE[i][3], 10) || 0
          };
        }
      }
    }

    const sheetQ = ss.getSheetByName('คำถาม');
    if (sheetQ) {
      const qData = sheetQ.getDataRange().getValues();
      for (let i = 1; i < qData.length; i++) {
        const row = qData[i];
        if (row[0]) {
          config.questions.push({
            q: row[0].toString(),
            choices: [
              row[1] !== undefined ? row[1].toString() : '',
              row[2] !== undefined ? row[2].toString() : '',
              row[3] !== undefined ? row[3].toString() : '',
              row[4] !== undefined ? row[4].toString() : ''
            ],
            ans: parseInt(row[5], 10) || 0,
            reward: parseInt(row[6], 10) || 500
          });
        }
      }
    }
  } catch (e) {
    config.error = '🚨 เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ' + e.message;
  }

  if (config.characters.length === 0 && config.error === '') {
    config.characters = [{
      name: 'นักรบสำรอง',
      img: 'https://cdn-icons-png.flaticon.com/512/616/616430.png',
      skillType: '',
      skillVal: 0,
      basePower: 200,
      skill_name: '',
      effect_type: '',
      activeSkillVal: 0,
      activeCost: 0,
      activeCooldown: 0
    }];
  }

  if (config.error === '') {
    try {
      cache.put('wow_game_config', JSON.stringify(config), CONFIG_CACHE_TTL);
    } catch (e) {}
  }
  return config;
}

function buildPlayerObject(id, playerName, characterObj) {
  const basePower = characterObj.basePower !== undefined ? Number(characterObj.basePower) : 200;
  let initPower = basePower;
  let initLives = 3;

  if (characterObj.skillType === 'start_power' || characterObj.skillType === 'rich_start') {
    initPower += (characterObj.skillVal || 1000);
  }
  if (characterObj.skillType === 'extra_life') {
    initLives += (characterObj.skillVal || 1);
  }

  return {
    id: id,
    name: playerName,
    pos: 1,
    power: initPower,
    basePower: basePower,
    lives: initLives,
    maxLives: initLives,
    icon: characterObj.img,
    skipTurns: 0,
    items: [],
    shield: false,
    evade: false,
    buffsInitialized: false,
    skillType: characterObj.skillType || '',
    skillVal: characterObj.skillVal || 0,
    skill_name: characterObj.skill_name || '',
    effect_type: characterObj.effect_type || '',
    activeSkillVal: characterObj.activeSkillVal || 0,
    activeCost: characterObj.activeCost || 0,
    activeCooldown: characterObj.activeCooldown || 0,
    currentCooldown: 0,
    status: []
  };
}

function cleanupExpiredRooms(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 20) return;
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  for (let i = data.length - 1; i >= 1; i--) {
    const updated = data[i][3] instanceof Date ? data[i][3].getTime() : 0;
    if (updated > 0 && updated < cutoff) {
      sheet.deleteRow(i + 1);
    }
  }
}

function createRoom(playerName, characterObj) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, msg: 'เซิร์ฟเวอร์กำลังยุ่ง กรุณาลองใหม่' };
  try {
    const ss = getSS();
    if (!ss) return { success: false, msg: 'เชื่อมต่อ Sheet ไม่ได้' };
    const sheet = getOrCreateRoomsSheet(ss);
    cleanupExpiredRooms(sheet);

    const existing = sheet.getDataRange().getValues().slice(1).map(r => String(r[0]).trim());
    let roomCode = '';
    for (let tries = 0; tries < 20; tries++) {
      const candidate = Math.floor(1000 + Math.random() * 9000).toString();
      if (existing.indexOf(candidate) === -1) {
        roomCode = candidate;
        break;
      }
    }
    if (!roomCode) roomCode = Math.floor(1000 + Math.random() * 9000).toString();

    const initialData = {
      v: 1,
      players: [buildPlayerObject(0, playerName, characterObj)],
      turn: 0,
      status: 'Waiting',
      bossLevels: {},
      bossNerfs: {},
      logs: []
    };
    pushGameLog(initialData, '🚩 <b>' + playerName + '</b> สร้างห้องและรอผู้เล่น...');

    sheet.appendRow([roomCode, 'Waiting', JSON.stringify(initialData), new Date()]);
    writeRoomCache(roomCode, 'Waiting', initialData);
    return { success: true, roomCode: roomCode, myId: 0 };
  } finally {
    lock.releaseLock();
  }
}

function joinRoom(roomCode, playerName, characterObj) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, msg: 'เซิร์ฟเวอร์กำลังยุ่ง กรุณาลองใหม่' };
  try {
    const ss = getSS();
    if (!ss) return { success: false, msg: 'เชื่อมต่อ Sheet ไม่ได้' };
    const sheet = getOrCreateRoomsSheet(ss);
    const found = findRoomRow(sheet, roomCode);

    if (!found || found.status !== 'Waiting') {
      return { success: false, msg: 'ไม่พบรหัสห้องนี้ หรือเกมเริ่มไปแล้ว' };
    }

    const gameData = JSON.parse(found.rawData);
    if (gameData.players.length >= 4) return { success: false, msg: 'ห้องนี้เต็มแล้ว (สูงสุด 4 คน)' };
    if (gameData.players.map(p => p.icon).indexOf(characterObj.img) !== -1) {
      return { success: false, msg: 'ตัวละครนี้มีคนเลือกไปแล้ว' };
    }

    const newId = gameData.players.length;
    gameData.players.push(buildPlayerObject(newId, playerName, characterObj));
    pushGameLog(gameData, '✨ <b>' + playerName + '</b> เข้าร่วมวงผจญภัย!');

    saveRoomState(sheet, found.rowIndex, roomCode, 'Waiting', gameData);
    return { success: true, roomCode: String(roomCode).trim(), myId: newId };
  } finally {
    lock.releaseLock();
  }
}

function startGame(roomCode) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    const ss = getSS();
    if (!ss) return false;
    const sheet = ss.getSheetByName('Rooms');
    if (!sheet) return false;
    const found = findRoomRow(sheet, roomCode);
    if (!found) return false;

    const gameData = JSON.parse(found.rawData);
    pushGameLog(gameData, '🚀 เกมเริ่มแล้ว! ขอให้โชคดี');
    saveRoomState(sheet, found.rowIndex, roomCode, 'Playing', gameData);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function getGameState(roomCode, clientVersion) {
  const code = String(roomCode).trim();
  const cache = CacheService.getScriptCache();
  try {
    const cachedStr = cache.get(roomCacheKey(code));
    if (cachedStr) {
      const cached = JSON.parse(cachedStr);
      if (clientVersion !== undefined && clientVersion !== null && Number(clientVersion) === Number(cached.data.v || 0)) {
        return { unchanged: true, v: cached.data.v || 0, status: cached.status };
      }
      return cached;
    }
  } catch (e) {}

  const ss = getSS();
  if (!ss) return null;
  const sheet = ss.getSheetByName('Rooms');
  if (!sheet) return null;
  const found = findRoomRow(sheet, code);
  if (!found) return null;

  const gameData = JSON.parse(found.rawData);
  gameData.status = found.status;
  if (!gameData.v) gameData.v = 1;
  writeRoomCache(code, found.status, gameData);

  if (clientVersion !== undefined && clientVersion !== null && Number(clientVersion) === Number(gameData.v)) {
    return { unchanged: true, v: gameData.v, status: found.status };
  }
  return { status: found.status, data: gameData };
}

function updateGameState(roomCode, gameDataStr) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    const ss = getSS();
    if (!ss) return false;
    const sheet = ss.getSheetByName('Rooms');
    if (!sheet) return false;
    const found = findRoomRow(sheet, roomCode);
    if (!found) return false;

    const gd = JSON.parse(gameDataStr);
    const status = gd.status === 'Finished' ? 'Finished' : (found.status || gd.status || 'Playing');
    const saved = saveRoomState(sheet, found.rowIndex, roomCode, status, gd);
    return { success: true, v: saved.v };
  } finally {
    lock.releaseLock();
  }
}

function useActiveSkill(roomCode, playerId, targetId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, msg: 'เซิร์ฟเวอร์กำลังยุ่ง กรุณาลองใหม่' };
  try {
    const ss = getSS();
    if (!ss) return { success: false, msg: 'เชื่อมต่อ Sheet ไม่ได้' };
    const sheet = ss.getSheetByName('Rooms');
    if (!sheet) return { success: false, msg: 'ไม่พบข้อมูลห้อง' };
    const found = findRoomRow(sheet, roomCode);
    if (!found) return { success: false, msg: 'ไม่พบห้องเกม' };

    let gameData = JSON.parse(found.rawData);
    const playerIndex = gameData.players.findIndex(p => p.id === playerId);
    if (playerIndex === -1) return { success: false, msg: 'ไม่พบข้อมูลผู้เล่น' };

    let player = gameData.players[playerIndex];
    if (!player.skill_name) return { success: false, msg: 'ตัวละครนี้ไม่มี Active Skill' };
    if (player.currentCooldown > 0) return { success: false, msg: 'สกิลยังติดคูลดาวน์อีก ' + player.currentCooldown + ' เทิร์น' };
    if (player.power < (player.activeCost || 0)) {
      return { success: false, msg: 'พลังเวทย์ไม่พอใช้สกิล (ต้องการพลัง ' + player.activeCost + ')' };
    }

    player.power -= (player.activeCost || 0);

    let effectType = (player.effect_type || '').toLowerCase();
    if (!effectType && player.skill_name === 'ควบม้าสีหมอก') {
      effectType = 'move_boost';
    }

    let target = (targetId !== null && targetId !== undefined)
      ? gameData.players.find(p => p.id === targetId)
      : null;

    player.status = player.status || [];
    if (target) target.status = target.status || [];

    const skillResult = executeActiveSkill(effectType, player, target, gameData);
    player = skillResult.caster;
    gameData = skillResult.gameState;
    gameData.players[playerIndex] = player;

    if (target && skillResult.target) {
      const tIndex = gameData.players.findIndex(p => p.id === target.id);
      if (tIndex !== -1) {
        gameData.players[tIndex] = skillResult.target;
        handleServerPlayerDeath(gameData.players[tIndex], gameData);
      }
    }
    handleServerPlayerDeath(gameData.players[playerIndex], gameData);

    player.currentCooldown = player.activeCooldown || 0;

    const targetLabel = target ? (' ใส่ ' + target.name) : '';
    pushGameLog(gameData, '💥 <b>' + player.name + '</b> ใช้สกิล [' + player.skill_name + ']' + targetLabel + '!');
    gameData.broadcast = {
      id: Date.now(),
      pId: playerId,
      title: '⚡ ' + player.name + ' ใช้สกิล' + targetLabel + '!',
      desc: player.skill_name,
      img: '✨'
    };

    // Check survivor win condition
    let nextStatus = found.status || gameData.status || 'Playing';
    const alive = gameData.players.filter(p => p.lives > 0);
    if (alive.length === 1 && gameData.players.length > 1) {
      nextStatus = 'Finished';
      gameData.winner = alive[0];
    }

    const saved = saveRoomState(sheet, found.rowIndex, roomCode, nextStatus, gameData);
    return { success: true, msg: 'ใช้สกิลสำเร็จ', data: saved };
  } finally {
    lock.releaseLock();
  }
}

function handleServerPlayerDeath(p, gameData) {
  if (!p || p.lives > 0) return;
  p.status = p.status || [];
  const autoRevIdx = p.status.indexOf('auto_revive');
  if (autoRevIdx !== -1) {
    p.status.splice(autoRevIdx, 1);
    p.lives = 1;
    pushGameLog(gameData, '👼 <b>' + p.name + '</b> ฟื้นคืนชีพจากสกิล!');
    return;
  }
  if (p.skillType === 'revive_once' && !p.intrinsicReviveUsed) {
    p.intrinsicReviveUsed = true;
    p.lives = p.maxLives || 3;
    pushGameLog(gameData, '👼 <b>' + p.name + '</b> คืนชีพด้วยสกิลพิเศษ!');
    return;
  }
  if (p.skillType === 'last_stand' && !p.intrinsicLastStandUsed) {
    p.intrinsicLastStandUsed = true;
    p.lives = 1;
    p.power += (p.skillVal || 500);
    pushGameLog(gameData, '🔥 <b>' + p.name + '</b> ฮึดสู้รอดตายหวุดหวิด!');
    return;
  }
  p.lives = 0;
  p.pos = 1;
  p.power = p.basePower !== undefined ? p.basePower : 200;
  p.items = [];
  p.shield = false;
  p.immune = false;
  p.reflect = false;
  p.tempBuff = 0;
  p.deaths = (p.deaths || 0) + 1;
  pushGameLog(gameData, '💀 <b>' + p.name + '</b> สิ้นชีพ!');
}

function executeActiveSkill(skillName, caster, target, gameState) {
  const skillVal = caster.activeSkillVal > 0 ? caster.activeSkillVal : 1;

  switch (skillName) {
    // ---------------- กลุ่มโจมตีและทำความเสียหาย ----------------
    case 'stun':
      if (target) target.skipTurns = (target.skipTurns || 0) + 1;
      break;
    case 'damage_pierce':
      if (target) target.lives = Math.max(0, target.lives - skillVal);
      break;
    case 'damage_stun':
      if (target) {
        applyDamage(target, skillVal, caster);
        target.skipTurns = (target.skipTurns || 0) + 1;
      }
      break;
    case 'crit_execute':
      if (target) {
        if (target.lives <= 1) {
          target.lives = 0;
        } else {
          applyDamage(target, 2, caster);
        }
      }
      break;
    case 'burn_area':
      if (target && !target.immune && target.skillType !== 'item_immunity') {
        target.status.push('burn');
      }
      break;
    case 'damage_water':
      if (target) applyDamage(target, skillVal > 3 ? 1 : skillVal, caster);
      break;
    case 'poison_blind':
      if (target && !target.immune && target.skillType !== 'item_immunity') {
        target.status.push('poison');
      }
      break;

    // ---------------- กลุ่มป้องกันและฟื้นฟู ----------------
    case 'evade':
      caster.evade = true;
      break;
    case 'immune':
      caster.immune = true;
      break;
    case 'heal_aoe':
      gameState.players.forEach(p => {
        if (p.lives > 0) p.lives = Math.min(p.lives + skillVal, p.maxLives || 3);
      });
      break;
    case 'shield_hp':
      caster.shield = true;
      break;
    case 'reflect':
      caster.reflect = true;
      break;
    case 'cleanse_heal':
      caster.status = caster.status.filter(s => ['auto_revive', 'tank', 'invisible', 'move_boost'].indexOf(s) !== -1);
      caster.skipTurns = 0;
      caster.silenced = 0;
      caster.lives = Math.min(caster.lives + skillVal, caster.maxLives || 3);
      break;
    case 'revive':
      if (caster.status.indexOf('auto_revive') === -1) caster.status.push('auto_revive');
      break;
    case 'summon_tank':
      caster.status.push('tank');
      break;

    // ---------------- กลุ่มป่วนเป้าหมายและบัฟ ----------------
    case 'steal_money':
      if (target && !target.immune && target.skillType !== 'item_immunity') {
        const stealAmount = caster.activeSkillVal > 0 ? caster.activeSkillVal : 200;
        const stolen = Math.min(target.power, stealAmount);
        target.power -= stolen;
        caster.power += stolen;
      }
      break;
    case 'break_buff':
      if (target) {
        target.shield = false;
        target.evade = false;
        target.immune = false;
        target.reflect = false;
        target.status = target.status.filter(s => ['burn', 'poison', 'reduce_dice'].indexOf(s) !== -1);
      }
      break;
    case 'reduce_dice':
      if (target && !target.immune && target.skillType !== 'item_immunity') {
        target.status.push('reduce_dice');
      }
      break;
    case 'steal_buff':
      if (target) {
        if (target.shield) { target.shield = false; caster.shield = true; }
        if (target.evade) { target.evade = false; caster.evade = true; }
        if (target.immune) { target.immune = false; caster.immune = true; }
        if (target.reflect) { target.reflect = false; caster.reflect = true; }
      }
      break;
    case 'shuffle_item': {
      const allItems = [];
      gameState.players.forEach(p => {
        if (p.items) allItems.push.apply(allItems, p.items);
        p.items = [];
      });
      for (let i = allItems.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const tmp = allItems[i];
        allItems[i] = allItems[j];
        allItems[j] = tmp;
      }
      let pIndex = 0;
      allItems.forEach(item => {
        gameState.players[pIndex].items.push(item);
        pIndex = (pIndex + 1) % gameState.players.length;
      });
      break;
    }
    case 'discard_item':
      if (target && target.items && target.items.length > 0 && !target.immune && target.skillType !== 'item_immunity') {
        const randIndex = Math.floor(Math.random() * target.items.length);
        target.items.splice(randIndex, 1);
      }
      break;
    case 'buff_stats':
      caster.power += (caster.activeSkillVal > 0 ? caster.activeSkillVal : 300);
      break;

    // ---------------- กลุ่มควบคุมกระดานและการเคลื่อนที่ ----------------
    case 'pull_target':
      if (target && !target.immune && target.skillType !== 'item_immunity') {
        target.pos = caster.pos;
      }
      break;
    case 'swap_pos':
      if (target && !target.immune && target.skillType !== 'item_immunity') {
        const tempPos = caster.pos;
        caster.pos = target.pos;
        target.pos = tempPos;
      }
      break;
    case 'invis':
      caster.evade = true;
      if (caster.status.indexOf('invisible') === -1) caster.status.push('invisible');
      break;
    case 'move_boost':
      caster.status.push('move_boost');
      break;
    case 'warp_check':
      if (skillVal > 1 && skillVal <= 32) {
        caster.pos = skillVal;
      } else {
        caster.pos = ((caster.pos - 1 + 5) % 32) + 1;
      }
      break;
    case 'cancel_combat':
      caster.evade = true;
      if (caster.status.indexOf('peace') === -1) caster.status.push('peace');
      break;
  }

  return { caster: caster, target: target, gameState: gameState };
}

function applyDamage(target, damage, caster) {
  if (target.immune) return;
  if (target.evade) {
    target.evade = false;
    return;
  }
  if (target.reflect && caster) {
    target.reflect = false;
    caster.lives = Math.max(0, caster.lives - damage);
    return;
  }
  if (target.status && target.status.indexOf('tank') !== -1) {
    target.status = target.status.filter(s => s !== 'tank');
    return;
  }
  if (target.shield) {
    target.shield = false;
    return;
  }
  if (damage > 0) {
    target.lives = Math.max(0, target.lives - damage);
  }
}
