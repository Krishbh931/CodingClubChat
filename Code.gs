// ============================================================
// CONFIGURATION
// ============================================================
const SPREADSHEET_ID = "14cnw6UIA4ifLgznC4rLEkbpHDDwNLtGk0-3CEtk6yq4";
const DRIVE_FOLDER_ID = "1pUlFOhJewuQpXLMZo22arsc9FFPtIM__";
const MAX_STORAGE_MB = 14848; // 14.5 GB storage limit
const GLOBAL_ADMIN = "nicole_favret@landon.net";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Landon Middle School Coding Club')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === "Users") {
      sheet.appendRow(["Email", "Username", "PasswordHash", "CreatedAt"]);
    } else if (sheetName === "Messages") {
      sheet.appendRow(["ID", "Timestamp", "Email", "Username", "TargetType", "TargetId", "Message", "FileUrl", "FileName", "FileId", "Reactions", "ReplyToId"]);
    } else if (sheetName === "Groups") {
      sheet.appendRow(["GroupId", "Name", "CreatedBy", "Admins", "Members", "PendingMembers"]);
    } else if (sheetName === "Presence") {
      sheet.appendRow(["Email", "Username", "LastSeen"]);
    }
  }
  return sheet;
}

function checkUserStatus(email) {
  email = email.trim().toLowerCase();
  if (!email.endsWith("@landon.net")) {
    return { success: false, error: "You must use a @landon.net email address." };
  }

  const username = email.split("@")[0].toUpperCase();
  const sheet = getOrCreateSheet("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return { success: true, isNewUser: false, email: email, username: username };
    }
  }

  return { success: true, isNewUser: true, email: email, username: username };
}

function authenticatePasscode(email, passcode, isSetup) {
  email = email.trim().toLowerCase();
  const username = email.split("@")[0].toUpperCase();
  const sheet = getOrCreateSheet("Users");
  const data = sheet.getDataRange().getValues();

  let userRow = -1;
  let storedPasscode = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      userRow = i + 1;
      storedPasscode = data[i][2];
      break;
    }
  }

  if (isSetup) {
    if (!passcode || passcode.length < 3) {
      return { success: false, error: "Passcode must be at least 3 characters." };
    }
    if (userRow === -1) {
      sheet.appendRow([email, username, passcode, new Date().toISOString()]);
    } else {
      sheet.getRange(userRow, 3).setValue(passcode);
    }
    updatePresence(email, username);
    return { success: true, email: email, username: username, isAdmin: (email === GLOBAL_ADMIN) };
  }

  if (storedPasscode === passcode) {
    updatePresence(email, username);
    return { success: true, email: email, username: username, isAdmin: (email === GLOBAL_ADMIN) };
  } else {
    return { success: false, error: "Incorrect passcode. Please try again." };
  }
}

function changeUserPassword(userEmail, currentPasscode, newPasscode) {
  userEmail = userEmail.trim().toLowerCase();
  if (!newPasscode || newPasscode.length < 3) {
    return { success: false, error: "New passcode must be at least 3 characters." };
  }

  const sheet = getOrCreateSheet("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userEmail) {
      const storedPasscode = data[i][2];
      if (storedPasscode !== currentPasscode) {
        return { success: false, error: "Current passcode is incorrect." };
      }
      sheet.getRange(i + 1, 3).setValue(newPasscode);
      return { success: true };
    }
  }
  return { success: false, error: "User record not found." };
}

function adminResetPassword(adminEmail, targetEmail, newPassword) {
  if (adminEmail !== GLOBAL_ADMIN) {
    return { success: false, error: "Unauthorized access." };
  }
  const sheet = getOrCreateSheet("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetEmail.trim().toLowerCase()) {
      sheet.getRange(i + 1, 3).setValue(newPassword);
      return { success: true };
    }
  }
  return { success: false, error: "User not found." };
}

function adminDeleteUser(adminEmail, targetEmail) {
  if (adminEmail !== GLOBAL_ADMIN) {
    return { success: false, error: "Unauthorized access." };
  }
  targetEmail = targetEmail.trim().toLowerCase();
  if (targetEmail === GLOBAL_ADMIN) {
    return { success: false, error: "Cannot delete the primary admin account." };
  }

  const sheet = getOrCreateSheet("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetEmail) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "User not found." };
}

function getAllUsers() {
  const sheet = getOrCreateSheet("Users");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      email: data[i][0],
      username: data[i][1]
    });
  }
  return users;
}

function getUserDMs(userEmail) {
  userEmail = userEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Messages");
  const data = sheet.getDataRange().getValues();

  const dmPartners = {};
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]).toLowerCase() === "dm") {
        const sender = String(data[i][2]).trim().toLowerCase();
        const target = String(data[i][5]).trim().toLowerCase();
        const timestamp = new Date(data[i][1]).getTime() || 0;

        let partner = null;
        if (sender === userEmail && target !== userEmail) partner = target;
        else if (target === userEmail && sender !== userEmail) partner = sender;

        if (partner) {
          if (!dmPartners[partner] || timestamp > dmPartners[partner].lastActive) {
            dmPartners[partner] = {
              email: partner,
              username: partner.split("@")[0].toUpperCase(),
              lastActive: timestamp
            };
          }
        }
      }
    }
  }

  const dms = [];
  for (const email in dmPartners) {
    dms.push(dmPartners[email]);
  }

  dms.sort((a, b) => b.lastActive - a.lastActive);
  return dms;
}

function getChatHistory(targetType, targetId, userEmail) {
  userEmail = userEmail.trim().toLowerCase();
  if (targetId) targetId = targetId.trim().toLowerCase();

  const sheet = getOrCreateSheet("Messages");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const messages = [];
  for (let i = 1; i < data.length; i++) {
    const msgSender = String(data[i][2]).trim().toLowerCase();
    const msgTargetType = String(data[i][4]).trim().toLowerCase();
    const msgTargetId = String(data[i][5]).trim().toLowerCase();

    let include = false;

    if (userEmail === GLOBAL_ADMIN && targetType === "ADMIN_VIEW") {
      include = true;
    } else if (msgTargetType === "channel" && targetType === "channel" && msgTargetId === targetId) {
      include = true;
    } else if (msgTargetType === "group" && targetType === "group" && msgTargetId === targetId) {
      include = true;
    } else if (msgTargetType === "dm" && targetType === "dm") {
      const pair = [msgSender, msgTargetId].sort().join("::");
      const currentPair = [userEmail, targetId].sort().join("::");
      if (pair === currentPair) include = true;
    }

    if (include) {
      let reactions = {};
      try { reactions = JSON.parse(data[i][10] || "{}"); } catch(e){}

      messages.push({
        id: data[i][0],
        timestamp: data[i][1],
        email: data[i][2],
        username: data[i][3],
        targetType: msgTargetType,
        targetId: msgTargetId,
        message: data[i][6],
        fileUrl: data[i][7],
        fileName: data[i][8],
        reactions: reactions,
        replyToId: data[i][11] || ""
      });
    }
  }
  return messages;
}

function getAdminDMMessages(adminEmail, user1, user2) {
  if (adminEmail !== GLOBAL_ADMIN) return [];
  user1 = user1.trim().toLowerCase();
  user2 = user2.trim().toLowerCase();

  const sheet = getOrCreateSheet("Messages");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const targetPair = [user1, user2].sort().join("::");
  const messages = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).toLowerCase() === "dm") {
      const sender = String(data[i][2]).trim().toLowerCase();
      const target = String(data[i][5]).trim().toLowerCase();
      const currentPair = [sender, target].sort().join("::");

      if (currentPair === targetPair) {
        let reactions = {};
        try { reactions = JSON.parse(data[i][10] || "{}"); } catch(e){}

        messages.push({
          id: data[i][0],
          timestamp: data[i][1],
          email: data[i][2],
          username: data[i][3],
          message: data[i][6],
          fileUrl: data[i][7],
          fileName: data[i][8],
          reactions: reactions,
          replyToId: data[i][11] || ""
        });
      }
    }
  }
  return messages;
}

function sendMessage(payload) {
  const sheet = getOrCreateSheet("Messages");
  let fileUrl = "";
  let fileName = "";
  let fileId = "";

  if (payload.file && payload.file.base64) {
    autoCleanupStorage();
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.file.base64),
      payload.file.mimeType,
      payload.file.name
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    fileUrl = file.getUrl();
    fileName = payload.file.name;
    fileId = file.getId();
  }

  const id = payload.id || ("msg_" + new Date().getTime());
  const timestamp = payload.timestamp || new Date().toISOString();

  sheet.appendRow([
    id, timestamp, payload.email.trim().toLowerCase(), payload.username,
    payload.targetType, payload.targetId.trim().toLowerCase(), payload.message || "",
    fileUrl, fileName, fileId, "{}", payload.replyToId || ""
  ]);

  updatePresence(payload.email, payload.username);
  return { success: true, id: id, fileUrl: fileUrl, fileName: fileName };
}

function toggleReaction(msgId, emoji, userEmail) {
  userEmail = userEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Messages");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === msgId) {
      let reactions = {};
      try { reactions = JSON.parse(data[i][10] || "{}"); } catch(e){}

      if (!reactions[emoji]) reactions[emoji] = [];
      const idx = reactions[emoji].indexOf(userEmail);
      if (idx > -1) {
        reactions[emoji].splice(idx, 1);
        if (reactions[emoji].length === 0) delete reactions[emoji];
      } else {
        reactions[emoji].push(userEmail);
      }

      sheet.getRange(i + 1, 11).setValue(JSON.stringify(reactions));
      return { success: true, reactions: reactions };
    }
  }
  return { success: false, error: "Message not found." };
}

function deleteMessage(msgId, userEmail) {
  const sheet = getOrCreateSheet("Messages");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === msgId) {
      if (data[i][2] === userEmail || userEmail === GLOBAL_ADMIN) {
        const fileId = data[i][9];
        if (fileId) {
          try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e){}
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      } else {
        return { success: false, error: "Unauthorized" };
      }
    }
  }
  return { success: false, error: "Message not found." };
}

function createGroup(groupName, creatorEmail) {
  creatorEmail = creatorEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Groups");
  const groupId = "grp_" + new Date().getTime();
  sheet.appendRow([groupId, groupName, creatorEmail, JSON.stringify([creatorEmail]), JSON.stringify([creatorEmail]), JSON.stringify([])]);
  return { success: true, groupId: groupId };
}

function deleteGroup(groupId, actorEmail) {
  actorEmail = actorEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Groups");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupId) {
      const createdBy = String(data[i][2]).toLowerCase();

      if (createdBy === actorEmail || actorEmail === GLOBAL_ADMIN) {
        sheet.deleteRow(i + 1);
        return { success: true };
      } else {
        return { success: false, error: "Only the group creator or Nicole Favret can delete this group." };
      }
    }
  }
  return { success: false, error: "Group not found." };
}

function renameGroup(groupId, newName, actorEmail) {
  actorEmail = actorEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Groups");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupId) {
      const createdBy = String(data[i][2]).toLowerCase();
      const admins = JSON.parse(data[i][3] || "[]").map(e => e.toLowerCase());

      if (createdBy === actorEmail || admins.includes(actorEmail) || actorEmail === GLOBAL_ADMIN) {
        sheet.getRange(i + 1, 2).setValue(newName);
        return { success: true };
      } else {
        return { success: false, error: "Only group admins or creators can rename the group." };
      }
    }
  }
  return { success: false, error: "Group not found." };
}

function addGroupMember(groupId, memberEmail, actorEmail) {
  memberEmail = memberEmail.trim().toLowerCase();
  actorEmail = actorEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Groups");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupId) {
      const createdBy = String(data[i][2]).toLowerCase();
      const admins = JSON.parse(data[i][3] || "[]").map(e => e.toLowerCase());
      let members = JSON.parse(data[i][4] || "[]").map(e => e.toLowerCase());

      if (createdBy === actorEmail || admins.includes(actorEmail) || actorEmail === GLOBAL_ADMIN) {
        if (!members.includes(memberEmail)) {
          members.push(memberEmail);
          sheet.getRange(i + 1, 5).setValue(JSON.stringify(members));
        }
        return { success: true };
      } else {
        return { success: false, error: "Only group admins or creators can add members." };
      }
    }
  }
  return { success: false, error: "Group not found." };
}

function removeGroupMember(groupId, memberEmail, actorEmail) {
  memberEmail = memberEmail.trim().toLowerCase();
  actorEmail = actorEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Groups");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === groupId) {
      const createdBy = String(data[i][2]).toLowerCase();
      let admins = JSON.parse(data[i][3] || "[]").map(e => e.toLowerCase());
      let members = JSON.parse(data[i][4] || "[]").map(e => e.toLowerCase());

      if (createdBy === actorEmail || admins.includes(actorEmail) || actorEmail === GLOBAL_ADMIN) {
        members = members.filter(e => e !== memberEmail);
        admins = admins.filter(e => e !== memberEmail);
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(admins));
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(members));
        return { success: true };
      } else {
        return { success: false, error: "Only group admins or creators can remove members." };
      }
    }
  }
  return { success: false, error: "Group not found." };
}

function getGroups(userEmail) {
  userEmail = userEmail.trim().toLowerCase();
  const sheet = getOrCreateSheet("Groups");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const msgSheet = getOrCreateSheet("Messages");
  const msgData = msgSheet.getDataRange().getValues();
  const groupLastActive = {};

  for (let i = 1; i < msgData.length; i++) {
    if (String(msgData[i][4]).toLowerCase() === "group") {
      const gId = msgData[i][5];
      const ts = new Date(msgData[i][1]).getTime();
      if (!groupLastActive[gId] || ts > groupLastActive[gId]) {
        groupLastActive[gId] = ts;
      }
    }
  }

  const groups = [];
  for (let i = 1; i < data.length; i++) {
    const groupId = data[i][0];
    const name = data[i][1];
    const createdBy = String(data[i][2]).toLowerCase();
    const admins = JSON.parse(data[i][3] || "[]").map(e => e.toLowerCase());
    const members = JSON.parse(data[i][4] || "[]").map(e => e.toLowerCase());

    if (createdBy === userEmail || members.includes(userEmail) || admins.includes(userEmail)) {
      groups.push({
        groupId: groupId,
        name: name,
        createdBy: createdBy,
        admins: admins,
        members: members,
        isAdmin: (userEmail === GLOBAL_ADMIN || createdBy === userEmail || admins.includes(userEmail)),
        isCreator: (userEmail === GLOBAL_ADMIN || createdBy === userEmail),
        lastActive: groupLastActive[groupId] || 0
      });
    }
  }

  groups.sort((a, b) => b.lastActive - a.lastActive);
  return groups;
}

function getAdminDashboardData(adminEmail) {
  if (adminEmail !== GLOBAL_ADMIN) return { success: false };

  const users = getAllUsers();
  const groupSheet = getOrCreateSheet("Groups");
  const groupData = groupSheet.getDataRange().getValues();
  const allGroups = [];
  for (let i = 1; i < groupData.length; i++) {
    allGroups.push({
      groupId: groupData[i][0],
      name: groupData[i][1],
      createdBy: groupData[i][2]
    });
  }

  const msgSheet = getOrCreateSheet("Messages");
  const msgData = msgSheet.getDataRange().getValues();
  const dmPairs = {};

  for (let i = 1; i < msgData.length; i++) {
    if (String(msgData[i][4]).toLowerCase() === "dm") {
      const sender = String(msgData[i][2]).trim().toLowerCase();
      const target = String(msgData[i][5]).trim().toLowerCase();
      const pairKey = [sender, target].sort().join("::");

      if (!dmPairs[pairKey]) {
        dmPairs[pairKey] = {
          user1: sender,
          user1Name: sender.split("@")[0].toUpperCase(),
          user2: target,
          user2Name: target.split("@")[0].toUpperCase()
        };
      }
    }
  }

  return {
    success: true,
    users: users,
    groups: allGroups,
    dms: Object.values(dmPairs)
  };
}

function updatePresence(email, username) {
  const sheet = getOrCreateSheet("Presence");
  const data = sheet.getDataRange().getValues();
  let found = false;
  const now = new Date().getTime();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 3).setValue(now);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([email, username, now]);
  }
}

function autoCleanupStorage() {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFiles();
    let fileList = [];
    let totalSizeMB = 0;

    while (files.hasNext()) {
      const file = files.next();
      const sizeMB = file.getSize() / (1024 * 1024);
      totalSizeMB += sizeMB;
      fileList.push({ file: file, id: file.getId(), created: file.getDateCreated().getTime(), sizeMB: sizeMB });
    }

    fileList.sort((a, b) => a.created - b.created);

    let deletedFileIds = [];
    while (totalSizeMB > MAX_STORAGE_MB && fileList.length > 0) {
      const oldest = fileList.shift();
      totalSizeMB -= oldest.sizeMB;
      deletedFileIds.push(oldest.id);
      oldest.file.setTrashed(true);
    }

    if (deletedFileIds.length > 0) {
      DriveApp.emptyTrash();
      const sheet = getOrCreateSheet("Messages");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const storedFileId = data[i][9];
        if (storedFileId && deletedFileIds.indexOf(storedFileId) !== -1) {
          sheet.getRange(i + 1, 8).setValue("");
          sheet.getRange(i + 1, 9).setValue("[File auto-removed during 14.5GB storage cleanup]");
        }
      }
    }
  } catch (err) {
    Logger.log("Storage cleanup error: " + err.toString());
  }
}
