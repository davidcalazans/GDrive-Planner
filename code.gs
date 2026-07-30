// Constante para o ID da sua planilha
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GDrive Planner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
* Registra logs de ações na aba Timeline_Logs
*/
function registerLog(userKey, action) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Timeline_Logs");
    if (sheet) {
      sheet.appendRow([new Date(), userKey, action]);
    }
  } catch(e) {}
}

/**
* Gerenciamento do Título do Projeto (Geral)
*/
function getProjectTitle() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('PROJECT_TITLE') || 'Construção Planner';
}

function updateProjectTitle(userObj, newTitle) {
  try {
    if (!userObj || (!userObj.isAdm && userObj.permissions && userObj.permissions.length === 0)) {
      return { success: false, message: "Sem permissão para alterar o nome do projeto." };
    }
    PropertiesService.getScriptProperties().setProperty('PROJECT_TITLE', newTitle);
    registerLog(userObj.userKey, `Alterou o nome do projeto para "${newTitle}"`);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
* Validação de Login
*/
function verifyLogin(key) {
  if (!key || key.length < 20) {
    return { success: false, message: "A chave deve ter no mínimo 20 caracteres." };
  }
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("senhas");
  if (!sheet) return { success: false, message: "Aba 'senhas' não encontrada." };
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let sheetKey = String(row[0]).trim();
    if (sheetKey === key) {
      let role = String(row[1]).trim().toLowerCase();
      let isAdm = (role === 'adm');
      let permissions = [];
      if (!isAdm) {
        for (let j = 1; j < row.length; j++) {
          if (row[j]) permissions.push(String(row[j]).trim());
        }
      }
      return {
        success: true,
        userKey: key,
        isAdm: isAdm,
        permissions: permissions
      };
    }
  }
  return { success: false, message: "Chave não encontrada." };
}

/**
* Busca Quadros Permitidos
*/
function getPermittedBoards(userObj) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Quadros");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length > 0) data.shift();
  let permittedBoards = [];
  data.forEach(row => {
    let boardId = String(row[0]);
    let boardName = String(row[1]);
    if (userObj.isAdm) {
      permittedBoards.push({ id: boardId, name: boardName, role: 'adm' });
    } else {
      let userRole = null;
      if (userObj.permissions) {
        userObj.permissions.forEach(perm => {
          if (perm === `lider+${boardName}`) userRole = 'lider';
          else if (perm === `ver+${boardName}` && !userRole) userRole = 'leitor';
          else if (perm === boardName && !userRole) userRole = 'editor';
        });
      }
      if (userRole) {
        permittedBoards.push({ id: boardId, name: boardName, role: userRole });
      }
    }
  });
  return permittedBoards;
}

/**
* CRUD de Quadros
*/
function createNewBoard(userObj, boardName) {
  try {
    if (!userObj || (!userObj.isAdm && userObj.permissions && userObj.permissions.length === 0)) {
      return { success: false, message: "Sem permissão para criar quadros." };
    }
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Quadros");
    const boardId = "B_" + new Date().getTime();
    sheet.appendRow([boardId, boardName]);
    registerLog(userObj.userKey, `Criou o quadro "${boardName}"`);
    return { success: true, boardId: boardId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateBoardName(userObj, role, boardId, newName) {
  try {
    if (role === 'leitor') return { success: false, message: "Apenas leitura." };
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Quadros");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(boardId)) {
        const oldName = data[i][1];
        sheet.getRange(i + 1, 2).setValue(newName);
        registerLog(userObj.userKey, `Renomeou o quadro de "${oldName}" para "${newName}"`);
        return { success: true };
      }
    }
    return { success: false, message: "Quadro não encontrado." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteBoard(userObj, role, boardId) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Apaga o Quadro
    const sheetQuadros = ss.getSheetByName("Quadros");
    const dataQ = sheetQuadros.getDataRange().getValues();
    let boardName = "";
    for (let i = 1; i < dataQ.length; i++) {
      if (String(dataQ[i][0]) === String(boardId)) {
        boardName = dataQ[i][1];
        sheetQuadros.deleteRow(i + 1);
        break;
      }
    }
    
    // Apaga Listas associadas
    const sheetLists = ss.getSheetByName("Listas");
    const dataL = sheetLists.getDataRange().getValues();
    let listIdsToDelete = [];
    for (let i = dataL.length - 1; i >= 1; i--) {
      if (String(dataL[i][1]) === String(boardId)) {
        listIdsToDelete.push(String(dataL[i][0]));
        sheetLists.deleteRow(i + 1);
      }
    }
    
    // Apaga Cartões associados
    if (listIdsToDelete.length > 0) {
      const sheetCards = ss.getSheetByName("Cartoes");
      const dataC = sheetCards.getDataRange().getValues();
      for (let i = dataC.length - 1; i >= 1; i--) {
        if (listIdsToDelete.includes(String(dataC[i][1]))) {
          sheetCards.deleteRow(i + 1);
        }
      }
    }
    registerLog(userObj.userKey, `Excluiu o quadro "${boardName}"`);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
* Busca Listas, Cartões e Checklists
*/
function getBoardListsAndCards(boardId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLists = ss.getSheetByName("Listas");
    if (!sheetLists) return { error: "Aba 'Listas' não encontrada." };
    const listsData = sheetLists.getDataRange().getValues();
    if (listsData.length > 0) listsData.shift();
    let lists = [];
    let listIds = [];
    
    listsData.forEach(row => {
      if (String(row[1]) == String(boardId)) {
        lists.push({
          id: String(row[0]),
          name: String(row[2]),
          position: Number(row[3]) || 0,
          color: String(row[4]) || '#1a73e8'
        });
        listIds.push(String(row[0]));
      }
    });
    lists.sort((a, b) => a.position - b.position);
    
    const sheetCards = ss.getSheetByName("Cartoes");
    if (!sheetCards) return { error: "Aba 'Cartoes' não encontrada." };
    const cardsData = sheetCards.getDataRange().getValues();
    if (cardsData.length > 0) cardsData.shift();
    let cards = [];
    let cardIds = [];
    
    cardsData.forEach(row => {
      if (listIds.includes(String(row[1]))) {
        let rawDeadline = row[4];
        let safeDeadline = "";
        if (rawDeadline instanceof Date) {
          safeDeadline = rawDeadline.toISOString();
        } else if (rawDeadline) {
          safeDeadline = String(rawDeadline);
        }
        let cId = String(row[0]);
        cards.push({
          id: cId,
          listId: String(row[1]),
          title: String(row[2]),
          body: String(row[3]),
          deadline: safeDeadline,
          tag: String(row[5] || "")
        });
        cardIds.push(cId);
      }
    });
    
    // Busca Checklists
    let checklists = [];
    const sheetCheck = ss.getSheetByName("Checklists");
    if (sheetCheck) {
      const checkData = sheetCheck.getDataRange().getValues();
      if (checkData.length > 0) checkData.shift();
      checkData.forEach(row => {
        let itemId = String(row[0]);
        let cardId = String(row[1]);
        let text = String(row[2]);
        let done = row[3] === true || String(row[3]).toUpperCase() === 'TRUE' || row[3] === 1;
        if (cardIds.includes(cardId)) {
          checklists.push({
            id: itemId,
            cardId: cardId,
            text: text,
            done: done
          });
        }
      });
    }
    return { lists: lists, cards: cards, checklists: checklists };
  } catch (e) {
    return { error: "Erro interno: " + e.message };
  }
}

/**
* CRUD de Checklists
*/
function addChecklistItem(userObj, role, cardId, text) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Checklists");
    if (!sheet) return { success: false, message: "Aba 'Checklists' não encontrada." };
    const itemId = "CHK_" + new Date().getTime();
    sheet.appendRow([itemId, String(cardId), String(text), false]);
    return { success: true, itemId: itemId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function toggleChecklistItem(userObj, role, itemId, isDone) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Checklists");
    if (!sheet) return { success: false, message: "Aba 'Checklists' não encontrada." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(itemId)) {
        sheet.getRange(i + 1, 4).setValue(isDone);
        return { success: true };
      }
    }
    return { success: false, message: "Item não encontrado." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteChecklistItem(userObj, role, itemId) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Checklists");
    if (!sheet) return { success: false, message: "Aba 'Checklists' não encontrada." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(itemId)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Item não encontrado." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateListOrder(userObj, role, listOrders) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheetLists = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Listas");
    const data = sheetLists.getDataRange().getValues();
    listOrders.forEach(item => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(item.id)) {
          sheetLists.getRange(i + 1, 4).setValue(item.position);
          break;
        }
      }
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateCardData(userObj, role, cardData) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheetCards = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Cartoes");
    const data = sheetCards.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(cardData.id)) {
        sheetCards.getRange(i+1, 3).setValue(cardData.title);
        sheetCards.getRange(i+1, 4).setValue(cardData.body);
        sheetCards.getRange(i+1, 5).setValue(cardData.deadline);
        sheetCards.getRange(i+1, 6).setValue(cardData.tag);
        registerLog(userObj.userKey, `Editou o cartão: "${cardData.title}"`);
        return { success: true };
      }
    }
    return { success: false, message: "Cartão não encontrado." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function createNewList(userObj, role, boardId, listName, listColor) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheetLists = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Listas");
    const listId = "L_" + new Date().getTime();
    const position = sheetLists.getLastRow();
    sheetLists.appendRow([String(listId), String(boardId), String(listName), position, String(listColor || '#1a73e8')]);
    registerLog(userObj.userKey, `Criou a lista: "${listName}"`);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateListName(userObj, role, listId, newName) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheetLists = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Listas");
    const data = sheetLists.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(listId)) {
        const oldName = data[i][2];
        sheetLists.getRange(i + 1, 3).setValue(newName);
        registerLog(userObj.userKey, `Renomeou a lista de "${oldName}" para "${newName}"`);
        return { success: true };
      }
    }
    return { success: false, message: "Lista não encontrada." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function createNewCard(userObj, role, listId, title, body, deadline, tag) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheetCards = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Cartoes");
    const cardId = "C_" + new Date().getTime();
    sheetCards.appendRow([cardId, listId, title, body || "", deadline || "", tag || ""]);
    registerLog(userObj.userKey, `Criou o cartão: "${title}"`);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function moveCardToList(userObj, role, cardId, newListId) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const sheetCards = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Cartoes");
    const data = sheetCards.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(cardId)) {
        sheetCards.getRange(i+1, 2).setValue(String(newListId));
        return { success: true };
      }
    }
    return { success: false, message: "Cartão não encontrado." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteCard(userObj, role, cardId) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetCards = ss.getSheetByName("Cartoes");
    const data = sheetCards.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(cardId)) {
        const cardTitle = data[i][2];
        sheetCards.deleteRow(i+1);
        // Apaga itens de checklist associados ao cartão
        const sheetCheck = ss.getSheetByName("Checklists");
        if (sheetCheck) {
          const checkData = sheetCheck.getDataRange().getValues();
          for (let c = checkData.length - 1; c >= 1; c--) {
            if (String(checkData[c][1]) === String(cardId)) {
              sheetCheck.deleteRow(c + 1);
            }
          }
        }
        registerLog(userObj.userKey, `Excluiu o cartão: "${cardTitle}"`);
        return { success: true };
      }
    }
    return { success: false, message: "Cartão não encontrado." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteList(userObj, role, listId) {
  try {
    if (role === 'leitor') return { success: false, message: "Sem permissão." };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLists = ss.getSheetByName("Listas");
    const sheetCards = ss.getSheetByName("Cartoes");
    const cardsData = sheetCards.getDataRange().getValues();
    let cardIdsToDelete = [];
    
    for (let i = cardsData.length - 1; i >= 1; i--) {
      if (String(cardsData[i][1]) === String(listId)) {
        cardIdsToDelete.push(String(cardsData[i][0]));
        sheetCards.deleteRow(i + 1);
      }
    }
    
    // Apaga checklists associados às cartas dessa lista
    if (cardIdsToDelete.length > 0) {
      const sheetCheck = ss.getSheetByName("Checklists");
      if (sheetCheck) {
        const checkData = sheetCheck.getDataRange().getValues();
        for (let c = checkData.length - 1; c >= 1; c--) {
          if (cardIdsToDelete.includes(String(checkData[c][1]))) {
            sheetCheck.deleteRow(c + 1);
          }
        }
      }
    }
    
    const listsData = sheetLists.getDataRange().getValues();
    for (let i = 1; i < listsData.length; i++) {
      if (String(listsData[i][0]) === String(listId)) {
        const listName = listsData[i][2];
        sheetLists.deleteRow(i + 1);
        registerLog(userObj.userKey, `Excluiu a lista: "${listName}"`);
        return { success: true };
      }
    }
    return { success: false, message: "Lista não encontrada." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
