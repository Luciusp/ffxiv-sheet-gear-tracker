function onOpen() {
  // Add the custom menu to the top along with their methods
  SpreadsheetApp.getUi()
      .createMenu("Manage Players")
      .addItem("Add Player", "addPlayerOption")
      .addItem("Remove Player", "removePlayerOption")
      .addToUi();

  // Fetch all api data every time the spreadsheet is opened up to a max
  // of once per hour
  var documentProperties = PropertiesService.getDocumentProperties();
  var lastUpdatedFromApi = documentProperties.getProperty("lastUpdatedFromApi");
  if (lastUpdatedFromApi == null || Date.now() - lastUpdatedFromApi > 3600000)
  {
    documentProperties.setProperty("lastUpdatedFromApi", Date.now());
    documentProperties.setProperty("cachedApiData", JSON.stringify(fetchApiData()));
  }
}

function onEdit()
{
  fetchEtro();
  renderRemaining();
}

function addPlayerOption() {
  var template = HtmlService.createTemplateFromFile("NewPlayer");
  template.data = getData("cachedApiData").jobs.sort();
  var html = template.evaluate()
    .setTitle("New Player Info")
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function processNewPlayer(formObject) {
  var ui = SpreadsheetApp.getUi();
  if (getPlayerList().length >= 8)
  {
    ui.alert("There are already 8 players. Remove one first.");
    return;
  }

  var playerResult = JSON.parse(UrlFetchApp.fetch(
    `https://xivapi.com/character/search?name=${encodeURIComponent(formObject.playerName)}&server=${encodeURIComponent(formObject.serverName)}`).getContentText()
  );

  if (playerResult.Results === null)
  {
    ui.alert("That player cannot be found");
    return;
  }

  var player = {
    "name" : playerResult.Results[0].Name,
    "id" : playerResult.Results[0].ID,
    "job" : formObject.job,
    "etro" : ""
  }
  addPlayer(player);

  // Close sidebar
  var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  ui.showSidebar(html);

  SpreadsheetApp.getActive().toast("Player added.");
  renderPlayers();
}

function addPlayer(player)
{
  var ui = SpreadsheetApp.getUi();
  var storedPlayers = JSON.parse(PropertiesService.getDocumentProperties().getProperty("players"));
  if (storedPlayers === null)
    storedPlayers = [];

  storedPlayers.push(player);

  // Sort players: tank => healer => dps
  var tanks = [];
  var healers = [];
  var dps = [];
  var apiData = getData("cachedApiData");

  storedPlayers.forEach(e => {
    if (apiData.jobCategories.tanks.includes(e.job))
      tanks.push(e);
    else if (apiData.jobCategories.healers.includes(e.job))
      healers.push(e);
    else dps.push(e);
  });

  var sortedPlayers = tanks.concat(healers.concat(dps));
  return PropertiesService.getDocumentProperties().setProperty("players", JSON.stringify(sortedPlayers));
}

function removePlayerOption() {
  var ui = SpreadsheetApp.getUi();
  if (getPlayerList().length <= 0)
  {
    ui.alert("No players to remove!")
    return;
  }
  var template = HtmlService.createTemplateFromFile("RemovePlayer");
  template.data = getPlayerList();
  var html = template.evaluate()
    .setTitle("Remove a Player")
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function processRemovePlayer(formObject) {
  var ui = SpreadsheetApp.getUi();
  removePlayer(formObject.name);

  // Close sidebar
  var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  ui.showSidebar(html);
  SpreadsheetApp.getActive().toast("Player deleted.");
  renderPlayers();
}

function removePlayer(name)
{
  var ui = SpreadsheetApp.getUi();
  var storedPlayers = JSON.parse(PropertiesService.getDocumentProperties().getProperty("players"));

  storedPlayers = storedPlayers.filter(e => e.name.toLowerCase() !== name.toLowerCase());
  return PropertiesService.getDocumentProperties().setProperty("players", JSON.stringify(storedPlayers));
}

function updatePlayer(player)
{
  var ui = SpreadsheetApp.getUi();
  var storedPlayers = JSON.parse(PropertiesService.getDocumentProperties().getProperty("players"));
  var index = storedPlayers.findIndex(i => i.id === player.id);
  if (index !== -1)
  {
    storedPlayers[index] = player;
    return PropertiesService.getDocumentProperties().setProperty("players", JSON.stringify(storedPlayers));
  }
}

function getData(propertyName)
{
  return JSON.parse(PropertiesService.getDocumentProperties().getProperty(propertyName));
}

function getPlayerList()
{
  var storedPlayers = JSON.parse(PropertiesService.getDocumentProperties().getProperty("players"));
  if (storedPlayers === null)
    return [];
    
  var flattenedRange = [];
  storedPlayers.forEach(item =>
  {
    return flattenedRange.push(item.name);
  });

  return flattenedRange;
}

function fetchApiData()
{
  var jobCategories = JSON.parse(UrlFetchApp.fetch("https://xivapi.com/ClassJobCategory?limit=500").getContentText());

  // Modify the item level for each new expansion
  var currentGear = JSON.parse(UrlFetchApp.fetch("https://etro.gg/api/equipment?minItemLevel=570").getContentText());

  // In order to dynamically color the classes of each job, we use https://xivapi.com/ClassJobCategory
  // which maintains all the job categories. The integers below specify the id of each job category
  // 59 = tanks, 64 = healers, 163 = dps
  var tanks = jobCategories.Results.filter(function(item) { return item.ID === 59; })[0].Name.split(" ");
  var healers = jobCategories.Results.filter(function(item) { return item.ID === 64; })[0].Name.split(" ");
  var dps = jobCategories.Results.filter(function(item) { return item.ID === 163; })[0].Name.split(" ");
  var allJobs = jobCategories.Results.filter(function(item) { return item.ID === 169; })[0].Name.split(" ");

  var result = {
    "jobs" : allJobs,
    "jobCategories" : {
      "tanks" : tanks,
      "healers" : healers,
      "dps" : dps
    },
    "gear" : currentGear.map(function (item) {
      return {
        "id" : item.id,
        "name" : item.name,
        "crafted" : item.advancedMelding,
        "slot" : item.slotName
      }
    })
  };
  return result;
}

function renderPlayers()
{
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive().getSheetByName("BiS");
  ss.setColumnWidths(1, 20, 116);
  ss.getRange("A1:V31")
  .clearFormat()
  .removeCheckboxes()
  .setBackground("#434343")
  .setFontColor("white")
  .setBorder(false, false, false, false, false, false)
  .setHorizontalAlignment("center")
  .clearContent();

  var apiData = getData("cachedApiData");
  var players = getData("players");

  var startingRow = 2;
  var startingCol = 2;
  
  var entryHeightInRows = 15;
  var entryWidthInColumns = 5;
  if (players === null || players.length <= 0)
    return;

  for (var i = 0; i < players.length; i++)
  {
    var output = [
      [players[i].job, players[i].name, ""],
      ["Item", "Needs", "Looted"],
      ["Weapon", "", ""],
      ["Helm", "", ""],
      ["Chest", "", ""],
      ["Gloves", "", ""],
      ["Pants", "", ""],
      ["Boots", "", ""],
      ["Earrings", "", ""],
      ["Necklace", "", ""],
      ["Bracelet", "", ""],
      ["Left Ring", "", ""],
      ["Right Ring", "", ""]
    ];

    var currentRow = startingRow + (Math.floor(i / 4) * entryHeightInRows);
    var currentCol = startingCol + (i % 4) * entryWidthInColumns;

    // Merge and format names
    ss.getRange(currentRow, currentCol + 1, 1, 2).mergeAcross();

    var jobBackgroundColor = ""
    var jobFontColor = "";

    if (apiData.jobCategories.tanks.includes(players[i].job))
    {
      jobBackgroundColor = "#073763";
      jobFontColor = "#9DBCCF";
    }
    else if (apiData.jobCategories.healers.includes(players[i].job))
    {
      jobBackgroundColor = "#274E13";
      jobFontColor = "#9FC88D";
    }
    else
    {
      jobBackgroundColor = "#660000";
      jobFontColor = "#EA5345";
    } 

    // Due to the way borders draw, the order of operations is important. Work from outside in
    // Reduce distance between player entries
    ss.setColumnWidths(currentCol + 3, 2, 5);
    // Reduce size of checkbox columns
    ss.setColumnWidth(currentCol + 2, 58);

    // Set the values and borders
    ss.getRange(currentRow, currentCol, entryHeightInRows - 2, 3).setValues(output)
      .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Format inner data
    ss.getRange(currentRow + 2, currentCol + 1, entryHeightInRows - 4).setBorder(false, true, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

    // Format inner header
    ss.getRange(currentRow +1, currentCol, 1, 3).setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID)
      .setFontWeight("bold");
    ss.getRange(currentRow + 1, currentCol - 1).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    ss.getRange(currentRow + 1, currentCol + 3).setBorder(false, true, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Fix bottom row
    ss.getRange(currentRow + entryHeightInRows - 2, currentCol + 1).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Format headers
    ss.getRange(currentRow, currentCol, 1, 3).setBackground(jobBackgroundColor)
      .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK).setFontColor(jobFontColor)
      .setFontWeight("bold");

    // Add checkboxes
    ss.getRange(currentRow + 2, currentCol + 2, entryHeightInRows - 4, 1).insertCheckboxes();

    // Add row banding
    var bandingRange = ss.getRange(currentRow + 2, currentCol, entryHeightInRows - 4, 3);
    bandingRange.getBandings().forEach(banding => banding.remove());
    bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREY, false, false)
      .setFirstRowColor("#666666")
      .setSecondRowColor("#434343");

    // Add etro ingest
    var etroRow = [["Etro Link =>", ""]];
    ss.getRange(currentRow + entryHeightInRows - 2, currentCol, 1, 2)
    .setValues(etroRow)
    .setBorder(true, true, true, true, false, true, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
}

function renderRemaining()
{
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive().getSheetByName("BiS");
  
  var players = getData("players");

  var startingRow = 36;
  var startingCol = 2;

  var playerGearStartingRow = 2;
  var playerGearStartingCol = 2;
  
  if (players === null || players.length <= 0)
    return;

  var output = [
    ["Drops Still Needed", "Twines", "Coatings"]
  ]
  for (var i = 0; i < players.length; i++)
  {

    var currentRow = playerGearStartingRow + (Math.floor(i / 4) * 15);
    var currentCol = playerGearStartingCol + (i % 4) * 5;

    var playerGearInfo = ss.getRange(currentRow + 2, currentCol + 1, 11, 2).getValues();
    var neededTwines = 0;
    var neededCoatings = 0;

    for (var j = 0; j < playerGearInfo.length; j++)
    {
      if (playerGearInfo[j][1])
        continue;
      if (playerGearInfo[j][0].toLowerCase() == "twine")
        neededTwines++;
      if (playerGearInfo[j][0].toLowerCase() == "coating")
        neededCoatings++;
    }

    output.push([players[i].name, neededTwines, neededCoatings]);
  }

  // Clear existing data
  ss.getRange(startingRow, startingCol, players.length + 1, 3)
  .clearFormat()
  .removeCheckboxes()
  .setBackground("#434343")
  .setFontColor("white")
  .setBorder(false, false, false, false, false, false)
  .setHorizontalAlignment("center")
  .clearContent();


  ss.getRange(startingRow, startingCol, players.length + 1, 3)
    .setValues(output)
      .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK)
      .applyRowBanding(SpreadsheetApp.BandingTheme.GREY, false, false)
      .setFirstRowColor("#666666")
      .setSecondRowColor("#434343");

  // Set header borders
  ss.getRange(startingRow, startingCol, 1, 3).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function fetchEtro()
{
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive().getSheetByName("BiS");
  
  var players = getData("players");

  var playerGearStartingRow = 2;
  var playerGearStartingCol = 2;
  
  if (players === null || players.length <= 0)
    return;

  for (var i = 0; i < players.length; i++)
  {
    var currentRow = playerGearStartingRow + (Math.floor(i / 4) * 15);
    var currentCol = playerGearStartingCol + (i % 4) * 5;

    var etroUri = ss.getRange(currentRow + 13, currentCol + 1).getValue();

    // If Etro link has been removed, remove it from the player object
    if (etroUri === null || etroUri === "" && players[i].etro !== null)
    {
      var player = {
        "name" : players[i].name,
        "id" : players[i].id,
        "job" : players[i].job,
        "etro" : ""
      }

      updatePlayer(player);
    }
    // If no Etro link is present or it hasn't been changed, quit out
    if (etroUri === null || etroUri == "" || etroUri == players[i].etro)
      continue;
    
    // Make it an API uri
    var etroApiUri = etroUri.replace("gearset", "api/gearsets");

    var etroGearset = {};

    try
    {
      etroGearset = JSON.parse(UrlFetchApp.fetch(etroApiUri).getContentText());
    }
    catch (error)
    {
      ui.alert(`${players[i].name}'s Etro link is invalid. Fix or delete it to stop this error.`)
    }
    
    var gearList = [
      etroGearset.weapon,
      etroGearset.head,
      etroGearset.body,
      etroGearset.hands,
      etroGearset.legs,
      etroGearset.feet,
      etroGearset.ears,
      etroGearset.neck,
      etroGearset.wrists,
      etroGearset.fingerL,
      etroGearset.fingerR
    ]

    var realGearNames = matchGearIds(gearList);

    var output = [];

    realGearNames.forEach(e => 
    {
      if (e.name.toLowerCase().includes("augmented"))
      {
        switch (e.slot.toLowerCase())
        {
          case "weapon":
          case "head":
          case "body":
          case "hands":
          case "legs":
          case "feet":
            output.push(["Twine"]);
            break;
          case "ears":
          case "neck":
          case "wrists":
          case "fingerL":
          case "fingerR":
            output.push(["Coating"]);
            break;
          default:
            break;
        }
      }
      else if (e.crafted)
        output.push(["Crafted"]);
      else output.push(["Coffer"])
    });

    ss.getRange(currentRow + 2, currentCol + 1, 11, 1).setValues(output);

    var player = {
      "name" : players[i].name,
      "id" : players[i].id,
      "job" : players[i].job,
      "etro" : etroUri
    }

    updatePlayer(player);
  }
}

function matchGearIds(idsToMatch)
{
  var apiData = getData("cachedApiData");

  var result = [];
  
  idsToMatch.forEach(e => {
    var item = apiData.gear.filter(x => x.id == e);
    result.push(item[0]);
  });

  return result;
}
