var games = [];
var properties = PropertiesService.getDocumentProperties();

var games_optimal_numplayers = [];
var games_recommended_numplayers = [];
var games_numplayers = [];

var sss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = sss.getSheetByName('Games');

  var username = sheet.getRange('B2').getValue();
  var includeexpansions = sheet.getRange('B3').getValue();
    var spieleranzahl = parseInt(sheet.getRange("B5").getValue());

var sortparameter = sheet.getRange("B7").getValue();
  var sortall = sheet.getRange("B8").getValue();

function getUserGames() {
  if (includeexpansions === 'no'){
  var bggCollectionUrl = "https://www.boardgamegeek.com/xmlapi2/collection?username=" + username + "&subtype=boardgame&own=1&excludesubtype=boardgameexpansion";
  }else{
    var bggCollectionUrl = "https://www.boardgamegeek.com/xmlapi2/collection?username=" + username + "&subtype=boardgame&own=1";
  }
  var collectionResponse = UrlFetchApp.fetch(bggCollectionUrl + "&stats=1");
  var collectionXmlData = collectionResponse.getContentText();
  var collectionDocument = XmlService.parse(collectionXmlData);
  var collectionRoot = collectionDocument.getRootElement();
  var collectionItems = collectionRoot.getChildren('item');

  clear();

  var gameIds = [];
  for (var i = 0; i < collectionItems.length; i++) {
    var itemId = collectionItems[i].getAttribute('objectid').getValue();
    gameIds.push(itemId);
  }

  var bggThingUrl = "https://www.boardgamegeek.com/xmlapi2/thing?id=" + gameIds.join(',') + "&stats=1";
  var thingResponse = UrlFetchApp.fetch(bggThingUrl);
  var thingXmlData = thingResponse.getContentText();
  var thingDocument = XmlService.parse(thingXmlData);
  var thingRoot = thingDocument.getRootElement();
  var thingItems = thingRoot.getChildren('item');


  for (var i = 0; i < thingItems.length; i++) {
    var item = thingItems[i];
    var game = {};
    game.id = item.getAttribute('id').getValue();
    game.name = item.getChild('name').getAttribute('value').getValue();
    game.minplayers = item.getChild('minplayers').getAttribute('value').getValue();
    game.maxplayers = item.getChild('maxplayers').getAttribute('value').getValue();
    game.numplayers = [game.minplayers, game.maxplayers];
    game.numplayers_display = game.numplayers[0] + ' - ' + game.numplayers[1];
    game.weight = item.getChild('statistics').getChild('ratings').getChild('averageweight').getAttribute('value').getValue();
    game.weight = parseFloat(game.weight).toFixed(1);
    game.rating = item.getChild('statistics').getChild('ratings').getChild('average').getAttribute('value').getValue();
    game.rating = parseFloat(game.rating).toFixed(1);
    game.year = item.getChild('yearpublished').getAttribute('value').getValue();
    game.mintime = item.getChild('minplaytime').getAttribute('value').getValue();
    game.maxtime = item.getChild('maxplaytime').getAttribute('value').getValue();
    if (game.mintime === game.maxtime){
        game.time = game.mintime;
    }else{
    game.time = game.mintime + ' - ' + game.maxtime;
    }
    var polls = item.getChildren('poll');
    for (var j = 0; j < polls.length; j++) {
      var poll = polls[j];
      if (poll.getAttribute('name').getValue() === 'suggested_numplayers') {
        var results = poll.getChildren('results');
        for (var k = 0; k < results.length; k++) {
          var result = results[k];
          var numplayers = parseInt(result.getAttribute('numplayers').getValue());
          var playerPollResults = result.getChildren('result');

          var bestplayers = 0;
          var recommendedvotes = 0;
          var notrecommendedvotes = 0;
          for (var l = 0; l < playerPollResults.length; l++) {
            var playerPollResult = playerPollResults[l];
            var numvotes = playerPollResult.getAttribute('numvotes').getValue();
            var value = playerPollResult.getAttribute('value').getValue();
            if (value === 'Best') {
              var bestplayers = numvotes;
            } else if (value === 'Recommended') {
              var recommendedvotes = numvotes;
            } else if (value === 'Not Recommended') {
              var notrecommendedvotes = numvotes;
            }
            var bestplayers_perc = 0;
            var recommendedvotes_perc = 0;
            var notrecommendedvotes_perc = 0;
            var totalvotes = bestplayers + recommendedvotes + notrecommendedvotes;
            var bestplayers_perc = (bestplayers * 100 / totalvotes);
            var recommendedvotes_perc = (recommendedvotes * 100 / totalvotes);
            var notrecommendedvotes_perc = (notrecommendedvotes * 100 / totalvotes);
          }
          if (bestplayers_perc > recommendedvotes_perc && bestplayers_perc > notrecommendedvotes_perc) {
            if (!game.optimal_numplayers) {
              game.optimal_numplayers = [numplayers];
            } else {
              game.optimal_numplayers.push(numplayers);
            }
          }
          if (bestplayers_perc > notrecommendedvotes_perc || recommendedvotes_perc > notrecommendedvotes_perc) {
            if (!game.recommended_numplayers) {
              game.recommended_numplayers = [numplayers];
            } else {
              game.recommended_numplayers.push(numplayers);
            }
          }
        }
      }
    }
    if (game.optimal_numplayers && game.optimal_numplayers.length > 1) {
      game.optimal_numplayers_display = game.optimal_numplayers[0] + ' - ' + game.optimal_numplayers[game.optimal_numplayers.length - 1];
    } else if (game.optimal_numplayers && game.optimal_numplayers.length === 1) {
      game.optimal_numplayers_display = game.optimal_numplayers[0];
    } else {
      game.optimal_numplayers_display = '-';
    }

    if (game.recommended_numplayers && game.recommended_numplayers.length > 1) {
      game.recommended_numplayers_display = game.recommended_numplayers[0] + ' - ' + game.recommended_numplayers[game.recommended_numplayers.length - 1];
    } else if (game.recommended_numplayers && game.recommended_numplayers.length === 1) {
      game.recommended_numplayers_display = game.recommended_numplayers[0];
    } else {
      game.recommended_numplayers_display = '-';
    }

    games.push(game);
    properties.setProperty("loadedGames", JSON.stringify(games));
  }
  sortgames();
}
function sortgames() {
  var arrayString = properties.getProperty("loadedGames");
  var games = JSON.parse(arrayString);

  clear();
  for (var i = 0; i < games.length; i++) {
    var game = games[i];
    if (game.optimal_numplayers) {
      if (spieleranzahl >= game.optimal_numplayers[0] && spieleranzahl <= game.optimal_numplayers[game.optimal_numplayers.length - 1]) {
        if (spieleranzahl >= game.recommended_numplayers[0] && spieleranzahl <= game.recommended_numplayers[game.recommended_numplayers.length - 1]) {
          if (spieleranzahl >= game.numplayers[0] && spieleranzahl <= game.numplayers[game.numplayers.length - 1]) {

            games_optimal_numplayers.push(game);
          }
        }
      }

    }
  } properties.setProperty("games_optimal_numplayers", JSON.stringify(games_optimal_numplayers));
  for (var i = 0; i < games.length; i++) {
    var game = games[i];
    if (game.recommended_numplayers) {
      if (!game.optimal_numplayers || spieleranzahl < game.optimal_numplayers[0] || spieleranzahl > game.optimal_numplayers[game.optimal_numplayers.length - 1]) {
        if (spieleranzahl >= game.recommended_numplayers[0] && spieleranzahl <= game.recommended_numplayers[game.recommended_numplayers.length - 1]) {
          if (spieleranzahl >= game.numplayers[0] && spieleranzahl <= game.numplayers[game.numplayers.length - 1]) {

            games_recommended_numplayers.push(game);
          }
        }
      }
    }
  } properties.setProperty("games_recommended_numplayers", JSON.stringify(games_recommended_numplayers));
  for (var i = 0; i < games.length; i++) {
    var game = games[i];
    if (game.numplayers) {
      if (!game.optimal_numplayers || spieleranzahl < game.optimal_numplayers[0] || spieleranzahl > game.optimal_numplayers[game.optimal_numplayers.length - 1]) {
        if (!game.recommended_numplayers || spieleranzahl < game.recommended_numplayers[0] || spieleranzahl > game.recommended_numplayers[game.recommended_numplayers.length - 1]) {
          if (spieleranzahl >= game.numplayers[0] && spieleranzahl <= game.numplayers[game.numplayers.length - 1]) {

            games_numplayers.push(game);
          }
        }
      }
    }
  } properties.setProperty("games_numplayers", JSON.stringify(games_numplayers));


  sort_by();
}



function sort_by() {
  

  var arrayString = properties.getProperty("games_optimal_numplayers");
  var games_optimal_numplayers = JSON.parse(arrayString);
  var arrayString = properties.getProperty("games_recommended_numplayers");
  var games_recommended_numplayers = JSON.parse(arrayString);
  var arrayString = properties.getProperty("games_numplayers");
  var games_numplayers = JSON.parse(arrayString);

  switch (sortparameter) {
    case 'name':
      sortparameter = "name";
      break;
    case 'published':
      sortparameter = "year";
      break;
    case 'rating':
      sortparameter = "rating";
      break;
    case 'time':
      sortparameter = "mintime";
      break;
    case 'weight':
      sortparameter = "weight";
      break;
  }

  clear();

  if (sortall == 'yes') {
    var games = [];
    for (var i = 0; i < Math.max(games_optimal_numplayers.length, games_recommended_numplayers.length, games_numplayers.length); i++) {
      if (i < games_optimal_numplayers.length) {
        games.push(games_optimal_numplayers[i]);
      }
      if (i < games_recommended_numplayers.length) {
        games.push(games_recommended_numplayers[i]);
      }
      if (i < games_numplayers.length) {
        games.push(games_numplayers[i]);
      }
    }
    games.sort(function (a, b) {
      return a[sortparameter] - b[sortparameter];
    });
    for (var i = 0; i < games.length; i++) {
      var game = games[i];
      addgames(game);
    }
  } else {
    if (games_optimal_numplayers && games_optimal_numplayers.length > 0) {
      games_optimal_numplayers.sort(function (a, b) {
        return a[sortparameter] - b[sortparameter];
      });
      for (var i = 0; i < games_optimal_numplayers.length; i++) {
        var game = games_optimal_numplayers[i];
        addgames(game);
      }
      sheet.getRange("I"+sheet.getLastRow()).setValue("↑ best for Playercount") 
    sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).setBorder(null, null, true, null, null, null);
    }
    
    if (games_recommended_numplayers && games_recommended_numplayers.length > 0) {
      games_recommended_numplayers.sort(function (a, b) {
        return a[sortparameter] - b[sortparameter];
      });
      for (var i = 0; i < games_recommended_numplayers.length; i++) {
        var game = games_recommended_numplayers[i];
        addgames(game);
      }
      sheet.getRange("I"+sheet.getLastRow()).setValue("↑ recommended for Playercount") 
    sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).setBorder(null, null, true, null, null, null);
    }
    
    if (games_numplayers && games_numplayers.length > 0) {
      games_numplayers.sort(function (a, b) {
        return a[sortparameter] - b[sortparameter];
      });
      for (var i = 0; i < games_numplayers.length; i++) {
        var game = games_numplayers[i];
        addgames(game);
      }
      sheet.getRange("I"+sheet.getLastRow()).setValue("↑ possible for Playercount")
    }
  }
  
  colorcodeweight();
}

function addgames(game) {
  sheet.appendRow([
    '=HYPERLINK("https://boardgamegeek.com/boardgame/' + game.id + '","' + game.name + '")',
    game.year,
    game.rating,
    game.time,
    game.weight,
    game.numplayers_display,
    game.recommended_numplayers_display,
    game.optimal_numplayers_display
  ]);
}

function colorcodeweight() {
  var rules = sheet.getConditionalFormatRules();
  var rulea = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(E7>=1, E7<=2)")
    .setBackground("#93c47d")
    .setRanges([sheet.getRange("E7:E")])
    .build();
  rules.push(rulea);
  var ruleb = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(E7>2, E7<=3)")
    .setBackground("#f6b26b")
    .setRanges([sheet.getRange("E7:E")])
    .build();
  rules.push(ruleb);
  var rulec = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(E7>3, E7<=5)")
    .setBackground("#e06666")
    .setRanges([sheet.getRange("E7:E")])
    .build();
  rules.push(rulec);
  sheet.setConditionalFormatRules(rules);
}

function clear() {
  var firstrowwithcontent = 11;
  sheet.getRange("A"+firstrowwithcontent+":I").clear();
   sheet.getRange("D"+firstrowwithcontent+":D").setNumberFormat("@").setHorizontalAlignment("center");
  sheet.getRange("F"+firstrowwithcontent+":H").setNumberFormat("@").setHorizontalAlignment("center");
}