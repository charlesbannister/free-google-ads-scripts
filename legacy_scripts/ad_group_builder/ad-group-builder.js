  /**
  * AutomatingAdWords.com - Ad Group Builder
  *
  * Go to automatingadwords.com for installation instructions and advice
  *
  * V 1.6.0 - Added H3 and D2 support (the settings sheet needs updating)
  * Version: 1.6.0
  **/
    
  //your spreadsheet URL
  var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1oQ7pCFk8fAMLwv7JOgK-kmS9aNfR5Jfyn1tYLYwZ_Ms/edit#gid=0";
  //your sheet (tab) name
  var SHEET_NAME = "Builder";

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  //Cell Locations - Only to be changed if the sheet changes
  var firstAdGroupRow = 9; //row of the first AdGroup, 9 is default
  var urlColumn = ss.getRangeByName("url").getColumn();
  var adGroupColumn = ss.getRangeByName("adGroup").getColumn();
  var keywordsColumn = ss.getRangeByName("keywords").getColumn()
  var negativeKeywordColumn = ss.getRangeByName("negativeKeywords").getColumn()
  var headline1Column = ss.getRangeByName("headline1").getColumn()
  var headline2Column = ss.getRangeByName("headline2").getColumn()
  var headline3Column = ss.getRangeByName("headline3").getColumn()
  var displayUrl1Column = ss.getRangeByName("path1").getColumn()
  var displayUrl2Column = ss.getRangeByName("path2").getColumn()
  var descriptionColumn = ss.getRangeByName("description").getColumn()
  var description2Column = ss.getRangeByName("description2").getColumn()

  //OPTIONS - update on sheet
  var campaignName = myTrim(sheet.getRange(2, 2).getValue());
  var createGroups = sheet.getRange(3, 2).getValue();
  var createAds_bool = sheet.getRange(4, 2).getValue();
  var createKeywords_bool = sheet.getRange(5, 2).getValue();
  var exactBid = sheet.getRange(2,4).getValue();
  var phraseBid = sheet.getRange(3, 4).getValue();
  var broadBid = sheet.getRange(4, 4).getValue();
  var addUrlsToKeywords = sheet.getRange(5, 4).getValue();
  var pauseExistingKeywords_bool = sheet.getRange(6, 2).getValue();
  var adGroupLabel = myTrim(sheet.getRange(2, 6).getValue());
  var adLabel = myTrim(sheet.getRange(3, 6).getValue());
  var keywordLabel = myTrim(sheet.getRange(4, 6).getValue());
  var keywordGroups = [];
  var adGroupsToAdd = [];

  var DEV_MODE = true;

  function main() {

    runChecks()
    
    //create the labels if they don't exist
    createLabel(adGroupLabel)
    createLabel(adLabel)
    createLabel(keywordLabel)
    
    if(createGroups == true){
      createAdGroups();
    }
    
    if(createAds_bool == true){
      createAds();
    }

    if(createKeywords_bool){
      createKeywords();
    }
    if(pauseExistingKeywords_bool == true){
      pauseExistingKeywords();
    }
    msg("All done!")
    
  }

  function runChecks(){

    if(!checkCampaignExists(campaignName)){
      throw("The campaign specified ('"+campaignName+"') doesn't exist within the account. Please check the settings.")
    }

  }

  function checkCampaignExists(campaignName){

    var cols = ["CampaignName"]
    
      var reportName = "CAMPAIGN_PERFORMANCE_REPORT"
      var where = " where CampaignName = '" + campaignName + "'";
    
      var OPTIONS = {
        includeZeroImpressions: true,
      };
      
        var query = ['select', cols.join(','), 'from', reportName, where].join(' ');
      
        var reportIter = AdWordsApp.report(query, OPTIONS).rows();
        if(reportIter.hasNext()) {
          return true
        }else{
          return false
        }

  }

  
  function checkAdGroupExists(campaignName,adGroupName){

    var cols = ["CampaignName", "AdGroupName"]
    
      var reportName = "ADGROUP_PERFORMANCE_REPORT"
      var where = " where CampaignName = '" + campaignName + "'";
      where += " and AdGroupName = '" + adGroupName + "'";
    
      var OPTIONS = {
        includeZeroImpressions: true,
      };
      
        var query = ['select', cols.join(','), 'from', reportName, where].join(' ');
      
        var reportIter = AdWordsApp.report(query, OPTIONS).rows();
        if(reportIter.hasNext()) {
          return true
        }{
          return false
        }

  }

  function createLabel(labelName){
    if(labelName==""){return}
    
          var labels = AdWordsApp.labels().get();
          var iter = 0;
          while(labels.hasNext()){
            var label = labels.next();
            if(label.getName() == labelName){
              iter++;
            }
          }
          if(iter==0){
          AdWordsApp.createLabel(labelName); 
          }
  }

  function pauseExistingKeywords(){
    msg("Pausing existing keywords...");
    //msg("keyword groups: " + keywordGroups);
    
    for(var kwGroup_i in keywordGroups){
      var keywordsToAdd = keywordGroups[kwGroup_i];
    
    //msg("keywords to pause: " + keywordsToAdd);
    //msg("excluding this adgroup: " + adGroupsToAdd[kwGroup_i]);
    
    var keywords = AdWordsApp.keywords()
    .withCondition("Status = ENABLED")
    .withCondition('CampaignName = "'+campaignName+'"')
    .withCondition('AdGroupName != "'+adGroupsToAdd[kwGroup_i]+'"')
    .get();
    while(keywords.hasNext()){
    var keyword = keywords.next();

      if(keywordsToAdd.indexOf(keyword.getText())>-1){
        //msg("Keyword to pause (after adGroup check): " + keyword.getText());
        //msg(keyword.getAdGroup().getName());      
        if(adGroupsToAdd[kwGroup_i].indexOf(keyword.getAdGroup().getName())==-1){
        //msg("Keyword to pause, adgroup check done: " + keyword.getText());
        keyword.pause();
        }
      }
    }
  }
  }

  function createAds(){
    msg("Creating ads...");
    var row = firstAdGroupRow;
    while(sheet.getRange(row, adGroupColumn).getValue()){
      
      var adGroupName = sheet.getRange(row, adGroupColumn).getValue(); 
      var url = sheet.getRange(row, urlColumn).getValue();  
      var headline1 = sheet.getRange(row, headline1Column).getValue();
      var headline2 = sheet.getRange(row, headline2Column).getValue();
      var headline3 = sheet.getRange(row, headline3Column).getValue();
      var path1 = sheet.getRange(row, displayUrl1Column).getValue();
      var path2 = sheet.getRange(row, displayUrl2Column).getValue();
      var description = sheet.getRange(row, descriptionColumn).getValue();
      var description2 = sheet.getRange(row, description2Column).getValue();
      var fullAd = url+headline1+headline2+headline3+path1+path2+description+description2;
      
    // msg("campaign name: " + campaignName);
      var adGroupIterator = AdWordsApp.adGroups()
      .withCondition('Name = "'+adGroupName+'"')
      .withCondition('CampaignName = "'+campaignName+'"').get();
      
      if (adGroupIterator.hasNext()) {
        var adGroup = adGroupIterator.next();
        
        var currentAds = [];
      var ads = adGroup.ads().withCondition("Status = ENABLED").withCondition("Type = EXPANDED_TEXT_AD").get();
      while(ads.hasNext()){
      var ad = ads.next();

        var fullCurrentAd = ad.urls().getFinalUrl() + ad.getHeadlinePart1() + ad.getHeadlinePart2() + ad.getHeadlinePart3() + ad.getPath1() + ad.getPath2() +ad.getDescription()+ad.getDescription2();
        currentAds.push(fullCurrentAd);

      }
      
      if(currentAds.indexOf(fullAd)>-1){
        //msg("The ad already exists in the Ad Group so will not be created again");      
      }else{
      
      var build = adGroup.newAd().expandedTextAdBuilder()
          .withHeadlinePart1(headline1)
          .withHeadlinePart2(headline2)
          .withHeadlinePart3(headline3)
          .withDescription1(description)
          .withDescription2(description2)
          .withPath1(path1)
          .withPath2(path2)
          .withFinalUrl(url)
          .build()
      var result = build.getResult()
      
      if(build.isSuccessful()){
        if(adLabel!=""){
          result.applyLabel(adLabel);
        }
      }else{
        msg("There was a problem creating an ad. Please see the change logs for details (view errors).")
      }     
      
    }
    }
      row++;
    }  
    
  }

  function createKeywords(){
    msg("Creating keywords...");

      var row = firstAdGroupRow;
      var currentKeywords = [];

      while(sheet.getRange(row, adGroupColumn).getValue()){


      function trimToLowerCase(x){
        return x.trim().toLowerCase()
      }
      var keywords = sheet.getRange(row, keywordsColumn).getValue()
      .split(",")
      .map(trimToLowerCase);


      var negativeKeywords = sheet.getRange(row, negativeKeywordColumn).getValue()
      .split(",")
      .map(trimToLowerCase);

      var url = sheet.getRange(row, urlColumn).getValue();

      var adGroupName = sheet.getRange(row, adGroupColumn).getValue(); 
      adGroupsToAdd.push(adGroupName);
      
      var keywordsToAdd = [];
      var adGroupIterator = AdWordsApp.adGroups()
      .withCondition('Name = "'+adGroupName+'"')
      .withCondition('CampaignName = "'+campaignName+'"')
      .get();
      
    if (adGroupIterator.hasNext()) {
      var adGroup = adGroupIterator.next();

      //add negs
      for(var neg_i in negativeKeywords){
        if(negativeKeywords[neg_i]!=""){
          adGroup.createNegativeKeyword(negativeKeywords[neg_i])
        }
      }
      
      currentKeywords = [];
      var currentKeywordsGet = adGroup.keywords().withCondition("Status = ENABLED").get();
      while(currentKeywordsGet.hasNext()){
      currentKeywords.push(currentKeywordsGet.next().getText()); 
      }


      //msg("currentKeywords: " +currentKeywords);
      for(var keyword in keywords){
    
        
          var kwToAdd = myTrim(keywords[keyword]);

          if(exactBid != "" && exactBid > 0){
            //check if the keyword already exists in the adGroup, add it if not
            var exactKW = '['+kwToAdd+']';
            if(currentKeywords.indexOf(exactKW)>-1){
            msg("The keyword "+exactKW+" already exists in the Ad Group so will not be added or modified");
            }else{
            addKeyword(exactKW, url, exactBid);
            }
          }        
          if(phraseBid != "" && phraseBid > 0){
            var phraseKW = '"'+kwToAdd+'"';
              //check if the keyword already exists in the adGroup, add it if not
            if(currentKeywords.indexOf(phraseKW)>-1){
            msg("The keyword "+phraseKW+" already exists in the Ad Group so will not be added or modified");
            }else{
            addKeyword(phraseKW, url, phraseBid);
            }
          }
          if(broadBid != "" && broadBid > 0){
          var broadSplit = kwToAdd.split(" ");
            var kw = "";
            for(var b in broadSplit){

              if(b == broadSplit.length-1){
                kw += "+"+broadSplit[b]+"";
              }else{
                kw += "+"+broadSplit[b]+" ";
              }
            }
            //check if the keyword already exists in the adGroup, add it if not
            if(currentKeywords.indexOf(kw)>-1){
              msg("The keyword "+kw+" already exists in the Ad Group so will not be added or modified");
            }else{
            addKeyword(kw, url, broadBid); 
            }
          }
          
        
        
      }

    }
      keywordGroups[row-9] = keywordsToAdd;
      
      row++;
    }


    function addKeyword(keyword, url, bid){      
        
      if(createKeywords_bool==true){
        //msg("adding keyword " + keyword.toLowerCase())        
        if(url!="" && addUrlsToKeywords){
        var result = adGroup.newKeywordBuilder()
        .withText(keyword.toLowerCase())
        .withCpc(bid)                         
        .withFinalUrl(url)
        .build().getResult();
      
        }else{
        var result = adGroup.newKeywordBuilder()
        .withText(keyword.toLowerCase())
        .withCpc(bid)                         
        .build()
        .getResult();
  
        }
  
        if(keywordLabel!=""){
          result.applyLabel(keywordLabel);
        }
  
      }
  
      keywordsToAdd.push(keyword.toLowerCase());
    }
    
  }

  

  function createAdGroups(){
  msg("Creating adgroups...");
    var defaultBid = sheet.getRange(2, 4).getValue();
  // msg("Campaign name: " + campaignName);
    var row = firstAdGroupRow;
    while(sheet.getRange(row, adGroupColumn).getValue()){

      var adGroupName = sheet.getRange(row, adGroupColumn).getValue();
      var campaignSelector = AdWordsApp
      .campaigns()
      .withCondition('Name = "' + campaignName +'"');
      var campaignIterator = campaignSelector.get();

  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
        
    if(campaign.getName() != campaignName){
      continue;
    }
    
    var agMaxCPC = .5//default
    if(exactBid > 0 && exactBid != ""){
        agMaxCPC = exactBid;
    }else if(phraseBid > 0 && phraseBid != ""){
        agMaxCPC = phraseBid
    }else{
        agMaxCPC = broadBid
    }

    if(!checkAdGroupExists(campaignName,adGroupName)){

      var adGroupBuilder = campaign.newAdGroupBuilder();
      var adGroupOperation = adGroupBuilder
      .withName(adGroupName)
      .withCpc(agMaxCPC)
      .build();
      var adGroup = adGroupOperation.getResult();
      if(adGroupOperation.isSuccessful()){
        if(adGroupLabel!=""){
          adGroup.applyLabel(adGroupLabel)
        }
      }else{
        msg("There was a problem creating the Ad Group '"+adGroupName+"'. Please check the change logs (view errors).")
      }

    }else{
      msg("The Ad Group '"+adGroupName+"' already exists so will not be created")
    }

  }
      
      row++;
      
    }  
  }

  function log(log){
    if(!DEV_MODE)return
    Logger.log(log)
  }
  function msg(log){
    Logger.log(log)
  }

  function myTrim(x) {
    return x.replace(/^\s+|\s+$/gm,'');
  }