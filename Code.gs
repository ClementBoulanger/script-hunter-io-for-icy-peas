function welcomeMessage() {

  let ui = SpreadsheetApp.getUi();
  ui.alert("Bienvenue, pour pouvoir trouver les adresses mails correspondantes aux noms, prénoms et nom de domaines entrées dans les trois premières colonnes, allez dans le menu déroulant Hunter.io et cliquez sur Find Emails. Attention il ne faut pas modifier les noms des colonnes pour que le script puisse s'effectuer correctement");
}


//Custom menu creation at the opening of the file - Création du menu custom à l'ouverture du fichier

function onOpen() {

  welcomeMessage();

  SpreadsheetApp.getUi()
      .createMenu('Hunter.io')
      .addItem('Set API Key', 'showApiKeySidebar')
      .addItem('Find Emails', 'processEmailRequest')
      .addToUi();
}

//Calling the HTML.html file to generate the Sidebar when clicking on 'Set API Key' - Appel du fichier HTML.html pour générer la barre latérale quand on clique sur 'Set API Key'

function showApiKeySidebar() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('HTML')
      .setTitle('Enter Hunter.io API Key')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// Saving the API Key - Sauvegarde de la clé API 
function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
}

//Retrieve the API Key from the Google Apps Script properties, construct the API Request URL,Make the API call and Process the API response, returns the found email address or raises an error if no email is found - Récupére la clé API des propriétés du Google Apps Script, création de l'URL de requête API, faire l'appel de l'API et traiter la réponse de l'API, et retourne l'adresse email trouvée ou une erreur si l'email aucun email n'est trouvé

function findEmail(firstName, lastName, company) {
  Logger.log("findEmail called with: " + firstName + ", " + lastName + ", " + company);
  let apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  if (!apiKey) {
    throw new Error('API key not set. Please use the Hunter.io menu to set the API key.');
  }
  
  let url = 'https://api.hunter.io/v2/email-finder?domain=' + encodeURIComponent(company) +
            '&first_name=' + encodeURIComponent(firstName) +
            '&last_name=' + encodeURIComponent(lastName) +
            '&api_key=' + apiKey;
  
  try {
    let response = UrlFetchApp.fetch(url);
    let result = JSON.parse(response.getContentText());
    
    Logger.log("API response: " + JSON.stringify(result));
    if (result.data && result.data.email) {
      return result.data.email;
    } else if (result.errors && result.errors.length > 0) {
      throw new Error(result.errors[0].details);
    } else {
      return 'Email not found';
    }
  } catch (e) {
    // Mask the API key in the error message
    let maskedUrl = url.replace(apiKey, 'API_KEY_HIDDEN');
    let maskedError = e.message.replace(apiKey, 'API_KEY_HIDDEN');
    Logger.log('Error with masked URL: ' + maskedUrl);
    throw new Error(maskedError);
  }
}


/* 
Function designed to read the active sheet, search for the columns "First Name", "Last Name", "Company", and "Email" by their header names. For each row, it uses the values from these columns to call findEmail. It then writes the result (the found email or an error message) in the "Email" column of the same row. :
-
Fonction destinée à lire la feuille active, chercher les colonnes  "First Name", "Last Name", "Company" et "Email" par leur nom d'en-tête. Pour chaque ligne, elle utilise les valeurs des colonnes pour appeler findEmail. Et elle écrit le résultat (l'email trouvé ou un message d'erreur) dans la colonne "Email" de la même ligne. :
*/
function processEmailRequest() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  
  // Initialize column indices - Initialiser les indices des colonnes
  let firstNameCol = -1;
  let lastNameCol = -1;
  let companyCol = -1;
  let emailCol = -1;
  
  // Find the column indices - Trouver les indices des colonnes
  for (let i = 0; i < headers.length; i++) {
    let header = headers[i].toString().trim().toLowerCase();
    if (header === "first name") firstNameCol = i;
    else if (header === "last name") lastNameCol = i;
    else if (header === "company") companyCol = i;
    else if (header === "email") emailCol = i;
  }
  
  // Check if all required columns are found - Vérifier si toutes les colonnes requises sont trouvées
  if (firstNameCol === -1 || lastNameCol === -1 || companyCol === -1 || emailCol === -1) {
    let missingColumns = [];
    if (firstNameCol === -1) missingColumns.push("First Name");
    if (lastNameCol === -1) missingColumns.push("Last Name");
    if (companyCol === -1) missingColumns.push("Company");
    if (emailCol === -1) missingColumns.push("Email");
    
    let errorMessage = "One or more required columns not found. Missing: " + missingColumns.join(", ") + ". Please make sure you have these columns with exact names.";
    SpreadsheetApp.getUi().alert(errorMessage);
    return;
  }
  
  // Process each row - Traiter chaque ligne
  for (let i = 1; i < data.length; i++) {
    let firstName = data[i][firstNameCol];
    let lastName = data[i][lastNameCol];
    let company = data[i][companyCol];
    
    if (firstName && lastName && company) {
      // Trim whitespace from inputs
      firstName = firstName.toString().trim();
      lastName = lastName.toString().trim();
      company = company.toString().trim();
      
      try {
        let email = findEmail(firstName, lastName, company);
        sheet.getRange(i + 1, emailCol + 1).setValue(email);
      } catch (e) {
        // Use a generic error message that doesn't expose sensitive information
        let safeErrorMessage = "Error: Unable to find email. Please check the input and try again.";
        if (e.message.includes("API key not set")) {
          safeErrorMessage = "Error: API key not set. Please use the Hunter.io menu to set the API key.";
        } else if (e.message.includes("Email not found")) {
          safeErrorMessage = "Error: Email not found for this combination.";
        }
        sheet.getRange(i + 1, emailCol + 1).setValue(safeErrorMessage);
        
        // Log the full error for debugging, but not in the sheet
        Logger.log("Full error for row " + (i + 1) + ": " + e.message);
      }
    }
  }
}



