/**
 * @fileoverview Automates the processing of a Google Form submission for Purchase orders.
 * It creates a folder in a Shared Drive, renames and moves uploaded files based on a template, and sends a notification email.
 */

// =========================================================================================
// SCRIPT CONFIGURATION
// All user-configurable settings are located in this section for easy management.
// =========================================================================================

const CONFIG = {
  // The unique ID of the target Shared Drive where the final folder will be created.
  // To find this, open the Shared Drive in your browser; the ID is the string of characters in the URL.
  TARGET_SHARED_DRIVE_ID: 'DriveID',
  
  // A comma-separated list of email addresses that will receive the notification email.
  EMAIL_RECIPIENTS: 'email1@email.com,email2@email.com',
  
  // The email address to notify if the script encounters a critical error.
  ADMIN_EMAIL: 'admin@email.com',

  /**
   * Defines the structure of the email subject line.
   * Placeholders in the format {Question Title} will be replaced with the corresponding answer from the form.
   * The text inside the curly braces must be an EXACT match to the question title on your Google Form.
   */
  EMAIL_SUBJECT_TEMPLATE: 'Purchase Order Submission: {Funding Source} - {Vendor} - {Last Name}',

  /**
   * Defines the standard naming convention for uploaded files.
   * Use placeholders for form answers. The special placeholder {QuestionTitle} will be replaced
   * with the title of the file upload question itself (e.g., "Upload Quote").
   */
  FILE_NAME_TEMPLATE: '{Funding Source} - {Vendor} - {Last Name} - {QuestionTitle}',

  /**
   * Defines the prefix for files uploaded to the "special" question defined below.
   * This allows a different naming convention for specific uploads.
   */
  FILE_NAME_PREFIX_TEMPLATE: '{Funding Source} - {Vendor} - {Last Name}',

  /**
   * The exact title of the file upload question that should use a different, special naming rule.
   * Files uploaded to this question will be named using the prefix above, followed by their original filename.
   */
  SPECIAL_NAMING_QUESTION_TITLE: 'Supporting Docs (SDS, Justification, etc)',

  // A list of question titles to exclude from being used in the generated folder name.
  EXCLUDE_FROM_FOLDER_NAME: [
    "First Name", 
    "Do you know the funding code or name of the fund?",
    "Is this purchase supporting a Capstone Project, Honors Research Project, or Student Independent Research Project?",
    "If you answered anything but no, please give the sponsoring dept, project, and Student name",
    "Description of the items being ordered and a justification for the order (e.g. resistors and sensors for use in EW309 coursework)", 
    "Does this order require an ITPRA/ITPR or RFR (Radio Frequency Review)?", 
    "Does this order contain any hazardous materials (HAZMAT)? Examples include glue, paint, solvents, etc.  (If so, provide SDS sheet)",
    "Total purchase price"
  ],

  // A list of question titles to exclude from the body of the notification email.
  EXCLUDE_FROM_EMAIL_BODY: [
    "One or Two keywords word describing the items being ordered (this is used for folder naming and text in email subject lines)",
    "Do you know the funding code or name of the fund?",
    "Is this purchase supporting a Capstone Project, Honors Research Project, or Student Independent Research Project?",
    "If you answered anything but no, please give the sponsoring dept, project, and Student name"
  ]
};


/**
 * Initializes the script by creating an "onFormSubmit" trigger.
 * This function should be run once manually from the script editor to set up the automation.
 * It intelligently cleans up any old triggers for this function to prevent duplicate executions.
 */
function initialize() {
  const form = FormApp.getActiveForm();
  
  // Loop through all existing triggers for this script project.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    // If a trigger is found that already calls the 'onFormSubmit' function, delete it.
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create a new trigger that will automatically run the 'onFormSubmit' function
  // whenever a new response is submitted to the form.
  ScriptApp.newTrigger('onFormSubmit')
.forForm(form)
.onFormSubmit()
.create();
  Logger.log('Trigger created successfully.');
}


/**
 * The main function that executes automatically upon every form submission.
 * It orchestrates the entire workflow from data collection to file processing and notification.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e The event object passed by the trigger, containing all response data.
 */
function onFormSubmit(e) {
  // Use a try...catch block to gracefully handle any unexpected errors during execution.
  try {
    // Generate a consistent timestamp for this submission.
    const timestamp = Utilities.formatDate(new Date(), "America/New_York", "yyyyMMdd-HHmm");
    const itemResponses = e.response.getItemResponses();

    // Create a "map" of question titles to their answers. This is done once for efficiency
    // and makes it easy to retrieve any answer by its question title.
    const responseMap = itemResponses.reduce((map, response) => {
      const title = response.getItem().getTitle();
      const answer = response.getResponse()?? ''; // Use empty string for unanswered questions.
      map[title] = Array.isArray(answer)? answer.join(", ") : answer.toString();
      return map;
    }, {});

    // Build the names for the folder and email subject using the response data.
    const folderName = buildFolderName(itemResponses, timestamp);
    const emailSubject = buildFromTemplate(CONFIG.EMAIL_SUBJECT_TEMPLATE, responseMap);

    // Create the new folder in the specified Shared Drive.
    const targetDrive = DriveApp.getFolderById(CONFIG.TARGET_SHARED_DRIVE_ID);
    const newFolder = targetDrive.createFolder(folderName);
    Logger.log(`Successfully created folder: "${newFolder.getName()}" (ID: ${newFolder.getId()})`);

    // Process all uploaded files: rename them according to the rules and move them to the new folder.
    processAndRenameFiles(itemResponses, newFolder, responseMap);
    
    // Send the final notification email.
    sendNotificationEmail(e.response, emailSubject);
    
  } catch (error) {
    // If any error occurs in the 'try' block, this code will run.
    const errorMessage = `Error in onFormSubmit: ${error.message}\nStack:\n${error.stack}`;
    Logger.log(errorMessage); // Log the detailed error for debugging.
    
    // Send an immediate email alert to the administrator about the failure.
    MailApp.sendEmail(
      CONFIG.ADMIN_EMAIL, 
      'CRITICAL: Google Script Failure in Purchase Order Form', 
      `The Purchase order form automation script has failed.\n\n${errorMessage}`
    );
  }
}


/**
 * Constructs the name for the new Google Drive folder.
 * @param {FormApp.ItemResponse} itemResponses An array of the form's item responses.
 * @param {string} timestamp The formatted timestamp string for this submission.
 * @returns {string} The complete folder name (e.g., "20250710-0930_PENDING_...").
 */
function buildFolderName(itemResponses, timestamp) {
  const parts = itemResponses
    // Filter out file upload questions and any questions listed in the exclusion list.
   .filter(response => 
      response.getItem().getType()!== FormApp.ItemType.FILE_UPLOAD &&
     !CONFIG.EXCLUDE_FROM_FOLDER_NAME.includes(response.getItem().getTitle())
    )
    // Get the answer for each remaining question.
   .map(response => {
      const answer = response.getResponse()?? '';
      return Array.isArray(answer)? answer.join(", ") : answer.toString().trim();
    })
    // Remove any parts that are empty (from unanswered questions).
   .filter(part => part!== '');

  // Join the parts together with underscores to create the final name.
  return `${timestamp}_PENDING_${parts.join("_")}`;
}


/**
 * A generic and reusable function to build a string from a template.
 * It replaces {placeholders} in a template string with values from a data map.
 * @param {string} template The template string (e.g., "Order for {Vendor}").
 * @param {Object} responseMap A map of question titles to answers.
 * @param {Object} [additionalPlaceholders={}] Optional, for special placeholders not in the form (like {QuestionTitle}).
 * @returns {string} The final, formatted string.
 */
function buildFromTemplate(template, responseMap, additionalPlaceholders = {}) {
  // Combine the form responses with any additional special placeholders.
  const allPlaceholders = {...responseMap,...additionalPlaceholders };
  // Use a regular expression to find all instances of {placeholder} and replace them.
  return template.replace(/\{([^}]+)\}/g, (match, placeholder) => {
    // If a placeholder exists in our map, use its value; otherwise, use 'N/A'.
    return allPlaceholders[placeholder]?? 'N/A';
  });
}


/**
 * Renames and moves all uploaded files, applying conditional logic for special cases.
 * @param {FormApp.ItemResponse} itemResponses An array of the form's item responses.
 * @param {DriveApp.Folder} destinationFolder The newly created folder to move files into.
 * @param {Object} responseMap A map of question titles to answers, used for template filling.
 */
function processAndRenameFiles(itemResponses, destinationFolder, responseMap) {
  itemResponses
    // First, get only the file upload responses.
   .filter(itemResponse => itemResponse.getItem().getType() === FormApp.ItemType.FILE_UPLOAD)
    // Process each file upload question one by one.
   .forEach(itemResponse => {
      const fileIds = itemResponse.getResponse(); // This can be one or more file IDs.
      const questionTitle = itemResponse.getItem().getTitle();

      if (fileIds && fileIds.length > 0) {
        // Process each individual file uploaded to this question.
        fileIds.forEach((fileId, index) => {
          try {
            const file = DriveApp.getFileById(fileId);
            const originalName = file.getName();
            const extension = originalName.slice(originalName.lastIndexOf('.'));
            let newFileName;

            // Check if the current file upload question is the one designated for special naming.
            if (questionTitle === CONFIG.SPECIAL_NAMING_QUESTION_TITLE) {
              // --- SPECIAL NAMING LOGIC ---
              // Build the prefix from the specified template (e.g., "Gift - Vendor - Lastname").
              const prefix = buildFromTemplate(CONFIG.FILE_NAME_PREFIX_TEMPLATE, responseMap);
              const sanitizedPrefix = prefix.replace(/[\\/:"*?<>|]/g, '-'); // Ensure prefix is filename-safe.
              
              // Clean the " - Submitter Name" suffix that Google Forms automatically adds.
              const separatorIndex = originalName.lastIndexOf(' - ');
              const originalBaseName = (separatorIndex!== -1)
               ? originalName.substring(0, separatorIndex)
                : originalName.substring(0, originalName.lastIndexOf('.'));

              // Combine the prefix and the cleaned original name.
              newFileName = `${sanitizedPrefix} - ${originalBaseName}${extension}`;

            } else {
              // --- STANDARD NAMING LOGIC ---
              // Build the filename from the main template, passing in the current question's title.
              const baseFileName = buildFromTemplate(CONFIG.FILE_NAME_TEMPLATE, responseMap, { QuestionTitle: questionTitle });
              const sanitizedBaseName = baseFileName.replace(/[\\/:"*?<>|]/g, '-'); // Ensure filename is safe.
              newFileName = sanitizedBaseName;
              // If multiple files were uploaded to this single question, add a number to differentiate them.
              if (fileIds.length > 1) {
                newFileName += ` (${index + 1})`;
              }
              newFileName += extension;
            }

            // Apply the new name and move the file to its final destination.
            file.setName(newFileName);
            file.moveTo(destinationFolder);
            
            Logger.log(`Renamed "${originalName}" to "${newFileName}" and moved to "${destinationFolder.getName()}"`);

          } catch (error) {
            // If a single file fails, log the error and continue with the next one.
            Logger.log(`Failed to process file with ID "${fileId}". Reason: ${error.message}`);
          }
        });
      }
    });
}


/**
 * Constructs and sends the final notification email.
 * @param {FormApp.FormResponse} formResponse The complete form response object.
 * @param {string} emailSubject The dynamically generated subject for the email.
 */
function sendNotificationEmail(formResponse, emailSubject) {
  // Start with an empty email body.
  let emailBody = '';
  
  // Iterate through all form responses to build the email content.
  formResponse.getItemResponses().forEach(itemResponse => {
    const item = itemResponse.getItem();
    // Include the response unless it's a file upload or is on the exclusion list.
    if (item.getType()!== FormApp.ItemType.FILE_UPLOAD &&!CONFIG.EXCLUDE_FROM_EMAIL_BODY.includes(item.getTitle())) {
      const questionTitle = item.getTitle();
      const answer = itemResponse.getResponse()?? '<em>No response provided</em>';
      
      // Sanitize the text to prevent HTML issues in the email.
      const safeTitle = escapeHtml(questionTitle);
      const formattedAnswer = Array.isArray(answer)? escapeHtml(answer.join(", ")) : escapeHtml(String(answer));
      
      // Add the question and answer to the email body.
      emailBody += `<strong>${safeTitle}:</strong> ${formattedAnswer}<br/><br/>`;
    }
  });

  // Send the composed email.
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: emailSubject,
    htmlBody: emailBody
  });
  Logger.log(`Notification email sent to ${CONFIG.EMAIL_RECIPIENTS}.`);
}


/**
 * A utility function to escape HTML special characters in a string.
 * This prevents form answers from breaking the HTML structure of the email body.
 * @param {string} text The text to escape.
 * @returns {string} The HTML-safe text.
 */
function escapeHtml(text) {
  if (typeof text!== 'string') return text;
  return text
   .replace(/&/g, "&amp;")
   .replace(/</g, "&lt;")
   .replace(/>/g, "&gt;")
   .replace(/"/g, "&quot;")
   .replace(/'/g, "&#039;");
}
