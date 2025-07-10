# Google-Forms-Submission-Script (the sorting part is further down this readme)

This Google Apps Script provides a robust, end-to-end automation solution for processing purchase order (PO) requests submitted through a Google Form. It is designed to streamline administrative workflows by systematically organizing submission data and associated files within a Google Workspace environment. The script automatically creates a dedicated folder for each PO submission in a specified Shared Drive, intelligently renames and moves all uploaded documents (such as quotes and supporting materials), and dispatches a comprehensive notification email to key stakeholders. This eliminates manual file handling, reduces the potential for human error, and establishes a consistent, searchable digital archive for all procurement activities.

## Key Features

  - **Automated Folder Creation:** Upon each form submission, the script generates a uniquely named folder within a designated Google Shared Drive, ensuring that all materials related to a single purchase order are centralized and isolated.
  - **Dynamic Naming Conventions:** Folder and file names are constructed dynamically using answers provided in the form. This creates a standardized and predictable organizational structure that is easy to search and navigate.
  - **Intelligent File Handling:** The script identifies all files uploaded with a form submission, renames them according to a configurable template, and moves them from the default Google Forms upload location into the newly created purchase order folder.
  - **Conditional Naming Logic:** A sophisticated conditional logic system allows for different naming rules to be applied to different types of documents. Standard files like quotes are renamed for uniformity, while supporting documents can retain their original, descriptive filenames alongside a dynamic prefix.
  - **Automated Email Notifications:** A detailed HTML email summarizing the submission is automatically sent to a configurable list of recipients (e.g., procurement team, managers), ensuring timely awareness and action.
  - **Robust Error Handling:** The entire process is wrapped in a comprehensive error-handling block. In the event of a critical failure, a detailed alert is immediately sent to a designated administrator, ensuring system issues are promptly addressed.

## Prerequisites

Before deploying this script, ensure the following requirements are met [1, 2]:

1.  A **Google Workspace account** with administrative or sufficient permissions to:
      * Create and edit Google Forms.
      * Access and manage content within a Google Shared Drive.
      * Create and manage Google Apps Script projects and their associated triggers.
2.  A pre-existing **Google Form** designed for collecting purchase order information. This form must include one or more "File upload" questions for attaching documents like quotes or justifications.
3.  A target **Google Shared Drive** where the purchase order folders and files will be permanently stored.
4.  The **email addresses** of all personnel who need to receive submission notifications and the email address of the administrator who will receive error alerts.

## Setup and Installation Guide

Follow these steps carefully to link the script to your Google Form and configure the automation. This process involves four main stages: linking the script, configuring its settings, and performing a one-time initialization.

### Step 1: Link the Script to Your Google Form

The script must be "container-bound" to your specific Google Form to function correctly. This direct link allows it to use `FormApp.getActiveForm()` and respond to the form's submission events.[3, 4]

1.  Open your purchase order Google Form.
2.  In the top-right corner, click the **More** icon ( **⋮** ).
3.  From the dropdown menu, select **Script editor**.
4.  A new browser tab will open with the Google Apps Script editor. If you see a default `myFunction` block, you are in the right place.

### Step 2: Paste and Save the Script Code

1.  In the Script editor, delete any boilerplate code (like `function myFunction() {... }`) that may be present in the `Code.gs` file.
2.  Copy the entire contents of the provided `.js` file for this project.
3.  Paste the code into the empty `Code.gs` editor window.
4.  At the top of the editor, click the **Save project** icon (a floppy disk).

### Step 3: Configure the Script

All user-specific settings are consolidated into a single `CONFIG` object at the top of the script. This design separates configuration from the core application logic, making the script easier to manage and update. You must populate these settings with values specific to your workflow.

For a detailed explanation of each parameter, refer to the **Configuration Deep Dive** section below.

### Step 4: Perform First-Time Initialization

This is a critical, one-time manual step to activate the automation. The `initialize()` function creates the `onFormSubmit` trigger that will automatically run the script every time a new response is submitted to your form.[5, 6]

1.  In the Apps Script editor, locate the function selection dropdown menu in the toolbar at the top (it may default to `onFormSubmit`).
2.  Click the dropdown and select the `initialize` function.
3.  Click the **Run** button directly to the left of the dropdown.
4.  **Authorization Prompt:** The first time you run any function, Google will prompt you to grant the script the necessary permissions.
      * A dialog box titled "Authorization required" will appear. Click **Review permissions**.
      * Choose the Google account you want to authorize the script with.
      * You will see a screen detailing the permissions the script needs (e.g., "See, edit, create, and delete your Google Drive files," "Send email as you"). This is expected.
      * Click **Allow** to grant the permissions and proceed.

The `initialize` function is designed to be safely re-run at any time. It includes logic to first find and delete any old triggers associated with the `onFormSubmit` function before creating a new one. This idempotent design is a crucial safeguard that prevents the creation of multiple triggers, which would otherwise cause the script to execute multiple times for a single form submission, leading to duplicate folders and notifications. Once the function completes, the script is live and will process the next form submission automatically.

## Configuration Deep Dive

Properly configuring the `CONFIG` object is essential for the script to function as intended. This section provides a detailed breakdown of each parameter.

A foundational principle of this script is the "contract" between your Google Form and the script's configuration. The script uses placeholders in the format `{Question Title}` to dynamically insert data into folder names, file names, and email subjects. For this to work, the text inside the curly braces `{}` in any template string **must be an exact, character-for-character match** to the corresponding question title on your Google Form. Any mismatch will result in the placeholder being replaced with `N/A`.

### The CONFIG Object Parameters

The following table details each configuration variable, its purpose, and an example value.

| Parameter | Type | Description | Example |
| :--- | :--- | :--- | :--- |
| `TARGET_SHARED_DRIVE_ID` | `String` | The unique identifier for the Google Shared Drive where all purchase order folders will be created. **This is not the name of the drive.** See section 5.2 for instructions. | `'0A1B2c3d4E5f6Uk9PVA'` |
| `EMAIL_RECIPIENTS` | `String` | A comma-separated list of email addresses that will receive the notification email upon successful submission. | `'procurement@example.com,manager@example.com'` |
| `ADMIN_EMAIL` | `String` | The single email address to be notified if the script encounters a critical, unrecoverable error. This ensures system failures are immediately flagged for attention. | `'it-support@example.com'` |
| `EMAIL_SUBJECT_TEMPLATE` | `String` | Defines the subject line for the notification email. Use `{Question Title}` placeholders to dynamically insert answers from the form. The text inside the braces must be an **exact match** to your form's question titles. | `'PO Submission: {Vendor} - {Last Name}'` |
| `FILE_NAME_TEMPLATE` | `String` | The standard naming convention for most uploaded files. In addition to form question placeholders, you can use the special placeholder `{QuestionTitle}` which will be replaced by the title of the file upload question itself (e.g., "Upload Quote"). | `'{Funding Source} - {Vendor} - {QuestionTitle}'` |
| `FILE_NAME_PREFIX_TEMPLATE` | `String` | The prefix template used **only** for files uploaded to the question defined in `SPECIAL_NAMING_QUESTION_TITLE`. This allows for a different naming scheme for certain documents. | `'{Funding Source} - {Vendor} - {Last Name}'` |
| `SPECIAL_NAMING_QUESTION_TITLE` | `String` | The **exact title** of the file upload question on your form that should use the special naming rule. Files uploaded here will be named `[Prefix] - [Original Filename]`. | `'Supporting Docs (SDS, Justification, etc)'` |
| `EXCLUDE_FROM_FOLDER_NAME` | `Array<String>` | A comma-separated list of question titles inside single quotes to **omit** from the generated folder name. Useful for excluding metadata or non-descriptive fields. | `'Your Email Address', 'Timestamp'` |
| `EXCLUDE_FROM_EMAIL_BODY` | `Array<String>` | A comma-separated list of question titles inside single quotes to **omit** from the body of the notification email. Useful for excluding confirmation checkboxes or instructions. | `'I confirm the details are correct'` |

### How to Find Your Shared Drive ID

Google Drive uses unique IDs, not names, to identify folders and drives programmatically. Follow these steps to find the correct ID for your target Shared Drive [7, 8]:

1.  In a web browser, navigate to Google Drive and open the Shared Drive you intend to use.
2.  Click on the Shared Drive in the left-hand navigation pane.
3.  Look at the URL in your browser's address bar. It will have the following structure:
    `https://drive.google.com/drive/folders/{shared-drive-id}`
4.  The Shared Drive ID is the long string of alphanumeric characters that appears after `/folders/`.
5.  Copy this entire string (e.g., `0ALkFKuSgYKhQUk9PVA`) and paste it into the `TARGET_SHARED_DRIVE_ID` field in the `CONFIG` object, enclosed in single quotes.

## Automated Workflow Explained

Understanding the sequence of events that occurs after a user clicks "Submit" provides context for the script's architecture. The entire process is orchestrated by the `onFormSubmit` function.

1.  **Trigger:** A user completes and submits the Google Form. This action fires the `onFormSubmit` trigger that was created during initialization.
2.  **Data Collection:** The script receives the submission data as an event object. It immediately processes this data into a `responseMap`, which is an efficient key-value store (or dictionary) mapping each question title to its corresponding answer. This map is used repeatedly throughout the script, avoiding redundant data lookups.
3.  **Folder Naming:** The `buildFolderName` function is called. It constructs the name for the new folder by combining a precise timestamp, the status "PENDING", and a concatenated string of relevant answers from the form.
4.  **Folder Creation:** Using the generated name, the script creates a new folder within the target Shared Drive specified by `TARGET_SHARED_DRIVE_ID`.
5.  **File Processing:** The `processAndRenameFiles` function is executed. It systematically finds all file upload questions in the submission, iterates through each uploaded file, and applies the appropriate naming logic. It determines whether to use the "standard" or "special" naming rule based on the question's title, renames the file accordingly, and moves it from its temporary location to the newly created PO folder.
6.  **Notification:** The `sendNotificationEmail` function is called. It constructs a formatted HTML email containing the submission details (excluding any fields specified in the `EXCLUDE_FROM_EMAIL_BODY` list) and sends it to all recipients listed in `EMAIL_RECIPIENTS`.
7.  **Completion:** The script's execution concludes. If any unrecoverable error occurred during steps 2-6, the `catch` block would have immediately halted the process and sent a detailed error report to the `ADMIN_EMAIL`.

## Technical Reference: Function-by-Function Analysis

This section provides a detailed analysis of each function within the script, intended for developers or administrators who may need to modify or debug the code.

### `initialize()`

  * **Purpose:** To programmatically set up the automation by creating the necessary `onFormSubmit` trigger. This function is designed to be run manually one time from the script editor.
  * **Logic:** The function demonstrates a key principle of robust script design: idempotency. Before creating a new trigger, it first gets a list of all existing triggers in the project. It then iterates through this list and deletes any trigger that is already configured to call the `onFormSubmit` handler. This ensures that running the function multiple times will not create duplicate triggers, which would cause the entire workflow to run multiple times per submission. After the cleanup, it creates a single, new trigger bound to the form's submit event.

### `onFormSubmit(e)`

  * **Purpose:** This is the main function and the entry point for the entire automated workflow. It is executed automatically by the trigger each time a form response is submitted.
  * **Parameter:** `e` (`GoogleAppsScript.Events.FormsOnFormSubmit`) - The event object passed by the trigger, which contains all information about the submission, including the responses.[3]
  * **Logic:** The function acts as an orchestrator, coordinating calls to the various helper functions in the correct sequence. Its entire logic is enclosed within a `try...catch` block, which is a critical feature for production-level stability. It first generates a timestamp and creates the efficient `responseMap`. It then calls other functions to build names, create the folder, process files, and send the notification. If any of these steps fail, execution jumps to the `catch` block, which logs the detailed error (including the stack trace) and sends an immediate alert email to the administrator via `MailApp`.

### `buildFolderName(itemResponses, timestamp)`

  * **Purpose:** To construct a standardized and descriptive name for the Google Drive folder that will house all artifacts for a single PO.
  * **Logic:** The function receives the array of item responses and filters them based on two criteria: it excludes any response that is a file upload (`FormApp.ItemType.FILE_UPLOAD`) and any response whose question title is listed in the `CONFIG.EXCLUDE_FROM_FOLDER_NAME` array. The remaining answers are then converted to strings, trimmed of whitespace, and filtered again to remove any empty parts (from unanswered optional questions). Finally, these parts are joined together with underscores and prefixed with the `timestamp` and the static string `_PENDING_`.

### `buildFromTemplate(template, responseMap, additionalPlaceholders)`

  * **Purpose:** A generic and highly reusable utility function for performing mail-merge-style string replacements. This function embodies the DRY (Don't Repeat Yourself) principle.
  * **Logic:** It takes a template string (e.g., `PO for {Vendor}`), a map of form responses, and an optional object for special placeholders. It uses a regular expression (`/\{([^}]+)\}/g`) to find all occurrences of `{placeholder}`. For each match, it looks up the placeholder's name (the part inside the braces) in a combined data object. If a corresponding value exists, it is substituted; otherwise, it defaults to the string `'N/A'`, which makes troubleshooting misconfigured placeholders straightforward.

### `processAndRenameFiles(itemResponses, destinationFolder, responseMap)`

  * **Purpose:** To manage the core logic of renaming and moving all files uploaded with the form submission.
  * **Logic:** This function contains the most complex conditional logic in the script. It begins by filtering the item responses to isolate only the file upload questions. It then iterates through each of these questions. For each file, it checks if the title of the question it belongs to is the same as the one defined in `CONFIG.SPECIAL_NAMING_QUESTION_TITLE`.
      * **If it is the special question:** It applies the special naming rule. It builds a prefix using `FILE_NAME_PREFIX_TEMPLATE` and combines it with the file's original name. The logic is sophisticated enough to first remove the `  - Submitter Name ` suffix that Google Forms automatically appends to uploaded files, ensuring a clean final name. This rule is designed for documents where the original filename provides important context (e.g., `Safety Data Sheet for Acetone.pdf`).
      * **If it is any other file upload question:** It applies the standard naming rule. It builds a new base name using `FILE_NAME_TEMPLATE`, which often includes the `{QuestionTitle}` placeholder to differentiate between, for example, a quote and an invoice. If multiple files were uploaded to a single standard question, it appends a sequential number (e.g., `(1)`, `(2)`) to prevent name collisions.
      * In both cases, the function sanitizes the generated names to remove characters that are illegal in filenames and then performs the final `setName()` and `moveTo()` operations.

### `sendNotificationEmail(formResponse, emailSubject)`

  * **Purpose:** To construct and dispatch the final HTML notification email to stakeholders.
  * **Logic:** The function iterates through all the form's item responses. For each response, it checks that it is not a file upload and that its title is not in the `CONFIG.EXCLUDE_FROM_EMAIL_BODY` list. For all other responses, it appends the question and its corresponding answer to an HTML string, using `<strong>` tags for the question title to improve readability. Crucially, before adding any user-supplied text to the email body, it passes the text through the `escapeHtml` utility function.

### `escapeHtml(text)`

  * **Purpose:** A small but vital security and utility function to prevent malformed data from breaking the notification email.
  * **Logic:** It takes a string and replaces key HTML characters (`&`, `<`, `>`, `"`, `'`) with their corresponding HTML entities. This ensures that if a user enters text like `<script>alert('test')</script>` into a form field, it will be rendered as plain text in the email rather than being executed as HTML, which could corrupt the email's layout or pose a security risk.

## Error Handling and Debugging

The script is designed with monitoring and fault tolerance in mind.

  - **Runtime Error Capture:** The primary `onFormSubmit` function is wrapped in a `try...catch` block. This means that any unexpected error during the execution—such as an invalid Drive ID, permissions issues, or an API service outage—will be caught gracefully instead of causing the script to crash silently.
  - **Administrator Alerts:** When an error is caught, the script immediately sends a detailed alert email to the address specified in `ADMIN_EMAIL`. This email contains the error message and a full stack trace, providing the administrator with the necessary information to diagnose the problem quickly.
  - **Execution Logs:** Google Apps Script provides a detailed execution log. To view it, open the script editor and click on the **Executions** tab in the left-hand navigation pane. Each time the script runs (either successfully or with an error), a new entry is created. Clicking on an entry reveals all the `Logger.log()` outputs from that run, which provides a step-by-step trail of the script's operations. This is the primary tool for debugging non-critical issues.

## Troubleshooting FAQ

This section addresses common issues that may arise during setup or operation.[2, 9]

  - **Q: The script ran, but my email subject or filenames have "N/A" in them. Why?**

      - **A:** This is the most frequent issue and is almost always caused by a mismatch between a placeholder in your `CONFIG` templates (e.g., `{Vendor}`) and the corresponding question title on your Google Form. They must be an **exact match**, including capitalization, punctuation, and spacing. Carefully compare every placeholder in your `CONFIG` strings with your form's question titles.

  - **Q: I submitted the form, but nothing happened. What should I do?**

      - **A:** Follow these diagnostic steps:
        1.  Check the **Executions** log in the Apps Script editor. Look for the most recent run and see if it failed or completed successfully. If it failed, the error message will be displayed there.
        2.  Verify that you ran the `initialize()` function **one time** to create the trigger. You can confirm this by going to the **Triggers** tab (clock icon) in the script editor. You should see one trigger for the `onFormSubmit` function.
        3.  Check the inbox for the `ADMIN_EMAIL` to see if a critical failure alert was sent.

  - **Q: I'm getting a "permission denied" or "authorization" error in the logs.**

      - **A:** This indicates an issue with the permissions you granted the script. This can happen if you initially denied a required permission or if you modified the script to use a new Google service. To fix this, simply re-run the `initialize()` function from the editor. This will re-initiate the authorization flow, allowing you to review and grant the correct permissions.

  - **Q: My files are not being moved to a Shared Drive, or I get an error about the folder ID.**

      - **A:** Double-check the `TARGET_SHARED_DRIVE_ID` in your configuration. Ensure it is the ID for a **Shared Drive** and not a regular folder in "My Drive," as the APIs can differ.[10] Also, verify that the Google account under which the script is authorized has at least **Content manager** access to that Shared Drive, as this level of permission is required to add and move files.


# Google-Drive-Sorting-Script

This Google Drive App Script is designed to automate the organization of subfolders within a specific Google Drive folder. It utilizes Google Apps Script, a JavaScript cloud scripting language that provides easy ways to automate tasks across Google products and third party services. The script sorts subfolders based on predefined criteria found in the subfolder names, then moves them into a structured folder hierarchy. Below is a detailed explanation of each part of the script:

### Entry Point: `organizeSubfolders()`

-   Purpose: This function serves as the entry point of the script. It specifies the folder to be organized and initiates the organization process.
-   How it works: It starts by defining a `folderId` which is the ID of the Google Drive folder you want to organize. Then, it retrieves this folder as a `parentFolder` object using `DriveApp.getFolderById(folderId)`. Finally, it calls `scanAndOrganizeSubfolders(parentFolder, parentFolder)` to begin organizing the folder's contents.

### Core Functionality: `scanAndOrganizeSubfolders(folder, baseParentFolder)`

-   Purpose: To recursively scan through all subfolders of the given folder, organizing them based on predefined criteria.
-   How it works: This function goes through each subfolder, checking if the subfolder name matches the predefined sorting criteria (e.g., specific keywords). If a match is found and the subfolder isn't already organized correctly, it moves the subfolder into the appropriate location within a structured hierarchy based on the sorting words and the date string extracted from the subfolder name.

### Helper Functions:

1.  `isSubfolderAlreadySorted(subfolder, baseParentFolder, expectedParentFolderName, month)`

    -   Purpose: Checks if the subfolder is already in the correct location to avoid unnecessary moves.
    -   How it works: It looks for the expected parent folder and month subfolder within the `baseParentFolder`. If the subfolder's current parent matches the expected location, it returns true; otherwise, false.
2.  `moveSubfolder(subfolder, newParentFolder)`

    -   Purpose: Moves the subfolder, including all its contents, to a new parent folder.
    -   How it works: It creates a new subfolder under the `newParentFolder` with the same name as the original subfolder, moves all files and sub-subfolders to this new subfolder, and finally trashes the original subfolder.
3.  `createSubfolders(folder, subfolderPath)`

    -   Purpose: Ensures the necessary folder structure exists within the `baseParentFolder` based on a given path.
    -   How it works: It iteratively checks for the existence of each subfolder in the path, creating any that don't exist, and navigates down the folder hierarchy to the final folder in the path.
4.  `findOrCreateFolder(parentFolder, folderName)`

    -   Purpose: Finds or creates a subfolder by name under a specified parent folder.
    -   How it works: It searches for a folder by name under the `parentFolder`. If the folder exists, it returns this folder; if not, it creates a new folder with the specified name.

### Automation Trigger: `createTrigger()`

-   Purpose: To set up an automatic trigger that runs the `organizeSubfolders` function periodically.
-   How it works: It creates a new time-based trigger using `ScriptApp.newTrigger("organizeSubfolders")` that executes the `organizeSubfolders` function every 15 minutes.

### Summary

The script efficiently organizes folders in Google Drive by sorting them into a predefined hierarchy based on their names, specifically looking for sorting keywords and date strings. It is recursive, ensuring even nested folders are organized. This can greatly enhance file management, especially for users dealing with a large number of folders that follow a consistent naming convention. The script also demonstrates the power of Google Apps Script for automating repetitive tasks within Google Drive, saving time and reducing the potential for human error.

created by [John Seargeant](https://github.com/John-Sarge)   03Feb2024
