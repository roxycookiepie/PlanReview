// Sanitized for public sharing: sensitive identifiers, environment-specific config keys,
// internal library/site references, database names, and hardcoded test file identifiers have been redacted.

﻿using AI.Models;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Ajax.Utilities;
using Microsoft.SharePoint.Client.DocumentManagement;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace AI
{
    public class PlanReview
    {

        public LogFile logFile;
        readonly private static string sqlDatasource = ConfigurationManager.AppSettings["DB_SOURCE"];
        readonly private static string sqlUserID = ConfigurationManager.AppSettings["DB_USER"];
        readonly private static string sqlPassword = ConfigurationManager.AppSettings["DB_PASSWORD"];
        readonly private static int landPlanSetPageCategoryAttributeId = int.Parse(ConfigurationManager.AppSettings["LAND_PLAN_SET_PAGE_CATEGORY_ATTRIBUTE_ID"] ?? "0");
        readonly private static int transportationPageCategoryAttributeId = int.Parse(ConfigurationManager.AppSettings["TRANSPORTATION_PAGE_CATEGORY_ATTRIBUTE_ID"] ?? "0");
        readonly private static int platTitleBlockAttributeId = int.Parse(ConfigurationManager.AppSettings["PLAT_TITLE_BLOCK_ATTRIBUTE_ID"] ?? "0");
        readonly private static int platBlockNumberAttributeId = int.Parse(ConfigurationManager.AppSettings["PLAT_BLOCK_NUMBER_ATTRIBUTE_ID"] ?? "0");
        readonly private static int transportationTitleBlockDetailSectionAttributeId = int.Parse(ConfigurationManager.AppSettings["TRANSPORTATION_TITLE_BLOCK_DETAIL_SECTION_ATTRIBUTE_ID"] ?? "0");
        private readonly SharePointAPI spApi;
        private readonly AgileDDApi agileDDApi;
        private readonly string siteID;
        private readonly string libraryName;
        private readonly bool quicktest = true; //*****
        private TitleBlock titleBlock; // Platting

        public PlanReview(LogFile lf)
        {
            logFile = lf;
            spApi = new SharePointAPI(logFile);
            agileDDApi = new AgileDDApi(logFile);
            siteID = "REDACTED_SITE_ID";
            libraryName = "REDACTED_LIBRARY";
        }

        public async Task<PlanReviewResponse> ProcessPlanReviewAsync(int runId)
        {
            logFile.MyLogFile($"Starting Plan Review process for runId {runId} at {DateTime.Now}.");

            // Get request data from SQL PlanReview_Run based on runId
            var planReviewRun = await GetPlanReviewRunAsync(runId);
            if (planReviewRun == null)
            {
                logFile.MyLogFile($"No Plan Review Run found for runId {runId}.");
                return new PlanReviewResponse
                {
                    Status = "Error",
                    Message = $"No Plan Review Run found for runId {runId}."
                };
            }
            logFile.MyLogFile($"Plan Review Run found: ID={planReviewRun.ID}, CreatedBy={planReviewRun.CreatedBy}, Timestamp={planReviewRun.Timestamp}, Status={planReviewRun.Status}, PlanReviewType={planReviewRun.PlanReviewType}.");

            // Get documents from SharePoint library based on runId *****
            var documents = await spApi.GetDocumentLibraryMetadataAsync(siteID, libraryName, runId.ToString());
            if (documents == null || !documents.Any())
            {
                logFile.MyLogFile($"No documents found for runId {runId} in library {libraryName}.");
                return new PlanReviewResponse
                {
                    Status = "Error",
                    Message = $"No documents found for runId {runId} in library {libraryName}."
                };
            }
            logFile.MyLogFile($"Found {documents.Count} documents for runId {runId} in library {libraryName}.");
            #region TESTING - Define hardcoded documents*****
            //var hardcodedDocuments = new List<(int DocumentId, string FileCategory, string FileName)>
            //{
            //    // Format: (DocumentId, FileCategory, FileName)
            //    (0, "Land Plan Set", "REDACTED_TEST_FILE.pdf")
            //    //(0, "Land Plan Set", "REDACTED_TEST_FILE.pdf")
            //};
            var hardcodedDocuments = new List<PlanReviewDocument>
            {
                //new PlanReviewDocument(0, "CPC101", "REDACTED_TEST_FILE.pdf"),
                //new PlanReviewDocument(0, "Plat", "REDACTED_TEST_FILE.pdf")
                //new PlanReviewDocument(0, "CPC101", "REDACTED_TEST_FILE.pdf"), //block issue example [REDACTED]
                new PlanReviewDocument(0, "Transportation", "REDACTED_TEST_FILE.pdf") //Transportation [REDACTED]
                //new PlanReviewDocument(0, "Transportation", "REDACTED_TEST_FILE.pdf") //Transportation [REDACTED]
                //new PlanReviewDocument(0, "Transportation", "REDACTED_TEST_FILE.pdf") //Transportation [REDACTED]
                //new PlanReviewDocument(0, "Transportation", "REDACTED_TEST_FILE.pdf") //Transportation [REDACTED]
                //new PlanReviewDocument(0, "Transportation", "REDACTED_TEST_FILE.pdf") //Transportation [REDACTED]
            };
            #endregion

            // Initialize task list and attributes for Plan Review
            var taskList = await GetPlanReviewTasksWithAttributesAsync(planReviewRun.PlanReviewType, planReviewRun.TaskSet);
            if (taskList == null | !taskList.Any())
            {
                logFile.MyLogFile($"No tasks found for Plan Review Type {planReviewRun.PlanReviewType} and Task Set {planReviewRun.TaskSet}.");
                return new PlanReviewResponse
                {
                    Status = "Error",
                    Message = $"No tasks found for Plan Review Type {planReviewRun.PlanReviewType} and Task Set {planReviewRun.TaskSet}."
                };
            }

            // Dictionary to store sheet numbers per page (for Transportation Sheet Index validation)
            Dictionary<int, string> sheetNumbersByPage = new Dictionary<int, string>();

            // Upload and Process each document in parallel
            List<Task> tasks = null;
            if (quicktest) //***** make sure to set back to false for production and both sets
            {
                tasks = hardcodedDocuments.Select(async docInfo =>
                {
                    int documentId = docInfo.DocumentId;
                    string fileCategory = docInfo.FileCategory;
                    string fileName = docInfo.FileName;

                    // Skip the upload and processing steps since we're using existing documents for testing

                    // Get attributes for the file
                    var attributes = await agileDDApi.ReadValuesAsync(documentId);
                    if (!attributes.Success)
                    {
                        logFile.MyLogFile($"Error retrieving attributes for document ID {documentId}: {attributes.ErrorMessage}");
                        return;
                    }
                    logFile.MyLogFile($"Attributes for document ID {documentId} retrieved successfully. Found {attributes.Content.Count} attributes.");

                    logFile.MyLogFile("Associating attributes with tasks.");
                    foreach (var task in taskList)
                    {
                        foreach (var attributeToUpdate in task.Attributes.Where(a => a.FileCategory == fileCategory))
                        {
                            // Look for a matching API attribute based on the attribute code
                            var matchedAttribute = attributes.Content.FirstOrDefault(a => a.attribute_id == attributeToUpdate.AttributeCode);
                            string newValue = "N/A";

                            if (matchedAttribute != null)
                            {
                                if (fileCategory == "Land Plan Set") //*****special case for land plan set to get page category
                                {
                                    // Find the PageCategory attribute for the same page as the matched attribute
                                    var pageCategoryAttr = attributes.Content.FirstOrDefault(a =>
                                        a.attribute_id == landPlanSetPageCategoryAttributeId &&
                                        a.page_number == matchedAttribute.page_number);
                                    attributeToUpdate.PageCategory = pageCategoryAttr?.data;
                                }

                                if (fileCategory == "Transportation") //*****special case for land plan set to get page category
                                {
                                    // Find the PageCategory attribute for the same page as the matched attribute
                                    var pageCategoryAttr = attributes.Content.FirstOrDefault(a =>
                                        a.attribute_id == transportationPageCategoryAttributeId &&
                                        a.page_number == matchedAttribute.page_number);
                                    attributeToUpdate.PageCategory = pageCategoryAttr?.data;
                                }

                                // If matching attribute found and it has classifications, perform extra API call
                                if (matchedAttribute.classifications != null &&
                                    matchedAttribute.classifications.Count > 0 &&
                                    matchedAttribute.position != null &&
                                    matchedAttribute.position.Any())
                                {
                                    newValue = await agileDDApi.GetTextFromBoxAsync(documentId, matchedAttribute.page_number, matchedAttribute.position);
                                    logFile.MyLogFile($"Extra API call for '{attributeToUpdate.AttributeName}' returned value: {newValue}");
                                }
                                else
                                {
                                    newValue = matchedAttribute.data ?? "N/A";
                                }
                            }
                            attributeToUpdate.Value = newValue;
                            logFile.MyLogFile($"Updated attribute '{attributeToUpdate.AttributeName}' with value: {newValue}");
                        }

                        // Track which documents contributed to this task
                        if (!task.DocumentIds.Contains(documentId))
                        {
                            task.DocumentIds.Add(documentId);
                        }

                        // Track document ID by category
                        if (!task.DocumentIdsByCategory.ContainsKey(fileCategory))
                        {
                            task.DocumentIdsByCategory[fileCategory] = documentId;
                        }
                    }
                }).ToList();
            }
            else
            {
                tasks = documents.Select(async doc =>
                {
                    if (doc.DriveItem == null || doc.Metadata == null)
                    {
                        logFile.MyLogFile($"Document {doc.DriveItem?.Name} is missing DriveItem or Metadata.");
                        return;
                    }

                    // Get the file stream from SharePoint
                    string fileName = doc.DriveItem.Name;
                    var stream = await spApi.GetFileStreamAsync(siteID, libraryName, fileName);
                    if (stream == null)
                    {
                        logFile.MyLogFile($"Failed to get stream for file {fileName}.");
                        return;
                    }

                    // Upload file to batch in AgileDD
                    string fileCategory = doc.Metadata.ContainsKey("FileCategory") ? doc.Metadata["FileCategory"].ToString() : "DefaultCategory";
                    var uploadResult = await agileDDApi.UploadFileToBatchAsync(fileCategory, stream, fileName);
                    stream.Dispose();
                    if (!uploadResult.Success)
                    {
                        logFile.MyLogFile($"Error uploading file {fileName}: {uploadResult.ErrorMessage}");
                        return;
                    }
                    var uploadResponse = uploadResult.Content;
                    string uploadJobId = uploadResponse.job_id;
                    logFile.MyLogFile($"File {fileName} started uploading. Upload Job ID: {uploadJobId}.");

                    // Poll until the upload job status is "finished"
                    while (true)
                    {
                        var status = await agileDDApi.CheckJobStatusAsync(uploadJobId);
                        if (status != null && status == "finished")
                        {
                            logFile.MyLogFile($"File {fileName} upload processing finished.");
                            break;
                        }
                        logFile.MyLogFile($"File {fileName} upload job status: {status}. Waiting...");
                        await Task.Delay(10000); // wait 10 seconds before polling again
                    }

                    // Use the first returned document's ID from the upload response
                    int documentId = uploadResponse.documents.First().id;
                    var processResult = await agileDDApi.ProcessDocumentAsync(documentId);
                    if (!processResult.Success)
                    {
                        logFile.MyLogFile($"Error starting process for file {fileName}: {processResult.ErrorMessage}");
                        return;
                    }
                    string processJobId = processResult.Content;
                    logFile.MyLogFile($"File {fileName} process started. process Job ID: {processJobId}.");

                    // Poll until the process job status is "complete"
                    while (true)
                    {
                        var processStatus = await agileDDApi.CheckDocumentJobStatusAsync(processJobId);
                        if (processStatus != null && processStatus == "complete")
                        {
                            logFile.MyLogFile($"File {fileName} process processing finished.");
                            break;
                        }
                        logFile.MyLogFile($"File {fileName} process job status: {processStatus}. Waiting...");
                        await Task.Delay(60000); // wait 60 seconds before polling again
                    }

                    // Get attributes for the file
                    var attributes = await agileDDApi.ReadValuesAsync(documentId);
                    if (!attributes.Success)
                    {
                        logFile.MyLogFile($"Error retrieving attributes for document ID {documentId}: {attributes.ErrorMessage}");
                        return;
                    }
                    logFile.MyLogFile($"Attributes for document ID {documentId} retrieved successfully. Found {attributes.Content.Count} attributes.");

                    logFile.MyLogFile("Associating attributes with tasks.");
                    foreach (var task in taskList)
                    {
                        foreach (var attributeToUpdate in task.Attributes.Where(a => a.FileCategory == fileCategory))
                        {
                            // Look for a matching API attribute based on the attribute code
                            var matchedAttribute = attributes.Content.FirstOrDefault(a => a.attribute_id == attributeToUpdate.AttributeCode);
                            string newValue = "N/A";

                            if (matchedAttribute != null)
                            {
                                if (fileCategory == "Land Plan Set") //*****special case for land plan set to get page category
                                {
                                    // Find the PageCategory attribute for the same page as the matched attribute
                                    var pageCategoryAttr = attributes.Content.FirstOrDefault(a =>
                                        a.attribute_id == landPlanSetPageCategoryAttributeId &&
                                        a.page_number == matchedAttribute.page_number);
                                    attributeToUpdate.PageCategory = pageCategoryAttr?.data;
                                }

                                if (fileCategory == "Transportation") //*****special case for land plan set to get page category
                                {
                                    // Find the PageCategory attribute for the same page as the matched attribute
                                    var pageCategoryAttr = attributes.Content.FirstOrDefault(a =>
                                        a.attribute_id == transportationPageCategoryAttributeId &&
                                        a.page_number == matchedAttribute.page_number);
                                    attributeToUpdate.PageCategory = pageCategoryAttr?.data;
                                }

                                // If matching attribute found and it has classifications, perform extra API call
                                if (matchedAttribute.classifications != null &&
                                    matchedAttribute.classifications.Count > 0 &&
                                    matchedAttribute.position != null &&
                                    matchedAttribute.position.Any())
                                {
                                    newValue = await agileDDApi.GetTextFromBoxAsync(documentId, matchedAttribute.page_number, matchedAttribute.position);
                                    logFile.MyLogFile($"Extra API call for '{attributeToUpdate.AttributeName}' returned value: {newValue}");
                                }
                                else
                                {
                                    newValue = matchedAttribute.data ?? "N/A";
                                }
                            }
                            attributeToUpdate.Value = newValue;
                            logFile.MyLogFile($"Updated attribute '{attributeToUpdate.AttributeName}' with value: {newValue}");
                        }

                        // Track which documents contributed to this task
                        if (!task.DocumentIds.Contains(documentId))
                        {
                            task.DocumentIds.Add(documentId);
                        }

                        // Track document ID by category
                        if (!task.DocumentIdsByCategory.ContainsKey(fileCategory))
                        {
                            task.DocumentIdsByCategory[fileCategory] = documentId;
                        }
                    }
                }).ToList();
            }
            await Task.WhenAll(tasks);

            // If Platting type get the TitleBlock attribute from the document
            // If Platting type get the TitleBlock attribute from the document
            if (planReviewRun.PlanReviewType == "Platting")
            {
                // Get the Plat document ID using DocumentIdsByCategory for the "Plat Name" task since it definitely has a Plat document
                int documentId = 0;
                var titleBlockTask = taskList.FirstOrDefault(t => t.TaskName == "Plat Name");
                if (titleBlockTask != null && titleBlockTask.DocumentIdsByCategory.ContainsKey("Plat"))
                {
                    documentId = titleBlockTask.DocumentIdsByCategory["Plat"];
                    logFile.MyLogFile($"Using Plat document ID: {documentId} for TitleBlock extraction");
                }
                else
                {
                    logFile.MyLogFile($"Task 'Plat Name' not found or no Plat document category available");
                }

                if (documentId == 0)
                {
                    logFile.MyLogFile("No document ID available for TitleBlock extraction.");
                }
                else
                {
                    logFile.MyLogFile("Processing TitleBlock extraction.");

                    // Extract titleblock attributes from the Plat document
                    var attributes = await agileDDApi.ReadValuesAsync(documentId);

                    if (!attributes.Success)
                    {
                        logFile.MyLogFile($"Error retrieving attributes for document ID {documentId}: {attributes.ErrorMessage}");
                    }
                    else
                    {
                        logFile.MyLogFile($"Attributes retrieved successfully. Found {attributes.Content.Count} attributes.");

                        // Get the TitleBlock attribute ID from configuration
                        int titleBlockAttributeId = platTitleBlockAttributeId;

                        // Prefer page 1, then page 2 (index sheet), then earliest page
                        var titleBlockAttribute = attributes.Content
                            .Where(a => a.attribute_id == titleBlockAttributeId
                                     && a.position != null
                                     && a.position.Any())
                            .OrderBy(a => a.page_number)
                            .FirstOrDefault(a => a.page_number <= 5)  // cover pages only
                            ?? attributes.Content
                                .Where(a => a.attribute_id == titleBlockAttributeId
                                         && a.position != null
                                         && a.position.Any())
                                .OrderBy(a => a.page_number)
                                .FirstOrDefault();

                        if (titleBlockAttribute != null)
                        {
                            logFile.MyLogFile($"Using page {titleBlockAttribute.page_number} for TitleBlock extraction");
                            string titleBlockText = "N/A";

                            // If classifications exist and position data is available, extract text from box
                            if (titleBlockAttribute.classifications != null &&
                                titleBlockAttribute.classifications.Count > 0 &&
                                titleBlockAttribute.position != null &&
                                titleBlockAttribute.position.Any())
                            {
                                logFile.MyLogFile("TitleBlock has classifications and position data. Extracting text from box.");
                                titleBlockText = await agileDDApi.GetTextFromBoxAsync(documentId, titleBlockAttribute.page_number, titleBlockAttribute.position);
                                logFile.MyLogFile($"Extracted TitleBlock text from box: {titleBlockText}");

                                ////Test *****
                                //// Initialize with Phi-4 visual model for image processing
                                //var azurePhiApi = new AzureOpenAIApi(
                                //    logFile,
                                //    "Extract information from engineering drawings.",
                                //    Convert.ToDecimal(0.0),
                                //    usePhiVisual: true
                                //);

                                //// Pass classifications to get cropped image
                                //if (titleBlockAttribute.position != null && titleBlockAttribute.position.Any())
                                //{
                                //    var position = titleBlockAttribute.position[0];
                                //    var imageBytes = await agileDDApi.GetCroppedImageAsync(
                                //        documentId,
                                //        titleBlockAttribute.page_number,
                                //        position.left,
                                //        position.top,
                                //        position.width,
                                //        position.height);

                                //    if (imageBytes != null)
                                //    {
                                //        var prompt = "Extract all fields from this title block. Return ONLY a JSON object with keys like acreage, blocks, lots, reserves, scale, action_date, owner, location, abstract_number, city_etj, county.";
                                //        var result = await azurePhiApi.QueryImageAsync(imageBytes, prompt);

                                //        if (result.Success)
                                //        {
                                //            logFile.MyLogFile($"Successfully extracted title block image data: {result.Content}");
                                //            var titleBlockData = JsonConvert.DeserializeObject<dynamic>(result.Content);
                                //            // Process the extracted data as needed
                                //        }
                                //        else
                                //        {
                                //            logFile.MyLogFile($"Error querying title block image: {result.ErrorMessage}");
                                //        }
                                //    }
                                //    else
                                //    {
                                //        logFile.MyLogFile("Failed to retrieve cropped image from AgileDD");
                                //    }
                                //}
                                ////end test
                            }
                            else
                            {
                                logFile.MyLogFile("TitleBlock attribute found but has no usable data.");
                            }

                            if (!string.IsNullOrEmpty(titleBlockText) && titleBlockText != "N/A")
                            {
                                logFile.MyLogFile("Processing TitleBlock extraction.");
                                await ProcessTitleBlockExtractionAsync(titleBlockText);
                            }
                            else
                            {
                                logFile.MyLogFile("No TitleBlock text available for extraction.");
                            }
                        }
                        else
                        {
                            logFile.MyLogFile("No TitleBlock attribute found or value is empty.");
                        }
                    }
                }
            }

            // Process review logic for each task
            logFile.MyLogFile("Processing review logic.");
            foreach (var task in taskList)
            {
                var attr1 = task.Attributes[0];
                attr1.Value = NormalizeValue(attr1.Value);

                if (task.LogicType == "Comparison")
                {
                    //comparing 2 values for each task
                    string comparisonValue = null;
                    string comparisonValueName = null;

                    // If Platting type, compare CPC101 attribute to titleBlock property. Also, plat name has its own attribute to compare against instead of title block.
                    if (planReviewRun.PlanReviewType == "Platting" && titleBlock != null && task.TaskName != "Plat Name")
                    {
                        // Get the CPC101 attribute (should be attr1)
                        var cpc101Attr = task.Attributes.FirstOrDefault(a => a.FileCategory == "CPC101");
                        if (cpc101Attr == null)
                        {
                            logFile.MyLogFile($"No CPC101 attribute found for task '{task.TaskName}' in Platting review.");
                            task.Status = "Not enough data";
                            continue;
                        }

                        // Get corresponding value from titleBlock
                        comparisonValue = GetTitleBlockPropertyValue(task.TaskName, titleBlock);
                        comparisonValueName = "Plat (TitleBlock)";
                        attr1 = cpc101Attr;
                    }
                    else
                    {
                        // Original logic: compare two task attributes
                        if (task.Attributes.Count < 2)
                        {
                            logFile.MyLogFile($"Not enough attributes for task '{task.TaskName}' to compare.");
                            task.Status = "Not enough data";
                            continue;
                        }
                        comparisonValue = task.Attributes[1].Value;
                        comparisonValueName = task.Attributes[1].AttributeName;
                    }

                    // Normalize values before comparison and logging
                    attr1.Value = NormalizeValue(attr1.Value);
                    comparisonValue = NormalizeValue(comparisonValue);

                    if (string.IsNullOrEmpty(attr1.Value) || string.IsNullOrEmpty(comparisonValue))
                    {
                        logFile.MyLogFile($"'{attr1.AttributeName}' or '{comparisonValueName}' values are missing for one or both sources.");
                        task.Status = "Notice";
                        continue;
                    }

                    // Treat "0" vs "N/A" (in either order) as a match (mainly for Lots)
                    if ((attr1.Value == "0" && comparisonValue == "n/a") ||
                        (attr1.Value == "n/a" && comparisonValue == "0"))
                    {
                        logFile.MyLogFile($"'{attr1.AttributeName}' match: {attr1.Value} vs {comparisonValue} (0 and N/A treated as match)");
                        task.Status = "Match";
                        continue;
                    }

                    // handle city/etj separately
                    if (task.TaskName == "City/ETJ")
                    {
                        bool v1HasCityOrEtj = attr1.Value.Contains("city") || attr1.Value.Contains("etj");
                        bool v2HasCityOrEtj = comparisonValue.Contains("city") || comparisonValue.Contains("etj");

                        if (!(v1HasCityOrEtj && v2HasCityOrEtj))
                        {
                            logFile.MyLogFile($"'City/ETJ' mismatch: {attr1.Value} vs {comparisonValue} (both must contain 'city' or 'etj')");
                            task.Status = "Mismatch";
                            continue;
                        }
                        else
                        {
                            logFile.MyLogFile($"'{attr1.AttributeName}' match: {attr1.Value}");
                            task.Status = "Match";
                            continue;
                        }
                    }

                    if (!ValuesMatch(attr1.Value, comparisonValue))
                    {
                        logFile.MyLogFile($"'{attr1.AttributeName}' mismatch: {attr1.Value} vs {comparisonValue}");
                        task.Status = "Mismatch";
                        continue;
                    }
                    else
                    {
                        logFile.MyLogFile($"'{attr1.AttributeName}' match: {attr1.Value}");
                        task.Status = "Match";
                        continue;
                    }
                } //LP-HTX

                if (task.LogicType == "GeneralNotes" && task.TaskSet == "CPC101/Plat")
                {
                    logFile.MyLogFile($"Handling General Notes");

                    //call GetGeneralNotesTasksAsync to get the list of General Notes tasks
                    var generalNotesTasks = await GetGeneralNotesTasksAsync();
                    if (!generalNotesTasks.Any())
                    {
                        logFile.MyLogFile("No general notes tasks found.");
                        task.Status = "Error";
                        continue;
                    }

                    bool allNotesFound = true;
                    task.GeneralNoteResults = new List<GeneralNoteCheckResult>();
                    var generalNotesAttrs = task.Attributes.Where(a => a.AttributeName == "GeneralNotes").ToList();
                    var dedicationAttrs = task.Attributes.Where(a => a.AttributeName == "Dedication").ToList();
                    string generalNotesValue = NormalizeValue(string.Join(" ", generalNotesAttrs.Select(a => a.Value ?? ""))) ?? "";
                    string dedicationValue = NormalizeValue(string.Join(" ", dedicationAttrs.Select(a => a.Value ?? ""))) ?? "";

                    //get document IDs for CPC101 and Plat categories
                    var documentIdCPC101 = task.DocumentIdsByCategory.ContainsKey("CPC101") ? task.DocumentIdsByCategory["CPC101"] : task.DocumentIds.FirstOrDefault();
                    var documentIdPlat = task.DocumentIdsByCategory.ContainsKey("Plat") ? task.DocumentIdsByCategory["Plat"] : task.DocumentIds.FirstOrDefault();

                    foreach (var noteTask in generalNotesTasks)
                    {
                        logFile.MyLogFile($"Processing general note task: {noteTask.Task} with LogicCode: {noteTask.LogicCode}");

                        bool noteFound = false;
                        string noteResult = "";
                        string noteDetails = "";

                        //Test *****
                        if (noteTask.Task == "UnobstructedVisibilityEasements")
                        {

                        }

                        try
                        {
                            if (noteTask.LogicCode.StartsWith("SearchPlat"))
                            {
                                var searchTerms = noteTask.SearchTerms;

                                // Check jurisdiction eligibility first
                                string currentJurisdiction = titleBlock?.CityETJ ?? "City";
                                bool jurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, currentJurisdiction);
                                if (!jurisdictionMatches)
                                {
                                    // Jurisdiction doesn't match - check if note exists anyway (would be orphaned/wrong jurisdiction)
                                    bool noteExistsInAttribute = !string.IsNullOrEmpty(generalNotesValue) &&
                                                                FuzzyMatch(generalNotesValue, noteTask.NoteTerm);
                                    // Also try search API if fuzzy match didn't find it
                                    if (!noteExistsInAttribute)
                                    {
                                        var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                        if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                        {
                                            logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                            noteExistsInAttribute = true;
                                        }
                                    }
                                    if (noteExistsInAttribute)
                                    {
                                        // WRONG JURISDICTION: Note exists but shouldn't for this jurisdiction
                                        logFile.MyLogFile($"✗ INCORRECT: '{noteTask.Task}' - Note present in General Notes BUT only applicable to {noteTask.Jurisdiction}, current jurisdiction is {currentJurisdiction}");
                                        noteResult = $"{noteTask.Task}: Incorrect";
                                        noteDetails = $"Note present BUT only applicable to {noteTask.Jurisdiction} (current: {currentJurisdiction})";
                                        allNotesFound = false;
                                    }
                                    else
                                    {
                                        // Note correctly absent for this jurisdiction
                                        logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - Not applicable for jurisdiction '{currentJurisdiction}' (Requires: {noteTask.Jurisdiction}) AND note is absent");
                                        noteResult = $"{noteTask.Task}: Correct";
                                        noteDetails = $"Correctly absent - only applicable to {noteTask.Jurisdiction}";
                                    }
                                }
                                else
                                {
                                    // Step 1: Search for the keyword in the document
                                    var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, searchTerms);
                                    bool searchTermFoundInDocument = searchResponse.Success &&
                                                                    searchResponse.Content?.results != null &&
                                                                    searchResponse.Content.results.Count > 0;

                                    int occurrences = searchTermFoundInDocument ? searchResponse.Content.results.Count : 0;

                                    //for visibility triangle occurence count needs to be greater than 2 to be considered detected in the document
                                    //this is bc the term triangle is pretty generic and at least shows as a note
                                    //will use detection instead down the road*****
                                    if (noteTask.Task == "VisibilityTriangle" && occurrences <= 2)
                                    {
                                        searchTermFoundInDocument = false;
                                    }

                                    // Step 2: Check if this note exists in the GeneralNotes attribute
                                    // Use fuzzy matching to check if the note (or similar text) is in GeneralNotes
                                    bool noteExistsInAttribute = !string.IsNullOrEmpty(generalNotesValue) &&
                                                                FuzzyMatch(generalNotesValue, noteTask.NoteTerm);
                                    // Also try search API if fuzzy match didn't find it
                                    if (!noteExistsInAttribute)
                                    {
                                        var searchResponseGeneralNotes = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                        if (searchResponseGeneralNotes.Success && searchResponseGeneralNotes.Content?.results != null && searchResponseGeneralNotes.Content.results.Count > 0)
                                        {
                                            logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                            noteExistsInAttribute = true;
                                        }
                                    }
                                    // Step 3: Cross-check logic
                                    if (searchTermFoundInDocument && noteExistsInAttribute)
                                    {
                                        // VALID: Search term found in document AND note exists in GeneralNotes
                                        noteFound = true;
                                        logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - Search term '{string.Join(", ", searchTerms)}' found in document AND note exists in General Notes");
                                        noteResult = $"{noteTask.Task}: Correct";
                                        noteDetails = $"Search terms '{string.Join(", ", searchTerms)}' found in document AND note present in General Notes";
                                    }
                                    else if (searchTermFoundInDocument && !noteExistsInAttribute)
                                    {
                                        // MISSING: Search term found but note NOT in GeneralNotes
                                        logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' - Search term '{string.Join(", ", searchTerms)}' found in document BUT note MISSING from General Notes");
                                        noteResult = $"{noteTask.Task}: Missing";
                                        noteDetails = $"Search term '{string.Join(", ", searchTerms)}' found in document BUT note is MISSING from General Notes";
                                        allNotesFound = false;
                                    }
                                    else if (!searchTermFoundInDocument && noteExistsInAttribute)
                                    {
                                        // ORPHANED: Note exists in GeneralNotes but search term NOT found in document
                                        logFile.MyLogFile($"✗ ORPHANED: '{noteTask.Task}' - Note exists in General Notes BUT search term '{string.Join(", ", searchTerms)}' NOT found in document");
                                        noteResult = $"{noteTask.Task}: Orphaned";
                                        noteDetails = $"Note present in General Notes BUT search term '{string.Join(", ", searchTerms)}' NOT found in document";
                                        allNotesFound = false;
                                    }
                                    else
                                    {
                                        // NOT FOUND: Neither search term in document NOR note in GeneralNotes
                                        logFile.MyLogFile($"✗ CORRECT: '{noteTask.Task}' - Search term '{string.Join(", ", searchTerms)}' not found in document AND note not in General Notes");
                                        noteResult = $"{noteTask.Task}: Correct";
                                        noteDetails = $"Search term '{string.Join(", ", searchTerms)}' not found in document AND note missing from General Notes";
                                        allNotesFound = false;
                                    }
                                }


                            }
                            else if (noteTask.LogicCode.StartsWith("SearchGeneralNotes"))
                            {
                                var notesToFind = noteTask.SearchTerms;

                                // Check jurisdiction eligibility first
                                string currentJurisdiction = titleBlock?.CityETJ ?? "City";
                                bool jurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, currentJurisdiction);
                                if (!jurisdictionMatches)
                                {
                                    // Jurisdiction doesn't match - check if note exists anyway (would be wrong jurisdiction)
                                    if (generalNotesValue != null && !string.IsNullOrEmpty(generalNotesValue))
                                    {
                                        bool noteExistsInAttribute = FuzzyMatch(generalNotesValue, noteTask.NoteTerm);

                                        // Also try search API if fuzzy match didn't find it
                                        if (!noteExistsInAttribute)
                                        {
                                            var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                            if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                            {
                                                logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                                noteExistsInAttribute = true;
                                            }
                                        }

                                        if (noteExistsInAttribute)
                                        {
                                            // WRONG JURISDICTION: Note exists but shouldn't for this jurisdiction
                                            logFile.MyLogFile($"✗ INCORRECT: '{noteTask.Task}' - Note present in General Notes BUT only applicable to {noteTask.Jurisdiction}, current jurisdiction is {currentJurisdiction}");
                                            noteResult = $"{noteTask.Task}: Incorrect";
                                            noteDetails = $"Note present BUT only applicable to {noteTask.Jurisdiction} (current: {currentJurisdiction})";
                                            allNotesFound = false;
                                        }
                                        else
                                        {
                                            // Note correctly absent for this jurisdiction
                                            logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - Not applicable for jurisdiction '{currentJurisdiction}' (Requires: {noteTask.Jurisdiction}) AND note is absent");
                                            noteResult = $"{noteTask.Task}: Correct";
                                            noteDetails = $"Correctly absent - only applicable to {noteTask.Jurisdiction}";
                                        }
                                    }
                                    else
                                    {
                                        // No general notes attribute - correctly absent for this jurisdiction
                                        logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - Not applicable for jurisdiction '{currentJurisdiction}' AND no general notes present");
                                        noteResult = $"{noteTask.Task}: Correct";
                                        noteDetails = $"Correctly absent - only applicable to {noteTask.Jurisdiction}";
                                    }
                                }
                                else
                                {
                                    // Jurisdiction matches - proceed with normal fuzzy matching
                                    if (generalNotesValue != null && !string.IsNullOrEmpty(generalNotesValue))
                                    {
                                        // Use fuzzy matching due to OCR errors
                                        noteFound = FuzzyMatch(generalNotesValue, noteTask.NoteTerm);

                                        // Also try search API if fuzzy match didn't find it
                                        if (!noteFound)
                                        {
                                            var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                            if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                            {
                                                logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                                noteFound = true;
                                            }
                                        }
                                        if (noteFound)
                                        {
                                            logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' found in General Notes");
                                            noteResult = $"{noteTask.Task}: Correct";
                                            noteDetails = $"'{noteTask.Task}' found in document";
                                        }
                                        else
                                        {
                                            logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' not found in General Notes");
                                            noteResult = $"{noteTask.Task}: Missing";
                                            noteDetails = $"'{noteTask.Task}' not found in document";
                                            allNotesFound = false;
                                        }
                                    }
                                    else
                                    {
                                        logFile.MyLogFile($"✗ ERROR: No General Notes found for task {task.TaskName}");
                                        noteResult = $"{noteTask.Task}: Error";
                                        noteDetails = "General Notes not found or empty";
                                        allNotesFound = false;
                                    }
                                }

                            }
                            else if (noteTask.LogicCode.StartsWith("SearchDedication"))
                            {
                                var searchTerms = noteTask.SearchTerms;
                                // Step 1: Search for the keyword in the Dedication attribute
                                bool searchTermFoundInDedication = !string.IsNullOrEmpty(dedicationValue) &&
                                                                   FuzzyMatch(dedicationValue, searchTerms);
                                // Step 2: Check if this note exists in the GeneralNotes attribute
                                bool noteExistsInAttribute = !string.IsNullOrEmpty(generalNotesValue) &&
                                                            FuzzyMatch(generalNotesValue, noteTask.NoteTerm);
                                // Also try search API if fuzzy match didn't find it
                                if (!noteExistsInAttribute)
                                {
                                    var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                    if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                    {
                                        logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                        noteExistsInAttribute = true;
                                    }
                                }

                                // Step 3: Cross-check logic
                                if (searchTermFoundInDedication && noteExistsInAttribute)
                                {
                                    // VALID: Search term found in Dedication AND note exists in GeneralNotes
                                    noteFound = true;
                                    logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - Search term '{string.Join(", ", searchTerms)}' found in Dedication AND note exists in General Notes");
                                    noteResult = $"{noteTask.Task}: Correct";
                                    noteDetails = $"Search terms '{string.Join(", ", searchTerms)}' found in Dedication AND note present in General Notes";
                                }
                                else if (searchTermFoundInDedication && !noteExistsInAttribute)
                                {
                                    // MISSING: Search term found but note NOT in GeneralNotes
                                    logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' - Search term '{string.Join(", ", searchTerms)}' found in Dedication BUT note MISSING from General Notes");
                                    noteResult = $"{noteTask.Task}: Missing";
                                    noteDetails = $"Search term '{string.Join(", ", searchTerms)}' found in Dedication BUT note is MISSING from General Notes";
                                    allNotesFound = false;
                                }
                                else if (!searchTermFoundInDedication && noteExistsInAttribute)
                                {
                                    // ORPHANED: Note exists in GeneralNotes but search term NOT found in Dedication
                                    logFile.MyLogFile($"✗ ORPHANED: '{noteTask.Task}' - Note exists in General Notes BUT search term '{string.Join(", ", searchTerms)}' NOT found in Dedication");
                                    noteResult = $"{noteTask.Task}: Orphaned";
                                    noteDetails = $"Note present in General Notes BUT search term '{string.Join(", ", searchTerms)}' NOT found in Dedication";
                                    allNotesFound = false;
                                }
                                else
                                {
                                    // NOT FOUND: Neither search term in Dedication NOR note in GeneralNotes
                                    logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - Search term '{string.Join(", ", searchTerms)}' not found in Dedication AND note not in General Notes");
                                    noteResult = $"{noteTask.Task}: Correct";
                                    noteDetails = $"Search term '{string.Join(", ", searchTerms)}' not found in Dedication AND note missing from General Notes";
                                    allNotesFound = false;
                                }
                            }
                            else if (noteTask.LogicCode.StartsWith("Misc"))
                            {
                                //switch on task name
                                switch (noteTask.Task)
                                {
                                    //PAE
                                    case "WasteCollectionInelligble":
                                    case "WasteCollectionElligble":
                                    case "PermanantAccessEasement":
                                    case "PrivateWaterSystem":
                                        //use ispermanentaccesseasementasync method to determine if the easement is present. If true, note should exist. If false, note should not exist.
                                        bool easementDetected = await IsPermanentAccessEasementAsync(documentIdCPC101, task);

                                        // Check if this note applies to the current jurisdiction
                                        string privateWaterJurisdiction = titleBlock?.CityETJ ?? "City";
                                        bool privateWaterJurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, privateWaterJurisdiction);

                                        // Combined condition: easement must be detected AND jurisdiction matches
                                        bool privateWaterCondition = easementDetected && privateWaterJurisdictionMatches;
                                        string privateWaterConditionDescription = privateWaterJurisdictionMatches
                                            ? $"PermanentAccessEasement detected (Jurisdiction: {privateWaterJurisdiction})"
                                            : $"Not applicable to {privateWaterJurisdiction} jurisdiction";

                                        var pwResult = await ValidateGeneralNoteByCondition(
                                            privateWaterCondition,
                                            privateWaterConditionDescription,
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = pwResult.NoteFound;
                                        noteResult = pwResult.NoteResult;
                                        noteDetails = pwResult.NoteDetails;
                                        allNotesFound = pwResult.AllFound;
                                        break;
                                    //Single Family Residential
                                    case "DedicationFeePercentage":
                                    case "UtilityEasementExtensions":
                                    case "DedicationFeeUnits":
                                    case "Lots":
                                        //If lot count > 0 and jurisdiction matches, note should exist. If lot count = 0 or jurisdiction doesn't match, note should not exist.
                                        int lotsCount = titleBlock?.Lots ?? 0;
                                        string lotsJurisdiction = titleBlock?.CityETJ ?? "City";
                                        string taskType = noteTask.Task == "UtilityEasementExtensions" ? "Utility Easement Extensions" : "Lots";

                                        // Check if this note applies to the current jurisdiction
                                        bool lotsJurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, lotsJurisdiction);

                                        // Condition: lots > 0 AND jurisdiction matches
                                        bool lotsConditionMet = lotsCount > 0 && lotsJurisdictionMatches;

                                        var lotsResult = await ValidateGeneralNoteByCondition(
                                            lotsConditionMet,
                                            lotsJurisdictionMatches ? $"{taskType} condition met (Lots: {lotsCount}, Jurisdiction: {lotsJurisdiction})" : $"Not applicable to {lotsJurisdiction} jurisdiction",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = lotsResult.NoteFound;
                                        noteResult = lotsResult.NoteResult;
                                        noteDetails = lotsResult.NoteDetails;
                                        allNotesFound = lotsResult.AllFound;
                                        break;
                                    //Misc
                                    case "BlockNumber":
                                        //get count of BlockNumber attribute from Plat document using API to retrive attibutes and if count of attributes named block number > 0 then the note should exist. If the note exists and there are no BlockNumbers then it's an orphaned note.
                                        var attributes = await agileDDApi.ReadValuesAsync(documentIdPlat);
                                        var blockNumberCount = attributes.Content.Count(a => a.attribute_id == platBlockNumberAttributeId && a.position != null && a.position.Count > 0);
                                        bool blockConditionMet = blockNumberCount > 0;
                                        var blockResult = await ValidateGeneralNoteByCondition(
                                            blockConditionMet,
                                            "BlockNumber condition met",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = blockResult.NoteFound;
                                        noteResult = blockResult.NoteResult;
                                        noteDetails = blockResult.NoteDetails;
                                        allNotesFound = blockResult.AllFound;
                                        break;
                                    case "GarageBuildingLine2":
                                        //If search term is detected in the plat and single family residential (lots > 0), must have note. search term not detected, note should not exist.
                                        //use the search api to see if the search terms is present in the plat document
                                        var searchResponseGarage = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.SearchTerms);
                                        bool garageLine2Detected = searchResponseGarage.Success &&
                                                                  searchResponseGarage.Content?.results != null &&
                                                                  searchResponseGarage.Content.results.Count > 0;
                                        int lotCount = titleBlock?.Lots ?? 0;
                                        bool garageConditionMet = garageLine2Detected && lotCount > 0;
                                        var garageResult = await ValidateGeneralNoteByCondition(
                                            garageConditionMet,
                                            "GarageBuildingLine2 condition met",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = garageResult.NoteFound;
                                        noteResult = garageResult.NoteResult;
                                        noteDetails = garageResult.NoteDetails;
                                        allNotesFound = garageResult.AllFound;
                                        break;
                                    case "SharedDriveway":
                                        //use issharedDrivewayasync method to determine if shared driveway is present. If true, note should exist. If false, note should not exist.
                                        bool sharedDrivewayDetected = await IsSharedDrivewayAsync(documentIdCPC101);
                                        var sdResult = await ValidateGeneralNoteByCondition(
                                            sharedDrivewayDetected,
                                            "SharedDriveway condition met",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = sdResult.NoteFound;
                                        noteResult = sdResult.NoteResult;
                                        noteDetails = sdResult.NoteDetails;
                                        allNotesFound = sdResult.AllFound;
                                        break;
                                    case "PrivateWaterSystem2":
                                        //use IsPermanentAccessEasementAsync method and isshareddrivewayasync to determine if there is a permanent access easement and shared driveway for private water system. If either is true, note should exist. If both false, note should not exist.
                                        bool permanentAccessEasementDetected = await IsPermanentAccessEasementAsync(documentIdCPC101, task);
                                        bool sharedDrivewayForPrivateWaterDetected = await IsSharedDrivewayAsync(documentIdCPC101);
                                        bool privateWater2Condition = permanentAccessEasementDetected || sharedDrivewayForPrivateWaterDetected;
                                        var pw2Result = await ValidateGeneralNoteByCondition(
                                            privateWater2Condition,
                                            "PrivateWaterSystem2 condition met",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = pw2Result.NoteFound;
                                        noteResult = pw2Result.NoteResult;
                                        noteDetails = pw2Result.NoteDetails;
                                        allNotesFound = pw2Result.AllFound;
                                        break;
                                    case "Drainage":
                                        //if the reserve from titleblock is > 0 and the lots = 0 and the jurisdiction matches, then the note should exist. If not, note should not exist.
                                        double reserveArea = titleBlock?.Reserves ?? 0;
                                        string drainageJurisdiction = titleBlock?.CityETJ ?? "City";
                                        int lots = titleBlock?.Lots ?? 0;

                                        // Check if this note applies to the current jurisdiction
                                        bool drainageJurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, drainageJurisdiction);

                                        // Condition: reserves > 0 AND lots = 0 AND jurisdiction matches
                                        bool drainageCondition = reserveArea > 0 && lots == 0 && drainageJurisdictionMatches;

                                        var dragResult = await ValidateGeneralNoteByCondition(
                                            drainageCondition,
                                            drainageJurisdictionMatches ? $"Drainage condition met (Reserves: {reserveArea}, Lots: {lots}, Jurisdiction: {drainageJurisdiction})" : $"Not applicable to {drainageJurisdiction} jurisdiction",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = dragResult.NoteFound;
                                        noteResult = dragResult.NoteResult;
                                        noteDetails = dragResult.NoteDetails;
                                        allNotesFound = dragResult.AllFound;
                                        break;
                                    case "StormWaterQuality":
                                        //if the titleblock has acreage > 5 and the jurisdiction matches, then the note should exist. If not, note should not exist.
                                        decimal acreage = titleBlock?.Acreage ?? 0;
                                        string stormqualityJurisdiction = titleBlock?.CityETJ ?? "City";

                                        // Check if this note applies to the current jurisdiction
                                        bool stormqualityJurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, stormqualityJurisdiction);

                                        // Condition: acreage > 5 AND jurisdiction matches
                                        bool stormqualityCondition = acreage > 5 && stormqualityJurisdictionMatches;

                                        var swqResult = await ValidateGeneralNoteByCondition(
                                            stormqualityCondition,
                                            stormqualityJurisdictionMatches ? $"StormWaterQuality condition met (Acreage: {acreage}, Jurisdiction: {stormqualityJurisdiction})" : $"Not applicable to {stormqualityJurisdiction} jurisdiction",
                                            noteTask.NoteTerm,
                                            documentIdPlat,
                                            generalNotesValue,
                                            noteTask,
                                            allNotesFound);
                                        noteFound = swqResult.NoteFound;
                                        noteResult = swqResult.NoteResult;
                                        noteDetails = swqResult.NoteDetails;
                                        allNotesFound = swqResult.AllFound;
                                        break;
                                    default:
                                        logFile.MyLogFile($"Unhandled Misc LogicCode task: {noteTask.Task}");
                                        noteResult = $"{noteTask.Task}: Unhandled misc logic";
                                        noteDetails = $"Misc LogicCode task '{noteTask.Task}' not implemented";
                                        allNotesFound = false;
                                        break;
                                }
                            }
                            else
                            {
                                // Handle on a case by case basis - for now, mark as unhandled
                                logFile.MyLogFile($"Unhandled LogicCode for GeneralNotes: {noteTask.LogicCode}");
                                noteResult = $"{noteTask.Task}: Unhandled logic";
                                noteDetails = $"LogicCode '{noteTask.LogicCode}' not implemented";
                                allNotesFound = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            logFile.MyLogFile($"Error processing general note task '{noteTask.Task}': {ex.Message}");
                            noteResult = $"{noteTask.Task}: Error";
                            noteDetails = $"Exception: {ex.Message}";
                            allNotesFound = false;
                        }

                        // Store detailed result
                        task.GeneralNoteResults.Add(new GeneralNoteCheckResult
                        {
                            GeneralNoteID = noteTask.GeneralNoteID,
                            Task = noteTask.Task,
                            LogicCode = noteTask.LogicCode,
                            Found = noteFound,
                            Details = noteDetails,
                            Result = noteResult
                        });
                    }

                    // Set task status based on results
                    if (allNotesFound)
                    {
                        task.Status = "Correct";
                        logFile.MyLogFile("All general notes checks passed");
                    }
                    else
                    {
                        task.Status = "Error";
                        logFile.MyLogFile("Some general notes checks failed");
                    }
                } //LP-HTX
                if (task.LogicType == "GeneralNotes" && task.TaskSet == "Transportation")
                {
                    logFile.MyLogFile($"Handling Transportation General Notes");

                    //call GetGeneralNotesTasksAsync to get the list of General Notes tasks
                    var generalNotesTasks = await GetGeneralNotesTasksAsync();
                    generalNotesTasks = generalNotesTasks
                        .Where(t => t.LogicCode == "SearchTransportationGeneralNotes")
                        .ToList();
                    logFile.MyLogFile($"Found {generalNotesTasks.Count} Transportation General Notes tasks from database");

                    // List out all general notes tasks found
                    foreach (var noteTask in generalNotesTasks)
                    {
                        logFile.MyLogFile($"  - Task: {noteTask.Task}, LogicCode: {noteTask.LogicCode}, ID: {noteTask.GeneralNoteID}, NoteTerm: {noteTask.NoteTerm ?? "N/A"}, SearchTerms: {(noteTask.SearchTerms.Any() ? string.Join(", ", noteTask.SearchTerms) : "N/A")}");
                    }
                    if (!generalNotesTasks.Any())
                    {
                        logFile.MyLogFile("No general notes tasks found.");
                        task.Status = "Error";
                        continue;
                    }

                    bool allNotesFound = true;
                    task.GeneralNoteResults = new List<GeneralNoteCheckResult>();

                    //get document IDs
                    var documentId = task.DocumentIds.FirstOrDefault();

                    if (documentId > 0)
                    {
                        var attributes = await agileDDApi.ReadValuesAsync(documentId);
                        if (attributes.Success && attributes.Content != null)
                        {
                            logFile.MyLogFile($"Retrieved {attributes.Content.Count} attributes for Transportation document {documentId}");

                            // Update task attributes with actual values from API
                            foreach (var attr in task.Attributes)
                            {
                                var matchedAttribute = attributes.Content.FirstOrDefault(a => a.attribute_id == attr.AttributeCode);
                                if (matchedAttribute != null)
                                {
                                    if (matchedAttribute.classifications != null &&
                                        matchedAttribute.classifications.Count > 0 &&
                                        matchedAttribute.position != null &&
                                        matchedAttribute.position.Any())
                                    {
                                        attr.Value = await agileDDApi.GetTextFromBoxAsync(documentId, matchedAttribute.page_number, matchedAttribute.position);
                                    }
                                    else
                                    {
                                        attr.Value = matchedAttribute.data ?? "N/A";
                                    }
                                    logFile.MyLogFile($"Updated Transportation attribute '{attr.AttributeName}' with value: {attr.Value}");
                                }
                            }
                        }
                        else
                        {
                            logFile.MyLogFile($"Failed to retrieve attributes for Transportation document {documentId}: {attributes.ErrorMessage}");
                        }
                    }

                    // Extract generalNotesValue
                    var generalNotesAttrs = task.Attributes.Where(a => a.AttributeName == "General Notes").ToList();
                    string generalNotesValue = NormalizeValue(string.Join(" ", generalNotesAttrs.Select(a => a.Value ?? ""))) ?? "";
                    logFile.MyLogFile($"Extracted General Notes attribute value: '{generalNotesValue}'");

                    foreach (var noteTask in generalNotesTasks)
                    {
                        logFile.MyLogFile($"Processing general note task: {noteTask.Task} with LogicCode: {noteTask.LogicCode}");

                        bool noteFound = false;
                        string noteResult = "";
                        string noteDetails = "";

                        try
                        {
                            if (noteTask.LogicCode.StartsWith("SearchTransportationGeneralNotes"))
                            {
                                if (generalNotesValue != null && !string.IsNullOrEmpty(generalNotesValue))
                                {
                                    // Use fuzzy matching due to OCR errors
                                    noteFound = FuzzyMatch(generalNotesValue, noteTask.NoteTerm);

                                    // Also try search API if fuzzy match didn't find it
                                    if (!noteFound)
                                    {
                                        // Trim and split NoteTerm by pipe to create a list of search terms
                                        List<string> searchTerms = (noteTask.NoteTerm ?? "")
                                            .Trim()
                                            .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(term => term.Trim())
                                            .Where(term => !string.IsNullOrEmpty(term))
                                            .ToList();

                                        logFile.MyLogFile($"Searching for note '{noteTask.Task}' with {searchTerms.Count} term(s): {string.Join(", ", searchTerms)}");

                                        // Call SearchDocumentAsync with the list of terms
                                        var searchResponse = await agileDDApi.SearchDocumentAsync(documentId, searchTerms);
                                        if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                        {
                                            logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentId}");
                                            noteFound = true;
                                        }
                                    }
                                    if (noteFound)
                                    {
                                        logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' found in General Notes");
                                        noteResult = $"{noteTask.Task}: Correct";
                                        noteDetails = $"'{noteTask.Task}' found in document";
                                    }
                                    else
                                    {
                                        logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' not found in General Notes");
                                        noteResult = $"{noteTask.Task}: Missing";
                                        noteDetails = $"'{noteTask.Task}' not found in document";
                                        allNotesFound = false;
                                    }
                                }
                                else
                                {
                                    logFile.MyLogFile($"✗ ERROR: No General Notes found for task {task.TaskName}");
                                    noteResult = $"{noteTask.Task}: Error";
                                    noteDetails = "General Notes not found or empty";
                                    allNotesFound = false;
                                }
                            }

                            else
                            {
                                // Handle on a case by case basis - for now, mark as unhandled
                                logFile.MyLogFile($"Unhandled LogicCode for GeneralNotes: {noteTask.LogicCode}");
                                noteResult = $"{noteTask.Task}: Unhandled logic";
                                noteDetails = $"LogicCode '{noteTask.LogicCode}' not implemented";
                                allNotesFound = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            logFile.MyLogFile($"Error processing general note task '{noteTask.Task}': {ex.Message}");
                            noteResult = $"{noteTask.Task}: Error";
                            noteDetails = $"Exception: {ex.Message}";
                            allNotesFound = false;
                        }

                        // Store detailed result
                        task.GeneralNoteResults.Add(new GeneralNoteCheckResult
                        {
                            GeneralNoteID = noteTask.GeneralNoteID,
                            Task = noteTask.Task,
                            LogicCode = noteTask.LogicCode,
                            Found = noteFound,
                            Details = noteDetails,
                            Result = noteResult
                        });
                    }

                    // Set task status based on results
                    if (allNotesFound)
                    {
                        task.Status = "Correct";
                        logFile.MyLogFile("All general notes checks passed");
                    }
                    else
                    {
                        task.Status = "Error";
                        logFile.MyLogFile("Some general notes checks failed");
                    }
                }

                if (task.LogicType == "Dedication")
                {
                    logFile.MyLogFile($"Handling Dedication tasks");

                    //call GetDedicationTasksAsync to get the list of Dedication tasks
                    var dedicationTasks = await GetDedicationTasksAsync();
                    if (!dedicationTasks.Any())
                    {
                        logFile.MyLogFile("No Dedication tasks found.");
                        task.Status = "Error";
                        continue;
                    }

                    bool allFound = true;
                    task.DedicationResults = new List<DedicationCheckResult>();
                    var dedicationAttrs = task.Attributes.Where(a => a.AttributeName == "Dedication").ToList();
                    string dedicationValue = NormalizeValue(string.Join(" ", dedicationAttrs.Select(a => a.Value ?? ""))) ?? "";

                    //get document IDs for CPC101 and Plat categories
                    var documentIdCPC101 = task.DocumentIdsByCategory.ContainsKey("CPC101") ? task.DocumentIdsByCategory["CPC101"] : task.DocumentIds.FirstOrDefault();
                    var documentIdPlat = task.DocumentIdsByCategory.ContainsKey("Plat") ? task.DocumentIdsByCategory["Plat"] : task.DocumentIds.FirstOrDefault();

                    //get values for county, acreage and lots from dedication attribute by passing the text of the dedication attribute to an LLM
                    var dedicationInfo = await ParseDedicationAttributeAsync(dedicationValue, logFile); //***** consider a form of redundancy if dedication attribute fails
                    string county = dedicationInfo.County;
                    decimal acreage = dedicationInfo.Acreage;
                    string developmentName = dedicationInfo.DevelopmentName;

                    foreach (var noteTask in dedicationTasks)
                    {
                        logFile.MyLogFile($"Processing general note task: {noteTask.Task} with LogicCode: {noteTask.LogicCode}");

                        bool noteFound = false;
                        string noteResult = "";
                        string noteDetails = "";

                        try
                        {
                            if (noteTask.LogicCode.StartsWith("SearchDedication"))
                            {
                                // Check jurisdiction eligibility first
                                string currentJurisdiction = titleBlock?.CityETJ ?? "City";
                                bool jurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, currentJurisdiction);

                                noteFound = FuzzyMatch(dedicationValue, noteTask.NoteTerm);

                                // Also try search API if fuzzy match didn't find it
                                if (!noteFound)
                                {
                                    var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                    if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                    {
                                        logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                        noteFound = true;
                                    }
                                }

                                // Also try search for the note term in the document using the search API
                                if (!noteFound && jurisdictionMatches)
                                {
                                    var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTask.NoteTerm);
                                    if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                                    {
                                        logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                                        noteFound = true;
                                    }
                                }

                                if (!jurisdictionMatches)
                                {
                                    // Jurisdiction doesn't match - check if note exists anyway (would be wrong jurisdiction)
                                    if (dedicationValue != null && !string.IsNullOrEmpty(dedicationValue))
                                    {
                                        if (noteFound)
                                        {
                                            // WRONG JURISDICTION: Note exists but shouldn't for this jurisdiction
                                            logFile.MyLogFile($"✗ INCORRECT: '{noteTask.Task}' present in Dedication BUT only applicable to {noteTask.Jurisdiction}, current jurisdiction is {currentJurisdiction}");
                                            noteResult = $"{noteTask.Task}: Incorrect";
                                            noteDetails = $"Note present BUT only applicable to {noteTask.Jurisdiction} (current: {currentJurisdiction})";
                                            allFound = false;
                                        }
                                        else
                                        {
                                            // Note correctly absent for this jurisdiction
                                            logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' applicable for jurisdiction '{currentJurisdiction}' (Requires: {noteTask.Jurisdiction}) and note is absent");
                                            noteResult = $"{noteTask.Task}: Correct";
                                            noteDetails = $"Correctly absent - only applicable to {noteTask.Jurisdiction}";
                                        }
                                    }
                                    else
                                    {
                                        // No dedication attribute - correctly absent for this jurisdiction
                                        logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' applicable for jurisdiction '{currentJurisdiction}' and no Dedication present");
                                        noteResult = $"{noteTask.Task}: Correct";
                                        noteDetails = $"Correctly absent - only applicable to {noteTask.Jurisdiction}";
                                    }
                                }
                                else // Jurisdiction matches
                                {
                                    if (dedicationValue != null && !string.IsNullOrEmpty(dedicationValue))
                                    {
                                        if (noteFound)
                                        {
                                            logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' found in Dedication");
                                            noteResult = $"{noteTask.Task}: Correct";
                                            noteDetails = $"'{noteTask.Task}' found in document";
                                        }
                                        else
                                        {
                                            logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' not found in Dedication");
                                            noteResult = $"{noteTask.Task}: Missing";
                                            noteDetails = $"'{noteTask.Task}' not found in document";
                                            allFound = false;
                                        }
                                    }
                                    else
                                    {
                                        logFile.MyLogFile($"✗ ERROR: No dedication attribute found for task {task.TaskName}");
                                        noteResult = $"{noteTask.Task}: No attribute data";
                                        noteDetails = "Dedication attribute not found or empty";
                                        allFound = false;
                                    }
                                }

                            }
                            else if (noteTask.LogicCode.StartsWith("CrossCheck"))
                            {
                                // using the relevant field from the dedicationInfo object, compare the value against the relevant titleblock value and the relevant CPC101 attribute value
                                // for example, if the noteTask.Task is "County", compare dedicationInfo.County against titleBlock.County and the CPC101 attribute for County
                                // the results should indicate if all three match, or if there is a mismatch in a specific area
                                string dedicationFieldValue = "";
                                string titleBlockFieldValue = "";
                                string cpc101FieldValue = "";
                                switch (noteTask.Task)
                                {
                                    case "County":
                                        dedicationFieldValue = county;
                                        titleBlockFieldValue = titleBlock?.County ?? "";
                                        var countyAttr = task.Attributes.FirstOrDefault(a => a.FileCategory == "CPC101" && a.AttributeName == "County");
                                        cpc101FieldValue = countyAttr?.Value ?? "";
                                        break;
                                    case "Acreage":
                                        dedicationFieldValue = acreage.ToString();
                                        titleBlockFieldValue = (titleBlock?.Acreage ?? 0).ToString();
                                        var acreageAttr = task.Attributes.FirstOrDefault(a => a.FileCategory == "CPC101" && a.AttributeName == "Acreage");
                                        cpc101FieldValue = acreageAttr?.Value ?? "";
                                        break;
                                    case "DevelopmentName":
                                        dedicationFieldValue = developmentName;
                                        var devNamePlatAttr = task.Attributes.FirstOrDefault(a => a.FileCategory == "Plat" && a.AttributeName == "PlatName");
                                        titleBlockFieldValue = devNamePlatAttr?.Value ?? "";
                                        var devNameCPC101Attr = task.Attributes.FirstOrDefault(a => a.FileCategory == "CPC101" && a.AttributeName == "PlatName");
                                        cpc101FieldValue = devNameCPC101Attr?.Value ?? "";
                                        break;
                                    default:
                                        logFile.MyLogFile($"Unhandled CrossCheck task: {noteTask.Task}");
                                        noteResult = $"{noteTask.Task}: Unhandled crosscheck task";
                                        noteDetails = $"CrossCheck task '{noteTask.Task}' not implemented";
                                        allFound = false;
                                        continue;
                                }

                                // Normalize and compare values
                                bool valuesMatch = CompareFieldValues(
                                    dedicationFieldValue,
                                    titleBlockFieldValue,
                                    cpc101FieldValue,
                                    out string normalizedDedication,
                                    out string normalizedTitleBlock,
                                    out string normalizedCpc101);

                                if (valuesMatch)
                                {
                                    logFile.MyLogFile($"✓ VALID: {noteTask.Task} matches across Dedication, TitleBlock, and CPC101");
                                    noteResult = $"{noteTask.Task}: Match";
                                    noteDetails = $"{noteTask.Task} matches in Dedication ('{normalizedDedication}'), TitleBlock ('{normalizedTitleBlock}'), and CPC101 ('{normalizedCpc101}')";
                                }
                                else
                                {
                                    logFile.MyLogFile($"✗ MISMATCH: {noteTask.Task} does not match across Dedication, TitleBlock, and CPC101");
                                    noteResult = $"{noteTask.Task}: Mismatch";
                                    noteDetails = $"{noteTask.Task} values - Dedication: '{normalizedDedication}', TitleBlock: '{normalizedTitleBlock}', CPC101: '{normalizedCpc101}'";
                                    allFound = false;
                                }
                            }
                            else if (noteTask.LogicCode.StartsWith("Misc"))
                            {
                                //switch on task name
                                switch (noteTask.Task)
                                {
                                    case "LotsParagraph":
                                        //If the title block has lots > 0, must have paragraph based on note term
                                        int lotsCount = titleBlock?.Lots ?? 0;
                                        bool lotsConditionMet = lotsCount > 0;
                                        var lotsResult = await ValidateDedicationByCondition(
                                            lotsConditionMet,
                                            $"LotsParagraph condition met (Lots: {lotsCount})",
                                            noteTask.NoteTerm,
                                            dedicationValue,
                                            noteTask,
                                            documentIdPlat,
                                            allFound);
                                        noteFound = lotsResult.NoteFound;
                                        noteResult = lotsResult.NoteResult;
                                        noteDetails = lotsResult.NoteDetails;
                                        allFound = lotsResult.AllFound;
                                        break;
                                    case "PAEParagraph":
                                        //use ispermanentaccesseasementasync method to determine if the easement is present. If true, note should exist. If false, note should not exist.
                                        bool easementDetected = await IsPermanentAccessEasementAsync(documentIdCPC101, task);

                                        // Check if this note applies to the current jurisdiction
                                        string privateWaterJurisdiction = titleBlock?.CityETJ ?? "City";
                                        bool privateWaterJurisdictionMatches = IsNoteApplicableToJurisdiction(noteTask.Jurisdiction, privateWaterJurisdiction);

                                        // Combined condition: easement must be detected AND jurisdiction matches
                                        bool privateWaterCondition = easementDetected && privateWaterJurisdictionMatches;
                                        string privateWaterConditionDescription = privateWaterJurisdictionMatches
                                            ? $"PermanentAccessEasement detected (Jurisdiction: {privateWaterJurisdiction})"
                                            : $"Not applicable to {privateWaterJurisdiction} jurisdiction";

                                        var paeResult = await ValidateDedicationByCondition(
                                            privateWaterCondition,
                                            privateWaterConditionDescription,
                                            noteTask.NoteTerm,
                                            dedicationValue,
                                            noteTask,
                                            documentIdPlat,
                                            allFound);
                                        noteFound = paeResult.NoteFound;
                                        noteResult = paeResult.NoteResult;
                                        noteDetails = paeResult.NoteDetails;
                                        allFound = paeResult.AllFound;
                                        break;
                                    case "ReasonForReplatParagraph":
                                        //the text 'replat' is anywhere in the titleblock attribute value, then the note should exist. If 'replat' is not present, note should not exist.
                                        var titleBlockNotesAttr = task.Attributes.FirstOrDefault(a => a.FileCategory == "Plat" && a.AttributeName == "TitleBlock");
                                        string titleBlockNotesValue = NormalizeValue(titleBlockNotesAttr?.Value) ?? "";
                                        bool replatDetected = titleBlockNotesValue.Contains("replat");
                                        var replatResult = await ValidateDedicationByCondition(
                                            replatDetected,
                                            "ReasonForReplatParagraph condition met",
                                            noteTask.NoteTerm,
                                            dedicationValue,
                                            noteTask,
                                            documentIdPlat,
                                            allFound);
                                        noteFound = replatResult.NoteFound;
                                        noteResult = replatResult.NoteResult;
                                        noteDetails = replatResult.NoteDetails;
                                        allFound = replatResult.AllFound;
                                        break;
                                    default:
                                        logFile.MyLogFile($"Unhandled Misc LogicCode task: {noteTask.Task}");
                                        noteResult = $"{noteTask.Task}: Unhandled misc logic";
                                        noteDetails = $"Misc LogicCode task '{noteTask.Task}' not implemented";
                                        allFound = false;
                                        break;
                                }
                            }
                            else
                            {
                                // Handle on a case by case basis - for now, mark as unhandled
                                logFile.MyLogFile($"Unhandled LogicCode for GeneralNotes: {noteTask.LogicCode}");
                                noteResult = $"{noteTask.Task}: Unhandled logic";
                                noteDetails = $"LogicCode '{noteTask.LogicCode}' not implemented";
                                allFound = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            logFile.MyLogFile($"Error processing general note task '{noteTask.Task}': {ex.Message}");
                            noteResult = $"{noteTask.Task}: Error";
                            noteDetails = $"Exception: {ex.Message}";
                            allFound = false;
                        }

                        // Store detailed result
                        task.DedicationResults.Add(new DedicationCheckResult
                        {
                            Task = noteTask.Task,
                            LogicCode = noteTask.LogicCode,
                            Found = noteFound,
                            Details = noteDetails,
                            Result = noteResult
                        });
                    }

                    // Set task status based on results
                    if (allFound)
                    {
                        task.Status = "Correct";
                        logFile.MyLogFile("All dedication checks passed");
                    }
                    else
                    {
                        task.Status = "Error";
                        logFile.MyLogFile("Some dedication checks failed");
                    }
                } //LP-HTX

                if (task.LogicType == "TitleBlockDetailSection")
                {
                    logFile.MyLogFile($"Handling Title Block Detail Section per-page validation using LLM");

                    // Get document ID
                    var documentId = task.DocumentIds.FirstOrDefault();

                    // Expected values from cover sheet (will be extracted from first valid page)
                    TitleBlockDetailSection expectedValues = null;

                    try
                    {
                        // Get total number of pages in the document
                        var docAttributes = await agileDDApi.ReadValuesAsync(documentId);
                        if (!docAttributes.Success)
                        {
                            logFile.MyLogFile($"Failed to get document attributes: {docAttributes.ErrorMessage}");
                            task.Status = "Error";
                            continue;
                        }

                        int maxPage = docAttributes.Content.Any() ? docAttributes.Content.Max(a => a.page_number) : 1;
                        logFile.MyLogFile($"Document has {maxPage} pages. Extracting Title Block Detail Section from each page using LLM...");

                        // Title Block Detail Section attribute ID from configuration
                        int titleBlockAttrId = transportationTitleBlockDetailSectionAttributeId;

                        // Track Title Block Detail Section presence
                        int totalPages = maxPage;
                        int pagesWithTitleBlock = 0;
                        List<int> pagesWithoutTitleBlock = new List<int>();

                        // Define all fields to check (REMOVED SheetNumber from validation fields)
                        var fieldsToCheck = new Dictionary<string, string>();

                        // Track results per field
                        var fieldResults = new Dictionary<string, (int PagesChecked, int PagesWithValue, List<int> MissingPages)>();

                        // Initialize Azure OpenAI API with GPT-4o vision model
                        var azureVisionApi = new AzureOpenAIApi(
                            logFile,
                            "You are an expert at extracting structured data from title block detail sections on transportation and engineering drawings.",
                            Convert.ToDecimal(0.2),
                            useVision: true  // Enable GPT-4o vision model
                        );

                        // Vision prompt for extraction
                        string visionPrompt = @"Extract the following values from this transportation engineering title block detail section image. Return ONLY a valid JSON object.

                            CRITICAL: Return ONLY raw JSON. Do NOT wrap in markdown code blocks. Do NOT include any explanatory text.

                            Fields to extract:
                            - CONT (Control Number)
                            - SECT (Section Number)
                            - JOB (Job Number)
                            - SHEET (Sheet Number) - IMPORTANT: Extract this separately
                            - DIST (District)
                            - COUNTY (County)
                            - HIGHWAY (Highway Number)

                            IMPORTANT - CONT/SECT/JOB extraction:
                            These three fields may appear in TWO different formats:

                            FORMAT A - Separate cells (extract each value individually):
                                CONT.  SECT.  JOB
                                0902   20     110
                                → ControlNumber=""0902"", SectionNumber=""20"", JobNumber=""110""

                            FORMAT B - Combined as PROJECT NO. with dashes (split on the dashes):
                                PROJECT NO.
                                0922-11-043
                                → ControlNumber=""0922"", SectionNumber=""11"", JobNumber=""043""

                            In BOTH cases, always output ControlNumber, SectionNumber, and JobNumber as SEPARATE fields.

                            SHEET NUMBER EXTRACTION:
                            The sheet number appears in various locations:
                            - Look for ""SHEET NO."" or ""SHEET"" label followed by a number
                            - May appear after PROJECT NO. on same row
                            - In standard sheets: look for ""DIST  COUNTY  SHEET NO."" label row, then values row
                            - Sheet number is the page number of this specific sheet in the set
                            - Examples: ""15"", ""20A"", ""33""

                            Field Name Variations:
                            - CONT may appear as: ""CONT."", ""CONTROL NO."", ""CONTROL""
                            - SECT may appear as: ""SECT."", ""SECTION"", ""SECTION NO.""
                            - JOB may appear as: ""JOB."", ""JOB NO.""
                            - SHEET may appear as: ""SHEET NO."", ""SHEET"", ""SHT.""
                            - DIST may appear as: ""STATE DIST. NO"", ""DIST."", ""DIST NO."", district abbreviation like ""FTW""
                            - COUNTY may appear as: ""CO."", ""COUNTY OF""
                            - HIGHWAY may appear as: ""HIGHWAY NO."", ""HWY""

                            Normalize formatting:
                            - CONT → preserve as-is (e.g., ""0902"")
                            - SECT → preserve as-is (e.g., ""20"")
                            - JOB → preserve as-is (e.g., ""110"")
                            - SHEET → preserve as-is (e.g., ""15"", ""20A"")
                            - DIST → full district name if present, otherwise abbreviation
                            - COUNTY → Proper case (e.g., ""Wise"")
                            - HIGHWAY → preserve as-is

                            Null Value Rules:
                            If a field is present → extract its value normally.
                            If a field is not present in the image → return null.

                            Output Format (return ONLY this JSON, nothing else):
                            {
                                ""ControlNumber"": null,
                                ""SectionNumber"": null,    
                                ""JobNumber"": null,
                                ""SheetNumber"": null,
                                ""District"": null,
                                ""County"": null,
                                ""HighwayNumber"": null
                            }";

                        // Check each page
                        for (int pageNumber = 1; pageNumber <= maxPage; pageNumber++)
                        {
                            var pageAttributes = await agileDDApi.ReadValuesAsync(documentId, pageNumber, titleBlockAttrId);

                            if (!pageAttributes.Success || pageAttributes.Content == null || !pageAttributes.Content.Any())
                            {
                                logFile.MyLogFile($"Page {pageNumber}: Title Block Detail Section not found");
                                pagesWithoutTitleBlock.Add(pageNumber);
                                continue;
                            }

                            // Get title block attribute with position data
                            var titleBlockAttr = pageAttributes.Content.FirstOrDefault(a =>
                                a.attribute_id == titleBlockAttrId &&
                                a.position != null &&
                                a.position.Any());

                            if (titleBlockAttr == null)
                            {
                                logFile.MyLogFile($"Page {pageNumber}: Title Block Detail Section attribute found but no position data");
                                pagesWithoutTitleBlock.Add(pageNumber);
                                continue;
                            }

                            // This page HAS Title Block Detail Section
                            pagesWithTitleBlock++;

                            // Get cropped image of the Title Block Detail Section
                            var pos = titleBlockAttr.position[0];
                            var croppedImageBytes = await agileDDApi.GetCroppedImageAsync(
                                documentId,
                                pageNumber,
                                pos.left,
                                pos.top,
                                pos.width,
                                pos.height
                            );

                            if (croppedImageBytes == null || croppedImageBytes.Length == 0)
                            {
                                logFile.MyLogFile($"Page {pageNumber}: Failed to crop Title Block Detail Section image");
                                pagesWithoutTitleBlock.Add(pageNumber);
                                continue;
                            }

                            // Query GPT-4o vision model with the cropped image
                            var visionResult = await azureVisionApi.QueryImageAsync(croppedImageBytes, visionPrompt);

                            if (!visionResult.Success)
                            {
                                logFile.MyLogFile($"Page {pageNumber}: Error querying GPT-4o vision: {visionResult.ErrorMessage}");
                                pagesWithoutTitleBlock.Add(pageNumber);
                                continue;
                            }

                            // Parse the JSON response
                            try
                            {
                                string cleanedJson = Regex.Replace(visionResult.Content, @"```(?:json)?\s*", "").Trim();
                                string jsonContent = ExtractFirstJsonObject(cleanedJson) ?? cleanedJson;
                                var extractedValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonContent);

                                if (extractedValues == null || extractedValues.Count == 0)
                                {
                                    logFile.MyLogFile($"Page {pageNumber}: Failed to extract values from GPT-4o vision response");
                                    pagesWithoutTitleBlock.Add(pageNumber);
                                    continue;
                                }

                                // Create TitleBlockDetailSection object from extracted values
                                var pageValues = new TitleBlockDetailSection
                                {
                                    ControlNumber = GetStringValue(extractedValues, "ControlNumber"),
                                    SectionNumber = GetStringValue(extractedValues, "SectionNumber"),
                                    JobNumber = GetStringValue(extractedValues, "JobNumber"),
                                    SheetNumber = GetStringValue(extractedValues, "SheetNumber"),
                                    District = GetStringValue(extractedValues, "District"),
                                    County = GetStringValue(extractedValues, "County"),
                                    HighwayNumber = GetStringValue(extractedValues, "HighwayNumber")
                                };

                                logFile.MyLogFile($"Page {pageNumber} extracted: CONT={pageValues.ControlNumber}, SECT={pageValues.SectionNumber}, " +
                                    $"JOB={pageValues.JobNumber}, SHEET={pageValues.SheetNumber}, DIST={pageValues.District}, " +
                                    $"COUNTY={pageValues.County}, HWY={pageValues.HighwayNumber}");

                                // Store sheet number for this page (for later use in SheetIndex validation)
                                if (!string.IsNullOrEmpty(pageValues.SheetNumber))
                                {
                                    sheetNumbersByPage[pageNumber] = pageValues.SheetNumber;
                                }

                                // Use first valid cover page (within first 5 pages) as expected values
                                if (expectedValues == null && pageNumber <= 5)
                                {
                                    expectedValues = new TitleBlockDetailSection
                                    {
                                        ControlNumber = pageValues.ControlNumber,
                                        SectionNumber = pageValues.SectionNumber,
                                        JobNumber = pageValues.JobNumber,
                                        District = pageValues.District,
                                        County = pageValues.County,
                                        HighwayNumber = pageValues.HighwayNumber
                                        // NOTE: SheetNumber is NOT included in expected values
                                    };

                                    logFile.MyLogFile($"Using page {pageNumber} as reference for expected values");

                                    // Initialize fieldsToCheck with expected values
                                    fieldsToCheck = new Dictionary<string, string>
                                    {
                                        { "Control Number", expectedValues.ControlNumber },
                                        { "Section Number", expectedValues.SectionNumber },
                                        { "Job Number", expectedValues.JobNumber },
                                        { "District", expectedValues.District },
                                        { "County", expectedValues.County },
                                        { "Highway Number", expectedValues.HighwayNumber }
                                    };

                                    // Initialize fieldResults
                                    foreach (var field in fieldsToCheck)
                                    {
                                        fieldResults[field.Key] = (0, 0, new List<int>());
                                    }
                                }

                                // Validate against expected values (if we have them)
                                if (expectedValues != null && fieldsToCheck.Any())
                                {
                                    foreach (var field in fieldsToCheck)
                                    {
                                        string fieldName = field.Key;
                                        string expectedValue = field.Value;

                                        var current = fieldResults[fieldName];
                                        current.PagesChecked++;

                                        // Get the actual value from pageValues
                                        string actualValue = null;
                                        switch (fieldName)
                                        {
                                            case "Control Number":
                                                actualValue = pageValues.ControlNumber;
                                                break;
                                            case "Section Number":
                                                actualValue = pageValues.SectionNumber;
                                                break;
                                            case "Job Number":
                                                actualValue = pageValues.JobNumber;
                                                break;
                                            case "District":
                                                actualValue = pageValues.District;
                                                break;
                                            case "County":
                                                actualValue = pageValues.County;
                                                break;
                                            case "Highway Number":
                                                actualValue = pageValues.HighwayNumber;
                                                break;
                                        }

                                        // Compare values
                                        if (!string.IsNullOrEmpty(actualValue) && !string.IsNullOrEmpty(expectedValue))
                                        {
                                            string normalizedActual = NormalizeValue(actualValue);
                                            string normalizedExpected = NormalizeValue(expectedValue);
                                            bool valueValid = normalizedActual.Contains(normalizedExpected) || normalizedExpected.Contains(normalizedActual);

                                            if (valueValid)
                                            {
                                                current.PagesWithValue++;
                                                logFile.MyLogFile($"Page {pageNumber}: {fieldName} = '{actualValue}' - VALID");
                                            }
                                            else
                                            {
                                                logFile.MyLogFile($"Page {pageNumber}: {fieldName} - MISMATCH (expected '{expectedValue}', got '{actualValue}')");
                                                current.MissingPages.Add(pageNumber);
                                            }
                                        }
                                        else
                                        {
                                            logFile.MyLogFile($"Page {pageNumber}: {fieldName} - MISSING");
                                            current.MissingPages.Add(pageNumber);
                                        }

                                        fieldResults[fieldName] = current;
                                    }
                                }
                            }
                            catch (JsonException ex)
                            {
                                logFile.MyLogFile($"Page {pageNumber}: Error parsing JSON response: {ex.Message}");
                                pagesWithoutTitleBlock.Add(pageNumber);
                            }
                        }

                        // Store sheet numbers dictionary for later use in SheetIndex validation
                        // Log sheet numbers extraction success
                        if (sheetNumbersByPage.Any())
                        {
                            logFile.MyLogFile($"Extracted sheet numbers for {sheetNumbersByPage.Count} pages");
                        }

                        // Build summary for each field and store as attributes
                        // Build summary for each field and store as attributes
                        task.Attributes = new List<PlanReviewAttribute>();
                        bool allFieldsValid = true;

                        // Add overall Title Block Detail Section summary FIRST
                        StringBuilder overallSummary = new StringBuilder();
                        overallSummary.AppendLine($"Title Block Detail Section found on {pagesWithTitleBlock}/{totalPages} pages.");

                        if (pagesWithoutTitleBlock.Any())
                        {
                            overallSummary.AppendLine($"Missing Title Block Detail Section on pages: {FormatPageNumbers(pagesWithoutTitleBlock)}");
                        }

                        overallSummary.AppendLine();

                        task.Attributes.Insert(0, new PlanReviewAttribute
                        {
                            AttributeName = "Title Block Detail Section Summary",
                            FileCategory = "Transportation",
                            Value = overallSummary.ToString().TrimEnd(),
                            AttributeCode = titleBlockAttrId
                        });

                        // Add individual field summaries
                        if (expectedValues != null && fieldsToCheck.Any())
                        {
                            foreach (var field in fieldsToCheck)
                            {
                                string fieldName = field.Key;
                                var result = fieldResults[fieldName];
                                string expectedValue = field.Value;

                                string summary = $"{fieldName} found on {result.PagesWithValue}/{result.PagesChecked} pages.";

                                if (!string.IsNullOrEmpty(expectedValue))
                                {
                                    summary += $"\nExpected value: {expectedValue}";
                                }

                                if (result.MissingPages.Any())
                                {
                                    summary += $"\nMissing/different {fieldName} on pages: {FormatPageNumbers(result.MissingPages)}";
                                    allFieldsValid = false;
                                }

                                logFile.MyLogFile($"{fieldName}: {summary}");

                                task.Attributes.Add(new PlanReviewAttribute
                                {
                                    AttributeName = fieldName,
                                    FileCategory = "Transportation",
                                    Value = summary,
                                    AttributeCode = titleBlockAttrId
                                });
                            }
                        }

                        // Set overall task status
                        if (allFieldsValid && pagesWithTitleBlock > 0)
                        {
                            task.Status = "Correct";
                            logFile.MyLogFile("All Title Block Detail Section fields are valid on sufficient pages");
                        }
                        else
                        {
                            task.Status = "Error";
                            logFile.MyLogFile("Some Title Block Detail Section fields are missing or invalid on multiple pages");
                        }

                    }
                    catch (Exception ex)
                    {
                        logFile.MyLogFile($"Error during Title Block Detail Section per-page validation: {ex.Message}");
                        task.Status = "Error";
                    }
                } //Transportation

                if (task.LogicType == "DateWithinYear")
                {
                    logFile.MyLogFile($"Handling special case for ActionDate in task '{task.TaskName}'");

                    // Get attributes for both CPC101 and Plat
                    var cpc101Attr = task.Attributes.FirstOrDefault(a => a.FileCategory == "CPC101");
                    var platAttr = titleBlock.ActionDate;

                    // Normalize values
                    if (cpc101Attr != null) cpc101Attr.Value = NormalizeValue(cpc101Attr.Value);
                    if (platAttr != null) platAttr = NormalizeValue(platAttr);

                    // Check if both values are present
                    if ((cpc101Attr == null || string.IsNullOrEmpty(cpc101Attr.Value)) ||
                        (platAttr == null || string.IsNullOrEmpty(platAttr)))
                    {
                        logFile.MyLogFile($"ActionDate values missing for one or both files.");
                        task.Status = "Error";
                        task.StatusDescription = "ActionDate missing in CPC101 or Plat";
                    }
                    else
                    {
                        // Try to parse dates
                        bool cpc101Valid = TryParseAndValidateDate(cpc101Attr.Value, out DateTime cpc101Date);
                        bool platValid = TryParseAndValidateDate(platAttr, out DateTime platDate);

                        if (!cpc101Valid || !platValid)
                        {
                            logFile.MyLogFile($"Invalid ActionDate format: CPC101={cpc101Attr.Value}, Plat={platAttr}");
                            task.Status = "Error";
                            task.StatusDescription = $"Invalid date format in CPC101 or Plat: CPC101={cpc101Attr.Value}, Plat={platAttr}";
                            continue;
                        }
                        else if (DateTime.Now.AddYears(-1) > cpc101Date || DateTime.Now.AddYears(-1) > platDate)
                        {
                            logFile.MyLogFile($"ActionDate is older than one year: CPC101={cpc101Date:MM/dd/yyyy}, Plat={platDate:MM/dd/yyyy}");
                            task.Status = "Error";
                            task.StatusDescription = $"ActionDate older than one year: CPC101={cpc101Date:MM/dd/yyyy}, Plat={platDate:MM/dd/yyyy}";
                            continue;
                        }
                        else if (DateTime.Now.AddYears(-1) <= cpc101Date || DateTime.Now.AddYears(-1) <= platDate)
                        {
                            logFile.MyLogFile($"ActionDate is within one year: CPC101={cpc101Date:MM/dd/yyyy}, Plat={platDate:MM/dd/yyyy}");
                            task.Status = "Correct";
                            task.StatusDescription = $"ActionDate within one year: CPC101={cpc101Date:MM/dd/yyyy}, Plat={platDate:MM/dd/yyyy}";
                            continue;
                        }
                        else
                        {
                            logFile.MyLogFile($"ActionDate error: CPC101={cpc101Date:MM/dd/yyyy}, Plat={platDate:MM/dd/yyyy}");
                            task.Status = "Error";
                            task.StatusDescription = $"Unknown error in ActionDate validation: CPC101={cpc101Date:MM/dd/yyyy}, Plat={platDate:MM/dd/yyyy}";
                            continue;
                        }
                    }
                } //LP-HTX

                if (task.LogicType == "SheetIndex" && planReviewRun.PlanReviewType == "Land")
                {
                    logFile.MyLogFile($"Handling SheetIndex");

                    // Get document information
                    var documentId = task.DocumentIds.FirstOrDefault();
                    logFile.MyLogFile($"Processing sheet index for document ID: {documentId}");

                    try
                    {
                        // Step 1: Extract the sheet index from the cover page (page 1)
                        var coverPageAttributes = await agileDDApi.ReadValuesAsync(documentId, 1);
                        if (!coverPageAttributes.Success)
                        {
                            logFile.MyLogFile($"Failed to get attributes for cover page: {coverPageAttributes.ErrorMessage}");
                            task.Status = "Error";
                            continue;
                        }

                        // Find all sheet index attributes
                        var sheetIndexAttributes = coverPageAttributes.Content
                            .Where(a => a.attribute_id == task.Attributes.FirstOrDefault(attr => attr.AttributeName == "SheetIndex")?.AttributeCode)
                            .Where(a => a.position != null && a.position.Count > 0)
                            .ToList();

                        if (!sheetIndexAttributes.Any())
                        {
                            logFile.MyLogFile("Sheet index not found on cover page.");
                            task.Status = "Notice";
                            continue;
                        }

                        // Extract and append all sheet index contents
                        var sheetIndexTexts = new List<string>();
                        foreach (var attr in sheetIndexAttributes)
                        {
                            string text = await agileDDApi.GetTextFromBoxAsync(documentId, 1, attr.position);
                            if (!string.IsNullOrWhiteSpace(text))
                                sheetIndexTexts.Add(text);
                        }
                        string sheetIndexText = string.Join(Environment.NewLine, sheetIndexTexts);

                        logFile.MyLogFile($"Sheet index text from cover page: {sheetIndexText}");

                        // Step 2: Parse the sheet index into a dictionary of sheet numbers and titles
                        var sheetIndex = new Dictionary<string, string>();

                        //add a call to azure openai to parse the sheet index text into a dictionary
                        AzureOpenAIApi azureOpenAIApi = new AzureOpenAIApi(
                            logFile,
                            "You are an automation assistant that extracts structured data from unstructured or semi-structured text. When given a table or list, parse each row into a key-value pair, using the first column as the key and the second column as the value. Output the result as a valid JSON dictionary. If the text is noisy or contains extra information, focus only on the relevant table rows.",
                            Convert.ToDecimal(0.2)
                        );
                        sheetIndex = await azureOpenAIApi.ParseSheetIndexAsync(sheetIndexText);
                        if (sheetIndex.Count == 0)
                        {
                            logFile.MyLogFile("Failed to parse any sheet numbers and titles from the index.");
                            task.Status = "Error";
                            continue;
                        }

                        // Step 3: Check each page for its sheet number and title
                        // Get total page count
                        var docAttributes = await agileDDApi.ReadValuesAsync(documentId);
                        if (!docAttributes.Success)
                        {
                            logFile.MyLogFile($"Failed to get document attributes: {docAttributes.ErrorMessage}");
                            task.Status = "Error";
                            continue;
                        }

                        int maxPage = docAttributes.Content.Any() ? docAttributes.Content.Max(a => a.page_number) : 1;
                        logFile.MyLogFile($"Document has {maxPage} pages. Checking each page for sheet number and title...");

                        int matchedNumberPages = 0;
                        int matchedTitlePages = 0;
                        int matchedBothPages = 0;
                        int totalPages = 0;

                        // Validate sheets starting from page 1
                        for (int pageNumber = 1; pageNumber <= maxPage; pageNumber++)
                        {
                            var pageAttributes = await agileDDApi.ReadValuesAsync(documentId, pageNumber);
                            if (!pageAttributes.Success)
                            {
                                logFile.MyLogFile($"Failed to get attributes for page {pageNumber}: {pageAttributes.ErrorMessage}");
                                continue;
                            }

                            totalPages++;

                            // Find sheet number attribute and extract value
                            var sheetNumberAttr = pageAttributes.Content.FirstOrDefault(a =>
                                a.attribute_id == task.Attributes.FirstOrDefault(attr => attr.AttributeName == "SheetNumber")?.AttributeCode);
                            string pageSheetNumber = "N/A";
                            if (sheetNumberAttr != null)
                            {
                                if (sheetNumberAttr.position != null && sheetNumberAttr.position.Any())
                                {
                                    pageSheetNumber = await agileDDApi.GetTextFromBoxAsync(documentId, pageNumber, sheetNumberAttr.position);
                                }
                                else
                                {
                                    pageSheetNumber = sheetNumberAttr.data ?? "N/A";
                                }
                            }
                            pageSheetNumber = NormalizeValue(pageSheetNumber);

                            // Find sheet title attribute and extract value
                            var sheetTitleAttr = pageAttributes.Content.FirstOrDefault(a =>
                                a.attribute_id == task.Attributes.FirstOrDefault(attr => attr.AttributeName == "SheetTitle")?.AttributeCode);
                            string pageSheetTitle = "N/A";
                            if (sheetTitleAttr != null)
                            {
                                if (sheetTitleAttr.position != null && sheetTitleAttr.position.Any())
                                {
                                    pageSheetTitle = await agileDDApi.GetTextFromBoxAsync(documentId, pageNumber, sheetTitleAttr.position);
                                }
                                else
                                {
                                    pageSheetTitle = sheetTitleAttr.data ?? "N/A";
                                }
                            }
                            pageSheetTitle = NormalizeValue(pageSheetTitle);

                            logFile.MyLogFile($"Page {pageNumber} - Sheet Number: {pageSheetNumber}, Title: {pageSheetTitle}");

                            // Verify sheet number and title separately
                            bool numberMatches = sheetIndex.ContainsKey(pageSheetNumber);
                            bool titleMatches = false;
                            string expectedTitle = numberMatches ? sheetIndex[pageSheetNumber] : null;

                            if (numberMatches)
                            {
                                logFile.MyLogFile($"Page {pageNumber} sheet number {pageSheetNumber} found in index.");
                                // Title matching is fuzzy since formatting and spacing can differ
                                titleMatches = NormalizeValue(expectedTitle).Contains(pageSheetTitle) ||
                                               pageSheetTitle.Contains(NormalizeValue(expectedTitle));

                                if (titleMatches)
                                {
                                    logFile.MyLogFile($"Page {pageNumber} title matches index entry: {pageSheetNumber} - {expectedTitle}");
                                    matchedTitlePages++;
                                    matchedBothPages++;
                                }
                                else
                                {
                                    logFile.MyLogFile($"Page {pageNumber} sheet number matches but title doesn't match. Expected: {expectedTitle}, Found: {pageSheetTitle}");
                                }
                            }
                            else
                            {
                                logFile.MyLogFile($"Page {pageNumber} sheet number {pageSheetNumber} not found in index.");
                            }
                        }

                        // Calculate percentages
                        double numberMatchPct = totalPages > 0 ? (double)matchedNumberPages / totalPages * 100 : 0;
                        double titleMatchPct = totalPages > 0 ? (double)matchedTitlePages / totalPages * 100 : 0;
                        double bothMatchPct = totalPages > 0 ? (double)matchedBothPages / totalPages * 100 : 0;

                        // Decide status
                        if (totalPages > 0)
                        {
                            logFile.MyLogFile($"Sheet number match: {matchedNumberPages}/{totalPages} ({numberMatchPct:F1}%)");
                            logFile.MyLogFile($"Sheet title match: {matchedTitlePages}/{totalPages} ({titleMatchPct:F1}%)");
                            logFile.MyLogFile($"Both match: {matchedBothPages}/{totalPages} ({bothMatchPct:F1}%)");

                            if (bothMatchPct >= 90)
                                task.Status = "Correct";
                            else if (bothMatchPct >= 50)
                                task.Status = "Partial Match";
                            else
                                task.Status = "Mismatch";

                            var detailsAttribute = task.Attributes.FirstOrDefault();
                            if (detailsAttribute != null)
                            {
                                detailsAttribute.Value = $"Sheet number match: {matchedNumberPages}/{totalPages} ({numberMatchPct:F1}%), " +
                                                        $"Sheet title match: {matchedTitlePages}/{totalPages} ({titleMatchPct:F1}%), " +
                                                        $"Both match: {matchedBothPages}/{totalPages} ({bothMatchPct:F1}%)";
                            }
                        }
                        else
                        {
                            task.Status = "Error";
                            logFile.MyLogFile("No pages processed during sheet index validation.");
                        }
                    }
                    catch (Exception ex)
                    {
                        logFile.MyLogFile($"Error during sheet index validation: {ex.Message}");
                        task.Status = "Error";
                    }
                } //LD-HTX

                if (task.LogicType == "SheetIndex" && planReviewRun.PlanReviewType == "Transportation")
                {
                    logFile.MyLogFile($"Handling SheetIndex");

                    // Get document information
                    var documentId = task.DocumentIds.FirstOrDefault();
                    logFile.MyLogFile($"Processing sheet index for document ID: {documentId}");

                    try
                    {
                        // Step 1: Get the attribute code for Sheet Index
                        int? sheetIndexAttrCode = task.Attributes.FirstOrDefault(attr => attr.AttributeName == "Sheet Index")?.AttributeCode;
                        if (sheetIndexAttrCode == null)
                        {
                            logFile.MyLogFile("Sheet Index attribute code not found in task attributes.");
                            task.Status = "Error";
                            continue;
                        }

                        // Step 2: Find ALL Sheet Index attributes across the document to determine which pages have them
                        var allDocAttributes = await agileDDApi.ReadValuesAsync(documentId);
                        if (!allDocAttributes.Success)
                        {
                            logFile.MyLogFile($"Failed to get document attributes: {allDocAttributes.ErrorMessage}");
                            task.Status = "Error";
                            continue;
                        }

                        // Find all pages that contain Sheet Index attributes
                        var sheetIndexAttributes = allDocAttributes.Content
                            .Where(a => a.attribute_id == sheetIndexAttrCode.Value)
                            .Where(a => a.position != null && a.position.Count > 0)
                            .ToList();

                        if (!sheetIndexAttributes.Any())
                        {
                            logFile.MyLogFile("Sheet index not found in the document.");
                            task.Status = "Notice";
                            continue;
                        }

                        // Get unique page numbers that have Sheet Index attributes
                        var sheetIndexPageNumbers = sheetIndexAttributes
                            .Select(a => a.page_number)
                            .Distinct()
                            .OrderBy(p => p)
                            .ToList();

                        logFile.MyLogFile($"Sheet Index found on {sheetIndexPageNumbers.Count} page(s): {string.Join(", ", sheetIndexPageNumbers)}");

                        // Step 3: Extract sheet index text using GPT-4o vision from cropped images
                        var allSheetIndexTexts = new List<string>();

                        // Initialize Azure OpenAI API with GPT-4o vision model
                        var azureVisionApi = new AzureOpenAIApi(
                            logFile,
                            "You are an expert at extracting structured data from engineering drawing sheet indexes.",
                            Convert.ToDecimal(0.2),
                            useVision: true  // Enable GPT-4o vision model
                        );

                        foreach (var sheetIndexAttr in sheetIndexAttributes)
                        {
                            int pageNum = sheetIndexAttr.page_number;
                            var pos = sheetIndexAttr.position[0];

                            logFile.MyLogFile($"Cropping Sheet Index from page {pageNum} at position: L={pos.left:F2}, T={pos.top:F2}, W={pos.width:F2}, H={pos.height:F2}");

                            // Get cropped image of the sheet index
                            var croppedImageBytes = await agileDDApi.GetCroppedImageAsync(
                                documentId,
                                pageNum,
                                pos.left,
                                pos.top,
                                pos.width,
                                pos.height
                            );

                            if (croppedImageBytes == null || croppedImageBytes.Length == 0)
                            {
                                logFile.MyLogFile($"Failed to crop Sheet Index image from page {pageNum}");
                                continue;
                            }

                            logFile.MyLogFile($"Successfully cropped Sheet Index image from page {pageNum}, size: {croppedImageBytes.Length} bytes");

                            // Create prompt for GPT-4o vision to extract sheet index data

                            string visionPrompt = @"Extract the sheet index table from this image. Return ONLY a valid JSON dictionary mapping sheet numbers to sheet titles.
                                
                                CRITICAL: Return ONLY raw JSON. Do NOT wrap in markdown code blocks. Do NOT include any explanatory text.

                                Rules:
                                1. Sheet numbers can be numeric (1, 2, 10) or alphanumeric (3A, 3B, 12C)
                                2. If a range is shown (e.g., '3A - 3C' or '6 - 9'), expand it to individual entries.
                                3. CRITICAL - Only increment the title designator when the label explicitly contains 'thru' or sequential designators:
                                   - '9-20  BC(1)-14 thru BC(12)-14' → increment because label says 'BC(1)-14 THRU BC(12)-14'
                                   - '21-22  TCP(2-1)-18 thru TCP(2-2)-18' → increment because label says 'TCP(2-1)-18 THRU TCP(2-2)-18'
                                   - '195-197  EC(9)-16' → do NOT increment, same title for all: '195': 'EC(9)-16', '196': 'EC(9)-16', '197': 'EC(9)-16'
                                   - '3A-3C  GENERAL NOTES' → do NOT increment, same title for all
                                4. The word 'thru' or 'through' in the label is the ONLY signal to increment. Without it, repeat the title unchanged.
                                5. Ignore section headers (GENERAL, ROADWAY, DRAINAGE, etc.)
                                6. Return ONLY the JSON object, no additional text or explanation

                                Example input row:  3A-3C   GENERAL NOTES
                                Example output:
                                {
                                  ""3A"": ""GENERAL NOTES"",
                                  ""3B"": ""GENERAL NOTES"",
                                  ""3C"": ""GENERAL NOTES""
                                }

                                Example input row:  9-20   BC(1)-14 thru BC(12)-14
                                Example output:
                                {
                                  ""9"":  ""BC(1)-14"",
                                  ""10"": ""BC(2)-14"",
                                  ""11"": ""BC(3)-14"",
                                  ""12"": ""BC(4)-14"",
                                  ""13"": ""BC(5)-14"",
                                  ""14"": ""BC(6)-14"",
                                  ""15"": ""BC(7)-14"",
                                  ""16"": ""BC(8)-14"",
                                  ""17"": ""BC(9)-14"",
                                  ""18"": ""BC(10)-14"",
                                  ""19"": ""BC(11)-14"",
                                  ""20"": ""BC(12)-14""
                                }

                                Example input row:  21-22   TCP(2-1)-18 thru TCP(2-2)-18
                                Example output:
                                {
                                  ""21"": ""TCP(2-1)-18"",
                                  ""22"": ""TCP(2-2)-18""
                                }
                                
                                Example input row: 195 - 197   EC(9) - 16
                                Example output:
                                {
                                    ""195"": ""EC(9) - 16"",
                                    ""196"": ""EC(9) - 16"",
                                    ""197"": ""EC(9) - 16""
                                }

                                Example input row:  195-197   EC(9)-16                  
                                Example output format:
                                {
                                  ""1"": ""COVER SHEET"",
                                  ""2"": ""LOCATION MAP"",
                                  ""3A"": ""GENERAL NOTES"",
                                  ""3B"": ""GENERAL NOTES"",
                                  ""3C"": ""GENERAL NOTES"",
                                  ""6"": ""TYPICAL SECTIONS"",
                                  ""7"": ""TYPICAL SECTIONS""
                                }";

                            // Query GPT-4o vision model with the cropped image
                            var visionResult = await azureVisionApi.QueryImageAsync(croppedImageBytes, visionPrompt);

                            if (!visionResult.Success)
                            {
                                logFile.MyLogFile($"Error querying GPT-4o vision for page {pageNum}: {visionResult.ErrorMessage}");
                                continue;
                            }

                            logFile.MyLogFile($"GPT-4o vision extraction successful for page {pageNum}");
                            logFile.MyLogFile($"[GPT-4o Vision Output] Page {pageNum}: {visionResult.Content}");
                            allSheetIndexTexts.Add(visionResult.Content);
                        }

                        // Combine all extracted texts
                        string sheetIndexText = string.Join("\n", allSheetIndexTexts);

                        if (string.IsNullOrWhiteSpace(sheetIndexText))
                        {
                            logFile.MyLogFile("ERROR: GPT-4o vision extraction failed for all pages. Cannot proceed with sheet index parsing.");
                            task.Status = "Error";
                            continue;
                        }

                        logFile.MyLogFile($"Combined sheet index text from GPT-4o vision ({sheetIndexText.Length} characters)");

                        // Step 4: Parse the combined JSON responses into a single dictionary
                        var sheetIndex = new Dictionary<string, string>();

                        foreach (var jsonText in allSheetIndexTexts)
                        {
                            try
                            {
                                // Extract JSON from response (in case there's extra text)
                                string cleanedJson = Regex.Replace(jsonText, @"```(?:json)?\s*", "").Trim();
                                string jsonContent = ExtractFirstJsonObject(cleanedJson) ?? cleanedJson;
                                var extractedDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonContent);

                                if (extractedDict != null)
                                {
                                    foreach (var kvp in extractedDict)
                                    {
                                        if (!sheetIndex.ContainsKey(kvp.Key))
                                        {
                                            sheetIndex[kvp.Key] = kvp.Value;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                logFile.MyLogFile($"Error parsing JSON from GPT-4o vision response: {ex.Message}");
                            }
                        }

                        if (sheetIndex.Count == 0)
                        {
                            logFile.MyLogFile("Failed to parse any sheet numbers and titles from GPT-4o vision responses.");
                            task.Status = "Error";
                            continue;
                        }

                        logFile.MyLogFile($"Parsed sheet index with {sheetIndex.Count} entries from GPT-4o vision");
                        logFile.MyLogFile($"Sheet Index Entries: {string.Join(", ", sheetIndex.Select(kvp => $"{kvp.Key}: {kvp.Value}"))}");

                        // Build a dictionary of actual sheet numbers and titles from each page
                        // Get total page count
                        var docAttributes = await agileDDApi.ReadValuesAsync(documentId);
                        if (!docAttributes.Success)
                        {
                            logFile.MyLogFile($"Failed to get document attributes: {docAttributes.ErrorMessage}");
                            task.Status = "Error";
                            continue;
                        }

                        int maxPage = docAttributes.Content.Any() ? docAttributes.Content.Max(a => a.page_number) : 1;
                        logFile.MyLogFile($"Document has {maxPage} pages. Extracting sheet numbers and titles from each page...");

                        // Get attribute codes
                        int? sheetTitleAttrCode = task.Attributes.FirstOrDefault(attr => attr.AttributeName == "Sheet Title")?.AttributeCode;

                        if (sheetTitleAttrCode == null)
                        {
                            logFile.MyLogFile("Sheet Title attribute code not found in task attributes.");
                            task.Status = "Error";
                            continue;
                        }

                        // Get the sheet number from TitleBlockDetailSection per page
                        // Lists to track pages with mismatches
                        List<int> pagesWithMismatch = new List<int>();

                        // Track which sheet index entries were matched
                        HashSet<string> matchedSheetNumbers = new HashSet<string>();

                        int matchedBothPages = 0;
                        int totalPages = 0;

                        // Extract sheet number and title for each page
                        for (int pageNumber = 1; pageNumber <= maxPage; pageNumber++)
                        {
                            totalPages++;
                            string pageSheetNumber = "N/A";
                            string pageSheetTitle = "N/A";

                            // Get sheet number from the local dictionary (populated by TitleBlockDetailSection task)
                            if (sheetNumbersByPage != null && sheetNumbersByPage.ContainsKey(pageNumber))
                            {
                                pageSheetNumber = sheetNumbersByPage[pageNumber];
                                logFile.MyLogFile($"Page {pageNumber}: Retrieved sheet number '{pageSheetNumber}' from dictionary");
                            }
                            else
                            {
                                logFile.MyLogFile($"Page {pageNumber}: No sheet number found in dictionary");
                            }

                            // Get sheet title from Sheet Title attribute
                            var sheetTitlePageAttrs = await agileDDApi.ReadValuesAsync(documentId, pageNumber, sheetTitleAttrCode.Value);
                            if (sheetTitlePageAttrs.Success && sheetTitlePageAttrs.Content != null && sheetTitlePageAttrs.Content.Any())
                            {
                                var sheetTitleAttr = sheetTitlePageAttrs.Content.FirstOrDefault(a => a.attribute_id == sheetTitleAttrCode.Value);
                                if (sheetTitleAttr != null)
                                {
                                    if (sheetTitleAttr.position != null && sheetTitleAttr.position.Any())
                                    {
                                        pageSheetTitle = await agileDDApi.GetTextFromBoxAsync(documentId, pageNumber, sheetTitleAttr.position);
                                    }
                                    else
                                    {
                                        pageSheetTitle = sheetTitleAttr.data ?? "N/A";
                                    }
                                    pageSheetTitle = NormalizeValue(pageSheetTitle);
                                }
                            }

                            logFile.MyLogFile($"Page {pageNumber}: Extracted Sheet Number='{pageSheetNumber}', Sheet Title='{pageSheetTitle}'");

                            // Check if sheet number exists in index
                            bool numberMatches = sheetIndex.ContainsKey(pageSheetNumber);
                            bool titleMatches = false;
                            string expectedTitle = numberMatches ? sheetIndex[pageSheetNumber] : null;

                            if (numberMatches)
                            {
                                logFile.MyLogFile($"✓ Page {pageNumber}: Sheet number '{pageSheetNumber}' found in index");

                                // Check if title matches (fuzzy matching)
                                string normalizedExpectedTitle = NormalizeValue(expectedTitle);
                                logFile.MyLogFile($"Comparing page title '{pageSheetTitle}' to expected title '{normalizedExpectedTitle}'");

                                // Before the comparison, create space-stripped versions for contains checks
                                string compactPageTitle = Regex.Replace(pageSheetTitle, @"\s+", "");
                                string compactExpectedTitle = Regex.Replace(normalizedExpectedTitle, @"\s+", "");

                                titleMatches = normalizedExpectedTitle.Contains(pageSheetTitle) ||
                                               pageSheetTitle.Contains(normalizedExpectedTitle) ||
                                               compactPageTitle.Contains(compactExpectedTitle) ||       // ← catches "tcp(2-2)-18" vs "tcp (2-2) -18"
                                               compactExpectedTitle.Contains(compactPageTitle) ||
                                               FuzzyMatch(pageSheetTitle, normalizedExpectedTitle, maxDistance: 15);

                                if (titleMatches)
                                {
                                    matchedBothPages++;
                                    logFile.MyLogFile($"  ✓ Title also matches: '{pageSheetTitle}' ≈ '{expectedTitle}'");

                                    // Track that this sheet number was successfully matched
                                    matchedSheetNumbers.Add(pageSheetNumber);
                                }
                                else
                                {
                                    logFile.MyLogFile($"  ✗ Sheet number/sheet title mismatch on page {pageNumber}");
                                    pagesWithMismatch.Add(pageNumber);

                                    // Still add to matched numbers since the number exists (only title mismatched)
                                    matchedSheetNumbers.Add(pageSheetNumber);
                                }
                            }
                            else
                            {
                                logFile.MyLogFile($"✗ Sheet number/sheet title mismatch on page {pageNumber}");
                                pagesWithMismatch.Add(pageNumber);
                            }
                        }

                        // Calculate unused sheet titles from sheet index
                        var unusedSheetNumbers = sheetIndex.Keys.Except(matchedSheetNumbers).OrderBy(n => n).ToList();

                        if (unusedSheetNumbers.Any())
                        {
                            logFile.MyLogFile($"⚠ WARNING: {unusedSheetNumbers.Count} sheet(s) in sheet index were NOT found in plan set:");
                            foreach (var unusedNumber in unusedSheetNumbers)
                            {
                                string missingSheetTitle = sheetIndex[unusedNumber];
                                logFile.MyLogFile($"  - Missing sheet: {missingSheetTitle} (Sheet #{unusedNumber})");
                            }
                        }
                        else
                        {
                            logFile.MyLogFile($"✓ All {sheetIndex.Count} sheets in sheet index were found in plan set");
                        }

                        // Calculate percentage
                        double bothMatchPct = totalPages > 0 ? (double)matchedBothPages / totalPages * 100 : 0;

                        // Build summary output
                        task.Attributes = new List<PlanReviewAttribute>();
                        bool allValid = pagesWithMismatch.Count == 0;

                        // Overall summary
                        StringBuilder overallSummary = new StringBuilder();
                        overallSummary.AppendLine($"Sheet Index validation completed for {totalPages} pages.");
                        overallSummary.AppendLine($"Both match: {matchedBothPages}/{totalPages} ({bothMatchPct:F1}%)");

                        if (pagesWithMismatch.Any())
                        {
                            overallSummary.AppendLine($"Sheet number/sheet title mismatch on page {FormatPageNumbers(pagesWithMismatch)}");
                        }

                        // NEW: Add unused sheets information to summary
                        if (unusedSheetNumbers.Any())
                        {
                            overallSummary.AppendLine($"\n({unusedSheetNumbers.Count} sheet(s) not found in plan set):");
                            foreach (var unusedNumber in unusedSheetNumbers)
                            {
                                string missingSheetTitle = sheetIndex[unusedNumber];
                                overallSummary.AppendLine($"  - Missing sheet: {missingSheetTitle} (Sheet #{unusedNumber})");
                            }
                        }
                        else
                        {
                            overallSummary.AppendLine($"\nAll {sheetIndex.Count} sheets listed in Index of Sheets found in plan set");
                        }

                        task.Attributes.Add(new PlanReviewAttribute
                        {
                            AttributeName = "Sheet Index Summary",
                            FileCategory = "Transportation",
                            Value = overallSummary.ToString().TrimEnd()
                        });

                        // Set overall task status
                        if (totalPages > 0)
                        {
                            logFile.MyLogFile($"=== SUMMARY ===");
                            logFile.MyLogFile($"Both match: {matchedBothPages}/{totalPages} ({bothMatchPct:F1}%)");
                            logFile.MyLogFile($"Unused sheets: {unusedSheetNumbers.Count}");

                            if (bothMatchPct >= 90 && allValid && !unusedSheetNumbers.Any())
                                task.Status = "Correct";
                            else if (bothMatchPct >= 50 || unusedSheetNumbers.Any())
                                task.Status = "Partial Match";
                            else
                                task.Status = "Mismatch";
                        }
                        else
                        {
                            task.Status = "Error";
                            logFile.MyLogFile("No pages processed during sheet index validation.");
                        }
                    }
                    catch (Exception ex)
                    {
                        logFile.MyLogFile($"Error during sheet index validation: {ex.Message}");
                        task.Status = "Error";
                    }
                } //Transportation

                if (task.LogicType == "CheckValuePerDocument")
                {
                    logFile.MyLogFile($"Handling document-level value validation for task '{task.TaskName}'");

                    bool isValueValid = false;
                    string expectedPattern = "";
                    string logicDescription = "";

                    switch (task.TaskName)
                    {
                        case "Location": //Plat
                            // Location just needs to have a meaningful value
                            isValueValid = !string.IsNullOrWhiteSpace(titleBlock.Location) && titleBlock.Location.Length > 3;
                            expectedPattern = "meaningful value";
                            logicDescription = "Location should be a meaningful value.";
                            task.Attributes.FirstOrDefault().Value = titleBlock.Location;
                            break;

                        case "Reserves": //Plat
                            // Reserves should be a whole number greater than or equal to 0
                            isValueValid = titleBlock.Reserves.HasValue && titleBlock.Reserves >= 0;
                            expectedPattern = "whole number >= 0";
                            logicDescription = "Reserves should be a whole number greater than or equal to 0.";
                            task.Attributes.FirstOrDefault().Value = titleBlock.Reserves.HasValue ? titleBlock.Reserves.Value.ToString() : "N/A";
                            break;

                        case "Block": //Plat
                            // Blocks should be a whole number greater than or equal to 0
                            isValueValid = titleBlock.Blocks.HasValue && titleBlock.Blocks >= 0;
                            expectedPattern = "whole number >= 0";
                            logicDescription = "Blocks should be a whole number greater than or equal to 0.";
                            task.Attributes.FirstOrDefault().Value = titleBlock.Blocks.HasValue ? titleBlock.Blocks.Value.ToString() : "N/A";
                            break;

                        case "Owner": //Plat
                            // Owner just needs to have a meaningful value
                            isValueValid = !string.IsNullOrWhiteSpace(titleBlock.Owner) && titleBlock.Owner.Length > 3;
                            expectedPattern = "meaningful value";
                            logicDescription = "Owner should be a meaningful value.";
                            task.Attributes.FirstOrDefault().Value = titleBlock.Owner;
                            break;

                        case "Scale": //Plat
                            // Scale should allow any number, with optional escaping for quotes/apostrophes
                            expectedPattern = @"1\s*\\?[""]\\?\s*=\s*\\?\d+\\?\s*[']";
                            isValueValid = System.Text.RegularExpressions.Regex.IsMatch(
                                titleBlock.Scale, expectedPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            logicDescription = "Scale should match the format 1\"=###' with optional escaping for quotes/apostrophes.";
                            task.Attributes.FirstOrDefault().Value = titleBlock.Scale;
                            break;

                        case "Abstract Number": //Plat
                                                // Abstract Number should match the format A-000 where zeros can be any integer
                                                // OR contain the keyword "abstract" for more flexible matching (e.g., "Abstract 832", "abstract a-323")

                            if (string.IsNullOrWhiteSpace(titleBlock.AbstractNumber))
                            {
                                isValueValid = false;
                                expectedPattern = "non-empty value containing 'Abstract' or matching A-### format";
                                logicDescription = "Abstract Number should not be empty and should either contain the keyword 'abstract' or match the format A-###.";
                                task.Attributes.FirstOrDefault().Value = titleBlock.AbstractNumber;
                            }
                            else
                            {
                                string normalizedAbstract = NormalizeValue(titleBlock.AbstractNumber);

                                // Check if it contains the keyword "abstract" (fuzzy match for OCR errors)
                                bool hasAbstractKeyword = normalizedAbstract.Contains("abstract");

                                // Check if it matches the strict A-### pattern
                                bool matchesPattern = System.Text.RegularExpressions.Regex.IsMatch(
                                    titleBlock.AbstractNumber, @"A-\d{3}",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                                isValueValid = hasAbstractKeyword || matchesPattern;
                                expectedPattern = "contain 'abstract' keyword or match format A-### (e.g., A-323)";
                                logicDescription = "Abstract Number should either contain the keyword 'abstract' or match the format A-###.";
                                task.Attributes.FirstOrDefault().Value = titleBlock.AbstractNumber;
                            }
                            break;

                        case "COH Park Sector": //CPC101
                            // COH Park Sector should only be validated if in the City of Houston (not ETJ)
                            string cohJurisdiction = titleBlock?.CityETJ ?? "City";
                            bool isCity = NormalizeValue(cohJurisdiction).Contains("city");

                            if (isCity)
                            {
                                // City jurisdiction: COH Park Sector should be a whole number greater than or equal to 0
                                isValueValid = int.TryParse(attr1.Value, out int parksectorValue) && parksectorValue >= 0;
                                expectedPattern = "whole number >= 0";
                                logFile.MyLogFile($"'{task.TaskName}' validation: Value='{attr1.Value}', Valid={isValueValid}");
                                logicDescription = "COH Park Sector should be a whole number greater than or equal to 0 for City of Houston jurisdiction.";
                            }
                            else
                            {
                                // ETJ jurisdiction: COH Park Sector should be empty or N/A
                                isValueValid = string.IsNullOrEmpty(attr1.Value) || attr1.Value == "n/a";
                                expectedPattern = "empty or N/A (not applicable for ETJ)";
                                logFile.MyLogFile($"'{task.TaskName}' validation skipped - not in City jurisdiction. CityETJ value: {cohJurisdiction}. Value should be empty but is: {attr1.Value}");
                                logicDescription = "COH Park Sector is not applicable for ETJ jurisdiction and should be empty or N/A.";
                            }
                            break;

                        case "Commission Action": //CPC101
                            // Commission Action should contain keywords related to commission approval
                            if (attr1 != null && !string.IsNullOrEmpty(attr1.Value) && attr1.Value != "N/A")
                            {
                                // Keywords to look for in Commission Action text
                                string[] commissionKeywords = new string[] {
                                    "commission action", "planning commission", "approved", "approve", "unanimously",
                                    "motion", "carried", "vote", "granted"
                                };

                                // Normalize the text to lower case for case-insensitive comparison
                                string normalizedText = NormalizeValue(attr1.Value);
                                isValueValid = commissionKeywords.Any(keyword => normalizedText.Contains(keyword));
                                expectedPattern = "contains commission-related keywords";
                                logicDescription = "Commission Action should contain keywords related to commission approval.";
                            }
                            else
                            {
                                isValueValid = false;
                                expectedPattern = "CPC101 attribute missing or N/A";
                                logicDescription = "Commission Action attribute is missing or marked as N/A.";
                            }
                            break;

                        case "ILMS Number": //LD
                            // ILMS Number should be an 8-digit number
                            expectedPattern = @"^\d{8}$";
                            // Remove any non-numeric characters before validation
                            string ilmsValue = System.Text.RegularExpressions.Regex.Replace(attr1.Value, @"[^\d]", "");
                            isValueValid = ilmsValue.Length == 8 &&
                                           System.Text.RegularExpressions.Regex.IsMatch(
                                               ilmsValue, expectedPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                            // Log the validation result with details
                            logFile.MyLogFile($"ILMS Number validation: Original value='{attr1.Value}', " +
                                            $"Cleaned value='{ilmsValue}', Valid={isValueValid}");
                            expectedPattern = "8-digit number";
                            logicDescription = "ILMS Number should be an 8-digit number.";
                            break;

                        case "One Call Note": //LD
                            // One Call Note should contain keywords related to One Call Note Notice
                            if (attr1 != null && !string.IsNullOrEmpty(attr1.Value) && attr1.Value != "N/A")
                            {
                                // Keywords to look for in Commission Action text
                                string[] oneCallKeywords = new string[] {
                                    "one-call", "call before you dig"
                                };

                                // Normalize the text to lower case for case-insensitive comparison
                                string normalizedText = NormalizeValue(attr1.Value);
                                isValueValid = oneCallKeywords.Any(keyword => normalizedText.Contains(keyword));
                                expectedPattern = "contains one-call note-related keywords";
                                logicDescription = "One Call Note should contain keywords related to One Call Note Notice.";
                            }
                            else
                            {
                                isValueValid = false;
                                expectedPattern = "attribute missing or N/A";
                                logicDescription = "One Call Note attribute is missing or marked as N/A.";
                            }
                            break;

                        case "Private Utility Notes": //LD
                            // Private Utility Notes should contain keywords related to Private Utility Notes
                            if (attr1 != null && !string.IsNullOrEmpty(attr1.Value) && attr1.Value != "N/A")
                            {
                                // Keywords to look for in Commission Action text
                                string[] commissionKeywords = new string[] {
                                    "underground utilities", "buried cables", "underground lines", "private utility"
                                };

                                // Normalize the text to lower case for case-insensitive comparison
                                string normalizedText = NormalizeValue(attr1.Value);
                                isValueValid = commissionKeywords.Any(keyword => normalizedText.Contains(keyword));
                                expectedPattern = "contains private utility-related keywords";
                                logicDescription = "Private Utility Notes should contain keywords related to Private Utility Notes.";
                            }
                            else
                            {
                                isValueValid = false;
                                expectedPattern = "attribute missing or N/A";
                                logicDescription = "Private Utility Notes attribute is missing or marked as N/A.";
                            }
                            break;

                        case "Pipe Construction Note": //LD
                            // Pipe Construction Note should contain keywords related to Pipe Construction Notes
                            if (attr1 != null && !string.IsNullOrEmpty(attr1.Value) && attr1.Value != "N/A")
                            {
                                // Keywords to look for in Commission Action text
                                string[] commissionKeywords = new string[] {
                                    "water line", "minimum pressure", "pipe construction"
                                };

                                // Normalize the text to lower case for case-insensitive comparison
                                string normalizedText = NormalizeValue(attr1.Value);
                                isValueValid = commissionKeywords.Any(keyword => normalizedText.Contains(keyword));
                                expectedPattern = "contains one-call note-related keywords";
                                logicDescription = "Pipe Construction Note should contain keywords related to Pipe Construction Notes.";
                            }
                            else
                            {
                                isValueValid = false;
                                expectedPattern = "attribute missing or N/A";
                                logicDescription = "Pipe Construction Note attribute is missing or marked as N/A.";
                            }
                            break;

                        case "Location Map":
                            // For Location Map, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Location Map should be detected on the cover sheet.";
                            break; //Transportation
                        case "Approval Notes":
                            // For Approval Notes, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Approval Notes should be detected on the index sheet.";
                            break; //Transportation
                        case "Design Speed ADT":
                            // For design speed ADT, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Design Speed ADT should be detected on the cover sheet.";
                            break; //Transportation
                        case "Copyright":
                            // For Copyright, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Copyright should be detected on the cover sheet.";
                            break; //Transportation
                        case "Accessibility Notes":
                            // For Accessibility Notes, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Accessibility Notes should be detected on the cover sheet.";
                            break; //Transportation
                        case "Equations":
                            // For Equations, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Equations should be detected on the cover sheet.";
                            break; //Transportation
                        case "Exceptions":
                            // For Exceptions, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Exceptions should be detected on the cover sheet.";
                            break; //Transportation
                        case "RRCrossing":
                            // For RR Crossing, just check if the attribute was detected (i.e., attr1.Value is not null or empty)
                            isValueValid = !string.IsNullOrWhiteSpace(attr1.Value) && attr1.Value != "n/a";
                            expectedPattern = "meaningful value";
                            logicDescription = "Railroad Crossing should be detected on the cover sheet.";
                            break; //Transportation

                        default:
                            // For any other fields, just ensure there's meaningful content (not just a few characters)
                            isValueValid = attr1.Value.Length > 3;
                            logicDescription = "Value should be a meaningful content.";
                            break;
                    }

                    if (isValueValid)
                    {
                        logFile.MyLogFile($"'{task.TaskName}' validation passed: {attr1.Value}");
                        task.Status = "Correct";
                        task.StatusDescription = logicDescription;
                        continue;
                    }
                    else
                    {
                        logFile.MyLogFile($"'{task.TaskName}' validation failed. Value '{attr1.Value}' doesn't match expected pattern: {expectedPattern}");
                        task.Status = "Error";
                        continue;
                    }
                } //LD-HTX & LP-HTX & Transportation

                if (task.LogicType == "CheckValuePerPage")
                {
                    logFile.MyLogFile($"Processing per-page validation for task '{task.TaskName}'");

                    // Get document information from hardcoded documents
                    int documentId = task.DocumentIds.FirstOrDefault();

                    try
                    {
                        // Get total number of pages in the document
                        var docAttributes = await agileDDApi.ReadValuesAsync(documentId);
                        if (!docAttributes.Success)
                        {
                            logFile.MyLogFile($"Failed to get document attributes: {docAttributes.ErrorMessage}");
                            task.Status = "Error";
                            continue;
                        }
                        int maxPage = docAttributes.Content.Any() ? docAttributes.Content.Max(a => a.page_number) : 1;
                        logFile.MyLogFile($"Document has {maxPage} pages. Checking '{task.TaskName}' on each page...");

                        // Get the attribute code for this task
                        var attributeToCheck = task.Attributes.FirstOrDefault();
                        if (attributeToCheck == null)
                        {
                            logFile.MyLogFile($"No matching attribute found for task {task.TaskName}.");
                            task.Status = "Error";
                            continue;
                        }

                        //Get relevant page categories
                        var allowedCategories = ParseRelevantCategories(attributeToCheck.RelevantPageCategories);

                        // Initialize counters and result tracking
                        int pagesChecked = 0;
                        int pagesWithValue = 0;
                        int pagesWithValidValue = 0;
                        StringBuilder detailedReport = new StringBuilder();
                        detailedReport.AppendLine($"<h3>Per-Page Validation Report for {task.TaskName}</h3>");
                        detailedReport.AppendLine("<table border='1' cellpadding='3' style='border-collapse: collapse;'>");
                        detailedReport.AppendLine("<tr><th>Page</th><th>Value Present</th><th>Valid</th><th>Details</th></tr>");

                        // New structure to store page-specific results for PDF report
                        List<PageValidationResult> pageResults = new List<PageValidationResult>();

                        // Before the per-page validation loop, declare a HashSet to track checked pages:
                        HashSet<int> checkedPageNumbers = new HashSet<int>();

                        // Check each page
                        for (int pageNumber = 1; pageNumber <= maxPage; pageNumber++)
                        {
                            // Get all attributes for the current page
                            var pageAttributesResponse = await agileDDApi.ReadValuesAsync(documentId, pageNumber, transportationPageCategoryAttributeId);
                            if (!pageAttributesResponse.Success)
                                continue;
                            var pageCategory = pageAttributesResponse.Content.FirstOrDefault()?.data ?? "";

                            //Validate the current page is relevant
                            if (!IsPageCategoryRelevant(pageCategory, allowedCategories))
                                continue;

                            // Get all text content from the current page
                            var pageContentResponse = await agileDDApi.ReadValuesAsync(documentId, pageNumber);

                            if (pageContentResponse.Success && pageContentResponse.Content != null)
                            {
                                // Extract and normalize all text from the page
                                string pageText = string.Join(" ", pageContentResponse.Content
                                    .Where(attr => !string.IsNullOrEmpty(attr.data))
                                    .Select(attr => attr.data));

                                string normalizedPageText = NormalizeValue(pageText);

                                // Check if page contains all four required words
                                bool hasDivision = normalizedPageText.Contains("division");
                                bool hasDistrict = normalizedPageText.Contains("district");
                                bool hasStandard = normalizedPageText.Contains("standard");


                                if ((hasDivision || hasDistrict) && hasStandard)
                                {
                                    logFile.MyLogFile($"Page {pageNumber}: Skipping - this is a standard sheet");
                                    continue; // Skip this page and move to the next
                                }
                            }

                            // Only increment pagesChecked if this page number hasn't already been counted
                            if (!checkedPageNumbers.Contains(pageNumber))
                            {
                                pagesChecked++;
                                checkedPageNumbers.Add(pageNumber);
                            }

                            bool valuePresent = false;
                            bool valueValid = false;
                            string valueDetails = "N/A";
                            string extractedValue = string.Empty;

                            var attrResponse = await agileDDApi.ReadValuesAsync(documentId, pageNumber, attributeToCheck.AttributeCode);

                            if (attrResponse.Success && attrResponse.Content?.Any() == true)
                            {
                                // For image attributes, presence = detection exists
                                valuePresent = true;

                                // Optional: keep raw data if needed
                                extractedValue = attrResponse.Content.FirstOrDefault()?.data ?? string.Empty;
                            }


                            if (valuePresent)
                            {
                                pagesWithValue++;

                                // Apply task-specific validation based on task name
                                switch (task.TaskName)
                                {
                                    case "Engineering Seal":
                                        valueValid = valuePresent; // present == valid
                                        valueDetails = valuePresent ? "Engineering seal detected on page" : "Not found";
                                        break;

                                    case "Client":
                                        valueValid = valuePresent; // present == valid
                                        valueDetails = valuePresent ? "Client detected on page" : "Not found";
                                        break;

                                    case "North Arrow":
                                        valueValid = valuePresent; // present == valid
                                        valueDetails = valuePresent ? "North arrow detected on page" : "Not found";
                                        break;

                                    case "Scale Bar":
                                        valueValid = valuePresent; // present == valid
                                        valueDetails = valuePresent ? "Scale bar detected on page" : "Not found";
                                        break;

                                    case "General Notes":
                                        var notesKeywords = new[] { "general notes", "notes", "general" };
                                        var normalizedNotesValue = NormalizeValue(extractedValue ?? string.Empty);
                                        valueValid = notesKeywords.Any(keyword => normalizedNotesValue.Contains(keyword));
                                        valueDetails = valueValid ? "General notes text detected" : "No general notes keywords found";
                                        break;

                                    default:
                                        valueValid = extractedValue.Length > 3;
                                        valueDetails = extractedValue;
                                        break;
                                }

                                if (valueValid)
                                    pagesWithValidValue++;
                            }
                            else
                            {
                                valueValid = false;
                                valueDetails = "Not found";
                            }


                            // Store page validation result
                            pageResults.Add(new PageValidationResult
                            {
                                PageNumber = pageNumber,
                                ValuePresent = valuePresent,
                                ValueValid = valueValid,
                                ValueDetails = valueDetails
                            });

                            // Log the result for this page
                            //string resultMessage = valuePresent
                            //    ? (valueValid ? "Valid" : "Invalid")
                            //    : "Not found";
                            //logFile.MyLogFile($"Page {pageNumber}: {task.TaskName} - {resultMessage} - {valueDetails}");

                            // Add to detailed report
                            string rowColor = !valuePresent ? "#FFEEEE" : (valueValid ? "#EEFFEE" : "#FFEEEE");
                            detailedReport.AppendLine($"<tr style='background-color: {rowColor};'>");
                            detailedReport.AppendLine($"<td>{pageNumber}</td>");
                            detailedReport.AppendLine($"<td>{valuePresent}</td>");
                            detailedReport.AppendLine($"<td>{valueValid}</td>");
                            detailedReport.AppendLine($"<td>{HttpUtility.HtmlEncode(valueDetails)}</td>");
                            detailedReport.AppendLine("</tr>");
                        }

                        detailedReport.AppendLine("</table>");
                        detailedReport.AppendLine($"<p>Summary: Found on {pagesWithValue} of {pagesChecked} pages. Valid on {pagesWithValidValue} pages.</p>");

                        // Calculate page coverage percentage
                        double coveragePercentage = pagesChecked > 0
                            ? (double)pagesWithValidValue / pagesChecked * 100
                            : 0;

                        // Determine task status based on validation results
                        if (pagesChecked == 0)
                        {
                            task.Status = "Error";
                            logFile.MyLogFile($"No pages could be checked for '{task.TaskName}'.");
                        }
                        else if (coveragePercentage >= 90)  // If valid on 90% or more of pages
                        {
                            task.Status = "Correct";
                            logFile.MyLogFile($"'{task.TaskName}' validation passed with {coveragePercentage:F1}% coverage.");
                        }
                        else if (coveragePercentage >= 50)  // If valid on 50-90% of pages
                        {
                            task.Status = "Partial Match";
                            logFile.MyLogFile($"'{task.TaskName}' validation partially passed with {coveragePercentage:F1}% coverage.");
                        }
                        else
                        {
                            task.Status = "Error";
                            logFile.MyLogFile($"'{task.TaskName}' validation failed with only {coveragePercentage:F1}% coverage.");
                        }

                        // Store summary in attribute value for email reporting
                        var summaryAttribute = task.Attributes.FirstOrDefault();
                        if (summaryAttribute != null)
                        {
                            summaryAttribute.Value = $"Found on {pagesWithValue}/{pagesChecked} pages. Valid on {pagesWithValidValue} pages ({coveragePercentage:F1}%).";
                        }

                        // Save the detailed report for potential attachment or inclusion in email
                        string detailedReportPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PerPageReport_Run{runId}_Task{task.TaskId}.html");
                        File.WriteAllText(detailedReportPath, detailedReport.ToString());

                        // Add the file path to a property that can be used when sending the email
                        task.Attributes.Add(new PlanReviewAttribute
                        {
                            AttributeName = "DetailedReportPath",
                            Value = detailedReportPath
                        });

                        // Add another attribute to store page results in serialized format
                        task.Attributes.Add(new PlanReviewAttribute
                        {
                            AttributeName = "PageValidationResults",
                            Value = Newtonsoft.Json.JsonConvert.SerializeObject(pageResults)
                        });

                        logFile.MyLogFile($"Detailed report for '{task.TaskName}' saved to {detailedReportPath}");
                    }
                    catch (Exception ex)
                    {
                        logFile.MyLogFile($"Error during per-page validation for '{task.TaskName}': {ex.Message}");
                        task.Status = "Error";
                    }
                } //LD-HTX & Transportation
            }

            // Update the status of the PlanReview_Run in SQL
            using (var conn = new SqlConnection(new SqlConnectionStringBuilder
            {
                DataSource = sqlDatasource,
                UserID = sqlUserID,
                Password = sqlPassword,
                InitialCatalog = "REDACTED_DATABASE"
            }.ConnectionString))
            using (var cmd = new SqlCommand("UPDATE PlanReview_Run SET Status = @Status WHERE ID = @RunId", conn))
            {
                cmd.Parameters.AddWithValue("@Status", "Completed");
                cmd.Parameters.AddWithValue("@RunId", runId);
                await conn.OpenAsync();
                int rowsAffected = await cmd.ExecuteNonQueryAsync();
                if (rowsAffected > 0)
                {
                    logFile.MyLogFile($"PlanReview_Run status updated to 'Completed' for runId {runId}.");
                }
                else
                {
                    logFile.MyLogFile($"No PlanReview_Run found with runId {runId} to update.");
                }
            }
            logFile.MyLogFile($"Plan Review process for runId {runId} completed successfully at {DateTime.Now}.");

            // Send email notification
            try
            {
                if (!string.IsNullOrEmpty(planReviewRun.CreatedByEmail))
                {
                    string emailBody = await GenerateEmailBodyAsync(runId, taskList);
                    string subject = "";
                    if (planReviewRun.PlanReviewType == "Platting")
                    {
                        // Get Plat Name from the task list
                        var platNameTask = taskList.FirstOrDefault(t => t.TaskName == "Plat Name");
                        string platName = "Unknown";

                        if (platNameTask != null && platNameTask.Attributes.Any())
                        {
                            var platNameAttr = platNameTask.Attributes.FirstOrDefault(a => !string.IsNullOrEmpty(a.Value));
                            if (platNameAttr != null && !string.IsNullOrEmpty(platNameAttr.Value))
                            {
                                platName = platNameAttr.Value;
                                // Clean the plat name for use in subject/filename (remove invalid characters)
                                platName = System.Text.RegularExpressions.Regex.Replace(platName, @"[^\w\s-]", "");
                                platName = platName.Trim();
                            }
                        }

                        subject = $"{platName} - Pre-QAQC Review Results";
                    }
                    else
                    {
                        subject = $"Plan Review Results - Run #{runId} - {planReviewRun.PlanReviewType}";
                    }

                    await SendEmailNotificationAsync(runId, planReviewRun.CreatedByEmail, subject, emailBody, taskList);
                    logFile.MyLogFile($"Email notification sent to {planReviewRun.CreatedByEmail} for runId {runId}.");
                }
                else
                {
                    logFile.MyLogFile($"No email address available to send notification for runId {runId}.");
                }
            }
            catch (Exception ex)
            {
                // Don't let email errors affect the response
                logFile.MyLogFile($"Error sending email notification: {ex.Message}");
            }

            return new PlanReviewResponse
            {
                Status = "Success",
                Message = $"Plan Review process completed for runId {runId}.",
                Tasks = taskList
            };
        }

        #region Platting Helper Methods
        private async Task<ValidationResult> ValidateGeneralNoteByCondition(
            bool conditionMet,
            string conditionDescription,
            string noteTerm,
            int documentIdPlat,
            string generalNotesValue,
            PlanReviewGeneralNote noteTask,
            bool currentAllNotesFound)
        {
            // Check if note exists in GeneralNotes
            bool noteFound = FuzzyMatch(generalNotesValue, noteTerm);

            // Also try search API if fuzzy match didn't find it
            if (!noteFound)
            {
                var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTerm);
                if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                {
                    logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                    noteFound = true;
                }
            }

            var result = new ValidationResult { AllFound = currentAllNotesFound };

            if (conditionMet)
            {
                // Condition met: note should exist
                if (noteFound)
                {
                    logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - {conditionDescription} and note exists in General Notes");
                    result.NoteResult = $"{noteTask.Task}: Correct";
                    result.NoteDetails = $"{conditionDescription} AND note present in General Notes";
                    result.NoteFound = true;
                }
                else
                {
                    logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' - {conditionDescription} but note MISSING from General Notes");
                    result.NoteResult = $"{noteTask.Task}: Missing";
                    result.NoteDetails = $"{conditionDescription} BUT note is MISSING from General Notes";
                    result.AllFound = false;
                    result.NoteFound = false;
                }
            }
            else
            {
                // Condition not met: note should not exist
                if (noteFound)
                {
                    logFile.MyLogFile($"✗ ORPHANED: '{noteTask.Task}' - {conditionDescription} NOT met but note exists in General Notes");
                    result.NoteResult = $"{noteTask.Task}: Orphaned";
                    result.NoteDetails = $"{conditionDescription} NOT met BUT note present in General Notes";
                    result.AllFound = false;
                    result.NoteFound = false;
                }
                else
                {
                    logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - {conditionDescription} NOT met and note does not exist in General Notes");
                    result.NoteResult = $"{noteTask.Task}: Correct";
                    result.NoteDetails = $"{conditionDescription} NOT met AND note not present in General Notes";
                    result.NoteFound = false;
                }
            }

            return result;
        }
        private async Task<ValidationResult> ValidateDedicationByCondition(
            bool conditionMet,
            string conditionDescription,
            string noteTerm,
            string dedicationValue,
            PlanReviewDedication noteTask,
            int documentIdPlat,
            bool currentAllFound)
        {
            // Check if note exists in Dedication
            bool noteFound = FuzzyMatch(dedicationValue, noteTerm);

            // Also try search API if fuzzy match didn't find it
            if (!noteFound)
            {
                var searchResponse = await agileDDApi.SearchDocumentAsync(documentIdPlat, noteTerm);
                if (searchResponse.Success && searchResponse.Content?.results != null && searchResponse.Content.results.Count > 0)
                {
                    logFile.MyLogFile($"Search API found note '{noteTask.Task}' in document {documentIdPlat}");
                    noteFound = true;
                }
            }

            var result = new ValidationResult { AllFound = currentAllFound };

            if (conditionMet)
            {
                // Condition met: note should exist
                if (noteFound)
                {
                    logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - {conditionDescription} and note exists in Dedication");
                    result.NoteResult = $"{noteTask.Task}: Correct";
                    result.NoteDetails = $"{conditionDescription} AND note present in Dedication";
                    result.NoteFound = true;
                }
                else
                {
                    logFile.MyLogFile($"✗ MISSING: '{noteTask.Task}' - {conditionDescription} but note MISSING from Dedication");
                    result.NoteResult = $"{noteTask.Task}: Missing";
                    result.NoteDetails = $"{conditionDescription} BUT note is MISSING from Dedication";
                    result.AllFound = false;
                    result.NoteFound = false;
                }
            }
            else
            {
                // Condition not met: note should not exist
                if (noteFound)
                {
                    logFile.MyLogFile($"✗ ORPHANED: '{noteTask.Task}' - {conditionDescription} NOT met but note exists in Dedication");
                    result.NoteResult = $"{noteTask.Task}: Orphaned";
                    result.NoteDetails = $"{conditionDescription} NOT met BUT note present in Dedication";
                    result.AllFound = false;
                    result.NoteFound = false;
                }
                else
                {
                    logFile.MyLogFile($"✓ CORRECT: '{noteTask.Task}' - {conditionDescription} NOT met and note does not exist in Dedication");
                    result.NoteResult = $"{noteTask.Task}: Correct";
                    result.NoteDetails = $"{conditionDescription} NOT met AND note not present in Dedication";
                    result.NoteFound = false;
                }
            }

            return result;
        }
        private bool IsNoteApplicableToJurisdiction(string applicableJurisdiction, string currentJurisdiction)
        {
            if (string.IsNullOrEmpty(applicableJurisdiction) || applicableJurisdiction == "All")
                return true;

            string normalized = NormalizeValue(currentJurisdiction);
            string normalizedApplicable = NormalizeValue(applicableJurisdiction);

            return normalizedApplicable.Contains(normalized) || normalized.Contains(normalizedApplicable);
        }
        public async Task<bool> IsPermanentAccessEasementAsync(int documentId, PlanReviewTask task)
        {
            var cpc101Term = "PAE";
            var dedicationTerm = "permanent access easements shall be hereby established";
            try
            {
                //first search the CP101 document using the search api for the search term
                var searchResponse = await agileDDApi.SearchDocumentAsync(documentId, cpc101Term);
                if (!searchResponse.Success)
                {
                    logFile.MyLogFile($"Search API call failed for document {documentId} with term '{cpc101Term}': {searchResponse.ErrorMessage}");
                    return false;
                }
                if (searchResponse.Content != null && searchResponse.Content.results != null && searchResponse.Content.results.Count > 0)
                {
                    logFile.MyLogFile($"Search term '{cpc101Term}' found in document {documentId}.");
                    return true;
                }
                else
                {
                    if (task.TaskName != "PAEParagraph") //this is to avoid checking dedication attribute for the PAEParagraph task bc it would be searching for itself
                    {
                        //now check the dedication attribute for the search term
                        var dedicationAttribute = task.Attributes.FirstOrDefault(attr => attr.AttributeName == "Dedication");
                        if (dedicationAttribute != null && !string.IsNullOrEmpty(dedicationAttribute.Value))
                        {
                            //normalize and fuzzy match
                            string normalizedDedicationText = NormalizeValue(dedicationAttribute.Value);
                            string normalizedSearchTerm = NormalizeValue(dedicationTerm);
                            if (normalizedDedicationText.Contains(normalizedSearchTerm))
                            {
                                logFile.MyLogFile($"Search term '{dedicationTerm}' found in Dedication attribute of document {documentId}.");
                                return true;
                            }
                        }
                    }
                    logFile.MyLogFile($"Search term '{dedicationTerm}' not found in document {documentId}.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                logFile.MyLogFile($"Error during search in document {documentId} for term '{dedicationTerm}': {ex.Message}");
                return false;
            }
        }
        public async Task<bool> IsSharedDrivewayAsync(int documentId)
        {
            var searchTerm = "shared driveway";
            try
            {
                // Search the CPC101 document using the search API for the search term
                var searchResponse = await agileDDApi.SearchDocumentAsync(documentId, searchTerm);
                if (!searchResponse.Success)
                {
                    logFile.MyLogFile($"Search API call failed for document {documentId} with term '{searchTerm}': {searchResponse.ErrorMessage}");
                    return false;
                }
                if (searchResponse.Content != null && searchResponse.Content.results != null && searchResponse.Content.results.Count > 0)
                {
                    logFile.MyLogFile($"Search term '{searchTerm}' found in document {documentId}.");
                    return true;
                }
                else
                {
                    logFile.MyLogFile($"Search term '{searchTerm}' not found in document {documentId}.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                logFile.MyLogFile($"Error during search in document {documentId} for term '{searchTerm}': {ex.Message}");
                return false;
            }
        }
        #endregion

        #region LLM Extractor Helper Methods
        private async Task ProcessTitleBlockExtractionAsync(string titleBlockText)
        {
            try
            {
                if (string.IsNullOrEmpty(titleBlockText))
                {
                    logFile.MyLogFile("No TitleBlock text provided for extraction.");
                    return;
                }

                logFile.MyLogFile($"Processing TitleBlock with value: {titleBlockText}");

                // Initialize Azure OpenAI API
                string systemPrompt = @"You are an expert at extracting structured data from title blocks on architectural and engineering drawings. 
                                        You will analyze the provided title block text and extract key fields in the exact JSON format specified.
                                        Ensure all values are accurately captured and properly formatted.";

                AzureOpenAIApi azureOpenAIApi = new AzureOpenAIApi(
                    logFile,
                    systemPrompt,
                    Convert.ToDecimal(0.2)
                );

                // Create the prompt for the LLM
                string userPrompt = $@"Task: Extract the following values from the text of a subdivision title block. Only use information explicitly present in the text.

                            Text to extract from:
                            {titleBlockText}

                            Fields to extract:
                            Acreage
                            Blocks
                            Lots
                            Reserves
                            Scale
                            Date
                            Owner Name
                            Location
                            Abstract Number (Format: A-###)
                            City/ETJ
                            - Logic: If the text contains ""City of <X>"", return City; otherwise return ETJ.
                            County

                            Extraction Rules:
                            Always return one value per field.
                            Use the first matching value unless a later value is clearly a correction.
                            If a field is missing, return an empty string (do not guess).
                            Normalize formatting:
                            Acreage → numeric (e.g., 1.435)
                            Scale → preserve original formatting (e.g., 1"" = 50')
                            Abstract → return only the first abstract (e.g., A-323)
                            Owner → text exactly as it appears
                            County → Proper case (e.g., Harris)

                            Optional / Nullable Fields:
                            Some fields may not appear in the text (e.g., “lots”).
                            If a field is missing, do not infer or guess.
                            Instead, return an empty string ("""") or null based on the required output format.

                            Null Value Rules:
                            If a field is present → extract its value normally.
                            If a field is not present in the text → return null.
                            Do NOT generate values, estimate, or interpret implied data.

                            Output Format:
                            Return the result as clean JSON and nothing else:
                            {{
                              ""Acreage"": null,
                              ""Block"": null,
                              ""Reserves"": null,
                              ""Lots"": null,
                              ""Scale"": null,
                              ""Action Date"": null,
                              ""Owner"": null,
                              ""Location"": null,
                              ""Abstract Number"": null,
                              ""City/ETJ"": null,
                              ""County"": null
                            }}

                            Example Title Block:
                            leeland at emancipation reserve a subdivision of 1.435 acres of land located in the j.s. holman survey, a-323 and the james wells survey, abstract 832 city of houston, harris county, texas also being of all of block 579, south side buffalo bayou an unrecorded subdivision reserves: 1 blocks : 2 scale: 1=50 date: november, 2024 owner: bayview usa holdings, llc ravis street, suite 1555 iston tv 77002 713.568.3055 bge, inc. bge 10777 westheimer, suite 500, houston, tx 77042 tel: 281-558-8700 · www.bgeinc.com tbpe registration no. f-1046 tbpls licensed surveying firm no. 101065-00

                            Correct Value Based on Above example:
                            {{
                              ""Acreage"": 1.435,
                              ""Block"": 2,
                              ""Reserves"": 1,
                              ""Lots"": null,
                              ""Scale"": ""1=50"",
                              ""Action Date"": ""november, 2024"",
                              ""Owner"": ""bayview usa holdings, llc"",
                              ""Location"": ""j.s. holman survey"",
                              ""Abstract Number"": ""a-323"",
                              ""City/ETJ"": ""city"",
                              ""County"": ""Harris""
                            }}";

                // Call the LLM to extract title block values
                var extractionResult = await azureOpenAIApi.GetCompletionAsync(userPrompt, 5000);

                if (!extractionResult.Success)
                {
                    logFile.MyLogFile($"Error extracting TitleBlock values: {extractionResult.ErrorMessage}");
                    return;
                }

                logFile.MyLogFile($"TitleBlock extraction response: {extractionResult.Content}");

                // Parse the JSON response and create TitleBlock object
                try
                {
                    // Extract JSON from response (in case there's extra text)
                    var jsonMatch = System.Text.RegularExpressions.Regex.Match(
                        extractionResult.Content,
                        @"\{(?:[^{}]|""(?:[^""]|\\.)*"")*\}",
                        System.Text.RegularExpressions.RegexOptions.Singleline);

                    string jsonContent = jsonMatch.Success ? jsonMatch.Value : extractionResult.Content;
                    var extractedValues = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonContent);

                    if (extractedValues == null || extractedValues.Count == 0)
                    {
                        logFile.MyLogFile("Failed to extract any TitleBlock values from LLM response");
                        return;
                    }

                    logFile.MyLogFile($"Successfully extracted {extractedValues.Count} TitleBlock fields");

                    // Create TitleBlock object from extracted values
                    TitleBlock curTitleBlock = CreateTitleBlockFromExtractedValues(extractedValues);

                    // Store for later reference
                    titleBlock = curTitleBlock;

                    logFile.MyLogFile("TitleBlock extraction completed successfully");
                }
                catch (Newtonsoft.Json.JsonException ex)
                {
                    logFile.MyLogFile($"Error parsing TitleBlock JSON response: {ex.Message}");
                    logFile.MyLogFile($"Response content: {extractionResult.Content}");
                    return;
                }
            }
            catch (Exception ex)
            {
                logFile.MyLogFile($"Error during TitleBlock processing: {ex.Message}");
            }
        }

        private TitleBlock CreateTitleBlockFromExtractedValues(Dictionary<string, object> extractedValues)
        {
            var titleBlock = new TitleBlock();

            // Helper function for string conversion
            Func<string, string> GetStringValue = (key) =>
            {
                if (extractedValues.ContainsKey(key))
                {
                    var value = extractedValues[key];
                    if (value != null && !string.IsNullOrEmpty(value.ToString()))
                    {
                        return value.ToString();
                    }
                }
                return null;
            };

            // Helper function for numeric conversion
            Func<string, decimal?> GetNumericValue = (key) =>
            {
                if (extractedValues.ContainsKey(key))
                {
                    var value = extractedValues[key];
                    if (value != null)
                    {
                        if (decimal.TryParse(value.ToString(), out decimal result))
                            return result;
                    }
                }
                return null;
            };

            // Helper function for integer conversion
            Func<string, int?> GetIntValue = (key) =>
            {
                if (extractedValues.ContainsKey(key))
                {
                    var value = extractedValues[key];
                    if (value != null)
                    {
                        if (int.TryParse(value.ToString(), out int result))
                            return result;
                    }
                }
                return null;
            };

            // Map extracted values to TitleBlock properties
            titleBlock.Acreage = GetNumericValue("Acreage");
            titleBlock.Blocks = GetIntValue("Block");
            titleBlock.Lots = GetIntValue("Lots");
            titleBlock.Reserves = GetIntValue("Reserves");
            titleBlock.Scale = GetStringValue("Scale");
            titleBlock.ActionDate = GetStringValue("Action Date");
            titleBlock.Owner = GetStringValue("Owner");
            titleBlock.Location = GetStringValue("Location");
            titleBlock.AbstractNumber = GetStringValue("Abstract Number");
            titleBlock.CityETJ = GetStringValue("City/ETJ");
            titleBlock.County = GetStringValue("County");

            // Log any missing or null fields
            if (titleBlock.Acreage == null || titleBlock.Acreage == 0)
            {
                logFile.MyLogFile("Warning: Acreage field is missing or zero from TitleBlock extraction");
            }
            if (titleBlock.Blocks == null || titleBlock.Blocks == 0)
            {
                logFile.MyLogFile("Warning: Blocks field is missing or zero from TitleBlock extraction");
            }
            if (titleBlock.Lots == null)
            {
                logFile.MyLogFile("Warning: Lots field is missing from TitleBlock extraction (nullable field)");
            }
            if (titleBlock.Reserves == null || titleBlock.Reserves == 0)
            {
                logFile.MyLogFile("Warning: Reserves field is missing or zero from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.Scale))
            {
                logFile.MyLogFile("Warning: Scale field is missing from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.ActionDate))
            {
                logFile.MyLogFile("Warning: ActionDate field is missing from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.Owner))
            {
                logFile.MyLogFile("Warning: Owner field is missing from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.Location))
            {
                logFile.MyLogFile("Warning: Location field is missing from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.AbstractNumber))
            {
                logFile.MyLogFile("Warning: AbstractNumber field is missing from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.CityETJ))
            {
                logFile.MyLogFile("Warning: CityETJ field is missing from TitleBlock extraction");
            }
            if (string.IsNullOrEmpty(titleBlock.County))
            {
                logFile.MyLogFile("Warning: County field is missing from TitleBlock extraction");
            }

            return titleBlock;
        }

        private string GetTitleBlockPropertyValue(string taskName, TitleBlock titleBlock)
        {
            if (titleBlock == null)
                return null;

            switch (taskName)
            {
                case "Acreage":
                    return titleBlock.Acreage?.ToString();
                case "Block":
                    return titleBlock.Blocks?.ToString();
                case "Lots":
                    return titleBlock.Lots?.ToString();
                case "Reserves":
                    return titleBlock.Reserves?.ToString();
                case "Scale":
                    return titleBlock.Scale;
                case "Action Date":
                    return titleBlock.ActionDate;
                case "Owner":
                    return titleBlock.Owner;
                case "Location":
                    return titleBlock.Location;
                case "Abstract Number":
                    return titleBlock.AbstractNumber;
                case "City/ETJ":
                    return titleBlock.CityETJ;
                case "County":
                    return titleBlock.County;
                default:
                    logFile.MyLogFile($"Warning: Unknown task name '{taskName}' for TitleBlock comparison.");
                    return null;
            }
        }

        //ParseDedicationAttributeAsync(dedicationValue, logFile); goal is to get specific values from the dedication attribute by parsing it with LLM and storing those values in an object for later use.
        //the values are County, Acreage, Development Name, and Lots. I have an example of what text to look for and where those values would be in the dedication text.
        private async Task<DedicationInfo> ParseDedicationAttributeAsync(string dedicationValue, LogFile logFile)
        {
            try
            {
                if (string.IsNullOrEmpty(dedicationValue))
                {
                    logFile.MyLogFile("No Dedication text provided for extraction.");
                    return new DedicationInfo();
                }

                logFile.MyLogFile($"Processing Dedication attribute with value: {dedicationValue}");

                // Initialize Azure OpenAI API
                string systemPrompt = @"You are an expert at extracting structured data from dedication text on architectural and engineering drawings. 
                                        You will analyze the provided dedication text and extract key fields in the exact JSON format specified.
                                        Ensure all values are accurately captured and properly formatted.";

                AzureOpenAIApi azureOpenAIApi = new AzureOpenAIApi(
                    logFile,
                    systemPrompt,
                    Convert.ToDecimal(0.2)
                );

                // Create the prompt for the LLM
                string userPrompt = $@"Task: Extract the following values from dedication text on a plat. Only use information explicitly present in the text.

                            Text to extract from:
                            {dedicationValue}

                            Fields to extract:
                            County
                            Acreage
                            Development Name

                            Extraction Rules:
                            Always return one value per field.
                            Use the first matching value unless a later value is clearly a correction.
                            If a field is missing, return null (do not guess).
                            Normalize formatting:
                            County → Proper case (e.g., Harris)
                            Acreage → numeric (e.g., 5.25)
                            Development Name → text exactly as it appears

                            Null Value Rules:
                            If a field is present → extract its value normally.
                            If a field is not present in the text → return null.
                            Do NOT generate values, estimate, or interpret implied data.

                            Output Format:
                            Return the result as clean JSON and nothing else:
                            {{
                              ""County"": null,
                              ""Acreage"": null,
                              ""DevelopmentName"": null
                            }}

                            Example Dedication Text:
                            APPENDIX A:	OWNERS ACKNOWLEDGEMENT STATE OF TEXAS COUNTY OF (county) We, (name of owner or owners), acting by and through (name and title of officer), being officers of (name of company or corporation), owner (or owners) hereinafter referred to as Owners (whether one or more) of the (number of Acreage)  tract described in the above and foregoing map of (name of DevelopmentName), do hereby...

                            Example Dedication Text with Values:
                            APPENDIX A:	OWNERS ACKNOWLEDGEMENT STATE OF TEXAS COUNTY OF HARRIS We, BAYVIEW USA HOLDINGS, LLC, acting by and through Dennis Thompson Managing Director, being officers of BAYVIEW USA HOLDINGS, LLC, owner (or owners) hereinafter referred to as Owners (whether one or more) of the 1.435 tract described in the above and foregoing map of LEELAND AT EMANCIPATION RESERVE, do hereby...

                            Correct Value Based on Above example:
                            {{
                              ""County"": ""Harris"",
                              ""Acreage"": 1.435,
                              ""DevelopmentName"": ""LEELAND AT EMANCIPATION RESERVE""
                            }}";

                // Call the LLM to extract dedication values
                var extractionResult = await azureOpenAIApi.GetCompletionAsync(userPrompt, 2000);

                if (!extractionResult.Success)
                {
                    logFile.MyLogFile($"Error extracting Dedication values: {extractionResult.ErrorMessage}");
                    return new DedicationInfo();
                }

                logFile.MyLogFile($"Dedication extraction response: {extractionResult.Content}");

                // Parse the JSON response and create DedicationInfo object
                try
                {
                    // Extract JSON from response (in case there's extra text)
                    var jsonMatch = System.Text.RegularExpressions.Regex.Match(
                        extractionResult.Content,
                        @"\{(?:[^{}]|""(?:[^""]|\\.)*"")*\}",
                        System.Text.RegularExpressions.RegexOptions.Singleline);

                    string jsonContent = jsonMatch.Success ? jsonMatch.Value : extractionResult.Content;
                    var extractedValues = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonContent);

                    if (extractedValues == null || extractedValues.Count == 0)
                    {
                        logFile.MyLogFile("Failed to extract any Dedication values from LLM response");
                        return new DedicationInfo();
                    }

                    logFile.MyLogFile($"Successfully extracted {extractedValues.Count} Dedication fields");

                    // Create DedicationInfo object from extracted values
                    var dedicationInfo = new DedicationInfo
                    {
                        County = GetStringValue(extractedValues, "County"),
                        Acreage = GetDecimalValue(extractedValues, "Acreage") ?? 0,
                        DevelopmentName = GetStringValue(extractedValues, "DevelopmentName")
                    };

                    logFile.MyLogFile("Dedication extraction completed successfully");
                    return dedicationInfo;
                }
                catch (Newtonsoft.Json.JsonException ex)
                {
                    logFile.MyLogFile($"Error parsing Dedication JSON response: {ex.Message}");
                    logFile.MyLogFile($"Response content: {extractionResult.Content}");
                    return new DedicationInfo();
                }
            }
            catch (Exception ex)
            {
                logFile.MyLogFile($"Error during Dedication processing: {ex.Message}");
                return new DedicationInfo();
            }
        }

        private string GetStringValue(Dictionary<string, object> values, string key)
        {
            if (values.ContainsKey(key))
            {
                var value = values[key];
                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    return value.ToString();
                }
            }
            return null;
        }

        private decimal? GetDecimalValue(Dictionary<string, object> values, string key)
        {
            if (values.ContainsKey(key))
            {
                var value = values[key];
                if (value != null && decimal.TryParse(value.ToString(), out decimal result))
                {
                    return result;
                }
            }
            return null;
        }
        #endregion

        #region SQL Handlers
        public async Task<PlanReviewRun> GetPlanReviewRunAsync(int runId)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder
            {
                DataSource = sqlDatasource,
                UserID = sqlUserID,
                Password = sqlPassword,
                InitialCatalog = "REDACTED_DATABASE"
            };

            var query = @"SELECT ID, CreatedBy, CreatedByEmail, Timestamp, Status, PlanReviewType, TaskSet, ProjectNumber
                  FROM PlanReview_Run 
                  WHERE ID = @RunId";

            using (var conn = new SqlConnection(builder.ConnectionString))
            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@RunId", runId);
                await conn.OpenAsync();

                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    if (await reader.ReadAsync())
                    {
                        return new PlanReviewRun
                        {
                            ID = reader.GetInt32(reader.GetOrdinal("ID")),
                            CreatedBy = reader.GetString(reader.GetOrdinal("CreatedBy")),
                            CreatedByEmail = !reader.IsDBNull(reader.GetOrdinal("CreatedByEmail")) ? reader.GetString(reader.GetOrdinal("CreatedByEmail")) : null,
                            Timestamp = reader.GetDateTime(reader.GetOrdinal("Timestamp")),
                            Status = reader.GetString(reader.GetOrdinal("Status")),
                            PlanReviewType = reader.GetString(reader.GetOrdinal("PlanReviewType")),
                            TaskSet = reader.GetString(reader.GetOrdinal("TaskSet")),
                            ProjectNumber = !reader.IsDBNull(reader.GetOrdinal("ProjectNumber")) ? reader.GetString(reader.GetOrdinal("ProjectNumber")) : null
                        };
                    }
                }
            }
            return null;
        }
        public async Task<List<PlanReviewTask>> GetPlanReviewTasksWithAttributesAsync(string planReviewType, string taskSet)
        {
            var result = new List<PlanReviewTask>();
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = sqlDatasource,
                UserID = sqlUserID,
                Password = sqlPassword,
                InitialCatalog = "REDACTED_DATABASE"
            };

            string query = @"
            SELECT t.ID AS TaskId, t.PlanReviewType, t.TaskName, t.TaskSet, t.LogicType,
                   a.ID AS AttributeId, a.AttributeName, a.FileCategory, a.AttributeCode, a.RelevantPageCategories
            FROM PlanReview_Task t
            LEFT JOIN PlanReview_Attribute a ON a.PlanReviewTask = t.ID
            WHERE t.PlanReviewType = @PlanReviewType AND t.TaskSet = @TaskSet
            ORDER BY t.ID, a.ID";

            using (var conn = new SqlConnection(builder.ConnectionString))
            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@PlanReviewType", planReviewType);
                cmd.Parameters.AddWithValue("@TaskSet", taskSet);
                await conn.OpenAsync();

                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    int? lastTaskId = null;
                    PlanReviewTask currentTask = null;

                    while (await reader.ReadAsync())
                    {
                        int taskId = reader.GetInt32(reader.GetOrdinal("TaskId"));
                        if (lastTaskId != taskId)
                        {
                            currentTask = new PlanReviewTask
                            {
                                TaskId = taskId,
                                PlanReviewType = reader.GetString(reader.GetOrdinal("PlanReviewType")),
                                TaskName = reader.GetString(reader.GetOrdinal("TaskName")),
                                TaskSet = reader.GetString(reader.GetOrdinal("TaskSet")),
                                LogicType = reader.GetString(reader.GetOrdinal("LogicType"))
                            };
                            result.Add(currentTask);
                            lastTaskId = taskId;
                        }

                        if (!reader.IsDBNull(reader.GetOrdinal("AttributeId")))
                        {
                            currentTask.Attributes.Add(new PlanReviewAttribute
                            {
                                AttributeId = reader.GetInt32(reader.GetOrdinal("AttributeId")),
                                AttributeName = reader.GetString(reader.GetOrdinal("AttributeName")),
                                FileCategory = reader.GetString(reader.GetOrdinal("FileCategory")),
                                AttributeCode = reader.GetInt32(reader.GetOrdinal("AttributeCode")),
                                RelevantPageCategories = !reader.IsDBNull(reader.GetOrdinal("RelevantPageCategories")) ? reader.GetString(reader.GetOrdinal("RelevantPageCategories")) : null
                            });
                        }
                    }
                }
            }
            return result;
        }
        public async Task<List<PlanReviewGeneralNote>> GetGeneralNotesTasksAsync()
        {
            var result = new List<PlanReviewGeneralNote>();
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = sqlDatasource,
                UserID = sqlUserID,
                Password = sqlPassword,
                InitialCatalog = "REDACTED_DATABASE"
            };

            string query = @"
            SELECT GeneralNoteID, Source, Task, LogicCode, SearchTerm, NoteTerm, LogicDescription, Jurisdiction
            FROM dbo.PlanReview_GeneralNotes
            ORDER BY GeneralNoteID";

            using (var conn = new SqlConnection(builder.ConnectionString))
            using (var cmd = new SqlCommand(query, conn))
            {
                await conn.OpenAsync();

                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        result.Add(new PlanReviewGeneralNote
                        {
                            GeneralNoteID = reader.GetInt32(reader.GetOrdinal("GeneralNoteID")),
                            Task = reader.GetString(reader.GetOrdinal("Task")),
                            LogicCode = reader.GetString(reader.GetOrdinal("LogicCode")),
                            SearchTermRaw = !reader.IsDBNull(reader.GetOrdinal("SearchTerm"))
                                ? reader.GetString(reader.GetOrdinal("SearchTerm"))
                                : null,
                            NoteTerm = !reader.IsDBNull(reader.GetOrdinal("NoteTerm"))
                                ? reader.GetString(reader.GetOrdinal("NoteTerm"))
                                : null,
                            LogicDescription = !reader.IsDBNull(reader.GetOrdinal("LogicDescription"))
                                ? reader.GetString(reader.GetOrdinal("LogicDescription"))
                                : null,
                            Jurisdiction = !reader.IsDBNull(reader.GetOrdinal("Jurisdiction"))
                                ? reader.GetString(reader.GetOrdinal("Jurisdiction"))
                                : "All"
                        });
                    }
                }
            }

            return result;
        }
        public async Task<List<PlanReviewDedication>> GetDedicationTasksAsync()
        {
            var result = new List<PlanReviewDedication>();
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = sqlDatasource,
                UserID = sqlUserID,
                Password = sqlPassword,
                InitialCatalog = "REDACTED_DATABASE"
            };

            string query = @"
            SELECT Task, LogicCode, SearchTerm, NoteTerm, LogicDescription, Jurisdiction
            FROM dbo.PlanReview_Dedication";

            using (var conn = new SqlConnection(builder.ConnectionString))
            using (var cmd = new SqlCommand(query, conn))
            {
                await conn.OpenAsync();

                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        result.Add(new PlanReviewDedication
                        {
                            Task = reader.GetString(reader.GetOrdinal("Task")),
                            LogicCode = reader.GetString(reader.GetOrdinal("LogicCode")),
                            SearchTerm = !reader.IsDBNull(reader.GetOrdinal("SearchTerm"))
                                ? reader.GetString(reader.GetOrdinal("SearchTerm"))
                                : null,
                            NoteTerm = !reader.IsDBNull(reader.GetOrdinal("NoteTerm"))
                                ? reader.GetString(reader.GetOrdinal("NoteTerm"))
                                : null,
                            LogicDescription = !reader.IsDBNull(reader.GetOrdinal("LogicDescription"))
                                ? reader.GetString(reader.GetOrdinal("LogicDescription"))
                                : null,
                            Jurisdiction = !reader.IsDBNull(reader.GetOrdinal("Jurisdiction"))
                                ? reader.GetString(reader.GetOrdinal("Jurisdiction"))
                                : "All"
                        });
                    }
                }
            }

            return result;
        }

        #endregion

        #region Helper methods for value normalization and comparison
        private string NormalizeValue(string value)
        {

            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string cleaned = value.ToLower()
                                  .Replace("\n", "")
                                  .Replace("\r", "")
                                  .Replace("\\n", "")
                                  .Replace("\\r", "")
                                  .Replace("\\", "")           // Remove all backslashes
                                  .Replace("\"", "")           // Remove all double quotes
                                  .Replace("'", "")            // Remove all single quotes
                                  .Replace("preliminary plat", "")
                                  .Trim();

            // Remove extra spaces
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\s+", " ");

            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\s*,\s*", ",");

            return cleaned;
        }
        private bool ValuesMatch(string v1, string v2)
        {
            string n1 = NormalizeValue(v1);
            string n2 = NormalizeValue(v2);

            // Try to compare as decimals if possible
            if (decimal.TryParse(n1, out decimal d1) && decimal.TryParse(n2, out decimal d2))
                return d1 == d2;

            //try to compare using ExtractYearMonth for date-like strings
            if (IsSameMonth(n1, n2))
                return true;

            // Otherwise, compare as strings
            return n1 == n2;
        }
        private (int? Year, int? Month) ExtractYearMonth(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return (null, null);

            // Try standard parsing first
            if (DateTime.TryParse(input, out var dt))
                return (dt.Year, dt.Month);

            // Try to extract month and year from text like "april, 2025"
            var lower = input.ToLower();
            var months = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat.MonthNames
                .Where(m => !string.IsNullOrEmpty(m)).ToArray();

            for (int i = 0; i < months.Length; i++)
            {
                if (lower.Contains(months[i]))
                {
                    // Try to find a 4-digit year
                    var yearMatch = System.Text.RegularExpressions.Regex.Match(lower, @"\b\d{4}\b");
                    if (yearMatch.Success && int.TryParse(yearMatch.Value, out int year))
                        return (year, i + 1);
                }
            }

            // Try to extract year only
            var yearOnlyMatch = System.Text.RegularExpressions.Regex.Match(lower, @"\b\d{4}\b");
            if (yearOnlyMatch.Success && int.TryParse(yearOnlyMatch.Value, out int yearOnly))
                return (yearOnly, null);

            return (null, null);
        }
        private bool TryParseAndValidateDate(string dateString, out DateTime result)
        {
            // Try standard parsing with various formats
            if (DateTime.TryParse(dateString, out result))
                return true;

            // Custom parsing for text formats like "April, 2025"
            var lower = dateString.ToLower();
            var months = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat.MonthNames
                .Where(m => !string.IsNullOrEmpty(m)).ToArray();

            for (int i = 0; i < months.Length; i++)
            {
                if (lower.Contains(months[i]))
                {
                    // Try to find a 4-digit year
                    var yearMatch = System.Text.RegularExpressions.Regex.Match(lower, @"\b\d{4}\b");
                    if (yearMatch.Success && int.TryParse(yearMatch.Value, out int year))
                    {
                        // Create a date using the first day of the month
                        result = new DateTime(year, i + 1, 1);
                        return true;
                    }
                }
            }

            // If we get here, we couldn't parse the date
            result = DateTime.MinValue;
            return false;
        }
        private bool IsSameMonth(string v1, string v2)
        {
            var (year1, month1) = ExtractYearMonth(v1);
            var (year2, month2) = ExtractYearMonth(v2);

            return year1.HasValue && month1.HasValue &&
                   year2.HasValue && month2.HasValue &&
                   year1 == year2 && month1 == month2;
        }
        private bool FuzzyMatch(string source, string target, int maxDistance = 10)
        {
            source = NormalizeValue(source);
            target = NormalizeValue(target);
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
                return false;

            // For short strings that look like acronyms (all letters, possibly with dots)
            // Use strict word-boundary matching instead of fuzzy matching
            if (IsLikelyAcronym(target))
            {
                return AcronymExistsAsWord(source, target);
            }

            // For longer strings, use standard fuzzy matching
            if (source.Contains(target) || target.Contains(source))
                return true;

            if (source.Length < 5 || target.Length < 5)
                return source == target;

            if (source.Length > 10 && target.Length > 10)
            {
                var sourceWords = source.Split(' ').Where(w => w.Length > 2).ToHashSet();
                var targetWords = target.Split(' ').Where(w => w.Length > 2).ToHashSet();
                if (targetWords.Count < 3)
                    return false;
                double overlap = (double)sourceWords.Intersect(targetWords).Count() / Math.Min(sourceWords.Count, targetWords.Count);
                if (overlap >= 0.5)
                    return true;
            }

            return LevenshteinDistance(source, target) <= maxDistance;
        }
        private bool FuzzyMatch(string source, List<string> targets, int maxDistance = 10)
        {
            if (string.IsNullOrEmpty(source) || targets == null || targets.Count == 0)
                return false;

            // Return true if ANY of the target terms match the source
            return targets.Any(target => FuzzyMatch(source, target, maxDistance));
        }
        private bool IsLikelyAcronym(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length > 15)
                return false;

            // Remove periods and spaces to check the core
            string core = Regex.Replace(text, @"[\s.\-]", ""); // to allow hyphens

            return Regex.IsMatch(core, @"^[a-zA-Z0-9]{1,10}$");
        }
        private bool AcronymExistsAsWord(string source, string acronym)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(acronym))
                return false;

            // Normalize the acronym: remove periods and spaces to get the core letters
            // "A.E." → "ae", "B. L." → "bl"
            string acronymCore = System.Text.RegularExpressions.Regex.Replace(acronym.ToLower(), @"[\s.]", "");

            // Create a pattern that matches the acronym with any spacing/period variations
            // surrounded by word boundaries
            // For "ae": matches "ae", "a.e", "a.e.", "a. e.", "a .e" etc. at word boundaries
            string letterPattern = string.Join(@"\s*\.?\s*", acronymCore.ToCharArray());
            string pattern = $@"\b{letterPattern}\b";

            return System.Text.RegularExpressions.Regex.IsMatch(source, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
        private int LevenshteinDistance(string source, string target)
        {
            if (source.Length == 0) return target.Length;
            if (target.Length == 0) return source.Length;

            int[,] distance = new int[source.Length + 1, target.Length + 1];

            for (int i = 0; i <= source.Length; i++)
                distance[i, 0] = i;
            for (int j = 0; j <= target.Length; j++)
                distance[0, j] = j;

            for (int i = 1; i <= source.Length; i++)
            {
                for (int j = 1; j <= target.Length; j++)
                {
                    int cost = (source[i - 1] == target[j - 1]) ? 0 : 1;
                    distance[i, j] = Math.Min(Math.Min(
                        distance[i - 1, j] + 1,
                        distance[i, j - 1] + 1),
                        distance[i - 1, j - 1] + cost);
                }
            }

            return distance[source.Length, target.Length];
        }
        private bool CompareFieldValues(
        string value1,
        string value2,
        string value3,
        out string normalized1,
        out string normalized2,
        out string normalized3)
        {
            // Normalize all values
            normalized1 = NormalizeValue(value1 ?? "");
            normalized2 = NormalizeValue(value2 ?? "");
            normalized3 = NormalizeValue(value3 ?? "");

            // Try numeric comparison first (handles decimals with trailing zeros)
            if (decimal.TryParse(normalized1, out decimal d1) &&
                decimal.TryParse(normalized2, out decimal d2) &&
                decimal.TryParse(normalized3, out decimal d3))
            {
                return d1 == d2 && d2 == d3;
            }

            // Fall back to string comparison
            return normalized1 == normalized2 && normalized2 == normalized3;
        }
        #endregion

        #region Helper methods to help parse and check page categories
        private string ExtractFirstJsonObject(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            // First, try to remove markdown code fences if present
            string cleaned = text;

            // Remove markdown JSON code blocks: ```json ... ``` or ``` ... ```
            var markdownPattern = @"```(?:json)?\s*([\s\S]*?)\s*```";
            var markdownMatch = System.Text.RegularExpressions.Regex.Match(text, markdownPattern);
            if (markdownMatch.Success && markdownMatch.Groups.Count > 1)
            {
                cleaned = markdownMatch.Groups[1].Value.Trim();
                logFile.MyLogFile($"Removed markdown code fences from response");
            }

            int braceDepth = 0;
            int startIndex = -1;
            bool inString = false;
            bool escaped = false;

            for (int i = 0; i < cleaned.Length; i++)
            {
                char c = cleaned[i];

                // Handle string escaping
                if (escaped)
                {
                    escaped = false;
                    continue;
                }

                if (c == '\\' && inString)
                {
                    escaped = true;
                    continue;
                }

                if (c == '"')
                {
                    inString = !inString;
                    continue;
                }

                // Only process braces outside of strings
                if (!inString)
                {
                    if (c == '{')
                    {
                        if (braceDepth == 0)
                            startIndex = i;
                        braceDepth++;
                    }
                    else if (c == '}')
                    {
                        braceDepth--;
                        if (braceDepth == 0 && startIndex >= 0)
                        {
                            string jsonCandidate = cleaned.Substring(startIndex, i - startIndex + 1);

                            // Validate and clean the JSON before returning
                            try
                            {
                                // Try to parse it to validate
                                var testParse = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonCandidate);
                                return jsonCandidate;
                            }
                            catch (Newtonsoft.Json.JsonException)
                            {
                                // If parsing fails, try to fix common issues
                                string fixedJson = FixCommonJsonIssues(jsonCandidate);
                                try
                                {
                                    var testParse = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(fixedJson);
                                    logFile.MyLogFile($"Fixed JSON parsing issues in extracted object");
                                    return fixedJson;
                                }
                                catch
                                {
                                    // Continue searching for next JSON object
                                    continue;
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }
        private string FixCommonJsonIssues(string json)
        {
            if (string.IsNullOrEmpty(json))
                return json;

            // Fix JavaScript Infinity/NaN values (not valid JSON)
            json = System.Text.RegularExpressions.Regex.Replace(json, @"\bInfinity\b", "null");
            json = System.Text.RegularExpressions.Regex.Replace(json, @"\bNaN\b", "null");
            json = System.Text.RegularExpressions.Regex.Replace(json, @"\bundefined\b", "null");

            // Fix unquoted property names (common GPT error)
            json = System.Text.RegularExpressions.Regex.Replace(json, @"(\{|,)\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*:", "$1\"$2\":");

            // Fix single quotes to double quotes (if they're being used for strings)
            // This is tricky - only do it if there are no double quotes already
            if (!json.Contains("\"") && json.Contains("'"))
            {
                json = json.Replace("'", "\"");
            }

            return json;
        }
        private static string[] ParseRelevantCategories(string relevantPageCategories)
        {
            return string.IsNullOrWhiteSpace(relevantPageCategories)
                ? new string[0]
                : relevantPageCategories.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim()).ToArray();
        }
        private static bool IsPageCategoryRelevant(string pageCategory, string[] allowedCategories)
        {
            // If allowedCategories contains "All", always count as relevant
            if (allowedCategories.Any(c => string.Equals(c, "All", StringComparison.OrdinalIgnoreCase)))
                return true;

            return allowedCategories.Contains(pageCategory, StringComparer.OrdinalIgnoreCase);
        }
        #endregion

        #region Email Notification Methods
        public async Task SendEmailNotificationAsync(int runId, string toEmailAddress, string subject, string body, List<PlanReviewTask> tasks)
        {
            try
            {
                // Get tasks for the run
                var run = await GetPlanReviewRunAsync(runId);

                logFile.MyLogFile($"Sending email with {tasks.Count} tasks. " +
                    $"Status counts: Match={tasks.Count(t => t.Status == "Match")}, " +
                    $"Correct={tasks.Count(t => t.Status == "Correct")}, " +
                    $"Mismatch={tasks.Count(t => t.Status == "Mismatch")}, " +
                    $"Error={tasks.Count(t => t.Status == "Error")}, " +
                    $"Notice={tasks.Count(t => t.Status == "Notice" || t.Status == "Not enough data")}");

                // Generate PDF report
                string pdfReportPath = await GeneratePdfReportAsync(runId, tasks);

                // Get SMTP settings from config
                string smtpServer = ConfigurationManager.AppSettings["SMTP_SERVER"];
                int smtpPort = int.Parse(ConfigurationManager.AppSettings["SMTP_PORT"]);
                string smtpUsername = ConfigurationManager.AppSettings["SMTP_USERNAME"];
                string smtpPassword = ConfigurationManager.AppSettings["SMTP_PASSWORD"];
                string fromEmail = ConfigurationManager.AppSettings["SMTP_USERNAME"];
                string fromName = ConfigurationManager.AppSettings["SMTP_DISPLAY_NAME"];
                bool enableSsl = true;

                using (var client = new SmtpClient(smtpServer, smtpPort)
                {
                    Credentials = new NetworkCredential(smtpUsername, smtpPassword),
                    EnableSsl = enableSsl
                })
                {
                    var mailMessage = new MailMessage
                    {
                        From = new MailAddress(fromEmail, fromName),
                        Subject = subject,
                        Body = body,
                        IsBodyHtml = true
                    };

                    mailMessage.To.Add(toEmailAddress);

                    // Add PDF attachment
                    if (File.Exists(pdfReportPath))
                    {
                        try
                        {
                            var attachment = new Attachment(pdfReportPath);
                            mailMessage.Attachments.Add(attachment);
                            logFile.MyLogFile($"Added PDF report attachment to email: {pdfReportPath}");
                        }
                        catch (Exception ex)
                        {
                            logFile.MyLogFile($"Error attaching PDF file: {ex.Message}");
                        }
                    }
                    else
                    {
                        logFile.MyLogFile($"PDF report not found at path: {pdfReportPath}");
                    }

                    await client.SendMailAsync(mailMessage);
                    logFile.MyLogFile($"Successfully sent email notification with PDF attachment to {toEmailAddress} for run ID {runId}.");
                }
            }
            catch (Exception ex)
            {
                logFile.MyLogFile($"Error sending email notification: {ex.Message}");
                throw;
            }
        }

        public async Task<string> GenerateEmailBodyAsync(int runId, List<PlanReviewTask> tasks)
        {
            // Get run details
            var run = await GetPlanReviewRunAsync(runId);
            if (run == null)
            {
                return "Error: Plan Review run not found.";
            }

            // Sort tasks: GeneralNotes and Dedication at the end
            var orderedTasks = tasks
                .OrderBy(t => t.LogicType == "GeneralNotes" || t.LogicType == "Dedication" ? 1 : 0)
                .ThenBy(t => t.TaskName)
                .ToList();

            StringBuilder htmlBuilder = new StringBuilder();

            // Build email header without style block
            htmlBuilder.AppendLine("<!DOCTYPE html>");
            htmlBuilder.AppendLine("<html lang=\"en\">");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine("<meta charset=\"UTF-8\">");
            htmlBuilder.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body style=\"font-family: Arial, sans-serif; color: #333; line-height: 1.6;\">");

            // Email header
            if (run.PlanReviewType == "Platting")
            {
                // Get Plat Name from the task list
                var platNameTask = tasks.FirstOrDefault(t => t.TaskName == "Plat Name");
                string platName = "Unknown";

                if (platNameTask != null && platNameTask.Attributes.Any())
                {
                    var platNameAttr = platNameTask.Attributes.FirstOrDefault(a => !string.IsNullOrEmpty(a.Value));
                    if (platNameAttr != null && !string.IsNullOrEmpty(platNameAttr.Value))
                    {
                        platName = platNameAttr.Value;
                        // Clean the plat name for display
                        platName = System.Text.RegularExpressions.Regex.Replace(platName, @"[^\w\s-]", "");
                        platName = platName.Trim();
                    }
                }

                htmlBuilder.AppendLine($"<h1 style=\"color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;\">{platName} - Pre-QAQC Review Results</h1>");
            }
            else
            {
                htmlBuilder.AppendLine($"<h1 style=\"color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;\">Plan Review Results - Run #{run.ID}</h1>");
            }

            // Summary section
            htmlBuilder.AppendLine("<div style=\"background-color: #f8f9fa; border-left: 4px solid #3498db; padding: 15px; margin: 20px 0;\">");
            htmlBuilder.AppendLine($"<p><strong>Plan Review Type:</strong> {run.PlanReviewType}</p>");
            htmlBuilder.AppendLine($"<p><strong>Task Set:</strong> {run.TaskSet}</p>");
            htmlBuilder.AppendLine($"<p><strong>Created By:</strong> {run.CreatedBy}</p>");
            htmlBuilder.AppendLine($"<p><strong>Date:</strong> {run.Timestamp:MMMM dd, yyyy HH:mm tt}</p>");
            htmlBuilder.AppendLine($"<p><strong>Status:</strong> {run.Status}</p>");

            // Count tasks by status
            int matchCount = tasks.Count(t => t.Status == "Match");
            int correctCount = tasks.Count(t => t.Status == "Correct");
            int mismatchCount = tasks.Count(t => t.Status == "Mismatch");
            int errorCount = tasks.Count(t => t.Status == "Error");
            int noticeCount = tasks.Count(t => t.Status == "Notice" || t.Status == "Not enough data");

            htmlBuilder.AppendLine($"<p><strong>Summary:</strong> " +
                $"<span style=\"display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #d4edda; color: #155724;\">{matchCount + correctCount} Matches/Correct</span> " +
                $"<span style=\"display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #f8d7da; color: #721c24;\">{mismatchCount + errorCount} Mismatches/Errors</span> " +
                $"<span style=\"display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #fff3cd; color: #856404;\">{noticeCount} Notice</span></p>");
            htmlBuilder.AppendLine("</div>");

            // Results table
            htmlBuilder.AppendLine("<h2 style=\"color: #2c3e50; margin-top: 20px;\">Detailed Results</h2>");
            htmlBuilder.AppendLine("<table style=\"border-collapse: collapse; width: 100%; margin: 20px 0;\">");
            htmlBuilder.AppendLine("<tr>");
            htmlBuilder.AppendLine("<th style=\"border: 1px solid #ddd; padding: 10px; text-align: left; background-color: #f2f2f2; color: #333;\">Task Name</th>");
            htmlBuilder.AppendLine("<th style=\"border: 1px solid #ddd; padding: 10px; text-align: left; background-color: #f2f2f2; color: #333;\">Status</th>");
            htmlBuilder.AppendLine("<th style=\"border: 1px solid #ddd; padding: 10px; text-align: left; background-color: #f2f2f2; color: #333;\">Details</th>");
            htmlBuilder.AppendLine("</tr>");

            bool isEvenRow = false;
            foreach (var task in orderedTasks)
            {
                string statusStyle = "display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #d1ecf1; color: #0c5460;";
                if (task.Status != null)
                {
                    string status = task.Status.ToLower();
                    if (status == "match" || status == "correct")
                        statusStyle = "display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #d4edda; color: #155724;";
                    else if (status == "mismatch" || status == "error" || status == "incorrect")
                        statusStyle = "display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #f8d7da; color: #721c24;";
                    else if (status == "notice" || status == "not enough data" || status == "missing" || status == "orphaned")
                        statusStyle = "display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 14px; background-color: #fff3cd; color: #856404;";
                }

                string rowBgColor = isEvenRow ? " background-color: #f9f9f9;" : "";
                string cellStyle = $"border: 1px solid #ddd; padding: 10px; text-align: left;{rowBgColor}";

                htmlBuilder.AppendLine("<tr>");
                htmlBuilder.AppendLine($"<td style=\"{cellStyle}\">{HttpUtility.HtmlEncode(task.TaskName)}</td>");
                htmlBuilder.AppendLine($"<td style=\"{cellStyle}\"><span style=\"{statusStyle}\">{task.Status ?? "N/A"}</span></td>");

                // Details column with attribute values
                htmlBuilder.AppendLine($"<td style=\"{cellStyle}\">");
                if (task.Attributes.Count > 0)
                {
                    if (task.LogicType == "GeneralNotes" && task.GeneralNoteResults.Any())
                    {
                        RenderCheckResultsToHtml(
                            task.GeneralNoteResults.Cast<dynamic>().ToList(),
                            "general notes",
                            htmlBuilder);
                    }
                    else if (task.LogicType == "Dedication" && task.DedicationResults.Any())
                    {
                        RenderCheckResultsToHtml(
                            task.DedicationResults.Cast<dynamic>().ToList(),
                            "dedication",
                            htmlBuilder);
                    }
                    else if (task.LogicType == "TitleBlockDetailSection")
                    {
                        // For TitleBlockDetailSection, show all attributes (each field summary)
                        foreach (var attr in task.Attributes.Where(a => a.AttributeName != "DetailedReportPath" &&
                                                                         a.AttributeName != "PageValidationResults"))
                        {
                            if (!string.IsNullOrEmpty(attr.Value))
                            {
                                // Show each field's summary (e.g., "Found on 50/50 pages.\nMissing Control Number on pages: 3, 7")
                                var lines = attr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                htmlBuilder.AppendLine($"<strong>{HttpUtility.HtmlEncode(attr.AttributeName)}:</strong><br>");
                                foreach (var line in lines)
                                {
                                    htmlBuilder.AppendLine($"{HttpUtility.HtmlEncode(line)}<br>");
                                }
                                htmlBuilder.AppendLine("<br>");
                            }
                        }
                    }
                    else if (task.Status == "Mismatch" && task.Attributes.Count >= 2)
                    {
                        var attr1 = task.Attributes[0];
                        var attr2 = task.Attributes[1];
                        htmlBuilder.AppendLine($"<strong>{HttpUtility.HtmlEncode(attr1.FileCategory)}:</strong> {HttpUtility.HtmlEncode(attr1.Value ?? "N/A")}<br>");
                        htmlBuilder.AppendLine($"<strong>{HttpUtility.HtmlEncode(attr2.FileCategory)}:</strong> {HttpUtility.HtmlEncode(attr2.Value ?? "N/A")}<br>");
                        htmlBuilder.AppendLine("<strong>Explanation:</strong> The values don't match in the documents.");
                    }
                    else if (run.PlanReviewType == "Land" && (task.LogicType == "CheckValuePerPage" || task.LogicType == "SheetIndex" || task.TaskName == "Client" || task.TaskName == "Scale"))
                    {
                        // Only show the first line of the first attribute's value (the summary)
                        var landAttr = task.Attributes.FirstOrDefault(a => a.AttributeName != "DetailedReportPath" && a.AttributeName != "PageValidationResults");
                        if (landAttr != null && !string.IsNullOrEmpty(landAttr.Value))
                        {
                            var firstLine = landAttr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                            htmlBuilder.AppendLine($"{HttpUtility.HtmlEncode(firstLine)}<br>");
                        }
                    }
                    else if (run.PlanReviewType == "Transportation" && task.LogicType == "CheckValuePerPage")
                    {
                        // Only show the first line of the first attribute's value (the summary)
                        var transportationCheckValueAttr = task.Attributes.FirstOrDefault(a => a.AttributeName != "DetailedReportPath" && a.AttributeName != "PageValidationResults");
                        if (transportationCheckValueAttr != null && !string.IsNullOrEmpty(transportationCheckValueAttr.Value))
                        {
                            var firstLine = transportationCheckValueAttr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                            htmlBuilder.AppendLine($"{HttpUtility.HtmlEncode(firstLine)}<br>");
                        }
                    }
                    else if (run.PlanReviewType == "Transportation" && task.LogicType == "SheetIndex")
                    {
                        // Show all lines for SheetIndex to include mismatch page information
                        var transportationSheetIndexAttr = task.Attributes.FirstOrDefault(a => a.AttributeName != "DetailedReportPath" && a.AttributeName != "PageValidationResults");
                        if (transportationSheetIndexAttr != null && !string.IsNullOrEmpty(transportationSheetIndexAttr.Value))
                        {
                            var lines = transportationSheetIndexAttr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var line in lines)
                            {
                                // Color-code lines containing "mismatch" - MATCHING PDF LOGIC
                                string lineColor = line.ToLower().Contains("mismatch") ? "color: #e74c3c;" : "color: #333;";
                                htmlBuilder.AppendLine($"<span style='{lineColor}'>{HttpUtility.HtmlEncode(line)}</span><br>");
                            }
                        }
                    }
                    else if (task.StatusDescription != null)
                    {
                        htmlBuilder.AppendLine(HttpUtility.HtmlEncode(task.StatusDescription));
                    }
                    else if (task.Status == "Notice")
                    {
                        htmlBuilder.AppendLine("Value is missing from the documents.");
                    }
                    else if (task.Status == "Error")
                    {
                        htmlBuilder.AppendLine("One or more required values have an error.");
                    }
                    else if (task.Status == "Not enough data")
                    {
                        htmlBuilder.AppendLine("Insufficient data to complete the comparison.");
                    }
                    else if (task.Status == "Match")
                    {
                        htmlBuilder.AppendLine($"Value matches in the documents: {HttpUtility.HtmlEncode(task.Attributes[0].Value ?? "N/A")}");
                    }
                    else if (task.Status == "Correct")
                    {
                        htmlBuilder.AppendLine($"Value: {HttpUtility.HtmlEncode(task.Attributes[0].Value ?? "N/A")}");
                    }
                    else
                    {
                        foreach (var attr in task.Attributes.Where(a => a.AttributeName != "DetailedReportPath"))
                        {
                            htmlBuilder.AppendLine($"<strong>{HttpUtility.HtmlEncode(attr.FileCategory)}:</strong> {HttpUtility.HtmlEncode(attr.Value ?? "N/A")}<br>");
                        }
                    }
                }
                else
                {
                    htmlBuilder.AppendLine("No attributes available");
                }
                htmlBuilder.AppendLine("</td>");
                htmlBuilder.AppendLine("</tr>");

                isEvenRow = !isEvenRow;
            }
            htmlBuilder.AppendLine("</table>");

            // Recommendations section
            if (mismatchCount > 0 || noticeCount > 0 || errorCount > 0)
            {
                htmlBuilder.AppendLine("<h2 style=\"color: #2c3e50; margin-top: 20px;\">Recommendations</h2>");
                htmlBuilder.AppendLine("<div style=\"margin-bottom: 30px;\">");
                htmlBuilder.AppendLine("<p>Based on the results of this plan review, the following recommendations are provided:</p>");
                htmlBuilder.AppendLine("<ul>");

                if (mismatchCount > 0)
                {
                    htmlBuilder.AppendLine("<li>Review and reconcile the mismatched values between documents.</li>");
                }

                if (noticeCount > 0)
                {
                    htmlBuilder.AppendLine("<li>Ensure all required information is present in the documents.</li>");
                }

                if (errorCount > 0)
                {
                    htmlBuilder.AppendLine("<li>Investigate and correct any errors found in the extracted data.</li>");
                }

                htmlBuilder.AppendLine("<li>Consider re-running the plan review after making necessary corrections.</li>");
                htmlBuilder.AppendLine("</ul>");
                htmlBuilder.AppendLine("</div>");
            }

            // Add note about PDF attachment
            var perPageTasks = tasks.Where(t => t.LogicType == "CheckValuePerPage").ToList();
            var sheetIndexTasks = tasks.Where(t => t.LogicType == "SheetIndex" && (t.TaskName == "Sheet Number & Title" || t.TaskName == "Sheet Index & Sheet Title")).ToList();

            if (perPageTasks.Any() || sheetIndexTasks.Any())
            {
                htmlBuilder.AppendLine("<div style=\"margin-top: 30px; padding: 15px; background-color: #f0f7fb; border-left: 5px solid #3498db;\">");
                htmlBuilder.AppendLine("<h3 style='margin-top: 0;'>Detailed Results Attachment</h3>");
                htmlBuilder.AppendLine("<p>Please find attached a PDF file with detailed results for:</p>");
                htmlBuilder.AppendLine("<ul>");

                if (perPageTasks.Any())
                {
                    htmlBuilder.AppendLine("<li>Per-page validation results</li>");
                }

                if (sheetIndexTasks.Any())
                {
                    htmlBuilder.AppendLine("<li>Sheet index validation results</li>");
                }

                htmlBuilder.AppendLine("</ul>");
                htmlBuilder.AppendLine("<p>The PDF file contains additional details that can be reviewed for a more comprehensive analysis.</p>");
                htmlBuilder.AppendLine("</div>");
            }

            // Email footer
            htmlBuilder.AppendLine("<p style=\"margin-top: 30px; font-size: 12px; color: #777; border-top: 1px solid #ddd; padding-top: 10px;\">");
            htmlBuilder.AppendLine("This is an automated message from the Plan Review System. Please do not reply to this email.<br>");
            htmlBuilder.AppendLine("For assistance, please contact your system administrator.");
            htmlBuilder.AppendLine("</p>");

            htmlBuilder.AppendLine("</body>");
            htmlBuilder.AppendLine("</html>");

            return htmlBuilder.ToString();
        }

        // New method to generate PDF report
        private async Task<string> GeneratePdfReportAsync(int runId, List<PlanReviewTask> tasks)
        {
            logFile.MyLogFile($"Generating PDF for {tasks.Count} tasks. Status counts: " +
                $"Match={tasks.Count(t => t.Status == "Match")}, " +
                $"Correct={tasks.Count(t => t.Status == "Correct")}, " +
                $"Mismatch={tasks.Count(t => t.Status == "Mismatch")}, " +
                $"Error={tasks.Count(t => t.Status == "Error")}, " +
                $"Notice={tasks.Count(t => t.Status == "Notice" || t.Status == "Not enough data")}");

            var run = await GetPlanReviewRunAsync(runId);
            if (run == null)
            {
                logFile.MyLogFile($"Error: Plan Review run not found for runId {runId}.");
                return null;
            }

            // Sort tasks: GeneralNotes and Dedication at the end, then by TaskName
            var orderedTasks = tasks
                .OrderBy(t => t.LogicType == "GeneralNotes" || t.LogicType == "Dedication" ? 1 : 0)
                .ThenBy(t => t.TaskName)
                .ToList();

            // Generate filename based on plan review type
            string filename;
            if (run.PlanReviewType == "Platting")
            {
                // Get Plat Name from the task list
                var platNameTask = tasks.FirstOrDefault(t => t.TaskName == "Plat Name");
                string platName = "Unknown";

                if (platNameTask != null && platNameTask.Attributes.Any())
                {
                    var platNameAttr = platNameTask.Attributes.FirstOrDefault(a => !string.IsNullOrEmpty(a.Value));
                    if (platNameAttr != null && !string.IsNullOrEmpty(platNameAttr.Value))
                    {
                        platName = platNameAttr.Value;
                        // Clean the plat name for use in filename (remove invalid filename characters)
                        platName = System.Text.RegularExpressions.Regex.Replace(platName, @"[^\w\s-]", "");
                        platName = platName.Trim().Replace(" ", "_");
                    }
                }

                filename = $"{platName}_Pre-QAQC Review Results_{DateTime.Now:yyyyMMdd}.pdf";
            }
            else
            {
                filename = $"PlanReviewReport_{runId}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            }

            string outputPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), filename);
            logFile.MyLogFile($"Generating PDF report to: {outputPath}");

            try
            {
                if (!File.Exists(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "itextsharp.dll")))
                {
                    logFile.MyLogFile("Warning: iTextSharp DLL not found in application directory.");
                }

                using (var document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 50, 50))
                {
                    iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(outputPath, FileMode.Create));
                    document.Open();

                    iTextSharp.text.Font titleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 18, iTextSharp.text.Font.BOLD);
                    iTextSharp.text.Font subtitleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 14, iTextSharp.text.Font.BOLD);
                    iTextSharp.text.Font normalFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10);
                    iTextSharp.text.Font boldFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD);
                    iTextSharp.text.Font smallFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8);
                    iTextSharp.text.Font greenFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, new iTextSharp.text.BaseColor(39, 174, 96));
                    iTextSharp.text.Font redFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, new iTextSharp.text.BaseColor(231, 76, 60));
                    iTextSharp.text.Font orangeFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, new iTextSharp.text.BaseColor(243, 156, 18));

                    // Title
                    if (run.PlanReviewType == "Platting")
                    {
                        // Get Plat Name from the task list
                        var platNameTask = tasks.FirstOrDefault(t => t.TaskName == "Plat Name");
                        string platName = "Unknown";

                        if (platNameTask != null && platNameTask.Attributes.Any())
                        {
                            var platNameAttr = platNameTask.Attributes.FirstOrDefault(a => !string.IsNullOrEmpty(a.Value));
                            if (platNameAttr != null && !string.IsNullOrEmpty(platNameAttr.Value))
                            {
                                platName = platNameAttr.Value;
                                // Clean the plat name for display
                                platName = System.Text.RegularExpressions.Regex.Replace(platName, @"[^\w\s-]", "");
                                platName = platName.Trim();
                            }
                        }

                        document.Add(new iTextSharp.text.Paragraph($"{platName} - Pre-QAQC Review Results", titleFont));
                    }
                    else
                    {
                        document.Add(new iTextSharp.text.Paragraph($"Plan Review Results - Run #{run.ID}", titleFont));
                    }
                    document.Add(new iTextSharp.text.Paragraph(" "));

                    // Summary Information
                    document.Add(new iTextSharp.text.Paragraph("Summary Information", subtitleFont));
                    document.Add(new iTextSharp.text.Paragraph(" "));

                    var summaryTable = new iTextSharp.text.pdf.PdfPTable(2) { WidthPercentage = 100 };
                    summaryTable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    summaryTable.AddCell(new iTextSharp.text.Phrase("Run ID:", boldFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase(run.ID.ToString(), normalFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase("Plan Review Type:", boldFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase(run.PlanReviewType, normalFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase("Task Set:", boldFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase(run.TaskSet, normalFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase("Created By:", boldFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase(run.CreatedBy, normalFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase("Date:", boldFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase(run.Timestamp.ToString("MMMM dd, yyyy HH:mm tt"), normalFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase("Status:", boldFont));
                    summaryTable.AddCell(new iTextSharp.text.Phrase(run.Status, normalFont));

                    document.Add(summaryTable);
                    document.Add(new iTextSharp.text.Paragraph(" "));

                    int matchCount = orderedTasks.Count(t => t.Status == "Match");
                    int correctCount = orderedTasks.Count(t => t.Status == "Correct");
                    int mismatchCount = orderedTasks.Count(t => t.Status == "Mismatch");
                    int errorCount = orderedTasks.Count(t => t.Status == "Error");
                    int noticeCount = orderedTasks.Count(t => t.Status == "Notice" || t.Status == "Not enough data");

                    document.Add(new iTextSharp.text.Paragraph("Task Status Summary:", boldFont));
                    document.Add(new iTextSharp.text.Paragraph($"• Matches/Correct: {matchCount + correctCount}", greenFont));
                    document.Add(new iTextSharp.text.Paragraph($"• Mismatches/Errors: {mismatchCount + errorCount}", redFont));
                    document.Add(new iTextSharp.text.Paragraph($"• Notice: {noticeCount}", orangeFont));
                    document.Add(new iTextSharp.text.Paragraph(" "));

                    var detailsTable = new iTextSharp.text.pdf.PdfPTable(3) { WidthPercentage = 100 };
                    float[] columnWidths = new float[] { 30f, 15f, 55f };
                    detailsTable.SetWidths(columnWidths);

                    // Table headers
                    var headerCell1 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Task Name", boldFont));
                    var headerCell2 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Status", boldFont));
                    var headerCell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Details", boldFont));
                    headerCell1.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                    headerCell2.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                    headerCell3.BackgroundColor = new iTextSharp.text.BaseColor(242, 242, 242);
                    headerCell1.Padding = 5;
                    headerCell2.Padding = 5;
                    headerCell3.Padding = 5;
                    detailsTable.AddCell(headerCell1);
                    detailsTable.AddCell(headerCell2);
                    detailsTable.AddCell(headerCell3);

                    bool isAlternateRow = false;
                    foreach (var task in orderedTasks)
                    {
                        var cell1 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(task.TaskName ?? "Unknown Task", normalFont));

                        iTextSharp.text.Font statusFont;
                        switch (task.Status?.ToLower())
                        {
                            case "match":
                            case "correct":
                                statusFont = greenFont;
                                break;
                            case "mismatch":
                            case "error":
                                statusFont = redFont;
                                break;
                            case "notice":
                            case "not enough data":
                            case "partial match":
                                statusFont = orangeFont;
                                break;
                            default:
                                statusFont = normalFont;
                                break;
                        }
                        var cell2 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(task.Status ?? "N/A", statusFont));

                        // --- Details column logic - NOW WITH COLORED STATUS ONLY (NO ICONS) ---
                        iTextSharp.text.pdf.PdfPCell cell3;

                        if (task.Attributes != null && task.Attributes.Count > 0)
                        {
                            if (task.LogicType == "GeneralNotes" && task.GeneralNoteResults.Any())
                            {
                                // Create a Paragraph to hold multiple colored phrases
                                var detailsParagraph = new iTextSharp.text.Paragraph();

                                foreach (var noteResult in task.GeneralNoteResults)
                                {
                                    // Determine the status word and its color
                                    iTextSharp.text.Font statusFont2 = normalFont;
                                    string statusWord = "";

                                    if (noteResult.Result.Contains("Correct"))
                                    {
                                        statusWord = "Correct";
                                        statusFont2 = greenFont;
                                    }
                                    else if (noteResult.Result.Contains("Missing"))
                                    {
                                        statusWord = "Missing";
                                        statusFont2 = redFont;
                                    }
                                    else if (noteResult.Result.Contains("Incorrect"))
                                    {
                                        statusWord = "Incorrect";
                                        statusFont2 = redFont;
                                    }
                                    else if (noteResult.Result.Contains("Orphaned"))
                                    {
                                        statusWord = "Orphaned";
                                        statusFont2 = orangeFont;
                                    }
                                    else if (noteResult.Result.Contains("Error"))
                                    {
                                        statusWord = "Error";
                                        statusFont2 = redFont;
                                    }
                                    else
                                    {
                                        statusWord = "Unhandled";
                                        statusFont2 = orangeFont;
                                    }

                                    // Add colored status word, then task name in normal font
                                    detailsParagraph.Add(new iTextSharp.text.Phrase(statusWord, statusFont2));
                                    detailsParagraph.Add(new iTextSharp.text.Phrase($" - {noteResult.Task}\n", normalFont));
                                    // Add details in normal font
                                    detailsParagraph.Add(new iTextSharp.text.Phrase($"{noteResult.Details}\n\n", normalFont));
                                }

                                cell3 = new iTextSharp.text.pdf.PdfPCell(detailsParagraph);
                            }
                            else if (task.LogicType == "Dedication" && task.DedicationResults.Any())
                            {
                                // Create a Paragraph to hold multiple colored phrases
                                var detailsParagraph = new iTextSharp.text.Paragraph();

                                foreach (var dedicationResult in task.DedicationResults)
                                {
                                    // Determine the status word and its color
                                    iTextSharp.text.Font statusFont2 = normalFont;
                                    string statusWord = "";

                                    if (dedicationResult.Result.Contains("Correct"))
                                    {
                                        statusWord = "Correct";
                                        statusFont2 = greenFont;
                                    }
                                    else if (dedicationResult.Result.Contains("Match"))
                                    {
                                        statusWord = "Match";
                                        statusFont2 = greenFont;
                                    }
                                    else if (dedicationResult.Result.Contains("Missing"))
                                    {
                                        statusWord = "Missing";
                                        statusFont2 = redFont;
                                    }
                                    else if (dedicationResult.Result.Contains("Incorrect"))
                                    {
                                        statusWord = "Incorrect";
                                        statusFont2 = redFont;
                                    }
                                    else if (dedicationResult.Result.Contains("Mismatch"))
                                    {
                                        statusWord = "Mismatch";
                                        statusFont2 = redFont;
                                    }
                                    else if (dedicationResult.Result.Contains("Error"))
                                    {
                                        statusWord = "Error";
                                        statusFont2 = redFont;
                                    }
                                    else
                                    {
                                        statusWord = "Unhandled";
                                        statusFont2 = orangeFont;
                                    }

                                    // Add colored status word, then task name in normal font
                                    detailsParagraph.Add(new iTextSharp.text.Phrase(statusWord, statusFont2));
                                    detailsParagraph.Add(new iTextSharp.text.Phrase($" - {dedicationResult.Task}\n", normalFont));
                                    // Add details in normal font
                                    detailsParagraph.Add(new iTextSharp.text.Phrase($"{dedicationResult.Details}\n\n", normalFont));
                                }

                                cell3 = new iTextSharp.text.pdf.PdfPCell(detailsParagraph);
                            }
                            else if (task.LogicType == "TitleBlockDetailSection")  // ADD THIS BLOCK
                            {
                                // Create a Paragraph to hold field summaries
                                var detailsParagraph = new iTextSharp.text.Paragraph();

                                foreach (var attr in task.Attributes.Where(a => a.AttributeName != "DetailedReportPath" &&
                                                                                 a.AttributeName != "PageValidationResults"))
                                {
                                    if (!string.IsNullOrEmpty(attr.Value))
                                    {
                                        // Add field name in bold
                                        detailsParagraph.Add(new iTextSharp.text.Phrase($"{attr.AttributeName}:\n", boldFont));

                                        // Split multi-line summary and add each line
                                        var lines = attr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                        foreach (var line in lines)
                                        {
                                            // Color-code based on content
                                            iTextSharp.text.Font lineFont = normalFont;
                                            if (line.Contains("Missing"))
                                            {
                                                lineFont = redFont;
                                            }
                                            else if (line.Contains("Found on"))
                                            {
                                                lineFont = normalFont; // Summary line stays normal
                                            }

                                            detailsParagraph.Add(new iTextSharp.text.Phrase($"{line}\n", lineFont));
                                        }
                                        detailsParagraph.Add(new iTextSharp.text.Phrase("\n", normalFont)); // Extra spacing
                                    }
                                }

                                cell3 = new iTextSharp.text.pdf.PdfPCell(detailsParagraph);
                            }
                            else if (task.Status == "Mismatch" && task.Attributes.Count >= 2)
                            {
                                var attr1 = task.Attributes[0];
                                var attr2 = task.Attributes[1];
                                var detailsText = $"{attr1.FileCategory}: {attr1.Value ?? "N/A"}\n{attr2.FileCategory}: {attr2.Value ?? "N/A"}\nExplanation: The values don't match between documents.";
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(detailsText, normalFont));
                            }
                            else if (run.PlanReviewType == "Land" && (task.LogicType == "CheckValuePerPage" || task.LogicType == "SheetIndex" || task.TaskName == "Client" || task.TaskName == "Scale"))
                            {
                                var landPdfAttr = task.Attributes.FirstOrDefault(a => a.AttributeName != "DetailedReportPath" && a.AttributeName != "PageValidationResults");
                                if (landPdfAttr != null && !string.IsNullOrEmpty(landPdfAttr.Value))
                                {
                                    var firstLine = landPdfAttr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                                    cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(firstLine, normalFont));
                                }
                                else
                                {
                                    cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("No details available", normalFont));
                                }
                            }
                            else if (run.PlanReviewType == "Transportation" && task.LogicType == "CheckValuePerPage")
                            {
                                var transportationCheckValuePdfAttr = task.Attributes.FirstOrDefault(a => a.AttributeName != "DetailedReportPath" && a.AttributeName != "PageValidationResults");
                                if (transportationCheckValuePdfAttr != null && !string.IsNullOrEmpty(transportationCheckValuePdfAttr.Value))
                                {
                                    var firstLine = transportationCheckValuePdfAttr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                                    cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(firstLine, normalFont));
                                }
                                else
                                {
                                    cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("No details available", normalFont));
                                }
                            }
                            else if (run.PlanReviewType == "Transportation" && task.LogicType == "SheetIndex")
                            {
                                // Show all lines for SheetIndex to include mismatch page information
                                var transportationSheetIndexPdfAttr = task.Attributes.FirstOrDefault(a => a.AttributeName != "DetailedReportPath" && a.AttributeName != "PageValidationResults");
                                if (transportationSheetIndexPdfAttr != null && !string.IsNullOrEmpty(transportationSheetIndexPdfAttr.Value))
                                {
                                    var detailsParagraph = new iTextSharp.text.Paragraph();
                                    var lines = transportationSheetIndexPdfAttr.Value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var line in lines)
                                    {
                                        // Color-code lines containing "mismatch"
                                        iTextSharp.text.Font lineFont = line.ToLower().Contains("mismatch") ? redFont : normalFont;
                                        detailsParagraph.Add(new iTextSharp.text.Phrase($"{line}\n", lineFont));
                                    }
                                    cell3 = new iTextSharp.text.pdf.PdfPCell(detailsParagraph);
                                }
                                else
                                {
                                    cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("No details available", normalFont));
                                }
                            }
                            else if (task.StatusDescription != null)
                            {
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(task.StatusDescription, normalFont));
                            }
                            else if (task.Status == "Notice")
                            {
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Value is missing from the documents.", normalFont));
                            }
                            else if (task.Status == "Error")
                            {
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("One or more required values have an error.", normalFont));
                            }
                            else if (task.Status == "Not enough data")
                            {
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("Insufficient data to complete the comparison.", normalFont));
                            }
                            else if (task.Status == "Match")
                            {
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase($"Value matches in the documents: {task.Attributes[0].Value ?? "N/A"}", normalFont));
                            }
                            else if (task.Status == "Correct")
                            {
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase($"Value: {task.Attributes[0].Value ?? "N/A"}", normalFont));
                            }
                            else
                            {
                                var detailsText = new StringBuilder();
                                foreach (var attr in task.Attributes.Where(a => a.AttributeName != "DetailedReportPath"))
                                {
                                    detailsText.AppendLine($"{attr.FileCategory}: {attr.Value ?? "N/A"}");
                                }
                                cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(detailsText.ToString(), normalFont));
                            }
                        }
                        else
                        {
                            cell3 = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase("No attributes available", normalFont));
                        }
                        // --- End details column logic ---

                        if (isAlternateRow)
                        {
                            cell1.BackgroundColor = new iTextSharp.text.BaseColor(249, 249, 249);
                            cell2.BackgroundColor = new iTextSharp.text.BaseColor(249, 249, 249);
                            cell3.BackgroundColor = new iTextSharp.text.BaseColor(249, 249, 249);
                        }
                        cell1.Padding = 5;
                        cell2.Padding = 5;
                        cell3.Padding = 5;

                        detailsTable.AddCell(cell1);
                        detailsTable.AddCell(cell2);
                        detailsTable.AddCell(cell3);

                        isAlternateRow = !isAlternateRow;
                    }

                    document.Add(detailsTable);

                    // Recommendations section
                    if (mismatchCount > 0 || noticeCount > 0 || errorCount > 0)
                    {
                        document.Add(new iTextSharp.text.Paragraph("Recommendations", subtitleFont));
                        document.Add(new iTextSharp.text.Paragraph(" "));
                        document.Add(new iTextSharp.text.Paragraph("Based on the results of this plan review, the following recommendations are provided:", normalFont));
                        document.Add(new iTextSharp.text.Paragraph(" "));

                        if (mismatchCount > 0)
                            document.Add(new iTextSharp.text.Paragraph("• Review and reconcile the mismatched values between documents.", normalFont));
                        if (noticeCount > 0)
                            document.Add(new iTextSharp.text.Paragraph("• Ensure all required information is present in the documents.", normalFont));
                        if (errorCount > 0)
                            document.Add(new iTextSharp.text.Paragraph("• Investigate and correct any errors found in the extracted data.", normalFont));
                        document.Add(new iTextSharp.text.Paragraph("• Consider re-running the plan review after making necessary corrections.", normalFont));
                        document.Add(new iTextSharp.text.Paragraph(" "));
                    }

                    // Per-Page Validation Details
                    var perPageTasks = orderedTasks.Where(t => t.LogicType == "CheckValuePerPage").ToList();
                    if (perPageTasks.Any())
                    {
                        document.Add(new iTextSharp.text.Paragraph("Per-Page Validation Details", subtitleFont));
                        document.Add(new iTextSharp.text.Paragraph(" "));

                        foreach (var perPageTask in perPageTasks)
                        {
                            document.Add(new iTextSharp.text.Paragraph($"{perPageTask.TaskName} Validation", boldFont));
                            var summaryAttribute = perPageTask.Attributes?.FirstOrDefault();
                            if (summaryAttribute != null && !string.IsNullOrEmpty(summaryAttribute.Value))
                                document.Add(new iTextSharp.text.Paragraph(summaryAttribute.Value, normalFont));

                            var reportPathAttr = perPageTask.Attributes?.FirstOrDefault(a => a.AttributeName == "DetailedReportPath");
                            if (reportPathAttr != null && File.Exists(reportPathAttr.Value))
                            {
                                try
                                {
                                    string htmlContent = File.ReadAllText(reportPathAttr.Value);
                                    var missingPages = new List<int>();
                                    var invalidPages = new List<int>();
                                    var validPages = new List<int>();
                                    var regex = new System.Text.RegularExpressions.Regex(
                                        @"<tr.*?>\s*<td>(\d+)</td>\s*<td>(True|False)</td>\s*<td>(True|False)</td>",
                                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                                    var matches = regex.Matches(htmlContent);

                                    foreach (System.Text.RegularExpressions.Match match in matches)
                                    {
                                        if (match.Groups.Count >= 4)
                                        {
                                            int pageNumber = int.Parse(match.Groups[1].Value);
                                            bool valuePresent = match.Groups[2].Value.Equals("True", StringComparison.OrdinalIgnoreCase);
                                            bool valueValid = match.Groups[3].Value.Equals("True", StringComparison.OrdinalIgnoreCase);

                                            if (!valuePresent)
                                                missingPages.Add(pageNumber);
                                            else if (!valueValid)
                                                invalidPages.Add(pageNumber);
                                            else
                                                validPages.Add(pageNumber);
                                        }
                                    }

                                    if (validPages.Any())
                                    {
                                        document.Add(new iTextSharp.text.Paragraph("Valid on pages:", boldFont));
                                        document.Add(new iTextSharp.text.Paragraph(FormatPageNumbers(validPages), greenFont));
                                    }
                                    if (invalidPages.Any())
                                    {
                                        document.Add(new iTextSharp.text.Paragraph("Invalid values on pages:", boldFont));
                                        document.Add(new iTextSharp.text.Paragraph(FormatPageNumbers(invalidPages), redFont));
                                    }
                                    if (missingPages.Any())
                                    {
                                        document.Add(new iTextSharp.text.Paragraph("Missing values on pages:", boldFont));
                                        document.Add(new iTextSharp.text.Paragraph(FormatPageNumbers(missingPages), orangeFont));
                                    }
                                }
                                catch (Exception ex)
                                {
                                    logFile.MyLogFile($"Error processing detailed HTML report for task {perPageTask.TaskName}: {ex.Message}");
                                    document.Add(new iTextSharp.text.Paragraph($"Error processing detailed report: {ex.Message}", redFont));
                                }
                            }
                            document.Add(new iTextSharp.text.Paragraph(" "));
                        }
                    }

                    // Sheet Index Validation Details
                    var sheetIndexTasks = orderedTasks.Where(t => t.LogicType == "SheetIndex").ToList();
                    if (sheetIndexTasks.Any())
                    {
                        document.Add(new iTextSharp.text.Paragraph("Sheet Index Validation", subtitleFont));
                        document.Add(new iTextSharp.text.Paragraph(" "));

                        foreach (var sheetTask in sheetIndexTasks)
                        {
                            document.Add(new iTextSharp.text.Paragraph($"{sheetTask.TaskName}", boldFont));
                            var detailsAttribute = sheetTask.Attributes?.FirstOrDefault();
                            if (detailsAttribute != null && !string.IsNullOrEmpty(detailsAttribute.Value))
                            {
                                document.Add(new iTextSharp.text.Paragraph(detailsAttribute.Value, normalFont));
                                string summaryText = detailsAttribute.Value;
                                var matchRegex = new System.Text.RegularExpressions.Regex(@"(\d+) out of (\d+) pages matched");
                                var match = matchRegex.Match(summaryText);

                                if (match.Success && match.Groups.Count >= 3)
                                {
                                    int matchedPages = int.Parse(match.Groups[1].Value);
                                    int totalPages = int.Parse(match.Groups[2].Value);

                                    if (matchedPages < totalPages)
                                    {
                                        document.Add(new iTextSharp.text.Paragraph("Sheet index discrepancies found. The following pages may have incorrect or missing sheet numbers/titles:", normalFont));
                                        var unmatchedPages = Enumerable.Range(1, totalPages).Except(Enumerable.Range(1, matchedPages)).ToList();
                                        document.Add(new iTextSharp.text.Paragraph($"Pages with potential issues: {FormatPageNumbers(unmatchedPages)}", orangeFont));
                                    }
                                }
                            }
                            document.Add(new iTextSharp.text.Paragraph(" "));
                        }
                    }

                    document.Add(new iTextSharp.text.Paragraph(" "));
                    document.Add(new iTextSharp.text.Paragraph("This is an automated report from the Plan Review System.", smallFont));
                    document.Add(new iTextSharp.text.Paragraph($"Generated on: {DateTime.Now:MMMM dd, yyyy HH:mm tt}", smallFont));

                    document.Close();
                }

                logFile.MyLogFile($"PDF report successfully generated: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                logFile.MyLogFile($"Error generating PDF report: {ex.Message}");
                logFile.MyLogFile($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        // Helper method to format lists of page numbers in a readable way
        private string FormatPageNumbers(List<int> pageNumbers)
        {
            if (pageNumbers == null || !pageNumbers.Any())
            {
                return "None";
            }

            // Sort page numbers
            pageNumbers.Sort();

            // Group consecutive page numbers
            var groups = new List<(int Start, int End)>();
            int start = pageNumbers[0];
            int end = start;

            for (int i = 1; i < pageNumbers.Count; i++)
            {
                if (pageNumbers[i] == end + 1)
                {
                    end = pageNumbers[i];
                }
                else
                {
                    groups.Add((start, end));
                    start = pageNumbers[i];
                    end = start;
                }
            }

            groups.Add((start, end));

            // Format the groups
            var parts = new List<string>();
            foreach (var (groupStart, groupEnd) in groups)
            {
                if (groupStart == groupEnd)
                {
                    parts.Add(groupStart.ToString());
                }
                else
                {
                    parts.Add($"{groupStart}-{groupEnd}");
                }
            }

            return string.Join(", ", parts);
        }

        // Helper methods for rendering check results
        private void RenderCheckResultsToHtml(
            List<dynamic> results,
            string resultTypeName,
            StringBuilder htmlBuilder)
        {
            if (results == null || results.Count == 0)
            {
                htmlBuilder.AppendLine($"All {resultTypeName} checks passed.");
                return;
            }

            foreach (var result in results)
            {
                string statusIcon = GetStatusIcon(result.Result);
                string statusColor = GetStatusColor(result.Result);

                htmlBuilder.AppendLine($"<span style='{statusColor}'>{statusIcon}</span> - {HttpUtility.HtmlEncode(result.Task)}<br>");
                htmlBuilder.AppendLine($"<em style='font-size: 12px;'>{HttpUtility.HtmlEncode(result.Details)}</em><br><br>");
            }
        }

        private string GetStatusIcon(string result)
        {
            if (result.Contains("Match")) return "✓ Match";
            if (result.Contains("Correct")) return "✓ Correct";
            if (result.Contains("Missing")) return "⚠ Missing";
            if (result.Contains("Incorrect")) return "✗ Incorrect";
            if (result.Contains("Mismatch")) return "✗ Mismatch";
            if (result.Contains("Error")) return "✗ Error";
            if (result.Contains("Orphaned")) return "⚠ Orphaned";
            return "⚠ Unhandled";
        }

        private string GetStatusColor(string result)
        {
            if (result.Contains("Correct") || result.Contains("Match")) return "color: #27ae60;";
            if (result.Contains("Missing") || result.Contains("Incorrect") || result.Contains("Error") || result.Contains("Mismatch"))
                return "color: #e74c3c;";
            if (result.Contains("Orphaned")) return "color: #f39c12;";
            return "color: #f39c12;";
        }
        #endregion

    }

    #region Models
    public class PlanReviewDocument
    {
        public int DocumentId { get; set; }
        public string FileCategory { get; set; }
        public string FileName { get; set; }

        public PlanReviewDocument(int documentId, string fileCategory, string fileName)
        {
            DocumentId = documentId;
            FileCategory = fileCategory;
            FileName = fileName;
        }
    }

    public class PlanReviewResponse
    {
        public string Status { get; set; }
        public string Message { get; set; }
        public List<PlanReviewTask> Tasks { get; set; } = new List<PlanReviewTask>();
    }

    public class PlanReviewRun
    {
        public int ID { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedByEmail { get; set; }
        public DateTime Timestamp { get; set; }
        public string Status { get; set; }
        public string PlanReviewType { get; set; }
        public string TaskSet { get; set; }
        public string ProjectNumber { get; set; }
    }

    public class PlanReviewTask
    {
        public int TaskId { get; set; }
        public string PlanReviewType { get; set; }
        public string TaskName { get; set; }
        public string TaskSet { get; set; }
        public string LogicType { get; set; }
        public List<PlanReviewAttribute> Attributes { get; set; } = new List<PlanReviewAttribute>();
        public string Status { get; set; }
        public string StatusDescription { get; set; }
        public List<int> DocumentIds { get; set; } = new List<int>();
        public Dictionary<string, int> DocumentIdsByCategory { get; set; } = new Dictionary<string, int>();

        public List<GeneralNoteCheckResult> GeneralNoteResults { get; set; } = new List<GeneralNoteCheckResult>();
        public List<DedicationCheckResult> DedicationResults { get; set; } = new List<DedicationCheckResult>();
    }

    public class PlanReviewAttribute
    {
        public int AttributeId { get; set; }
        public string AttributeName { get; set; }
        public string FileCategory { get; set; }
        public int AttributeCode { get; set; }
        public string Value { get; set; }
        public string PageCategory { get; set; }
        public string RelevantPageCategories { get; set; }
    }

    public class PageValidationResult
    {
        public int PageNumber { get; set; }
        public bool ValuePresent { get; set; }
        public bool ValueValid { get; set; }
        public string ValueDetails { get; set; }
    }
    public class PlanReviewGeneralNote
    {
        public int GeneralNoteID { get; set; }
        public string Task { get; set; }
        public string LogicCode { get; set; }
        private string _searchTermRaw;
        public string SearchTermRaw
        {
            get { return _searchTermRaw; }
            set
            {
                _searchTermRaw = value;
                // Parse comma-separated values into list
                SearchTerms = string.IsNullOrWhiteSpace(value)
                    ? new List<string>()
                    : value.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Trim())
                        .ToList();
            }
        }

        public List<string> SearchTerms { get; set; } = new List<string>();
        public string NoteTerm { get; set; }
        public string LogicDescription { get; set; }
        public string Jurisdiction { get; set; }
    }
    public class GeneralNoteCheckResult
    {
        public int GeneralNoteID { get; set; }
        public string Task { get; set; }
        public string LogicCode { get; set; }
        public bool Found { get; set; }
        public string Details { get; set; }
        public string Result { get; set; }
    }
    public class PlanReviewDedication
    {
        public string Task { get; set; }
        public string LogicCode { get; set; }
        public string SearchTerm { get; set; }
        public string NoteTerm { get; set; }
        public string LogicDescription { get; set; }
        public string Jurisdiction { get; set; }
    }
    public class DedicationCheckResult
    {
        public string Task { get; set; }
        public string LogicCode { get; set; }
        public bool Found { get; set; }
        public string Details { get; set; }
        public string Result { get; set; }
    }
    public class TitleBlock
    {
        public decimal? Acreage { get; set; }
        public int? Blocks { get; set; }
        public int? Lots { get; set; }
        public int? Reserves { get; set; }
        public string Scale { get; set; }
        public string ActionDate { get; set; }
        public string Owner { get; set; }
        public string Location { get; set; }
        public string AbstractNumber { get; set; }
        public string CityETJ { get; set; }
        public string County { get; set; }
    }
    public class TitleBlockDetailSection
    {
        public string ControlNumber { get; set; }
        public string SectionNumber { get; set; }
        public string JobNumber { get; set; }
        public string SheetNumber { get; set; }
        public string District { get; set; }
        public string County { get; set; }
        public string HighwayNumber { get; set; }
    }
    public class DedicationInfo
    {
        public string County { get; set; }
        public decimal Acreage { get; set; }
        public string DevelopmentName { get; set; }
    }
    public class ValidationResult
    {
        public bool NoteFound { get; set; }
        public string NoteResult { get; set; }
        public string NoteDetails { get; set; }
        public bool AllFound { get; set; }
    }
    #endregion
}