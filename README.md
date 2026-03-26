# Plan Review Automation Project

## Overview

This project leverages a third-party ML platform to process and analyze plan set documents. It combines trained machine learning models with custom backend logic to extract structured data and generate meaningful outputs for end users.

The system is designed to automate plan review workflows by identifying key attributes within documents and organizing them into usable results.

---

## Key Features

* Integration with a third-party ML platform for document processing
* Machine learning–driven attribute extraction from plan sets
* Automated classification of pages and document sections
* Configurable attribute mappings via application settings
* Backend logic to transform extracted data into user-friendly outputs

---

## How It Works

1. **Document Input**
   Plan set documents are uploaded and processed through the ML platform.

2. **Model Processing**
   Machine learning models analyze the documents and extract relevant attributes.

3. **Attribute Retrieval**
   The application retrieves extracted values via API calls.

4. **Business Logic Processing**
   Custom C# logic organizes and filters the extracted data.

5. **Output Generation**
   Results are formatted and returned for downstream use or display.

---

## Configuration

All environment-specific and sensitive values have been externalized to configuration settings.

Example keys in `App.config` or `Web.config`:

```
LAND_PLAN_SET_PAGE_CATEGORY_ATTRIBUTE_ID
TRANSPORTATION_PAGE_CATEGORY_ATTRIBUTE_ID
PLAT_TITLE_BLOCK_ATTRIBUTE_ID
PLAT_BLOCK_NUMBER_ATTRIBUTE_ID
TRANSPORTATION_TITLE_BLOCK_DETAIL_SECTION_ATTRIBUTE_ID
```

Ensure these values are set correctly for your environment.

---

## Tech Stack

* C# (.NET)
* Third-party ML platform API
* ConfigurationManager for environment configuration

---

## Security Considerations

* No credentials or sensitive identifiers are stored in source code
* All IDs and environment-specific values are externalized
* Sanitized version prepared for public repository

---

## Use Cases

* Automated plan review workflows
* Extraction of structured data from engineering drawings
* Classification of plan sheets (e.g., transportation, plat, land use)

---

## Getting Started

1. Clone the repository
2. Configure required AppSettings values
3. Build the project in Visual Studio
4. Run and integrate with your ML platform environment

---

## Notes

This repository contains a sanitized version of the original implementation. Internal schemas, identifiers, and environment-specific configurations have been abstracted.

---

## Author

Developed as part of a plan review automation initiative using machine learning capabilities.
