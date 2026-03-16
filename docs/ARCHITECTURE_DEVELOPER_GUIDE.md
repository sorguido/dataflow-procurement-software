# DataFlow – Developer Guide

Technical overview of the DataFlow codebase and project architecture.

This document is intended for developers who want to understand the internal structure of the application and contribute to the project.

------------------------------------------------------------

1. PROJECT OVERVIEW

DataFlow is a desktop application written in Python using Tkinter for the graphical interface and SQLite as the database backend.

The software is designed to manage procurement workflows and RFQ (Request for Quotation) processes.

Main features include:

- RFQ creation and management
- supplier quotations
- document attachments
- Excel import/export
- SQDC analysis
- multilingual interface (internationalization)

The application was originally developed as a single-file monolithic script (~8500 lines) and later refactored into a modular architecture to improve:

- maintainability
- readability
- extensibility
- developer onboarding

------------------------------------------------------------

2. HIGH-LEVEL ARCHITECTURE

The project now follows a modular structure.

dataflow/
│
├─ DataFlow 2.0.0.py        Application entry point (bootstrap)
├─ constants.py             Global UI constants and configuration
├─ database_manager.py      Main database interaction layer
│
├─ utils/                   Generic reusable helper functions
├─ services/                Application services
├─ database/                Database helpers and query utilities
├─ ui/                      User interface modules
│   ├─ dialogs/
│   └─ windows/
│
├─ locale/                  Translation files
├─ docs/                    Documentation
└─ add_data/                Static resources (help files etc.)

------------------------------------------------------------

3. APPLICATION ENTRY POINT

File: DataFlow 2.0.0.py

This file acts as the application bootstrap.

Responsibilities:

- application startup
- logging configuration
- loading configuration files
- initializing internationalization
- loading services
- starting the main UI

Logical startup flow:

initialize application
↓
load configuration
↓
initialize i18n
↓
initialize services
↓
launch MainWindow

This file should remain as lightweight as possible and mainly orchestrate the application startup.

------------------------------------------------------------

4. CORE MODULES

constants.py

Centralized configuration values used across the UI.

Typical contents:

- window sizes
- layout constants
- UI spacing
- default configuration values

Purpose:

Avoid hardcoding UI constants across multiple modules.

------------------------------------------------------------

database_manager.py

Main interface to the SQLite database.

Responsibilities:

- opening database connections
- executing queries
- reading RFQ data
- writing RFQ data
- managing attachments metadata
- transaction control

This module acts as the main data access layer.

------------------------------------------------------------

5. UTILITIES (utils)

The utils package contains reusable helper functions independent from business logic.

These modules are intentionally stateless.

------------------------------------------------------------

utils/string_utils.py

Handles string manipulation utilities.

Examples:

- username generation
- accent stripping
- string normalization

Purpose:

Standardize string handling across the application.

------------------------------------------------------------

utils/format_utils.py

Handles formatting of numeric values.

Examples:

- quantity formatting
- parsing numbers with comma decimal separators

Purpose:

Ensure consistent formatting between UI and database.

------------------------------------------------------------

utils/validation_utils.py

Handles input validation and formatting.

Examples:

- filename sanitization
- date conversion for database storage
- price formatting

Purpose:

Centralize input validation logic.

------------------------------------------------------------

utils/window_utils.py

Contains helper functions for Tkinter window management.

Examples:

- centering windows
- calculating optimal window sizes
- setting window icons

Purpose:

Standardize window behavior across all dialogs.

------------------------------------------------------------

utils/resource_utils.py

Handles file path resolution.

This module ensures resources can be loaded both:

- in development mode
- when bundled with PyInstaller

Typical usage:

resource_path("docs/manual.pdf")

------------------------------------------------------------

utils/user_utils.py

Handles user-specific configuration and identity.

Examples:

- locating configuration files
- loading user identity
- saving user settings

------------------------------------------------------------

utils/i18n_utils.py

Handles internationalization via gettext.

Responsibilities:

- initialize language system
- load translation files
- provide helper functions for localized UI elements

Supported languages:

- Italian
- English

The _() translation function is installed globally during application startup.

------------------------------------------------------------

6. SERVICES

The services package contains application-level services.

Typical responsibilities include:

- application startup procedures
- filesystem structure initialization
- application paths
- operational services not tied to UI

Services encapsulate operational logic independent from the user interface.

------------------------------------------------------------

7. DATABASE HELPERS

The database directory contains lower-level helpers related to database interaction.

These modules support:

- database structure
- query utilities
- specialized data operations

The database_manager module orchestrates these helpers.

------------------------------------------------------------

8. USER INTERFACE (ui)

The UI is built with Tkinter.

The UI layer is divided into two main parts:

ui/
├─ dialogs/
└─ windows/

------------------------------------------------------------

9. UI WINDOWS

ui/windows/

Contains large application windows.

Example:

view_request_window.py

This module implements the RFQ detail window.

Responsibilities:

- viewing RFQ details
- editing RFQ data
- managing suppliers
- managing attachments
- exporting RFQ data
- grid-based editing of RFQ items

This window represents one of the core workflows of the application.

------------------------------------------------------------

10. UI DIALOGS

ui/dialogs/

Contains smaller dialogs and modal windows.

Examples include:

- language selection dialogs
- progress dialogs
- reference editing dialogs
- user identity dialogs

These dialogs are designed to be self-contained UI components.

------------------------------------------------------------

11. INTERNATIONALIZATION

The application uses gettext for translations.

Structure:

locale/
    it/
    en/

Translation files are loaded during application startup using:

init_i18n(language_code)

The function _() is available globally for translating UI strings.

Example usage:

_("Save")

------------------------------------------------------------

12. ATTACHMENTS AND FILE MANAGEMENT

The application supports attachments associated with RFQs.

Supported attachment types include:

- internal documents
- supplier quotations

Operations supported:

- upload
- download
- export
- filename sanitization

------------------------------------------------------------

13. REFACTORING HISTORY

Originally the application consisted of:

DataFlow.py (~8500 lines)

The codebase was refactored to introduce:

- modular utilities
- separated UI components
- service layer
- improved maintainability

Goals of the refactoring:

- reduce monolithic code
- improve readability
- enable easier contributions

------------------------------------------------------------

14. DEVELOPMENT GUIDELINES

When contributing to the project:

Prefer:

- small focused modules
- reusable utilities
- minimal duplication
- separation of UI and logic

Avoid:

- large logic blocks inside the main file
- mixing UI code and business logic
- duplicating helper functions

------------------------------------------------------------

15. TESTING STRATEGY

Manual testing should verify:

- RFQ creation
- RFQ editing
- supplier management
- attachment management
- Excel export
- language switching
- help system
- application startup

------------------------------------------------------------

16. VERSIONING

Example version scheme:

2.0.0  original release
2.0.1  refactoring + stability improvements

Semantic versioning guidelines:

PATCH   bug fixes
MINOR   new features
MAJOR   breaking changes

------------------------------------------------------------

17. CONTRIBUTING

Contributors are encouraged to:

- open GitHub issues
- submit pull requests
- improve documentation
- propose architectural improvements

------------------------------------------------------------

18. LICENSE

The Linux version of DataFlow is released under the GNU GPLv3 license.

See the LICENSE file for details.

------------------------------------------------------------

End of document
