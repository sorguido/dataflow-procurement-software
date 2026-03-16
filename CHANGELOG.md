# DataFlow – Changelog

Version 2.0.1
Release date: YYYY-MM-DD

Refactoring release focused on codebase modularization, stability improvements, and internal architecture cleanup.  
No major user-facing features were introduced, but the internal structure of the application has been significantly improved.

## Added
- Modular project structure with separated packages for utilities, services, database helpers, and UI components.
- Dedicated modules for reusable utilities (string, formatting, validation, window management, resource paths, user configuration).
- HelpWindow extracted into a standalone module.
- Improved Help search functionality with robust Python-based text search to prevent Tkinter crashes.

## Changed
- Major internal refactoring of the application to reduce the monolithic structure.
- Extraction of multiple utility and helper functions into dedicated modules.
- Improved path management and resource resolution.
- Improved attachment filename sanitization.
- Improved date conversion utilities for database storage.
- Improved numeric formatting utilities.

## Fixed
- Crash in HelpWindow search functionality (segmentation fault).
- File dialog appearing behind the attachment window when adding files.
- Various minor issues related to window focus and modal dialogs.

## Refactoring
- Extraction of utility modules:
  - string_utils
  - format_utils
  - window_utils
  - user_utils
  - resource_utils
  - i18n_utils
  - validation_utils
- Extraction of HelpWindow into `ui/help_window.py`
- Additional separation of services and database helpers
- Significant reduction of the main application file size

Internal Improvements
- Cleaner module imports
- Better separation of responsibilities
- Reduced code duplication
- Improved maintainability for future development
- Improved code readability for external contributors

## Notes
This release focuses on internal improvements and codebase maintainability.  
Application behaviour and user workflows remain unchanged compared to version 2.0.0.
