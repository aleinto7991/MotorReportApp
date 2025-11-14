# GUI Module Documentation

This directory contains the source code for the graphical user interface (GUI) of the Motor Report Generator application, built using the Flet framework.

## Modular Architecture

The GUI has been refactored into a modular architecture to improve organization, maintainability, and scalability. The key components of this architecture are:

- **`main_gui.py`**: The main entry point for the GUI. It initializes the application, orchestrates the different components, and builds the main window.

- **`core/`**: This directory contains the core logic and managers that drive the application's functionality.
  - `state_manager.py`: Manages the application's state, including user selections, file paths, and workflow progress.
  - `event_handlers.py`: Contains all the event handling logic for UI interactions (e.g., button clicks, file picking).
  - `search_manager.py`: Handles the logic for searching tests and displaying the results.
  - `report_manager.py`: Manages the interaction with the backend for report generation.
  - `workflow_manager.py`: Controls the application's workflow, managing tab navigation and step validation.

- **`components/`**: Contains reusable UI components that are used across different parts of the GUI.
  - `base.py`: Base classes for components and tabs.
  - `dialogs.py`: Dialog windows for user interaction (e.g., SAP selection).

- **`tabs/`**: Each file in this directory represents a tab in the main interface.
  - `setup_tab.py`: The first tab for setting up file paths.
  - `search_select_tab.py`: The second tab for searching and selecting tests.
  - `config_tab.py`: The third tab for configuring the report generation.
  - `generate_tab.py`: The final tab for generating the report.

- **`utils/`**: Contains utility modules and helper classes.
  - `helpers.py`: Helper classes like `StatusManager`.
  - `selection_cache.py`: Manages caching of user selections.

## Workflow

The application guides the user through a four-step process, with each step corresponding to a tab:

1.  **Setup**: The user verifies the auto-detected paths for the test data folder and registry file.
2.  **Search & Select**: The user searches for tests by SAP code or test number and selects the tests to include in the report.
3.  **Configure**: The user configures options for the report, such as including noise data or generating a comparison sheet.
4.  **Generate**: The user initiates the report generation process.

The `WorkflowManager` ensures that the user proceeds through the steps in the correct order and that each step is validated before moving to the next.
