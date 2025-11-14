"""Tests for controllers and services added during refactoring."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from src.ui.core.file_picker_controller import FilePickerController
from src.ui.core.configuration_controller import ConfigurationController
from src.ui.core.state_manager import StateManager
from src.services.noise_registry_loader import NoiseRegistryLoader


# ============================================================================
# FilePickerController Tests
# ============================================================================


class TestFilePickerController:
    """Tests for FilePickerController extracted from EventHandlers."""

    @pytest.fixture
    def state_manager(self) -> StateManager:
        """Create a StateManager instance for testing."""
        return StateManager()

    @pytest.fixture
    def mock_gui(self, state_manager: StateManager) -> MagicMock:
        """Create a mock GUI with required attributes."""
        gui = MagicMock()
        gui.state_manager = state_manager
        gui.page = MagicMock()
        gui.workflow_manager = MagicMock()
        return gui

    @pytest.fixture
    def controller(self, mock_gui: MagicMock, state_manager: StateManager) -> FilePickerController:
        """Create a FilePickerController instance."""
        return FilePickerController(mock_gui, state_manager)

    def test_controller_initialization(self, controller: FilePickerController, state_manager: StateManager):
        """Test that controller initializes correctly with GUI and state references."""
        assert controller.state_manager == state_manager
        assert controller.state == state_manager.state

    def test_update_noise_folder_updates_state(
        self, controller: FilePickerController, state_manager: StateManager, tmp_path: Path
    ):
        """Test that updating noise folder modifies state correctly."""
        controller._update_noise_folder(str(tmp_path))
        assert state_manager.state.selected_noise_folder == str(tmp_path)

    def test_update_tests_folder_updates_state(
        self, controller: FilePickerController, state_manager: StateManager, tmp_path: Path
    ):
        """Test that updating tests folder modifies state correctly."""
        controller._update_tests_folder(str(tmp_path))
        assert state_manager.state.selected_tests_folder == str(tmp_path)

    def test_update_noise_registry_updates_state(
        self, controller: FilePickerController, state_manager: StateManager, tmp_path: Path
    ):
        """Test that updating noise registry modifies state correctly."""
        registry_file = tmp_path / "REGISTRO RUMORE.xlsx"
        registry_file.touch()
        
        controller._update_noise_registry(str(registry_file))
        assert state_manager.state.selected_noise_registry == str(registry_file)

    def test_update_registry_file_updates_state(
        self, controller: FilePickerController, state_manager: StateManager, tmp_path: Path
    ):
        """Test that updating registry file modifies state correctly."""
        registry_file = tmp_path / "registry.xlsx"
        registry_file.touch()
        
        controller._update_registry_file(str(registry_file))
        assert state_manager.state.selected_registry_file == str(registry_file)


# ============================================================================
# ConfigurationController Tests
# ============================================================================


class TestConfigurationController:
    """Tests for ConfigurationController extracted from EventHandlers."""

    @pytest.fixture
    def state_manager(self) -> StateManager:
        """Create a StateManager instance for testing."""
        return StateManager()

    @pytest.fixture
    def mock_gui(self, state_manager: StateManager) -> MagicMock:
        """Create a mock GUI with required attributes."""
        gui = MagicMock()
        gui.state_manager = state_manager
        gui.page = MagicMock()
        # Mock the config_tab
        gui.config_tab = MagicMock()
        gui.config_tab.sap_test_lab_containers = {}
        return gui

    @pytest.fixture
    def controller(self, mock_gui: MagicMock, state_manager: StateManager) -> ConfigurationController:
        """Create a ConfigurationController instance."""
        return ConfigurationController(mock_gui, state_manager)

    def test_controller_initialization(self, controller: ConfigurationController, state_manager: StateManager):
        """Test that controller initializes correctly with GUI and state references."""
        assert controller.state == state_manager

    def test_on_sap_checked_returns_event_handler(self, controller: ConfigurationController):
        """Test that on_sap_checked returns a callable event handler."""
        handler = controller.on_sap_checked("comparison")
        assert callable(handler)


# ============================================================================
# NoiseRegistryLoader Tests
# ============================================================================


class TestNoiseRegistryLoader:
    """Tests for NoiseRegistryLoader service extracted from config_tab."""

    @pytest.fixture
    def loader(self) -> NoiseRegistryLoader:
        """Create a NoiseRegistryLoader instance."""
        return NoiseRegistryLoader()

    def test_loader_initialization(self, loader: NoiseRegistryLoader):
        """Test that loader initializes with correct defaults."""
        assert loader.CACHE_VALIDITY_SECONDS == 600  # 10 minutes
        assert loader.LOAD_TIMEOUT_SECONDS == 15.0  # 15 seconds
        assert loader._cached_sap_codes is None
        assert loader._cache_timestamp is None

    def test_get_sap_codes_with_missing_file_returns_empty(self, loader: NoiseRegistryLoader):
        """Test that missing file returns empty list."""
        result = loader.get_sap_codes("/nonexistent/file.xlsx")

        assert result == []

    def test_clear_cache_resets_cache_state(self, loader: NoiseRegistryLoader):
        """Test that clearing cache resets internal state."""
        # Manually set cache
        loader._cached_sap_codes = ["612057"]
        loader._cache_timestamp = 12345.0
        
        # Clear cache
        loader.clear_cache()
        
        # Verify reset
        assert loader._cached_sap_codes is None
        assert loader._cache_timestamp is None

    def test_get_cache_age_returns_none_when_empty(self, loader: NoiseRegistryLoader):
        """Test that get_cache_age returns None when cache is empty."""
        assert loader.get_cache_age() is None

    def test_is_loading_returns_false_initially(self, loader: NoiseRegistryLoader):
        """Test that is_loading returns False initially."""
        assert loader.is_loading() is False

    def test_force_reload_bypasses_cache(self, loader: NoiseRegistryLoader, tmp_path: Path):
        """Test that force_reload parameter bypasses cache."""
        test_file = tmp_path / "REGISTRO RUMORE.xlsx"
        test_file.touch()

        # Mock the internal loading method
        with patch.object(loader, '_load_sap_codes', return_value=["612057"]) as mock_load:
            # First call
            loader.get_sap_codes(str(test_file))

            # Second call with force_reload - should bypass cache
            loader.get_sap_codes(str(test_file), force_reload=True)

            # Load should be called twice
            assert mock_load.call_count == 2

    def test_cache_validity_check(self, loader: NoiseRegistryLoader):
        """Test that cache validity is checked correctly."""
        # New cache should be valid
        loader._cache_timestamp = 12345.0
        loader._cached_sap_codes = ["612057"]
        
        # Mock time to be within validity period
        with patch('time.time', return_value=12345.0 + 300):  # 5 minutes later
            assert loader._is_cache_valid() is True
        
        # Mock time to be beyond validity period
        with patch('time.time', return_value=12345.0 + 700):  # 11+ minutes later
            assert loader._is_cache_valid() is False
