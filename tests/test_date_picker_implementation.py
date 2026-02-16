"""
Tests for Date Picker Implementation

Tests the Python VBA injection system for the new date picker components.

Usage:
    python -m pytest tests/test_date_picker_implementation.py -v
"""
import os
import sys
import unittest
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.vba_injection.ui_helpers import add_date_entry_control
from src.vba_injection.utils import get_vba_path, read_vba_file


class TestVBAFileStructure(unittest.TestCase):
    """Test that all required VBA files exist and are readable"""

    def test_modDateUtils_exists(self):
        """Test that modDateUtils.bas file exists"""
        path = get_vba_path("modDateUtils.bas", "modules")
        self.assertTrue(os.path.exists(path), f"modDateUtils.bas not found at {path}")

    def test_modDateUtils_readable(self):
        """Test that modDateUtils.bas can be read"""
        path = get_vba_path("modDateUtils.bas", "modules")
        content = read_vba_file(path)
        self.assertIsInstance(content, str)
        self.assertGreater(len(content), 0)

    def test_modDateUtils_contains_functions(self):
        """Test that modDateUtils contains expected function signatures"""
        path = get_vba_path("modDateUtils.bas", "modules")
        content = read_vba_file(path)

        # Check for key functions
        self.assertIn("Public Function ParseDate", content)
        self.assertIn("Public Function ValidateDate", content)
        self.assertIn("Public Function FormatDateDisplay", content)
        self.assertIn("Public Function FormatDateStorage", content)
        self.assertIn("Public Function ShowDatePicker", content)

    def test_frmCalendarPicker_exists(self):
        """Test that frmCalendarPicker.vba file exists"""
        path = get_vba_path("frmCalendarPicker.vba", "forms")
        self.assertTrue(os.path.exists(path), f"frmCalendarPicker.vba not found at {path}")

    def test_frmCalendarPicker_readable(self):
        """Test that frmCalendarPicker.vba can be read"""
        path = get_vba_path("frmCalendarPicker.vba", "forms")
        content = read_vba_file(path)
        self.assertIsInstance(content, str)
        self.assertGreater(len(content), 0)

    def test_frmCalendarPicker_contains_interface(self):
        """Test that frmCalendarPicker has the public interface"""
        path = get_vba_path("frmCalendarPicker.vba", "forms")
        content = read_vba_file(path)

        self.assertIn("Public Function ShowCalendar", content)
        self.assertIn("Private Sub btnSelect_Click", content)
        self.assertIn("Private Sub btnCancel_Click", content)


class TestModDateUtilsCode(unittest.TestCase):
    """Test the VBA code quality in modDateUtils"""

    def setUp(self):
        """Read the modDateUtils file once for all tests"""
        path = get_vba_path("modDateUtils.bas", "modules")
        self.content = read_vba_file(path)

    def test_no_null_assignments(self):
        """Test that we use Empty instead of Null (VBA compatibility)"""
        # VBA doesn't allow direct Null assignments to function returns
        # We should use Empty instead
        self.assertNotIn("ParseDate = Null", self.content)
        self.assertNotIn("ValidateDate = Null", self.content)
        self.assertNotIn("GetDateFromString = Null", self.content)

    def test_uses_empty_for_invalid(self):
        """Test that functions return Empty for invalid values"""
        self.assertIn("ParseDate = Empty", self.content)

    def test_uses_isempty_checks(self):
        """Test that we check IsEmpty instead of IsNull"""
        self.assertIn("IsEmpty", self.content)

    def test_has_error_handling(self):
        """Test that functions have error handling"""
        self.assertIn("On Error GoTo", self.content)

    def test_has_locale_independent_parsing(self):
        """Test that date parsing is locale-independent"""
        self.assertIn("DateSerial", self.content)
        self.assertIn("Split", self.content)

    def test_date_range_validation(self):
        """Test that date range 2020-2030 is enforced"""
        self.assertIn("2020", self.content)
        self.assertIn("2030", self.content)


class TestFormUpdates(unittest.TestCase):
    """Test that all forms were updated correctly"""

    def test_frmAdmission_uses_modDateUtils(self):
        """Test that frmAdmission uses modDateUtils functions"""
        path = get_vba_path("frmAdmission.vba", "forms")
        content = read_vba_file(path)

        self.assertIn("modDateUtils.ParseDate", content)
        self.assertIn("modDateUtils.FormatDateDisplay", content)
        self.assertIn("txtDate_picker_Click", content)

    def test_frmAdmission_removed_duplicate(self):
        """Test that ParseDateAdm function was removed"""
        path = get_vba_path("frmAdmission.vba", "forms")
        content = read_vba_file(path)

        # Should have a comment saying it was removed
        self.assertIn("ParseDateAdm function removed", content)
        # Should not have the function definition
        self.assertNotIn("Private Function ParseDateAdm(", content)

    def test_frmDeath_uses_modDateUtils(self):
        """Test that frmDeath uses modDateUtils functions"""
        path = get_vba_path("frmDeath.vba", "forms")
        content = read_vba_file(path)

        self.assertIn("modDateUtils.ParseDate", content)
        self.assertIn("modDateUtils.FormatDateDisplay", content)
        self.assertIn("txtDate_picker_Click", content)

    def test_frmDeath_removed_duplicate(self):
        """Test that ParseDateDth function was removed"""
        path = get_vba_path("frmDeath.vba", "forms")
        content = read_vba_file(path)

        # Should have a comment saying it was removed
        self.assertIn("ParseDateDth function removed", content)
        # Should not have the function definition
        self.assertNotIn("Private Function ParseDateDth(", content)

    def test_frmAgesEntry_uses_modDateUtils(self):
        """Test that frmAgesEntry uses modDateUtils functions"""
        path = get_vba_path("frmAgesEntry.vba", "forms")
        content = read_vba_file(path)

        self.assertIn("modDateUtils.ParseDate", content)
        self.assertIn("modDateUtils.FormatDateDisplay", content)
        self.assertIn("txtDate_picker_Click", content)

    def test_all_forms_use_isempty(self):
        """Test that all forms use IsEmpty instead of IsNull"""
        forms = ["frmAdmission.vba", "frmDeath.vba", "frmAgesEntry.vba"]

        for form_name in forms:
            path = get_vba_path(form_name, "forms")
            content = read_vba_file(path)

            # Should not have IsNull checks for date variables
            # Context: checking parsed dates from modDateUtils
            if "IsNull" in content:
                # Check if it's in the context of date validation
                lines = content.split('\n')
                for i, line in enumerate(lines):
                    if "IsNull" in line and "Date" in line:
                        self.fail(f"{form_name} still uses IsNull for date checks on line {i+1}")


class TestPythonComponents(unittest.TestCase):
    """Test Python helper functions and builders"""

    def test_calendar_form_builder_imports(self):
        """Test that calendar form builder can be imported"""
        try:
            from src.vba_injection.calendar_form_builder import create_calendar_picker_form
            self.assertTrue(callable(create_calendar_picker_form))
        except ImportError as e:
            self.fail(f"Failed to import calendar_form_builder: {e}")

    def test_ui_helpers_has_date_control(self):
        """Test that add_date_entry_control exists"""
        try:
            from src.vba_injection.ui_helpers import add_date_entry_control
            self.assertTrue(callable(add_date_entry_control))
        except ImportError as e:
            self.fail(f"Failed to import add_date_entry_control: {e}")

    def test_core_imports_calendar_builder(self):
        """Test that core.py imports calendar builder"""
        path = os.path.join(project_root, "src", "vba_injection", "core.py")
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn("from .calendar_form_builder import create_calendar_picker_form", content)
        self.assertIn("create_calendar_picker_form(vbproj)", content)

    def test_userform_builder_imports_updated(self):
        """Test that userform_builder imports are updated"""
        path = os.path.join(project_root, "src", "vba_injection", "userform_builder.py")
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn("add_date_entry_control", content)
        self.assertIn("from .calendar_form_builder import create_calendar_picker_form", content)

    def test_modDateUtils_in_injection_list(self):
        """Test that modDateUtils is in the module injection list"""
        path = os.path.join(project_root, "src", "vba_injection", "core.py")
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Should be in the modules list
        self.assertIn('"modDateUtils.bas"', content)
        self.assertIn('("modDateUtils", "modDateUtils.bas")', content)


class TestCalendarFormBuilder(unittest.TestCase):
    """Test calendar form builder code"""

    def test_calendar_builder_creates_all_controls(self):
        """Test that calendar builder creates all required controls"""
        path = os.path.join(project_root, "src", "vba_injection", "calendar_form_builder.py")
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Check for essential controls
        required_controls = [
            'btnPrev', 'btnNext', 'btnToday', 'btnSelect', 'btnCancel',
            'cmbMonth', 'cmbYear', 'lblMonthYear'
        ]

        for control in required_controls:
            self.assertIn(f'"{control}"', content, f"Missing control: {control}")

    def test_calendar_builder_creates_day_grid(self):
        """Test that 42 day labels are created (6 rows x 7 days)"""
        path = os.path.join(project_root, "src", "vba_injection", "calendar_form_builder.py")
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Should create lblDay_X_Y labels
        self.assertIn("range(6)", content)  # 6 rows
        self.assertIn("range(7)", content)  # 7 columns
        self.assertIn("lblDay_", content)


class TestCodeDuplicationRemoval(unittest.TestCase):
    """Test that code duplication was successfully eliminated"""

    def test_no_parsedateadm_in_forms(self):
        """Test that ParseDateAdm doesn't exist as a function anymore"""
        forms = ["frmAdmission.vba", "frmDeath.vba", "frmAgesEntry.vba"]

        for form_name in forms:
            path = get_vba_path(form_name, "forms")
            content = read_vba_file(path)

            # Should not have the actual function (only a comment)
            self.assertNotIn("Private Function ParseDateAdm(", content)

    def test_no_parsedatedth_in_forms(self):
        """Test that ParseDateDth doesn't exist as a function anymore"""
        path = get_vba_path("frmDeath.vba", "forms")
        content = read_vba_file(path)

        self.assertNotIn("Private Function ParseDateDth(", content)

    def test_all_forms_use_centralized_validation(self):
        """Test that all forms call modDateUtils.ParseDate"""
        forms = ["frmAdmission.vba", "frmDeath.vba", "frmAgesEntry.vba"]

        for form_name in forms:
            path = get_vba_path(form_name, "forms")
            content = read_vba_file(path)

            self.assertIn("modDateUtils.ParseDate(", content,
                         f"{form_name} doesn't use centralized ParseDate")


def run_tests():
    """Run all tests and print results"""
    print("=" * 70)
    print("Date Picker Implementation Test Suite")
    print("=" * 70)
    print()

    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    # Add all test classes
    suite.addTests(loader.loadTestsFromTestCase(TestVBAFileStructure))
    suite.addTests(loader.loadTestsFromTestCase(TestModDateUtilsCode))
    suite.addTests(loader.loadTestsFromTestCase(TestFormUpdates))
    suite.addTests(loader.loadTestsFromTestCase(TestPythonComponents))
    suite.addTests(loader.loadTestsFromTestCase(TestCalendarFormBuilder))
    suite.addTests(loader.loadTestsFromTestCase(TestCodeDuplicationRemoval))

    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    print()
    print("=" * 70)
    if result.wasSuccessful():
        print(f"SUCCESS! All {result.testsRun} tests passed.")
    else:
        print(f"FAILURE! {len(result.failures)} failures, {len(result.errors)} errors")
    print("=" * 70)

    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
