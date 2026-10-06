"""
Tests for README.md content as changed in this PR.

The PR replaced the detailed technical fix description with "new branch woo yeah".
These tests verify the README reflects that change.
"""

import os
import unittest

README_PATH = os.path.join(os.path.dirname(__file__), "README.md")


class TestReadmeContent(unittest.TestCase):
    def setUp(self):
        with open(README_PATH, "r", encoding="utf-8") as f:
            self.content = f.read()

    def test_readme_exists(self):
        """README.md file must exist at the repo root."""
        self.assertTrue(os.path.isfile(README_PATH))

    def test_readme_new_content_present(self):
        """README.md must contain the updated branch note added in this PR."""
        self.assertIn("new branch woo yeah", self.content)

    def test_readme_old_technical_content_removed(self):
        """Old detailed fix description must no longer be present."""
        self.assertNotIn("The Fix — Two Places Per Module", self.content)

    def test_readme_old_fix1_removed(self):
        """References to 'Fix 1: Compilation files' from old README must be gone."""
        self.assertNotIn("Fix 1: Compilation files", self.content)

    def test_readme_old_fix2_removed(self):
        """References to 'Fix 2: Upload files' from old README must be gone."""
        self.assertNotIn("Fix 2: Upload files", self.content)

    def test_readme_old_clean_value_reference_removed(self):
        """Reference to clean_value function from old README must be gone."""
        self.assertNotIn("clean_value", self.content)

    def test_readme_is_not_empty(self):
        """README.md must not be empty after the PR change."""
        self.assertTrue(len(self.content.strip()) > 0)

    def test_readme_content_is_exactly_new_text(self):
        """README.md content (stripped) must exactly match the new PR text."""
        self.assertEqual(self.content.strip(), "new branch woo yeah")

    def test_readme_does_not_contain_column_mapping_details(self):
        """Old column_mapping loop description must not appear in README."""
        self.assertNotIn("column_mapping", self.content)

    def test_readme_does_not_reference_batch_number_fix(self):
        """Old BATCH_NUMBER fix instructions must not appear in README."""
        self.assertNotIn("BATCH_NUMBER", self.content)


if __name__ == "__main__":
    unittest.main()
