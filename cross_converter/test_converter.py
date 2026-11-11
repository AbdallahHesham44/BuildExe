import os
import unittest
import pandas as pd
import tempfile
import shutil

from cross_converter import (
    TYPE1_MAP,
    TYPE2_MAP,
    convert,
    build_frame,
    resolve_columns,
)

class TestCrossConverter(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.sample_cols = [
            "ProductName", "PartId", "PartNumber", "CompanyName", "Lifecycle", "GroupDescription", "FamilyName",
            "TempGradeCode", "TempGrade", "OrgPartRisk", "RiskScore", "CrossPartID", "CrossPartNumber",
            "CrossCompanyName", "CrossLifecycle", "CrossGroupDescription", "CrossFamilyName",
            "CrossTempGradeCode", "CrossTempGrade", "CrossRohsStatus", "CrossRiskScore", "Grade", "DisplayOrder",
            "FroecedCrossGrade", "ManualRank", "SourceName", "Type", "GroupId", "IsApproved", "Action", "actions"
        ]

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _create_sample_excel(self, filename="sample.xlsx", col_modifier=None, row_count=3):
        cols = list(self.sample_cols)
        if col_modifier:
            cols = [col_modifier(c) for c in cols]
        
        data = {c: [f"val_{c}_{i}" for i in range(row_count)] for c in cols}
        df = pd.DataFrame(data)
        path = os.path.join(self.test_dir, filename)
        df.to_excel(path, index=False)
        return path

    def test_successful_conversion(self):
        in_path = self._create_sample_excel("sample.xlsx", row_count=5)
        out_dir = os.path.join(self.test_dir, "output")
        rows, p1, p2 = convert(in_path, out_dir)

        self.assertEqual(rows, 5)
        self.assertTrue(os.path.exists(p1))
        self.assertTrue(os.path.exists(p2))

        # Check type1
        df1 = pd.read_excel(p1)
        self.assertEqual(list(df1.columns), list(TYPE1_MAP.keys()))
        self.assertEqual(len(df1), 5)
        self.assertEqual(df1["PartNumberC"].iloc[0], "val_PartNumber_0")
        self.assertTrue(pd.isna(df1["FeatureName"].iloc[0]) or df1["FeatureName"].iloc[0] == "")

        # Check type2
        df2 = pd.read_excel(p2)
        self.assertEqual(list(df2.columns), list(TYPE2_MAP.keys()))
        self.assertEqual(len(df2), 5)
        self.assertEqual(df2["ParIDC"].iloc[0], "val_PartId_0")
        self.assertEqual(df2["CrossGrade"].iloc[0], "val_Grade_0")

    def test_case_insensitive_column_matching(self):
        # Upper/lower case mix: e.g. PARTID, crosspartnumber
        def modifier(col):
            if col == "PartId":
                return "PARTID"
            if col == "CrossPartNumber":
                return "crosspartnumber"
            return col

        in_path = self._create_sample_excel("case_test.xlsx", col_modifier=modifier, row_count=2)
        rows, p1, p2 = convert(in_path, self.test_dir)
        self.assertEqual(rows, 2)
        df1 = pd.read_excel(p1)
        self.assertEqual(df1["PartIDc"].iloc[0], "val_PARTID_0")
        self.assertEqual(df1["PartNumberX"].iloc[0], "val_crosspartnumber_0")

    def test_missing_column_raises_error(self):
        # Create excel with missing column
        data = {"PartNumber": ["PN1"], "CompanyName": ["Comp1"]}
        df = pd.DataFrame(data)
        path = os.path.join(self.test_dir, "missing.xlsx")
        df.to_excel(path, index=False)

        with self.assertRaises(ValueError) as ctx:
            convert(path, self.test_dir)
        self.assertIn("Missing columns in input file", str(ctx.exception))

if __name__ == "__main__":
    unittest.main()
