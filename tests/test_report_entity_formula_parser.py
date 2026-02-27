from __future__ import annotations

import unittest

from report_entity_registry import validate_formula_expression


class ReportEntityFormulaParserTests(unittest.TestCase):
    def test_accepts_valid_expression(self):
        out = validate_formula_expression(
            "sum(planned_hours) - max(actual_hours)",
            known_entity_keys={"planned_hours", "actual_hours", "capacity"},
            current_entity_key="capacity",
        )
        self.assertEqual(out["references"], ["actual_hours", "planned_hours"])

    def test_rejects_unknown_function(self):
        with self.assertRaises(ValueError) as ctx:
            validate_formula_expression("median(planned_hours)", known_entity_keys={"planned_hours"}, current_entity_key="")
        self.assertIn("Unknown function", str(ctx.exception))

    def test_rejects_unknown_entity(self):
        with self.assertRaises(ValueError) as ctx:
            validate_formula_expression("sum(missing_entity)", known_entity_keys={"planned_hours"}, current_entity_key="")
        self.assertIn("Unknown entity", str(ctx.exception))

    def test_rejects_self_reference(self):
        with self.assertRaises(ValueError) as ctx:
            validate_formula_expression("capacity + planned_hours", known_entity_keys={"capacity", "planned_hours"}, current_entity_key="capacity")
        self.assertIn("Self reference", str(ctx.exception))

    def test_rejects_comma_multi_argument_call(self):
        with self.assertRaises(ValueError) as ctx:
            validate_formula_expression("sum(planned_hours, actual_hours)", known_entity_keys={"planned_hours", "actual_hours"}, current_entity_key="")
        self.assertIn("accepts one argument", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
