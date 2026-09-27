import unittest

from processor import ALLOWED_STATUSES, process_file


class ProcessorTests(unittest.TestCase):
    def test_csv_bytes(self):
        data = b"supplier,product,price\nAcme,Widget,9.99"
        records = process_file(data)
        self.assertIsInstance(records, list)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["title"], "Acme")
        self.assertIn(records[0]["status"], ALLOWED_STATUSES)
        self.assertIsInstance(records[0]["details"], dict)
        self.assertNotIn("due_date", records[0]["details"])

    def test_multiple_csv_rows(self):
        data = b"supplier,product,price\nAcme,Widget,9.99\nGlobex,Gadget,12.50"
        records = process_file(data)
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["title"], "Acme")
        self.assertEqual(records[1]["title"], "Globex")

    def test_empty_bytes(self):
        records = process_file(b"")
        self.assertIsInstance(records, list)
        self.assertEqual(records[0]["status"], "Missing:critical")
        self.assertEqual(records[0]["title"], "Unnamed Entity")

    def test_status_never_nested_due_date(self):
        data = b"supplier,product,price,due date\nAcme,Widget,9.99,2026-01-31"
        records = process_file(data)
        self.assertEqual(len(records), 1)
        self.assertNotIn("due_date", records[0]["details"])
        self.assertEqual(records[0]["due_date"], "2026-01-31")


if __name__ == "__main__":
    unittest.main()
