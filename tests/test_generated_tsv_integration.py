import argparse
import csv
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import addressmend as cleaner


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class GeneratedTsvIntegrationTest(unittest.TestCase):
    def test_generated_tsv_standard_procedure(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            temporary = Path(folder)
            database = temporary / "generated_test.sqlite"
            output = temporary / "generated_test_output.tsv"
            audit_path = temporary / "generated_test_audit.tsv"
            cleaner.build_index(
                argparse.Namespace(
                    sources=[str(PROJECT_ROOT / "generated_address_reference.tsv")],
                    db=str(database),
                    profile="generic",
                    postcode_column="postcode",
                    address_columns=["address"],
                    source_rank=None,
                    quiet=True,
                )
            )
            args = argparse.Namespace(
                input=str(PROJECT_ROOT / "generated_test_input.tsv"),
                output=str(output),
                audit=str(audit_path),
                db=str(database),
                memory=str(temporary / "memory.sqlite"),
                online_validate=True,
                validate_email_domains=True,
                doogal=True,
                doogal_delay=1.05,
                getaddress_key_env=None,
                address_threshold=0.84,
                auto_name=True,
                header=True,
                explain=False,
                quiet=True,
                fail_on_unresolved=False,
            )

            dns_response = {
                "Status": 0,
                "Answer": [
                    {"name": "example.org.", "type": 15, "data": "0 mail.example.org."}
                ],
            }
            with (
                patch.object(
                    cleaner,
                    "postcodes_io_lookup",
                    side_effect=cleaner.normalise_postcode,
                ),
                patch.object(cleaner, "postcodes_io_bulk", return_value=[]),
                patch.object(cleaner, "doogal_candidates", return_value=[]),
                patch.object(cleaner, "http_json", return_value=dns_response),
            ):
                self.assertEqual(cleaner.clean_records(args), 0)

            with output.open(encoding="utf-8-sig", newline="") as stream:
                rows = list(csv.DictReader(stream, delimiter="\t"))
            self.assertEqual(len(rows), 12)
            by_name = {row["first_name"]: row for row in rows}
            expected = {
                "Alex": ("Example", "18 Woodstock Close", "HP21 7HY"),
                "Taylor": ("Example", "Burnett Barn, Orcop", "HR2 8SF"),
                "Morgan": ("Example", "32 Maes-y-Bedw", "CF46 6UA"),
                "Jordan": ("Example", "24 Coldhams Lane", "CB1 8DZ"),
                "Casey": ("Domain", "5 Meadow Way", "PO22 7RJ"),
                "Avery": ("Test", "60 Aragon Way", "HP22 7DJ"),
                "Robin": ("Example", "Lower Burlorne, Tregoose", "PL30 3AJ"),
                "Cameron": ("Example", "173 Huntingdon Road", "CB3 0DH"),
                "OCR": ("Example", "28 Malvern Avenue", "FY6 7TL"),
                "Formatting": ("Example", "8 Constable Lee", "BB4 8EN"),
                "Jamie": ("Mitchener", "Hurst Crescent", "BB4 7SX"),
            }
            for first_name, values in expected.items():
                row = by_name[first_name]
                actual = (row["last_name"], row["address"], row["postcode"])
                self.assertEqual(actual, values)

            self.assertEqual(by_name["Jordan"]["email"], "jordan@example.org")
            self.assertEqual(by_name["Casey"]["email"], "person@gmail.com")
            self.assertEqual(by_name["Formatting"]["email"], "formatting_user@hotmail.com")
            self.assertEqual(by_name["Ambiguous"]["address"], "12")
            self.assertEqual(by_name["Ambiguous"]["email"], "hola@spanishwithmonica")

            with audit_path.open(encoding="utf-8-sig", newline="") as stream:
                audit = list(csv.DictReader(stream, delimiter="\t"))
            unresolved = [row for row in audit if row["confidence"] == "unresolved"]
            verified = [row for row in audit if row["confidence"] == "verified"]
            self.assertEqual(len(unresolved), 2)
            self.assertEqual(len(verified), 9)


if __name__ == "__main__":
    unittest.main()
