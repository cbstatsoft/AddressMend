import argparse
import csv
import tempfile
import unittest
import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import addressmend as cleaner


class NormalisationTests(unittest.TestCase):
    def test_postcode_spacing_and_case(self) -> None:
        self.assertEqual(cleaner.normalise_postcode("hr11hh"), "HR1 1HH")
        self.assertTrue(cleaner.valid_postcode("SW1A 1AA"))
        self.assertFalse(cleaner.valid_postcode("LL22 7G"))

    def test_markdown_email_is_unwrapped(self) -> None:
        email, problem = cleaner.clean_email(
            "[sample\\_person@example.org](mailto\\:sample_person@example.org)"
        )
        self.assertEqual(email, "sample_person@example.org")
        self.assertIsNone(problem)
        self.assertEqual(
            cleaner.email_change_confidence(
                "[sample\\_person@example.org](mailto\\:sample_person@example.org)"
            ),
            "formatting",
        )

    def test_email_ocr_repairs_are_conservative(self) -> None:
        cases = {
            "sample.person@hotmail.co,uk": "sample.person@hotmail.co.uk",
            "person@gmai1.com": "person@gmail.com",
            "person (at) example (dot) org": "person@example.org",
            "sample.person@1112@yahoo.com": "sample.person1112@yahoo.com",
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                email, problem = cleaner.clean_email(raw)
                self.assertEqual(email, expected)
                self.assertIsNone(problem)

        ambiguous, problem = cleaner.clean_email("hola@spanishwithmonica")
        self.assertEqual(ambiguous, "hola@spanishwithmonica")
        self.assertEqual(problem, "malformed email")

    def test_markdown_table_parser_preserves_six_columns(self) -> None:
        row = cleaner.split_markdown_row(
            "| Mrs | Anne | Example | 12 | CV37 0AA | anne@example.org |"
        )
        self.assertEqual(row, ["Mrs", "Anne", "Example", "12", "CV37 0AA", "anne@example.org"])

    def test_email_supported_surname_correction(self) -> None:
        audit: list[cleaner.Audit] = []
        record = cleaner.basic_clean(
            cleaner.Record("Mr", "Alex", "Exemple", "18", "HP21 7HY", "alex.example@example.org"),
            1,
            audit,
            True,
        )
        self.assertEqual(record.last_name, "Example")
        self.assertTrue(any(item.field == "last_name" for item in audit))

    def test_full_and_partial_address_ocr_matching(self) -> None:
        candidates = [
            ("32 Maes-y-Bedw", "CF46 6UA"),
            ("31 Maes-y-Bedw", "CF46 6UA"),
        ]
        full = cleaner.choose_address("32 Maehbed", candidates)
        partial = cleaner.choose_address("32", candidates)
        self.assertIsNotNone(full)
        self.assertIsNotNone(partial)
        assert full is not None and partial is not None
        self.assertEqual(full[0], "32 Maes-y-Bedw")
        self.assertEqual(partial[0], "32 Maes-y-Bedw")
        self.assertTrue(full[3] and partial[3])

    def test_building_name_ocr_is_compared_without_locality_suffix(self) -> None:
        choice = cleaner.choose_address(
            "Burneit Barn",
            [("Burnett Barn, Orcop", "HR2 8SF"), ("Church Barn, Orcop", "HR2 8SF")],
        )
        self.assertIsNotNone(choice)
        assert choice is not None
        self.assertEqual(choice[0], "Burnett Barn, Orcop")
        self.assertTrue(choice[3])

    def test_address_audit_reason_mentions_postcode_harmonisation(self) -> None:
        reason = cleaner.address_match_reason(
            "32 Maehbed", "32 Maes-y-Bedw", "offline address index", "CF46 6UA"
        )
        self.assertIn("OCR-corrected", reason)
        self.assertIn("CF46 6UA", reason)

    def test_doogal_street_consensus_preserves_unverified_number(self) -> None:
        candidates = [
            ("1 Langdon Avenue", "HP21 9UX"),
            ("3 Langdon Avenue", "HP21 9UX"),
            ("5 Langdon Avenue", "HP21 9UX"),
        ]
        self.assertEqual(
            cleaner.street_consensus_suggestion("2", candidates),
            ("2 Langdon Avenue", "HP21 9UX", 0.9),
        )
        damaged = cleaner.street_consensus_suggestion(
            "96 Hillwry",
            [("4 Hillary Close", "HP21 9TN"), ("8 Hillary Close", "HP21 9TN")],
        )
        self.assertIsNotNone(damaged)
        assert damaged is not None
        self.assertEqual(damaged[:2], ("96 Hillary Close", "HP21 9TN"))

    def test_doogal_flat_list_collapses_to_shared_base(self) -> None:
        candidates = [
            ("Flat 1, 70 Walton Street", "HP21 7QP"),
            ("Flat 8, 70 Walton Street", "HP21 7QP"),
        ]
        self.assertEqual(
            cleaner.base_address_consensus("70", candidates),
            ("70 Walton Street", "HP21 7QP"),
        )

    def test_approved_person_spelling_is_reused_by_email(self) -> None:
        memory = cleaner.connect_memory(":memory:")
        assert memory is not None
        approved = cleaner.Record("Mr", "Alex", "Example", "", "", "person@example.org")
        memory.execute(
            "INSERT INTO people VALUES(?,?,?)",
            (approved.email, json.dumps(approved.values()), 1),
        )
        audit: list[cleaner.Audit] = []
        result = cleaner.apply_person_memory(
            cleaner.Record("Mr", "Alex", "Exampel", "24", "HP21 9UB", approved.email),
            1,
            audit,
            memory,
        )
        self.assertEqual(result.last_name, "Example")
        self.assertEqual(audit[0].confidence, "learned")
        memory.close()

    def test_nominatim_finds_missing_postcode_conservatively(self) -> None:
        response = MagicMock()
        response.__enter__.return_value.read.return_value = json.dumps(
            [
                {
                    "address": {
                        "house_number": "72",
                        "road": "Example Drive",
                        "postcode": "AB12 3EF",
                    }
                }
            ]
        ).encode()
        with (
            patch.object(cleaner.urllib.request, "urlopen", return_value=response),
            patch.object(cleaner.time, "sleep"),
        ):
            found = cleaner.nominatim_address_lookup("72 Example Drive", None, 1.05)
        self.assertEqual(found, ("72 Example Drive", "AB12 3EF"))

    def test_x11_clipboard_owner_receives_tab_separated_text(self) -> None:
        with (
            patch.object(cleaner.os, "name", "posix"),
            patch.dict(cleaner.os.environ, {"DISPLAY": ":1"}, clear=True),
            patch.object(cleaner.shutil, "which", side_effect=lambda name: "/usr/bin/xclip" if name == "xclip" else None),
            patch.object(cleaner, "_start_clipboard_owner") as owner,
        ):
            route = cleaner.clipboard_set("A\tB\n1\t2\n")
        self.assertEqual(route, "X11 clipboard through xclip")
        owner.assert_called_once_with(
            ("xclip", "-selection", "clipboard", "-in"), "A\tB\n1\t2\n"
        )

    def test_review_flags_do_not_duplicate_missing_postcode(self) -> None:
        record = cleaner.Record(address="72 Example Drive")
        audit = [
            cleaner.Audit(1, "postcode", "", "", "unresolved", "missing postcode")
        ]
        cleaner.add_record_review_flags(record, record, 1, audit)
        postcode_flags = [item for item in audit if item.field == "postcode"]
        self.assertEqual(len(postcode_flags), 1)

    def test_download_filename_is_sanitised(self) -> None:
        self.assertEqual(
            cleaner.safe_download_name("https://example.org/files/address%20data.zip"),
            "address data.zip",
        )

    def test_uncommon_email_domain_mx_lookup_is_cached(self) -> None:
        cache = cleaner.connect_memory(":memory:")
        assert cache is not None
        response = {
            "Status": 0,
            "Answer": [{"name": "example.org.", "type": 15, "data": "0 mail.example.org."}],
        }
        with patch.object(cleaner, "http_json", return_value=response) as lookup:
            status = cleaner.uncommon_email_domain_status("person@example.org", cache)
            self.assertEqual(status[0], "valid")
            self.assertEqual(lookup.call_count, 1)
        with patch.object(cleaner, "http_json") as cached_lookup:
            cached = cleaner.uncommon_email_domain_status("person@example.org", cache)
            self.assertEqual(cached[0], "valid")
            cached_lookup.assert_not_called()
        cache.close()

    def test_nonexistent_uncommon_email_domain_is_flagged(self) -> None:
        with patch.object(cleaner, "http_json", return_value={"Status": 3}):
            status = cleaner.uncommon_email_domain_status("person@not-a-real-domain.invalid", None)
        self.assertEqual(status[0], "invalid")


class OfflineIndexTests(unittest.TestCase):
    def test_hmlr_import_and_address_selection(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "hmlr.csv"
            database = root / "addresses.sqlite"
            with source.open("w", encoding="utf-8", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(
                    [
                        "id",
                        "100000",
                        "2020-01-01",
                        "HP21 7HY",
                        "D",
                        "N",
                        "F",
                        "18",
                        "",
                        "WOODSTOCK CLOSE",
                        "AYLESBURY",
                        "AYLESBURY",
                        "BUCKINGHAMSHIRE",
                        "AYLESBURY VALE",
                        "A",
                        "A",
                    ]
                )
            args = argparse.Namespace(
                sources=[str(source)],
                db=str(database),
                profile="hmlr",
                postcode_column="postcode",
                address_columns=["address"],
                source_rank=None,
                quiet=True,
            )
            cleaner.build_index(args)
            index = cleaner.AddressIndex(str(database))
            choice = cleaner.choose_address("18", index.by_postcode("HP21 7HY"))
            self.assertIsNotNone(choice)
            assert choice is not None
            self.assertEqual(choice[0], "18 Woodstock Close")
            self.assertTrue(choice[3])


class ReleaseMetadataTests(unittest.TestCase):
    def test_release_identity(self) -> None:
        self.assertEqual(cleaner.VERSION, "1.2.0")
        self.assertIn("Connor Baird", cleaner.COPYRIGHT)


if __name__ == "__main__":
    unittest.main()
