import importlib.util
import io
import sys
import unittest
from pathlib import Path

from openpyxl import Workbook


MODULE_PATH = Path(__file__).parents[1] / "monitor.py"
SPEC = importlib.util.spec_from_file_location("nav_monitor", MODULE_PATH)
monitor = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
sys.modules[SPEC.name] = monitor
SPEC.loader.exec_module(monitor)


class MonitorTests(unittest.TestCase):
    def test_finds_latest_workbook(self):
        html = """
        <a href="/july.xlsx">Arbeidssøkere og ledige stillinger. Kommune og kjennetegn. Juli 2026 (xls)</a>
        <a href="/august.xlsx">Arbeidssøkere og ledige stillinger. Kommune og kjennetegn. August 2026 (xls)</a>
        """
        found = monitor.find_latest_workbook(html, "https://www.nav.no/source")
        self.assertEqual(found["period"], "2026-08")
        self.assertEqual(found["url"], "https://www.nav.no/august.xlsx")

    def test_extracts_only_configured_municipalities(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Oversikt"
        sheet.append([None, None, None, None, None, None, None, None, None, None])
        sheet.append([None, "5059 Orkland", 279, 164, 69, 46, 1.7, 0.7, 0.5, 69])
        sheet.append([None, "0301 Oslo", 999, 888, 1, 2, 3.4, 0.1, 0.2, 700])
        output = io.BytesIO()
        workbook.save(output)
        values = monitor.extract_values(output.getvalue(), {"5059": "Orkland"})
        self.assertEqual(set(values), {"5059"})
        self.assertEqual(values["5059"]["fully_unemployed"], 164)
        self.assertEqual(values["5059"]["new_vacancies"], 69)

    def test_rss_contains_source_links_and_no_person_data(self):
        config = {
            "source_page": "https://www.nav.no/statistikk",
            "feed": {"title": "Test", "description": "Test"},
        }
        event = {
            "id": "nav-kommune-2026-08",
            "period_label": "August 2026",
            "source_url": "https://www.nav.no/table.xlsx",
            "observed_at": "2026-09-02T08:00:00+00:00",
            "values": {
                "5059": {
                    "name": "Orkland",
                    "fully_unemployed": 164,
                    "fully_unemployed_pct": 1.7,
                    "new_vacancies": 69,
                }
            },
            "previous_values": {"5059": {"fully_unemployed": 160}},
        }
        xml = monitor.build_rss([event], config).decode()
        self.assertIn("NAV-kommunepuls: August 2026", xml)
        self.assertIn("table.xlsx", xml)
        self.assertIn("Orkland", xml)
        self.assertIn("+4 fra forrige måned", xml)


if __name__ == "__main__":
    unittest.main()
