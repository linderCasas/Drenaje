import unittest

from proyecto_drenaje.web.app import app, load_report_data


class MachineryReportTests(unittest.TestCase):
    def test_reads_report_metrics_from_workbook(self) -> None:
        report = load_report_data()

        self.assertEqual(report.cost_total, "S/ 928.529,50")
        self.assertEqual(report.hours, "2.736,9 h")
        self.assertEqual(report.pending, "S/ 176.180,50")
        self.assertEqual(report.scheduled, 8)
        self.assertEqual(report.operational, 3)
        self.assertEqual(report.availability, "37,5%")

    def test_home_page_renders_workbook_data(self) -> None:
        response = app.test_client().get("/")

        self.assertEqual(response.status_code, 200)
        self.assertIn("S/ 928.529,50".encode(), response.data)
        self.assertIn("37,5%".encode(), response.data)
        self.assertIn("Actualizar reporte".encode(), response.data)


if __name__ == "__main__":
    unittest.main()
