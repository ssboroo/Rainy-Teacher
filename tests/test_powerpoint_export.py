import tempfile
import unittest
from pathlib import Path
from io import BytesIO

from PIL import Image
from pptx import Presentation
from powerpoint_export import save_certificate_pptx, save_presentation_pptx


class PowerPointExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def test_presentation_is_editable_and_preserves_every_point(self):
        deck = {
            "title": "Монгол хэлний хичээл",
            "subject": "Монгол хэл · 5-р анги",
            "slides": [{"title": "Эгшиг авиаг таних", "points": [
                "Эгшиг авиаг ялгана", "Үгийн үеийг тоолно", "Жишээгээр бататгана",
                "Хосоор ажиллана", "Дүгнэлт гаргана",
            ]}],
        }
        target = self.folder / "lesson.pptx"
        save_presentation_pptx(deck, target, {
            "accent": "#1C5D9D", "bg": "#F5F8FC", "ink": "#102B48",
            "soft": "#DCEAF7", "mode": "steppe",
        })
        reopened = Presentation(target)
        self.assertEqual(len(reopened.slides), 3)  # cover plus two editable content slides
        text = "\n".join(shape.text for slide in reopened.slides for shape in slide.shapes if shape.has_text_frame)
        for point in deck["slides"][0]["points"]:
            self.assertIn(point, text)
        self.assertEqual(reopened.slide_width, int(13.333 * 914400))

    def test_certificate_exports_editable_mongolian_text_and_frame(self):
        logo = BytesIO()
        Image.new("RGBA", (64, 64), (25, 60, 120, 255)).save(logo, format="PNG")
        record = {
            "style": "navy", "title": "ӨРГӨМЖЛӨЛ", "subtitle": "Амжилтаа үнэлүүлэв",
            "recipient": "Бат-Эрдэнэ", "achievement": "Улсын олимпиадад амжилттай оролцсон",
            "body": "{achievement} амжилт гаргасныг үнэлж {name}-д гардуулав.",
            "organization": "Эрдмийн оргил сургууль", "teacher": "Б. Саруул",
            "issue_date": "2026-10-15", "logo": logo.getvalue(), "logo_mime": "image/png",
        }
        target = self.folder / "certificate.pptx"
        save_certificate_pptx(record, target, {
            "navy": {"colors": ("#183F73", "#EFDFB3", "#FFFAF0")},
            "gold": {"colors": ("#A77920", "#F5E8BC", "#FFFAF0")},
        })
        reopened = Presentation(target)
        self.assertEqual(len(reopened.slides), 1)
        slide = reopened.slides[0]
        text = "\n".join(shape.text for shape in slide.shapes if shape.has_text_frame)
        self.assertIn("ӨРГӨМЖЛӨЛ", text)
        self.assertIn("Бат-Эрдэнэ", text)
        self.assertIn("Эрдмийн оргил сургууль", text)
        self.assertGreaterEqual(len(slide.shapes), 35)  # editable frame, corner ornaments, rules, and text
        pictures = [shape for shape in slide.shapes if shape.shape_type == 13]
        self.assertEqual(len(pictures), 1)
        self.assertAlmostEqual(pictures[0].width / pictures[0].height, 1.0, places=2)


if __name__ == "__main__":
    unittest.main()
