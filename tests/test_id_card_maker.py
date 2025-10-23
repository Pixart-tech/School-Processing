import math
import types
import unittest
from pathlib import Path
from xml.dom.minidom import Document

from PIL import ImageFont

doc_maker_stub = types.ModuleType("doc_maker")
doc_maker_stub.callInkscape = lambda *args, **kwargs: None

import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

sys.modules.setdefault("doc_maker", doc_maker_stub)

from id_card_maker import (
    MIN_FONT_SIZE,
    MULTILINE_MIN_FONT_SIZE,
    _extract_font_size,
    _measure_text_width,
    _parse_length,
    _resolve_font_path,
    _update_text_group,
)


class MultilineFallbackTests(unittest.TestCase):
    def _collect_lines(self, text_element):
        lines = []
        first = text_element.firstChild
        if first is not None and first.nodeType == first.TEXT_NODE:
            lines.append(first.data)
        for tspan in text_element.getElementsByTagName("tspan"):
            if tspan.firstChild is not None:
                lines.append(tspan.firstChild.data)
            else:
                lines.append("")
        return lines

    def test_overflow_text_shrinks_within_rect(self):
        doc = Document()
        svg = doc.createElement("svg")
        doc.appendChild(svg)

        group = doc.createElement("g")
        group.setAttribute("id", "name")
        svg.appendChild(group)

        rect = doc.createElement("rect")
        rect.setAttribute("x", "0")
        rect.setAttribute("y", "0")
        rect.setAttribute("width", "70")
        rect.setAttribute("height", "40")
        group.appendChild(rect)

        text_element = doc.createElement("text")
        text_element.setAttribute(
            "style", "font-size:38px;font-family:PlaypenSans-Medium"
        )
        text_element.setAttribute("x", "35")
        text_element.setAttribute("y", "20")
        text_element.setAttribute("text-anchor", "middle")
        text_element.appendChild(doc.createTextNode("AB"))
        group.appendChild(text_element)

        long_text = "Alexandria Maximillian Robertson"
        _update_text_group(group, long_text)

        lines = self._collect_lines(text_element)
        self.assertGreaterEqual(len(lines), 2, "Two-line fallback was not applied")

        final_font_size = _extract_font_size(text_element)
        self.assertIsNotNone(final_font_size)
        self.assertGreaterEqual(final_font_size, MULTILINE_MIN_FONT_SIZE)

        font_path = _resolve_font_path(text_element)
        self.assertTrue(font_path.exists())

        font = ImageFont.truetype(
            str(font_path), max(1, int(math.floor(final_font_size)))
        )
        measured_width = max(_measure_text_width(font, line) for line in lines)
        rect_width = float(rect.getAttribute("width"))

        self.assertLessEqual(
            measured_width,
            rect_width + 0.5,
            "Multiline shrink routine did not fit text within rectangle",
        )

        self.assertLess(
            final_font_size,
            MIN_FONT_SIZE,
            "Expected multiline shrink routine to reduce below default minimum",
        )

        self.assertEqual(text_element.getAttribute("text-anchor"), "middle")

        original_y = 20.0
        first_baseline = _parse_length(text_element.getAttribute("y"))
        self.assertIsNotNone(first_baseline)
        baselines = [first_baseline]
        for tspan in text_element.getElementsByTagName("tspan"):
            dy = _parse_length(tspan.getAttribute("dy"))
            self.assertIsNotNone(dy)
            baselines.append(baselines[-1] + dy)

        self.assertGreaterEqual(len(baselines), 2, "Expected multiline layout baselines")
        centered_baseline = (baselines[0] + baselines[-1]) / 2.0
        self.assertAlmostEqual(centered_baseline, original_y, places=2)


class TransformPreservationTests(unittest.TestCase):
    def _build_group(self):
        doc = Document()
        svg = doc.createElement("svg")
        doc.appendChild(svg)

        group = doc.createElement("g")
        group.setAttribute("id", "name")
        svg.appendChild(group)

        rect = doc.createElement("rect")
        rect.setAttribute("x", "0")
        rect.setAttribute("y", "0")
        rect.setAttribute("width", "70")
        rect.setAttribute("height", "40")
        group.appendChild(rect)

        text_element = doc.createElement("text")
        text_element.setAttribute(
            "style", "font-size:38px;font-family:PlaypenSans-Medium"
        )
        text_element.appendChild(doc.createTextNode("AB"))
        group.appendChild(text_element)

        return group, text_element

    def test_translate_offsets_preserve_coordinates_and_other_transforms(self):
        group, text_element = self._build_group()
        text_element.setAttribute("x", "12")
        text_element.setAttribute("y", "30")
        text_element.setAttribute("transform", "translate(5,-7) rotate(30)")

        _update_text_group(group, "Noah")

        self.assertAlmostEqual(
            _parse_length(text_element.getAttribute("x")),
            17.0,
            places=4,
        )
        self.assertAlmostEqual(
            _parse_length(text_element.getAttribute("y")),
            23.0,
            places=4,
        )
        self.assertEqual(text_element.getAttribute("transform"), "rotate(30)")

    def test_pure_translate_transform_converted_to_coordinates(self):
        group, text_element = self._build_group()
        text_element.setAttribute("transform", "translate(8, 12)")

        _update_text_group(group, "Lia")

        self.assertAlmostEqual(
            _parse_length(text_element.getAttribute("x")),
            8.0,
            places=4,
        )
        self.assertAlmostEqual(
            _parse_length(text_element.getAttribute("y")),
            12.0,
            places=4,
        )
        self.assertFalse(text_element.hasAttribute("transform"))

    def test_single_line_restores_alignment_and_size(self):
        group, text_element = self._build_group()
        text_element.setAttribute("x", "35")
        text_element.setAttribute("y", "20")
        text_element.setAttribute("text-anchor", "middle")
        original_font_size = _extract_font_size(text_element)
        text_element.firstChild.data = "Template Name"

        _update_text_group(group, "Lia")

        self.assertEqual(text_element.getAttribute("x"), "35")
        self.assertEqual(text_element.getAttribute("text-anchor"), "middle")
        self.assertEqual(text_element.getAttribute("y"), "20")
        self.assertAlmostEqual(_extract_font_size(text_element), original_font_size)
        self.assertEqual(text_element.getAttribute("dominant-baseline"), "alphabetic")
        self.assertEqual(len(text_element.getElementsByTagName("tspan")), 0)


if __name__ == "__main__":
    unittest.main()
