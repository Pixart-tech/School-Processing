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
    MULTILINE_MIN_FONT_SIZE,
    _apply_alignment,
    _extract_font_size,
    _build_school_verification_label,
    _measure_text_width,
    _parse_length,
    _extract_outer_code_prefix,
    _resolve_font_path,
    _update_text_group,
    _resolve_address_value,
    _update_address_group,
)


class MultilineFallbackTests(unittest.TestCase):
    def test_middle_alignment_keyword_sets_expected_anchor(self):
        doc = Document()
        svg = doc.createElement("svg")
        doc.appendChild(svg)

        text_element = doc.createElement("text")
        text_element.setAttribute("x", "10")
        svg.appendChild(text_element)

        _apply_alignment(text_element, "middle")

        self.assertEqual(text_element.getAttribute("text-anchor"), "middle")


class VerificationLabelTests(unittest.TestCase):
    def test_extract_outer_code_prefix_prefers_digits(self):
        self.assertEqual(_extract_outer_code_prefix("1234567b"), "123")

    def test_extract_outer_code_prefix_pads_short_values(self):
        self.assertEqual(_extract_outer_code_prefix("9"), "009")

    def test_build_label_with_prefix(self):
        label = _build_school_verification_label("Sunrise Public School", "123")
        self.assertEqual(label, "123_Sunrise_Public_School")

    def test_build_label_without_prefix(self):
        label = _build_school_verification_label("Sunrise Public School", None)
        self.assertEqual(label, "Sunrise_Public_School")

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

    def test_rect_width_used_when_template_text_is_short(self):
        doc = Document()
        svg = doc.createElement("svg")
        doc.appendChild(svg)

        group = doc.createElement("g")
        group.setAttribute("id", "name")
        svg.appendChild(group)

        rect = doc.createElement("rect")
        rect.setAttribute("x", "0")
        rect.setAttribute("y", "0")
        rect.setAttribute("width", "120")
        rect.setAttribute("height", "40")
        group.appendChild(rect)

        text_element = doc.createElement("text")
        text_element.setAttribute(
            "style", "font-size:38px;font-family:PlaypenSans-Medium"
        )
        text_element.setAttribute("x", "60")
        text_element.setAttribute("y", "20")
        text_element.setAttribute("text-anchor", "middle")
        text_element.appendChild(doc.createTextNode("X"))
        group.appendChild(text_element)

        long_text = "Alexandria Maximillian Robertson"
        _update_text_group(group, long_text)

        lines = self._collect_lines(text_element)
        self.assertGreaterEqual(len(lines), 2, "Two-line fallback was not applied")

        final_font_size = _extract_font_size(text_element)
        self.assertIsNotNone(final_font_size)
        self.assertGreater(final_font_size, 10.0)

        font_path = _resolve_font_path(text_element)
        font = ImageFont.truetype(
            str(font_path), max(1, int(math.floor(final_font_size)))
        )
        measured_width = max(_measure_text_width(font, line) for line in lines)

        rect_width = float(rect.getAttribute("width"))
        self.assertLessEqual(
            measured_width,
            rect_width + 0.5,
            "Adjusted text width should respect enclosing rectangle",
        )

    def test_center_alignment_uses_measured_text_width_for_centering(self):
        doc = Document()
        svg = doc.createElement("svg")
        doc.appendChild(svg)

        group = doc.createElement("g")
        group.setAttribute("inkscape:label", "M name")
        svg.appendChild(group)

        rect = doc.createElement("rect")
        rect.setAttribute("x", "30")
        rect.setAttribute("y", "0")
        rect.setAttribute("width", "140")
        rect.setAttribute("height", "40")
        group.appendChild(rect)

        text_element = doc.createElement("text")
        text_element.setAttribute(
            "style", "font-size:38px;font-family:PlaypenSans-Medium"
        )
        text_element.setAttribute("x", "30")
        text_element.setAttribute("y", "20")
        text_element.setAttribute("text-anchor", "start")
        text_element.appendChild(doc.createTextNode("Centered"))
        group.appendChild(text_element)

        _update_text_group(group, "Centered")

        self.assertEqual(text_element.getAttribute("text-anchor"), "middle")
        final_x = _parse_length(text_element.getAttribute("x"))
        self.assertIsNotNone(final_x)

        font_size = _extract_font_size(text_element)
        self.assertIsNotNone(font_size)
        font_path = _resolve_font_path(text_element)
        self.assertTrue(font_path.exists())
        font = ImageFont.truetype(
            str(font_path), max(1, int(math.floor(font_size)))
        )
        measured_width = _measure_text_width(font, "Centered")
        expected_center = 30.0 + (measured_width / 2.0)

        self.assertAlmostEqual(final_x, expected_center, places=4)


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


class AddressResolutionTests(unittest.TestCase):
    def test_prefers_new_address_field(self):
        record = {"address": "New Address", "current_address": "Old Address"}
        self.assertEqual(_resolve_address_value(record), "New Address")

    def test_falls_back_to_current_address(self):
        record = {"current_address": "Legacy Address"}
        self.assertEqual(_resolve_address_value(record), "Legacy Address")

    def test_handles_missing_values(self):
        record = {"address": "", "current_address": None}
        self.assertEqual(_resolve_address_value(record), "")


class AddressLayoutTests(unittest.TestCase):
    def test_middle_aligned_fallback_preserves_center_position(self):
        doc = Document()
        svg = doc.createElement("svg")
        doc.appendChild(svg)

        group = doc.createElement("g")
        group.setAttribute("id", "address_middle")
        svg.appendChild(group)

        text_element = doc.createElement("text")
        text_element.setAttribute("style", "font-size:10px")
        group.appendChild(text_element)

        for index in range(4):
            tspan = doc.createElement("tspan")
            tspan.setAttribute("x", "100")
            if index > 0:
                tspan.setAttribute("dy", "5")
            tspan.appendChild(doc.createTextNode(f"Template Line {index+1}"))
            text_element.appendChild(tspan)

        _update_address_group(group, "First Line\nSecond Line")

        self.assertEqual(text_element.getAttribute("text-anchor"), "middle")

        tspans = text_element.getElementsByTagName("tspan")
        self.assertEqual(tspans.length, 2)
        for tspan in tspans:
            self.assertEqual(tspan.getAttribute("x"), "100")


if __name__ == "__main__":
    unittest.main()
