import unittest
from dataclasses import dataclass, field
from pathlib import Path
from typing import List
from unittest.mock import Mock, patch

from order_automation import (
    ORIENTATION_HORIZONTAL,
    ORIENTATION_VERTICAL,
    OrderAutomation,
    OrderAutomationError,
    build_size_mappings,
    determine_orientation,
    expand_ui_text_variants,
    expand_size_text_variants,
    extract_match_key_from_folder_name,
    extract_match_key_from_remark,
    find_size_mapping,
    index_local_assets,
    infer_orientation_from_image,
    resolve_asset_from_remark,
)


@dataclass(order=True)
class FakePath:
    name: str
    is_dir_value: bool = False
    is_file_value: bool = False
    children: List["FakePath"] = field(default_factory=list, compare=False)

    def exists(self) -> bool:
        return True

    def is_dir(self) -> bool:
        return self.is_dir_value

    def is_file(self) -> bool:
        return self.is_file_value

    def iterdir(self):
        return iter(self.children)

    def __str__(self) -> str:
        return self.name


def make_file(name: str) -> FakePath:
    return FakePath(name=name, is_file_value=True)


def make_folder(name: str, children: List[FakePath]) -> FakePath:
    return FakePath(name=name, is_dir_value=True, children=children)


class FakeImage:
    def __init__(self, size):
        self.size = size

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return None


class FakeLocator:
    def __init__(self, count: int, visible: bool = False, text: str = ""):
        self._count = count
        self._visible = visible
        self._text = text

    def count(self) -> int:
        return self._count

    @property
    def first(self):
        return self

    def is_visible(self) -> bool:
        return self._visible

    def text_content(self) -> str:
        return self._text


class FakePage:
    def __init__(self, states):
        self.states = list(states)
        self.state_index = 0
        self.wait_calls = 0

    def locator(self, selector: str):
        state = self.states[min(self.state_index, len(self.states) - 1)]
        count, visible, text = state.get(selector, (0, False, ""))
        return FakeLocator(count, visible, text)

    def wait_for_timeout(self, _ms: int) -> None:
        self.wait_calls += 1
        if self.state_index < len(self.states) - 1:
            self.state_index += 1

    def wait_for_load_state(self, _state: str, timeout: int | None = None) -> None:
        return None


class FakeClickLocator(FakeLocator):
    def __init__(self, count: int, visible: bool = False, text: str = ""):
        super().__init__(count, visible, text)
        self.clicked = False

    def nth(self, _index: int):
        return self

    def click(self, force: bool = False) -> None:
        self.clicked = True


class FakeRow:
    def __init__(self, mapping):
        self.mapping = mapping

    def get_by_text(self, text: str, exact: bool = False):
        return self.mapping.get(text, FakeClickLocator(0, False, ""))


class FakeAncestorLocator(FakeClickLocator):
    def locator(self, _selector: str):
        return self


class FakeButtonLocator(FakeClickLocator):
    def __init__(self, count: int, visible: bool = False, text: str = "", ancestor: FakeAncestorLocator | None = None):
        super().__init__(count, visible, text)
        self.ancestor = ancestor

    def nth(self, _index: int):
        return self

    def locator(self, _selector: str):
        return self.ancestor or FakeClickLocator(0, False, "")


class FakeRowsLocator:
    def __init__(self, rows):
        self._rows = rows

    def count(self):
        return len(self._rows)

    def nth(self, index: int):
        return self._rows[index]


class OrderAutomationTests(unittest.TestCase):
    def make_automation(self) -> OrderAutomation:
        config = {
            "base_order_dir": ".",
            "target_url": "https://example.com",
            "store_name": "Amazon_Beilo_US",
            "browser_profile_dir": ".order_automation/browser_profile",
            "size_mappings": [
                {
                    "external_spec_pattern": '8X10"',
                    "model_name": "帆布画4:5（多尺寸）",
                    "size_option_text": "8x10inch",
                }
            ],
        }
        return OrderAutomation(config=config, logger=Mock(), date_str="2026.5.7", headful=True)

    def test_extract_match_key_from_folder_name(self) -> None:
        folder_name = "1-111-4243692-4662650-001-8x10-None-横图"
        self.assertEqual(extract_match_key_from_folder_name(folder_name), "4662650-001")

    def test_extract_match_key_from_resend_folder_name(self) -> None:
        folder_name = "1-resend-114-1771439-0493831-001-16x24"
        self.assertEqual(extract_match_key_from_folder_name(folder_name), "0493831-001")

    def test_determine_orientation_from_folder_name(self) -> None:
        self.assertEqual(determine_orientation("abc-横图"), ORIENTATION_HORIZONTAL)
        self.assertEqual(determine_orientation("abc-竖图"), ORIENTATION_VERTICAL)

    @patch("order_automation.Image.open", return_value=FakeImage((1600, 900)))
    def test_determine_orientation_falls_back_to_image_for_horizontal(self, _mock_open) -> None:
        orientation = determine_orientation("1-resend-114-1771439-0493831-001-16x24", Path("demo.jpg"))
        self.assertEqual(orientation, ORIENTATION_HORIZONTAL)

    @patch("order_automation.Image.open", return_value=FakeImage((900, 1600)))
    def test_infer_orientation_from_image_for_vertical(self, _mock_open) -> None:
        orientation = infer_orientation_from_image(Path("demo.jpg"))
        self.assertEqual(orientation, ORIENTATION_VERTICAL)

    def test_extract_match_key_from_remark_prefers_composite_key(self) -> None:
        remark = "用户备注 0505B4662650-001"
        self.assertEqual(extract_match_key_from_remark(remark), "4662650-001")

    def test_find_size_mapping(self) -> None:
        mappings = build_size_mappings(
            [
                {
                    "external_spec_pattern": '8X10"',
                    "model_name": "帆布画4:5（多尺寸）",
                    "size_option_text": "8x10inch",
                }
            ]
        )
        mapping = find_size_mapping('外部订单规格：8X10"', mappings)
        self.assertEqual(mapping.model_name, "帆布画4:5（多尺寸）")
        self.assertEqual(mapping.size_option_text, "8x10inch")

    @patch("order_automation.Image.open", return_value=FakeImage((1500, 900)))
    def test_index_local_assets_success(self, _mock_open) -> None:
        folder = make_folder(
            "1-111-4243692-4662650-001-8x10-None-横图",
            [make_file("image-processed.jpg")],
        )
        day_dir = make_folder("2026.5.5", [folder])

        assets = index_local_assets(day_dir)

        self.assertIn("4662650-001", assets)
        asset = assets["4662650-001"]
        self.assertEqual(asset.orientation, ORIENTATION_HORIZONTAL)
        self.assertEqual(asset.processed_image_path.name, "image-processed.jpg")

    @patch("order_automation.Image.open", side_effect=[FakeImage((1500, 900)), FakeImage((900, 1500))])
    def test_index_local_assets_allows_same_id_with_different_sequence(self, _mock_open) -> None:
        first = make_folder(
            "1-111-4243692-4662650-001-8x10-None-横图",
            [make_file("a-processed.jpg")],
        )
        second = make_folder(
            "1-111-5555555-4662650-002-8x10-None-竖图",
            [make_file("b-processed.jpg")],
        )
        day_dir = make_folder("2026.5.5", [first, second])

        assets = index_local_assets(day_dir)

        self.assertIn("4662650-001", assets)
        self.assertIn("4662650-002", assets)

    @patch("order_automation.Image.open", return_value=FakeImage((900, 1500)))
    def test_index_local_assets_uses_image_orientation_when_folder_suffix_missing(self, _mock_open) -> None:
        folder = make_folder(
            "1-resend-114-1771439-0493831-001-16x24",
            [make_file("image-processed.jpg")],
        )
        day_dir = make_folder("2026.5.5", [folder])

        assets = index_local_assets(day_dir)

        self.assertEqual(assets["0493831-001"].orientation, ORIENTATION_VERTICAL)

    @patch("order_automation.Image.open", return_value=FakeImage((900, 1500)))
    def test_index_local_assets_ignores_non_order_folders(self, _mock_open) -> None:
        order_folder = make_folder(
            "1-resend-114-1771439-0493831-001-16x24",
            [make_file("image-processed.jpg")],
        )
        helper_folder = make_folder("Canvas Print", [make_file("readme.txt")])
        day_dir = make_folder("2026.5.5", [order_folder, helper_folder])

        assets = index_local_assets(day_dir)

        self.assertEqual(list(assets.keys()), ["0493831-001"])

    def test_index_local_assets_skips_folder_with_multiple_processed_images(self) -> None:
        folder = make_folder(
            "1-111-4243692-4662650-001-8x10-None-横图",
            [make_file("a-processed.jpg"), make_file("b-processed.jpg")],
        )
        valid_folder = make_folder(
            "1-111-5555555-4662650-002-8x10-None-竖图",
            [make_file("ok-processed.jpg")],
        )
        day_dir = make_folder("2026.5.5", [folder, valid_folder])

        assets = index_local_assets(day_dir)

        self.assertEqual(list(assets.keys()), ["4662650-002"])

    def test_resolve_asset_from_remark_with_composite_key(self) -> None:
        first = make_folder(
            "1-111-4243692-4662650-001-8x10-None-横图",
            [make_file("a-processed.jpg")],
        )
        second = make_folder(
            "1-111-5555555-4662650-002-8x10-None-竖图",
            [make_file("b-processed.jpg")],
        )
        assets = index_local_assets(make_folder("2026.5.5", [first, second]))

        asset = resolve_asset_from_remark("用户备注 0505B4662650-002", assets)

        self.assertEqual(asset.match_key, "4662650-002")

    def test_resolve_asset_from_remark_rejects_ambiguous_id_only(self) -> None:
        first = make_folder(
            "1-111-4243692-4662650-001-8x10-None-横图",
            [make_file("a-processed.jpg")],
        )
        second = make_folder(
            "1-111-5555555-4662650-002-8x10-None-竖图",
            [make_file("b-processed.jpg")],
        )
        assets = index_local_assets(make_folder("2026.5.5", [first, second]))

        with self.assertRaises(OrderAutomationError):
            resolve_asset_from_remark("用户备注 4662650", assets)

    def test_find_size_mapping_raises_for_unknown_spec(self) -> None:
        mappings = build_size_mappings(
            [
                {
                    "external_spec_pattern": '8X10"',
                    "model_name": "帆布画4:5（多尺寸）",
                    "size_option_text": "8x10inch",
                }
            ]
        )
        with self.assertRaises(OrderAutomationError):
            find_size_mapping('外部订单规格：11X14"', mappings)

    def test_is_login_required_v2_detects_login_form(self) -> None:
        automation = self.make_automation()
        page = Mock()
        page.locator.side_effect = [FakeLocator(1, True)]

        self.assertTrue(automation.is_login_required_v2(page))

    def test_is_login_required_v2_detects_login_register_header(self) -> None:
        automation = self.make_automation()
        page = Mock()
        page.locator.side_effect = [
            FakeLocator(0, False),
            FakeLocator(0, False),
            FakeLocator(1, True, "登录 / 注册"),
        ]

        self.assertTrue(automation.is_login_required_v2(page))

    def test_wait_for_manual_login_if_needed_allows_page_when_login_is_absent(self) -> None:
        automation = self.make_automation()
        page = Mock()
        page.locator.side_effect = [
            FakeLocator(0, False),
            FakeLocator(0, False),
            FakeLocator(0, False),
            FakeLocator(0, False),
            FakeLocator(0, False),
        ]

        automation.wait_for_manual_login_if_needed(page)

    def test_wait_for_manual_login_if_needed_waits_for_manual_login(self) -> None:
        automation = self.make_automation()
        page = FakePage(
            [
                {
                    ".login-wrap": (0, False, ""),
                    ".login-form": (0, False, ""),
                    ".header_right_userbox .user-name": (1, True, "登录 / 注册"),
                    "input[type='password']": (0, False, ""),
                    "button:has-text('登录')": (0, False, ""),
                },
                {
                    ".login-wrap": (0, False, ""),
                    ".login-form": (0, False, ""),
                    ".header_right_userbox .user-name": (1, True, "User"),
                    "input[type='password']": (0, False, ""),
                    "button:has-text('登录')": (0, False, ""),
                },
            ]
        )

        automation.wait_for_manual_login_if_needed(page, wait_timeout_seconds=5)

        self.assertEqual(page.wait_calls, 1)

    @patch("builtins.input", return_value="")
    def test_wait_for_manual_login_confirmation_prompts_and_continues(self, _mock_input) -> None:
        automation = self.make_automation()
        page = Mock()
        page.wait_for_load_state.return_value = None

        automation.wait_for_manual_login_confirmation(page)

        page.wait_for_load_state.assert_called()

    def test_wait_for_manual_login_confirmation_skips_in_headless_mode(self) -> None:
        automation = self.make_automation()
        automation.headful = False
        page = Mock()

        automation.wait_for_manual_login_confirmation(page)

        page.wait_for_load_state.assert_not_called()

    def test_row_has_customize_action_accepts_alternate_button_text(self) -> None:
        automation = self.make_automation()
        row = FakeRow({"去定制": FakeClickLocator(1, True, "去定制")})

        self.assertTrue(automation.row_has_customize_action(row))

    def test_open_customization_modal_clicks_first_matching_action(self) -> None:
        automation = self.make_automation()
        locator = FakeClickLocator(1, True, "去定制")
        row = FakeRow({"去定制": locator})

        automation.open_customization_modal(row)

        self.assertTrue(locator.clicked)

    def test_find_row_from_action_buttons_returns_ancestor_container(self) -> None:
        automation = self.make_automation()
        ancestor = FakeAncestorLocator(1, True, "用户备注 0505B4752234-001")
        button = FakeButtonLocator(1, True, "去定制", ancestor=ancestor)
        page = Mock()
        page.get_by_text.side_effect = [button]

        row = automation.find_row_from_action_buttons(page)

        self.assertIs(row, ancestor)


    def test_expand_ui_text_variants_supports_fullwidth_punctuation(self) -> None:
        variants = expand_ui_text_variants("帆布画4:5（多尺寸）")

        self.assertIn("帆布画4：5（多尺寸）", variants)
        self.assertIn("帆布画4：5", variants)

    def test_locator_by_text_tries_text_variants(self) -> None:
        automation = self.make_automation()
        page = Mock()

        def fake_get_by_text(text: str, exact: bool = False):
            if text == "帆布画4：5（多尺寸）":
                return FakeClickLocator(1, True, text)
            return FakeClickLocator(0, False, text)

        page.get_by_text.side_effect = fake_get_by_text

        locator = automation.locator_by_text(page, ["帆布画4:5（多尺寸）"])

        self.assertEqual(locator.text_content(), "帆布画4：5（多尺寸）")

    def test_locator_by_text_prefers_visible_match(self) -> None:
        automation = self.make_automation()
        hidden = FakeClickLocator(1, False, "设计")
        visible = FakeClickLocator(1, True, "设计")

        class FakeMultiLocator:
            def __init__(self, locators):
                self.locators = locators

            def count(self):
                return len(self.locators)

            def nth(self, index: int):
                return self.locators[index]

            @property
            def first(self):
                return self.locators[0]

        page = Mock()
        page.get_by_text.side_effect = lambda text, exact=False: FakeMultiLocator([hidden, visible])

        locator = automation.locator_by_text(page, ["设计"])

        self.assertIs(locator, visible)

    def test_expand_size_text_variants_supports_multiplication_symbol(self) -> None:
        variants = expand_size_text_variants("16x20inch")

        self.assertIn("16×20inch", variants)
        self.assertIn("16X20inch", variants)

    def test_try_select_model_row_clicks_choose_button(self) -> None:
        automation = self.make_automation()
        choose = FakeClickLocator(1, True, "选择")
        row = Mock()
        row.text_content.return_value = "帆布画4：5（多尺寸） 16×20inch 选择"

        def fake_get_by_text(text: str, exact: bool = False):
            if text == "选择":
                return choose
            return FakeClickLocator(0, False, text)

        row.get_by_text.side_effect = fake_get_by_text
        page = Mock()
        page.locator.return_value = FakeRowsLocator([row])

        mapping = build_size_mappings(
            [
                {
                    "external_spec_pattern": '16X20"',
                    "model_name": "帆布画4：5（多尺寸）",
                    "size_option_text": "16x20inch",
                }
            ]
        )[0]

        result = automation.try_select_model_row(page, mapping)

        self.assertTrue(result)
        self.assertTrue(choose.clicked)


if __name__ == "__main__":
    unittest.main()
