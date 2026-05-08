from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

from PIL import Image

try:
    from playwright.sync_api import BrowserContext, Error as PlaywrightError, Page, sync_playwright
except ImportError:  # pragma: no cover
    BrowserContext = Any  # type: ignore[assignment]
    Page = Any  # type: ignore[assignment]
    PlaywrightError = Exception  # type: ignore[assignment]
    sync_playwright = None  # type: ignore[assignment]


DEFAULT_CONFIG_PATH = Path(__file__).with_name("order_automation.config.json")
IGNORED_FOLDERS = {"new", "output", "automation_logs"}
MATCH_KEY_PATTERN = re.compile(r"(?<!\d)(\d{7}-\d{3})(?!\d)")
MATCH_ID_PATTERN = re.compile(r"(?<!\d)(\d{7})(?!\d)")
FOLDER_MATCH_KEY_PATTERN = re.compile(r"(\d{7}-\d{3})(?=-|$)")
ORIENTATION_HORIZONTAL = "horizontal"
ORIENTATION_VERTICAL = "vertical"


class OrderAutomationError(RuntimeError):
    pass


@dataclass(frozen=True)
class SizeMapping:
    external_spec_pattern: str
    model_name: str
    size_option_text: str


@dataclass(frozen=True)
class LocalOrderAsset:
    match_key: str
    folder_path: Path
    processed_image_path: Path
    orientation: str
    folder_name: str


def str_to_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    raise argparse.ArgumentTypeError(f"Invalid boolean value: {value}")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Automate Canvro order customization workflow.")
    parser.add_argument("--date", default=None, help="Target date folder in yyyy.M.d format.")
    parser.add_argument("--headful", type=str_to_bool, default=True, help="Show browser window.")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="Path to config JSON.")
    args = parser.parse_args(argv)
    if args.date is None:
        now = datetime.now()
        args.date = f"{now.year}.{now.month}.{now.day}"
    return args


def load_config(config_path: Path) -> Dict[str, Any]:
    if not config_path.exists():
        raise OrderAutomationError(f"Configuration file not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as file:
        config = json.load(file)

    required_keys = {
        "base_order_dir",
        "target_url",
        "store_name",
        "browser_profile_dir",
        "size_mappings",
    }
    missing_keys = sorted(required_keys - set(config))
    if missing_keys:
        raise OrderAutomationError(f"Configuration is missing keys: {', '.join(missing_keys)}")
    if not isinstance(config["size_mappings"], list) or not config["size_mappings"]:
        raise OrderAutomationError("Configuration key 'size_mappings' must be a non-empty list.")
    return config


def setup_logger(log_dir: Path) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"order_automation_{timestamp}.log"

    logger = logging.getLogger("order_automation")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    logger.info("Logging to %s", log_file)
    return logger


def extract_match_key_from_folder_name(folder_name: str) -> str:
    matches = FOLDER_MATCH_KEY_PATTERN.findall(folder_name)
    if not matches:
        raise OrderAutomationError(f"Cannot extract match key from folder name: {folder_name}")
    return matches[-1]


def infer_orientation_from_image(processed_image_path: Path) -> str:
    try:
        with Image.open(processed_image_path) as image:
            width, height = image.size
    except Exception as exc:  # pragma: no cover
        raise OrderAutomationError(
            f"Cannot read image dimensions from processed image: {processed_image_path}"
        ) from exc
    return ORIENTATION_HORIZONTAL if width >= height else ORIENTATION_VERTICAL


def determine_orientation(folder_name: str, processed_image_path: Optional[Path] = None) -> str:
    if "横图" in folder_name:
        return ORIENTATION_HORIZONTAL
    if "竖图" in folder_name:
        return ORIENTATION_VERTICAL
    if processed_image_path is not None:
        return infer_orientation_from_image(processed_image_path)
    raise OrderAutomationError(f"Cannot determine orientation from folder name: {folder_name}")


def find_processed_image(folder_path: Path) -> Path:
    matches = [
        path for path in folder_path.iterdir()
        if path.is_file() and "processed" in path.name.lower()
    ]
    if not matches:
        raise OrderAutomationError(f"No processed image found in folder: {folder_path}")
    if len(matches) > 1:
        raise OrderAutomationError(f"Multiple processed images found in folder: {folder_path}")
    return matches[0]


def index_local_assets(day_dir: Path) -> Dict[str, LocalOrderAsset]:
    if not day_dir.exists():
        raise OrderAutomationError(f"Date folder not found: {day_dir}")

    assets: Dict[str, LocalOrderAsset] = {}
    for folder_path in sorted(path for path in day_dir.iterdir() if path.is_dir()):
        if folder_path.name in IGNORED_FOLDERS:
            continue
        matches = FOLDER_MATCH_KEY_PATTERN.findall(folder_path.name)
        if not matches:
            continue
        match_key = matches[-1]
        if match_key in assets:
            raise OrderAutomationError(
                f"Duplicate match key {match_key} found in folders: "
                f"{assets[match_key].folder_path} and {folder_path}"
            )
        processed_image_path = find_processed_image(folder_path)
        assets[match_key] = LocalOrderAsset(
            match_key=match_key,
            folder_path=folder_path,
            processed_image_path=processed_image_path,
            orientation=determine_orientation(folder_path.name, processed_image_path),
            folder_name=folder_path.name,
        )

    if not assets:
        raise OrderAutomationError(f"No order folders found under {day_dir}")
    return assets


def extract_match_key_from_remark(remark: str) -> str:
    composite_matches = MATCH_KEY_PATTERN.findall(remark)
    if composite_matches:
        return max(composite_matches, key=len)
    id_matches = MATCH_ID_PATTERN.findall(remark)
    if id_matches:
        return max(id_matches, key=len)
    raise OrderAutomationError(f"Cannot extract match key from remark: {remark}")


def resolve_asset_from_remark(remark: str, assets: Dict[str, LocalOrderAsset]) -> LocalOrderAsset:
    extracted_key = extract_match_key_from_remark(remark)
    if extracted_key in assets:
        return assets[extracted_key]

    if "-" not in extracted_key:
        matching_assets = [asset for key, asset in assets.items() if key.startswith(f"{extracted_key}-")]
        if len(matching_assets) == 1:
            return matching_assets[0]
        if len(matching_assets) > 1:
            raise OrderAutomationError(
                f"Remark {remark} only contains ID {extracted_key}, but multiple local folders match it."
            )
    raise OrderAutomationError(f"Local folder not found for remark {remark} using key {extracted_key}")


def normalize_spec(spec_text: str) -> str:
    normalized = spec_text.upper().replace("INCH", '"').replace(" ", "")
    normalized = normalized.replace("”", '"').replace("“", '"')
    return normalized


def build_size_mappings(entries: Iterable[Dict[str, str]]) -> List[SizeMapping]:
    mappings: List[SizeMapping] = []
    for entry in entries:
        try:
            mappings.append(
                SizeMapping(
                    external_spec_pattern=entry["external_spec_pattern"],
                    model_name=entry["model_name"],
                    size_option_text=entry["size_option_text"],
                )
            )
        except KeyError as exc:
            raise OrderAutomationError(f"Size mapping is missing key: {exc.args[0]}") from exc
    return mappings


def find_size_mapping(spec_text: str, mappings: Sequence[SizeMapping]) -> SizeMapping:
    normalized_spec = normalize_spec(spec_text)
    for mapping in mappings:
        if normalize_spec(mapping.external_spec_pattern) in normalized_spec:
            return mapping
    raise OrderAutomationError(f"No size mapping configured for external spec: {spec_text}")


class OrderAutomation:
    def __init__(self, config: Dict[str, Any], logger: logging.Logger, date_str: str, headful: bool = True) -> None:
        self.config = config
        self.logger = logger
        self.date_str = date_str
        self.headful = headful
        self.base_order_dir = Path(config["base_order_dir"])
        self.day_dir = self.base_order_dir / date_str
        self.browser_profile_dir = Path(config["browser_profile_dir"])
        self.log_dir = self.day_dir / "automation_logs"
        self.artifacts_dir = self.log_dir / "artifacts"
        self.size_mappings = build_size_mappings(config["size_mappings"])
        self.processed_remarks: set[str] = set()

    def run(self) -> None:
        self.validate_runtime_paths()
        assets = index_local_assets(self.day_dir)
        self.logger.info("Indexed %s local assets from %s", len(assets), self.day_dir)
        self.process_orders(assets)

    def validate_runtime_paths(self) -> None:
        if not self.day_dir.exists():
            raise OrderAutomationError(f"Date folder does not exist: {self.day_dir}")
        if not self.browser_profile_dir.exists():
            raise OrderAutomationError(f"Browser profile directory does not exist: {self.browser_profile_dir}")

    def process_orders(self, assets: Dict[str, LocalOrderAsset]) -> None:
        if sync_playwright is None:
            raise OrderAutomationError(
                "Playwright is not installed. Install it with 'pip install playwright' and "
                "then run 'playwright install chromium'."
            )

        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(self.browser_profile_dir),
                headless=not self.headful,
                accept_downloads=False,
            )
            try:
                page = context.pages[0] if context.pages else context.new_page()
                page.set_default_timeout(20_000)
                self.open_import_page(page)
                self.wait_for_manual_login_confirmation(page)
                self.select_store(page, self.config["store_name"])
                self.process_unfinished_orders(page, assets)
            except Exception as exc:
                self.capture_failure_artifacts(context, page if "page" in locals() else None, exc)
                raise
            finally:
                context.close()

    def open_import_page(self, page: Page) -> None:
        self.logger.info("Opening %s", self.config["target_url"])
        page.goto(self.config["target_url"], wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")

    def is_login_required_v2(self, page: Page) -> bool:
        login_indicators = [
            ".login-wrap",
            ".login-form",
            ".header_right_userbox .user-name",
            "input[type='password']",
            "button:has-text('登录')",
        ]
        for selector in login_indicators:
            locator = page.locator(selector)
            if not locator.count():
                continue
            try:
                if not locator.first.is_visible():
                    continue
                text = (locator.first.text_content() or "").strip()
                if selector == ".header_right_userbox .user-name" and "登录" not in text:
                    continue
                return True
            except PlaywrightError:
                continue
        return False

    def wait_for_manual_login_if_needed(self, page: Page, wait_timeout_seconds: int = 600) -> None:
        if not self.is_login_required_v2(page):
            return
        if not self.headful:
            raise OrderAutomationError(
                "Login session is not active, and headless mode cannot wait for manual login. "
                "Run with --headful true, log in manually, then rerun the script."
            )

        self.logger.warning(
            "Login session is not active. Please complete the login in the opened browser window. "
            "The script will continue automatically after login succeeds."
        )
        deadline = datetime.now().timestamp() + wait_timeout_seconds
        while datetime.now().timestamp() < deadline:
            page.wait_for_timeout(2000)
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except PlaywrightError:
                pass
            if not self.is_login_required_v2(page):
                self.logger.info("Manual login detected. Continuing automation.")
                return
        raise OrderAutomationError("Timed out waiting for manual login. Please log in and rerun the script.")

    def wait_for_manual_login_confirmation(self, page: Page) -> None:
        if not self.headful:
            self.logger.info("Headless mode: skipping manual login confirmation prompt.")
            return

        self.logger.info(
            "Please complete login in the opened browser window if needed, then press Enter here to continue."
        )
        try:
            input("Login complete? Press Enter to continue...")
        except EOFError:
            self.logger.warning("No interactive stdin available; continuing without manual confirmation input.")

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except PlaywrightError:
            pass
        self.logger.info("Manual confirmation received. Continuing automation.")

    def select_store(self, page: Page, store_name: str) -> None:
        self.logger.info("Selecting store %s", store_name)
        dropdown = self.first_visible(page, ["div.el-select", ".el-select", "div.ant-select", "[role='combobox']"])
        dropdown.click()
        option = self.locator_by_text(page, [store_name])
        option.wait_for(state="visible")
        option.click()
        page.wait_for_timeout(500)

    def process_unfinished_orders(self, page: Page, assets: Dict[str, LocalOrderAsset]) -> None:
        while True:
            row = self.find_next_order_row(page)
            if row is None:
                self.log_order_list_diagnostics(page)
                self.logger.info("No more unfinished orders found. Run completed successfully.")
                return

            remark = self.extract_row_remark(row)
            if remark in self.processed_remarks:
                self.logger.info("Encountered already-processed remark %s, stopping loop.", remark)
                return

            asset = resolve_asset_from_remark(remark, assets)
            self.logger.info(
                "Processing remark=%s match_key=%s folder=%s image=%s",
                remark,
                asset.match_key,
                asset.folder_name,
                asset.processed_image_path.name,
            )
            self.open_customization_modal(row)
            spec_text = self.read_external_spec(page)
            mapping = find_size_mapping(spec_text, self.size_mappings)
            self.logger.info(
                "Matched external spec '%s' -> model '%s' size '%s'",
                spec_text,
                mapping.model_name,
                mapping.size_option_text,
            )
            self.apply_size_mapping(page, mapping)
            self.upload_and_design(page, asset)
            self.save_design(page)
            self.processed_remarks.add(remark)
            self.return_to_order_list(page)

    def find_next_order_row(self, page: Page) -> Optional[Any]:
        row_selectors = [
            ".el-table__body-wrapper tbody tr",
            ".el-table__row",
            ".ant-table-tbody tr",
            ".order-list tr",
            "tbody tr",
        ]
        for selector in row_selectors:
            rows = page.locator(selector)
            count = rows.count()
            for index in range(count):
                row = rows.nth(index)
                if not row.is_visible():
                    continue
                if self.row_has_customize_action(row):
                    return row
        fallback_row = self.find_row_from_action_buttons(page)
        if fallback_row is not None:
            return fallback_row
        return None

    def row_has_customize_action(self, row: Any) -> bool:
        for text in ["立即定制", "去定制", "定制", "设计", "立即设计"]:
            try:
                locator = row.get_by_text(text, exact=False)
            except TypeError:
                locator = row.get_by_text(text)
            if locator.count():
                return True
        return False

    def find_row_from_action_buttons(self, page: Page) -> Optional[Any]:
        ancestor_xpaths = [
            "ancestor::tr[1]",
            "ancestor::li[1]",
            "ancestor::*[contains(@class,'row')][1]",
            "ancestor::*[contains(@class,'item')][1]",
            "ancestor::*[contains(@class,'card')][1]",
            "ancestor::*[contains(@class,'list')][1]",
        ]
        for text in ["立即定制", "去定制", "定制", "设计", "立即设计"]:
            try:
                buttons = page.get_by_text(text, exact=False)
            except TypeError:
                buttons = page.get_by_text(text)
            for index in range(buttons.count()):
                button = buttons.nth(index)
                try:
                    if not button.is_visible():
                        continue
                except Exception:
                    continue
                for xpath in ancestor_xpaths:
                    container = button.locator(f"xpath={xpath}")
                    if not container.count():
                        continue
                    try:
                        content = (container.first.text_content() or "").strip()
                    except Exception:
                        continue
                    if MATCH_KEY_PATTERN.search(content) or MATCH_ID_PATTERN.search(content) or "备注" in content:
                        return container.first
        return None

    def log_order_list_diagnostics(self, page: Page) -> None:
        row_counts: Dict[str, int] = {}
        for selector in [
            ".el-table__body-wrapper tbody tr",
            ".el-table__row",
            ".ant-table-tbody tr",
            ".order-list tr",
            "tbody tr",
        ]:
            try:
                row_counts[selector] = page.locator(selector).count()
            except Exception:
                row_counts[selector] = -1

        action_counts: Dict[str, int] = {}
        for text in ["立即定制", "去定制", "定制", "设计", "立即设计"]:
            try:
                action_counts[text] = page.get_by_text(text, exact=False).count()
            except TypeError:
                action_counts[text] = page.get_by_text(text).count()
            except Exception:
                action_counts[text] = -1

        user_text = ""
        try:
            user_locator = page.locator(".header_right_userbox .user-name")
            if user_locator.count():
                user_text = (user_locator.first.text_content() or "").strip()
        except Exception:
            pass

        self.logger.info(
            "Order list diagnostics: user_header=%r row_counts=%s action_counts=%s title=%r url=%s",
            user_text,
            row_counts,
            action_counts,
            page.title(),
            page.url,
        )
        self.save_debug_snapshot(page, prefix="no_rows")

    def save_debug_snapshot(self, page: Page, prefix: str) -> None:
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = self.artifacts_dir / f"{prefix}_{timestamp}.png"
        html_path = self.artifacts_dir / f"{prefix}_{timestamp}.html"
        try:
            page.screenshot(path=str(screenshot_path), full_page=True)
            self.logger.info("Saved diagnostic screenshot to %s", screenshot_path)
        except Exception as screenshot_exc:  # pragma: no cover
            self.logger.info("Failed to save diagnostic screenshot: %s", screenshot_exc)
        try:
            html_path.write_text(page.content(), encoding="utf-8")
            self.logger.info("Saved diagnostic HTML to %s", html_path)
        except Exception as html_exc:  # pragma: no cover
            self.logger.info("Failed to save diagnostic HTML: %s", html_exc)

    def extract_row_remark(self, row: Any) -> str:
        candidate_texts = []
        for selector in [".remark", "[data-label='备注']", "td", ".cell"]:
            locator = row.locator(selector)
            for index in range(locator.count()):
                text = (locator.nth(index).text_content() or "").strip()
                if text:
                    candidate_texts.append(text)

        for text in candidate_texts:
            if MATCH_KEY_PATTERN.search(text) or MATCH_ID_PATTERN.search(text):
                return text

        full_row_text = (row.text_content() or "").strip()
        if MATCH_KEY_PATTERN.search(full_row_text) or MATCH_ID_PATTERN.search(full_row_text):
            return full_row_text
        raise OrderAutomationError(f"Could not locate remark text in row: {full_row_text}")

    def open_customization_modal(self, row: Any) -> None:
        for text in ["立即定制", "去定制", "定制", "立即设计", "设计"]:
            try:
                button = row.get_by_text(text, exact=False)
            except TypeError:
                button = row.get_by_text(text)
            if button.count():
                button.first.click()
                return
        raise OrderAutomationError("Could not find a customization action button in the order row.")

    def read_external_spec(self, page: Page) -> str:
        for label in ["外部订单规格", "订单规格", "规格"]:
            labelled = page.get_by_text(label, exact=False)
            if labelled.count():
                text = labelled.first.locator("xpath=..").text_content() or labelled.first.text_content() or ""
                match = re.search(r"([0-9]{1,2}\s*[xX]\s*[0-9]{1,2}\s*(?:INCH|\"))", text, re.IGNORECASE)
                if match:
                    return match.group(1)

        modal_text = self.first_visible(page, [".el-dialog", ".ant-modal", "[role='dialog']"]).text_content() or ""
        match = re.search(r"([0-9]{1,2}\s*[xX]\s*[0-9]{1,2}\s*(?:INCH|\"))", modal_text, re.IGNORECASE)
        if not match:
            raise OrderAutomationError("Could not read external order spec from the customization modal.")
        return match.group(1)

    def apply_size_mapping(self, page: Page, mapping: SizeMapping) -> None:
        model = self.locator_by_text(page, [mapping.model_name])
        model.wait_for(state="visible")
        model.click()
        size_option = self.locator_by_text(page, [mapping.size_option_text])
        size_option.wait_for(state="visible")
        size_option.click()
        confirm = self.locator_by_text(page, ["确认", "确定"])
        confirm.wait_for(state="visible")
        confirm.click()
        page.wait_for_load_state("networkidle")

    def upload_and_design(self, page: Page, asset: LocalOrderAsset) -> None:
        self.logger.info("Uploading image %s", asset.processed_image_path)
        self.locator_by_text(page, ["设计"]).click()
        self.locator_by_text(page, ["上传"]).click()
        self.set_upload_file(page, asset.processed_image_path)
        self.wait_for_uploaded_material(page)
        self.click_uploaded_material(page)
        self.apply_orientation_adjustment(page, asset.orientation)

    def set_upload_file(self, page: Page, image_path: Path) -> None:
        for selector in ["input[type='file']", "input.el-upload__input"]:
            locator = page.locator(selector)
            if locator.count():
                locator.first.set_input_files(str(image_path))
                return
        raise OrderAutomationError("Upload input was not found on the design page.")

    def wait_for_uploaded_material(self, page: Page) -> None:
        for selector in [".material-list img", ".my-material img", ".el-image img", "img"]:
            locator = page.locator(selector)
            try:
                locator.first.wait_for(state="visible", timeout=30_000)
                return
            except PlaywrightError:
                continue
        raise OrderAutomationError("Uploaded material did not appear in the material panel.")

    def click_uploaded_material(self, page: Page) -> None:
        for selector in [".material-list img", ".my-material img", ".el-image img"]:
            locator = page.locator(selector)
            if locator.count():
                locator.first.click()
                return
        raise OrderAutomationError("Could not click uploaded material thumbnail.")

    def apply_orientation_adjustment(self, page: Page, orientation: str) -> None:
        if orientation == ORIENTATION_HORIZONTAL:
            action = self.locator_by_text(page, ["逆向旋转"])
            action.click()
            page.wait_for_timeout(300)
            action.click()
            self.logger.info("Applied horizontal adjustment via reverse rotate x2.")
            return
        if orientation == ORIENTATION_VERTICAL:
            self.locator_by_text(page, ["适应"]).click()
            self.logger.info("Applied vertical adjustment via fit x1.")
            return
        raise OrderAutomationError(f"Unsupported orientation: {orientation}")

    def save_design(self, page: Page) -> None:
        self.locator_by_text(page, ["保存"]).click()
        confirm = self.locator_by_text(page, ["确定", "确认"])
        confirm.wait_for(state="visible", timeout=20_000)
        confirm.click()
        page.wait_for_load_state("networkidle")
        self.logger.info("Design saved successfully.")

    def return_to_order_list(self, page: Page) -> None:
        for label in ["返回订单列表", "返回", "关闭"]:
            locator = page.get_by_text(label, exact=False)
            if locator.count():
                locator.first.click()
                page.wait_for_load_state("networkidle")
                return
        page.goto(self.config["target_url"], wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")

    def locator_by_text(self, page: Page, texts: Sequence[str]) -> Any:
        for text in texts:
            locator = page.get_by_text(text, exact=False)
            if locator.count():
                return locator.first
        raise OrderAutomationError(f"Could not find any locator by text: {texts}")

    def first_visible(self, page: Page, selectors: Sequence[str]) -> Any:
        for selector in selectors:
            locator = page.locator(selector)
            if locator.count():
                try:
                    locator.first.wait_for(state="visible", timeout=10_000)
                    return locator.first
                except PlaywrightError:
                    continue
        raise OrderAutomationError(f"None of the selectors were visible: {selectors}")

    def capture_failure_artifacts(self, context: BrowserContext, page: Optional[Page], exc: Exception) -> None:
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = self.artifacts_dir / f"failure_{timestamp}.png"
        html_path = self.artifacts_dir / f"failure_{timestamp}.html"
        self.logger.error("Automation failed: %s", exc)

        if page is not None:
            try:
                page.screenshot(path=str(screenshot_path), full_page=True)
                self.logger.error("Saved screenshot to %s", screenshot_path)
            except Exception as screenshot_exc:  # pragma: no cover
                self.logger.error("Failed to save screenshot: %s", screenshot_exc)
            try:
                html_path.write_text(page.content(), encoding="utf-8")
                self.logger.error("Saved HTML to %s", html_path)
            except Exception as html_exc:  # pragma: no cover
                self.logger.error("Failed to save HTML snapshot: %s", html_exc)

        for browser_page in context.pages:
            browser_page.close()


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    config = load_config(args.config)
    day_dir = Path(config["base_order_dir"]) / args.date
    logger = setup_logger(day_dir / "automation_logs")
    try:
        automation = OrderAutomation(config=config, logger=logger, date_str=args.date, headful=args.headful)
        automation.run()
    except OrderAutomationError as exc:
        logger.error("Fatal automation error: %s", exc)
        return 1
    except Exception as exc:  # pragma: no cover
        logger.exception("Unexpected automation failure: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
